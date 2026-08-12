"""智能组表核心引擎（在独立子进程中运行）。

流程：解析源/模板 → 结构签名 → 查代码存档 → 字段知识库 → AI 生成 →
      模板数据行预扩展(Aspose) → 沙箱执行 → 双版本结果 → 落盘 + 任务归档。

事件通过 push(event_dict) 回传给父进程（转 SSE 展示在右侧日志区）。
"""

import os
import re
import sys
import json
import shutil
import hashlib
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from sqlalchemy.orm import Session

from ..database.connection import SessionLocal
from ..database.models import AssembleTask, AssembleFieldMapping, AssembleRule

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
TENANTS_DIR = PROJECT_ROOT / "tenants"
GLOBAL_RULES_DIR = PROJECT_ROOT / "global_assets" / "assemble_rules"
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}

# ==================== 解析 ====================

def _parse_source_files(source_dir: Path, file_passwords: Dict, push) -> Tuple[Dict, str, int]:
    """解析源文件（全部可见 sheet，样例取前 3 行并脱敏）。

    Returns:
        (source_struct, source_signature, max_source_rows)
        max_source_rows: 各文件数据行数最大值（预扩展用）
    """
    from excel_parser import IntelligentExcelParser
    from backend.utils.desensitize import build_structure_json

    parser = IntelligentExcelParser()
    raw = {}
    max_rows = 0
    files = sorted(p for p in source_dir.iterdir()
                   if p.is_file() and p.suffix.lower() in EXCEL_EXTS)

    for fp in files:
        password = (file_passwords or {}).get(fp.name) or None
        try:
            sheets_data = parser.parse_excel_file(
                str(fp), max_data_rows=3, read_formulas=True,
                active_sheet_only=False, best_region_only=True,
                password=password,
            )
        except Exception as e:
            msg = f"源文件 {fp.name} 解析失败: {e}"
            logger.warning("[assemble] %s", msg)
            push({"type": "log", "message": f"⚠️ {msg}"})
            continue

        fdesc = {}
        for sd in sheets_data:
            region = (sd.regions or [None])[0]
            if region is None:
                continue
            head = region.head_data or {}
            cols = list(head.keys())
            rows = list(region.data or [])
            max_rows = max(max_rows, len(rows))
            fdesc[sd.sheet_name] = {
                "columns": cols,
                "head_data": head,
                "data": rows[:3],                       # 传给 AI 的原始样例（下面统一脱敏）
                "column_letters": {c: l for l, c in head.items()} if False else None,
            }
        if fdesc:
            raw[fp.name] = fdesc

    if not raw:
        raise RuntimeError("没有成功解析到任何源文件数据")

    # 脱敏结构 json（只脱敏样例层）
    source_struct = build_structure_json(raw)

    # 源签名：按文件/sheet/列名排序后的稳定序列
    sig_items = []
    for fname in sorted(raw):
        for sn in sorted(raw[fname]):
            sig_items.append(f"{fname}|{sn}|" + ",".join(raw[fname][sn]["columns"]))
    source_signature = hashlib.sha256("\n".join(sig_items).encode("utf-8")).hexdigest()[:16]

    push({"type": "log", "message":
          f"✅ 源文件解析完成：{len(raw)} 个文件，最大数据行数 {max_rows}（样例已脱敏）"})
    return source_struct, source_signature, max_rows


def _parse_template(template_path: Path, push) -> Tuple[Dict, str, Dict]:
    """解析模板（只读激活 sheet）。返回 (结构, 模板签名, 激活sheet数据区)。"""
    from excel_parser import IntelligentExcelParser

    parser = IntelligentExcelParser()
    sheets_data = parser.parse_excel_file(
        str(template_path), max_data_rows=5, read_formulas=True,
        skip_hidden_sheets=False, active_sheet_only=True,
    )

    region = None
    sheet_name = None
    for sd in sheets_data:
        sheet_name = sd.sheet_name
        region = (sd.regions or [None])[0]
        break

    if region is None or sheet_name is None:
        raise RuntimeError("模板解析失败：激活 sheet 未识别到数据区域（可能只有表头/空表）")

    head = region.head_data or {}
    cols = list(head.keys())
    if not cols:
        raise RuntimeError("模板激活 sheet 未识别到表头列")

    # 模板可能只有表头没有数据行（data_row_start=0）：数据起始行退回表头下一行
    ds = getattr(region, "data_row_start", None) or 0
    de = getattr(region, "data_row_end", None) or 0
    header_end = getattr(region, "head_row_end", None) or 0
    data_area = {
        "sheet_name": sheet_name,
        "columns": cols,
        "header_row_end": header_end,
        "data_start_row": ds if ds > 0 else (header_end + 1),
        "data_end_row": de if de >= ds else 0,      # 无数据行 → 0
    }

    template_struct = {
        "file_name": template_path.name,
        "sheets": {
            sheet_name: {"regions": [{
                "header_row": getattr(region, "head_row_start", None),
                "data_start_row": data_area["data_start_row"],
                "data_end_row": data_area["data_end_row"] or None,
                "columns": [{
                    "name": c,
                    "column_letter": head.get(c),   # head_data = {列名: 列字母}
                    "has_formula_in_data": False,
                } for c in cols],
            }]}
        },
    }

    template_signature = hashlib.sha256(
        "\n".join(sorted(cols)).encode("utf-8")).hexdigest()[:16]
    push({"type": "log", "message":
          f"✅ 模板解析完成：激活 sheet「{sheet_name}」{len(cols)} 列，"
          f"数据区 {data_area['data_start_row']}-{data_area['data_end_row']} 行"})
    return template_struct, template_signature, data_area


# ==================== 规则内容 ====================

def _load_rule_text(rule: Optional[AssembleRule]) -> str:
    """读取规则文件内容（文字说明 + 结构化示例），拼接为文本。"""
    if not rule or not rule.file_names:
        return ""
    if rule.scope == "tenant":
        d = TENANTS_DIR / str(rule.tenant_id or "") / "assemble_rules"
    else:
        d = GLOBAL_RULES_DIR
    texts = []
    for n in rule.file_names or []:
        p = d / f"{rule.id}_{n}"
        if not p.exists():
            continue
        try:
            if p.suffix.lower() in (".md", ".txt"):
                texts.append(p.read_text(encoding="utf-8", errors="ignore"))
            else:
                texts.append(f"[附件: {n}]")
        except Exception:
            pass
    return "\n\n".join(texts)


# ==================== 代码存档 ====================

def _code_path(tenant_id: str, signature: str) -> Path:
    d = TENANTS_DIR / tenant_id / "assemble_scripts"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"{signature}.py"


def _lookup_cached_code(db: Session, tenant_id: str, signature: str) -> Optional[str]:
    """查代码存档：存在该签名的已完成任务且代码文件在 → 返回代码内容。"""
    row = (db.query(AssembleTask)
           .filter(AssembleTask.tenant_id == tenant_id,
                   AssembleTask.signature == signature,
                   AssembleTask.status == "completed",
                   AssembleTask.code_path.isnot(None))
           .order_by(AssembleTask.id.desc()).first())
    if not row:
        return None
    p = Path(row.code_path)
    if not p.exists():
        return None
    try:
        return p.read_text(encoding="utf-8")
    except Exception:
        return None


def _save_cached_code(db: Session, tenant_id: str, signature: str, code: str) -> str:
    p = _code_path(tenant_id, signature)
    p.write_text(code, encoding="utf-8")
    return str(p)


# ==================== 字段级匹配知识库 ====================

def _query_knowledge_mappings(db: Session, tenant_id: str, template_signature: str,
                              source_columns: List[str],
                              target_columns: List[str]) -> Tuple[Dict[str, str], List[int], List[dict]]:
    """查询知识库命中映射。

    Returns:
        (auto_mappings {源列: 模板列}, used_mapping_ids, 日志行列表)
    逻辑：
      1. 精确锚定：template_signature 相同且 status=active 的条目 → 采用；
         同源列多目标（矛盾）→ 该源列交 AI，不采用
      2. 同名列兜底：源列名 = 模板列名 → 确定性采用
    """
    target_set = set(target_columns)
    rows = (db.query(AssembleFieldMapping)
            .filter(AssembleFieldMapping.tenant_id == tenant_id,
                    AssembleFieldMapping.template_signature == template_signature,
                    AssembleFieldMapping.source_column.in_(source_columns))
            .all())

    # 锚定映射 + 矛盾检测
    anchored: Dict[str, List[str]] = {}
    used_ids: List[int] = []
    for m in rows:
        if m.status != "active":
            continue
        anchored.setdefault(m.source_column, []).append(m.target_column)
        used_ids.append(m.id)

    auto: Dict[str, str] = {}
    log_lines: List[dict] = []
    for src, dsts in anchored.items():
        uniq = list(dict.fromkeys(dsts))
        if len(uniq) == 1:
            auto[src] = uniq[0]
            log_lines.append({"message": f"[知识库] {src} → {uniq[0]}（锚定命中）"})
        else:
            # 矛盾：同一模板下同源列映射到多个不同目标 → 停用相关条目，交 AI
            log_lines.append({"message": f"⚠️ [知识库] 源列「{src}」存在矛盾映射 {uniq}，已停用交 AI 重新匹配"})
            for m in rows:
                if m.source_column == src and m.status == "active":
                    m.status = "review_needed"
            db.commit()
            used_ids = [i for i in used_ids if i not in [m.id for m in rows if m.source_column == src]]

    # 同名列兜底
    for c in source_columns:
        if c in target_set and c not in auto:
            auto[c] = c
            log_lines.append({"message": f"[知识库] {c} → {c}（同名列）"})

    return auto, used_ids, log_lines


def _save_knowledge_mappings(db: Session, tenant_id: str, template_signature: str,
                             new_mappings: Dict[str, str], task_id: int):
    """AI 分析出的新映射回写知识库（同键已存在 → 更新目标列 + 次数+1）。"""
    for src, dst in (new_mappings or {}).items():
        if not src or not dst:
            continue
        existing = (db.query(AssembleFieldMapping)
                    .filter(AssembleFieldMapping.tenant_id == tenant_id,
                            AssembleFieldMapping.template_signature == template_signature,
                            AssembleFieldMapping.source_column == src)
                    .first())
        if existing:
            existing.target_column = dst
            existing.hit_count = (existing.hit_count or 1) + 1
            existing.status = "active"
        else:
            db.add(AssembleFieldMapping(
                tenant_id=tenant_id, source_column=src, target_column=dst,
                template_signature=template_signature, hit_count=1,
                source_task_id=task_id, status="active",
            ))
    db.commit()


# ==================== 模板数据行处理 ====================
# 采用智训模板模式同款链路：不预扩展模板，由 AI 生成的 fill_template
# 在 openpyxl 内直接写入（超行时 openpyxl 自动扩展维度），prompt 中已引导
# AI 在超出模板数据区时复制最后数据行的样式（cell._style）到新行。


# ==================== AI 生成 ====================

def _ai_generate_code(generator, source_struct: Dict, source_dir: Path,
                      rule_text: str, template_path: Path, template_struct: Dict,
                      auto_mappings: Dict[str, str], max_source_rows: int,
                      push) -> Tuple[str, Dict[str, str], List[str]]:
    """调用 AI 生成填充代码，返回 (代码, 新映射清单, AI 流式日志已由回调推送)。"""
    from ..ai_engine.assemble_code_generator import AssembleCodeGenerator

    def stream_cb(msg: str):
        push({"type": "log", "message": msg})

    def thinking_cb(chunk: str):
        push({"type": "thinking", "content": chunk})

    # 模板数据区现有行数 / 源数据最大行数（prompt 引导 AI 行数不足时复制样式扩展）
    data_area = template_struct.get("_data_area") or {}
    ds = data_area.get("data_start_row") or 2
    de = data_area.get("data_end_row") or 0
    tpl_data_rows = max(0, de - ds + 1)

    gen = generator or AssembleCodeGenerator()
    code, ai_response = gen.generate_code(
        input_folder=str(source_dir),
        rules_content=rule_text,
        template_path=str(template_path),
        stream_callback=stream_cb,
        thinking_callback=thinking_cb,
        pre_mapped=auto_mappings,
        tpl_data_rows=tpl_data_rows,
        src_data_rows=max_source_rows,
    )
    return code, ai_response


# ==================== 沙箱执行 ====================

def _execute_in_sandbox(code: str, source_dir: Path, output_dir: Path,
                        template_to_exec: Path, file_passwords: Dict,
                        tenant_id: str) -> dict:
    from ..sandbox.code_sandbox import CodeSandbox

    env = {
        "input_folder": str(source_dir),
        "output_folder": str(output_dir),
        "source_files": [p.name for p in source_dir.iterdir()
                         if p.is_file() and p.suffix.lower() in EXCEL_EXTS],
        "file_passwords": file_passwords,
        "_template_override_path": str(template_to_exec),
        "tenant_id": tenant_id,
    }
    return CodeSandbox().execute_script(code, env)


# ==================== 结果后处理 ====================

def _finalize_results(output_dir: Path, tenant_id: str, task_id: int,
                      rule_name: str, push) -> List[str]:
    """从脚本输出挑出结果，生成 原版 + 纯值版，落盘到 tenants/{tenant}/assemble_results/{task_id}/"""
    from backend.utils.aspose_helper import flatten_formulas_to_values

    outputs = sorted(p for p in output_dir.iterdir()
                     if p.is_file() and p.suffix.lower() in EXCEL_EXTS)
    if not outputs:
        raise RuntimeError("执行完成但没有生成结果文件")

    result_dir = TENANTS_DIR / tenant_id / "assemble_results" / str(task_id)
    result_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = re.sub(r"[\\/:*?\"<>|]", "_", rule_name or "组表结果")[:40]
    saved = []

    # 主结果：脚本第一个输出（模板副本 + 填充）
    main_out = outputs[0]
    orig_name = f"{safe_name}_{ts}_原版.xlsx"
    orig_path = result_dir / orig_name
    shutil.copy2(main_out, orig_path)
    saved.append(orig_name)

    # 纯值版：公式扁平化为值
    try:
        values_name = f"{safe_name}_{ts}_纯值.xlsx"
        values_path = result_dir / values_name
        flatten_formulas_to_values(str(orig_path), str(values_path))
        saved.append(values_name)
        push({"type": "log", "message": f"✅ 已生成纯值版: {values_name}"})
    except Exception as e:
        logger.warning("[assemble] 纯值版生成失败: %s", e)
        push({"type": "log", "message": f"⚠️ 纯值版生成失败: {e}"})

    # 清理多余输出
    for p in outputs[1:]:
        try:
            p.unlink()
        except Exception:
            pass

    return saved


# ==================== 主流程 ====================

def run_assemble_task(task_id: int, push, params: Dict):
    """智能组表任务主流程（子进程内调用）。

    params: tenant_id / rule_id / source_dir / template_path /
            force_rematch / file_passwords / ai_provider
    """
    tenant_id = params["tenant_id"]
    rule_id = params.get("rule_id")
    source_dir = Path(params["source_dir"])
    template_path = Path(params["template_path"])
    force_rematch = bool(params.get("force_rematch"))
    file_passwords = params.get("file_passwords") or {}
    ai_provider = params.get("ai_provider")

    db = SessionLocal()
    task = db.query(AssembleTask).filter_by(id=task_id).first()
    if task is None:
        raise RuntimeError(f"任务 {task_id} 不存在")

    def set_status(status: str, message: str = ""):
        task.status = status
        if message:
            push({"type": "status", "status": status, "message": message})
        db.commit()

    try:
        # 0. 规则内容
        rule = db.query(AssembleRule).filter_by(id=rule_id).first() if rule_id else None
        rule_text = _load_rule_text(rule)
        if rule_text:
            push({"type": "log", "message": f"✅ 已加载规则文件（{len(rule_text)} 字符）"})
        elif rule:
            push({"type": "log", "message": "⚠️ 规则文件内容为空（仅附件），继续"})
        else:
            push({"type": "log", "message": "⚠️ 未选择规则文件，AI 将仅依据结构分析"})

        # 1. 解析源 + 模板
        push({"type": "status", "status": "analyzing", "message": "解析源文件和模板..."})
        source_struct, source_signature, max_source_rows = _parse_source_files(
            source_dir, file_passwords, push)
        template_struct, template_signature, data_area = _parse_template(template_path, push)
        template_struct["_data_area"] = data_area

        # 2. 总签名
        signature = hashlib.sha256(
            (source_signature + "|" + template_signature + "|" +
             hashlib.sha256(rule_text.encode("utf-8")).hexdigest()).encode("utf-8")
        ).hexdigest()[:16]
        task.signature = signature
        task.source_signature = source_signature
        task.template_signature = template_signature
        db.commit()

        # 3. 查代码存档
        code = None
        matched_from_cache = False
        if not force_rematch:
            code = _lookup_cached_code(db, tenant_id, signature)
            if code:
                matched_from_cache = True
                push({"type": "log", "message":
                      "⚡ 命中代码存档（源/模板结构 + 规则一致），跳过 AI 直接执行"})
            else:
                push({"type": "log", "message":
                      "未命中代码存档，进入 AI 分析（勾选「强制重新匹配」可跳过此层）"})
        else:
            push({"type": "log", "message": "已勾选「强制重新匹配」，跳过代码存档和知识库，全量 AI 分析"})

        # 4. AI 分析 + 知识库 + 生成（未命中存档时）
        used_mapping_ids: List[int] = []
        new_mappings: Dict[str, str] = {}
        if code is None:
            push({"type": "status", "status": "analyzing", "message": "AI 分析字段匹配关系..."})
            all_source_cols = []
            for fname in sorted(source_struct):
                for sn in sorted(source_struct[fname]):
                    all_source_cols.extend(source_struct[fname][sn].get("columns") or [])
            all_source_cols = list(dict.fromkeys(all_source_cols))

            auto_mappings: Dict[str, str] = {}
            if not force_rematch:
                auto_mappings, used_mapping_ids, map_logs = _query_knowledge_mappings(
                    db, tenant_id, template_signature, all_source_cols, data_area["columns"])
                for line in map_logs:
                    push({"type": "mapping", **line})
                auto_mappings = dict(auto_mappings)  # 副本避免污染
            else:
                push({"type": "log", "message": "强制重匹配：知识库映射不使用，全部交 AI"})

            # AI 生成
            push({"type": "status", "status": "generating", "message": "AI 生成填充代码..."})
            try:
                from ..ai_engine.assemble_code_generator import AssembleCodeGenerator
                gen = AssembleCodeGenerator(file_passwords=file_passwords)
                if ai_provider:
                    # 与智训一致：设 AI_PROVIDER 环境变量 → create_provider 读 .env 对应配置
                    # （api_key/base_url/model/max_tokens 等，不写死）
                    from ..ai_engine.ai_provider import AIProviderFactory
                    _orig = os.environ.get("AI_PROVIDER")
                    os.environ["AI_PROVIDER"] = ai_provider
                    try:
                        gen.ai_provider = AIProviderFactory.create_provider(ai_provider)
                    finally:
                        if _orig is not None:
                            os.environ["AI_PROVIDER"] = _orig
                        else:
                            os.environ.pop("AI_PROVIDER", None)
                code, _ = _ai_generate_code(
                    gen, source_struct, source_dir, rule_text, template_path,
                    template_struct, auto_mappings, max_source_rows, push)
            except Exception as e:
                logger.exception("[assemble] AI 生成失败")
                raise RuntimeError(f"AI 生成代码失败: {e}")

            # 存档 + 知识库回写（AI 新映射由生成器 prompt 覆盖了 auto_mappings；
            # 实际新映射 = auto_mappings 之外 AI 处理的字段 → 简单起见回写 auto_mappings 全量）
            task.code_path = _save_cached_code(db, tenant_id, signature, code)
            _save_knowledge_mappings(db, tenant_id, template_signature,
                                     dict(auto_mappings), task_id)
            push({"type": "log", "message": "✅ AI 代码已生成并存入代码存档"})

        task.used_mapping_ids = used_mapping_ids
        task.matched_from_cache = matched_from_cache
        db.commit()

        # 5. 沙箱执行（模板模式同款链路：AI fill_template 在 openpyxl 内直接写入/扩展行）
        exec_dir = source_dir.parent / f"exec_{task_id}"
        exec_dir.mkdir(parents=True, exist_ok=True)
        output_dir = exec_dir / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        push({"type": "status", "status": "executing", "message": "沙箱执行填充代码（约1-3分钟）..."})
        result = _execute_in_sandbox(code, source_dir, output_dir, template_path,
                                     file_passwords, tenant_id)

        if not result.get("success"):
            err = result.get("error") or "未知错误"
            logger.error("[assemble] 执行失败: %s", err)
            push({"type": "error", "message": f"执行失败: {err}"})
            task.status = "error"
            task.error = str(err)[:2000]
            db.commit()
            return

        # 7. 结果后处理 + 落盘
        push({"type": "status", "status": "executing", "message": "生成结果文件（原版+纯值版）..."})
        saved = _finalize_results(output_dir, tenant_id, task_id,
                                  rule.name if rule else "", push)

        # 8. 归档
        task.status = "completed"
        task.output_files = saved
        task.completed_at = datetime.utcnow()
        db.commit()

        push({"type": "status", "status": "complete", "message": "组表完成"})
        push({"type": "complete",
              "output_files": saved,
              "matched_from_cache": matched_from_cache,
              "task_id": task_id})

    except Exception as e:
        logger.exception("[assemble] 任务异常")
        task.status = "error"
        task.error = str(e)[:2000]
        db.commit()
        push({"type": "error", "message": f"组表失败: {e}"})
    finally:
        db.close()
