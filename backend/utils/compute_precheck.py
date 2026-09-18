"""智算事前校验模块

在 compute_submit 同步阶段拦截，避免事后脚本运行时炸出难定位的二手错误。

实现委托给 `compute_ingest`：源文件只被解析一次，改名候选 / 低置信列名 /
目标表歧义 / 历史缺口在**同一轮**里全部收集返回（前端一次弹窗即可确认完）。
本模块保留 PrecheckResult 结构、目标表与历史校验，以及给前端的默认建议算法。

confirmed_mapping 透传：用户在前端弹窗确认 AI 建议后，重提时携带，本模块直接使用、跳过 AI 步骤。
"""

import os
import re
import json
import logging
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# 历史数据关键字（脚本含任一即视为依赖历史数据）
_HISTORY_KEYWORDS = [
    "history_provider",
    "load_history",
    "get_available_months",
    "get_employee_history",
    "HistoricalDataProvider",
    "历史数据",  # sheet 名引用
]


@dataclass
class PrecheckResult:
    ok: bool = True
    missing_files: List[str] = field(default_factory=list)
    auto_filled: List[Dict[str, Any]] = field(default_factory=list)
    auto_renamed: List[Dict[str, Any]] = field(default_factory=list)
    rename_candidates: List[Dict[str, Any]] = field(default_factory=list)
    missing_columns: List[Dict[str, Any]] = field(default_factory=list)
    ai_suggestions: List[Dict[str, Any]] = field(default_factory=list)
    # 上传文件的实际列路径全集（file > sheet > col），供前端手动选择下拉全量列出
    actual_paths: List[str] = field(default_factory=list)
    history_warnings: List[str] = field(default_factory=list)
    # 目标模板表（②模板目标侧）：歧义时的候选，交前端人工选择；确认后的映射透传给计算
    target_candidates: List[Dict[str, Any]] = field(default_factory=list)
    target_map: Optional[Dict[str, str]] = None
    file_mapping: Optional[Dict[str, Any]] = None
    unmatched_columns: List[list] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


def precheck_compute(
    source_dir: str,
    source_structure: dict,
    manual_headers: Optional[dict],
    script_content: str,
    tenant_id: str,
    salary_year: Optional[int],
    salary_month: Optional[int],
    db_session,
    ai_provider_name: Optional[str] = None,
    confirmed_mapping: Optional[Dict[str, Any]] = None,
    confirmed_renames: Optional[Dict[str, str]] = None,
    use_history: Optional[bool] = None,
    expected_structure: Optional[dict] = None,
    template_override_path: Optional[str] = None,
    confirmed_target_map: Optional[Dict[str, str]] = None,
    in_worker: bool = False,
    session_dir: Optional[str] = None,
    skipped_missing_files: Optional[List[str]] = None,
) -> PrecheckResult:
    """智算事前校验主入口（compute_ingest 的薄封装）

    整个校验只解析源文件一次：ingest 负责解密/解析/改名/兜底，
    resolve_with_confirmations 负责纯内存的映射解算与待确认项收集。

    use_history: 训练时记录的"使用历史数据"开关
        - True/False: 显式开关，覆盖关键字检测
        - None     : 未训练标记(存量脚本)，回退到关键字扫描
    session_dir: 传入时把解析产物（meta + sources）落盘，供人工确认轮复用，
        确认时无需重传文件、无需重新解析。
    """
    result = PrecheckResult()
    # 前端只确认历史提示或未选择列时会提交空封装，不能当作映射已通过。
    if isinstance(confirmed_mapping, dict) and 'file_mapping' in confirmed_mapping and not confirmed_mapping.get('unmatched_columns'):
        confirmed_mapping = confirmed_mapping.get('file_mapping') or None

    if not source_structure:
        # 老脚本可能没存 source_structure，无法校验，直接放行
        logger.info("[Precheck] 缺少 source_structure，跳过校验")
        return result

    payload = {
        "source_dir": source_dir,
        "source_structure": source_structure,
        "manual_headers": manual_headers,
        "expected_structure": expected_structure,
        "tenant_id": tenant_id,
        "ai_provider_name": ai_provider_name,
        "salary_year": salary_year,
        "salary_month": salary_month,
        "confirmed_renames": confirmed_renames,
        "confirmed_mapping": confirmed_mapping,
        "confirmed_target_map": confirmed_target_map,
        "script_content": script_content,
        "use_history": use_history,
        "template_override_path": template_override_path,
        "session_dir": session_dir,
        "skipped_missing_files": skipped_missing_files,
    }
    if in_worker:
        # 外层已有总超时/内存护栏，直接解析避免嵌套进程和大表 pickle 往返。
        return _ingest_and_resolve(payload, keep_preload=True, db_session=db_session)

    # 全量解析（多文件 Aspose 打开）在【独立子进程】执行：
    # Aspose 持 GIL 会冻结主进程（其他请求全卡 → IIS 502 / 服务器假死），
    # 日志实证：主进程跑几十个文件解析可卡 10-30 分钟。子进程内爆只炸自己。
    from backend.utils.subprocess_runner import (
        run_in_subprocess, default_max_memory_mb, default_timeout,
    )
    _r = run_in_subprocess(
        "backend.utils.compute_precheck:_ingest_subprocess", (payload,),
        timeout=default_timeout("parse"), max_memory_mb=default_max_memory_mb())
    if _r.success:
        return _r.result
    reason = "超时" if _r.timed_out else ("内存超限" if _r.killed_by_memory else _r.error)
    result.ok = False
    result.missing_columns = [{"file": "", "sheet": "", "expected_columns": [],
                              "error": f"解析子进程失败（{reason}）"}]
    return result


def _ingest_subprocess(payload: dict) -> PrecheckResult:
    """模块级包装（subprocess_runner 定位入口）：解析 + 解算在独立子进程执行。

    预加载数据不回传（避免把大表 pickle 进 API 进程），只落盘给计算进程用。
    """
    return _ingest_and_resolve(payload, keep_preload=False)


def _ingest_and_resolve(payload: dict, keep_preload: bool, db_session=None) -> PrecheckResult:
    """解析一次 + 解算映射；有 session_dir 时把产物落盘供确认轮复用。"""
    own_db = db_session is None
    if own_db:
        db_session = _open_db()
    try:
        return _ingest_and_resolve_inner(payload, keep_preload, db_session)
    finally:
        if own_db and db_session is not None:
            try:
                db_session.close()
            except Exception:
                pass


def _ingest_and_resolve_inner(payload: dict, keep_preload: bool, db_session) -> PrecheckResult:
    from .compute_ingest import (
        ingest_source_dir, resolve_with_confirmations, build_preload, write_ingest,
    )

    meta, parsed_sheets_map = ingest_source_dir(
        source_dir=payload["source_dir"],
        source_structure=payload["source_structure"],
        manual_headers=payload.get("manual_headers"),
        expected_structure=payload.get("expected_structure"),
        tenant_id=payload.get("tenant_id"),
        db_session=db_session,
        ai_provider_name=payload.get("ai_provider_name"),
        salary_year=payload.get("salary_year"),
        salary_month=payload.get("salary_month"),
        confirmed_renames=payload.get("confirmed_renames"),
    )
    result = resolve_with_confirmations(
        meta,
        confirmed_renames=payload.get("confirmed_renames"),
        confirmed_mapping=payload.get("confirmed_mapping"),
        confirmed_target_map=payload.get("confirmed_target_map"),
        script_content=payload.get("script_content"),
        tenant_id=payload.get("tenant_id"),
        salary_year=payload.get("salary_year"),
        salary_month=payload.get("salary_month"),
        use_history=payload.get("use_history"),
        template_override_path=payload.get("template_override_path"),
        skipped_missing_files=payload.get("skipped_missing_files"),
    )
    session_dir = payload.get("session_dir")
    wrote_session = False
    if session_dir and parsed_sheets_map:
        try:
            write_ingest(session_dir, meta, parsed_sheets_map)
            wrote_session = True
        except Exception as e:
            logger.warning(f"[Precheck] 写会话产物失败（退回重传路径）: {e}", exc_info=True)

    # 会话产物已落盘时不再重复写一份预加载 pickle（同一批数据落两次盘），
    # 计算进程按最终映射直接 build_preload。
    if result.ok and result.file_mapping and not wrote_session:
        try:
            preload = build_preload(meta, parsed_sheets_map, result.file_mapping)
        except Exception as e:
            logger.error(f"[Precheck] 构建预加载数据失败: {e}", exc_info=True)
            preload = None
        if preload:
            if keep_preload:
                result._pre_loaded_source_data = preload
            else:
                from .compute_preload_cache import save_preload
                save_preload(payload["source_dir"], preload, result.file_mapping,
                             (payload.get("source_structure"), payload.get("manual_headers"),
                              payload.get("expected_structure")))
    return result


def _open_db():
    """子进程/线程内独立 DB 会话；失败时返回 None（基础资料兜底自动跳过）。"""
    try:
        from backend.database.connection import SessionLocal
        return SessionLocal()
    except Exception as e:
        logger.warning(f"[Precheck] 打开数据库会话失败，跳过基础资料兜底: {e}")
        return None



def _extract_colmap(script_content: str) -> Optional[Dict[str, Any]]:
    """从脚本里安全抽取 _COL_MAP 字面量（模板模式脚本训练时固化的目标表结构）。"""
    try:
        import ast
        tree = ast.parse(script_content)
        for node in tree.body:
            if isinstance(node, ast.Assign):
                for t in node.targets:
                    if isinstance(t, ast.Name) and t.id == "_COL_MAP":
                        return ast.literal_eval(node.value)
    except Exception as e:
        logger.warning(f"[Precheck/Target] 提取 _COL_MAP 失败: {e}")
    return None


def _locate_template(tenant_id: str, script_content: str, template_override_path: Optional[str]) -> Optional[str]:
    """定位当前环境下的模板：优先用户上传的覆盖模板，否则按名/哈希在本环境解析。"""
    if template_override_path and os.path.exists(template_override_path):
        return template_override_path
    try:
        from .template_resolver import resolve_template_path
        proj_root = str(Path(__file__).resolve().parent.parent.parent)
        p = resolve_template_path(tenant_id=tenant_id, script_code=script_content, project_root=proj_root)
        if p and os.path.exists(p):
            return p
    except Exception as e:
        logger.warning(f"[Precheck/Target] 定位模板失败: {e}")
    return None


def _check_target_sheets(
    script_content: Optional[str],
    tenant_id: str,
    template_override_path: Optional[str],
    confirmed_target_map: Optional[Dict[str, str]],
    result: PrecheckResult,
) -> None:
    """校验模板里的目标表能否唯一对到训练固化的 _COL_MAP 键。

    歧义（多候选并列）→ result.ok=False + target_candidates（交前端人工选择）；
    唯一解析 → 把非同名映射记入 result.target_map（计算时注入 _target_sheet_manual_map）；
    无候选的键（本月缺该表）仅记日志、不阻断。仅对模板模式脚本生效。
    """
    if not script_content or "_COL_MAP" not in script_content or "def fill_template" not in script_content:
        return
    col_map = _extract_colmap(script_content)
    if not col_map:
        return
    tpl = _locate_template(tenant_id, script_content, template_override_path)
    if not tpl:
        logger.info("[Precheck/Target] 未定位到模板，跳过目标表校验（运行时兜底/报错）")
        return
    try:
        import openpyxl
        from .target_sheet_resolver import resolve_target_sheets
        wb = openpyxl.load_workbook(tpl, read_only=False, data_only=True)
        try:
            resolved, ambiguous, unresolved = resolve_target_sheets(
                wb, col_map, manual_map=(confirmed_target_map or {})
            )
            all_sheets = [sn for sn in wb.sheetnames if not sn.startswith("源_")]
        finally:
            wb.close()
    except Exception as e:
        logger.warning(f"[Precheck/Target] 目标表校验异常（不阻断）: {e}", exc_info=True)
        return

    if ambiguous:
        result.ok = False
        result.target_candidates = [
            {"key": k, "candidates": v, "all_sheets": all_sheets}
            for k, v in ambiguous.items()
        ]
        logger.warning(f"[Precheck/Target] 目标表歧义需人工确认: {list(ambiguous.keys())}")
    else:
        # 只记非同名映射（同名的运行时精确匹配即可），供计算注入 _target_sheet_manual_map
        result.target_map = {k: v for k, v in resolved.items() if v != k}
        if unresolved:
            logger.info(f"[Precheck/Target] 目标表本月无对应（运行时落空跳过）: {unresolved}")


# ==================== 内部工具 ====================

def _extract_missing_columns(
    source_structure: dict,
    input_files: List[str],
    error_msg: Optional[str],
) -> List[Dict[str, Any]]:
    """从训练 source_structure 与上传文件名集合的差集，提取缺失列明细

    简化策略：枚举训练每个文件每个 sheet 的列，标注「未在上传文件里出现的」
    """
    uploaded_names = {os.path.basename(f).lower() for f in input_files}
    missing = []
    files = source_structure.get("files", {}) if isinstance(source_structure, dict) else {}

    for file_name, file_data in files.items():
        if not isinstance(file_data, dict):
            continue
        # 训练文件带 error 但仍有 sheets 结构时也尽力提取列（避免训练侧解析失败的文件
        # 在手动选择里整体消失、期望列全漏）
        if "error" in file_data and not file_data.get("sheets"):
            continue
        for sheet_name, sheet_info in (file_data.get("sheets") or {}).items():
            headers = sheet_info.get("headers") if isinstance(sheet_info, dict) else None
            cols = []
            if isinstance(headers, dict):
                cols = list(headers.keys())
            elif isinstance(headers, list):
                cols = list(headers)

            entry = {
                "file": file_name,
                "sheet": sheet_name,
                "expected_columns": cols,
                "uploaded_present": file_name.lower() in uploaded_names,
            }
            missing.append(entry)

    if error_msg:
        missing.append({"file": "(matcher)", "sheet": "", "expected_columns": [], "error": error_msg})
    return missing


def _suggest_structural_columns(expected_entries: list, actual_paths: list,
                                file_mapping=None, file_renames=None) -> list:
    """为确认界面提供保守的默认值，不改写文件，也不重复解析/请求 AI。"""
    def columns_index(columns):
        index = {}
        for column in columns:
            key = ''.join(str(column).split()).casefold()
            if key:
                index.setdefault(key, []).append(column)
        # 规范化后重名的列不能自动选择。
        return {key: values[0] for key, values in index.items() if len(values) == 1}

    actual = {}
    for path in dict.fromkeys(actual_paths):
        parts = path.split(' > ', 2)
        if len(parts) == 3:
            actual.setdefault(tuple(parts[:2]), []).append(parts[2])
    actual = {key: columns_index(cols) for key, cols in actual.items()}
    expected = {(entry['file'], entry['sheet']): columns_index(entry['expected_columns'])
                for entry in expected_entries if entry.get('expected_columns')}
    choices = {}
    claimed = {}
    target_sources = {}
    file_targets = dict(file_renames or {})
    for filename, info in (file_mapping or {}).items():
        file_targets[filename] = info.get('expected_file', filename)
        for sheet, target_sheet in info.get('sheet_mapping', {}).items():
            source = (filename, sheet)
            target = (file_targets[filename], target_sheet)
            claimed[source] = target
            target_sources[target] = source
    for target, columns in expected.items():
        candidates = []
        for source, source_columns in actual.items():
            if source in claimed and claimed[source] != target:
                continue
            if target in target_sources and target_sources[target] != source:
                continue
            if source[0] in file_targets and file_targets[source[0]] != target[0]:
                continue
            common = columns.keys() & source_columns.keys()
            score = len(common) / max(len(columns), len(source_columns), 1)
            if score < 0.7 or (len(common) < 2 and score != 1):
                continue
            # 完整路径相同优先，但不能只凭名称忽略列结构。
            candidates.append((score + (1 if source == target else 0), source, common))
        candidates.sort(key=lambda item: item[0], reverse=True)
        if candidates and (len(candidates) == 1 or candidates[0][0] - candidates[1][0] >= 0.1):
            choices[target] = candidates[0]

    suggestions = []
    for target, (score, source, common) in choices.items():
        # 不允许两张训练表抢用同一张表，或产生文件级映射冲突。
        if any(other != target and (candidate[1] == source or
               (other[0] == target[0] and candidate[1][0] != source[0]) or
               (other[0] != target[0] and candidate[1][0] == source[0]))
               for other, candidate in choices.items()):
            continue
        for key, column in expected[target].items():
            if key in common:
                suggestions.append({
                    'expected_path': ' > '.join((*target, str(column))),
                    'suggested_path': ' > '.join((*source, str(actual[source][key]))),
                    'confidence': min(score, 1.0),
                    'reason': '列结构匹配（表名可不同），请确认',
                })
    return suggestions


def _ai_suggest_column_mapping(
    missing: List[Dict[str, Any]],
    source_structure: dict,
    input_files: List[str],
    ai_provider_name: str,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """调用 AI 在「训练期望列」与「上传文件实际列」之间给出映射建议。

    Returns: (suggestions, actual_paths)
        - suggestions: AI 建议的映射列表（可能为空/只覆盖部分期望列）
        - actual_paths: 上传文件实际列路径全集，前端手动选择下拉全量列出
          （AI 未建议的期望列用户也要能手动指定，不能只给 AI 提到过的路径）
    """
    expected_paths = []
    for file_name, file_data in (source_structure.get("files") or {}).items():
        if not isinstance(file_data, dict) or "error" in file_data:
            continue
        for sheet_name, sheet_info in (file_data.get("sheets") or {}).items():
            headers = sheet_info.get("headers") if isinstance(sheet_info, dict) else None
            cols = list(headers.keys()) if isinstance(headers, dict) else (list(headers) if isinstance(headers, list) else [])
            for col in cols:
                expected_paths.append(f"{file_name} > {sheet_name} > {col}")

    actual_paths = _collect_uploaded_columns(input_files)
    if not expected_paths or not actual_paths:
        return [], actual_paths

    prompt = _build_column_match_prompt(expected_paths, actual_paths)

    try:
        from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout
        provider = AIProviderFactory.create_provider(ai_provider_name)
        messages = [
            {"role": "system", "content": "你是一个 Excel 表头匹配专家，擅长在不同命名习惯之间找到语义等价的列对应关系。"},
            {"role": "user", "content": prompt},
        ]
        # 带总超时：AI 卡住时跳过建议（提交请求同步等待，超网关读超时会被判 504）
        raw = chat_with_timeout(provider, messages, max_tokens=2000)
        if raw is None:
            return [], actual_paths
    except Exception as e:
        logger.warning(f"[Precheck/AI] 调用失败: {e}")
        return [], actual_paths

    return _parse_ai_response(raw), actual_paths


def _collect_uploaded_columns(input_files: List[str]) -> List[str]:
    """解析上传文件，列出 {file > sheet > col} 路径列表。

    解析失败的文件用 openpyxl 兜底提取列，保证所有上传文件都进手动选择下拉
    （上传文件多时，个别文件解析失败若被跳过，用户就选不到该文件）。
    """
    from .fast_header_matcher import fallback_headers_openpyxl
    paths = []
    try:
        from excel_parser import IntelligentExcelParser
        parser = IntelligentExcelParser()
        for fp in input_files:
            fname = os.path.basename(fp)
            got = False
            try:
                results = parser.parse_excel_file(fp, max_data_rows=1, read_formulas=False)
                for sheet_data in results:
                    sheet_name = sheet_data.sheet_name
                    for region in (sheet_data.regions or []):
                        head = region.head_data or {}
                        for col in head.keys():
                            paths.append(f"{fname} > {sheet_name} > {col}")
                got = True
            except Exception as e:
                logger.warning(f"[Precheck/AI] 解析上传文件 {fp} 失败: {e}")
            if not got:
                # openpyxl 兜底，保证文件不消失
                try:
                    fh = fallback_headers_openpyxl(fp)
                    for sheet_name, cols in fh.items():
                        for col in cols:
                            paths.append(f"{fname} > {sheet_name} > {col}")
                except Exception:
                    pass
    except Exception as e:
        logger.warning(f"[Precheck/AI] 加载 parser 失败: {e}")
    return paths


def _build_column_match_prompt(expected: List[str], actual: List[str]) -> str:
    return (
        "下面是两个列清单，每条格式为 `文件名 > Sheet名 > 列名`（Sheet名可能含 banner 后缀，如 `数据-合同工`）。\n\n"
        f"## 训练期望的列（共 {len(expected)} 项）\n"
        + "\n".join(f"- {p}" for p in expected[:200])
        + "\n\n"
        f"## 用户上传文件实际有的列（共 {len(actual)} 项）\n"
        + "\n".join(f"- {p}" for p in actual[:200])
        + "\n\n"
        "请在期望列与实际列之间做语义匹配，对每个期望列给出最可能的实际列对应（如果完全找不到合理对应可省略该项）。\n"
        "**严格只输出 JSON 数组**，每项格式：\n"
        '  {"expected_path": "...", "suggested_path": "...", "confidence": 0.0-1.0, "reason": "简短中文原因"}\n'
        "不要输出任何 JSON 之外的文字、解释、代码块标记。"
    )


def _parse_ai_response(raw: str) -> List[Dict[str, Any]]:
    if not raw:
        return []
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)

    m = re.search(r"\[[\s\S]*\]", text)
    if m:
        text = m.group(0)
    try:
        data = json.loads(text)
        if isinstance(data, list):
            cleaned = []
            for item in data:
                if not isinstance(item, dict):
                    continue
                cleaned.append({
                    "expected_path": str(item.get("expected_path", "")),
                    "suggested_path": str(item.get("suggested_path", "")),
                    "confidence": float(item.get("confidence", 0.0) or 0.0),
                    "reason": str(item.get("reason", "")),
                })
            return cleaned
    except Exception as e:
        logger.warning(f"[Precheck/AI] 响应 JSON 解析失败: {e}; raw={raw[:300]}")
    return []


def _check_history(
    script_content: Optional[str],
    tenant_id: str,
    salary_year: Optional[int],
    salary_month: Optional[int],
    result: PrecheckResult,
    use_history: Optional[bool] = None,
) -> None:
    """历史数据齐全性校验

    判定是否依赖历史数据：
      - use_history is True/False  → 直接采用(训练时显式开关)
      - use_history is None        → 回退到脚本关键字扫描(兼容存量脚本)
    """
    # 决定是否依赖历史数据
    if use_history is True:
        depends_on_history = True
    elif use_history is False:
        depends_on_history = False
    else:
        if not script_content:
            return
        depends_on_history = any(kw in script_content for kw in _HISTORY_KEYWORDS)

    if not depends_on_history:
        return
    if not salary_year or not salary_month or salary_month <= 1:
        return

    try:
        from .historical_data import HistoricalDataProvider
        provider = HistoricalDataProvider(tenant_id)
        available = set(provider.get_available_months(salary_year))
        expected = set(range(1, salary_month))
        missing = sorted(expected - available)
        if missing:
            warning = (
                f"脚本依赖历史数据，但 {salary_year} 年缺少 "
                f"{', '.join(str(m) + '月' for m in missing)} 的历史结果"
            )
            result.history_warnings.append(warning)
            logger.warning(f"[Precheck/History] {warning}")
    except Exception as e:
        logger.warning(f"[Precheck/History] 校验异常: {e}", exc_info=True)
