"""
对话式训练 API - 支持交互式代码调试
"""

import os
import json
import asyncio
import logging
import shutil
import tempfile
import time
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any, Tuple
from concurrent.futures import ThreadPoolExecutor

from fastapi import APIRouter, Depends, HTTPException, Query, Form, File, UploadFile
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from sqlalchemy.orm.attributes import flag_modified
from pydantic import BaseModel

from ..database.connection import get_db, SessionLocal
from ..database.models import (
    TrainingSession, TrainingIteration, TrainingMessage, Script, DataAsset,
)
from ..auth.dependencies import get_current_user

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/training/chat", tags=["对话式训练"])

_executor = ThreadPoolExecutor(max_workers=4)


def _create_formula_generator(ai_provider_name: str, stream_callback=None):
    """创建 FormulaCodeGenerator 实例（复用原有智训管线）"""
    from ..ai_engine.ai_provider import AIProviderFactory
    from ..ai_engine.formula_code_generator import FormulaCodeGenerator
    from ..ai_engine.training_logger import TrainingLogger

    # 设置 AI provider
    original = os.environ.get("AI_PROVIDER")
    if ai_provider_name:
        os.environ["AI_PROVIDER"] = ai_provider_name

    try:
        provider = AIProviderFactory.create_provider(ai_provider_name)
    finally:
        if original is not None:
            os.environ["AI_PROVIDER"] = original
        elif ai_provider_name:
            os.environ.pop("AI_PROVIDER", None)

    # 创建简单 logger（不写文件）
    tl = TrainingLogger("chat_training")
    if stream_callback:
        tl.set_stream_callback(stream_callback)

    generator = FormulaCodeGenerator(ai_provider=provider, training_logger=tl)
    return generator, provider


def _analyze_expected_structure(expected_file: str) -> Dict[str, Any]:
    """分析预期文件结构（与 TrainingEngine._analyze_expected_structure 一致）"""
    from excel_parser import IntelligentExcelParser

    parser = IntelligentExcelParser()
    parsed_data = parser.parse_excel_file(
        expected_file,
        max_data_rows=10,
        active_sheet_only=True,
        best_region_only=True,
    )

    structure = {
        "file_name": Path(expected_file).name,
        "sheets": {},
        "total_regions": 0,
    }

    for sheet_data in parsed_data:
        sheet_structure = {
            "sheet_name": sheet_data.sheet_name,
            "regions": len(sheet_data.regions),
            "headers": {},
            "data_sample": [],
        }
        for region in sheet_data.regions:
            sheet_structure["headers"].update(region.head_data)
            if region.data and len(sheet_structure["data_sample"]) < 3:
                sheet_structure["data_sample"].append(region.data[0])

        structure["sheets"][sheet_data.sheet_name] = sheet_structure
        structure["total_regions"] += len(sheet_data.regions)

    return structure

# ==================== 后台全量数据加载 ====================


def _load_full_source_data(source_dir: str, manual_headers: Dict = None) -> Dict:
    """全量加载源文件数据（无 max_data_rows 限制），供脚本执行时使用。
    该函数设计为在后台线程中运行，与 AI 代码生成并行执行。

    返回格式与模板代码中 load_source_data() 一致：
    {"文件名_sheet名": {"df": DataFrame, "columns": [列名]}}
    """
    import pandas as pd
    from excel_parser import IntelligentExcelParser

    source_data = {}
    parser = IntelligentExcelParser()

    for filename in sorted(os.listdir(source_dir)):
        if not filename.endswith(('.xlsx', '.xls')) or filename.startswith('~'):
            continue
        file_path = os.path.join(source_dir, filename)
        file_base = filename.replace('.xlsx', '').replace('.xls', '')

        try:
            results = parser.parse_excel_file(
                file_path,
                manual_headers=manual_headers,
                active_sheet_only=True,
                best_region_only=True,
                # 不传 max_data_rows → 加载全量数据
            )
            if not results:
                continue

            for sheet_data in results:
                dfs = []
                columns = None
                for region in sheet_data.regions:
                    # 将 ExcelRegion 转换为 DataFrame（与模板代码逻辑一致）
                    col_letter_to_name = {v: k for k, v in region.head_data.items()}
                    cols = list(region.head_data.keys())
                    if not region.data:
                        df = pd.DataFrame(columns=cols)
                    else:
                        converted = []
                        for row in region.data:
                            new_row = {col_letter_to_name.get(cl, cl): val for cl, val in row.items()}
                            converted.append(new_row)
                        df = pd.DataFrame(converted, columns=cols)

                    if df.empty and len(df.columns) == 0:
                        continue
                    if columns is None:
                        columns = list(df.columns)
                    dfs.append(df)

                if not dfs:
                    continue

                merged_df = dfs[0] if len(dfs) == 1 else pd.concat(dfs, ignore_index=True)
                sheet_name = f"{file_base}_{sheet_data.sheet_name}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                source_data[sheet_name] = {"df": merged_df, "columns": columns}
                logger.info(f"[后台全量加载] {sheet_name}: {len(merged_df)} 行")

        except Exception as e:
            logger.warning(f"[后台全量加载] 解析 {filename} 失败: {e}")

    return source_data


# ==================== 辅助函数 ====================


def _add_message(db: Session, session_id: int, role: str, content: str,
                 msg_type: str = "chat", metadata: dict = None) -> TrainingMessage:
    """添加一条消息到会话"""
    msg = TrainingMessage(
        session_id=session_id,
        role=role,
        content=content,
        msg_type=msg_type,
        metadata_=metadata,
    )
    db.add(msg)
    db.commit()
    db.refresh(msg)
    return msg


def _get_session_context(db: Session, session_id: int) -> Dict[str, Any]:
    """构建结构化上下文（固定区）"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        return {}

    # 获取最新迭代（最佳代码）
    best_iteration = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .filter(TrainingIteration.accuracy.isnot(None))
        .order_by(TrainingIteration.accuracy.desc(), TrainingIteration.iteration_num.desc())
        .first()
    )

    # 获取最近的迭代（最新执行结果）
    latest_iteration = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .order_by(TrainingIteration.iteration_num.desc())
        .first()
    )

    # 获取最近 5 条对话消息
    recent_messages = (
        db.query(TrainingMessage)
        .filter_by(session_id=session_id)
        .filter(TrainingMessage.role.in_(["user", "assistant"]))
        .order_by(TrainingMessage.created_at.desc())
        .limit(5)
        .all()
    )
    recent_messages.reverse()  # 时间正序

    context = {
        "tenant_id": session.tenant_id,
        "mode": session.mode,
        "config": session.config or {},
        "best_code": best_iteration.generated_code if best_iteration else None,
        "best_accuracy": best_iteration.accuracy if best_iteration else None,
        "latest_code": latest_iteration.generated_code if latest_iteration else None,
        "latest_accuracy": latest_iteration.accuracy if latest_iteration else None,
        "latest_diff": latest_iteration.error_details if latest_iteration else None,
        "latest_execution_result": latest_iteration.execution_result if latest_iteration else None,
        "total_iterations": session.total_iterations or 0,
        "recent_messages": [
            {"role": m.role, "content": m.content} for m in recent_messages
        ],
    }
    return context


def _build_chat_system_prompt(context: Dict, config: Dict, rules: str) -> str:
    """构建 AI 对话的 system prompt（分析/讨论模式）"""
    parts = [
        "你是专业的人力资源薪资计算顾问和 Excel/Python 自动化专家。",
        "你正在帮助用户分析、讨论和优化一个薪资数据处理脚本。",
        "请根据用户的问题进行专业的分析和讨论。",
        "如果用户确认了修改方案，请清楚总结需要修改的内容，用户会在准备好后触发代码生成。",
        "请用中文回答。",
    ]

    # 当前代码信息
    if context.get("latest_code"):
        code_lines = context["latest_code"].strip().split("\n")
        parts.append(f"\n当前代码共 {len(code_lines)} 行。")

    # 准确率信息
    if context.get("latest_accuracy") is not None:
        acc = context["latest_accuracy"]
        parts.append(f"当前最新准确率: {acc*100:.1f}%")
    if context.get("best_accuracy") is not None:
        parts.append(f"最佳准确率: {context['best_accuracy']*100:.1f}%")

    # 差异信息
    if context.get("latest_diff"):
        diff = context["latest_diff"]
        if isinstance(diff, dict):
            diff_text = json.dumps(diff, ensure_ascii=False, indent=2)[:3000]
        else:
            diff_text = str(diff)[:3000]
        parts.append(f"\n最新差异详情:\n{diff_text}")

    # 规则
    if rules:
        parts.append(f"\n计算规则（参考）:\n{rules[:5000]}")

    # 源数据结构
    src_desc = config.get("source_structure_desc", "")
    if src_desc:
        parts.append(f"\n源数据结构:\n{src_desc[:3000]}")

    return "\n".join(parts)


def _build_chat_messages(context: Dict, current_message: str) -> list:
    """构建对话消息列表（包含最近对话历史）"""
    messages = []
    for m in context.get("recent_messages", []):
        messages.append({"role": m["role"], "content": m["content"]})
    messages.append({"role": "user", "content": current_message})
    return messages


def _persist_iteration_files(tenant_id: str, session_id: int, iteration_num: int,
                              code: str, run_result: Dict) -> Dict[str, str]:
    """将迭代产物保存到持久化目录，返回文件路径字典"""
    try:
        from ..storage.storage_manager import StorageManager
        sm = StorageManager()
        tenant_dir = sm.get_tenant_dir(tenant_id)
        iter_dir = tenant_dir / "training_chat" / str(session_id) / f"iter_{iteration_num}"
        iter_dir.mkdir(parents=True, exist_ok=True)

        paths = {}

        # 保存脚本
        script_path = iter_dir / "script.py"
        script_path.write_text(code, encoding="utf-8")
        paths["script_file"] = str(script_path)

        # 复制生成的 Excel
        output_dir = run_result.get("output_dir", "")
        if output_dir and os.path.isdir(output_dir):
            for fn in os.listdir(output_dir):
                if fn.endswith((".xlsx", ".xls")) and not fn.startswith("~"):
                    src = os.path.join(output_dir, fn)
                    if "diff" in fn.lower() or "_diff" in fn:
                        dst = iter_dir / f"diff_{fn}"
                        shutil.copy2(src, dst)
                        paths["diff_file"] = str(dst)
                    else:
                        dst = iter_dir / fn
                        shutil.copy2(src, dst)
                        paths["output_file"] = str(dst)

        return paths
    except Exception as e:
        logger.warning(f"保存迭代文件失败: {e}")
        return {}



def _run_single_iteration(
    session_id: int,
    code: str,
    tenant_id: str,
    source_dir: str,
    expected_file: str,
    iteration_num: int,
    salary_year: int = None,
    salary_month: int = None,
    monthly_standard_hours: float = None,
    manual_headers: Dict = None,
    file_passwords: Dict = None,
    pre_loaded_source_data: Dict = None,
) -> Dict[str, Any]:
    """执行单轮训练：运行代码 → 对比 → 返回结果（与 TrainingEngine._execute_and_validate 一致）"""
    from ..sandbox.code_sandbox import CodeSandbox
    from ..utils.excel_comparator import compare_excel_files

    sandbox = CodeSandbox()

    # 创建独立的临时目录（与原训练引擎一致）
    temp_dir = tempfile.mkdtemp(prefix="train_chat_")
    temp_dir = str(Path(temp_dir).resolve())
    input_dir = Path(temp_dir) / "input"
    output_dir = Path(temp_dir) / "output"
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    try:
        # 复制源文件到临时输入目录
        source_file_names = []
        for fn in os.listdir(source_dir):
            if fn.endswith((".xlsx", ".xls")) and not fn.startswith("~"):
                shutil.copy(os.path.join(source_dir, fn), input_dir / fn)
                source_file_names.append(fn)

        # 准备执行环境（与原训练引擎一致）
        execution_env = {
            "input_folder": str(input_dir),
            "output_folder": str(output_dir),
            "source_files": source_file_names,
            "manual_headers": manual_headers or {},
            "file_passwords": file_passwords or {},
        }
        if salary_year is not None:
            execution_env["salary_year"] = salary_year
        if salary_month is not None:
            execution_env["salary_month"] = salary_month
        if monthly_standard_hours is not None:
            execution_env["monthly_standard_hours"] = monthly_standard_hours
        # 注入 tenant_id，让沙箱能创建 HistoricalDataProvider
        execution_env["tenant_id"] = tenant_id

        # 注入预加载源数据（后台全量解析完成后的缓存，避免脚本内重复解析）
        if pre_loaded_source_data is not None:
            execution_env["_pre_loaded_source_data"] = pre_loaded_source_data

        # 执行代码
        start_time = time.time()
        exec_result = sandbox.execute_script(code, execution_env)
        execution_time = time.time() - start_time

        if not exec_result.get("success"):
            return {
                "success": False,
                "error": exec_result.get("error", "执行失败"),
                "accuracy": 0,
                "diff_details": None,
                "output_dir": str(output_dir),
                "execution_time": execution_time,
            }

        # 查找输出文件（排除临时文件和对比文件）
        output_files = [
            f for f in output_dir.glob("*.xlsx")
            if not f.name.startswith("~") and "diff" not in f.name.lower()
            and "comparison" not in f.name.lower()
        ]
        if not output_files:
            return {
                "success": False,
                "error": "脚本执行成功但未生成输出文件",
                "accuracy": 0,
                "diff_details": None,
                "output_dir": str(output_dir),
                "execution_time": execution_time,
            }
        result_file = str(output_files[0])

        # 对比
        diff_output = str(output_dir / "_diff.xlsx")
        comparison = compare_excel_files(result_file, expected_file, diff_output)

        total = comparison.get("total_cells", 1)
        matched = comparison.get("matched_cells", 0)
        accuracy = matched / total if total > 0 else 0

        # 格式化差异摘要
        diff_summary = {}
        field_diffs = comparison.get("field_diff_samples", {})
        if field_diffs:
            for col, info in field_diffs.items():
                diff_summary[col] = {
                    "count": info.get("count", 0),
                    "sample": info.get("formula", info.get("sample", "")),
                }

        # 构建详细差异文本（供 generate_correction_code 使用）
        detailed_diff = comparison.get("detailed_text", "")
        if not detailed_diff and diff_summary:
            lines = []
            for col, info in diff_summary.items():
                lines.append(f"列 '{col}': {info['count']}处差异, 示例: {info.get('sample', '')}")
            detailed_diff = "\n".join(lines)

        return {
            "success": True,
            "accuracy": accuracy,
            "total_cells": total,
            "matched_cells": matched,
            "total_differences": comparison.get("total_differences", 0),
            "diff_details": diff_summary,
            "detailed_diff": detailed_diff,
            "diff_file": diff_output if os.path.exists(diff_output) else None,
            "output_dir": str(output_dir),
            "execution_time": execution_time,
        }

    except Exception as e:
        logger.error(f"单轮执行失败: {e}", exc_info=True)
        return {
            "success": False,
            "error": str(e),
            "accuracy": 0,
            "diff_details": None,
            "output_dir": str(output_dir),
        }


# ==================== 会话管理 ====================


@router.get("/sessions")
def list_chat_sessions(
    tenant_id: str = Query(...),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """列出租户的训练会话（含最新准确率）"""
    sessions = (
        db.query(TrainingSession)
        .filter(TrainingSession.tenant_id == tenant_id)
        .order_by(TrainingSession.started_at.desc())
        .limit(50)
        .all()
    )

    result = []
    for s in sessions:
        # 获取最新脚本信息
        script = None
        if s.final_script_id:
            script = db.query(Script).filter_by(id=s.final_script_id).first()

        cfg = s.config or {}
        latest_files = cfg.get("latest_files", {})
        result.append({
            "id": s.id,
            "session_key": s.session_key,
            "mode": s.mode,
            "status": s.status,
            "total_iterations": s.total_iterations or 0,
            "best_accuracy": s.best_accuracy,
            "has_script": s.final_script_id is not None,
            "script_version": script.version if script else None,
            "started_at": s.started_at.isoformat() if s.started_at else None,
            "finished_at": s.finished_at.isoformat() if s.finished_at else None,
            "has_output": bool(latest_files.get("output_file")),
            "has_diff": bool(latest_files.get("diff_file")),
            "has_code": bool(latest_files.get("script_file")),
        })

    return {"sessions": result}


@router.get("/sessions/{session_id}/messages")
def get_session_messages(
    session_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取会话的所有消息"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    messages = (
        db.query(TrainingMessage)
        .filter_by(session_id=session_id)
        .order_by(TrainingMessage.created_at)
        .all()
    )

    # 也返回最新迭代的代码和准确率
    latest_iteration = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .order_by(TrainingIteration.iteration_num.desc())
        .first()
    )

    best_iteration = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .filter(TrainingIteration.accuracy.isnot(None))
        .order_by(TrainingIteration.accuracy.desc())
        .first()
    )

    cfg = session.config or {}
    latest_files = cfg.get("latest_files", {})
    src_dir = cfg.get("source_dir", "")
    exp_file = cfg.get("expected_file", "")

    # 获取原始文件名称列表
    source_file_names = []
    expected_file_name = None
    try:
        if src_dir and os.path.isdir(src_dir):
            source_file_names = [f for f in os.listdir(src_dir) if not f.startswith("~") and os.path.isfile(os.path.join(src_dir, f))]
        if exp_file and os.path.exists(exp_file):
            expected_file_name = os.path.basename(exp_file)
    except Exception:
        pass

    # 获取所有迭代记录（用于补充对话历史中缺失的代码和执行细节）
    iterations = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .order_by(TrainingIteration.iteration_num)
        .all()
    )

    return {
        "session": {
            "id": session.id,
            "tenant_id": session.tenant_id,
            "mode": session.mode,
            "status": session.status,
            "best_accuracy": session.best_accuracy,
            "total_iterations": session.total_iterations or 0,
            "has_script": session.final_script_id is not None,
            "has_source_files": bool(src_dir and os.path.isdir(src_dir)),
            "has_expected_file": bool(exp_file and os.path.exists(exp_file)),
        },
        "messages": [
            {
                "id": m.id,
                "role": m.role,
                "content": m.content,
                "msg_type": m.msg_type,
                "metadata": m.metadata_,
                "created_at": m.created_at.isoformat() if m.created_at else None,
            }
            for m in messages
        ],
        "iterations": [
            {
                "iteration_num": it.iteration_num,
                "status": it.status,
                "accuracy": it.accuracy,
                "generated_code": it.generated_code,
                "ai_response": (it.ai_response or "")[:2000],
                "execution_result": it.execution_result,
                "error_details": it.error_details,
                "created_at": it.started_at.isoformat() if it.started_at else None,
            }
            for it in iterations
        ],
        "current_code": best_iteration.generated_code if best_iteration else None,
        "current_accuracy": best_iteration.accuracy if best_iteration else None,
        "latest_files": {
            "script_file": bool(latest_files.get("script_file")),
            "output_file": bool(latest_files.get("output_file")),
            "diff_file": bool(latest_files.get("diff_file")),
        },
        "source_file_names": source_file_names,
        "expected_file_name": expected_file_name,
        "has_rules": bool(cfg.get("rules_content")),
    }


# ==================== 开始训练 (首轮) ====================


@router.post("/start")
async def start_training(
    tenant_id: str = Form(...),
    source_files: List[UploadFile] = File(...),
    expected_result: UploadFile = File(None),
    target_file: UploadFile = File(None),
    rule_files: List[UploadFile] = File(default=[]),
    ai_provider: str = Form("deepseek"),
    mode: str = Form("formula"),
    salary_year_month: Optional[str] = Form(None),
    monthly_standard_hours: Optional[float] = Form(None),
    manual_headers: Optional[str] = Form(None),
    force_retrain: bool = Form(False),
    session_id: Optional[int] = Form(None),  # 传入已有 session_id 则继续
    current_user=Depends(get_current_user),
):
    """开始训练（首轮），返回 SSE 流"""

    # 保存上传文件到临时目录
    work_dir = tempfile.mkdtemp(prefix=f"train_{tenant_id}_")
    source_dir = os.path.join(work_dir, "source")
    os.makedirs(source_dir)

    for f in source_files:
        content = await f.read()
        with open(os.path.join(source_dir, f.filename), "wb") as fp:
            fp.write(content)

    expected_file = None
    ef = expected_result or target_file
    if ef:
        content = await ef.read()
        expected_file = os.path.join(work_dir, ef.filename)
        with open(expected_file, "wb") as fp:
            fp.write(content)

    # 保存规则文件到磁盘，然后用 document_parser 解析（支持 docx/xlsx/pdf 等格式）
    rules_content = ""
    saved_rule_paths = []
    rules_dir = os.path.join(work_dir, "rules")
    os.makedirs(rules_dir, exist_ok=True)
    for rf in rule_files:
        try:
            content = await rf.read()
            rule_path = os.path.join(rules_dir, rf.filename)
            with open(rule_path, "wb") as fp:
                fp.write(content)
            saved_rule_paths.append(rule_path)
        except Exception:
            pass

    if saved_rule_paths:
        try:
            from ..ai_engine.document_parser import get_document_parser
            doc_parser = get_document_parser()
            for rp in saved_rule_paths:
                parsed = doc_parser.parse_document(rp)
                rules_content += f"=== 规则文件: {os.path.basename(rp)} ===\n{parsed}\n"
        except Exception as e:
            logger.warning(f"规则文件解析失败，回退到文本读取: {e}")
            for rp in saved_rule_paths:
                try:
                    with open(rp, "r", encoding="utf-8", errors="replace") as f:
                        rules_content += f.read() + "\n"
                except Exception:
                    pass

    # 解析薪资年月
    salary_year, salary_month = None, None
    if salary_year_month:
        try:
            parts = salary_year_month.replace("/", "-").split("-")
            salary_year = int(parts[0])
            salary_month = int(parts[1]) if len(parts) > 1 else None
        except Exception:
            pass

    # 解析手动表头
    manual_headers_dict = None
    if manual_headers:
        try:
            manual_headers_dict = json.loads(manual_headers)
        except Exception:
            logger.warning(f"手动表头 JSON 解析失败: {manual_headers}")

    queue = asyncio.Queue()
    loop = asyncio.get_event_loop()

    async def event_generator():
        while True:
            event = await queue.get()
            if event is None:
                break
            yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"

    def _emit(event):
        """线程安全地往 queue 推事件"""
        loop.call_soon_threadsafe(queue.put_nowait, event)

    def _run_first_iteration():
        """在线程中执行首轮训练（使用 FormulaCodeGenerator）"""
        db = SessionLocal()
        try:
            from ..api.training_persistence import TrainingPersistence
            persistence = TrainingPersistence(db)

            # 创建或获取 session
            if session_id:
                ts = persistence.get_session(session_id)
                if not ts:
                    _emit({"type": "error", "message": "会话不存在"})
                    return
                ts.status = "running"
                db.commit()
            else:
                session_key = f"{tenant_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

                # 确保租户目录存在（使新租户出现在租户列表中）
                try:
                    from ..storage.storage_manager import StorageManager
                    sm = StorageManager()
                    sm.get_tenant_dir(tenant_id)
                except Exception:
                    pass

                # 分析预期文件结构（与 TrainingEngine 一致）
                expected_struct_dict = {}
                try:
                    if expected_file:
                        _emit({"type": "status", "message": "正在分析文件结构..."})
                        expected_struct_dict = _analyze_expected_structure(expected_file)
                except Exception as e:
                    logger.warning(f"预期文件结构分析失败: {e}")

                config = {
                    "source_dir": source_dir,
                    "expected_file": expected_file,
                    "work_dir": work_dir,
                    "rules_content": rules_content[:10000],
                    "expected_structure": expected_struct_dict,
                    "ai_provider": ai_provider,
                    "salary_year": salary_year,
                    "salary_month": salary_month,
                    "monthly_standard_hours": monthly_standard_hours,
                    "manual_headers": manual_headers_dict,
                }
                ts = persistence.create_session(
                    tenant_id=tenant_id,
                    session_key=session_key,
                    mode=mode,
                    user_id=current_user.id,
                    config=config,
                )
                # 填充训练元数据到正式列
                ts.ai_provider = ai_provider
                ts.salary_year = salary_year
                ts.salary_month = salary_month
                ts.manual_headers = manual_headers_dict
                ts.rules_content = rules_content[:10000] if rules_content else None
                ts.expected_structure = expected_struct_dict or None
                db.commit()

            sid = ts.id

            # 发送 session 创建事件
            _emit({"type": "session_created", "session_id": sid})

            # 持久化训练文件到租户目录（防止临时文件被清理后无法继续训练）
            try:
                from ..storage.storage_manager import StorageManager
                _sm = StorageManager()
                session_persist_dir = Path(_sm.get_tenant_dir(tenant_id)) / "training_chat" / str(sid)
                session_persist_dir.mkdir(parents=True, exist_ok=True)

                # 复制源文件
                p_source = session_persist_dir / "source"
                p_source.mkdir(exist_ok=True)
                for fn in os.listdir(source_dir):
                    fp = os.path.join(source_dir, fn)
                    if os.path.isfile(fp):
                        shutil.copy2(fp, str(p_source / fn))

                # 复制预期文件
                if expected_file and os.path.exists(expected_file):
                    p_expected = str(session_persist_dir / Path(expected_file).name)
                    shutil.copy2(expected_file, p_expected)
                else:
                    p_expected = expected_file

                # 保存规则文本
                if rules_content:
                    (session_persist_dir / "rules.txt").write_text(rules_content, encoding="utf-8")

                # 更新 session config 为持久化路径
                # 必须创建新 dict，否则 SQLAlchemy JSON 列不检测 in-place 变异
                _cfg = dict(ts.config) if ts.config else {}
                _cfg["source_dir"] = str(p_source)
                if p_expected:
                    _cfg["expected_file"] = p_expected
                ts.config = _cfg
                from sqlalchemy.orm.attributes import flag_modified
                flag_modified(ts, "config")
                db.commit()
                logger.info(f"训练文件已持久化到: {session_persist_dir}")
            except Exception as e:
                logger.warning(f"持久化训练文件失败: {e}")

            # 添加系统消息
            _add_message(db, sid, "system", "训练开始，正在生成代码...", "status",
                         {"ai_provider": ai_provider, "mode": mode})

            config = dict(ts.config) if ts.config else {}

            # ========== 直接导入模式：固定脚本，跳过 AI，直接 100% ==========
            if mode == "direct":
                _emit({"type": "status", "message": "直接导入模式 — 源文件即输出，无需AI生成..."})

                passthrough_code = '''"""直接导入模式 - 源文件即输出文件，无需数据转换"""
import shutil
import os
import glob

def main(source_dir, output_dir, **kwargs):
    """直接复制源文件到输出目录"""
    os.makedirs(output_dir, exist_ok=True)
    files = glob.glob(os.path.join(source_dir, "*.xls*"))
    for src in files:
        dest = os.path.join(output_dir, os.path.basename(src))
        shutil.copy2(src, dest)
        print(f"直接导入: {os.path.basename(src)}")
    return {"success": True, "files": [os.path.basename(f) for f in files]}
'''
                src_dir = config.get("source_dir", source_dir)
                # 列出源文件名
                try:
                    src_files = [f for f in os.listdir(src_dir) if f.lower().endswith(('.xls', '.xlsx'))]
                except Exception:
                    src_files = []

                _add_message(db, sid, "assistant",
                             f"直接导入模式：源文件直接作为输出（{len(src_files)} 个文件）", "code",
                             {"has_code": True, "direct_import": True})

                iteration_num = 1
                accuracy = 1.0

                persistence.record_iteration(
                    session_id=sid,
                    iteration_num=iteration_num,
                    prompt_text="[direct_import]",
                    ai_response="",
                    generated_code=passthrough_code,
                    accuracy=accuracy,
                    execution_result={"success": True, "direct_import": True,
                                      "files": src_files},
                    error_details=None,
                    status="completed",
                )
                persistence.update_session_best(sid, accuracy, iteration_num)

                # 保存脚本到 storage
                try:
                    from ..storage.storage_manager import StorageManager
                    _sm = StorageManager()
                    _sm.save_script(
                        tenant_id, passthrough_code,
                        {"success": True, "best_score": 1.0, "total_iterations": 1,
                         "best_code": passthrough_code, "mode": "direct"},
                        {}
                    )
                except Exception as e:
                    logger.warning(f"直接导入保存脚本失败: {e}")

                # 保存脚本到 DB（正式列）
                try:
                    persistence.save_script(
                        tenant_id=tenant_id,
                        name=f"script_{tenant_id}",
                        code=passthrough_code,
                        mode="direct",
                        source_session_id=sid,
                        accuracy=1.0,
                        created_by=current_user.id if current_user else None,
                        manual_headers=config.get("manual_headers"),
                    )
                except Exception as e:
                    logger.warning(f"直接导入 DB save_script 失败: {e}")

                # 持久化脚本文件
                saved_files = {}
                try:
                    from ..storage.storage_manager import StorageManager
                    _sm2 = StorageManager()
                    persist_dir = Path(_sm2.get_tenant_dir(tenant_id)) / "training_chat" / str(sid)
                    persist_dir.mkdir(parents=True, exist_ok=True)
                    script_path = persist_dir / f"iter_{iteration_num}_script.py"
                    script_path.write_text(passthrough_code, encoding="utf-8")
                    saved_files["script_file"] = str(script_path)
                except Exception as e:
                    logger.warning(f"直接导入持久化失败: {e}")

                if saved_files:
                    config["latest_files"] = saved_files
                    ts.config = config
                    flag_modified(ts, "config")
                    db.commit()

                _emit({
                    "type": "iteration_complete",
                    "iteration": iteration_num,
                    "accuracy": 1.0,
                    "success": True,
                    "diff_details": None,
                    "files": saved_files,
                })
                _add_message(db, sid, "system",
                             f"直接导入完成，共 {len(src_files)} 个文件，准确率 100%", "status",
                             {"iteration": iteration_num, "accuracy": 1.0})

                ts.status = "completed"
                db.commit()

                _emit({"type": "done"})
                return
            # ========== 结束直接导入模式 ==========

            # 创建 FormulaCodeGenerator（复用原有训练管线）
            def stream_cb(msg):
                _emit({"type": "log", "message": msg})

            _emit({"type": "status", "message": "正在调用 AI 生成代码..."})

            generator, provider = _create_formula_generator(ai_provider, stream_callback=stream_cb)

            # 获取 expected_structure
            expected_struct = config.get("expected_structure", {})
            if not expected_struct and expected_file:
                try:
                    expected_struct = _analyze_expected_structure(expected_file)
                except Exception:
                    expected_struct = {}

            rules = config.get("rules_content", "")
            src_dir = config.get("source_dir", source_dir)

            # 【后台全量加载】在 AI 代码生成期间并行加载全量源数据
            # 这样 AI 生成代码时（耗时最长），全量数据同时解析
            _full_data_future = _executor.submit(
                _load_full_source_data, src_dir, config.get("manual_headers")
            )

            # 使用 FormulaCodeGenerator.generate_code()（与原训练引擎一致）
            _emit({"type": "status", "message": "AI 正在生成代码（使用公式模式）..."})

            code, ai_response = generator.generate_code(
                input_folder=src_dir,
                rules_content=rules,
                expected_structure=expected_struct,
                manual_headers=config.get("manual_headers"),
                stream_callback=stream_cb,
            )

            if not code:
                _add_message(db, sid, "system", "AI 未能生成有效代码", "status",
                             {"error": "no_code"})
                _emit({"type": "error", "message": "AI 未能生成有效代码"})
                return

            # 保存代码生成的 assistant 消息
            code_lines = code.strip().split("\n")
            _add_message(db, sid, "assistant",
                         f"已生成代码（{len(code_lines)} 行），正在执行验证...", "code",
                         {"has_code": True})

            # 保存 source_structure_desc 供后续修正使用
            source_structure_desc = ""
            try:
                source_structure_desc = generator.formula_builder.get_source_structure_for_prompt()
                config["source_structure_desc"] = source_structure_desc[:5000]
                ts.config = config
                # 写入正式列
                ts.source_structure = {"desc": source_structure_desc[:5000]}
                db.commit()
            except Exception as e:
                logger.warning(f"获取源数据结构描述失败: {e}")

            _emit({"type": "status", "message": "代码生成完成，正在执行验证..."})

            # 【后台全量加载】等待全量数据就绪（通常 AI 生成代码耗时更长，此时已完成）
            _full_source_data = None
            try:
                _full_source_data = _full_data_future.result(timeout=300)
                if _full_source_data:
                    logger.info(f"[后台全量加载] 完成，共 {len(_full_source_data)} 个sheet")
                    _emit({"type": "log", "message": f"全量源数据加载完成（{len(_full_source_data)} 个sheet）"})
            except Exception as e:
                logger.warning(f"[后台全量加载] 失败，脚本将自行解析: {e}")

            # 执行并验证
            iteration_num = (ts.total_iterations or 0) + 1
            run_result = _run_single_iteration(
                sid, code, tenant_id,
                src_dir,
                config.get("expected_file", expected_file),
                iteration_num,
                salary_year=salary_year,
                salary_month=salary_month,
                monthly_standard_hours=monthly_standard_hours,
                pre_loaded_source_data=_full_source_data,
            )

            # 记录迭代
            accuracy = run_result.get("accuracy", 0)
            persistence.record_iteration(
                session_id=sid,
                iteration_num=iteration_num,
                prompt_text="[FormulaCodeGenerator.generate_code]",
                ai_response=(ai_response or "")[:10000],
                generated_code=code,
                accuracy=accuracy,
                execution_result={"success": run_result.get("success"),
                                  "total_cells": run_result.get("total_cells"),
                                  "matched_cells": run_result.get("matched_cells")},
                error_details=run_result.get("diff_details") if not run_result.get("success") or accuracy < 1.0 else None,
                status="completed" if run_result.get("success") else "failed",
            )

            # 保存详细差异文本到 config（供后续修正使用）
            if run_result.get("detailed_diff"):
                config["latest_detailed_diff"] = run_result["detailed_diff"][:5000]
                ts.config = config
                db.commit()

            # 更新 session
            persistence.update_session_best(sid, accuracy, iteration_num)

            # 保存脚本到 storage（使智算页面可见）
            try:
                from ..storage.storage_manager import StorageManager
                _sm = StorageManager()
                _sm.save_script(
                    tenant_id, code,
                    {"success": run_result.get("success", False),
                     "best_score": accuracy,
                     "total_iterations": iteration_num,
                     "best_code": code, "mode": mode,
                     "manual_headers": config.get("manual_headers")},
                    {}
                )
            except Exception as e:
                logger.warning(f"save_script 失败: {e}")

            # 保存脚本到 DB（正式列）
            try:
                persistence.save_script(
                    tenant_id=tenant_id,
                    name=f"script_{tenant_id}",
                    code=code,
                    mode=mode,
                    source_session_id=sid,
                    accuracy=accuracy,
                    created_by=current_user.id if current_user else None,
                    config={"manual_headers": config.get("manual_headers"),
                            "source_structure": config.get("source_structure_desc", ""),
                            "rules_content": config.get("rules_content", "")},
                    manual_headers=config.get("manual_headers"),
                    source_structure=ts.source_structure,
                    rules_content=config.get("rules_content", ""),
                    expected_structure=config.get("expected_structure"),
                )
            except Exception as e:
                logger.warning(f"DB save_script 失败: {e}")

            # 持久化训练产物（脚本、输出文件、差异文件）
            saved_files = _persist_iteration_files(tenant_id, sid, iteration_num, code, run_result)
            if saved_files:
                config["latest_files"] = saved_files
                ts.config = config
                db.commit()

            # 生成差异描述消息
            if run_result.get("success"):
                acc_pct = f"{accuracy * 100:.1f}%"
                if accuracy >= 1.0:
                    msg_content = f"第 {iteration_num} 轮完成，准确率 {acc_pct}，所有数据匹配！"
                    _add_message(db, sid, "system", msg_content, "status",
                                 {"iteration": iteration_num, "accuracy": accuracy})
                else:
                    diff_text = _format_diff_for_chat(run_result.get("diff_details", {}))
                    msg_content = f"第 {iteration_num} 轮完成，准确率 {acc_pct}\n\n差异详情:\n{diff_text}"
                    _add_message(db, sid, "system", msg_content, "diff",
                                 {"iteration": iteration_num, "accuracy": accuracy,
                                  "diff_details": run_result.get("diff_details")})
            else:
                error = run_result.get("error", "未知错误")
                msg_content = f"第 {iteration_num} 轮执行失败: {error}"
                _add_message(db, sid, "system", msg_content, "status",
                             {"iteration": iteration_num, "error": error})

            # 发送完成事件（含文件路径）
            _emit({
                "type": "iteration_complete",
                "session_id": sid,
                "iteration": iteration_num,
                "accuracy": accuracy,
                "success": run_result.get("success", False),
                "diff_details": run_result.get("diff_details"),
                "error": run_result.get("error"),
                "files": saved_files,
            })

        except Exception as e:
            logger.error(f"首轮训练失败: {e}", exc_info=True)
            _emit({"type": "error", "message": f"训练失败: {str(e)}"})
        finally:
            db.close()
            _emit(None)

    loop.run_in_executor(_executor, _run_first_iteration)

    return StreamingResponse(event_generator(), media_type="text/event-stream")


# ==================== 对话发消息 ====================


@router.post("/sessions/{session_id}/message")
async def send_message(
    session_id: int,
    message: str = Form(...),
    action: str = Form("chat"),  # "chat" = 对话讨论, "generate" = 触发代码修正
    rule_files: List[UploadFile] = File(default=[]),
    current_user=Depends(get_current_user),
):
    """用户发送消息。action=chat 进行对话讨论，action=generate 触发代码修正+执行"""

    # 读取新的规则文件内容（用 document_parser 支持各种格式）
    new_rules = ""
    if rule_files:
        tmp_dir = tempfile.mkdtemp(prefix="chat_rules_")
        for rf in rule_files:
            try:
                content = await rf.read()
                rule_path = os.path.join(tmp_dir, rf.filename)
                with open(rule_path, "wb") as fp:
                    fp.write(content)
                from ..ai_engine.document_parser import get_document_parser
                parsed = get_document_parser().parse_document(rule_path)
                new_rules += f"=== 规则文件: {rf.filename} ===\n{parsed}\n"
            except Exception as e:
                logger.warning(f"规则文件 {rf.filename} 解析失败: {e}")
                try:
                    new_rules += content.decode("utf-8", errors="replace") + "\n"
                except Exception:
                    pass

    queue = asyncio.Queue()
    loop = asyncio.get_event_loop()

    async def event_generator():
        while True:
            event = await queue.get()
            if event is None:
                break
            yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"

    def _emit(event):
        """线程安全地往 queue 推事件"""
        loop.call_soon_threadsafe(queue.put_nowait, event)

    def _run_chat_conversation():
        """纯对话模式：AI 分析/讨论，不触发代码生成"""
        db = SessionLocal()
        try:
            from ..api.training_persistence import TrainingPersistence
            persistence = TrainingPersistence(db)

            session = persistence.get_session(session_id)
            if not session:
                _emit({"type": "error", "message": "会话不存在"})
                return

            # 保存用户消息
            _add_message(db, session_id, "user", message, "chat")

            # 构建上下文
            context = _get_session_context(db, session_id)
            config = dict(session.config) if session.config else {}

            # 合并规则
            rules = config.get("rules_content", "")
            if new_rules:
                rules = new_rules + "\n" + rules
                config["rules_content"] = rules[:10000]
                session.config = config
                db.commit()

            # 构建 AI 对话消息
            ai_provider_name = config.get("ai_provider", "deepseek")
            from ..ai_engine.ai_provider import AIProviderFactory
            provider = AIProviderFactory.create_provider(ai_provider_name)

            system_prompt = _build_chat_system_prompt(context, config, rules)
            chat_messages = _build_chat_messages(context, message)

            _emit({"type": "status", "message": "AI 正在分析..."})

            # 流式对话
            full_response = ""

            def chunk_cb(chunk):
                nonlocal full_response
                full_response += chunk
                _emit({"type": "chat_chunk", "content": chunk})

            try:
                result = provider.chat_stream(
                    [{"role": "system", "content": system_prompt}] + chat_messages,
                    chunk_callback=chunk_cb,
                )
                if not full_response:
                    full_response = result or ""
            except Exception as e:
                logger.error(f"AI 对话失败: {e}", exc_info=True)
                full_response = f"AI 对话出错: {str(e)}"
                _emit({"type": "chat_chunk", "content": full_response})

            # 保存 AI 回复
            _add_message(db, session_id, "assistant", full_response, "chat")
            _emit({"type": "chat_done", "content": full_response})

        except Exception as e:
            logger.error(f"对话失败: {e}", exc_info=True)
            _emit({"type": "error", "message": f"对话失败: {str(e)}"})
        finally:
            db.close()
            _emit(None)

    def _run_chat_iteration():
        db = SessionLocal()
        try:
            from ..api.training_persistence import TrainingPersistence
            persistence = TrainingPersistence(db)

            session = persistence.get_session(session_id)
            if not session:
                _emit({"type": "error", "message": "会话不存在"})
                return

            # 保存用户消息
            _add_message(db, session_id, "user", message, "chat")

            # 构建上下文
            context = _get_session_context(db, session_id)
            config = dict(session.config) if session.config else {}
            # 合并规则
            rules = config.get("rules_content", "")
            if new_rules:
                rules = new_rules + "\n" + rules
                config["rules_content"] = rules[:10000]
                session.config = config
                db.commit()

            _emit({"type": "status", "message": "正在根据反馈修正代码..."})

            # 获取最新代码（优先最新一轮，而非最佳准确率的一轮）
            original_code = context.get("latest_code") or context.get("best_code")
            if not original_code:
                _emit({"type": "error", "message": "没有可修正的代码，请先运行首轮训练"})
                return

            # 获取差异文本（优先从最新迭代获取，再回退到 config 缓存）
            comparison_result = ""
            diff = context.get("latest_diff")
            if diff:
                if isinstance(diff, dict):
                    comparison_result = json.dumps(diff, ensure_ascii=False, indent=2)
                else:
                    comparison_result = str(diff)
            if not comparison_result:
                comparison_result = config.get("latest_detailed_diff", "")

            # 将用户消息追加到差异说明中（用户的修正指导）
            comparison_result += f"\n\n用户反馈:\n{message}"

            # 获取源数据结构描述
            source_structure_desc = config.get("source_structure_desc", "")

            # 创建 FormulaCodeGenerator 并修正代码
            ai_provider_name = config.get("ai_provider", "deepseek")

            def stream_cb(msg):
                _emit({"type": "log", "message": msg})

            generator, provider = _create_formula_generator(ai_provider_name, stream_callback=stream_cb)

            # 【后台全量加载】修正轮次也需要全量数据
            src_dir = config.get("source_dir", "")
            _full_data_future = _executor.submit(
                _load_full_source_data, src_dir, config.get("manual_headers")
            ) if src_dir and os.path.isdir(src_dir) else None

            _emit({"type": "status", "message": "AI 正在修正代码..."})

            # 使用 FormulaCodeGenerator.generate_correction_code()（与原训练引擎修正逻辑一致）
            code = generator.generate_correction_code(
                original_code=original_code,
                comparison_result=comparison_result,
                rules_content=rules,
                source_structure=source_structure_desc,
                stream_callback=stream_cb,
            )

            if not code:
                ai_msg = "抱歉，我未能生成有效的修正代码。请提供更具体的指导。"
                _add_message(db, session_id, "assistant", ai_msg, "chat")
                _emit({"type": "assistant_message", "content": ai_msg})
                return

            # 保存 AI 回复
            _add_message(db, session_id, "assistant",
                         "已根据反馈修正代码，正在执行验证...", "code",
                         {"has_code": True})

            _emit({"type": "status", "message": "代码已修正，正在执行验证..."})

            # 检查训练文件是否存在（临时文件可能已被清理）
            # 注意：src_dir 已在上方定义（后台全量加载时使用）
            exp_file = config.get("expected_file", "")
            if not src_dir or not os.path.isdir(src_dir):
                _emit({"type": "error", "message": "训练源文件已丢失，请创建新会话并重新上传文件后再训练"})
                _add_message(db, session_id, "system", "训练源文件已丢失，无法继续训练。请新建会话并重新上传文件。", "status", {"error": "files_missing"})
                return
            if not exp_file or not os.path.exists(exp_file):
                _emit({"type": "error", "message": "预期结果文件已丢失，请创建新会话并重新上传文件后再训练"})
                _add_message(db, session_id, "system", "预期结果文件已丢失，无法继续训练。请新建会话并重新上传文件。", "status", {"error": "files_missing"})
                return

            # 【后台全量加载】等待全量数据就绪
            _full_source_data = None
            if _full_data_future:
                try:
                    _full_source_data = _full_data_future.result(timeout=300)
                except Exception as e:
                    logger.warning(f"[后台全量加载] 修正轮次失败: {e}")

            # 执行并验证
            iteration_num = (session.total_iterations or 0) + 1
            run_result = _run_single_iteration(
                session_id, code, session.tenant_id,
                config.get("source_dir", ""),
                config.get("expected_file", ""),
                iteration_num,
                salary_year=config.get("salary_year"),
                salary_month=config.get("salary_month"),
                monthly_standard_hours=config.get("monthly_standard_hours"),
                pre_loaded_source_data=_full_source_data,
            )

            # 保存详细差异文本到 config
            if run_result.get("detailed_diff"):
                config["latest_detailed_diff"] = run_result["detailed_diff"][:5000]
                session.config = config
                db.commit()

            # 准确率倒退检查
            accuracy = run_result.get("accuracy", 0)
            prev_best = context.get("best_accuracy") or 0
            rollback = False

            if accuracy < prev_best and prev_best > 0:
                rollback = True
                rollback_msg = (
                    f"本轮修改导致准确率从 {prev_best*100:.1f}% 下降到 {accuracy*100:.1f}%，"
                    f"已自动回滚到之前的最佳代码。请调整修改方向。"
                )
                _add_message(db, session_id, "system", rollback_msg, "status",
                             {"rollback": True, "old_accuracy": prev_best, "new_accuracy": accuracy})

                persistence.record_iteration(
                    session_id=session_id,
                    iteration_num=iteration_num,
                    generated_code=code,
                    accuracy=accuracy,
                    execution_result={"rollback": True},
                    status="rolled_back",
                )
                session.total_iterations = iteration_num
                db.commit()

                _emit({
                    "type": "iteration_complete",
                    "session_id": session_id,
                    "iteration": iteration_num,
                    "accuracy": prev_best,
                    "rollback": True,
                    "attempted_accuracy": accuracy,
                })
            else:
                persistence.record_iteration(
                    session_id=session_id,
                    iteration_num=iteration_num,
                    prompt_text="[FormulaCodeGenerator.generate_correction_code]",
                    ai_response="",
                    generated_code=code,
                    accuracy=accuracy,
                    execution_result={"success": run_result.get("success"),
                                      "total_cells": run_result.get("total_cells"),
                                      "matched_cells": run_result.get("matched_cells")},
                    error_details=run_result.get("diff_details"),
                    status="completed" if run_result.get("success") else "failed",
                )
                persistence.update_session_best(session_id, accuracy, iteration_num)

                # 保存脚本到 storage（使智算页面可见）
                try:
                    from ..storage.storage_manager import StorageManager
                    _sm = StorageManager()
                    _sm.save_script(
                        session.tenant_id, code,
                        {"success": run_result.get("success", False),
                         "best_score": accuracy,
                         "total_iterations": iteration_num,
                         "best_code": code, "mode": session.mode or "formula",
                         "manual_headers": config.get("manual_headers")},
                        {}
                    )
                except Exception as e:
                    logger.warning(f"save_script 失败: {e}")

                # 保存脚本到 DB（正式列）
                try:
                    persistence.save_script(
                        tenant_id=session.tenant_id,
                        name=f"script_{session.tenant_id}",
                        code=code,
                        mode=session.mode or "formula",
                        source_session_id=session_id,
                        accuracy=accuracy,
                        created_by=current_user.id if current_user else None,
                        config={"manual_headers": config.get("manual_headers"),
                                "source_structure": config.get("source_structure_desc", ""),
                                "rules_content": config.get("rules_content", "")},
                        manual_headers=config.get("manual_headers"),
                        source_structure=session.source_structure,
                        rules_content=config.get("rules_content", ""),
                        expected_structure=config.get("expected_structure"),
                    )
                except Exception as e:
                    logger.warning(f"DB save_script 失败: {e}")

                # 持久化训练产物
                saved_files = _persist_iteration_files(session.tenant_id, session_id, iteration_num, code, run_result)
                if saved_files:
                    config["latest_files"] = saved_files
                    session.config = config
                    db.commit()

                if run_result.get("success"):
                    acc_pct = f"{accuracy * 100:.1f}%"
                    if accuracy >= 1.0:
                        diff_msg = f"第 {iteration_num} 轮完成，准确率 {acc_pct}，所有数据匹配！"
                    else:
                        diff_text = _format_diff_for_chat(run_result.get("diff_details", {}))
                        diff_msg = f"第 {iteration_num} 轮完成，准确率 {acc_pct}\n\n差异详情:\n{diff_text}"
                    _add_message(db, session_id, "system", diff_msg, "diff",
                                 {"iteration": iteration_num, "accuracy": accuracy,
                                  "diff_details": run_result.get("diff_details")})
                else:
                    error = run_result.get("error", "未知错误")
                    diff_msg = f"第 {iteration_num} 轮执行失败: {error}"
                    _add_message(db, session_id, "system", diff_msg, "status",
                                 {"iteration": iteration_num, "error": error})

                _emit({
                    "type": "iteration_complete",
                    "session_id": session_id,
                    "iteration": iteration_num,
                    "accuracy": accuracy,
                    "success": run_result.get("success", False),
                    "diff_details": run_result.get("diff_details"),
                    "error": run_result.get("error"),
                    "files": saved_files,
                })

        except Exception as e:
            logger.error(f"对话迭代失败: {e}", exc_info=True)
            _emit({"type": "error", "message": f"处理失败: {str(e)}"})
        finally:
            db.close()
            _emit(None)

    if action == "generate":
        loop.run_in_executor(_executor, _run_chat_iteration)
    else:
        loop.run_in_executor(_executor, _run_chat_conversation)

    return StreamingResponse(event_generator(), media_type="text/event-stream")


# ==================== 设为最佳 / 上传代码 ====================


class SetBestRequest(BaseModel):
    iteration_id: Optional[int] = None


@router.post("/sessions/{session_id}/set-best")
def set_as_best(
    session_id: int,
    body: SetBestRequest = None,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """将当前最佳代码保存为正式脚本"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    # 找到最佳迭代
    if body and body.iteration_id:
        iteration = db.query(TrainingIteration).filter_by(id=body.iteration_id).first()
    else:
        iteration = (
            db.query(TrainingIteration)
            .filter_by(session_id=session_id)
            .filter(TrainingIteration.accuracy.isnot(None))
            .order_by(TrainingIteration.accuracy.desc(), TrainingIteration.iteration_num.desc())
            .first()
        )

    if not iteration or not iteration.generated_code:
        raise HTTPException(status_code=400, detail="没有可用的代码")

    from ..api.training_persistence import TrainingPersistence
    persistence = TrainingPersistence(db)

    script = persistence.save_script(
        tenant_id=session.tenant_id,
        name=f"script_{session.tenant_id}",
        code=iteration.generated_code,
        mode=session.mode,
        source_session_id=session_id,
        accuracy=iteration.accuracy,
        created_by=current_user.id,
    )

    # 更新 session
    session.final_script_id = script.id
    session.status = "completed"
    session.finished_at = datetime.utcnow()
    db.commit()

    _add_message(db, session_id, "system",
                 f"已设为最佳脚本 (v{script.version}，准确率 {iteration.accuracy*100:.1f}%)",
                 "status", {"script_id": script.id, "version": script.version})

    return {
        "ok": True,
        "script_id": script.id,
        "version": script.version,
        "accuracy": iteration.accuracy,
    }


@router.post("/sessions/{session_id}/upload-code")
async def upload_code(
    session_id: int,
    code: str = Form(None),
    code_file: UploadFile = File(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """直接上传代码，执行并验证"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    # 获取代码
    if code_file:
        code_content = (await code_file.read()).decode("utf-8", errors="replace")
    elif code:
        code_content = code
    else:
        raise HTTPException(status_code=400, detail="请提供代码内容或代码文件")

    config = session.config or {}

    # 执行验证
    iteration_num = (session.total_iterations or 0) + 1
    run_result = _run_single_iteration(
        session_id, code_content, session.tenant_id,
        config.get("source_dir", ""),
        config.get("expected_file", ""),
        iteration_num,
        salary_year=config.get("salary_year"),
        salary_month=config.get("salary_month"),
        monthly_standard_hours=config.get("monthly_standard_hours"),
    )

    from ..api.training_persistence import TrainingPersistence
    persistence = TrainingPersistence(db)

    accuracy = run_result.get("accuracy", 0)
    persistence.record_iteration(
        session_id=session_id,
        iteration_num=iteration_num,
        generated_code=code_content,
        accuracy=accuracy,
        execution_result={"success": run_result.get("success"),
                          "source": "manual_upload"},
        error_details=run_result.get("diff_details"),
        status="completed" if run_result.get("success") else "failed",
    )
    persistence.update_session_best(session_id, accuracy, iteration_num)

    # 消息
    if run_result.get("success"):
        acc_pct = f"{accuracy * 100:.1f}%"
        msg = f"手动上传代码已验证，准确率 {acc_pct}"
        if accuracy < 1.0:
            diff_text = _format_diff_for_chat(run_result.get("diff_details", {}))
            msg += f"\n\n差异详情:\n{diff_text}"
        _add_message(db, session_id, "system", msg, "diff" if accuracy < 1.0 else "status",
                     {"iteration": iteration_num, "accuracy": accuracy, "source": "upload"})
    else:
        msg = f"手动上传代码执行失败: {run_result.get('error', '未知错误')}"
        _add_message(db, session_id, "system", msg, "status",
                     {"iteration": iteration_num, "error": run_result.get("error")})

    return {
        "ok": True,
        "iteration": iteration_num,
        "accuracy": accuracy,
        "success": run_result.get("success", False),
        "diff_details": run_result.get("diff_details"),
        "error": run_result.get("error"),
    }


@router.get("/sessions/{session_id}/code")
def get_current_code(
    session_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取会话当前最佳代码"""
    best = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .filter(TrainingIteration.accuracy.isnot(None))
        .order_by(TrainingIteration.accuracy.desc(), TrainingIteration.iteration_num.desc())
        .first()
    )
    if not best:
        raise HTTPException(status_code=404, detail="暂无代码")

    return {
        "code": best.generated_code,
        "accuracy": best.accuracy,
        "iteration": best.iteration_num,
    }


# ==================== 下载训练产物 ====================


@router.get("/sessions/{session_id}/download/{file_type}")
def download_iteration_file(
    session_id: int,
    file_type: str,  # script / output / diff
    iteration: Optional[int] = Query(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """下载训练产物文件（脚本/生成Excel/差异Excel）"""
    from fastapi.responses import FileResponse

    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    config = session.config or {}

    # 支持按迭代号下载（不指定则使用最新文件）
    files = {}
    if iteration:
        try:
            from ..storage.storage_manager import StorageManager
            sm = StorageManager()
            td = sm.get_tenant_dir(session.tenant_id)
            iter_dir = td / "training_chat" / str(session_id) / f"iter_{iteration}"
            if iter_dir.exists():
                for f in iter_dir.iterdir():
                    if f.name == "script.py":
                        files["script_file"] = str(f)
                    elif "diff" in f.name.lower() and f.suffix in (".xlsx", ".xls"):
                        files["diff_file"] = str(f)
                    elif f.suffix in (".xlsx", ".xls") and not f.name.startswith("~"):
                        files.setdefault("output_file", str(f))
        except Exception:
            pass
    if not files:
        files = config.get("latest_files", {})

    if file_type == "script":
        file_path = files.get("script_file")
        media_type = "text/x-python"
        filename = f"script_{session.tenant_id}_{session.session_key}.py"
    elif file_type == "output":
        file_path = files.get("output_file")
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = f"output_{session.session_key}.xlsx"
    elif file_type == "diff":
        file_path = files.get("diff_file")
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = f"diff_{session.session_key}.xlsx"
    else:
        raise HTTPException(status_code=400, detail=f"未知文件类型: {file_type}")

    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"文件不存在: {file_type}")

    return FileResponse(file_path, media_type=media_type, filename=filename)


# ==================== 重命名版本 ====================


class RenameRequest(BaseModel):
    name: str


@router.post("/sessions/{session_id}/rename")
def rename_session(
    session_id: int,
    body: RenameRequest,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """重命名训练版本"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    new_name = body.name.strip()
    if not new_name:
        raise HTTPException(status_code=400, detail="名称不能为空")

    session.session_key = new_name
    db.commit()

    return {"ok": True, "session_key": new_name}


# ==================== 原始文件下载 ====================


@router.get("/sessions/{session_id}/original-files/{file_category}")
def download_original_file(
    session_id: int,
    file_category: str,  # source / expected / rules
    filename: str = Query(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """下载训练会话的原始文件（源文件/预期文件/规则）"""
    from fastapi.responses import FileResponse

    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    config = session.config or {}

    if file_category == "source":
        src_dir = config.get("source_dir", "")
        if not src_dir or not os.path.isdir(src_dir):
            raise HTTPException(status_code=404, detail="源文件目录不存在")
        if filename:
            file_path = os.path.join(src_dir, os.path.basename(filename))
        else:
            files = [f for f in os.listdir(src_dir) if not f.startswith("~") and os.path.isfile(os.path.join(src_dir, f))]
            if not files:
                raise HTTPException(status_code=404, detail="无源文件")
            file_path = os.path.join(src_dir, files[0])
            filename = files[0]
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="文件不存在")
        return FileResponse(file_path, filename=os.path.basename(file_path))

    elif file_category == "expected":
        file_path = config.get("expected_file", "")
        if not file_path or not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="预期文件不存在")
        return FileResponse(file_path, filename=os.path.basename(file_path))

    elif file_category == "rules":
        # 检查持久化的规则文件
        try:
            from ..storage.storage_manager import StorageManager
            sm = StorageManager()
            td = sm.get_tenant_dir(session.tenant_id)
            rules_file = td / "training_chat" / str(session_id) / "rules.txt"
            if rules_file.exists():
                return FileResponse(str(rules_file), media_type="text/plain", filename="rules.txt")
        except Exception:
            pass
        # 回退：从 config 中生成
        rules = config.get("rules_content", "")
        if not rules:
            raise HTTPException(status_code=404, detail="无规则内容")
        tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8')
        tmp.write(rules)
        tmp.close()
        from fastapi.responses import FileResponse as FR
        return FR(tmp.name, media_type="text/plain", filename="rules.txt")

    elif file_category == "prompt":
        # 生成训练上下文/提示词文件
        lines = [f"训练会话 #{session_id} - 训练上下文\n"]
        lines.append(f"租户: {session.tenant_id}")
        lines.append(f"模式: {session.mode}")
        lines.append(f"AI提供者: {config.get('ai_provider', 'unknown')}")
        lines.append(f"创建时间: {session.started_at}\n")

        rules = config.get("rules_content", "")
        if rules:
            lines.append("=" * 50)
            lines.append("规则内容")
            lines.append("=" * 50)
            lines.append(rules)
            lines.append("")

        src_struct = config.get("source_structure_desc", "")
        if src_struct:
            lines.append("=" * 50)
            lines.append("源数据结构")
            lines.append("=" * 50)
            lines.append(src_struct)
            lines.append("")

        exp_struct = config.get("expected_structure", {})
        if exp_struct:
            lines.append("=" * 50)
            lines.append("预期文件结构")
            lines.append("=" * 50)
            lines.append(json.dumps(exp_struct, ensure_ascii=False, indent=2))
            lines.append("")

        diff = config.get("latest_detailed_diff", "")
        if diff:
            lines.append("=" * 50)
            lines.append("最新差异")
            lines.append("=" * 50)
            lines.append(diff)
            lines.append("")

        tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8')
        tmp.write("\n".join(lines))
        tmp.close()
        from fastapi.responses import FileResponse as FR
        return FR(tmp.name, media_type="text/plain",
                  filename=f"prompt_{session.session_key}.txt")

    else:
        raise HTTPException(status_code=400, detail=f"未知文件类别: {file_category}")


# ==================== 辅助函数 ====================


def _extract_code(ai_response: str) -> Optional[str]:
    """从 AI 响应中提取 Python 代码"""
    if not ai_response:
        return None

    # 尝试提取 ```python ``` 块
    import re
    patterns = [
        r'```python\s*\n(.*?)```',
        r'```\s*\n(.*?)```',
    ]
    for pattern in patterns:
        matches = re.findall(pattern, ai_response, re.DOTALL)
        if matches:
            # 取最长的代码块
            code = max(matches, key=len).strip()
            if "def " in code or "import " in code:
                return code

    # 如果没有代码块标记，但看起来像代码
    if ("def main" in ai_response or "def process" in ai_response) and "import " in ai_response:
        # 去掉明显的非代码行
        lines = ai_response.split("\n")
        code_lines = []
        in_code = False
        for line in lines:
            if line.strip().startswith(("import ", "from ", "def ", "class ", "#")) or in_code:
                in_code = True
                code_lines.append(line)
            elif in_code and (line.strip() == "" or line.startswith(" ") or line.startswith("\t")):
                code_lines.append(line)
        if code_lines:
            return "\n".join(code_lines).strip()

    return None


def _format_diff_for_chat(diff_details: Dict) -> str:
    """将差异详情格式化为可读文本"""
    if not diff_details:
        return "无详细差异信息"

    lines = []
    for col, info in diff_details.items():
        count = info.get("count", 0)
        sample = info.get("sample", "")
        lines.append(f"- {col}: {count}处差异{f' (示例: {sample})' if sample else ''}")

    return "\n".join(lines) if lines else "无详细差异信息"
