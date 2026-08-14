"""智能组表执行 API：提交 / SSE 流 / 状态 / 结果 / 历史 / 结果反馈。

执行在独立子进程（backend.assemble.assemble_worker），事件经 TaskLogBuffer
（key 前缀 "a" 避免与智算任务冲突）推给 SSE。
"""

import os
import sys
import json
import asyncio
import shutil
import logging
import tempfile
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Query
from fastapi.responses import FileResponse, StreamingResponse

from ..auth.dependencies import require_permission, get_operable_tenants
from ..database.connection import SessionLocal
from ..database.models import AssembleTask, AssembleFieldMapping
from ..assemble.assemble_engine import _confirm_mapping, CONFIRM_THRESHOLD

router = APIRouter(prefix="/api/assemble", tags=["智能组表"])

logger = logging.getLogger(__name__)

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_TENANTS_DIR = _PROJECT_ROOT / "tenants"
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
_EVT_PREFIX = "@@EVT@@"
_DONE = "@@DONE@@"


def _buffer_key(task_id: int) -> str:
    """TaskLogBuffer key：加 "a" 前缀避免与智算自增 id 冲突。"""
    return f"a{task_id}"


def _check_task_access(task: AssembleTask, current_user, db):
    """校验当前用户对该组表任务的访问权，防跨租户越权（IDOR）。

    admin → 全部放行；工具租户 __assemble__ 的任务 → 仅创建者本人；
    其余租户任务 → 租户必须在本用户可访问集合内（与 submit 校验一致）。
    """
    if current_user.role and current_user.role.name == "admin":
        return
    from ..auth.dependencies import get_accessible_tenants
    if task.tenant_id == "__assemble__":
        if task.user_id and task.user_id == current_user.id:
            return
        raise HTTPException(status_code=403, detail="无权访问该任务")
    if task.tenant_id in get_accessible_tenants(current_user=current_user, db=db):
        return
    raise HTTPException(status_code=403, detail="无权访问该任务")


# ==================== 提交 ====================

@router.post("/submit")
async def assemble_submit(
    tenant_id: str = Form(""),        # 可空（空租户走通用全局规则）⚠️ multipart 空字符串会被 FastAPI 判缺失，必须给默认值
    rule_id: int = Form(0),
    force_rematch: str = Form("false"),
    ai_provider: str = Form(""),
    file_passwords: str = Form(""),
    source_files: List[UploadFile] = File(...),
    template_file: UploadFile = File(...),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """提交智能组表任务：保存文件到临时目录 → 建任务 → spawn 子进程。"""
    from ..auth.dependencies import get_operable_tenants
    from ..database.connection import SessionLocal as _SL
    from ..database.models import AssembleTask as _AT
    from backend.compute.task_log_buffer import TaskLogBuffer

    # 校验 AI provider 未被系统禁用（模型管理配置）
    from ..database.models import SystemConfig
    _db_cfg = SessionLocal()
    try:
        _cfg_row = _db_cfg.query(SystemConfig).filter_by(key="ai_providers_enabled").first()
    finally:
        _db_cfg.close()
    if ai_provider:
        if _cfg_row and _cfg_row.value and ai_provider not in (_cfg_row.value.get("enabled") or []):
            raise HTTPException(status_code=400, detail=f"AI 模型 {ai_provider} 已被系统禁用")
        if ai_provider not in {"openai", "claude", "deepseek", "ollama", "local"}:
            raise HTTPException(status_code=400, detail=f"不支持的 AI 模型: {ai_provider}")

    tenant_id = (tenant_id or "").strip()
    if not tenant_id:
        # 空租户（默认）：走工具租户 __assemble__（类似 __tools_sop__），不校验授权
        tenant_id = "__assemble__"
    else:
        # 校验租户可访问性
        from ..auth.dependencies import get_accessible_tenants
        _db_check = SessionLocal()
        try:
            accessible = get_accessible_tenants(current_user=current_user, db=_db_check)
        finally:
            _db_check.close()
        if tenant_id not in accessible:
            raise HTTPException(status_code=403, detail="无权访问该租户")

    if not source_files or not template_file:
        raise HTTPException(status_code=400, detail="必须上传源文件和模板文件")
    template_name = (template_file.filename or "").strip()
    if not template_name.lower().endswith(tuple(EXCEL_EXTS)):
        raise HTTPException(status_code=400, detail="模板文件必须是 Excel 文件")

    tmp_dir = Path(tempfile.mkdtemp(prefix="assemble_"))
    try:
        source_dir = tmp_dir / "source"
        source_dir.mkdir(exist_ok=True)
        saved_sources = []
        for f in source_files:
            if not f.filename:
                continue
            name = Path(f.filename).name
            if not name.lower().endswith(tuple(EXCEL_EXTS)):
                raise HTTPException(status_code=400, detail=f"源文件 {name} 不是 Excel 文件")
            (source_dir / name).write_bytes(f.file.read())
            saved_sources.append(name)
        if not saved_sources:
            raise HTTPException(status_code=400, detail="没有有效的源文件")

        tpl_path = tmp_dir / "template" / template_name
        tpl_path.parent.mkdir(exist_ok=True)
        tpl_path.write_bytes(template_file.file.read())
        # .xls 转 .xlsx（复用智算做法）
        if tpl_path.suffix.lower() == ".xls":
            from backend.utils.aspose_helper import convert_format
            xlsx = tpl_path.with_suffix(".xlsx")
            convert_format(str(tpl_path), str(xlsx))
            tpl_path = xlsx
            template_name = tpl_path.name

        # 建任务
        db = _SL()
        try:
            task = _AT(
                tenant_id=tenant_id,
                user_id=current_user.id,
                rule_id=rule_id if rule_id else None,
                status="pending",
                ai_provider=ai_provider or None,
            )
            db.add(task)
            db.commit()
            db.refresh(task)
            task_id = task.id
        finally:
            db.close()

        buffer = TaskLogBuffer.get_instance()
        buffer.create_task(_buffer_key(task_id))
        buffer.push(_buffer_key(task_id),
                    json.dumps({"type": "status", "status": "pending",
                                "message": "任务已提交，文件保存中..."}, ensure_ascii=False))

        # 参数文件
        params = {
            "task_id": task_id,
            "tenant_id": tenant_id,
            "rule_id": rule_id if rule_id else None,
            "source_dir": str(source_dir),
            "template_path": str(tpl_path),
            "force_rematch": force_rematch.lower() in ("true", "1", "on"),
            "file_passwords": json.loads(file_passwords) if file_passwords else {},
            "ai_provider": ai_provider or None,
        }
        params_file = tmp_dir / "params.json"
        params_file.write_text(json.dumps(params, ensure_ascii=False), encoding="utf-8")

        asyncio.create_task(_run_assemble_subprocess(
            _buffer_key(task_id), buffer, str(params_file)))

        return {"task_id": task_id}
    except HTTPException:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        logger.error(f"[assemble/submit] 提交失败: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


async def _run_assemble_subprocess(buffer_key: str, buffer, params_file: str):
    """独立子进程运行引擎；线程读 stdout 事件转推 TaskLogBuffer。

    Windows dev 模式(reload=True)是 SelectorEventLoop，不能用 asyncio 子进程
    （抛 NotImplementedError），用同步 Popen + 后台线程（与智算同一套方案）。
    """
    import subprocess as _subprocess
    import threading as _threading

    loop = asyncio.get_running_loop()
    _proot = str(_PROJECT_ROOT)
    done = asyncio.Event()
    state = {"returncode": None, "start_error": None}

    def _reader():
        proc = None
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"
            proc = _subprocess.Popen(
                [sys.executable, "-u", "-m", "backend.assemble.assemble_worker", params_file],
                cwd=_proot, stdout=_subprocess.PIPE, stderr=_subprocess.STDOUT, env=env,
            )
            for raw in proc.stdout:
                line = raw.decode("utf-8", "replace").rstrip("\r\n")
                if line.startswith(_EVT_PREFIX):
                    ev = line[len(_EVT_PREFIX):]
                    loop.call_soon_threadsafe(buffer.push, buffer_key, ev)
                elif line.startswith(_DONE):
                    pass
                elif line.strip():
                    logger.debug(f"[assemble/subproc] {line}")
            proc.wait()
            state["returncode"] = proc.returncode
        except Exception as e:
            state["start_error"] = e
        finally:
            loop.call_soon_threadsafe(done.set)

    _threading.Thread(target=_reader, name=f"assemble-subproc-{buffer_key}",
                      daemon=True).start()
    await done.wait()

    if state["start_error"] is not None:
        e = state["start_error"]
        logger.error(f"[assemble/subproc] 启动失败: {e}", exc_info=e)
        try:
            buffer.push(buffer_key, json.dumps(
                {"type": "error", "message": f"组表进程启动失败: {e}"}, ensure_ascii=False))
        except Exception:
            pass
        _mark_failed(buffer_key, str(e))
    elif state["returncode"] not in (0, None):
        _mark_failed(buffer_key, f"组表子进程异常退出(code={state['returncode']})")

    # 清理临时工作目录（源文件/模板/exec 输出都在其中；结果已复制到 tenants 结果目录，
    # 这里无论成功/失败/启动异常都删，防止每个任务永久泄漏一个 temp 目录）
    try:
        _tmp = Path(params_file).parent
        if _tmp.name.startswith("assemble_") and _tmp.exists():
            shutil.rmtree(_tmp, ignore_errors=True)
            logger.info(f"[assemble/subproc] 已清理临时目录: {_tmp}")
    except Exception:
        pass

    # 标记缓冲区完成（与 compute 路径 main.py finally 一致）：否则 SSE 生成器永不结束、
    # TaskLogBuffer 该任务的所有事件永久驻留内存（cleanup_expired 只清理 finish 过的任务）
    try:
        buffer.finish(buffer_key)
    except Exception:
        pass


def _mark_failed(buffer_key: str, msg: str):
    """子进程异常退出时兜底标记任务失败。"""
    try:
        from backend.compute.task_log_buffer import TaskLogBuffer
        TaskLogBuffer.get_instance().push(
            buffer_key, json.dumps({"type": "error", "message": msg}, ensure_ascii=False))
    except Exception:
        pass
    try:
        db = SessionLocal()
        task = db.query(AssembleTask).filter_by(id=int(buffer_key[1:])).first()
        if task and task.status not in ("completed", "error"):
            task.status = "error"
            task.error = msg[:2000]
            db.commit()
    except Exception:
        pass
    finally:
        try:
            db.close()
        except Exception:
            pass


# ==================== SSE 流 / 状态 ====================

@router.get("/tasks/{task_id}/stream")
async def assemble_task_stream(
    task_id: int,
    last_event_id: int = 0,
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """SSE 流：从 last_event_id 续读，支持断线重连。

    已加鉴权（前端改为 fetch+ReadableStream 携带 Bearer token 消费），并校验任务归属。
    """
    from backend.compute.task_log_buffer import TaskLogBuffer
    buffer = TaskLogBuffer.get_instance()
    key = _buffer_key(task_id)

    # 鉴权 + 归属校验（防未登录/跨租户按 task_id 探测他人任务）
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail=f"任务不存在: {task_id}")
        _check_task_access(task, current_user, db)
    finally:
        db.close()

    status = buffer.get_status(key)
    if status is None:
        # buffer 已过期 → 查 DB 返回最终状态
        db = SessionLocal()
        try:
            task = db.query(AssembleTask).filter_by(id=task_id).first()
        finally:
            db.close()
        if task and task.status in ("completed", "error"):
            if task.status == "completed":
                data = json.dumps({"type": "complete", "output_files": task.output_files or [],
                                   "matched_from_cache": bool(task.matched_from_cache),
                                   "message": "组表完成"}, ensure_ascii=False)
            else:
                data = json.dumps({"type": "error", "message": task.error or "组表失败"},
                                  ensure_ascii=False)
            async def _final_event():
                yield f"id: 0\ndata: {data}\n\n"
            return StreamingResponse(_final_event(), media_type="text/event-stream",
                                     headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})
        raise HTTPException(status_code=404, detail=f"任务不存在: {task_id}")

    async def _sse_generator():
        async for event_id, data in buffer.read_from(key, from_id=last_event_id):
            if event_id == -1:
                yield ": heartbeat\n\n"
            else:
                yield f"id: {event_id}\ndata: {data}\n\n"

    return StreamingResponse(_sse_generator(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@router.get("/tasks/{task_id}/status")
async def assemble_task_status(
    task_id: int,
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """轮询兜底。已加鉴权 + 归属校验。"""
    from backend.compute.task_log_buffer import TaskLogBuffer
    buffer = TaskLogBuffer.get_instance()
    key = _buffer_key(task_id)

    out = {"task_id": task_id}
    buf_status = buffer.get_status(key)
    out["stream_available"] = buf_status is not None
    if buf_status:
        out["event_count"] = buf_status["event_count"]

    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        _check_task_access(task, current_user, db)
        out.update({
            "status": task.status,
            "matched_from_cache": bool(task.matched_from_cache),
            "output_files": task.output_files or [],
            "error": task.error,
            "created_at": task.created_at.isoformat() if task.created_at else None,
        })
        return out
    finally:
        db.close()


# ==================== 结果 ====================

@router.get("/tasks/{task_id}/result")
async def assemble_task_result(task_id: int,
                               current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage"))):
    """任务结果文件列表。"""
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        _check_task_access(task, current_user, db)
        return {"task_id": task_id, "status": task.status,
                "output_files": task.output_files or [],
                "matched_from_cache": bool(task.matched_from_cache)}
    finally:
        db.close()


@router.get("/download/{task_id}/{file_name}")
async def assemble_download(task_id: int, file_name: str):
    """下载结果文件（原版/纯值版）。

    ⚠️ 无鉴权（与智算 /api/download-compute-result 一致）：前端用 <a> 标签直接跳转下载，
    EventSource/普通链接无法携带 Authorization header，加鉴权会导致下载 401。
    """
    name = Path(file_name).name   # 防目录穿越
    p = _TENANTS_DIR / "placeholder"  # 先占位，下面按任务查租户
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        if task.output_files and name not in task.output_files:
            raise HTTPException(status_code=404, detail="文件不在任务结果中")
        fp = _TENANTS_DIR / task.tenant_id / "assemble_results" / str(task_id) / name
        if not fp.exists():
            raise HTTPException(status_code=404, detail="结果文件不存在")
        return FileResponse(
            fp, filename=name,
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"},
        )
    finally:
        db.close()


# ==================== 历史 ====================

@router.get("/history")
async def assemble_history(
    tenant_id: str = Query(""),
    limit: int = Query(50, ge=1, le=200),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """任务历史列表（按租户过滤）。

    非 admin：历史限制为本人可访问租户 ∪ 自己的工具租户(__assemble__)任务，
    避免空 tenant_id 时返回全部租户数据（越权）。
    """
    db = SessionLocal()
    try:
        from sqlalchemy import and_, or_
        from ..auth.dependencies import get_accessible_tenants

        q = db.query(AssembleTask)
        is_admin = current_user.role and current_user.role.name == "admin"
        if tenant_id:
            if is_admin:
                q = q.filter(AssembleTask.tenant_id == tenant_id)
            elif tenant_id == "__assemble__":
                q = q.filter(and_(AssembleTask.tenant_id == "__assemble__",
                                  AssembleTask.user_id == current_user.id))
            else:
                accessible = set(get_accessible_tenants(current_user=current_user, db=db))
                if tenant_id in accessible:
                    q = q.filter(AssembleTask.tenant_id == tenant_id)
                else:
                    q = q.filter(False)   # 非可访问租户 → 空结果，不泄露任务存在性
        elif not is_admin:
            accessible = set(get_accessible_tenants(current_user=current_user, db=db))
            q = q.filter(or_(
                AssembleTask.tenant_id.in_(accessible),
                and_(AssembleTask.tenant_id == "__assemble__",
                     AssembleTask.user_id == current_user.id),
            ))
        rows = q.order_by(AssembleTask.id.desc()).limit(limit).all()
        items = [{
            "id": t.id,
            "tenant_id": t.tenant_id,
            "rule_id": t.rule_id,
            "status": t.status,
            "matched_from_cache": bool(t.matched_from_cache),
            "output_files": t.output_files or [],
            "error": t.error,
            "created_at": t.created_at.isoformat() if t.created_at else None,
            "completed_at": t.completed_at.isoformat() if t.completed_at else None,
        } for t in rows]
        return {"items": items}
    finally:
        db.close()


# ==================== 结果确认反馈（置信度闭环） ====================

def _load_review_samples(src_arc: Path, limit: int = 3) -> List[dict]:
    """读归档源文件前几行样例（不脱敏），供错误弹窗人工核对映射。"""
    from excel_parser import IntelligentExcelParser
    if not src_arc.exists():
        return []
    parser = IntelligentExcelParser()
    out = []
    files = sorted(p for p in src_arc.iterdir()
                   if p.is_file() and p.suffix.lower() in EXCEL_EXTS)
    for fp in files:
        try:
            sheets = parser.parse_excel_file(
                str(fp), max_data_rows=limit, read_formulas=False,
                active_sheet_only=False, best_region_only=True)
        except Exception:
            continue
        for sd in sheets:
            region = (sd.regions or [None])[0]
            if region is None:
                continue
            out.append({
                "file": fp.name,
                "sheet": sd.sheet_name,
                "columns": list((region.head_data or {}).keys()),
                "head_data": region.head_data or {},
                "rows": (region.data or [])[:limit],
            })
    return out


def _list_source_options(src_arc: Path) -> List[dict]:
    """解析归档源文件的全部可见 sheet 列，生成复核弹窗的候选源表清单。

    返回 [{source_sheet: "源_名", columns: [列名...]}]，source_sheet 与骨架追加后的
    实际 sheet 名一致（assign_sheet_keys + build_prefixed_sheet_names 同套命名），
    供弹窗下拉选择其它源表（修正后 field_mapping 只剩选中表，须从全表候选恢复）。
    """
    from excel_parser import IntelligentExcelParser
    from backend.utils.data_helpers import assign_sheet_keys, build_prefixed_sheet_names

    if not src_arc.exists():
        return []
    parser = IntelligentExcelParser()
    files = sorted(p for p in src_arc.iterdir()
                   if p.is_file() and p.suffix.lower() in EXCEL_EXTS)

    pairs: List[tuple] = []
    cols_by_pair: Dict[tuple, List[str]] = {}
    for fp in files:
        try:
            sheets = parser.parse_excel_file(
                str(fp), max_data_rows=1, read_formulas=False,
                active_sheet_only=False, best_region_only=True)
        except Exception:
            continue
        file_base = Path(fp.name).stem
        for sd in sheets:
            region = (sd.regions or [None])[0]
            if region is None:
                continue
            pairs.append((file_base, sd.sheet_name))
            cols_by_pair[(file_base, sd.sheet_name)] = list((region.head_data or {}).keys())

    if not pairs:
        return []
    pairs.sort(key=lambda p: (str(p[0]), str(p[1])))
    sk_map = assign_sheet_keys(pairs, reserved_names=set())
    ordered_sk = sorted(set(sk_map.values()))
    name_map = build_prefixed_sheet_names(ordered_sk, prefix="源_", reserved=set())

    out = []
    for (fb, sn) in pairs:
        out.append({
            "source_sheet": name_map[sk_map[(fb, sn)]],
            "columns": cols_by_pair[(fb, sn)],
        })
    return out


def _batch_confirm(db, task: AssembleTask) -> int:
    """对任务 field_mapping 里的语义列批量确认（同名列跳过）。返回确认的列数。"""
    n = 0
    for tgt, info in (task.field_mapping or {}).items():
        if not isinstance(info, dict):
            continue
        src = (info.get("source_column") or "").strip()
        if _confirm_mapping(db, task.tenant_id, task.template_signature, src, tgt) is not None:
            n += 1
    return n


@router.post("/tasks/{task_id}/feedback")
async def assemble_feedback(
    task_id: int,
    correct: str = Form("true"),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """结果确认：
    ✅ correct=true → 对 field_mapping 里所有语义列批量 confirm+1（达阈值转 active）
    ❌ correct=false → 返回逐列复核数据（映射清单 + 源样例不脱敏），不立即停用任何条目
    """
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        _check_task_access(task, current_user, db)
        if task.feedback_status:
            raise HTTPException(status_code=409, detail="该任务已反馈过，不能重复反馈")

        is_correct = correct.lower() in ("true", "1", "on")
        if is_correct:
            n = _batch_confirm(db, task)
            task.feedback_status = "confirmed"
            db.commit()
            return {"ok": True, "correct": True, "confirmed": n,
                    "message": f"已确认 {n} 列语义映射（累计确认次数 +1）"}

        # 错误：返回复核数据（映射清单 + 全部源表候选 + 上次修正预填），不连坐停用
        src_arc = _TENANTS_DIR / task.tenant_id / "assemble_results" / str(task_id) / "_source"
        return {"ok": True, "correct": False,
                "field_mapping": task.field_mapping or {},
                "corrected_mapping": task.corrected_mapping or {},
                "source_options": _list_source_options(src_arc),
                "samples": _load_review_samples(src_arc)}
    finally:
        db.close()


@router.post("/tasks/{task_id}/confirm-corrected")
async def assemble_confirm_corrected(
    task_id: int,
    corrected_mapping: str = Form(""),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """弹窗「确认」：对最终采用的映射（corrected_mapping，{目标列: 源列}）逐列 confirm+1。"""
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        _check_task_access(task, current_user, db)
        if task.feedback_status == "confirmed":
            raise HTTPException(status_code=409, detail="该任务已确认过，不能重复确认")
        try:
            cm = json.loads(corrected_mapping) if corrected_mapping else {}
        except Exception:
            raise HTTPException(status_code=400, detail="corrected_mapping 必须是 JSON")
        if not isinstance(cm, dict):
            raise HTTPException(status_code=400, detail="corrected_mapping 必须是对象")

        n = 0
        for tgt, src in cm.items():
            _src_col = str(src).strip()
            if "|" in _src_col:          # "源表|源列" → 知识库只存源列名
                _src_col = _src_col.split("|", 1)[1].strip()
            if _confirm_mapping(db, task.tenant_id, task.template_signature, _src_col, tgt) is not None:
                n += 1
        task.corrected_mapping = cm
        task.feedback_status = "confirmed"
        db.commit()
        return {"ok": True, "confirmed": n, "message": f"已确认 {n} 列映射（累计确认次数 +1）"}
    finally:
        db.close()


@router.post("/tasks/{task_id}/rematch")
async def assemble_rematch(
    task_id: int,
    corrected_mapping: str = Form(""),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """用修正映射重新组表：读归档源文件+模板 → 建新任务 → spawn 子进程。"""
    from backend.compute.task_log_buffer import TaskLogBuffer

    try:
        cm = json.loads(corrected_mapping) if corrected_mapping else {}
    except Exception:
        raise HTTPException(status_code=400, detail="corrected_mapping 必须是 JSON")
    if not isinstance(cm, dict):
        raise HTTPException(status_code=400, detail="corrected_mapping 必须是对象")

    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")
        _check_task_access(task, current_user, db)
        if task.status != "completed":
            raise HTTPException(status_code=400, detail="仅已完成的任务可重新组表")

        src_arc = _TENANTS_DIR / task.tenant_id / "assemble_results" / str(task_id) / "_source"
        tpl_arc = _TENANTS_DIR / task.tenant_id / "assemble_results" / str(task_id) / "_template"
        src_files = sorted(p for p in src_arc.iterdir()
                           if p.is_file() and p.suffix.lower() in EXCEL_EXTS) if src_arc.exists() else []
        tpl_files = sorted(p for p in tpl_arc.iterdir()
                           if p.is_file() and p.suffix.lower() in EXCEL_EXTS) if tpl_arc.exists() else []
        if not src_files or not tpl_files:
            raise HTTPException(status_code=404, detail="未找到归档的源文件/模板，无法重新组表")

        # 记录修正映射到原任务（供追溯），标记进入修正流程
        task.corrected_mapping = cm
        task.feedback_status = "corrected"
        db.commit()

        tenant_id = task.tenant_id
        rule_id = task.rule_id
        ai_provider = task.ai_provider
        tpl_path_src = tpl_files[0]
        # 复用原任务代码（匹配关系已定，不重新 AI 生成，直接覆盖 FIELD_MAPPING 执行）
        # 兜底：任务 code_path 为空时（如 rematch 产生的新任务未落库）按签名存档路径找
        prewritten_code = None
        _code_candidates = []
        if task.code_path:
            _code_candidates.append(task.code_path)
        if task.signature:
            _code_candidates.append(str(_TENANTS_DIR / task.tenant_id / "assemble_scripts"
                                         / f"{task.signature}.py"))
        for _cp in _code_candidates:
            if _cp and Path(_cp).exists():
                try:
                    prewritten_code = Path(_cp).read_text(encoding="utf-8")
                    break
                except Exception:
                    continue
        if prewritten_code is None:
            logger.warning(f"[assemble/rematch] 未找到原任务代码，回退 AI 生成: task={task_id}")
    finally:
        db.close()

    # 方向转换：{目标列: 源表|源列} → 两个结构
    #   field_mapping_override: {目标列: {source_sheet, source_column}}（骨架 FIELD_MAPPING 覆盖）
    #   pre_mapped: {源列: 目标列@源表}（供 AI 生成兜底，本次直接执行不需要，保留兼容）
    field_mapping_override = {}
    pre_mapped = {}
    for tgt, src in cm.items():
        if src and tgt:
            _v = str(src).strip()
            if "|" in _v:
                _sheet, _col = _v.split("|", 1)
                field_mapping_override[str(tgt).strip()] = {
                    "source_sheet": _sheet.strip(),
                    "source_column": _col.strip(),
                }
                pre_mapped[_col.strip()] = f"{str(tgt).strip()}@{_sheet.strip()}"
            else:
                field_mapping_override[str(tgt).strip()] = {
                    "source_sheet": "",
                    "source_column": _v,
                }
                pre_mapped[_v] = str(tgt).strip()

    tmp_dir = Path(tempfile.mkdtemp(prefix="assemble_"))
    try:
        source_dir = tmp_dir / "source"
        source_dir.mkdir(exist_ok=True)
        for p in src_files:
            shutil.copy2(p, source_dir / p.name)
        new_tpl = tmp_dir / "template" / tpl_path_src.name
        new_tpl.parent.mkdir(exist_ok=True)
        shutil.copy2(tpl_path_src, new_tpl)

        db = SessionLocal()
        try:
            new_task = AssembleTask(
                tenant_id=tenant_id, user_id=current_user.id,
                rule_id=rule_id, status="pending",
                ai_provider=ai_provider, corrected_mapping=cm,
            )
            db.add(new_task)
            db.commit()
            db.refresh(new_task)
            new_id = new_task.id
        finally:
            db.close()

        buffer = TaskLogBuffer.get_instance()
        buffer.create_task(_buffer_key(new_id))
        buffer.push(_buffer_key(new_id), json.dumps(
            {"type": "status", "status": "pending",
             "message": "重新组表任务已提交（使用人工修正映射）..."}, ensure_ascii=False))

        params = {
            "task_id": new_id,
            "tenant_id": tenant_id,
            "rule_id": rule_id,
            "source_dir": str(source_dir),
            "template_path": str(new_tpl),
            "force_rematch": True,          # 跳过存档/知识库
            "file_passwords": {},
            "ai_provider": ai_provider,
            "pre_mapped": pre_mapped,
            "prewritten_code": prewritten_code,          # 复用原代码直接执行，不走 AI
            "field_mapping_override": field_mapping_override,  # 修正映射覆盖 FIELD_MAPPING
        }
        params_file = tmp_dir / "params.json"
        params_file.write_text(json.dumps(params, ensure_ascii=False), encoding="utf-8")

        asyncio.create_task(_run_assemble_subprocess(_buffer_key(new_id), buffer, str(params_file)))
        return {"task_id": new_id}
    except HTTPException:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        logger.error(f"[assemble/rematch] 重新组表失败: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
