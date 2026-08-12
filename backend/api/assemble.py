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
async def assemble_task_stream(task_id: int, last_event_id: int = 0):
    """SSE 流：从 last_event_id 续读，支持断线重连。"""
    from backend.compute.task_log_buffer import TaskLogBuffer
    buffer = TaskLogBuffer.get_instance()
    key = _buffer_key(task_id)

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
async def assemble_task_status(task_id: int):
    """轮询兜底。"""
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
    """任务历史列表（按租户过滤）。"""
    db = SessionLocal()
    try:
        q = db.query(AssembleTask)
        if tenant_id:
            q = q.filter(AssembleTask.tenant_id == tenant_id)
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


# ==================== 结果确认反馈 ====================

@router.post("/tasks/{task_id}/feedback")
async def assemble_feedback(
    task_id: int,
    correct: str = Form("true"),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """结果确认：✅ correct=true → 命中条目 hit_count+1；
    ❌ correct=false → 该任务命中的知识库条目全部标 review_needed（不再自动采用）。"""
    db = SessionLocal()
    try:
        task = db.query(AssembleTask).filter_by(id=task_id).first()
        if not task:
            raise HTTPException(status_code=404, detail="任务不存在")

        used_ids = task.used_mapping_ids or []
        is_correct = correct.lower() in ("true", "1", "on")
        if is_correct:
            if used_ids:
                for m in db.query(AssembleFieldMapping).filter(AssembleFieldMapping.id.in_(used_ids)).all():
                    m.hit_count = (m.hit_count or 1) + 1
                db.commit()
            return {"ok": True, "correct": True,
                    "message": f"已确认结果正确，知识库 {len(used_ids)} 条映射可信度 +1"}
        else:
            if used_ids:
                for m in db.query(AssembleFieldMapping).filter(AssembleFieldMapping.id.in_(used_ids)).all():
                    m.status = "review_needed"
                db.commit()
            return {"ok": True, "correct": False,
                    "message": f"已停用 {len(used_ids)} 条知识库映射（待复核），"
                               "建议勾选「强制重新匹配」后重新组表"}
    finally:
        db.close()
