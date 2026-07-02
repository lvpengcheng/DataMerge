"""
训练历史查询 API
"""

from typing import Optional, List
from datetime import datetime

from fastapi import APIRouter, Depends, Query, HTTPException
from sqlalchemy.orm import Session

from ..database.connection import get_db
from ..database.models import TrainingSession, TrainingIteration, Script
from ..auth.dependencies import get_current_user, get_accessible_tenants, get_operable_tenants

router = APIRouter(prefix="/api/training", tags=["训练管理"])


@router.get("/sessions")
def list_training_sessions(
    tenant_id: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible_tenants: list = Depends(get_operable_tenants),
):
    """训练会话列表（按租户权限过滤；tenant_id 支持模糊包含）"""
    # 会话隐藏规则：会话“关联脚本” = 它生成的(source_session_id) ∪ 它绑定的(final_script_id)。
    # 关联脚本全部被停用 → 隐藏（含“生成该脚本的会话”）；任一仍启用或未产出脚本 → 保留。
    # 先算出需隐藏的会话 id 集合，再在 SQL 中排除（保持分页 total 正确）。
    hidden_session_ids = set()
    try:
        _id_active = {}
        _src_any = set()
        _src_active = set()
        for _sid, _src, _act in db.query(
            Script.id, Script.source_session_id, Script.is_active
        ).filter(Script.tenant_id.in_(accessible_tenants)).all():
            _id_active[_sid] = _act
            if _src is not None:
                _src_any.add(_src)
                if _act:
                    _src_active.add(_src)
        for _s_id, _final in db.query(
            TrainingSession.id, TrainingSession.final_script_id
        ).filter(TrainingSession.tenant_id.in_(accessible_tenants)).all():
            has_any = _s_id in _src_any
            has_active = _s_id in _src_active
            if _final is not None and _final in _id_active:
                has_any = True
                if _id_active[_final]:
                    has_active = True
            if has_any and not has_active:
                hidden_session_ids.add(_s_id)
    except Exception:
        hidden_session_ids = set()

    q = db.query(TrainingSession).filter(TrainingSession.tenant_id.in_(accessible_tenants))
    if tenant_id:
        q = q.filter(TrainingSession.tenant_id.like(f"%{tenant_id}%"))
    if status:
        q = q.filter(TrainingSession.status == status)
    if hidden_session_ids:
        q = q.filter(~TrainingSession.id.in_(hidden_session_ids))
    total = q.count()
    sessions = q.order_by(TrainingSession.started_at.desc()).offset(offset).limit(limit).all()

    return {
        "total": total,
        "items": [
            {
                "id": s.id,
                "tenant_id": s.tenant_id,
                "session_key": s.session_key,
                "mode": s.mode,
                "status": s.status,
                "total_iterations": s.total_iterations,
                "best_accuracy": s.best_accuracy,
                "error_message": s.error_message,
                "started_at": s.started_at.isoformat() if s.started_at else None,
                "finished_at": s.finished_at.isoformat() if s.finished_at else None,
            }
            for s in sessions
        ],
    }


@router.get("/sessions/{session_id}")
def get_training_session(
    session_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible_tenants: list = Depends(get_operable_tenants),
):
    """训练会话详情"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="训练会话不存在")
    if session.tenant_id not in accessible_tenants:
        raise HTTPException(status_code=403, detail="无权访问该租户")

    return {
        "id": session.id,
        "tenant_id": session.tenant_id,
        "session_key": session.session_key,
        "mode": session.mode,
        "status": session.status,
        "config": session.config,
        "total_iterations": session.total_iterations,
        "best_accuracy": session.best_accuracy,
        "error_message": session.error_message,
        "started_at": session.started_at.isoformat() if session.started_at else None,
        "finished_at": session.finished_at.isoformat() if session.finished_at else None,
    }


@router.get("/sessions/{session_id}/iterations")
def get_training_iterations(
    session_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible_tenants: list = Depends(get_operable_tenants),
):
    """训练会话的所有迭代记录"""
    session = db.query(TrainingSession).filter_by(id=session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="训练会话不存在")
    if session.tenant_id not in accessible_tenants:
        raise HTTPException(status_code=403, detail="无权访问该租户")

    iterations = (
        db.query(TrainingIteration)
        .filter_by(session_id=session_id)
        .order_by(TrainingIteration.iteration_num)
        .all()
    )

    return [
        {
            "id": it.id,
            "iteration_num": it.iteration_num,
            "status": it.status,
            "accuracy": it.accuracy,
            "generated_code": it.generated_code,
            "execution_result": it.execution_result,
            "error_details": it.error_details,
            "duration_seconds": it.duration_seconds,
            "started_at": it.started_at.isoformat() if it.started_at else None,
            "finished_at": it.finished_at.isoformat() if it.finished_at else None,
        }
        for it in iterations
    ]


@router.get("/scripts")
def list_scripts(
    tenant_id: Optional[str] = Query(None),
    is_active: bool = Query(True),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible_tenants: list = Depends(get_operable_tenants),
):
    """脚本列表（按租户权限过滤）"""
    q = db.query(Script).filter(Script.is_active == is_active)
    if tenant_id:
        if tenant_id not in accessible_tenants:
            raise HTTPException(status_code=403, detail="无权访问该租户")
        q = q.filter(Script.tenant_id == tenant_id)
    else:
        q = q.filter(Script.tenant_id.in_(accessible_tenants))
    scripts = q.order_by(Script.created_at.desc()).all()

    return [
        {
            "id": s.id,
            "tenant_id": s.tenant_id,
            "name": s.name,
            "mode": s.mode,
            "accuracy": s.accuracy,
            "version": s.version,
            "is_active": s.is_active,
            "created_at": s.created_at.isoformat() if s.created_at else None,
        }
        for s in scripts
    ]


@router.get("/scripts/{script_id}")
def get_script(
    script_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible_tenants: list = Depends(get_operable_tenants),
):
    """脚本详情（含代码）"""
    script = db.query(Script).filter_by(id=script_id).first()
    if not script:
        raise HTTPException(status_code=404, detail="脚本不存在")
    if script.tenant_id not in accessible_tenants:
        raise HTTPException(status_code=403, detail="无权访问该租户")

    return {
        "id": script.id,
        "tenant_id": script.tenant_id,
        "name": script.name,
        "code": script.code,
        "mode": script.mode,
        "config": script.config,
        "accuracy": script.accuracy,
        "version": script.version,
        "source_session_id": script.source_session_id,
        "is_active": script.is_active,
        "created_at": script.created_at.isoformat() if script.created_at else None,
    }
