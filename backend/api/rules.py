"""
规则整理会话 CRUD API
"""

import logging
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from pydantic import BaseModel

from ..database.connection import get_db
from ..database.models import RuleSession
from ..auth.dependencies import get_current_user

router = APIRouter(prefix="/api/rules", tags=["规则整理"])

logger = logging.getLogger(__name__)


# ==================== Pydantic 模型 ====================

class SessionUpdateRequest(BaseModel):
    title: Optional[str] = None
    status: Optional[str] = None


class SessionListItem(BaseModel):
    id: int
    title: str
    status: str
    ai_provider: str
    source_file_names: Optional[list] = None
    target_file_name: Optional[str] = None
    design_doc_names: Optional[list] = None
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    class Config:
        from_attributes = True


# ==================== 会话列表 ====================

@router.get("/sessions")
def list_sessions(
    current_user=Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """列出当前用户的会话（不含 messages / final_result 大字段）"""
    rows = (
        db.query(RuleSession)
        .filter(RuleSession.user_id == current_user.id)
        .order_by(RuleSession.updated_at.desc())
        .all()
    )
    result = []
    for r in rows:
        result.append({
            "id": r.id,
            "title": r.title,
            "status": r.status,
            "ai_provider": r.ai_provider or "",
            "source_file_names": r.source_file_names,
            "target_file_name": r.target_file_name,
            "design_doc_names": r.design_doc_names,
            "created_at": r.created_at.isoformat() if r.created_at else None,
            "updated_at": r.updated_at.isoformat() if r.updated_at else None,
        })
    return {"sessions": result}


# ==================== 会话详情 ====================

@router.get("/sessions/{session_id}")
def get_session(
    session_id: int,
    current_user=Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """获取会话详情（含完整 messages）"""
    session = (
        db.query(RuleSession)
        .filter(RuleSession.id == session_id, RuleSession.user_id == current_user.id)
        .first()
    )
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    return {
        "id": session.id,
        "title": session.title,
        "status": session.status,
        "ai_provider": session.ai_provider or "",
        "source_file_names": session.source_file_names,
        "target_file_name": session.target_file_name,
        "design_doc_names": session.design_doc_names,
        "messages": session.messages or [],
        "final_result": session.final_result,
        "created_at": session.created_at.isoformat() if session.created_at else None,
        "updated_at": session.updated_at.isoformat() if session.updated_at else None,
    }


# ==================== 更新会话 ====================

@router.put("/sessions/{session_id}")
def update_session(
    session_id: int,
    body: SessionUpdateRequest,
    current_user=Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """更新会话标题/状态"""
    session = (
        db.query(RuleSession)
        .filter(RuleSession.id == session_id, RuleSession.user_id == current_user.id)
        .first()
    )
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    if body.title is not None:
        session.title = body.title
    if body.status is not None:
        session.status = body.status
    session.updated_at = datetime.utcnow()
    db.commit()
    return {"ok": True}


# ==================== 删除会话 ====================

@router.delete("/sessions/{session_id}")
def delete_session(
    session_id: int,
    current_user=Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """删除会话"""
    session = (
        db.query(RuleSession)
        .filter(RuleSession.id == session_id, RuleSession.user_id == current_user.id)
        .first()
    )
    if not session:
        raise HTTPException(status_code=404, detail="会话不存在")

    db.delete(session)
    db.commit()
    return {"ok": True}
