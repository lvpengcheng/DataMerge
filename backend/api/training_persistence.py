"""
训练持久化服务 - 将训练过程数据写入数据库
"""

import logging
from datetime import datetime
from typing import Optional, Dict, Any, List, Union

from sqlalchemy.orm import Session

from ..database.models import (
    TrainingSession, TrainingIteration, Script, DataAsset,
)

logger = logging.getLogger(__name__)


class TrainingPersistence:
    """训练数据持久化服务"""

    def __init__(self, db: Session):
        self.db = db

    def create_session(
        self,
        tenant_id: str,
        session_key: str,
        mode: str = "formula",
        user_id: Optional[int] = None,
        config: Optional[Dict] = None,
        source_asset_ids: Optional[List[int]] = None,
        expected_asset_id: Optional[int] = None,
    ) -> TrainingSession:
        """创建训练会话（session_key重复时自动追加时间戳）"""
        # 如果session_key已存在，追加时间戳使其唯一
        existing = self.db.query(TrainingSession).filter(
            TrainingSession.session_key == session_key
        ).first()
        if existing:
            session_key = f"{session_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        session = TrainingSession(
            tenant_id=tenant_id,
            session_key=session_key,
            mode=mode,
            user_id=user_id,
            status="running",
            config=config,
            source_asset_ids=source_asset_ids,
            expected_asset_id=expected_asset_id,
        )
        self.db.add(session)
        self.db.commit()
        self.db.refresh(session)
        logger.info(f"训练会话已创建: id={session.id}, key={session_key}")
        return session

    def record_iteration(
        self,
        session_id: int,
        iteration_num: int,
        prompt_text: Optional[str] = None,
        ai_response: Optional[str] = None,
        generated_code: Optional[str] = None,
        accuracy: Optional[float] = None,
        execution_result: Optional[Dict] = None,
        error_details: Optional[Dict] = None,
        duration_seconds: Optional[float] = None,
        status: str = "completed",
    ) -> TrainingIteration:
        """记录一次训练迭代"""
        iteration = TrainingIteration(
            session_id=session_id,
            iteration_num=iteration_num,
            status=status,
            prompt_text=prompt_text,
            ai_response=ai_response,
            generated_code=generated_code,
            accuracy=accuracy,
            execution_result=execution_result,
            error_details=error_details,
            duration_seconds=duration_seconds,
            finished_at=datetime.utcnow() if status in ("completed", "failed") else None,
        )
        self.db.add(iteration)
        self.db.commit()
        self.db.refresh(iteration)
        return iteration

    def update_session_best(
        self,
        session_id: int,
        best_accuracy: float,
        total_iterations: int,
    ):
        """更新会话的最佳准确率"""
        session = self.db.query(TrainingSession).filter_by(id=session_id).first()
        if session:
            session.best_accuracy = best_accuracy
            session.total_iterations = total_iterations
            self.db.commit()

    def complete_session(
        self,
        session_id: int,
        status: str = "completed",
        best_accuracy: Optional[float] = None,
        total_iterations: Optional[int] = None,
        final_script_id: Optional[int] = None,
        error_message: Optional[str] = None,
    ):
        """完成训练会话"""
        session = self.db.query(TrainingSession).filter_by(id=session_id).first()
        if session:
            session.status = status
            session.finished_at = datetime.utcnow()
            if best_accuracy is not None:
                session.best_accuracy = best_accuracy
            if total_iterations is not None:
                session.total_iterations = total_iterations
            if final_script_id is not None:
                session.final_script_id = final_script_id
            if error_message is not None:
                session.error_message = error_message
            self.db.commit()
            logger.info(f"训练会话完成: id={session_id}, status={status}, accuracy={best_accuracy}")

    def save_script(
        self,
        tenant_id: str,
        name: str,
        code: str,
        mode: str = "formula",
        config: Optional[Dict] = None,
        source_session_id: Optional[int] = None,
        accuracy: Optional[float] = None,
        created_by: Optional[int] = None,
        manual_headers: Optional[Dict] = None,
        source_structure: Optional[Any] = None,
        rules_content: Optional[str] = None,
        expected_structure: Optional[Any] = None,
    ) -> Script:
        """保存训练产出的脚本"""
        # 查找同名脚本，自增版本号
        existing = (
            self.db.query(Script)
            .filter_by(tenant_id=tenant_id, name=name, is_active=True)
            .order_by(Script.version.desc())
            .first()
        )
        version = (existing.version + 1) if existing else 1
        if existing:
            existing.is_active = False

        script = Script(
            tenant_id=tenant_id,
            name=name,
            code=code,
            mode=mode,
            config=config,
            source_session_id=source_session_id,
            accuracy=accuracy,
            version=version,
            created_by=created_by,
            manual_headers=manual_headers,
            source_structure=source_structure,
            rules_content=rules_content,
            expected_structure=expected_structure,
        )
        self.db.add(script)
        self.db.commit()
        self.db.refresh(script)
        logger.info(f"脚本已保存: id={script.id}, name={name}, version={version}")
        return script

    def get_session(self, session_id: int) -> Optional[TrainingSession]:
        """获取训练会话"""
        return self.db.query(TrainingSession).filter_by(id=session_id).first()

    def get_session_by_key(self, session_key: str) -> Optional[TrainingSession]:
        """根据 session_key 获取训练会话"""
        return self.db.query(TrainingSession).filter_by(session_key=session_key).first()

    def list_sessions(
        self,
        tenant_id: str,
        limit: int = 50,
        offset: int = 0,
    ) -> List[TrainingSession]:
        """列出租户的训练会话"""
        return (
            self.db.query(TrainingSession)
            .filter_by(tenant_id=tenant_id)
            .order_by(TrainingSession.started_at.desc())
            .offset(offset)
            .limit(limit)
            .all()
        )

    def get_iterations(self, session_id: int) -> List[TrainingIteration]:
        """获取会话的所有迭代"""
        return (
            self.db.query(TrainingIteration)
            .filter_by(session_id=session_id)
            .order_by(TrainingIteration.iteration_num)
            .all()
        )

    def get_best_script(self, tenant_id: str) -> Optional[Script]:
        """获取租户最新的活跃脚本"""
        return (
            self.db.query(Script)
            .filter_by(tenant_id=tenant_id, is_active=True)
            .order_by(Script.created_at.desc())
            .first()
        )
