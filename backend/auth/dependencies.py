"""
FastAPI 认证依赖
"""

from typing import Optional, List
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from sqlalchemy.orm import Session

from ..database.connection import get_db
from ..database.models import User, TenantAuthorization
from .utils import decode_access_token

security = HTTPBearer(auto_error=False)


async def get_current_user(
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
    db: Session = Depends(get_db),
) -> User:
    """从 Authorization: Bearer <token> 中解析并验证用户"""
    if not credentials:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="未提供认证令牌")

    payload = decode_access_token(credentials.credentials)
    if payload is None:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="令牌无效或已过期")

    user_id = payload.get("user_id")
    if not user_id:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="令牌数据不完整")

    user = db.query(User).filter(User.id == user_id, User.is_active == True).first()
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="用户不存在或已禁用")

    return user


async def require_admin(current_user: User = Depends(get_current_user)) -> User:
    """要求管理员权限"""
    if not current_user.role or current_user.role.name != "admin":
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="需要管理员权限")
    return current_user


def get_accessible_tenants(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> List[str]:
    """获取当前用户可访问的所有租户 ID 列表"""
    # admin 可以访问所有租户
    if current_user.role and current_user.role.name == "admin":
        auths = db.query(TenantAuthorization).filter(
            TenantAuthorization.revoked_at.is_(None)
        ).all()
    else:
        # 普通用户只能访问自己组织的租户
        if not current_user.org_id:
            return []
        # 收集用户所属组织及所有子组织的 ID
        org_ids = _get_org_and_children_ids(db, current_user.org_id)
        auths = db.query(TenantAuthorization).filter(
            TenantAuthorization.org_id.in_(org_ids),
            TenantAuthorization.revoked_at.is_(None),
        ).all()

    return list(set(a.tenant_id for a in auths))


def get_operable_tenants(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> List[str]:
    """当前用户可"操作"的租户 = 已授权租户 ∪ 自己训练过的租户。

    用于驱动前端按钮可用性 + 后端写接口权限校验。admin 仍走 TenantAuthorization 全集。
    """
    from ..database.models import TrainingSession

    is_admin = current_user.role and current_user.role.name == "admin"

    # 1) 已授权
    if is_admin:
        auths = db.query(TenantAuthorization).filter(
            TenantAuthorization.revoked_at.is_(None)
        ).all()
    elif not current_user.org_id:
        auths = []
    else:
        org_ids = _get_org_and_children_ids(db, current_user.org_id)
        auths = db.query(TenantAuthorization).filter(
            TenantAuthorization.org_id.in_(org_ids),
            TenantAuthorization.revoked_at.is_(None),
        ).all()
    authorized = {a.tenant_id for a in auths}

    if is_admin:
        return list(authorized)

    # 2) 自训(以 TrainingSession.user_id 为准)
    own_rows = (
        db.query(TrainingSession.tenant_id)
        .filter(TrainingSession.user_id == current_user.id)
        .distinct()
        .all()
    )
    own = {r.tenant_id for r in own_rows}

    return list(authorized | own)


def _get_org_and_children_ids(db: Session, org_id: int) -> List[int]:
    """递归获取组织及所有子组织的 ID"""
    from ..database.models import Organization

    result = [org_id]
    children = db.query(Organization).filter(
        Organization.parent_id == org_id,
        Organization.is_active == True,
    ).all()
    for child in children:
        result.extend(_get_org_and_children_ids(db, child.id))
    return result
