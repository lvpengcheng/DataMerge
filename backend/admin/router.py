"""
管理后台路由 - /api/admin/*
所有接口需要管理员权限
"""

import os
from pathlib import Path
from datetime import datetime
from typing import Optional, List
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from ..database.connection import get_db
from ..database.models import User, Role, Organization, TenantAuthorization
from ..auth.dependencies import require_admin, get_current_user
from ..auth.schemas import (
    UserCreate, UserUpdate, UserResponse,
    RoleCreate, RoleUpdate, RoleResponse,
    OrgCreate, OrgUpdate, OrgResponse,
    TenantAuthCreate, TenantAuthResponse,
)
from ..auth.utils import get_password_hash

router = APIRouter(prefix="/api/admin", tags=["admin"])


# ========================= 用户管理 =========================

@router.get("/users", response_model=List[UserResponse])
async def list_users(
    org_id: Optional[int] = Query(None),
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取用户列表"""
    query = db.query(User)
    if org_id is not None:
        query = query.filter(User.org_id == org_id)
    users = query.order_by(User.id).all()
    return [_build_user_resp(u) for u in users]


@router.post("/users", response_model=UserResponse)
async def create_user(
    req: UserCreate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """创建用户"""
    if db.query(User).filter(User.username == req.username).first():
        raise HTTPException(status_code=400, detail="用户名已存在")

    user = User(
        username=req.username,
        password_hash=get_password_hash(req.password),
        display_name=req.display_name or req.username,
        email=req.email or "",
        phone=req.phone or "",
        org_id=req.org_id,
        role_id=req.role_id,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return _build_user_resp(user)


@router.get("/users/{user_id}", response_model=UserResponse)
async def get_user(
    user_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取用户详情"""
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="用户不存在")
    return _build_user_resp(user)


@router.put("/users/{user_id}", response_model=UserResponse)
async def update_user(
    user_id: int,
    req: UserUpdate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """更新用户"""
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="用户不存在")

    for field, value in req.model_dump(exclude_unset=True).items():
        setattr(user, field, value)
    db.commit()
    db.refresh(user)
    return _build_user_resp(user)


@router.delete("/users/{user_id}")
async def delete_user(
    user_id: int,
    db: Session = Depends(get_db),
    admin: User = Depends(require_admin),
):
    """禁用用户（软删除）"""
    if user_id == admin.id:
        raise HTTPException(status_code=400, detail="不能禁用自己")
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="用户不存在")
    user.is_active = False
    db.commit()
    return {"message": f"用户 {user.username} 已禁用"}


@router.post("/users/{user_id}/reset-password")
async def reset_password(
    user_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """重置用户密码为 123456"""
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="用户不存在")
    user.password_hash = get_password_hash("123456")
    db.commit()
    return {"message": f"用户 {user.username} 密码已重置为 123456"}


# ========================= 角色管理 =========================

@router.get("/roles", response_model=List[RoleResponse])
async def list_roles(
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取角色列表"""
    return db.query(Role).order_by(Role.id).all()


@router.post("/roles", response_model=RoleResponse)
async def create_role(
    req: RoleCreate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """创建角色"""
    if db.query(Role).filter(Role.name == req.name).first():
        raise HTTPException(status_code=400, detail="角色名已存在")
    role = Role(name=req.name, description=req.description or "", permissions=req.permissions or {})
    db.add(role)
    db.commit()
    db.refresh(role)
    return role


@router.put("/roles/{role_id}", response_model=RoleResponse)
async def update_role(
    role_id: int,
    req: RoleUpdate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """更新角色"""
    role = db.query(Role).filter(Role.id == role_id).first()
    if not role:
        raise HTTPException(status_code=404, detail="角色不存在")
    if role.is_system and req.name and req.name != role.name:
        raise HTTPException(status_code=400, detail="系统角色名称不可修改")

    for field, value in req.model_dump(exclude_unset=True).items():
        setattr(role, field, value)
    db.commit()
    db.refresh(role)
    return role


@router.delete("/roles/{role_id}")
async def delete_role(
    role_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """删除角色"""
    role = db.query(Role).filter(Role.id == role_id).first()
    if not role:
        raise HTTPException(status_code=404, detail="角色不存在")
    if role.is_system:
        raise HTTPException(status_code=400, detail="系统角色不可删除")
    # 检查是否有用户使用该角色
    user_count = db.query(User).filter(User.role_id == role_id).count()
    if user_count > 0:
        raise HTTPException(status_code=400, detail=f"该角色下有 {user_count} 个用户，无法删除")
    db.delete(role)
    db.commit()
    return {"message": f"角色 {role.name} 已删除"}


# ========================= 组织管理 =========================

@router.get("/organizations", response_model=List[OrgResponse])
async def list_organizations(
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取组织列表（树形结构）"""
    orgs = db.query(Organization).filter(Organization.is_active == True).order_by(Organization.id).all()
    return _build_org_tree(orgs)


@router.post("/organizations", response_model=OrgResponse)
async def create_organization(
    req: OrgCreate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """创建组织"""
    if req.parent_id:
        parent = db.query(Organization).filter(Organization.id == req.parent_id).first()
        if not parent:
            raise HTTPException(status_code=400, detail="父组织不存在")

    org = Organization(name=req.name, parent_id=req.parent_id, description=req.description or "")
    db.add(org)
    db.commit()
    db.refresh(org)
    return OrgResponse(
        id=org.id, name=org.name, parent_id=org.parent_id,
        description=org.description, is_active=org.is_active, children=[],
    )


@router.put("/organizations/{org_id}", response_model=OrgResponse)
async def update_organization(
    org_id: int,
    req: OrgUpdate,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """更新组织"""
    org = db.query(Organization).filter(Organization.id == org_id).first()
    if not org:
        raise HTTPException(status_code=404, detail="组织不存在")

    for field, value in req.model_dump(exclude_unset=True).items():
        setattr(org, field, value)
    db.commit()
    db.refresh(org)
    return OrgResponse(
        id=org.id, name=org.name, parent_id=org.parent_id,
        description=org.description, is_active=org.is_active, children=[],
    )


@router.delete("/organizations/{org_id}")
async def delete_organization(
    org_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """禁用组织（软删除）"""
    org = db.query(Organization).filter(Organization.id == org_id).first()
    if not org:
        raise HTTPException(status_code=404, detail="组织不存在")

    # 检查是否有用户
    user_count = db.query(User).filter(User.org_id == org_id, User.is_active == True).count()
    if user_count > 0:
        raise HTTPException(status_code=400, detail=f"该组织下有 {user_count} 个活跃用户，无法删除")

    org.is_active = False
    db.commit()
    return {"message": f"组织 {org.name} 已禁用"}


# ========================= 租户授权管理 =========================

@router.get("/tenant-auth", response_model=List[TenantAuthResponse])
async def list_tenant_auth(
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取所有租户授权"""
    auths = (
        db.query(TenantAuthorization)
        .filter(TenantAuthorization.revoked_at.is_(None))
        .order_by(TenantAuthorization.tenant_id)
        .all()
    )
    return [_build_auth_resp(a) for a in auths]


@router.post("/tenant-auth", response_model=TenantAuthResponse)
async def grant_tenant_auth(
    req: TenantAuthCreate,
    db: Session = Depends(get_db),
    admin: User = Depends(require_admin),
):
    """授权组织访问租户"""
    # 检查组织
    org = db.query(Organization).filter(Organization.id == req.org_id).first()
    if not org:
        raise HTTPException(status_code=400, detail="组织不存在")

    # 检查是否已授权
    existing = (
        db.query(TenantAuthorization)
        .filter(
            TenantAuthorization.tenant_id == req.tenant_id,
            TenantAuthorization.org_id == req.org_id,
            TenantAuthorization.revoked_at.is_(None),
        )
        .first()
    )
    if existing:
        raise HTTPException(status_code=400, detail="该组织已拥有该租户的访问权限")

    auth = TenantAuthorization(
        tenant_id=req.tenant_id,
        org_id=req.org_id,
        auth_type=req.auth_type,
        granted_by=admin.id,
    )
    db.add(auth)
    db.commit()
    db.refresh(auth)
    return _build_auth_resp(auth)


@router.delete("/tenant-auth/{auth_id}")
async def revoke_tenant_auth(
    auth_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """撤销租户授权"""
    auth = db.query(TenantAuthorization).filter(TenantAuthorization.id == auth_id).first()
    if not auth:
        raise HTTPException(status_code=404, detail="授权记录不存在")
    if auth.revoked_at:
        raise HTTPException(status_code=400, detail="该授权已被撤销")
    auth.revoked_at = datetime.utcnow()
    db.commit()
    return {"message": "授权已撤销"}


@router.get("/tenant-auth/tenants")
async def list_all_tenants(
    _admin: User = Depends(require_admin),
):
    """列出所有文件系统中的租户（供选择器使用）"""
    tenants_dir = Path(__file__).resolve().parent.parent.parent / "tenants"
    if not tenants_dir.exists():
        return []
    return sorted([d.name for d in tenants_dir.iterdir() if d.is_dir()])


# ========================= 工具函数 =========================

def _build_user_resp(user: User) -> UserResponse:
    return UserResponse(
        id=user.id,
        username=user.username,
        display_name=user.display_name or "",
        email=user.email or "",
        phone=user.phone or "",
        org_id=user.org_id,
        org_name=user.organization.name if user.organization else "",
        role_id=user.role_id,
        role_name=user.role.name if user.role else "",
        is_active=user.is_active,
    )


def _build_auth_resp(auth: TenantAuthorization) -> TenantAuthResponse:
    return TenantAuthResponse(
        id=auth.id,
        tenant_id=auth.tenant_id,
        org_id=auth.org_id,
        org_name=auth.organization.name if auth.organization else "",
        auth_type=auth.auth_type,
        granted_by=auth.granted_by,
        granted_at=str(auth.granted_at) if auth.granted_at else None,
        revoked_at=str(auth.revoked_at) if auth.revoked_at else None,
    )


def _build_org_tree(orgs: list) -> List[OrgResponse]:
    """将扁平组织列表构建为树形结构"""
    org_map = {}
    for org in orgs:
        org_map[org.id] = OrgResponse(
            id=org.id, name=org.name, parent_id=org.parent_id,
            description=org.description or "", is_active=org.is_active, children=[],
        )
    roots = []
    for org_resp in org_map.values():
        if org_resp.parent_id and org_resp.parent_id in org_map:
            org_map[org_resp.parent_id].children.append(org_resp)
        else:
            roots.append(org_resp)
    return roots
