"""
管理后台路由 - /api/admin/*
所有接口需要管理员权限
"""

import os
import re
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, List
from fastapi import APIRouter, Depends, HTTPException, Query, File, UploadFile, Form, Request
from fastapi.responses import FileResponse
from sqlalchemy.orm import Session
from sqlalchemy import or_

from ..database.connection import get_db
from ..database.models import User, Role, Organization, TenantAuthorization, Template, ComputeTask, DataAsset
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


# ========================= 模版管理 =========================

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


def _save_template_file(file: UploadFile, tenant_id: Optional[str]) -> tuple:
    """保存模版文件，返回 (file_path, file_name)"""
    if tenant_id:
        base_dir = _PROJECT_ROOT / "tenants" / tenant_id / "templates"
    else:
        base_dir = _PROJECT_ROOT / "global_assets" / "templates"
    base_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
    safe_name = file.filename.replace(" ", "_")
    saved_name = f"{timestamp}_{safe_name}"
    saved_path = base_dir / saved_name

    with open(saved_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    return str(saved_path), file.filename


def _build_template_resp(t: Template) -> dict:
    return {
        "id": t.id,
        "tenant_id": t.tenant_id,
        "name": t.name,
        "description": t.description or "",
        "file_name": t.file_name,
        "file_name_rule": t.file_name_rule or "",
        "encrypt_type": t.encrypt_type or "none",
        "encrypt_password": t.encrypt_password or "",
        "report_mode": getattr(t, "report_mode", "fill") or "fill",
        "group_by": getattr(t, "group_by", "") or "",
        "skip_rows": getattr(t, "skip_rows", 1) or 1,
        "name_field": getattr(t, "name_field", "") or "",
        "show_empty_period": getattr(t, "show_empty_period", True),
        "is_active": t.is_active,
        "created_by": t.created_by,
        "creator_name": t.creator.display_name if t.creator else "",
        "created_at": t.created_at.isoformat() if t.created_at else None,
        "updated_at": t.updated_at.isoformat() if t.updated_at else None,
    }


@router.get("/templates")
async def list_templates(
    tenant_id: Optional[str] = Query(None),
    include_global: bool = Query(False),
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取模版列表"""
    query = db.query(Template).filter(Template.is_active == True)
    if tenant_id:
        if tenant_id == "__global__":
            query = query.filter(Template.tenant_id.is_(None))
        elif include_global:
            query = query.filter(or_(Template.tenant_id == tenant_id, Template.tenant_id.is_(None)))
        else:
            query = query.filter(Template.tenant_id == tenant_id)
    templates = query.order_by(Template.id.desc()).all()
    return [_build_template_resp(t) for t in templates]


@router.post("/templates")
async def create_template(
    file: UploadFile = File(...),
    tenant_id: Optional[str] = Form(None),
    name: str = Form(...),
    description: str = Form(""),
    file_name_rule: str = Form(""),
    encrypt_type: str = Form("none"),
    encrypt_password: str = Form(""),
    report_mode: str = Form("fill"),
    group_by: str = Form(""),
    skip_rows: int = Form(1),
    name_field: str = Form(""),
    show_empty_period: bool = Form(True),
    db: Session = Depends(get_db),
    admin: User = Depends(require_admin),
):
    """创建模版（含文件上传）"""
    file_path, file_name = _save_template_file(file, tenant_id)

    tpl = Template(
        tenant_id=tenant_id or None,
        name=name,
        description=description,
        file_path=file_path,
        file_name=file_name,
        file_name_rule=file_name_rule,
        encrypt_type=encrypt_type,
        encrypt_password=encrypt_password,
        report_mode=report_mode,
        group_by=group_by,
        skip_rows=skip_rows,
        name_field=name_field,
        show_empty_period=show_empty_period,
        created_by=admin.id,
    )
    db.add(tpl)
    db.commit()
    db.refresh(tpl)
    return _build_template_resp(tpl)


@router.get("/templates/{template_id}")
async def get_template(
    template_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """获取模版详情"""
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")
    return _build_template_resp(tpl)


@router.put("/templates/{template_id}")
async def update_template(
    request: Request,
    template_id: int,
    file: Optional[UploadFile] = File(None),
    tenant_id: Optional[str] = Form(None),
    name: Optional[str] = Form(None),
    description: Optional[str] = Form(None),
    file_name_rule: Optional[str] = Form(None),
    encrypt_type: Optional[str] = Form(None),
    encrypt_password: Optional[str] = Form(None),
    report_mode: Optional[str] = Form(None),
    group_by: Optional[str] = Form(None),
    skip_rows: Optional[int] = Form(None),
    name_field: Optional[str] = Form(None),
    show_empty_period: Optional[bool] = Form(None),
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """更新模版（可选替换文件）"""
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")

    if file and file.filename:
        file_path, file_name = _save_template_file(file, tenant_id or tpl.tenant_id)
        tpl.file_path = file_path
        tpl.file_name = file_name

    # 读取原始表单数据，解决 FastAPI 把空字符串转为 None 的问题
    form = await request.form()

    if tenant_id is not None:
        tpl.tenant_id = tenant_id or None
    if name is not None:
        tpl.name = name
    if description is not None:
        tpl.description = description
    # 可清空的字段：检查原始表单是否包含该字段（即使值为空字符串）
    if 'file_name_rule' in form:
        tpl.file_name_rule = str(form.get('file_name_rule', ''))
    elif file_name_rule is not None:
        tpl.file_name_rule = file_name_rule
    if encrypt_type is not None:
        tpl.encrypt_type = encrypt_type
    if 'encrypt_password' in form:
        new_pwd = str(form.get('encrypt_password', ''))
        logger.info(f"[模版更新] encrypt_password: '{tpl.encrypt_password}' -> '{new_pwd}'")
        tpl.encrypt_password = new_pwd
    elif encrypt_password is not None:
        tpl.encrypt_password = encrypt_password
    if report_mode is not None:
        tpl.report_mode = report_mode
    if 'group_by' in form:
        tpl.group_by = str(form.get('group_by', ''))
    elif group_by is not None:
        tpl.group_by = group_by
    if skip_rows is not None:
        tpl.skip_rows = skip_rows
    if 'name_field' in form:
        tpl.name_field = str(form.get('name_field', ''))
    elif name_field is not None:
        tpl.name_field = name_field
    if show_empty_period is not None:
        tpl.show_empty_period = show_empty_period

    db.commit()
    db.refresh(tpl)
    return _build_template_resp(tpl)


@router.delete("/templates/{template_id}")
async def delete_template(
    template_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """停用模版（软删除）"""
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")
    tpl.is_active = False
    db.commit()
    return {"message": f"模版 {tpl.name} 已停用"}


@router.get("/templates/{template_id}/download")
async def download_template(
    template_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """下载模版文件"""
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")
    if not os.path.exists(tpl.file_path):
        raise HTTPException(status_code=404, detail="模版文件不存在")
    return FileResponse(
        path=tpl.file_path,
        filename=tpl.file_name,
        media_type="application/octet-stream",
    )


# ========================= 报表生成 =========================

logger = logging.getLogger(__name__)

# 匹配 {变量名} 或 {变量名[:N]} 或 {变量名[-N:]} 模式
_RULE_PATTERN = re.compile(r'\{([^{}]+)\}')
_SLICE_PATTERN = re.compile(r'^(.+?)\[(-?\d*):(-?\d*)\]$')


def _resolve_rule_pattern(pattern: str, data_row: dict, system_vars: dict) -> str:
    """解析规则表达式，替换 {变量} 为实际值。

    支持:
      {year} {month} {date} {tenant} — 系统变量
      {列名}       — 取数据行中该列的完整值
      {列名[:N]}   — 取前 N 位
      {列名[-N:]}  — 取后 N 位
    """
    def _replace(m):
        expr = m.group(1).strip()
        # 先检查系统变量
        if expr in system_vars:
            return str(system_vars[expr])
        # 检查切片语法
        slice_m = _SLICE_PATTERN.match(expr)
        if slice_m:
            col_name = slice_m.group(1).strip()
            start = slice_m.group(2)
            end = slice_m.group(3)
            val = str(data_row.get(col_name, ""))
            if start and end:
                return val[int(start):int(end)]
            elif start:
                return val[int(start):]
            elif end:
                return val[:int(end)]
            return val
        # 普通列名
        return str(data_row.get(expr, ""))

    return _RULE_PATTERN.sub(_replace, pattern)


@router.post("/templates/{template_id}/generate-report")
async def generate_report(
    template_id: int,
    task_id: int = Form(...),
    use_history: bool = Form(False),
    period_from: Optional[str] = Form(None),
    period_to: Optional[str] = Form(None),
    db: Session = Depends(get_db),
    admin: User = Depends(require_admin),
):
    """基于模版 + 计算结果生成报表"""
    import pandas as pd
    from ..utils import aspose_helper

    # 1. 查模版
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")
    if not os.path.exists(tpl.file_path):
        raise HTTPException(status_code=404, detail="模版文件不存在")

    # 2. 查当前任务
    task = db.query(ComputeTask).filter(ComputeTask.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="计算任务不存在")
    tenant_id = task.tenant_id

    # 3. 收集数据资产
    if use_history and period_from and period_to:
        # 解析周期范围 YYYY-MM
        try:
            from_y, from_m = int(period_from[:4]), int(period_from[5:7])
            to_y, to_m = int(period_to[:4]), int(period_to[5:7])
        except (ValueError, IndexError):
            raise HTTPException(status_code=400, detail="薪资周期格式错误，请使用 YYYY-MM")

        # 查询该租户所有已完成且有薪资周期的任务
        all_tasks = (
            db.query(ComputeTask)
            .filter(
                ComputeTask.tenant_id == tenant_id,
                ComputeTask.status == "completed",
                ComputeTask.salary_year.isnot(None),
                ComputeTask.salary_month.isnot(None),
            )
            .all()
        )

        # 周期范围内，每月取最后一次
        ym_from = from_y * 100 + from_m
        ym_to = to_y * 100 + to_m
        last_per_month = {}  # {(year, month): ComputeTask}
        for t in all_tasks:
            ym = t.salary_year * 100 + t.salary_month
            if ym_from <= ym <= ym_to:
                key = (t.salary_year, t.salary_month)
                if key not in last_per_month or t.created_at > last_per_month[key].created_at:
                    last_per_month[key] = t

        # 始终包含当前任务
        task_ids = set(t.id for t in last_per_month.values())
        task_ids.add(task_id)
        task_ids = list(task_ids)

        if not task_ids:
            raise HTTPException(status_code=400, detail="所选薪资周期范围内没有已完成的计算任务")

        # 构建 task_id → (salary_year, salary_month) 映射，用于自动补月份列
        task_ym_map = {}
        for ym_key, t_obj in last_per_month.items():
            task_ym_map[t_obj.id] = ym_key  # (year, month)
        # 当前任务也加入映射
        if task_id not in task_ym_map and task.salary_year and task.salary_month:
            task_ym_map[task_id] = (task.salary_year, task.salary_month)

        assets = (
            db.query(DataAsset)
            .filter(
                DataAsset.source_task_id.in_(task_ids),
                DataAsset.asset_type == "result",
                DataAsset.is_active == True,
            )
            .all()
        )
    else:
        # 不启用历史：仅当前任务的结果
        assets = (
            db.query(DataAsset)
            .filter(
                DataAsset.source_task_id == task_id,
                DataAsset.asset_type == "result",
                DataAsset.is_active == True,
            )
            .all()
        )

    if not assets:
        raise HTTPException(status_code=400, detail="未找到计算结果")

    # 4. 从 DB parsed_data 读取数据（优先），无 parsed_data 时回退到读文件
    #    只取 sheet0 的数据作为 dt
    #    use_history 模式下自动补 salary_year / salary_month / 月份 列
    is_multi_month = use_history and period_from and period_to
    all_dfs = []
    for asset in assets:
        asset_dfs = []
        if asset.parsed_data:
            # parsed_data: [{"sheet_name": "...", "regions": [...]}, ...]
            # 只取第一个 sheet
            first_sheet = asset.parsed_data[0] if asset.parsed_data else None
            if first_sheet:
                for region in (first_sheet.get("regions") or []):
                    head_data = region.get("head_data") or {}
                    data_rows = region.get("data") or []
                    if not head_data or not data_rows:
                        continue
                    col_map = {v: k for k, v in head_data.items()}
                    mapped_rows = [{col_map.get(c, c): val for c, val in row.items()} for row in data_rows]
                    df = pd.DataFrame(mapped_rows)
                    if not df.empty:
                        asset_dfs.append(df)
        elif os.path.exists(asset.file_path):
            # 回退：从文件读取，只取第一个 sheet
            try:
                sheets = aspose_helper.read_all_sheets_calculated(asset.file_path)
                if sheets:
                    first_df = list(sheets.values())[0]
                    if not first_df.empty:
                        asset_dfs.append(first_df)
            except Exception as e:
                logger.warning(f"读取结果文件失败 {asset.file_path}: {e}")

        # 多月合并时，自动补 salary_year / salary_month / 月份 列
        if is_multi_month and asset_dfs and asset.source_task_id in task_ym_map:
            y, m = task_ym_map[asset.source_task_id]
            for df in asset_dfs:
                if "salary_year" not in df.columns:
                    df["salary_year"] = y
                if "salary_month" not in df.columns:
                    df["salary_month"] = m
                if "月份" not in df.columns:
                    df["月份"] = f"{m}月"

        all_dfs.extend(asset_dfs)

    if not all_dfs:
        raise HTTPException(status_code=400, detail="无法读取计算结果数据")

    dataset = pd.concat(all_dfs, ignore_index=True) if len(all_dfs) > 1 else all_dfs[0]

    # 多月合并时：show_empty_period 补齐缺失月份的空行
    show_empty = getattr(tpl, "show_empty_period", True)
    if is_multi_month and show_empty and "月份" in dataset.columns:
        # 生成完整月份列表
        all_months = []
        cy, cm = from_y, from_m
        while cy * 100 + cm <= to_y * 100 + to_m:
            all_months.append(f"{cm}月")
            cm += 1
            if cm > 12:
                cm = 1
                cy += 1
        existing_months = set(dataset["月份"].unique())
        for month_label in all_months:
            if month_label not in existing_months:
                empty_row = {col: None for col in dataset.columns}
                empty_row["月份"] = month_label
                dataset = pd.concat([dataset, pd.DataFrame([empty_row])], ignore_index=True)
                logger.info(f"[多月合并] 补齐空月份: {month_label}")

    # 5. 构建模版数据字典
    #    - "DataSource" = 完整数据集（模版中写 &=DataSource.列名）
    #    - 系统变量（模版中写 &=$year &=$month 等）
    now = datetime.utcnow()
    system_vars = {
        "year": str(now.year),
        "month": f"{now.month:02d}",
        "date": now.strftime("%Y%m%d"),
        "tenant": tenant_id,
    }

    template_data = {
        "DT": dataset,
        "$year": system_vars["year"],
        "$month": system_vars["month"],
        "$date": system_vars["date"],
        "$tenant": tenant_id,
    }

    # 6. 用数据集第一行来解析文件名和加密规则
    first_row = dataset.iloc[0].to_dict() if len(dataset) > 0 else {}

    if tpl.file_name_rule:
        output_name = _resolve_rule_pattern(tpl.file_name_rule, first_row, system_vars)
        if not output_name.endswith(('.xlsx', '.xls')):
            output_name += '.xlsx'
    else:
        output_name = f"报表_{tpl.name}_{now.strftime('%Y%m%d%H%M%S')}.xlsx"

    # 7. 解析加密规则（可以是固定值或参数表达式）
    password = None
    logger.info(f"数据库 encrypt_password='{tpl.encrypt_password}', encrypt_type='{tpl.encrypt_type}'")
    if tpl.encrypt_password:
        password = _resolve_rule_pattern(tpl.encrypt_password, first_row, system_vars)
        if password:
            logger.info(f"报表加密: 模版={tpl.name}, 密码长度={len(password)}")
        else:
            logger.warning(f"加密规则解析为空: rule={tpl.encrypt_password}, columns={list(first_row.keys())[:10]}")

    # 8. 打印 template_data 前5条用于调试
    logger.info("=== template_data 调试信息 ===")
    for k, v in template_data.items():
        if isinstance(v, pd.DataFrame):
            logger.info(f"[{k}] DataFrame shape={v.shape}, columns={list(v.columns)}")
            logger.info(f"[{k}] 前5条:\n{v.head(5).to_string()}")
        else:
            logger.info(f"[{k}] = {v}")
    logger.info(f"文件名规则: {tpl.file_name_rule} -> {output_name if tpl.file_name_rule else '(默认)'}")
    logger.info(f"加密规则: {tpl.encrypt_password} -> {'***' if password else '无'}")
    logger.info("=== end ===")

    # 9. 生成报表
    report_mode = getattr(tpl, "report_mode", "fill") or "fill"
    group_by_field = getattr(tpl, "group_by", "") or ""
    skip_rows_val = getattr(tpl, "skip_rows", 1) or 1
    name_field_val = getattr(tpl, "name_field", "") or ""
    show_empty = getattr(tpl, "show_empty_period", True)

    # zip/block 模式前置校验：group_by 不能为空，且必须在数据列中
    if report_mode in ("zip", "block"):
        if not group_by_field:
            raise HTTPException(
                status_code=400,
                detail=f"报表模式为 {report_mode}，但模版未配置分组字段(group_by)，请在模版设置中指定分组列名",
            )
        available_cols = list(dataset.columns)
        # 模糊匹配：去空格、忽略大小写
        matched_col = None
        target = group_by_field.strip().lower()
        for col in available_cols:
            if str(col).strip().lower() == target:
                matched_col = col
                break
        if not matched_col:
            raise HTTPException(
                status_code=400,
                detail=f"分组字段 '{group_by_field}' 不在数据列中，可用列: {available_cols[:30]}",
            )
        # 如果匹配到的列名与配置不完全一致，使用实际列名
        if matched_col != group_by_field:
            logger.info(f"group_by 模糊匹配: '{group_by_field}' -> '{matched_col}'")
            group_by_field = matched_col

    # zip 模式输出 .zip，其余输出原始扩展名
    if report_mode == "zip":
        output_ext = ".zip"
        output_name_final = os.path.splitext(output_name)[0] + output_ext
    else:
        output_name_final = output_name

    output_dir = _PROJECT_ROOT / "tenants" / tenant_id / "reports"
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = now.strftime("%Y%m%d%H%M%S")
    output_path = str(output_dir / f"{timestamp}_{output_name_final}")

    logger.info(f"报表模式: {report_mode}, group_by={group_by_field}, skip_rows={skip_rows_val}")

    try:
        actual_output_path = aspose_helper.generate_from_template(
            output_path=output_path,
            template_path=tpl.file_path,
            data=template_data,
            password=password,
            mode=report_mode,
            group_by=group_by_field,
            skip_rows=skip_rows_val,
            name_field=name_field_val,
            show_empty_period=show_empty,
        )
        # 实际输出路径可能和请求路径不同（如 zip 回退到 fill 时扩展名变为 .xlsx）
        output_path = actual_output_path
    except Exception as e:
        logger.error(f"报表生成失败: {e}")
        raise HTTPException(status_code=500, detail=f"报表生成失败: {str(e)}")

    # 10. 留痕 — 保存为 DataAsset
    actual_filename = os.path.basename(output_path)
    try:
        report_asset = DataAsset(
            tenant_id=tenant_id,
            asset_type="report",
            name=f"报表_{tpl.name}_{now.strftime('%Y%m%d')}",
            file_path=output_path,
            file_name=actual_filename,
            file_size=os.path.getsize(output_path),
            source_task_id=task_id,
            uploaded_by=admin.id,
            tags={
                "template_id": template_id,
                "template_name": tpl.name,
                "period_from": period_from,
                "period_to": period_to,
                "use_history": use_history,
            },
        )
        db.add(report_asset)
        db.commit()
    except Exception as e:
        logger.warning(f"报表留痕失败: {e}")
        try:
            db.rollback()
        except Exception:
            pass

    # 11. 返回文件下载（根据实际文件扩展名决定 MIME 类型）
    is_zip = output_path.lower().endswith(".zip")
    media = "application/zip" if is_zip else "application/octet-stream"
    return FileResponse(
        path=output_path,
        filename=actual_filename,
        media_type=media,
    )
