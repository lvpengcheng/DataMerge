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
from ..database.models import (
    User, Role, Organization, TenantAuthorization, Template, ComputeTask, DataAsset, Script,
    TrainingSession, TrainingMessage, TrainingIteration, ScriptMigration,
)
from ..auth.dependencies import require_admin, get_current_user
from ..auth.schemas import (
    UserCreate, UserUpdate, UserResponse,
    RoleCreate, RoleUpdate, RoleResponse,
    OrgCreate, OrgUpdate, OrgResponse,
    TenantAuthCreate, TenantAuthResponse,
    AdminSetPasswordRequest,
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


@router.post("/users/{user_id}/set-password")
async def set_password(
    user_id: int,
    req: AdminSetPasswordRequest,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """管理员为指定用户设定新密码"""
    pwd = (req.new_password or "").strip()
    if len(pwd) < 6:
        raise HTTPException(status_code=400, detail="密码至少 6 位")
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="用户不存在")
    user.password_hash = get_password_hash(pwd)
    db.commit()
    return {"message": f"用户 {user.username} 密码已修改"}


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

    # 检查是否已授权（含已撤销的记录，避免 UniqueConstraint 冲突）
    existing = (
        db.query(TenantAuthorization)
        .filter(
            TenantAuthorization.tenant_id == req.tenant_id,
            TenantAuthorization.org_id == req.org_id,
        )
        .first()
    )
    if existing:
        if existing.revoked_at is None:
            raise HTTPException(status_code=400, detail="该组织已拥有该租户的访问权限")
        # 恢复已撤销的授权
        existing.revoked_at = None
        existing.auth_type = req.auth_type
        existing.granted_by = admin.id
        db.commit()
        db.refresh(existing)
        return _build_auth_resp(existing)

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
        "split_by": getattr(t, "split_by", "") or "",
        "show_empty_period": getattr(t, "show_empty_period", True),
        "carry_over_sheets": getattr(t, "carry_over_sheets", "") or "",
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
    split_by: str = Form(""),
    show_empty_period: bool = Form(True),
    carry_over_sheets: str = Form(""),
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
        split_by=split_by,
        show_empty_period=show_empty_period,
        carry_over_sheets=carry_over_sheets,
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
    split_by: Optional[str] = Form(None),
    show_empty_period: Optional[bool] = Form(None),
    carry_over_sheets: Optional[str] = Form(None),
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
    if 'split_by' in form:
        tpl.split_by = str(form.get('split_by', ''))
    elif split_by is not None:
        tpl.split_by = split_by
    if show_empty_period is not None:
        tpl.show_empty_period = show_empty_period
    if 'carry_over_sheets' in form:
        tpl.carry_over_sheets = str(form.get('carry_over_sheets', ''))
    elif carry_over_sheets is not None:
        tpl.carry_over_sheets = carry_over_sheets

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


def _coerce_df_for_report(df):
    """为 SmartMarker 填充做类型规整（优先让模板单元格自身的数字格式生效）：

    - 某列非空值【本就是数值类型 int/float】→ 整列保留数值（空值→None，单元格留空）。
      这样模板里该列若是数值格式(如 0.00) 就能正常生效，不再被写成 "0" 文本。
      整数值转 int 避免 3→3.0（数值格式仍会把 0 显示成 0.00）。
    - 其余列（含"数值型字符串"如工号/长编码/银行卡号，以及 #N/A 等错误码）→ 文本，
      NaN/NaT/None→''。保留字符串是刻意的：解析层已把 ID/长数字统一成文本，
      不能再数值化（否则丢前导零/精度）。
    """
    import pandas as pd

    def _isna(v):
        if v is None:
            return True
        try:
            return bool(pd.isna(v))
        except (TypeError, ValueError):
            return False

    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        s = out[col]
        non_na = [v for v in s.tolist() if not _isna(v)]
        is_num_col = bool(non_na) and all(
            isinstance(v, (int, float)) and not isinstance(v, bool) for v in non_na
        )
        if is_num_col:
            def _num(x):
                if _isna(x):
                    return None
                f = float(x)
                return int(f) if f.is_integer() else f
            out[col] = s.map(_num)
        else:
            s2 = s.where(s.notna(), '')
            out[col] = s2.astype(str).replace({'nan': '', 'NaT': '', 'None': ''})
    return out


def _col_idx_to_letter(idx: int) -> str:
    """0-indexed 列下标 → Excel 列字母 (0→A, 25→Z, 26→AA...)"""
    letters = ""
    n = idx
    while True:
        letters = chr(65 + n % 26) + letters
        n = n // 26 - 1
        if n < 0:
            break
    return letters


def _read_sheets_with_letter_columns(file_path: str):
    """用 excel_parser 解析文件，返回:
       sheets: {sheet_name: DataFrame(列名为 Excel 列字母 A/B/.../AA)}
       header_map: {sheet_name: {letter: 拼接表头文本}}

    数据起始行由 excel_parser 智能识别（跳过 title 行 + 多行表头）。
    """
    import pandas as pd
    from excel_parser import IntelligentExcelParser

    parser = IntelligentExcelParser()
    # calculate_formulas=True：报表数据源常是"计算结果文件"，含公式列但无缓存值
    # （openpyxl/模板模式产出的公式不会自动算）。不先算的话 .Value 读到空，
    # SmartMarker &=DT.[列] 就填成空白。先 CalculateFormula 再取值。
    sheet_data_list = parser.parse_excel_file(file_path, read_formulas=False, calculate_formulas=True)

    sheets: dict = {}
    header_map: dict = {}
    for sd in sheet_data_list:
        sheet_name = sd.sheet_name
        if not sd.regions:
            continue
        # 取该 sheet 的第一个数据区域（计算结果文件通常只有一个）
        region = sd.regions[0]
        # head_data: {header_text: col_letter}  → 反转为 {letter: header}
        letter_to_header = {v: k for k, v in (region.head_data or {}).items()}
        # 数据行已是 {letter: value}
        data_rows = region.data or []
        if not data_rows:
            continue
        df = pd.DataFrame(data_rows)
        # 列名规范化为大写字母
        df.columns = [str(c).upper() for c in df.columns]
        # 按字母自然顺序排列
        df = df.reindex(columns=sorted(df.columns, key=lambda x: (len(x), x)))
        sheets[sheet_name] = df
        header_map[sheet_name] = {str(k).upper(): v for k, v in letter_to_header.items()}
    return sheets, header_map


def _build_header_map_from_parsed(parsed_data) -> dict:
    """从 DataAsset.parsed_data 中提取 {sheet_name: {letter: header_text}}"""
    result: dict = {}
    for sheet_info in (parsed_data or []):
        sheet_name = sheet_info.get("sheet_name", "Sheet1")
        if sheet_name in ("参数", "历史数据"):
            continue
        for region in (sheet_info.get("regions") or []):
            head_data = region.get("head_data") or {}
            if not head_data:
                continue
            letter_to_header = {str(v).upper(): k for k, v in head_data.items()}
            # 同 sheet 多 region 时合并（后者不覆盖前者）
            existing = result.setdefault(sheet_name, {})
            for letter, header in letter_to_header.items():
                existing.setdefault(letter, header)
            break  # 取第一个 region
    return result


def _alias_letter_columns_with_headers(df, sheet_header_map: dict):
    """给 DataFrame 增加"原始拼接表头"作为别名列(指向同一列字母的数据)。

    使 SmartMarker 模板可以同时用 &=DT.A 和 &=DT.原始表头 两种写法。
    冲突规则:
      - 别名与已有列名(尤其是其他字母列)同名时跳过,以字母列优先
      - 别名为空/None 跳过
    """
    if df is None or df.empty or not sheet_header_map:
        return df
    out = df.copy()
    existing_cols = set(out.columns)
    for letter, header in sheet_header_map.items():
        letter_u = str(letter).upper()
        if not header:
            continue
        header_str = str(header).strip()
        if not header_str:
            continue
        if header_str in existing_cols:
            continue
        if letter_u not in out.columns:
            continue
        out[header_str] = out[letter_u]
        existing_cols.add(header_str)
    return out


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

    # 去重：双结果模式(COMPUTE_OUTPUT_VALUES_COPY)下，同一输出会注册"公式版"+"纯值版(_纯值)"
    # 两个 result 资产，二者同数据同 sheet。报表按 sheet 名拼接会使每行数据重复出现。
    # 同一逻辑文件只保留公式版，剔除其纯值副本；仅纯值版存在时(公式版缺失)才保留纯值版。
    try:
        from ..utils.output_postprocess import values_only_name
        _names = {a.file_name for a in assets}
        # 若某资产名是另一资产的"纯值版名"，说明它是重复副本 → 剔除
        _values_dupes = {a.id for a in assets
                         if any(a.file_name == values_only_name(other) for other in _names)}
        if _values_dupes:
            _before = len(assets)
            assets = [a for a in assets if a.id not in _values_dupes]
            logger.info(f"[报表去重] 剔除 {_before - len(assets)} 个纯值版重复资产，剩 {len(assets)} 个")
    except Exception as _dedup_e:
        logger.warning(f"[报表去重] 跳过（不阻断）: {_dedup_e}")

    # 4. 从 DB parsed_data 读取数据（优先），无 parsed_data 时回退到读文件
    #    读取所有 sheet，按 sheet 名分组，每个 sheet 作为独立数据源
    #    use_history 模式下自动补 salary_year / salary_month / 月份 列
    #    【列字母方案】DataFrame 列名统一为 Excel 列字母（A/B/.../AA），原表头作为 header_map 返回
    is_multi_month = use_history and period_from and period_to
    all_sheet_dfs: dict[str, list] = {}  # {sheet_name: [df1, df2, ...]}
    header_map: dict[str, dict] = {}      # {sheet_name: {letter: 拼接表头}}
    for asset in assets:
        asset_sheet_dfs: dict[str, list] = {}
        if asset.parsed_data:
            # parsed_data: [{"sheet_name": "...", "regions": [...]}, ...]
            # head_data = {header_text: col_letter}, region.data 已按 col_letter 索引
            for sheet_info in asset.parsed_data:
                sheet_name = sheet_info.get("sheet_name", "Sheet1")
                if sheet_name in ("参数", "历史数据"):
                    continue
                for region in (sheet_info.get("regions") or []):
                    head_data = region.get("head_data") or {}
                    data_rows = region.get("data") or []
                    if not head_data or not data_rows:
                        continue
                    # 直接用 col_letter 作为列名（大写），不再映射回中文表头
                    norm_rows = [
                        {str(c).upper(): val for c, val in row.items()}
                        for row in data_rows
                    ]
                    df = pd.DataFrame(norm_rows)
                    df = df.reindex(columns=sorted(df.columns, key=lambda x: (len(x), x)))
                    if not df.empty:
                        asset_sheet_dfs.setdefault(sheet_name, []).append(df)
            # 收集 header_map（同 sheet 取第一个 asset 的即可）
            for sn, m in _build_header_map_from_parsed(asset.parsed_data).items():
                header_map.setdefault(sn, m)
        elif os.path.exists(asset.file_path):
            # 回退：用 excel_parser 解析文件，返回字母列 DataFrame + header_map
            try:
                sheets, hm = _read_sheets_with_letter_columns(asset.file_path)
                for sheet_name, df in sheets.items():
                    if sheet_name in ("参数", "历史数据"):
                        continue
                    if not df.empty:
                        asset_sheet_dfs.setdefault(sheet_name, []).append(df)
                for sn, m in hm.items():
                    header_map.setdefault(sn, m)
            except Exception as e:
                logger.warning(f"读取结果文件失败 {asset.file_path}: {e}")

        # 多月合并时，自动补 salary_year / salary_month / 月份 列（命名列，保留原名）
        if is_multi_month and asset_sheet_dfs and asset.source_task_id in task_ym_map:
            y, m = task_ym_map[asset.source_task_id]
            for dfs in asset_sheet_dfs.values():
                for df in dfs:
                    if "salary_year" not in df.columns:
                        df["salary_year"] = y
                    if "salary_month" not in df.columns:
                        df["salary_month"] = m
                    if "月份" not in df.columns:
                        df["月份"] = f"{m}月"

        # 合入全局
        for sn, dfs in asset_sheet_dfs.items():
            all_sheet_dfs.setdefault(sn, []).extend(dfs)

    if not all_sheet_dfs:
        raise HTTPException(status_code=400, detail="无法读取计算结果数据")

    # 每个 sheet 合并为一个 DataFrame
    template_sheets: dict[str, pd.DataFrame] = {}
    for sheet_name, dfs in all_sheet_dfs.items():
        template_sheets[sheet_name] = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]

    # 主数据源 "DT" = 第一个 sheet（用于 group_by 分组，向后兼容）
    first_sheet_name = next(iter(all_sheet_dfs), None)

    # dataset 用于后续 show_empty_period / first_row 等逻辑
    dataset = template_sheets.get(first_sheet_name, pd.DataFrame())

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
    now = datetime.now()   # 北京时间（容器 TZ=Asia/Shanghai），不用 utcnow 以免月初跨日错月
    # {year}/{month} 优先用薪资周期（计算历史的年月），未填才回退当前时间
    _year = str(task.salary_year) if task.salary_year else str(now.year)
    _month = f"{int(task.salary_month):02d}" if task.salary_month else f"{now.month:02d}"
    system_vars = {
        "year": _year,
        "month": _month,
        "date": now.strftime("%Y%m%d"),
        "tenant": tenant_id,
    }

    # template_data 键序很重要：DT 必须是第一个非$键（_extract_datasource 取第一个做 group_by 主数据源）
    template_data = {"DT": dataset}
    for sheet_name, df in template_sheets.items():
        if sheet_name != first_sheet_name:  # 第一个 sheet 已作为 DT，不重复
            template_data[sheet_name] = df
    template_data["$year"] = system_vars["year"]
    template_data["$month"] = system_vars["month"]
    template_data["$date"] = system_vars["date"]
    template_data["$tenant"] = tenant_id

    # 【Step 1】按列规整类型：数值列保留数值（让模板数字格式生效），其余转文本兜底 #N/A/error
    for k in list(template_data.keys()):
        v = template_data[k]
        if isinstance(v, pd.DataFrame):
            template_data[k] = _coerce_df_for_report(v)

    # 【Step 2 - 双别名】给每个 sheet 的字母列追加"原始拼接表头"为别名列
    # 模板可同时用 &=DT.A 和 &=DT.原始表头 两种写法
    if header_map:
        # DT 对应 first_sheet_name 的 header_map
        if first_sheet_name and first_sheet_name in header_map and "DT" in template_data:
            template_data["DT"] = _alias_letter_columns_with_headers(
                template_data["DT"], header_map[first_sheet_name]
            )
        for sheet_name in list(template_data.keys()):
            if sheet_name in ("DT",) or sheet_name.startswith("$"):
                continue
            if sheet_name in header_map and isinstance(template_data[sheet_name], pd.DataFrame):
                template_data[sheet_name] = _alias_letter_columns_with_headers(
                    template_data[sheet_name], header_map[sheet_name]
                )

    # 同步刷新 dataset/template_sheets 引用
    dataset = template_data.get("DT", dataset)
    for sheet_name in list(template_sheets.keys()):
        if sheet_name == first_sheet_name:
            template_sheets[sheet_name] = template_data["DT"]
        elif sheet_name in template_data:
            template_sheets[sheet_name] = template_data[sheet_name]

    # 位置别名：按 sheet 顺序追加 DT1/DT2…（第 2 个 sheet=DT1、第 3 个=DT2，以此类推；
    # 第 1 个 sheet 已是 DT）。与真实 sheet 名【并存】——老模板写 &=<sheet名>.列 仍可用，
    # 新模板可写 &=DT1.列 按位置引用，不必关心 sheet 叫什么名字。指向已规整+双别名后的数据。
    for _idx, _sn in enumerate(template_sheets.keys()):
        if _idx == 0:
            continue   # 第一个就是 DT
        _alias = f"DT{_idx}"
        if _alias not in template_data:   # 不覆盖恰好真名叫 DT1 的 sheet
            template_data[_alias] = template_sheets[_sn]
            logger.info(f"[报表] 位置别名 {_alias} -> sheet '{_sn}'")

    # 【Step 2】header_map 用于调试和模板设计预览
    if header_map:
        logger.info("=== header_map（列字母 → 原始拼接表头）===")
        for sn, m in header_map.items():
            preview = ", ".join(f"{k}={v}" for k, v in list(m.items())[:10])
            logger.info(f"[{sn}] {preview}{' ...' if len(m) > 10 else ''}")

    # 6. 用数据集第一行来解析文件名和加密规则
    #    同时把"原始表头 → 值"也加进 first_row，使 {工号}/{身份证号} 等基于中文名的规则仍可解析
    first_row = dataset.iloc[0].to_dict() if len(dataset) > 0 else {}
    if first_row and first_sheet_name:
        sheet_hm = header_map.get(first_sheet_name, {})
        for letter, original_header in sheet_hm.items():
            if letter in first_row and original_header and original_header not in first_row:
                first_row[original_header] = first_row[letter]

    if tpl.file_name_rule:
        output_name = _resolve_rule_pattern(tpl.file_name_rule, first_row, system_vars)
        if not output_name.endswith(('.xlsx', '.xls')):
            output_name += '.xlsx'
    else:
        output_name = f"报表_{tpl.name}.xlsx"

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
    split_by_field = getattr(tpl, "split_by", "") or ""

    # zip/block/sheet 模式前置校验：group_by 不能为空，且必须在数据列中
    if report_mode in ("zip", "block", "sheet"):
        if not group_by_field:
            raise HTTPException(
                status_code=400,
                detail=f"报表模式为 {report_mode}，但模版未配置分组字段(group_by)，请在模版设置中指定分组列名",
            )
        available_cols = list(dataset.columns)
        # 模糊匹配：去空格、忽略大小写；若失败再尝试通过 header_map 反查（中文表头 → 列字母）
        matched_col = None
        target = group_by_field.strip().lower()
        for col in available_cols:
            if str(col).strip().lower() == target:
                matched_col = col
                break
        if not matched_col:
            sheet_hm = header_map.get(first_sheet_name, {}) if first_sheet_name else {}
            for letter, original_header in sheet_hm.items():
                if str(original_header).strip().lower() == target and letter in available_cols:
                    matched_col = letter
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

    # 有 split_by 或 zip 模式时输出 .zip，其余输出原始扩展名
    if report_mode == "zip" or split_by_field:
        output_ext = ".zip"
        output_name_final = os.path.splitext(output_name)[0] + output_ext
    else:
        output_name_final = output_name

    output_dir = _PROJECT_ROOT / "tenants" / tenant_id / "reports"
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = now.strftime("%Y%m%d%H%M%S")
    output_path = str(output_dir / f"{timestamp}_{output_name_final}")

    logger.info(f"报表模式: {report_mode}, group_by={group_by_field}, split_by={split_by_field}, skip_rows={skip_rows_val}")

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
            split_by=split_by_field,
            # sheet 名占位符：DT/DT1… → 第 idx 个数据源 sheet 名；{year}/{month}/{date}/{tenant} 子串替换
            sheet_source_names=list(all_sheet_dfs.keys()),
            sheet_vars=system_vars,
        )
        # 实际输出路径可能和请求路径不同（如 zip 回退到 fill 时扩展名变为 .xlsx）
        output_path = actual_output_path
    except Exception as e:
        logger.error(f"报表生成失败: {e}")
        raise HTTPException(status_code=500, detail=f"报表生成失败: {str(e)}")

    # 9.5 整表搬运：把计算结果中预设的 sheet(按名称/#序号)整表拷贝追加到报表末尾。
    # 仅对单文件 xlsx 报表生效（zip 打包模式跳过）。源取当前任务的结果文件。
    _carry_cfg = (getattr(tpl, "carry_over_sheets", "") or "").strip()
    if _carry_cfg and not output_path.lower().endswith(".zip"):
        try:
            _specs = [s.strip() for s in re.split(r"[,，;；\n]+", _carry_cfg) if s.strip()]
            _carry_src = next(
                (a.file_path for a in assets
                 if a.source_task_id == task_id and a.file_path and os.path.exists(a.file_path)),
                None,
            ) or next((a.file_path for a in assets if a.file_path and os.path.exists(a.file_path)), None)
            if _carry_src and _specs:
                _n = aspose_helper.append_carryover_sheets(
                    output_path, _carry_src, _specs, password=password,
                )
                logger.info(f"[整表搬运] 追加 {_n}/{len(_specs)} 个结果 sheet 到报表: {_specs}")
            else:
                logger.warning(f"[整表搬运] 跳过：源结果文件缺失或未配置 sheet（cfg={_carry_cfg}）")
        except Exception as _carry_e:
            logger.error(f"[整表搬运] 失败（不阻断报表下载）: {_carry_e}", exc_info=True)

    # 10. 留痕 — 保存为 DataAsset
    #     磁盘文件名带时间戳前缀（防同模版同任务重复生成时互相覆盖），
    #     但用户可见的下载名/展示名去掉该前缀，保持干净（如 aaaaaa-202605.xlsx）。
    actual_filename = os.path.basename(output_path)
    download_name = actual_filename
    if len(download_name) >= 15 and download_name[:14].isdigit() and download_name[14] == "_":
        download_name = download_name[15:]
    try:
        report_asset = DataAsset(
            tenant_id=tenant_id,
            asset_type="report",
            name=f"报表_{tpl.name}_{now.strftime('%Y%m%d')}",
            file_path=output_path,
            file_name=download_name,
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
        filename=download_name,
        media_type=media,
    )


@router.get("/templates/{template_id}/headers-preview")
async def preview_template_headers(
    template_id: int,
    task_id: int = Query(..., description="计算任务 ID（取该任务的结果文件做表头预览）"),
    sheet: Optional[str] = Query(None, description="只返回指定 sheet"),
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """根据指定计算任务的结果文件，返回各 sheet 的「列字母 → 原始拼接表头」对照。
    用于模板设计者编辑模板前确认应该写 &=DT.A 还是 &=DT.B 等。
    每次实时计算，不缓存。
    """
    tpl = db.query(Template).filter(Template.id == template_id).first()
    if not tpl:
        raise HTTPException(status_code=404, detail="模版不存在")

    task = db.query(ComputeTask).filter(ComputeTask.id == task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="计算任务不存在")

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
        raise HTTPException(status_code=400, detail="该任务没有结果文件，无法预览表头")

    merged: dict = {}
    for asset in assets:
        partial: dict = {}
        if asset.parsed_data:
            partial = _build_header_map_from_parsed(asset.parsed_data)
        elif os.path.exists(asset.file_path):
            try:
                _, partial = _read_sheets_with_letter_columns(asset.file_path)
            except Exception as e:
                logger.warning(f"预览解析失败 {asset.file_path}: {e}")
                continue
        for sn, m in partial.items():
            existing = merged.setdefault(sn, {})
            for letter, header in m.items():
                existing.setdefault(letter, header)

    if sheet:
        merged = {sheet: merged.get(sheet, {})}

    sheets_out = []
    for idx, (sn, mapping) in enumerate(merged.items()):
        items = sorted(mapping.items(), key=lambda kv: (len(kv[0]), kv[0]))
        # 位置别名：第1个=DT、第2个=DT1、第3个=DT2…；sheet 真名同样可用（并存，向后兼容）
        _pos_alias = "DT" if idx == 0 else f"DT{idx}"
        sheets_out.append({
            "sheet_name": sn,
            "is_primary": idx == 0,
            "alias": _pos_alias,
            "sheet_alias": sn,   # 老写法：&=<sheet名>.列 仍可用
            "columns": [{"letter": k, "header": v} for k, v in items],
        })
    return {"template_id": template_id, "task_id": task_id, "sheets": sheets_out}


# ========================= 脚本管理 =========================

@router.get("/scripts")
async def list_scripts(
    tenant_id: Optional[str] = Query(None, description="租户ID模糊匹配"),
    include_inactive: bool = Query(False, description="是否包含已停用脚本"),
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """列出所有脚本（管理员视图）"""
    q = db.query(Script)
    if tenant_id:
        q = q.filter(Script.tenant_id.like(f"%{tenant_id}%"))
    if not include_inactive:
        q = q.filter(Script.is_active == True)
    scripts = q.order_by(Script.tenant_id, Script.name, Script.version.desc()).all()
    result = []
    for s in scripts:
        result.append({
            "id": s.id,
            "tenant_id": s.tenant_id,
            "name": s.name,
            "description": s.description or "",
            "mode": s.mode,
            "version": s.version,
            "accuracy": s.accuracy,
            "is_active": bool(s.is_active),
            "source_session_id": s.source_session_id,
            "created_at": s.created_at.isoformat() if s.created_at else None,
            "updated_at": s.updated_at.isoformat() if s.updated_at else None,
        })
    return {"total": len(result), "items": result}


@router.post("/scripts/{script_id}/disable")
async def disable_script(
    script_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """停用脚本（软删除，智训和智算不再可见）"""
    s = db.query(Script).filter(Script.id == script_id).first()
    if not s:
        raise HTTPException(status_code=404, detail="脚本不存在")
    if not s.is_active:
        return {"success": True, "message": "脚本已是停用状态", "script_id": script_id}
    s.is_active = False
    s.updated_at = datetime.utcnow()
    db.commit()
    # 同步到 FS 层禁用列表，确保 /api/tenant-scripts 也过滤掉
    try:
        from ..storage.storage_manager import StorageManager
        StorageManager().set_script_disabled(s.tenant_id, s.name, True)
    except Exception:
        pass
    return {"success": True, "message": f"已停用脚本「{s.name}」(v{s.version})", "script_id": script_id}


@router.post("/scripts/{script_id}/enable")
async def enable_script(
    script_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """恢复脚本"""
    s = db.query(Script).filter(Script.id == script_id).first()
    if not s:
        raise HTTPException(status_code=404, detail="脚本不存在")
    if s.is_active:
        return {"success": True, "message": "脚本已是启用状态", "script_id": script_id}
    # 检查同租户同名脚本是否已有启用版本，避免冲突
    existing = db.query(Script).filter(
        Script.tenant_id == s.tenant_id,
        Script.name == s.name,
        Script.is_active == True,
        Script.id != script_id,
    ).first()
    if existing:
        raise HTTPException(
            status_code=400,
            detail=f"租户「{s.tenant_id}」已存在启用中的同名脚本「{s.name}」(v{existing.version})，请先停用后再恢复"
        )
    s.is_active = True
    s.updated_at = datetime.utcnow()
    db.commit()
    # 同步从 FS 禁用列表移除
    try:
        from ..storage.storage_manager import StorageManager
        StorageManager().set_script_disabled(s.tenant_id, s.name, False)
    except Exception:
        pass
    return {"success": True, "message": f"已启用脚本「{s.name}」(v{s.version})", "script_id": script_id}


@router.delete("/scripts/{script_id}")
async def delete_script(
    script_id: int,
    db: Session = Depends(get_db),
    _admin: User = Depends(require_admin),
):
    """物理删除脚本：删 DB 行 + FS 文件（.py/_info.json），并断开历史计算任务引用。

    引用该脚本的 compute_tasks.script_id 置空（保留计算历史与结果文件），不可恢复。
    """
    s = db.query(Script).filter(Script.id == script_id).first()
    if not s:
        raise HTTPException(status_code=404, detail="脚本不存在")
    name, version, tenant_id, code = s.name, s.version, s.tenant_id, s.code
    # 断开计算任务引用（保留计算历史/结果）
    try:
        db.query(ComputeTask).filter(ComputeTask.script_id == script_id).update(
            {ComputeTask.script_id: None}, synchronize_session=False)
    except Exception as e:
        logger.warning(f"[删脚本] 断开 compute_tasks 引用失败(忽略): script_id={script_id} - {e}")
    # 删 FS 文件
    try:
        from ..storage.storage_manager import StorageManager
        StorageManager().delete_script_files_by_code(tenant_id, code or "")
    except Exception as e:
        logger.warning(f"[删脚本] 删脚本文件失败(忽略): script_id={script_id} - {e}")
    db.delete(s)
    db.commit()
    return {"success": True, "message": f"已删除脚本「{name}」(v{version})", "script_id": script_id}


@router.post("/tenant-scripts/{tenant_id}/{script_id}/disable")
async def disable_tenant_script(
    tenant_id: str,
    script_id: str,
    _admin: User = Depends(require_admin),
):
    """按租户+脚本ID禁用（覆盖 FS-only 脚本，不依赖 DB 行）"""
    from ..storage.storage_manager import StorageManager
    StorageManager().set_script_disabled(tenant_id, script_id, True)
    return {"success": True, "tenant_id": tenant_id, "script_id": script_id, "disabled": True}


@router.post("/tenant-scripts/{tenant_id}/{script_id}/enable")
async def enable_tenant_script(
    tenant_id: str,
    script_id: str,
    _admin: User = Depends(require_admin),
):
    """按租户+脚本ID恢复"""
    from ..storage.storage_manager import StorageManager
    StorageManager().set_script_disabled(tenant_id, script_id, False)
    return {"success": True, "tenant_id": tenant_id, "script_id": script_id, "disabled": False}


# ========================= 测试环境脚本迁移 =========================

import hashlib as _hashlib
import json as _json
from pydantic import BaseModel as _BaseModel


def _code_hash(code: str) -> str:
    """跨环境稳定标识：script_<md5(code)[:12]>，与文件型 script_id 一致。"""
    return "script_" + _hashlib.md5((code or "").encode("utf-8")).hexdigest()[:12]


def _test_env_conf():
    base = (os.getenv("TEST_ENV_BASE_URL") or "").strip().rstrip("/")
    user = (os.getenv("TEST_ENV_USERNAME") or "").strip()
    pwd = os.getenv("TEST_ENV_PASSWORD") or ""
    return base, user, pwd


def _remote_login(base, user, pwd):
    import requests
    resp = requests.post(f"{base}/api/auth/login",
                         json={"username": user, "password": pwd}, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _remote_get(base, token, path):
    import requests
    resp = requests.get(f"{base}{path}",
                        headers={"Authorization": f"Bearer {token}"}, timeout=180)
    if resp.status_code >= 400:
        # 把远程返回的 detail/正文带出来，避免只看到 "500 Server Error" 这类无信息异常
        raise RuntimeError(f"HTTP {resp.status_code} {path}: {resp.text[:500]}")
    try:
        return resp.json()
    except Exception:
        raise RuntimeError(f"响应非 JSON {path}: {resp.text[:200]}")


# ---- 导出端（双环境同代码都有，被对方远程拉取）----

@router.get("/migration/export-list")
def migration_export_list(db: Session = Depends(get_db), _admin: User = Depends(require_admin)):
    """列出本环境可迁移的已训练脚本（供其它环境拉取）。"""
    scripts = (db.query(Script).filter(Script.is_active == True)
               .order_by(Script.tenant_id, Script.name).all())
    items = []
    for s in scripts:
        if not s.code:
            continue
        items.append({
            "db_id": s.id, "hash": _code_hash(s.code), "name": s.name,
            "tenant_id": s.tenant_id, "mode": s.mode, "accuracy": s.accuracy,
            "version": s.version, "source_session_id": s.source_session_id,
            "has_session": s.source_session_id is not None,
            "created_at": s.created_at.isoformat() if s.created_at else None,
        })
    return {"total": len(items), "items": items}


@router.get("/migration/export/{db_id}")
def migration_export(db_id: int, db: Session = Depends(get_db), _admin: User = Depends(require_admin)):
    """导出单脚本完整包：Script + 训练会话 + 对话 + 迭代。"""
    s = db.query(Script).filter(Script.id == db_id).first()
    if not s:
        raise HTTPException(status_code=404, detail="脚本不存在")
    _sc_cols = ("name", "description", "code", "mode", "config", "manual_headers",
                "source_structure", "rules_content", "expected_structure", "accuracy", "version", "tenant_id")
    script = {c: getattr(s, c) for c in _sc_cols}
    session, messages, iterations = None, [], []
    if s.source_session_id:
        ts = db.query(TrainingSession).filter(TrainingSession.id == s.source_session_id).first()
        if ts:
            _ss_cols = ("session_key", "mode", "status", "config", "ai_provider", "salary_year",
                        "salary_month", "manual_headers", "rules_content", "source_structure",
                        "expected_structure", "total_iterations", "best_accuracy")
            session = {c: getattr(ts, c) for c in _ss_cols}
            for m in (db.query(TrainingMessage).filter_by(session_id=ts.id)
                      .order_by(TrainingMessage.created_at.asc()).all()):
                messages.append({"role": m.role, "content": m.content, "msg_type": m.msg_type,
                                 "metadata": m.metadata_,
                                 "created_at": m.created_at.isoformat() if m.created_at else None})
            for it in (db.query(TrainingIteration).filter_by(session_id=ts.id)
                       .order_by(TrainingIteration.iteration_num.asc()).all()):
                iterations.append({c: getattr(it, c) for c in (
                    "iteration_num", "status", "prompt_text", "ai_response", "generated_code",
                    "execution_result", "accuracy", "error_details", "duration_seconds")})
    # 模板文件一并打包：模板模式脚本的模板文件跨环境不可复用（烘焙的是训练机绝对路径，
    # 且文件名带源 session id 前缀，如 13_太保上海.xlsx）。这里把文件字节 base64 塞进包，
    # 导入端按【原始烘焙文件名】落到目标租户 templates/ 下——resolve_template_path 按
    # 【文件名+哈希】在租户目录递归查找，与两边 session id 是否一致无关。
    template_blob = None
    try:
        from ..utils.template_resolver import extract_template_ref
        _tname, _thash, _tbaked = extract_template_ref(s.code)
        # 读取用的实际路径：优先会话 config 里的持久化路径，回退烘焙绝对路径（均为本机路径）
        _tpl_file = None
        if session and isinstance(session.get("config"), dict):
            _cfg_tp = session["config"].get("template_path")
            if _cfg_tp and os.path.exists(_cfg_tp):
                _tpl_file = _cfg_tp
        if not _tpl_file and _tbaked and os.path.exists(_tbaked):
            _tpl_file = _tbaked
        # 跨环境兜底：config/烘焙路径都是外机绝对路径时（如导出机是 Windows、烘焙的是 Docker
        # /app/... 路径），按【文件名+哈希】在当前环境的租户目录/全局资源里递归定位真实文件，
        # 确保迁移包能把模版带上（否则 template_blob 为空，导入端无模版可落盘）。
        if not _tpl_file:
            try:
                from ..utils.template_resolver import resolve_template_path
                _hit = resolve_template_path(
                    tenant_id=s.tenant_id, script_code=s.code, project_root=_PROJECT_ROOT)
                if _hit and os.path.exists(_hit):
                    _tpl_file = _hit
            except Exception as _rte:
                logging.getLogger(__name__).warning(
                    f"[迁移导出] 模板解析器兜底失败 db_id={s.id}: {_rte}")
        if _tpl_file and os.path.exists(_tpl_file):
            _sz = os.path.getsize(_tpl_file)
            if _sz <= 60 * 1024 * 1024:   # 60MB 上限，超大不打包（避免撑爆 JSON 传输）
                import base64 as _b64
                with open(_tpl_file, "rb") as _tf:
                    _tbytes = _tf.read()
                template_blob = {
                    "name": _tname or os.path.basename(_tpl_file),
                    "hash": _thash or _hashlib.md5(_tbytes).hexdigest(),
                    "data_b64": _b64.b64encode(_tbytes).decode("ascii"),
                }
            else:
                logging.getLogger(__name__).warning(
                    f"[迁移导出] 模板过大({_sz}B)未打包，导入端需手动上传 db_id={s.id}")
    except Exception as _te:
        logging.getLogger(__name__).warning(f"[迁移导出] 模板打包失败 db_id={db_id}: {_te}")

    return {"hash": _code_hash(s.code), "source_db_id": s.id, "script": script,
            "session": session, "messages": messages, "iterations": iterations,
            "template": template_blob}


# ---- 导入端（当前环境，UI 调用）----

class _MigrateItem(_BaseModel):
    db_id: int
    hash: str
    name: Optional[str] = ""
    tenant_id: Optional[str] = ""     # 来源租户（空目标时用作迁入租户）
    new_name: Optional[str] = ""      # 迁入后自定义脚本名（空 = 沿用源脚本名）


class _MigrateReq(_BaseModel):
    target_tenant_id: str = ""        # 空 = 沿用各脚本来源租户（自动创建）
    items: List[_MigrateItem]
    overwrite: bool = False


@router.get("/migration/remote-scripts")
def migration_remote_scripts(
    target_tenant: str = Query("", description="迁入的目标租户；空=沿用各脚本来源租户"),
    db: Session = Depends(get_db), _admin: User = Depends(require_admin),
):
    """登录测试环境→拉取其已训练脚本列表→标注已迁移/已存在。

    target_tenant 为空时，每个脚本的迁入目标即其来源租户，逐脚本按各自租户判断状态。
    """
    base, user, pwd = _test_env_conf()
    if not (base and user and pwd):
        raise HTTPException(status_code=400, detail="未配置测试环境（TEST_ENV_BASE_URL/USERNAME/PASSWORD）")
    try:
        token = _remote_login(base, user, pwd)
        data = _remote_get(base, token, "/api/admin/migration/export-list")
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"连接测试环境失败: {e}")
    items = data.get("items", [])
    fixed = (target_tenant or "").strip()
    # 本环境已存在的租户集合（目录 ∪ DB 引用），用于“撞名”预警
    existing_tenants = set()
    _tdir = Path(__file__).resolve().parent.parent.parent / "tenants"
    if _tdir.exists():
        existing_tenants.update(d.name for d in _tdir.iterdir() if d.is_dir())
    # 按租户建立“已迁移哈希”“已存在哈希”映射，兼容固定目标与沿用来源两种模式
    migrated_by_tenant = {}
    for r in db.query(ScriptMigration).all():
        migrated_by_tenant.setdefault(r.target_tenant_id, set()).add(r.source_script_hash)
    local_by_tenant = {}
    for s in db.query(Script).filter_by(is_active=True).all():
        existing_tenants.add(s.tenant_id)
        if s.code:
            local_by_tenant.setdefault(s.tenant_id, set()).add(_code_hash(s.code))
    for it in items:
        dest = fixed or (it.get("tenant_id") or "")
        h = it.get("hash")
        it["dest_tenant"] = dest
        it["dest_tenant_exists"] = dest in existing_tenants
        it["already_migrated"] = h in migrated_by_tenant.get(dest, set())
        it["exists_by_hash"] = h in local_by_tenant.get(dest, set())
    return {"total": len(items), "items": items, "source_url": base,
            "use_source_tenant": not fixed}


@router.post("/migration/import")
def migration_import(req: _MigrateReq, db: Session = Depends(get_db), _admin: User = Depends(require_admin)):
    """把选中的测试环境脚本+训练记录迁入当前环境的目标租户。"""
    base, user, pwd = _test_env_conf()
    if not (base and user and pwd):
        raise HTTPException(status_code=400, detail="未配置测试环境")
    try:
        token = _remote_login(base, user, pwd)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"测试环境登录失败: {e}")

    fixed_tenant = (req.target_tenant_id or "").strip()   # 空 = 沿用各脚本来源租户
    # (租户, 代码哈希) -> Script，覆盖判定按“各自租户”进行
    local_map = {}
    for s in db.query(Script).filter_by(is_active=True).all():
        if s.code:
            local_map.setdefault((s.tenant_id, _code_hash(s.code)), s)

    from ..storage.storage_manager import StorageManager
    sm = StorageManager()
    imported, conflicts, skipped = [], [], []

    for item in req.items:
        tenant = fixed_tenant or (item.tenant_id or "").strip()
        if not tenant:
            skipped.append({"name": item.name, "reason": "无法确定迁入租户（来源租户为空）"})
            continue
        existing = local_map.get((tenant, item.hash))
        if existing and not req.overwrite:
            conflicts.append({"db_id": item.db_id, "hash": item.hash,
                              "name": item.name, "tenant_id": tenant})
            continue
        # 迁移前先确保租户目录建成；失败则跳过该脚本，避免 DB 有行而文件系统无租户
        try:
            scripts_dir = sm.get_tenant_dir(tenant) / "scripts"
            scripts_dir.mkdir(parents=True, exist_ok=True)
        except Exception as de:
            skipped.append({"name": item.name, "reason": f"创建租户目录失败: {de}"})
            continue
        try:
            bundle = _remote_get(base, token, f"/api/admin/migration/export/{item.db_id}")
        except Exception as e:
            logging.getLogger(__name__).exception(
                f"[迁移] 拉取脚本包失败 db_id={item.db_id} name={item.name}")
            skipped.append({"name": item.name, "reason": f"拉取失败: {e}"})
            continue
        sc = bundle.get("script") or {}
        code = sc.get("code") or ""
        if not code:
            skipped.append({"name": item.name, "reason": "无代码"})
            continue

        # 1) 训练会话 + 对话 + 迭代
        new_session_id = None
        sess = bundle.get("session")
        if sess:
            skey = sess.get("session_key") or f"mig_{item.hash}"
            if db.query(TrainingSession).filter_by(session_key=skey).first():
                skey = f"{skey}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            ts = TrainingSession(
                tenant_id=tenant, session_key=skey,
                mode=sess.get("mode") or sc.get("mode") or "formula",
                status=sess.get("status") or "completed", config=sess.get("config"),
                ai_provider=sess.get("ai_provider"), salary_year=sess.get("salary_year"),
                salary_month=sess.get("salary_month"), manual_headers=sess.get("manual_headers"),
                rules_content=sess.get("rules_content"), source_structure=sess.get("source_structure"),
                expected_structure=sess.get("expected_structure"),
                total_iterations=sess.get("total_iterations") or 0, best_accuracy=sess.get("best_accuracy"),
            )
            db.add(ts)
            db.flush()
            new_session_id = ts.id
            for m in bundle.get("messages", []):
                db.add(TrainingMessage(
                    session_id=ts.id, role=m.get("role") or "user", content=m.get("content") or "",
                    msg_type=m.get("msg_type") or "chat", metadata_=m.get("metadata")))
            for itr in bundle.get("iterations", []):
                db.add(TrainingIteration(
                    session_id=ts.id, iteration_num=itr.get("iteration_num") or 0,
                    status=itr.get("status") or "completed", prompt_text=itr.get("prompt_text"),
                    ai_response=itr.get("ai_response"), generated_code=itr.get("generated_code"),
                    execution_result=itr.get("execution_result"), accuracy=itr.get("accuracy"),
                    error_details=itr.get("error_details"), duration_seconds=itr.get("duration_seconds")))

        # 迁入后脚本名：优先用户指定的 new_name，否则沿用源脚本名
        _new_name = (item.new_name or "").strip()

        # 2) Script（覆盖或新建）
        if existing:
            tgt = existing
            tgt.code = code
            tgt.name = _new_name or sc.get("name") or existing.name
            tgt.description = sc.get("description") or ""
            tgt.mode = sc.get("mode") or "formula"
            tgt.config = sc.get("config")
            tgt.manual_headers = sc.get("manual_headers")
            tgt.source_structure = sc.get("source_structure")
            tgt.rules_content = sc.get("rules_content")
            tgt.expected_structure = sc.get("expected_structure")
            tgt.accuracy = sc.get("accuracy")
            tgt.source_session_id = new_session_id
            tgt.is_active = True
            tgt.updated_at = datetime.utcnow()
        else:
            tgt = Script(
                tenant_id=tenant, name=_new_name or sc.get("name") or item.hash, description=sc.get("description") or "",
                code=code, mode=sc.get("mode") or "formula", config=sc.get("config"),
                manual_headers=sc.get("manual_headers"), source_structure=sc.get("source_structure"),
                rules_content=sc.get("rules_content"), expected_structure=sc.get("expected_structure"),
                accuracy=sc.get("accuracy"), source_session_id=new_session_id, version=1, is_active=True)
            db.add(tgt)
        db.flush()
        if new_session_id:
            ts2 = db.query(TrainingSession).filter_by(id=new_session_id).first()
            if ts2:
                ts2.final_script_id = tgt.id

        # 3) 落盘 scripts/script_<hash>.py + _info.json（scripts_dir 已在前面建好）
        try:
            (scripts_dir / f"{item.hash}.py").write_text(code, encoding="utf-8")
            info = {"script_id": item.hash, "tenant_id": tenant, "name": tgt.name,
                    "score": sc.get("accuracy"), "migrated_from": base}
            (scripts_dir / f"{item.hash}_info.json").write_text(
                _json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as fe:
            logging.getLogger(__name__).warning(f"[迁移] 落盘失败: {fe}")

        # 3b) 模板文件落盘：把导出包里的模板字节按【原始烘焙文件名】写到目标租户 templates/。
        # 文件名保持源环境烘焙的名字（含源 session id 前缀），resolve_template_path 按
        # 【文件名+哈希】在租户目录递归命中，故两边 session id 不一致也能定位到，迁移后
        # 无需手动上传模板即可直接智算。仅模板模式脚本有 template 包。
        tpl_blob = bundle.get("template")
        if tpl_blob and tpl_blob.get("data_b64") and tpl_blob.get("name"):
            try:
                import base64 as _b64
                templates_dir = sm.get_tenant_dir(tenant) / "templates"
                templates_dir.mkdir(parents=True, exist_ok=True)
                _tpl_dst = templates_dir / tpl_blob["name"]
                _tpl_bytes = _b64.b64decode(tpl_blob["data_b64"])
                _tpl_dst.write_bytes(_tpl_bytes)
                # 会话 config 指向新路径，训练/复算再跑也能定位（智算另有 resolver 兜底）
                if new_session_id:
                    _ts = db.query(TrainingSession).filter_by(id=new_session_id).first()
                    if _ts:
                        _c = dict(_ts.config) if _ts.config else {}
                        _c["template_path"] = str(_tpl_dst)
                        _ts.config = _c
                        from sqlalchemy.orm.attributes import flag_modified
                        flag_modified(_ts, "config")
                logging.getLogger(__name__).info(
                    f"[迁移] 模板落盘 -> {_tpl_dst} ({len(_tpl_bytes)}B)")
            except Exception as _tre:
                logging.getLogger(__name__).warning(f"[迁移] 模板落盘失败: {_tre}")

        # 4) 迁移记录
        db.add(ScriptMigration(source_url=base, source_script_hash=item.hash, source_db_id=item.db_id,
                               target_tenant_id=tenant, target_script_id=tgt.id, name=tgt.name))
        imported.append({"name": tgt.name, "hash": item.hash,
                         "tenant_id": tenant, "overwritten": bool(existing)})

    db.commit()
    return {"imported": imported, "conflicts": conflicts, "skipped": skipped}

