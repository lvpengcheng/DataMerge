"""
认证相关 Pydantic 模型
"""

from typing import Optional, List, Dict, Any
from pydantic import BaseModel


# ---- Auth ----
class LoginRequest(BaseModel):
    username: str
    password: str


class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: "UserResponse"


class ChangePasswordRequest(BaseModel):
    old_password: str
    new_password: str


# ---- User ----
class UserCreate(BaseModel):
    username: str
    password: str
    display_name: Optional[str] = ""
    email: Optional[str] = ""
    phone: Optional[str] = ""
    org_id: Optional[int] = None
    role_id: Optional[int] = None


class UserUpdate(BaseModel):
    display_name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    org_id: Optional[int] = None
    role_id: Optional[int] = None
    is_active: Optional[bool] = None


class UserResponse(BaseModel):
    id: int
    username: str
    display_name: str
    email: str
    phone: str
    org_id: Optional[int]
    org_name: Optional[str] = ""
    role_id: Optional[int]
    role_name: Optional[str] = ""
    is_active: bool

    class Config:
        from_attributes = True


# ---- Role ----
class RoleCreate(BaseModel):
    name: str
    description: Optional[str] = ""
    permissions: Optional[Dict[str, Any]] = {}


class RoleUpdate(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    permissions: Optional[Dict[str, Any]] = None


class RoleResponse(BaseModel):
    id: int
    name: str
    description: str
    permissions: Dict[str, Any]
    is_system: bool

    class Config:
        from_attributes = True


# ---- Organization ----
class OrgCreate(BaseModel):
    name: str
    parent_id: Optional[int] = None
    description: Optional[str] = ""


class OrgUpdate(BaseModel):
    name: Optional[str] = None
    parent_id: Optional[int] = None
    description: Optional[str] = None
    is_active: Optional[bool] = None


class OrgResponse(BaseModel):
    id: int
    name: str
    parent_id: Optional[int]
    description: str
    is_active: bool
    children: Optional[List["OrgResponse"]] = []

    class Config:
        from_attributes = True


# ---- Tenant Authorization ----
class TenantAuthCreate(BaseModel):
    tenant_id: str
    org_id: int
    auth_type: str = "shared"


class TenantAuthResponse(BaseModel):
    id: int
    tenant_id: str
    org_id: int
    org_name: Optional[str] = ""
    auth_type: str
    granted_by: Optional[int]
    granted_at: Optional[str]
    revoked_at: Optional[str] = None

    class Config:
        from_attributes = True


# ---- Template ----
class TemplateCreate(BaseModel):
    tenant_id: Optional[str] = None
    name: str
    description: Optional[str] = ""
    file_name_rule: Optional[str] = ""
    encrypt_type: Optional[str] = "none"
    encrypt_password: Optional[str] = ""


class TemplateUpdate(BaseModel):
    tenant_id: Optional[str] = None
    name: Optional[str] = None
    description: Optional[str] = None
    file_name_rule: Optional[str] = None
    encrypt_type: Optional[str] = None
    encrypt_password: Optional[str] = None


class TemplateResponse(BaseModel):
    id: int
    tenant_id: Optional[str]
    name: str
    description: str
    file_name: str
    file_name_rule: str
    encrypt_type: str
    is_active: bool
    created_by: Optional[int]
    creator_name: Optional[str] = ""
    created_at: Optional[str]
    updated_at: Optional[str]

    class Config:
        from_attributes = True


# 更新前向引用
LoginResponse.model_rebuild()
OrgResponse.model_rebuild()
