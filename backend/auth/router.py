"""
认证路由 - /api/auth/*
"""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from ..database.connection import get_db
from ..database.models import User
from .schemas import LoginRequest, LoginResponse, ChangePasswordRequest, UserResponse
from .utils import verify_password, get_password_hash, create_access_token
from .dependencies import get_current_user

router = APIRouter(prefix="/api/auth", tags=["auth"])


def _build_user_response(user: User) -> UserResponse:
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


@router.post("/login", response_model=LoginResponse)
async def login(req: LoginRequest, db: Session = Depends(get_db)):
    """用户登录"""
    user = db.query(User).filter(User.username == req.username).first()
    if not user or not verify_password(req.password, user.password_hash):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="用户名或密码错误")
    if not user.is_active:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="账号已被禁用")

    token = create_access_token({
        "sub": user.username,
        "user_id": user.id,
        "org_id": user.org_id,
        "role": user.role.name if user.role else "",
    })

    return LoginResponse(
        access_token=token,
        user=_build_user_response(user),
    )


@router.post("/logout")
async def logout():
    """登出（客户端删除 token 即可）"""
    return {"message": "已登出"}


@router.get("/me", response_model=UserResponse)
async def get_me(current_user: User = Depends(get_current_user)):
    """获取当前用户信息"""
    return _build_user_response(current_user)


@router.post("/change-password")
async def change_password(
    req: ChangePasswordRequest,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """修改密码"""
    if not verify_password(req.old_password, current_user.password_hash):
        raise HTTPException(status_code=400, detail="原密码错误")

    current_user.password_hash = get_password_hash(req.new_password)
    db.commit()
    return {"message": "密码修改成功"}
