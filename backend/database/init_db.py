"""
数据库初始化脚本 - 建表 + 种子数据
运行: python -m backend.database.init_db
"""

import os
import sys
from pathlib import Path

# 确保项目根目录在 sys.path 中
project_root = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv()


def ensure_mysql_database():
    """如果使用 MySQL，自动创建数据库（如果不存在）"""
    db_url = os.getenv("DATABASE_URL", "")
    if not db_url.startswith("mysql"):
        return

    # 从 URL 中提取数据库名，连接到 MySQL 服务器（不指定数据库）
    from urllib.parse import urlparse
    parsed = urlparse(db_url)
    db_name = parsed.path.lstrip("/").split("?")[0]
    server_url = db_url.replace(f"/{db_name}", "/", 1).split("?")[0]

    try:
        import pymysql
        # 从 parsed URL 中提取连接参数
        conn = pymysql.connect(
            host=parsed.hostname or "localhost",
            port=parsed.port or 3306,
            user=parsed.username,
            password=parsed.password,
        )
        cursor = conn.cursor()
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        conn.commit()
        cursor.close()
        conn.close()
        print(f"      数据库 '{db_name}' 已确认存在。")
    except Exception as e:
        print(f"      警告: 无法自动创建数据库: {e}")


ensure_mysql_database()

from backend.database.connection import engine, SessionLocal, Base
from backend.database.models import Role, Organization, User, TenantAuthorization
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


def init_database():
    """创建所有表"""
    print("[1/4] 创建数据库表...")
    Base.metadata.create_all(bind=engine)
    print("      表创建完成。")


def seed_roles(db):
    """种子角色数据"""
    print("[2/4] 初始化角色...")
    roles = [
        {
            "name": "admin",
            "description": "系统管理员",
            "permissions": {
                "admin": True,
                "can_train": True,
                "can_compute": True,
                "can_manage_users": True,
            },
            "is_system": True,
        },
        {
            "name": "user",
            "description": "普通用户",
            "permissions": {
                "can_train": True,
                "can_compute": True,
            },
            "is_system": True,
        },
        {
            "name": "readonly",
            "description": "只读用户",
            "permissions": {
                "can_view": True,
            },
            "is_system": True,
        },
    ]
    for role_data in roles:
        existing = db.query(Role).filter_by(name=role_data["name"]).first()
        if not existing:
            db.add(Role(**role_data))
            print(f"      创建角色: {role_data['name']}")
        else:
            print(f"      角色已存在: {role_data['name']}")
    db.commit()


def seed_default_org(db):
    """种子默认组织"""
    print("[3/4] 初始化默认组织...")
    org = db.query(Organization).filter_by(name="默认组织").first()
    if not org:
        org = Organization(name="默认组织", description="系统默认组织")
        db.add(org)
        db.commit()
        db.refresh(org)
        print(f"      创建默认组织 (id={org.id})")
    else:
        print(f"      默认组织已存在 (id={org.id})")
    return org


def seed_admin_user(db, org):
    """种子管理员用户"""
    print("[4/4] 初始化管理员账号...")
    admin_role = db.query(Role).filter_by(name="admin").first()
    admin_user = db.query(User).filter_by(username="admin").first()
    if not admin_user:
        admin_user = User(
            username="admin",
            password_hash=pwd_context.hash("admin123"),
            display_name="系统管理员",
            org_id=org.id,
            role_id=admin_role.id,
            is_active=True,
        )
        db.add(admin_user)
        db.commit()
        print("      创建管理员: admin / admin123")
    else:
        print(f"      管理员已存在: {admin_user.username}")
    return admin_user


def migrate_existing_tenants(db, org, admin_user):
    """将现有 tenants 目录下的租户自动关联到默认组织"""
    tenants_dir = project_root / "tenants"
    if not tenants_dir.exists():
        print("      tenants 目录不存在，跳过迁移。")
        return
    tenant_dirs = [d.name for d in tenants_dir.iterdir() if d.is_dir()]
    migrated = 0
    for tenant_id in tenant_dirs:
        existing = (
            db.query(TenantAuthorization)
            .filter_by(tenant_id=tenant_id, org_id=org.id)
            .first()
        )
        if not existing:
            auth = TenantAuthorization(
                tenant_id=tenant_id,
                org_id=org.id,
                auth_type="owner",
                granted_by=admin_user.id,
            )
            db.add(auth)
            migrated += 1
    db.commit()
    print(f"      迁移了 {migrated} 个现有租户到默认组织。")


def main():
    print("=" * 50)
    print("数据整合平台 - 数据库初始化")
    print("=" * 50)

    init_database()

    db = SessionLocal()
    try:
        seed_roles(db)
        org = seed_default_org(db)
        admin_user = seed_admin_user(db, org)
        migrate_existing_tenants(db, org, admin_user)
    finally:
        db.close()

    print("=" * 50)
    print("初始化完成！")
    print("默认管理员: admin / admin123")
    print("=" * 50)


if __name__ == "__main__":
    main()
