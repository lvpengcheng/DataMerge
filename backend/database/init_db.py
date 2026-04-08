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


def ensure_database():
    """自动创建数据库（如果不存在），支持 PostgreSQL 和 MySQL"""
    db_url = os.getenv("DATABASE_URL", "")

    if db_url.startswith("postgresql"):
        _ensure_postgresql_database(db_url)
    elif db_url.startswith("mysql"):
        _ensure_mysql_database(db_url)


def _ensure_postgresql_database(db_url):
    """PostgreSQL: 连接到默认 postgres 库，创建目标数据库"""
    from urllib.parse import urlparse
    parsed = urlparse(db_url)
    db_name = parsed.path.lstrip("/").split("?")[0]

    try:
        import psycopg2
        from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

        conn = psycopg2.connect(
            host=parsed.hostname or "localhost",
            port=parsed.port or 5432,
            user=parsed.username,
            password=parsed.password,
            dbname="postgres",
        )
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        cursor = conn.cursor()

        # 检查数据库是否存在
        cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (db_name,))
        if not cursor.fetchone():
            cursor.execute(f'CREATE DATABASE "{db_name}" ENCODING \'UTF8\'')
            print(f"      数据库 '{db_name}' 创建成功。")
        else:
            print(f"      数据库 '{db_name}' 已存在。")

        cursor.close()
        conn.close()
    except Exception as e:
        print(f"      警告: 无法自动创建数据库: {e}")


def _ensure_mysql_database(db_url):
    """MySQL: 连接到服务器，创建目标数据库"""
    from urllib.parse import urlparse
    parsed = urlparse(db_url)
    db_name = parsed.path.lstrip("/").split("?")[0]

    try:
        import pymysql
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


ensure_database()

from backend.database.connection import engine, SessionLocal, Base
from backend.database.models import (
    Role, Organization, User, TenantAuthorization,
    ReferenceCategory, DataAsset,
    TrainingSession, TrainingIteration, TrainingMessage, Script,
    ComputeTask, ComputeTaskInput, RuleSession,
)
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


def init_database():
    """创建所有表"""
    print("[1/6] 创建数据库表...")
    Base.metadata.create_all(bind=engine)
    print("      表创建完成。")

    # 增量迁移：为已有表添加新列
    _migrate_add_columns()


def _migrate_add_columns():
    """安全地为已有表添加新列（列已存在时跳过）"""
    from sqlalchemy import text, inspect
    insp = inspect(engine)

    migrations = [
        # ---------- compute_tasks ----------
        ("compute_tasks", "salary_year", "INTEGER"),
        ("compute_tasks", "salary_month", "INTEGER"),
        # ---------- templates ----------
        ("templates", "file_name_rule", "VARCHAR(500) DEFAULT ''"),
        ("templates", "encrypt_type", "VARCHAR(20) DEFAULT 'none'"),
        ("templates", "encrypt_password", "VARCHAR(100) DEFAULT ''"),
        ("templates", "report_mode", "VARCHAR(20) DEFAULT 'fill'"),
        ("templates", "group_by", "VARCHAR(100) DEFAULT ''"),
        ("templates", "skip_rows", "INTEGER DEFAULT 1"),
        ("templates", "name_field", "VARCHAR(100) DEFAULT ''"),
        ("templates", "show_empty_period", "BOOLEAN DEFAULT TRUE"),
        ("templates", "split_by", "VARCHAR(100) DEFAULT ''"),
        # ---------- training_sessions ----------
        ("training_sessions", "ai_provider", "VARCHAR(50)"),
        ("training_sessions", "salary_year", "INTEGER"),
        ("training_sessions", "salary_month", "INTEGER"),
        ("training_sessions", "manual_headers", "TEXT"),
        ("training_sessions", "rules_content", "TEXT"),
        ("training_sessions", "source_structure", "TEXT"),
        ("training_sessions", "expected_structure", "TEXT"),
        # ---------- scripts ----------
        ("scripts", "manual_headers", "TEXT"),
        ("scripts", "source_structure", "TEXT"),
        ("scripts", "rules_content", "TEXT"),
        ("scripts", "expected_structure", "TEXT"),
    ]

    with engine.connect() as conn:
        for table, column, col_type in migrations:
            if table not in insp.get_table_names():
                continue
            existing = [c["name"] for c in insp.get_columns(table)]
            if column not in existing:
                conn.execute(text(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}"))
                print(f"      迁移: {table}.{column} 已添加")
        conn.commit()


def seed_roles(db):
    """种子角色数据"""
    print("[2/6] 初始化角色...")
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
    print("[3/6] 初始化默认组织...")
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
    print("[4/6] 初始化管理员账号...")
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
    print("[5/6] 迁移现有租户...")
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


def seed_reference_categories(db):
    """种子基础数据分类"""
    print("[6/6] 初始化基础数据分类...")
    categories = [
        {"code": "min_wage", "name": "最低工资标准", "description": "各地区最低工资标准表", "scope": "global", "sort_order": 1},
        {"code": "salary_calendar", "name": "薪资日历", "description": "工作日/节假日/调休日历", "scope": "global", "sort_order": 2},
        {"code": "social_insurance", "name": "社保基数", "description": "社会保险缴纳基数与比例", "scope": "global", "sort_order": 3},
        {"code": "housing_fund", "name": "公积金基数", "description": "住房公积金缴纳基数与比例", "scope": "global", "sort_order": 4},
        {"code": "tax_bracket", "name": "个税税率", "description": "个人所得税累进税率表", "scope": "global", "sort_order": 5},
        {"code": "allowance", "name": "津贴补贴标准", "description": "各类津贴补贴标准表", "scope": "tenant", "sort_order": 6},
        {"code": "position_salary", "name": "岗位薪资标准", "description": "各岗位级别薪资标准", "scope": "tenant", "sort_order": 7},
        {"code": "other", "name": "其他基础数据", "description": "其他业务基础数据", "scope": "tenant", "sort_order": 99},
    ]
    for cat_data in categories:
        existing = db.query(ReferenceCategory).filter_by(code=cat_data["code"]).first()
        if not existing:
            db.add(ReferenceCategory(**cat_data))
            print(f"      创建分类: {cat_data['name']}")
        else:
            print(f"      分类已存在: {cat_data['name']}")
    db.commit()


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
        seed_reference_categories(db)
    finally:
        db.close()

    print("=" * 50)
    print("初始化完成！")
    print("默认管理员: admin / admin123")
    print("=" * 50)


if __name__ == "__main__":
    main()
