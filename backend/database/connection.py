"""
数据库连接模块
支持 PostgreSQL / MySQL / SQLite
"""

import os
import json
from functools import partial
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import make_url
from sqlalchemy.orm import sessionmaker, declarative_base

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./data.db")


def _bounded_int(name: str, default: int, minimum: int = 1, maximum: int = 120) -> int:
    try:
        return max(minimum, min(int(os.getenv(name, str(default))), maximum))
    except (TypeError, ValueError):
        return default


DATABASE_CONNECT_TIMEOUT = _bounded_int("DATABASE_CONNECT_TIMEOUT", 5)


def database_endpoint() -> str:
    """Return a credential-free endpoint for diagnostics."""
    url = make_url(DATABASE_URL)
    if url.get_backend_name() == "sqlite":
        return f"SQLite {url.database or './data.db'}"
    return f"{url.get_backend_name()} {url.host or 'localhost'}:{url.port or 'default'}/{url.database or ''}"

# JSON 序列化：中文不转义
_json_serializer = partial(json.dumps, ensure_ascii=False, default=str)

# 根据数据库类型配置连接参数
if DATABASE_URL.startswith("postgresql") or DATABASE_URL.startswith("mysql"):
    _connect_args = {"connect_timeout": DATABASE_CONNECT_TIMEOUT}
    DATABASE_CONNECT_ARGS = dict(_connect_args)
    engine = create_engine(
        DATABASE_URL,
        pool_size=10,
        max_overflow=20,
        pool_pre_ping=True,
        pool_recycle=3600,
        pool_timeout=_bounded_int("DATABASE_POOL_TIMEOUT", 10),
        connect_args=_connect_args,
        echo=False,
        json_serializer=_json_serializer,
    )
else:
    DATABASE_CONNECT_ARGS = {}
    engine = create_engine(DATABASE_URL, echo=False, json_serializer=_json_serializer)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()


def verify_database_connection() -> None:
    """Fail quickly with a safe, actionable error before application startup."""
    try:
        with engine.connect() as connection:
            connection.execute(text("SELECT 1"))
    except Exception as exc:
        message = str(exc).casefold()
        if "timeout" in message or "timed out" in message:
            reason = "连接超时：目标主机或端口不可达"
        elif "password authentication failed" in message or "access denied" in message:
            reason = "数据库账号或密码校验失败"
        elif "does not exist" in message and "database" in message:
            reason = "数据库不存在"
        elif "connection refused" in message or "actively refused" in message:
            reason = "目标主机可达，但数据库端口未监听或被防火墙拒绝"
        else:
            reason = f"{type(exc).__name__}: 数据库连接失败"
        raise RuntimeError(
            f"无法启动：{reason}；连接目标 {database_endpoint()}。"
            "请检查数据库服务器、虚拟交换机/网卡、防火墙和 DATABASE_URL。"
        ) from exc


def get_db():
    """FastAPI 依赖：获取数据库会话"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
