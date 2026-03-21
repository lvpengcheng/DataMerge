"""
数据库连接模块
支持 PostgreSQL / MySQL / SQLite
"""

import os
import json
from functools import partial
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./data.db")

# JSON 序列化：中文不转义
_json_serializer = partial(json.dumps, ensure_ascii=False, default=str)

# 根据数据库类型配置连接参数
if DATABASE_URL.startswith("postgresql") or DATABASE_URL.startswith("mysql"):
    engine = create_engine(
        DATABASE_URL,
        pool_size=10,
        max_overflow=20,
        pool_pre_ping=True,
        pool_recycle=3600,
        echo=False,
        json_serializer=_json_serializer,
    )
else:
    engine = create_engine(DATABASE_URL, echo=False, json_serializer=_json_serializer)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()


def get_db():
    """FastAPI 依赖：获取数据库会话"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
