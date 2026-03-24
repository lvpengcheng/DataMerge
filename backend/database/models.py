"""
数据库 ORM 模型
"""

from datetime import datetime
from sqlalchemy import (
    Column, Integer, BigInteger, String, Boolean, DateTime, Date, Text,
    Float, ForeignKey, JSON, UniqueConstraint, Index
)
from sqlalchemy.orm import relationship
from .connection import Base


class Role(Base):
    __tablename__ = "roles"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(50), unique=True, nullable=False)
    description = Column(String(200), default="")
    permissions = Column(JSON, default=dict)
    is_system = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)

    users = relationship("User", back_populates="role")


class Organization(Base):
    __tablename__ = "organizations"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), nullable=False)
    parent_id = Column(Integer, ForeignKey("organizations.id"), nullable=True)
    description = Column(String(500), default="")
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    parent = relationship("Organization", remote_side=[id], backref="children")
    users = relationship("User", back_populates="organization")
    tenant_authorizations = relationship("TenantAuthorization", back_populates="organization")


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    display_name = Column(String(100), default="")
    email = Column(String(100), default="")
    phone = Column(String(20), default="")
    org_id = Column(Integer, ForeignKey("organizations.id"), nullable=True)
    role_id = Column(Integer, ForeignKey("roles.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    organization = relationship("Organization", back_populates="users")
    role = relationship("Role", back_populates="users")


class TenantAuthorization(Base):
    __tablename__ = "tenant_authorizations"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=False, index=True)
    org_id = Column(Integer, ForeignKey("organizations.id"), nullable=False)
    auth_type = Column(String(20), nullable=False, default="owner")  # owner / shared
    granted_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    granted_at = Column(DateTime, default=datetime.utcnow)
    revoked_at = Column(DateTime, nullable=True)

    organization = relationship("Organization", back_populates="tenant_authorizations")
    granter = relationship("User", foreign_keys=[granted_by])

    __table_args__ = (
        UniqueConstraint("tenant_id", "org_id", name="uq_tenant_org"),
    )


# ==================== 模版管理 ====================

class Template(Base):
    __tablename__ = "templates"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=True, index=True)          # NULL = 全局模版
    name = Column(String(200), nullable=False)                           # 模版名称
    description = Column(Text, default="")                               # 描述
    file_path = Column(String(500), nullable=False)                      # 模版文件物理路径
    file_name = Column(String(200), nullable=False)                      # 原始上传文件名
    file_name_rule = Column(String(500), default="")                     # 输出文件命名规则
    encrypt_type = Column(String(20), default="none")                    # none / password / write_protect
    encrypt_password = Column(String(100), default="")                   # 加密密码
    is_active = Column(Boolean, default=True)
    created_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    creator = relationship("User", foreign_keys=[created_by])


# ==================== 基础数据分类 ====================

class ReferenceCategory(Base):
    __tablename__ = "reference_categories"

    id = Column(Integer, primary_key=True, autoincrement=True)
    code = Column(String(50), unique=True, nullable=False)          # min_wage / salary_calendar / ...
    name = Column(String(100), nullable=False)                      # 最低工资标准
    description = Column(String(500), default="")
    scope = Column(String(20), default="global")                    # global / tenant
    sort_order = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)

    assets = relationship("DataAsset", back_populates="category_rel")


# ==================== 数据资产 ====================

class DataAsset(Base):
    __tablename__ = "data_assets"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=True, index=True)      # NULL = 全局
    asset_type = Column(String(20), nullable=False, index=True)     # reference / source / result / import
    category_id = Column(Integer, ForeignKey("reference_categories.id"), nullable=True)
    name = Column(String(200), nullable=False)                      # 显示名
    description = Column(Text, default="")
    file_path = Column(String(500), nullable=False)                 # 物理路径
    file_name = Column(String(200), nullable=False)                 # 原始文件名
    file_size = Column(BigInteger, default=0)
    sheet_summary = Column(JSON, nullable=True)                     # [{sheet_name, rows, cols, headers}]
    parsed_headers = Column(JSON, nullable=True)                    # 解析出的表头结构
    parsed_data = Column(JSON, nullable=True)                       # 完整解析数据（基础数据用，避免重复读文件）
    version = Column(Integer, default=1)
    effective_from = Column(Date, nullable=True)                    # 基础数据生效日期
    effective_to = Column(Date, nullable=True)                      # 基础数据失效日期
    source_task_id = Column(Integer, ForeignKey("compute_tasks.id"), nullable=True)
    uploaded_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    is_active = Column(Boolean, default=True)
    tags = Column(JSON, nullable=True)                              # 自由标签
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    category_rel = relationship("ReferenceCategory", back_populates="assets")
    uploader = relationship("User", foreign_keys=[uploaded_by])
    source_task = relationship("ComputeTask", foreign_keys=[source_task_id], back_populates="output_assets")

    __table_args__ = (
        Index("ix_asset_tenant_type", "tenant_id", "asset_type"),
    )


# ==================== 训练会话 ====================

class TrainingSession(Base):
    __tablename__ = "training_sessions"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=False, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    session_key = Column(String(100), unique=True, nullable=False)  # 原有的 training_id
    mode = Column(String(20), default="formula")                    # formula / modular
    status = Column(String(20), default="running")                  # running / completed / failed / cancelled
    config = Column(JSON, nullable=True)                            # 训练配置参数
    source_asset_ids = Column(JSON, nullable=True)                  # 关联的源文件 asset id 列表
    expected_asset_id = Column(Integer, ForeignKey("data_assets.id"), nullable=True)
    final_script_id = Column(Integer, ForeignKey("scripts.id"), nullable=True)
    total_iterations = Column(Integer, default=0)
    best_accuracy = Column(Float, nullable=True)
    error_message = Column(Text, nullable=True)
    started_at = Column(DateTime, default=datetime.utcnow)
    finished_at = Column(DateTime, nullable=True)

    user = relationship("User", foreign_keys=[user_id])
    expected_asset = relationship("DataAsset", foreign_keys=[expected_asset_id])
    final_script = relationship("Script", foreign_keys=[final_script_id])
    iterations = relationship("TrainingIteration", back_populates="session", order_by="TrainingIteration.iteration_num")


# ==================== 训练迭代 ====================

class TrainingIteration(Base):
    __tablename__ = "training_iterations"

    id = Column(Integer, primary_key=True, autoincrement=True)
    session_id = Column(Integer, ForeignKey("training_sessions.id"), nullable=False, index=True)
    iteration_num = Column(Integer, nullable=False)                 # 第几轮
    status = Column(String(20), default="running")                  # running / completed / failed
    prompt_text = Column(Text, nullable=True)                       # 发给 AI 的 prompt
    ai_response = Column(Text, nullable=True)                       # AI 返回的原始内容
    generated_code = Column(Text, nullable=True)                    # 提取出的代码
    execution_result = Column(JSON, nullable=True)                  # 执行结果摘要
    accuracy = Column(Float, nullable=True)                         # 准确率
    error_details = Column(JSON, nullable=True)                     # 错误详情
    duration_seconds = Column(Float, nullable=True)                 # 本轮耗时
    started_at = Column(DateTime, default=datetime.utcnow)
    finished_at = Column(DateTime, nullable=True)

    session = relationship("TrainingSession", back_populates="iterations")


# ==================== 脚本 ====================

class Script(Base):
    __tablename__ = "scripts"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=False, index=True)
    name = Column(String(200), nullable=False)                      # 脚本名称
    description = Column(Text, default="")
    code = Column(Text, nullable=False)                             # Python 代码
    mode = Column(String(20), default="formula")                    # formula / modular
    config = Column(JSON, nullable=True)                            # 脚本配置（列映射等）
    source_session_id = Column(Integer, ForeignKey("training_sessions.id"), nullable=True)
    accuracy = Column(Float, nullable=True)                         # 训练时的最佳准确率
    version = Column(Integer, default=1)
    is_active = Column(Boolean, default=True)
    created_by = Column(Integer, ForeignKey("users.id"), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    creator = relationship("User", foreign_keys=[created_by])
    source_session = relationship("TrainingSession", foreign_keys=[source_session_id])
    compute_tasks = relationship("ComputeTask", back_populates="script")


# ==================== 计算任务 ====================

class ComputeTask(Base):
    __tablename__ = "compute_tasks"

    id = Column(Integer, primary_key=True, autoincrement=True)
    tenant_id = Column(String(100), nullable=False, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    script_id = Column(Integer, ForeignKey("scripts.id"), nullable=True)
    status = Column(String(20), default="pending")                  # pending / analyzing / analyzed / computing / completed / failed
    parent_task_id = Column(Integer, ForeignKey("compute_tasks.id"), nullable=True)
    analysis_report = Column(JSON, nullable=True)                   # 分析报告
    header_mapping = Column(JSON, nullable=True)                    # 表头映射（用户可修正）
    execution_log = Column(Text, nullable=True)                     # 执行日志
    result_summary = Column(JSON, nullable=True)                    # 结果摘要（行数/sheet数等）
    error_message = Column(Text, nullable=True)
    duration_seconds = Column(Float, nullable=True)
    salary_year = Column(Integer, nullable=True, index=True)        # 薪资年份
    salary_month = Column(Integer, nullable=True, index=True)       # 薪资月份
    created_at = Column(DateTime, default=datetime.utcnow)
    finished_at = Column(DateTime, nullable=True)

    user = relationship("User", foreign_keys=[user_id])
    script = relationship("Script", back_populates="compute_tasks")
    parent_task = relationship("ComputeTask", remote_side=[id], backref="child_tasks")
    inputs = relationship("ComputeTaskInput", back_populates="task")
    output_assets = relationship("DataAsset", foreign_keys="DataAsset.source_task_id", back_populates="source_task")


# ==================== 计算任务输入 ====================

class ComputeTaskInput(Base):
    __tablename__ = "compute_task_inputs"

    id = Column(Integer, primary_key=True, autoincrement=True)
    task_id = Column(Integer, ForeignKey("compute_tasks.id"), nullable=False, index=True)
    asset_id = Column(Integer, ForeignKey("data_assets.id"), nullable=False)
    role = Column(String(30), nullable=False)                       # source / reference / previous_result
    sheet_name = Column(String(100), nullable=True)                 # 指定 sheet（可选）
    created_at = Column(DateTime, default=datetime.utcnow)

    task = relationship("ComputeTask", back_populates="inputs")
    asset = relationship("DataAsset")
