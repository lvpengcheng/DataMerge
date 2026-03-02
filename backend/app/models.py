"""
数据模型定义
"""

from typing import Optional, List, Dict, Any
from pydantic import BaseModel
from datetime import datetime


class TrainingRequest(BaseModel):
    """训练请求模型"""
    rule_files: List[str]
    source_files: List[str]
    expected_result: str
    ai_provider: str = "openai"
    ai_model: str = "gpt-4"
    manual_headers: Optional[Dict[str, Any]] = None


class TrainingResponse(BaseModel):
    """训练响应模型"""
    tenant_id: str
    status: str
    training_result: Dict[str, Any]
    files_uploaded: Dict[str, int]


class CalculationRequest(BaseModel):
    """计算请求模型"""
    data_files: List[str]


class CalculationResponse(BaseModel):
    """计算响应模型"""
    tenant_id: str
    status: str
    batch_id: Optional[str] = None
    result_file: Optional[str] = None
    errors: Optional[List[str]] = None
    message: Optional[str] = None
    execution_result: Optional[Dict[str, Any]] = None


class StorageStats(BaseModel):
    """存储统计模型"""
    tenant_id: str
    total_size: int
    total_size_human: str
    directory_stats: Dict[str, Dict[str, Any]]
    file_counts: Dict[str, int]


class ScriptInfo(BaseModel):
    """脚本信息模型"""
    script_id: str
    tenant_id: str
    created_time: str
    score: float
    iterations: int
    success: bool
    file_path: str
    template_schema: Dict[str, Any]
    manual_headers: Optional[Dict[str, Any]] = None
    source_structure: Dict[str, Any]
    expected_structure: Dict[str, Any]


class BatchInfo(BaseModel):
    """批次信息模型"""
    batch_id: str
    tenant_id: str
    calculation_time: str
    input_file_count: int
    status: str
    input_dir: str
    output_dir: str
    completion_time: Optional[str] = None
    result_file: Optional[str] = None
    execution_result: Optional[Dict[str, Any]] = None


class ValidationResult(BaseModel):
    """验证结果模型"""
    is_valid: bool
    errors: List[str]
    warnings: List[str] = []


class ExecutionResult(BaseModel):
    """执行结果模型"""
    success: bool
    output: str
    error: str
    execution_time: float
    return_value: Optional[Any] = None