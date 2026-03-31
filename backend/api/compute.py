"""
计算任务 API（两步计算流程：分析 → 执行）
"""

import os
import shutil
import tempfile
import logging
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Optional, List

from fastapi import APIRouter, Depends, UploadFile, File, Form, Query, HTTPException
from sqlalchemy.orm import Session
from pydantic import BaseModel

from ..database.connection import get_db
from ..database.models import (
    ComputeTask, ComputeTaskInput, DataAsset, Script,
)
from ..auth.dependencies import get_current_user

router = APIRouter(prefix="/api/compute2", tags=["计算任务"])

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


def _load_script_code(script: Script) -> str:
    """加载脚本代码：优先读取磁盘文件（支持手动编辑），回退到DB。

    训练时脚本同时保存到 DB(Script.code) 和磁盘(tenants/{id}/scripts/script_{hash}.py)。
    用户可能手动编辑磁盘文件，此函数优先使用磁盘版本。
    """
    try:
        # 通过DB代码的hash推算磁盘文件名
        code_hash = hashlib.md5(script.code.encode('utf-8')).hexdigest()[:12]
        disk_script_id = f"script_{code_hash}"
        disk_path = PROJECT_ROOT / "tenants" / script.tenant_id / "scripts" / f"{disk_script_id}.py"

        if disk_path.exists():
            disk_code = disk_path.read_text(encoding='utf-8')
            if disk_code != script.code:
                logger.info(f"[脚本同步] 检测到磁盘文件已手动修改，使用磁盘版本: {disk_path.name}")
                # 同步回DB，保持一致
                script.code = disk_code
                try:
                    from ..database.connection import SessionLocal
                    with SessionLocal() as db_sync:
                        db_sync.query(Script).filter_by(id=script.id).update({"code": disk_code})
                        db_sync.commit()
                    logger.info(f"[脚本同步] 磁盘修改已同步回DB: Script.id={script.id}")
                except Exception as e:
                    logger.warning(f"[脚本同步] 回写DB失败（不影响本次执行）: {e}")
                return disk_code
            else:
                logger.debug(f"[脚本同步] 磁盘文件与DB一致: {disk_path.name}")
                return script.code
        else:
            logger.debug(f"[脚本同步] 磁盘文件不存在: {disk_path}, 使用DB版本")
            return script.code
    except Exception as e:
        logger.warning(f"[脚本同步] 读取磁盘脚本失败，使用DB版本: {e}")
        return script.code


# ==================== Pydantic 模型 ====================

class AnalyzeRequest(BaseModel):
    tenant_id: str
    script_id: int
    asset_ids: List[int] = []           # 已有数据资产 ID
    salary_year: Optional[int] = None
    salary_month: Optional[int] = None
    standard_hours: Optional[float] = None


class ExecuteRequest(BaseModel):
    header_mapping: Optional[dict] = None   # 用户修正的表头映射


# ==================== 辅助函数 ====================

def _task_to_dict(task: ComputeTask) -> dict:
    # 查询输出资产
    output_assets = []
    if task.id:
        from sqlalchemy import inspect
        session = inspect(task).session
        if session:
            assets = session.query(DataAsset).filter_by(source_task_id=task.id, is_active=True).all()
            output_assets = [
                {
                    "id": a.id,
                    "name": a.name,
                    "file_name": a.file_name,
                    "file_size": a.file_size,
                }
                for a in assets
            ]

    return {
        "id": task.id,
        "tenant_id": task.tenant_id,
        "script_id": task.script_id,
        "salary_year": task.salary_year,
        "salary_month": task.salary_month,
        "status": task.status,
        "parent_task_id": task.parent_task_id,
        "analysis_report": task.analysis_report,
        "header_mapping": task.header_mapping,
        "result_summary": task.result_summary,
        "error_message": task.error_message,
        "duration_seconds": task.duration_seconds,
        "created_at": task.created_at.isoformat() if task.created_at else None,
        "finished_at": task.finished_at.isoformat() if task.finished_at else None,
        "inputs": [
            {
                "id": inp.id,
                "asset_id": inp.asset_id,
                "role": inp.role,
                "sheet_name": inp.sheet_name,
                "asset_name": inp.asset.name if inp.asset else None,
                "asset_type": inp.asset.asset_type if inp.asset else None,
                "file_name": inp.asset.file_name if inp.asset else None,
            }
            for inp in task.inputs
        ] if task.inputs else [],
        "output_assets": output_assets,
    }


# ==================== 路由 ====================

@router.post("/analyze")
async def analyze_compute(
    req: AnalyzeRequest,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """第一步：分析数据源，生成匹配报告"""
    # 验证脚本
    script = db.query(Script).filter_by(id=req.script_id, is_active=True).first()
    if not script:
        raise HTTPException(status_code=404, detail="脚本不存在或已停用")

    # 验证资产
    assets = db.query(DataAsset).filter(
        DataAsset.id.in_(req.asset_ids),
        DataAsset.is_active == True,
    ).all()
    if len(assets) != len(req.asset_ids):
        raise HTTPException(status_code=400, detail="部分数据资产不存在或已停用")

    # 创建计算任务
    task = ComputeTask(
        tenant_id=req.tenant_id,
        user_id=current_user.id,
        script_id=req.script_id,
        salary_year=req.salary_year,
        salary_month=req.salary_month,
        status="analyzing",
    )
    db.add(task)
    db.commit()
    db.refresh(task)

    # 关联输入资产
    for asset in assets:
        role = "reference" if asset.asset_type == "reference" else "source"
        inp = ComputeTaskInput(task_id=task.id, asset_id=asset.id, role=role)
        db.add(inp)
    db.commit()

    # 生成分析报告
    analysis = {
        "script": {"id": script.id, "name": script.name, "mode": script.mode},
        "inputs": [],
        "params": {
            "salary_year": req.salary_year,
            "salary_month": req.salary_month,
            "standard_hours": req.standard_hours,
        },
    }

    for asset in assets:
        asset_info = {
            "id": asset.id,
            "name": asset.name,
            "type": asset.asset_type,
            "file_name": asset.file_name,
            "sheet_summary": asset.sheet_summary,
        }
        analysis["inputs"].append(asset_info)

    task.analysis_report = analysis
    task.status = "analyzed"
    db.commit()
    db.refresh(task)

    return _task_to_dict(task)


@router.post("/{task_id}/execute")
async def execute_compute(
    task_id: int,
    req: Optional[ExecuteRequest] = None,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """第二步：确认执行计算"""
    task = db.query(ComputeTask).filter_by(id=task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="计算任务不存在")
    if task.status not in ("analyzed", "failed"):
        raise HTTPException(status_code=400, detail=f"任务状态不允许执行: {task.status}")

    # 更新用户修正的表头映射
    if req and req.header_mapping:
        task.header_mapping = req.header_mapping

    task.status = "computing"
    db.commit()

    script = db.query(Script).filter_by(id=task.script_id).first()
    if not script:
        task.status = "failed"
        task.error_message = "脚本不存在"
        db.commit()
        raise HTTPException(status_code=404, detail="脚本不存在")

    # 准备执行环境
    start_time = datetime.utcnow()
    temp_dir = Path(tempfile.mkdtemp(prefix="compute2_"))

    try:
        source_dir = temp_dir / "source"
        source_dir.mkdir(parents=True, exist_ok=True)
        output_dir = temp_dir / "output"
        output_dir.mkdir(parents=True, exist_ok=True)

        # 复制输入文件到临时目录
        source_files = []
        reference_files = {}
        for inp in task.inputs:
            asset = db.query(DataAsset).filter_by(id=inp.asset_id).first()
            if not asset or not os.path.exists(asset.file_path):
                continue
            dest = source_dir / asset.file_name
            shutil.copy2(asset.file_path, dest)
            if inp.role == "reference":
                cat_code = asset.category_rel.code if asset.category_rel else "other"
                reference_files[cat_code] = str(dest)
            else:
                source_files.append(str(dest))

        # 写入脚本（优先磁盘文件，支持手动编辑）
        script_code = _load_script_code(script)
        script_path = temp_dir / "compute_script.py"
        script_path.write_text(script_code, encoding="utf-8")

        # 执行脚本
        from ..sandbox.code_sandbox import CodeSandbox
        sandbox = CodeSandbox()

        params = task.analysis_report.get("params", {}) if task.analysis_report else {}
        env_vars = {
            "SALARY_YEAR": str(params.get("salary_year", "")),
            "SALARY_MONTH": str(params.get("salary_month", "")),
            "MONTHLY_STANDARD_HOURS": str(params.get("standard_hours", "")),
        }

        exec_result = sandbox.execute(
            script_path=str(script_path),
            source_dir=str(source_dir),
            output_dir=str(output_dir),
            env_vars=env_vars,
        )

        duration = (datetime.utcnow() - start_time).total_seconds()

        if exec_result.get("success"):
            # 查找输出文件
            output_files = list(output_dir.glob("*.xlsx")) + list(output_dir.glob("*.xls"))

            result_summary = {
                "output_files": [f.name for f in output_files],
                "execution_time": duration,
            }

            # 保存结果到租户目录并注册为 data_asset
            saved_assets = []
            for output_file in output_files:
                # 保存到租户目录
                tenant_output_dir = PROJECT_ROOT / "tenants" / task.tenant_id / "assets" / "result"
                tenant_output_dir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                dest_name = f"{timestamp}_{output_file.name}"
                dest_path = tenant_output_dir / dest_name
                shutil.copy2(output_file, dest_path)

                # 注册为数据资产
                result_asset = DataAsset(
                    tenant_id=task.tenant_id,
                    asset_type="result",
                    name=f"计算结果_{output_file.name}",
                    file_path=str(dest_path),
                    file_name=output_file.name,
                    file_size=dest_path.stat().st_size,
                    source_task_id=task.id,
                    uploaded_by=current_user.id,
                )
                db.add(result_asset)
                saved_assets.append(dest_name)

            result_summary["saved_assets"] = saved_assets
            task.status = "completed"
            task.result_summary = result_summary
            task.duration_seconds = duration
            task.finished_at = datetime.utcnow()
        else:
            task.status = "failed"
            task.error_message = exec_result.get("error", "执行失败")
            task.duration_seconds = duration
            task.finished_at = datetime.utcnow()

        db.commit()
        db.refresh(task)
        return _task_to_dict(task)

    except Exception as e:
        task.status = "failed"
        task.error_message = str(e)
        task.finished_at = datetime.utcnow()
        db.commit()
        raise HTTPException(status_code=500, detail=f"计算执行失败: {str(e)}")
    finally:
        # 清理临时目录
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


@router.get("/tasks")
def list_compute_tasks(
    tenant_id: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    limit: int = Query(50, ge=1, le=200),
    offset: int = Query(0, ge=0),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """计算任务列表"""
    q = db.query(ComputeTask)
    if tenant_id:
        q = q.filter(ComputeTask.tenant_id == tenant_id)
    if status:
        q = q.filter(ComputeTask.status == status)
    total = q.count()
    tasks = q.order_by(ComputeTask.created_at.desc()).offset(offset).limit(limit).all()
    return {
        "total": total,
        "items": [_task_to_dict(t) for t in tasks],
    }


@router.get("/tasks/{task_id}")
def get_compute_task(
    task_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """计算任务详情"""
    task = db.query(ComputeTask).filter_by(id=task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="计算任务不存在")
    return _task_to_dict(task)


@router.get("/tasks/{task_id}/result")
def get_compute_result(
    task_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取计算结果（输出的数据资产列表）"""
    task = db.query(ComputeTask).filter_by(id=task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="计算任务不存在")

    output_assets = (
        db.query(DataAsset)
        .filter_by(source_task_id=task_id, is_active=True)
        .all()
    )
    return [
        {
            "id": a.id,
            "name": a.name,
            "file_name": a.file_name,
            "file_size": a.file_size,
            "file_path": a.file_path,
            "created_at": a.created_at.isoformat() if a.created_at else None,
        }
        for a in output_assets
    ]


@router.get("/history")
def compute_history(
    tenant_id: str = Query(...),
    limit: int = Query(20, ge=1, le=100),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """计算历史（含计算链路追溯）"""
    tasks = (
        db.query(ComputeTask)
        .filter_by(tenant_id=tenant_id)
        .order_by(ComputeTask.created_at.desc())
        .limit(limit)
        .all()
    )

    result = []
    for task in tasks:
        item = _task_to_dict(task)
        # 追溯输出资产被哪些后续任务使用
        output_assets = db.query(DataAsset).filter_by(source_task_id=task.id).all()
        downstream_tasks = []
        for oa in output_assets:
            usages = db.query(ComputeTaskInput).filter_by(asset_id=oa.id).all()
            for u in usages:
                downstream_tasks.append(u.task_id)
        item["downstream_task_ids"] = list(set(downstream_tasks))
        result.append(item)
    return result
