"""
Dashboard API - 首页租户计算状态总览
"""

import logging
from datetime import datetime
from fastapi import APIRouter, Depends
from sqlalchemy import func
from sqlalchemy.orm import Session

from ..database.connection import get_db
from ..database.models import ComputeTask
from ..auth.dependencies import get_current_user, get_accessible_tenants

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api/dashboard", tags=["首页"])


@router.get("/tenants")
def get_dashboard_tenants(
    tenant_ids: list = Depends(get_accessible_tenants),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取用户可访问租户的计算状态总览"""
    if not tenant_ids:
        return []

    now = datetime.now()
    cur_year, cur_month = now.year, now.month
    if cur_month == 1:
        last_year, last_month = cur_year - 1, 12
    else:
        last_year, last_month = cur_year, cur_month - 1

    # 当前月: 每个租户完成的计算任务数
    current_month_rows = (
        db.query(
            ComputeTask.tenant_id,
            func.count(ComputeTask.id).label("cnt"),
        )
        .filter(
            ComputeTask.tenant_id.in_(tenant_ids),
            ComputeTask.status == "completed",
            ComputeTask.salary_year == cur_year,
            ComputeTask.salary_month == cur_month,
        )
        .group_by(ComputeTask.tenant_id)
        .all()
    )
    current_month_counts = {row.tenant_id: row.cnt for row in current_month_rows}

    # 上月: 每个租户最新完成任务的时间和耗时
    last_month_tasks = {}
    for tid in tenant_ids:
        task = (
            db.query(ComputeTask)
            .filter(
                ComputeTask.tenant_id == tid,
                ComputeTask.status == "completed",
                ComputeTask.salary_year == last_year,
                ComputeTask.salary_month == last_month,
            )
            .order_by(ComputeTask.finished_at.desc())
            .first()
        )
        if task:
            last_month_tasks[tid] = task

    # 组装结果
    result = []
    for tid in sorted(tenant_ids):
        cur_count = current_month_counts.get(tid, 0)
        last_task = last_month_tasks.get(tid)
        result.append({
            "tenant_id": tid,
            "tenant_name": tid,
            "current_month_computed": cur_count > 0,
            "current_month_task_count": cur_count,
            "last_month_compute_time": (
                last_task.finished_at.isoformat()
                if last_task and last_task.finished_at
                else None
            ),
            "last_month_duration_seconds": (
                last_task.duration_seconds if last_task else None
            ),
        })

    return result
