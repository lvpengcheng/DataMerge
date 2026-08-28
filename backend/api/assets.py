"""
数据资产管理 API
"""

import asyncio
import json
import logging
import os
import shutil
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Optional, List

from fastapi import APIRouter, Depends, UploadFile, File, Form, Query, HTTPException
from sqlalchemy import or_
from sqlalchemy.orm import Session
from pydantic import BaseModel

from ..database.connection import SessionLocal, get_db
from ..database.models import (
    AssetUploadTask, DataAsset, ReferenceCategory, TenantAuthorization,
)
from ..auth.dependencies import get_current_user, get_accessible_tenants

router = APIRouter(prefix="/api/assets", tags=["数据资产"])

# 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
# data 在 Docker Compose 中是持久化卷；容器/服务重启后排队文件仍可恢复。
UPLOAD_STAGING_ROOT = PROJECT_ROOT / "data" / "asset_uploads"
_UPLOAD_CHUNK_SIZE = 1024 * 1024
_upload_queue = None
_upload_consumer_task = None
logger = logging.getLogger(__name__)

# Starlette 在进入端点前会先解析 multipart。把单文件的内存缓冲阈值
# 压到 64KiB，超过即转临时文件，避免“大量小文件”在调用端点前已堆满内存。
try:
    from starlette.formparsers import MultiPartParser
    _spool_bytes = int(os.getenv("UPLOAD_SPOOL_MEMORY_BYTES", str(64 * 1024)))
    if hasattr(MultiPartParser, "spool_max_size"):
        MultiPartParser.spool_max_size = _spool_bytes
    if hasattr(MultiPartParser, "max_file_size"):
        MultiPartParser.max_file_size = _spool_bytes
except Exception:
    pass


# ==================== Pydantic 模型 ====================

class AssetOut(BaseModel):
    id: int
    tenant_id: Optional[str] = None
    asset_type: str
    category_id: Optional[int] = None
    category_name: Optional[str] = None
    name: str
    description: str = ""
    file_name: str
    file_size: int = 0
    sheet_summary: Optional[list] = None
    version: int = 1
    effective_from: Optional[str] = None
    effective_to: Optional[str] = None
    is_active: bool = True
    tags: Optional[list] = None
    created_at: Optional[str] = None
    updated_at: Optional[str] = None

    class Config:
        from_attributes = True


class AssetUpdateIn(BaseModel):
    name: Optional[str] = None
    description: Optional[str] = None
    category_id: Optional[int] = None
    effective_from: Optional[str] = None
    effective_to: Optional[str] = None
    tags: Optional[list] = None
    is_active: Optional[bool] = None


class ReferenceCategoryOut(BaseModel):
    id: int
    code: str
    name: str
    description: str = ""
    scope: str = "global"
    sort_order: int = 0

    class Config:
        from_attributes = True


# ==================== 辅助函数 ====================

def _parse_sheet_summary(file_path: str) -> list:
    """解析 Excel 文件的 sheet 结构摘要"""
    try:
        from excel_parser import IntelligentExcelParser
        parser = IntelligentExcelParser()
        results = parser.parse_excel_file(file_path, max_data_rows=5)
        summary = []
        for sheet_data in results:
            regions = sheet_data.regions
            headers = []
            total_rows = 0
            for region in regions:
                headers.extend(list(region.head_data.keys()) if region.head_data else [])
                total_rows += len(region.data)
            summary.append({
                "sheet_name": sheet_data.sheet_name,
                "rows": total_rows,
                "headers": headers[:50],  # 最多 50 列
                "regions": len(regions),
            })
        return summary
    except Exception as e:
        return [{"error": str(e)}]


def _parse_full_data(file_path: str) -> list:
    """解析 Excel 完整数据（用于基础数据存入DB，避免重复读文件）"""
    try:
        import dataclasses
        from excel_parser import IntelligentExcelParser
        parser = IntelligentExcelParser()
        results = parser.parse_excel_file(file_path)
        return [dataclasses.asdict(sheet) for sheet in results]
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning(f"解析完整数据失败: {e}")
        return None


def _parse_reference_once(file_path: str):
    """基础数据只打开、解析一次，同时产生摘要和全量数据。"""
    import dataclasses
    from excel_parser import IntelligentExcelParser

    results = IntelligentExcelParser().parse_excel_file(
        file_path, calculate_formulas=False,
    )
    summary = []
    for sheet_data in results:
        headers = []
        total_rows = 0
        for region in sheet_data.regions:
            headers.extend(list(region.head_data.keys()) if region.head_data else [])
            total_rows += len(region.data)
        summary.append({
            "sheet_name": sheet_data.sheet_name,
            "rows": total_rows,
            "headers": headers[:50],
            "regions": len(sheet_data.regions),
        })
    return summary, [dataclasses.asdict(sheet) for sheet in results]


def _get_asset_storage_dir(tenant_id: Optional[str], asset_type: str) -> Path:
    """获取资产存储目录"""
    if tenant_id:
        base = PROJECT_ROOT / "tenants" / tenant_id / "assets" / asset_type
    else:
        base = PROJECT_ROOT / "global_assets" / asset_type
    base.mkdir(parents=True, exist_ok=True)
    return base


def _asset_to_out(asset: DataAsset) -> dict:
    """DataAsset ORM → 响应 dict"""
    cat_name = asset.category_rel.name if asset.category_rel else None
    return {
        "id": asset.id,
        "tenant_id": asset.tenant_id,
        "asset_type": asset.asset_type,
        "category_id": asset.category_id,
        "category_name": cat_name,
        "name": asset.name,
        "description": asset.description or "",
        "file_name": asset.file_name,
        "file_size": asset.file_size,
        "sheet_summary": asset.sheet_summary,
        "version": asset.version,
        "effective_from": str(asset.effective_from) if asset.effective_from else None,
        "effective_to": str(asset.effective_to) if asset.effective_to else None,
        "is_active": asset.is_active,
        "tags": asset.tags,
        "created_at": asset.created_at.isoformat() if asset.created_at else None,
        "updated_at": asset.updated_at.isoformat() if asset.updated_at else None,
    }


def _is_admin(user) -> bool:
    return bool(user.role and user.role.name == "admin")


def _assert_can_manage(tenant_id, accessible, is_admin):
    """写操作(上传/改/删)的租户校验：非管理员禁止操作全局(None)及未授权租户。"""
    if is_admin:
        return
    if tenant_id is None:
        raise HTTPException(status_code=403, detail="无权管理全局基础数据")
    if tenant_id not in accessible:
        raise HTTPException(status_code=403, detail=f"无权管理租户 '{tenant_id}' 的数据")


def _assert_can_view(tenant_id, accessible, is_admin):
    """读操作的租户校验：全局(None)对所有人可见；租户数据须在授权内。"""
    if is_admin or tenant_id is None:
        return
    if tenant_id not in accessible:
        raise HTTPException(status_code=403, detail="无权访问该租户数据")



# ==================== 静态路由（必须在 /{asset_id} 之前） ====================

@router.get("/reference-categories", response_model=List[ReferenceCategoryOut])
def list_reference_categories(
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """基础数据分类列表"""
    return db.query(ReferenceCategory).order_by(ReferenceCategory.sort_order).all()


@router.post("/reference-categories")
def create_reference_category(
    code: str = Form(...),
    name: str = Form(...),
    description: str = Form(""),
    scope: str = Form("global"),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """新增基础数据分类"""
    existing = db.query(ReferenceCategory).filter_by(code=code).first()
    if existing:
        raise HTTPException(status_code=400, detail=f"分类代码 '{code}' 已存在")

    max_order = db.query(ReferenceCategory).count()
    cat = ReferenceCategory(
        code=code, name=name, description=description,
        scope=scope, sort_order=max_order + 1,
    )
    db.add(cat)
    db.commit()
    db.refresh(cat)
    return {"id": cat.id, "code": cat.code, "name": cat.name}


@router.get("/reference")
def list_reference_assets(
    tenant_id: Optional[str] = Query(None),
    category_id: Optional[int] = Query(None),
    scope: Optional[str] = Query(None),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """全局 + 当前租户的基础数据列表"""
    q = db.query(DataAsset).filter(
        DataAsset.asset_type == "reference",
        DataAsset.is_active == True,
    )
    # 作用域筛选
    if scope == "global":
        q = q.filter(DataAsset.tenant_id.is_(None))
    elif scope == "tenant":
        q = q.filter(DataAsset.tenant_id.isnot(None))
        if tenant_id:
            q = q.filter(DataAsset.tenant_id == tenant_id)
    else:
        # 默认: 全局 + 指定租户
        if tenant_id:
            q = q.filter((DataAsset.tenant_id == tenant_id) | (DataAsset.tenant_id.is_(None)))
    if category_id:
        q = q.filter(DataAsset.category_id == category_id)
    # 非管理员: 仅授权租户 + 全局
    if not _is_admin(current_user):
        q = q.filter(or_(DataAsset.tenant_id.in_(accessible), DataAsset.tenant_id.is_(None)))
    assets = q.order_by(DataAsset.created_at.desc()).all()
    return [_asset_to_out(a) for a in assets]


@router.get("/tenants")
def list_available_tenants(
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """获取可用租户列表（用于上传基础数据时选择作用域）"""
    # 非管理员: 仅返回授权租户
    if not _is_admin(current_user):
        return [{"tenant_id": t} for t in sorted(accessible)]
    from sqlalchemy import union_all, literal_column
    # 从 tenant_authorizations 和 data_assets 汇总所有租户
    q1 = db.query(TenantAuthorization.tenant_id).distinct()
    q2 = db.query(DataAsset.tenant_id).filter(DataAsset.tenant_id.isnot(None)).distinct()
    all_tenants = set()
    for row in q1.all():
        all_tenants.add(row[0])
    for row in q2.all():
        all_tenants.add(row[0])
    return [{"tenant_id": t} for t in sorted(all_tenants)]


@router.get("")
def list_assets(
    tenant_id: Optional[str] = Query(None),
    asset_type: Optional[str] = Query(None),
    category_id: Optional[int] = Query(None),
    scope: Optional[str] = Query(None),
    is_active: bool = Query(True),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """资产列表（按 type/category/tenant/scope 筛选）"""
    q = db.query(DataAsset).filter(DataAsset.is_active == is_active)
    if asset_type:
        q = q.filter(DataAsset.asset_type == asset_type)
    # 作用域筛选
    if scope == "global":
        q = q.filter(DataAsset.tenant_id.is_(None))
    elif scope == "tenant":
        q = q.filter(DataAsset.tenant_id.isnot(None))
        if tenant_id:
            q = q.filter(DataAsset.tenant_id == tenant_id)
    else:
        if tenant_id:
            q = q.filter((DataAsset.tenant_id == tenant_id) | (DataAsset.tenant_id.is_(None)))
    if category_id:
        q = q.filter(DataAsset.category_id == category_id)
    # 非管理员: 仅授权租户 + 全局
    if not _is_admin(current_user):
        q = q.filter(or_(DataAsset.tenant_id.in_(accessible), DataAsset.tenant_id.is_(None)))
    assets = q.order_by(DataAsset.created_at.desc()).all()
    return [_asset_to_out(a) for a in assets]


@router.post("/upload", status_code=202)
async def upload_asset(
    file: UploadFile = File(...),
    tenant_id: Optional[str] = Form(None),
    asset_type: str = Form("source"),
    category_id: Optional[int] = Form(None),
    name: Optional[str] = Form(None),
    description: str = Form(""),
    effective_from: Optional[str] = Form(None),
    effective_to: Optional[str] = Form(None),
    tags: Optional[str] = Form(None),  # JSON string
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """单文件上传：流式落盘后返回异步任务。"""
    _assert_can_manage(tenant_id, accessible, _is_admin(current_user))
    task = await _create_asset_upload_task(
        uploads=[file], tenant_id=tenant_id, asset_type=asset_type,
        category_id=category_id, description=description,
        effective_from=effective_from, effective_to=effective_to, tags=tags,
        uploaded_by=current_user.id, db=db,
        names=[name] if name else None,
    )
    return _upload_task_to_out(task)


@router.post("/upload-batch", status_code=202)
async def upload_assets_batch(
    files: List[UploadFile] = File(...),
    tenant_id: Optional[str] = Form(None),
    asset_type: str = Form("reference"),
    category_id: Optional[int] = Form(None),
    description: str = Form(""),
    effective_from: Optional[str] = Form(None),
    effective_to: Optional[str] = Form(None),
    tags: Optional[str] = Form(None),  # JSON string
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """批量上传：不限人为文件数，流式落盘后立即返回异步任务。"""
    _assert_can_manage(tenant_id, accessible, _is_admin(current_user))
    task = await _create_asset_upload_task(
        uploads=files, tenant_id=tenant_id, asset_type=asset_type,
        category_id=category_id, description=description,
        effective_from=effective_from, effective_to=effective_to, tags=tags,
        uploaded_by=current_user.id, db=db,
    )
    return _upload_task_to_out(task)


def _parse_tags(tags):
    if not tags:
        return None
    try:
        return json.loads(tags)
    except Exception:
        return [t.strip() for t in tags.split(",") if t.strip()]


async def _stream_upload_to_disk(upload: UploadFile, destination: Path) -> int:
    """固定 1MiB 分块落盘，不调用 read() 全量读入内存。"""
    written = 0
    try:
        with destination.open("wb") as output:
            while True:
                chunk = await upload.read(_UPLOAD_CHUNK_SIZE)
                if not chunk:
                    break
                output.write(chunk)
                written += len(chunk)
    finally:
        await upload.close()
    return written


async def _create_asset_upload_task(*, uploads, tenant_id, asset_type, category_id,
                                    description, effective_from, effective_to, tags,
                                    uploaded_by, db, names=None, previous_asset_id=None):
    task_id = uuid.uuid4().hex
    staging_dir = UPLOAD_STAGING_ROOT / task_id
    staging_dir.mkdir(parents=True, exist_ok=False)
    manifest = []
    try:
        for index, upload in enumerate(uploads):
            original_name = Path(upload.filename or f"upload_{index + 1}.xlsx").name
            suffix = Path(original_name).suffix.lower()
            staged_path = staging_dir / f"{index:08d}_{uuid.uuid4().hex}{suffix}"
            size = await _stream_upload_to_disk(upload, staged_path)
            manifest.append({
                "path": str(staged_path),
                "filename": original_name,
                "size": size,
                "name": names[index] if names and index < len(names) else None,
                "previous_asset_id": previous_asset_id,
            })
        if not manifest:
            raise HTTPException(status_code=400, detail="没有可上传的文件")

        task = AssetUploadTask(
            id=task_id, tenant_id=tenant_id, asset_type=asset_type,
            category_id=category_id, description=description or "",
            effective_from=date.fromisoformat(effective_from) if effective_from else None,
            effective_to=date.fromisoformat(effective_to) if effective_to else None,
            tags=_parse_tags(tags), uploaded_by=uploaded_by, status="queued",
            total_files=len(manifest), staging_dir=str(staging_dir), files=manifest,
            result={"created": [], "failed": []},
        )
        db.add(task)
        db.commit()
        db.refresh(task)
        await _ensure_upload_dispatcher()
        _upload_queue.put_nowait(task.id)
        return task
    except Exception:
        db.rollback()
        shutil.rmtree(staging_dir, ignore_errors=True)
        raise


def _upload_task_to_out(task: AssetUploadTask) -> dict:
    result = task.result or {"created": [], "failed": []}
    return {
        "task_id": task.id,
        "status": task.status,
        "total_files": task.total_files or 0,
        "completed_files": task.completed_files or 0,
        "failed_files": task.failed_files or 0,
        "current_file": task.current_file,
        "created": result.get("created", []),
        "failed": result.get("failed", []),
        "error": task.error_message,
    }


@router.get("/upload-tasks/{task_id}")
def get_upload_task(
    task_id: str,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    task = db.query(AssetUploadTask).filter_by(id=task_id).first()
    if not task:
        raise HTTPException(status_code=404, detail="上传任务不存在")
    if not _is_admin(current_user) and task.uploaded_by != current_user.id:
        raise HTTPException(status_code=403, detail="无权查看该上传任务")
    return _upload_task_to_out(task)


def _process_staged_asset_subprocess(payload: dict) -> dict:
    """子进程入口：一个文件的预处理、解析和入库全部在这里完成。"""
    staged_path = Path(payload["path"])
    display_filename = payload["filename"]
    storage_dir = _get_asset_storage_dir(payload.get("tenant_id"), payload["asset_type"])
    final_name = f"{datetime.now():%Y%m%d_%H%M%S_%f}_{uuid.uuid4().hex[:8]}_{display_filename}"
    file_path = storage_dir / final_name
    shutil.move(str(staged_path), str(file_path))
    db = SessionLocal()
    try:
        from ..utils.source_normalizer import convert_xls_to_xlsx, normalize_misformatted_dates
        converted = convert_xls_to_xlsx(str(file_path))
        if converted != str(file_path):
            file_path = Path(converted)
            if display_filename.lower().endswith(".xls"):
                display_filename = display_filename[:-4] + ".xlsx"

        # 上传阶段不做 banner 拆分。该预处理会为了断开跨 sheet 引用而
        # 展平公式，不仅额外多次打开 Workbook，也违背“上传不计算公式”。

        # 上传阶段禁止全工作簿公式重算。
        normalize_misformatted_dates(str(file_path), calculate_formulas=False)

        if payload["asset_type"] == "reference":
            sheet_summary, parsed_data = _parse_reference_once(str(file_path))
        else:
            sheet_summary = _parse_sheet_summary(str(file_path))
            parsed_data = None

        previous_asset = None
        if payload.get("previous_asset_id"):
            previous_asset = db.query(DataAsset).filter_by(
                id=payload["previous_asset_id"]
            ).first()
            if not previous_asset:
                raise RuntimeError("待更新的旧版本资产不存在")

        asset = DataAsset(
            tenant_id=payload.get("tenant_id"), asset_type=payload["asset_type"],
            category_id=payload.get("category_id"),
            name=payload.get("name") or display_filename,
            description=payload.get("description") or "", file_path=str(file_path),
            file_name=display_filename, file_size=file_path.stat().st_size,
            sheet_summary=sheet_summary, parsed_data=parsed_data,
            effective_from=date.fromisoformat(payload["effective_from"])
                if payload.get("effective_from") else None,
            effective_to=date.fromisoformat(payload["effective_to"])
                if payload.get("effective_to") else None,
            uploaded_by=payload.get("uploaded_by"), tags=payload.get("tags"),
            version=(previous_asset.version + 1) if previous_asset else 1,
        )
        db.add(asset)
        if previous_asset:
            previous_asset.is_active = False
            previous_asset.updated_at = datetime.utcnow()
        db.commit()
        db.refresh(asset)
        return _asset_to_out(asset)
    except Exception:
        db.rollback()
        try:
            if file_path.exists():
                file_path.unlink()
        except Exception:
            pass
        raise
    finally:
        db.close()


async def _run_asset_upload_task(task_id: str):
    db = SessionLocal()
    try:
        task = db.query(AssetUploadTask).filter_by(id=task_id).first()
        if not task or task.status not in ("queued", "processing"):
            return
        task.status = "processing"
        task.started_at = task.started_at or datetime.utcnow()
        task.error_message = None
        db.commit()

        result = task.result or {"created": [], "failed": []}
        completed = task.completed_files or 0
        already_done = completed + (task.failed_files or 0)
        manifest = list(task.files or [])
        for item in manifest[already_done:]:
            task.current_file = item["filename"]
            db.commit()
            payload = dict(item)
            payload.update({
                "tenant_id": task.tenant_id, "asset_type": task.asset_type,
                "category_id": task.category_id, "description": task.description,
                "effective_from": str(task.effective_from) if task.effective_from else None,
                "effective_to": str(task.effective_to) if task.effective_to else None,
                "tags": task.tags, "uploaded_by": task.uploaded_by,
            })
            from ..utils.subprocess_runner import run_in_fresh_subprocess_async
            from ..utils.upload_stream import get_excel_work_semaphore
            async with get_excel_work_semaphore():
                sub = await run_in_fresh_subprocess_async(
                    "backend.api.assets:_process_staged_asset_subprocess",
                    args=(payload,), timeout=int(os.getenv("ASSET_UPLOAD_PARSE_TIMEOUT", "900")),
                    max_memory_mb=int(os.getenv("ASSET_UPLOAD_MAX_MEMORY_MB", "4096")),
                )
            if sub.success:
                result["created"].append(sub.result)
                task.completed_files = (task.completed_files or 0) + 1
            else:
                result["failed"].append({"filename": item["filename"], "error": sub.error})
                task.failed_files = (task.failed_files or 0) + 1
            task.result = dict(result)
            db.commit()

        task.status = "completed" if result["created"] else "failed"
        task.current_file = None
        task.finished_at = datetime.utcnow()
        if task.status == "failed":
            task.error_message = "所有文件处理失败"
        db.commit()
    except Exception as exc:
        db.rollback()
        task = db.query(AssetUploadTask).filter_by(id=task_id).first()
        if task:
            task.status = "failed"
            task.error_message = str(exc)
            task.finished_at = datetime.utcnow()
            db.commit()
        logger.exception("上传任务执行失败: %s", task_id)
    finally:
        task = db.query(AssetUploadTask).filter_by(id=task_id).first()
        if task:
            shutil.rmtree(task.staging_dir, ignore_errors=True)
        db.close()


async def _asset_upload_consumer():
    """唯一消费者：无论一次上传多少文件，上传解析并发恒为 1。"""
    while True:
        task_id = await _upload_queue.get()
        try:
            await _run_asset_upload_task(task_id)
        finally:
            _upload_queue.task_done()


async def _ensure_upload_dispatcher():
    global _upload_queue, _upload_consumer_task
    if _upload_queue is None:
        _upload_queue = asyncio.Queue()
    if _upload_consumer_task is None or _upload_consumer_task.done():
        _upload_consumer_task = asyncio.create_task(_asset_upload_consumer())


async def start_asset_upload_dispatcher():
    """应用启动时恢复未完成任务并启动单消费者。"""
    await _ensure_upload_dispatcher()
    db = SessionLocal()
    try:
        tasks = db.query(AssetUploadTask).filter(
            AssetUploadTask.status.in_(["queued", "processing"])
        ).order_by(AssetUploadTask.created_at).all()
        for task in tasks:
            task.status = "queued"
            _upload_queue.put_nowait(task.id)
        db.commit()
    finally:
        db.close()



# ==================== 动态路由 /{asset_id} ====================

@router.get("/{asset_id}")
def get_asset(
    asset_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """获取单个资产详情"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_view(asset.tenant_id, accessible, _is_admin(current_user))
    return _asset_to_out(asset)


@router.get("/{asset_id}/download")
def download_asset(
    asset_id: int,
    format: Optional[str] = Query(None, description="下载格式: 空=原始, pdf, encrypted"),
    password: Optional[str] = Query(None, description="加密密码（encrypted格式用）"),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """下载资产文件（支持原始/PDF/加密Excel）"""
    from fastapi.responses import FileResponse
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_view(asset.tenant_id, accessible, _is_admin(current_user))
    if not os.path.exists(asset.file_path):
        raise HTTPException(status_code=404, detail="文件不存在")

    # 原始下载
    if not format:
        # 根据实际文件扩展名决定 MIME 类型
        ext = Path(asset.file_path).suffix.lower()
        mime_map = {
            ".zip": "application/zip",
            ".pdf": "application/pdf",
            ".csv": "text/csv",
            ".xls": "application/vnd.ms-excel",
        }
        media_type = mime_map.get(ext, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return FileResponse(
            path=asset.file_path,
            filename=asset.file_name,
            media_type=media_type,
        )

    from ..utils.aspose_helper import convert_to_pdf, encrypt_excel

    base_name = Path(asset.file_name).stem

    if format == "pdf":
        pdf_path = convert_to_pdf(asset.file_path)
        return FileResponse(
            path=pdf_path,
            filename=f"{base_name}.pdf",
            media_type="application/pdf",
            background=None,
        )

    if format == "encrypted":
        pwd = password or "123456"
        enc_path = encrypt_excel(asset.file_path, password=pwd)
        return FileResponse(
            path=enc_path,
            filename=f"{base_name}_加密.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    raise HTTPException(status_code=400, detail=f"不支持的格式: {format}，可选: pdf, encrypted")


@router.get("/{asset_id}/preview")
def preview_asset(
    asset_id: int,
    rows: int = Query(20, ge=1, le=200),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """预览资产数据（前 N 行）"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_view(asset.tenant_id, accessible, _is_admin(current_user))
    if not os.path.exists(asset.file_path):
        raise HTTPException(status_code=404, detail="文件不存在")

    try:
        import pandas as pd
        xls = pd.ExcelFile(asset.file_path)
        result = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, nrows=rows)
            result[sheet_name] = {
                "headers": list(df.columns),
                "data": df.fillna("").values.tolist(),
            }
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"预览失败: {str(e)}")


@router.get("/{asset_id}/parsed-data")
def get_asset_parsed_data(
    asset_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """获取资产的解析数据（直接从DB读取，无需文件IO）"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_view(asset.tenant_id, accessible, _is_admin(current_user))
    if asset.parsed_data:
        return {"source": "database", "data": asset.parsed_data}
    # 如果DB中没有，尝试现场解析并回填
    if not os.path.exists(asset.file_path):
        raise HTTPException(status_code=404, detail="文件不存在且无缓存数据")
    parsed = _parse_full_data(asset.file_path)
    if parsed:
        asset.parsed_data = parsed
        db.commit()
    return {"source": "file_parsed", "data": parsed}


@router.put("/{asset_id}")
def update_asset(
    asset_id: int,
    data: AssetUpdateIn,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """更新资产信息"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_manage(asset.tenant_id, accessible, _is_admin(current_user))

    from datetime import date
    for field, value in data.model_dump(exclude_none=True).items():
        if field in ("effective_from", "effective_to") and value:
            value = date.fromisoformat(value)
        setattr(asset, field, value)

    asset.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(asset)
    return _asset_to_out(asset)


@router.delete("/{asset_id}")
def delete_asset(
    asset_id: int,
    hard: bool = Query(False),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """停用或物理删除资产"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_manage(asset.tenant_id, accessible, _is_admin(current_user))

    if hard:
        # 物理删除文件 + 数据库记录
        if os.path.exists(asset.file_path):
            os.remove(asset.file_path)
        db.delete(asset)
    else:
        asset.is_active = False
        asset.updated_at = datetime.utcnow()
    db.commit()
    return {"message": "删除成功" if hard else "已停用"}


@router.post("/{asset_id}/new-version", status_code=202)
async def upload_new_version(
    asset_id: int,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
    accessible: list = Depends(get_accessible_tenants),
):
    """上传新版本：流式落盘后交给同一个单并发子进程队列。"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    _assert_can_manage(asset.tenant_id, accessible, _is_admin(current_user))

    task = await _create_asset_upload_task(
        uploads=[file], tenant_id=asset.tenant_id, asset_type=asset.asset_type,
        category_id=asset.category_id, description=asset.description,
        effective_from=str(asset.effective_from) if asset.effective_from else None,
        effective_to=str(asset.effective_to) if asset.effective_to else None,
        tags=json.dumps(asset.tags, ensure_ascii=False) if asset.tags is not None else None,
        uploaded_by=current_user.id, db=db, names=[asset.name],
        previous_asset_id=asset.id,
    )
    return _upload_task_to_out(task)
