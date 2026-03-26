"""
数据资产管理 API
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional, List

from fastapi import APIRouter, Depends, UploadFile, File, Form, Query, HTTPException
from sqlalchemy.orm import Session
from pydantic import BaseModel

from ..database.connection import get_db
from ..database.models import DataAsset, ReferenceCategory, TenantAuthorization
from ..auth.dependencies import get_current_user

router = APIRouter(prefix="/api/assets", tags=["数据资产"])

# 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


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
    assets = q.order_by(DataAsset.created_at.desc()).all()
    return [_asset_to_out(a) for a in assets]


@router.get("/tenants")
def list_available_tenants(
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取可用租户列表（用于上传基础数据时选择作用域）"""
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
    assets = q.order_by(DataAsset.created_at.desc()).all()
    return [_asset_to_out(a) for a in assets]


@router.post("/upload")
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
):
    """上传文件 → 自动解析 sheet 结构 → 存入 data_assets"""
    # 保存文件
    storage_dir = _get_asset_storage_dir(tenant_id, asset_type)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = f"{timestamp}_{file.filename}"
    file_path = storage_dir / safe_name

    content = await file.read()
    with open(file_path, "wb") as f:
        f.write(content)

    # 解析 sheet 结构
    sheet_summary = _parse_sheet_summary(str(file_path))

    # 基础数据：解析完整数据存入DB，后续计算可直接从DB读取
    parsed_data = None
    if asset_type == "reference":
        parsed_data = _parse_full_data(str(file_path))

    # 解析标签
    import json
    parsed_tags = None
    if tags:
        try:
            parsed_tags = json.loads(tags)
        except Exception:
            parsed_tags = [t.strip() for t in tags.split(",") if t.strip()]

    # 解析日期
    from datetime import date
    eff_from = date.fromisoformat(effective_from) if effective_from else None
    eff_to = date.fromisoformat(effective_to) if effective_to else None

    asset = DataAsset(
        tenant_id=tenant_id,
        asset_type=asset_type,
        category_id=category_id,
        name=name or file.filename,
        description=description,
        file_path=str(file_path),
        file_name=file.filename,
        file_size=len(content),
        sheet_summary=sheet_summary,
        parsed_data=parsed_data,
        effective_from=eff_from,
        effective_to=eff_to,
        uploaded_by=current_user.id,
        tags=parsed_tags,
    )
    db.add(asset)
    db.commit()
    db.refresh(asset)
    return _asset_to_out(asset)


# ==================== 动态路由 /{asset_id} ====================

@router.get("/{asset_id}")
def get_asset(
    asset_id: int,
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """获取单个资产详情"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    return _asset_to_out(asset)


@router.get("/{asset_id}/download")
def download_asset(
    asset_id: int,
    format: Optional[str] = Query(None, description="下载格式: 空=原始, pdf, encrypted"),
    password: Optional[str] = Query(None, description="加密密码（encrypted格式用）"),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """下载资产文件（支持原始/PDF/加密Excel）"""
    from fastapi.responses import FileResponse
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
    if not os.path.exists(asset.file_path):
        raise HTTPException(status_code=404, detail="文件不存在")

    # 原始下载
    if not format:
        return FileResponse(
            path=asset.file_path,
            filename=asset.file_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
):
    """预览资产数据（前 N 行）"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
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
):
    """获取资产的解析数据（直接从DB读取，无需文件IO）"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")
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
):
    """更新资产信息"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")

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
):
    """停用或物理删除资产"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")

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


@router.post("/{asset_id}/new-version")
async def upload_new_version(
    asset_id: int,
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    current_user=Depends(get_current_user),
):
    """上传新版本（保留历史版本）"""
    asset = db.query(DataAsset).filter_by(id=asset_id).first()
    if not asset:
        raise HTTPException(status_code=404, detail="资产不存在")

    # 保存新文件
    storage_dir = Path(asset.file_path).parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = f"{timestamp}_{file.filename}"
    file_path = storage_dir / safe_name

    content = await file.read()
    with open(file_path, "wb") as f:
        f.write(content)

    # 创建新版本记录
    new_parsed_data = _parse_full_data(str(file_path)) if asset.asset_type == "reference" else None
    new_asset = DataAsset(
        tenant_id=asset.tenant_id,
        asset_type=asset.asset_type,
        category_id=asset.category_id,
        name=asset.name,
        description=asset.description,
        file_path=str(file_path),
        file_name=file.filename,
        file_size=len(content),
        sheet_summary=_parse_sheet_summary(str(file_path)),
        parsed_data=new_parsed_data,
        version=asset.version + 1,
        effective_from=asset.effective_from,
        effective_to=asset.effective_to,
        uploaded_by=current_user.id,
        tags=asset.tags,
    )
    db.add(new_asset)

    # 停用旧版本
    asset.is_active = False
    asset.updated_at = datetime.utcnow()

    db.commit()
    db.refresh(new_asset)
    return _asset_to_out(new_asset)
