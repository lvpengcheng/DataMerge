"""智能组表：规则文件管理与匹配知识库 API。

- 规则文件：scope=global 存 global_assets/assemble_rules/，scope=tenant 存 tenants/{tenant_id}/assemble_rules/
- 文件命名 {rule_id}_{原始文件名} 防重
- 知识库（assemble_field_mappings）：查看/删除/停用恢复，防错误映射污染
"""

import os
import shutil
import logging
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Query
from fastapi.responses import FileResponse

from ..auth.dependencies import require_permission
from ..database.connection import SessionLocal
from ..database.models import AssembleRule, AssembleFieldMapping

router = APIRouter(prefix="/api/assemble", tags=["智能组表"])

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
GLOBAL_RULES_DIR = PROJECT_ROOT / "global_assets" / "assemble_rules"
TENANTS_DIR = PROJECT_ROOT / "tenants"

ALLOWED_EXTS = {".md", ".txt", ".pdf", ".docx", ".doc", ".xlsx", ".xls"}


def _rule_dir(scope: str, tenant_id: Optional[str]) -> Path:
    """规则文件目录：global → global_assets/assemble_rules/，tenant → tenants/{id}/assemble_rules/"""
    if scope == "tenant":
        d = TENANTS_DIR / str(tenant_id or "") / "assemble_rules"
    else:
        d = GLOBAL_RULES_DIR
    d.mkdir(parents=True, exist_ok=True)
    return d


def _rule_to_dict(r: AssembleRule) -> dict:
    return {
        "id": r.id,
        "name": r.name,
        "scope": r.scope,
        "tenant_id": r.tenant_id,
        "file_names": r.file_names or [],
        "description": r.description or "",
        "uploader_name": r.uploader.display_name if r.uploader else "",
        "created_at": r.created_at.isoformat() if r.created_at else None,
        "updated_at": r.updated_at.isoformat() if r.updated_at else None,
    }


# ==================== 规则文件 CRUD ====================

@router.get("/rules")
async def assemble_rules_list(
    scope: str = Query("all"),            # all(管理端) / available(执行端: 全局+指定租户)
    tenant_id: str = Query("", description="scope=available 时的租户"),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """规则文件列表。

    scope=available：全局规则 + 该租户规则（执行页签下拉用）。
    scope=all：管理端全部。
    """
    db = SessionLocal()
    try:
        q = db.query(AssembleRule)
        if scope == "available":
            t = (tenant_id or "").strip()
            if not t:
                # 空租户：只看全局规则
                q = q.filter(AssembleRule.scope == "global")
            else:
                q = q.filter((AssembleRule.scope == "global") | (
                    (AssembleRule.scope == "tenant") & (AssembleRule.tenant_id == t)))
        items = [_rule_to_dict(r) for r in q.order_by(AssembleRule.updated_at.desc()).all()]
        return {"items": items}
    finally:
        db.close()


@router.post("/rules")
async def assemble_rules_create(
    files: List[UploadFile] = File(...),
    name: str = Form(""),
    scope: str = Form("global"),
    tenant_id: str = Form(""),
    description: str = Form(""),
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """上传规则文件（支持多文件）。scope=global 全局大规则 / scope=tenant 按租户规则。"""
    scope = (scope or "global").strip()
    if scope not in ("global", "tenant"):
        raise HTTPException(status_code=400, detail="scope 必须是 global 或 tenant")
    t = (tenant_id or "").strip()
    if scope == "tenant" and not t:
        raise HTTPException(status_code=400, detail="按租户规则必须选择租户")
    if not files or all(f.filename is None for f in files):
        raise HTTPException(status_code=400, detail="至少上传一个规则文件")

    db = SessionLocal()
    try:
        row = AssembleRule(
            name=name.strip() or files[0].filename or "未命名规则",
            scope=scope,
            tenant_id=t if scope == "tenant" else None,
            description=description,
            uploader_id=current_user.id,
        )
        db.add(row)
        db.commit()
        db.refresh(row)

        # 落盘：{rule_id}_{原始文件名}
        saved = []
        d = _rule_dir(scope, t)
        for f in files:
            orig = Path(f.filename or "").name or "rule"
            if Path(orig).suffix.lower() not in ALLOWED_EXTS:
                continue
            dest = d / f"{row.id}_{orig}"
            with open(dest, "wb") as out:
                out.write(f.file.read())
            saved.append(orig)

        if not saved:
            db.delete(row)
            db.commit()
            raise HTTPException(status_code=400, detail="没有可用的规则文件（支持 md/txt/pdf/docx/xlsx）")

        row.file_names = saved
        db.commit()
        return _rule_to_dict(row)
    finally:
        db.close()


@router.post("/rules/{rule_id}/upload")
async def assemble_rules_reupload(
    rule_id: int,
    files: List[UploadFile] = File(...),
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """重新上传：新文件替换原规则文件（保留规则 id/名称/作用域，删旧文件换新文件）。"""
    if not files or all(f.filename is None for f in files):
        raise HTTPException(status_code=400, detail="至少选择一个规则文件")
    db = SessionLocal()
    try:
        r = db.query(AssembleRule).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        d = _rule_dir(r.scope, r.tenant_id)

        # 删旧文件
        for n in r.file_names or []:
            try:
                p = d / f"{r.id}_{n}"
                if p.exists():
                    os.unlink(p)
            except Exception:
                logger.warning("[assemble] 删除旧规则文件失败: %s", n)

        # 存新文件
        saved = []
        for f in files:
            orig = Path(f.filename or "").name or "rule"
            if Path(orig).suffix.lower() not in ALLOWED_EXTS:
                continue
            dest = d / f"{r.id}_{orig}"
            with open(dest, "wb") as out:
                out.write(f.file.read())
            saved.append(orig)
        if not saved:
            raise HTTPException(status_code=400, detail="没有可用的规则文件（支持 md/txt/pdf/docx/xlsx）")

        r.file_names = saved
        db.commit()
        return _rule_to_dict(r)
    finally:
        db.close()


@router.get("/rules/{rule_id}/download")
async def assemble_rules_download(
    rule_id: int,
    file_name: str = Query("", description="要下载的文件名（多文件时指定）"),
    current_user=Depends(require_permission("tools.assemble", "tools.assemble.manage")),
):
    """下载规则文件。多文件规则未指定 file_name 时下载第一个。"""
    db = SessionLocal()
    try:
        r = db.query(AssembleRule).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        names = r.file_names or []
        if not names:
            raise HTTPException(status_code=404, detail="规则没有文件")
        target = file_name if file_name in names else names[0]
        p = _rule_dir(r.scope, r.tenant_id) / f"{r.id}_{target}"
        if not p.exists():
            raise HTTPException(status_code=404, detail="文件不存在")
        return FileResponse(
            p, filename=target,
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(target)}"},
        )
    finally:
        db.close()


@router.delete("/rules/{rule_id}")
async def assemble_rules_delete(
    rule_id: int,
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """删除规则 + 物理文件。

    ⚠️ 先解除任务历史的 rule_id 引用（assemble_tasks.rule_id 外键无 ondelete，
    直接删被引用规则会触发 FK violation 报错）；任务历史本身保留。
    """
    from ..database.models import AssembleTask
    db = SessionLocal()
    try:
        r = db.query(AssembleRule).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        d = _rule_dir(r.scope, r.tenant_id)
        names = r.file_names or []
        # 解除任务关联后再删（历史任务保留，仅 rule_id 置空）
        db.query(AssembleTask).filter(AssembleTask.rule_id == rule_id).update(
            {AssembleTask.rule_id: None})
        db.delete(r)
        db.commit()
        for n in names:
            p = d / f"{r.id}_{n}"
            try:
                if p.exists():
                    os.unlink(p)
            except Exception:
                logger.warning("[assemble] 删除规则文件失败: %s", p)
        return {"ok": True}
    finally:
        db.close()


# ==================== 匹配知识库管理 ====================

@router.get("/mappings")
async def assemble_mappings_list(
    keyword: str = Query("", description="源列/目标列模糊搜索"),
    status: str = Query("", description="all/active/review_needed"),
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """字段级匹配知识库列表（管理端）。"""
    db = SessionLocal()
    try:
        q = db.query(AssembleFieldMapping)
        kw = keyword.strip()
        if kw:
            q = q.filter(
                (AssembleFieldMapping.source_column.like(f"%{kw}%"))
                | (AssembleFieldMapping.target_column.like(f"%{kw}%"))
            )
        if status in ("active", "review_needed", "pending"):
            q = q.filter(AssembleFieldMapping.status == status)
        rows = q.order_by(AssembleFieldMapping.updated_at.desc()).limit(1000).all()
        items = [{
            "id": m.id,
            "tenant_id": m.tenant_id,
            "source_column": m.source_column,
            "target_column": m.target_column,
            "template_signature": m.template_signature,
            "match_type": m.match_type,
            "confirm_count": m.confirm_count,
            "hit_count": m.hit_count,
            "status": m.status,
            "created_at": m.created_at.isoformat() if m.created_at else None,
            "updated_at": m.updated_at.isoformat() if m.updated_at else None,
        } for m in rows]
        return {"items": items}
    finally:
        db.close()


@router.delete("/mappings/{mapping_id}")
async def assemble_mappings_delete(
    mapping_id: int,
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """删除知识库条目。"""
    db = SessionLocal()
    try:
        m = db.query(AssembleFieldMapping).filter_by(id=mapping_id).first()
        if not m:
            raise HTTPException(status_code=404, detail="映射条目不存在")
        db.delete(m)
        db.commit()
        return {"ok": True}
    finally:
        db.close()


@router.put("/mappings/{mapping_id}")
async def assemble_mappings_update(
    mapping_id: int,
    status: str = Form(""),
    current_user=Depends(require_permission("tools.assemble.manage")),
):
    """停用/恢复/候选知识库条目：status=active / review_needed / pending。"""
    if status not in ("active", "review_needed", "pending"):
        raise HTTPException(status_code=400, detail="status 必须是 active / review_needed / pending")
    from ..assemble.assemble_engine import CONFIRM_THRESHOLD
    db = SessionLocal()
    try:
        m = db.query(AssembleFieldMapping).filter_by(id=mapping_id).first()
        if not m:
            raise HTTPException(status_code=404, detail="映射条目不存在")
        m.status = status
        if status == "active":
            # 手动置 active 时补齐确认次数，确保能被查询自动采用
            if (m.confirm_count or 0) < CONFIRM_THRESHOLD:
                m.confirm_count = CONFIRM_THRESHOLD
            m.match_type = "anchored"
        db.commit()
        return {"id": m.id, "status": m.status, "confirm_count": m.confirm_count}
    finally:
        db.close()
