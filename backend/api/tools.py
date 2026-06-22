"""智能小工具 API"""

import io
import os
import sys
import json
import uuid
import shutil
import tempfile
import zipfile
import logging
import hashlib
from pathlib import Path
from typing import List, Optional

import pandas as pd
from fastapi import APIRouter, Depends, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from ..auth.dependencies import get_current_user

router = APIRouter(prefix="/api/tools", tags=["智能小工具"])

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
MERGE_SESSION_ROOT = PROJECT_ROOT / "temp" / "merge_sessions"


def _import_split_one_file():
    """延迟导入根目录的 split_by_banner.split_one_file"""
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    from split_by_banner import split_one_file
    return split_one_file


@router.post("/split-by-banner")
async def split_by_banner(
    files: List[UploadFile] = File(...),
    current_user=Depends(get_current_user),
):
    """按 banner 拆分 sheet:接收多个 Excel,逐个拆分,打包为 zip 返回"""
    if not files:
        raise HTTPException(status_code=400, detail="未上传文件")

    split_one_file = _import_split_one_file()
    work_dir = Path(tempfile.mkdtemp(prefix="split_banner_"))
    src_dir = work_dir / "src"
    out_dir = work_dir / "out"
    src_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    errors: List[str] = []
    success_files: List[Path] = []

    try:
        for uf in files:
            name = uf.filename or "unnamed.xlsx"
            ext = Path(name).suffix.lower()
            if ext not in EXCEL_EXTS:
                errors.append(f"{name}: 不支持的扩展名({ext})")
                continue
            src_path = src_dir / name
            try:
                content = await uf.read()
                src_path.write_bytes(content)
            except Exception as e:
                errors.append(f"{name}: 写入失败 {e}")
                continue

            out_name = f"{src_path.stem}_split.xlsx"
            out_path = out_dir / out_name
            try:
                split_one_file(src_path, out_path)
                if out_path.exists():
                    success_files.append(out_path)
                else:
                    errors.append(f"{name}: 拆分未生成输出")
            except Exception as e:
                logger.exception(f"拆分失败: {name}")
                errors.append(f"{name}: {e}")

        if not success_files:
            detail = "全部失败:\n" + "\n".join(errors) if errors else "未生成任何输出"
            raise HTTPException(status_code=400, detail=detail)

        # 打包 zip
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in success_files:
                zf.write(p, arcname=p.name)
            if errors:
                zf.writestr("_errors.txt", "\n".join(errors))
        buf.seek(0)

        # 同步清理(StreamingResponse 已读完 buffer)
        shutil.rmtree(work_dir, ignore_errors=True)

        headers = {"Content-Disposition": 'attachment; filename="split_results.zip"'}
        if errors:
            headers["X-Split-Errors"] = str(len(errors))
        return StreamingResponse(buf, media_type="application/zip", headers=headers)

    except HTTPException:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        logger.exception("拆分接口异常")
        raise HTTPException(status_code=500, detail=str(e))


# ==================== 多表数据合并 ====================

def _parse_file_to_df(path: str):
    """用 excel_parser 解析单文件（激活 sheet、最优区域、先算公式），返回 {sheet, columns, df}。

    df 列名为表头名（excel_parser 的 region.data 按列字母键，这里用 head_data 反查改回表头名）。
    """
    from excel_parser import IntelligentExcelParser
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(
        path, read_formulas=False, calculate_formulas=True,
        active_sheet_only=True, best_region_only=True,
    )
    if not results or not results[0].regions:
        return {"sheet": "", "columns": [], "df": pd.DataFrame()}
    sd = results[0]
    region = sd.regions[0]
    head = region.head_data or {}                       # {表头: 列字母}
    letter_to_header = {v: k for k, v in head.items()}
    df = pd.DataFrame(region.data or [])                # 列名为列字母
    if not df.empty:
        df = df.rename(columns=letter_to_header)
        ordered = [h for h in head.keys() if h in df.columns]
        if ordered:
            df = df[ordered]
    return {"sheet": sd.sheet_name, "columns": list(head.keys()), "df": df}


def _session_dir(session_id: str) -> Path:
    # 防目录穿越
    safe = "".join(ch for ch in session_id if ch.isalnum() or ch in "-_")
    return MERGE_SESSION_ROOT / safe


class MergeExecuteRequest(BaseModel):
    session_id: str
    tenant_id: Optional[str] = None
    key_map: dict                       # {file_name: 主键列名}
    result_columns: List[dict]          # [{"name", "sources":[{"file","col"}]}]
    merge_mode: str = "union"           # union | base | conflict_only
    base_file: Optional[str] = None
    normalize_keys: bool = True
    ai_provider: Optional[str] = None   # 仅缓存写回时记录用


@router.post("/merge/analyze")
async def merge_analyze(
    files: List[UploadFile] = File(...),
    tenant_id: Optional[str] = Form(None),
    current_user=Depends(get_current_user),
):
    """多表合并第一步：上传多文件 → 解析 → 返回各文件列（带来源）+ 列头指纹/缓存命中。

    不在此处调用 AI：用户先选字段，再决定是否调 AI 匹配同义列（见 /merge/match）。
    """
    if not files or len(files) < 2:
        raise HTTPException(status_code=400, detail="请至少上传 2 个文件")

    from ..utils.merge_engine import compute_header_fingerprint
    from ..database.connection import SessionLocal
    from ..database.models import MergeFieldMapping

    session_id = uuid.uuid4().hex
    sdir = _session_dir(session_id)
    sdir.mkdir(parents=True, exist_ok=True)

    files_meta = []          # [{name, sheet, columns, fingerprint}]
    cache_hit_files = []

    try:
        db = SessionLocal()
        try:
            for uf in files:
                name = uf.filename or "unnamed.xlsx"
                if Path(name).suffix.lower() not in EXCEL_EXTS:
                    raise HTTPException(status_code=400, detail=f"{name}: 不支持的扩展名")
                dest = sdir / name
                dest.write_bytes(await uf.read())
                info = _parse_file_to_df(str(dest))
                cols = info["columns"]
                fp = compute_header_fingerprint(cols)
                files_meta.append({"name": name, "sheet": info["sheet"], "columns": cols, "fingerprint": fp})
                if tenant_id:
                    row = (db.query(MergeFieldMapping)
                             .filter_by(tenant_id=tenant_id, header_fingerprint=fp).first())
                    if row and isinstance(row.mapping, dict):
                        cache_hit_files.append(name)
        finally:
            db.close()

        # 落 meta，供 /merge/match 读列头指纹（无需重新解析）
        (sdir / "_meta.json").write_text(
            json.dumps({"files": files_meta}, ensure_ascii=False), encoding="utf-8"
        )
        return {"session_id": session_id, "files": files_meta, "cache_hit_files": cache_hit_files}
    except HTTPException:
        shutil.rmtree(sdir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(sdir, ignore_errors=True)
        logger.exception("merge/analyze 异常")
        raise HTTPException(status_code=500, detail=str(e))


class MergeMatchRequest(BaseModel):
    session_id: str
    tenant_id: Optional[str] = None
    selected: List[dict]                 # [{"file","col"}] 用户勾选的字段
    use_ai: bool = False                 # 是否调用 AI 匹配差异命名的同义列
    ai_provider: Optional[str] = "deepseek"


@router.post("/merge/match")
async def merge_match(req: MergeMatchRequest, current_user=Depends(get_current_user)):
    """对【已勾选字段】做同义列归组：精确同名 + 缓存（免 AI）；use_ai 时再调 AI 补差异命名。

    返回 groups（可合并为多源结果列的建议）；前端据此把同义列并成一个结果列（多源→可能冲突标红）。
    """
    from ..utils.merge_engine import suggest_field_groups
    from ..database.connection import SessionLocal
    from ..database.models import MergeFieldMapping

    sdir = _session_dir(req.session_id)
    meta_path = sdir / "_meta.json"
    if not meta_path.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")

    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    fp_by_file = {f["name"]: f["fingerprint"] for f in meta.get("files", [])}

    # 仅对勾选字段归组
    files_columns: dict = {}
    for s in req.selected or []:
        files_columns.setdefault(s["file"], []).append(s["col"])
    if not files_columns:
        return {"groups": [], "ai_suggestions": []}

    # 缓存（按各文件全列指纹），免 AI
    cache_map: dict = {}
    if req.tenant_id:
        db = SessionLocal()
        try:
            for fname in files_columns:
                fp = fp_by_file.get(fname)
                if not fp:
                    continue
                row = (db.query(MergeFieldMapping)
                         .filter_by(tenant_id=req.tenant_id, header_fingerprint=fp).first())
                if row and isinstance(row.mapping, dict):
                    cache_map[fname] = row.mapping
        finally:
            db.close()

    suggestion = suggest_field_groups(
        files_columns,
        ai_provider_name=(req.ai_provider if req.use_ai else None),
        cache_map=cache_map,
    )
    return suggestion


@router.post("/merge/execute")
async def merge_execute(req: MergeExecuteRequest, current_user=Depends(get_current_user)):
    """多表合并第二步：按确认的映射归并 → 标红冲突 → 返回 xlsx；并写回字段匹配缓存。"""
    from ..utils.merge_engine import merge_tables, write_merged_xlsx, compute_header_fingerprint
    from ..database.connection import SessionLocal
    from ..database.models import MergeFieldMapping

    sdir = _session_dir(req.session_id)
    if not sdir.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")

    try:
        # 重新解析会话内文件 → {file: {df}}
        parsed_files = {}
        files_columns = {}
        for fp_path in sorted(sdir.iterdir()):
            if fp_path.suffix.lower() not in EXCEL_EXTS:
                continue
            info = _parse_file_to_df(str(fp_path))
            parsed_files[fp_path.name] = {"df": info["df"]}
            files_columns[fp_path.name] = info["columns"]

        if not parsed_files:
            raise HTTPException(status_code=400, detail="会话内无有效文件")

        result = merge_tables(
            parsed_files=parsed_files,
            key_map=req.key_map,
            result_columns=req.result_columns,
            merge_mode=req.merge_mode,
            base_file=req.base_file,
            normalize_keys=req.normalize_keys,
        )

        out_path = sdir / "合并结果.xlsx"
        write_merged_xlsx(result, str(out_path))
        data = out_path.read_bytes()

        # 写回缓存：每文件 {源列 -> 规范字段名}
        if req.tenant_id:
            db = SessionLocal()
            try:
                # 由 result_columns 反推每文件的 列->规范名
                per_file_map = {}
                for rc in req.result_columns:
                    for src in rc.get("sources", []):
                        per_file_map.setdefault(src["file"], {})[src["col"]] = rc["name"]
                for fname, mp in per_file_map.items():
                    fp = compute_header_fingerprint(files_columns.get(fname, []))
                    row = (db.query(MergeFieldMapping)
                             .filter_by(tenant_id=req.tenant_id, header_fingerprint=fp).first())
                    if row:
                        row.mapping = mp
                    else:
                        db.add(MergeFieldMapping(tenant_id=req.tenant_id, header_fingerprint=fp, mapping=mp))
                db.commit()
            except Exception as ce:
                logger.warning(f"[merge] 缓存写回失败（不阻断）: {ce}")
            finally:
                db.close()

        buf = io.BytesIO(data)
        buf.seek(0)
        shutil.rmtree(sdir, ignore_errors=True)
        headers = {
            "Content-Disposition": 'attachment; filename="merged_result.xlsx"',
            "X-Merge-Conflicts": str(result["report"].get("conflict_ids", 0)),
            "X-Merge-Rows": str(result["report"].get("output_rows", 0)),
        }
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("merge/execute 异常")
        raise HTTPException(status_code=500, detail=str(e))
