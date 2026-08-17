"""智能小工具 API"""

import io
import os
import re
import sys
import json
import uuid
import shutil
import tempfile
import zipfile
import logging
import hashlib
import asyncio
from pathlib import Path
from typing import List, Optional
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import pandas as pd
from fastapi import APIRouter, Depends, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

from ..auth.dependencies import get_current_user, require_permission
from ..database.connection import SessionLocal
from ..database.models import SopEntry, SopRound, SopRuleFile
from ..ai_engine.document_parser import DocumentParser

router = APIRouter(prefix="/api/tools", tags=["智能小工具"])

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
MERGE_SESSION_ROOT = PROJECT_ROOT / "temp" / "merge_sessions"
INTEGRATE_SESSION_ROOT = PROJECT_ROOT / "temp" / "integrate_sessions"
# 方案模版文件存储根目录（按 owner 分子目录），tenants/ 已 gitignore
MERGE_TPL_ROOT = PROJECT_ROOT / "tenants" / "__tools_merge__" / "merge_templates"


def _safe_seg(s: str) -> str:
    """路径分段安全化：仅保留字母数字和 -_，防目录穿越/非法字符（如 owner 里的冒号）。"""
    return "".join(ch for ch in str(s) if ch.isalnum() or ch in "-_") or "x"


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

def _regions_to_df(sd):
    """把一个 sheet 的多个【同结构】区域纵向拼接成一张 df。

    同结构判定按【列位置(字母)集合一致】——堆叠的多块通常占同一批列(A..N)，即使某列表头文字
    不同(如把人名/基本工资写进表头)也算同结构；据此按【列字母对齐】拼接，再统一用基准块的表头名，
    列不会错位。列位置不同的区域(汇总/说明块)忽略，避免乱拼。单区域文件行为与旧版一致(零回归)。
    像"按人分块堆叠"的文件(每人一块=一个区域)，旧版 best_region_only 只取一块会漏掉其余人的行
    (如某人的"2月26日"落在第二块)，此处把所有同结构块都读进来，按日期/姓名做主键才能命中全部。
    """
    regions = [r for r in (sd.regions or []) if (r.head_data or r.data)]
    if not regions:
        return {"sheet": sd.sheet_name, "columns": [], "df": pd.DataFrame()}
    # 基准：列数最多、其次行数最多的区域，定义列位置(字母)->表头名 及列顺序
    base = max(regions, key=lambda r: (len(r.head_data or {}), len(r.data or [])))
    base_head = base.head_data or {}                    # {表头: 字母}
    base_l2h = {v: k for k, v in base_head.items()}     # {字母: 表头}
    base_letters = frozenset(base_head.values())
    base_cols = list(base_head.keys())                  # 表头名(基准顺序)
    frames, n = [], 0
    for region in regions:
        if frozenset((region.head_data or {}).values()) != base_letters:
            continue                                    # 列位置不同 → 非同结构块
        n += 1
        df = pd.DataFrame(region.data or [])            # 列名=列字母
        if not df.empty:
            frames.append(df)
    if frames:
        out_df = pd.concat(frames, ignore_index=True)   # 按列字母对齐拼接
        out_df = out_df.rename(columns=base_l2h)        # 字母 → 基准表头名
        ordered = [c for c in base_cols if c in out_df.columns]
        if ordered:
            out_df = out_df[ordered]
    else:
        out_df = pd.DataFrame(columns=base_cols)
    if n > 1:
        logger.info(f"[parse] sheet「{sd.sheet_name}」按列对齐合并 {n} 个同结构区域 → 共 {len(out_df)} 行")
    return {"sheet": sd.sheet_name, "columns": base_cols, "df": out_df}


def _parse_file_to_df_impl(path: str):
    """（子进程执行体）用 excel_parser 解析单文件（激活 sheet、先算公式），返回 {sheet, columns, df}。

    读【全部区域】：同结构的多个区域（如按人堆叠的块）纵向拼接成一张 df，避免只命中一块。
    df 列名为表头名（excel_parser 的 region.data 按列字母键，反查 head_data 改回表头名）。
    """
    from excel_parser import IntelligentExcelParser
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(
        path, read_formulas=False, calculate_formulas=True,
        active_sheet_only=True, best_region_only=False,
    )
    if not results or not results[0].regions:
        return {"sheet": "", "columns": [], "df": pd.DataFrame()}
    return _regions_to_df(results[0])


def _parse_file_to_df(path: str):
    """用 excel_parser 解析单文件（激活 sheet、先算公式），返回 {sheet, columns, df}。

    在【独立子进程】执行：某些文件（公式密集/超大）会让 Aspose 解析长时间计算、
    内存暴涨 → VM 假死（swap 风暴）。子进程超时/超内存会被强杀，主进程安全。
    失败返回空 df（与 parse_excel_file 失败时的旧行为一致），详情记日志。
    """
    from backend.utils.subprocess_runner import run_in_subprocess, default_max_memory_mb, default_timeout

    r = run_in_subprocess(
        "backend.api.tools:_parse_file_to_df_impl", (str(path),),
        timeout=default_timeout("parse"), max_memory_mb=default_max_memory_mb(),
    )
    if r.success:
        return r.result
    reason = "超时" if r.timed_out else ("内存超限" if r.killed_by_memory else r.error)
    logger.error(f"[parse] 子进程解析失败（{reason}）: {path}")
    return {"sheet": "", "columns": [], "df": pd.DataFrame()}


def _session_dir(session_id: str) -> Path:
    # 防目录穿越
    safe = "".join(ch for ch in session_id if ch.isalnum() or ch in "-_")
    return MERGE_SESSION_ROOT / safe


class MergeExecuteRequest(BaseModel):
    session_id: str
    tenant_id: Optional[str] = None
    key_map: dict                       # {file_name: 主键列名 或 主键列名列表(复合主键)}
    result_columns: List[dict]          # [{"name", "sources":[{"file","col"}]}]
    merge_mode: str = "union"           # union | base | conflict_only
    base_file: Optional[str] = None
    normalize_keys: bool = True
    date_key_mode: str = "yearmonthday"  # 日期主键归一：off|yearmonthday|yearmonth|month|day（默认按年月日）
    ai_provider: Optional[str] = None   # 仅缓存写回时记录用
    template_id: Optional[int] = None   # 套用的方案 id；该方案带模版文件时按模版填充


class MergeSkeletonRequest(BaseModel):
    result_columns: List[dict]          # [{"name", ...}]  有序，输出列顺序
    session_id: Optional[str] = None    # 会话 id（配合 base_file 以某上传表为模版基准）
    base_file: Optional[str] = None     # 以哪个上传表为模版基准；空/找不到则回退空白骨架


def _blank_skeleton_bytes(names):
    """无基准表时：新建一个空白骨架（列名 + &=DT.列名 标记）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "模版"
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for ci, name in enumerate(names, start=1):
        h = ws.cell(row=1, column=ci, value=name)
        h.font = Font(bold=True)
        h.fill = header_fill
        ws.cell(row=2, column=ci, value=f"&=DT.{name}")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@router.post("/merge/template-skeleton")
async def merge_template_skeleton(req: MergeSkeletonRequest, current_user=Depends(get_current_user)):
    """生成模版骨架 xlsx：数据行写 `&=DT.列名` 标记，供用户下载编辑后作为方案模版上传。

    优先【以选中的某个上传表为基准】：保留其标题/表头行格式与列宽，删掉全部数据行，
    在表头下第一行按【选中结果列的顺序】写 `&=DT.列名` 标记（同名列复用其表头格式/列宽，
    新增列取对应位置格式兜底）。无基准表或解析失败时，回退到空白骨架。
    """
    names = [str(rc.get("name") or "").strip() for rc in (req.result_columns or [])]
    names = [n for n in names if n]
    if not names:
        raise HTTPException(status_code=400, detail="没有可用的结果列")

    buf = None
    # 尝试以基准表生成
    if req.session_id and req.base_file:
        base_path = _session_dir(req.session_id) / req.base_file
        if base_path.exists():
            try:
                buf = _skeleton_from_base(str(base_path), names)
            except Exception as e:
                logger.warning(f"[merge] 基准表生成骨架失败，回退空白骨架: {e}")
                buf = None
    if buf is None:
        buf = _blank_skeleton_bytes(names)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="merge_template_skeleton.xlsx"'},
    )


def _skeleton_from_base(base_path: str, names: List[str]) -> io.BytesIO:
    """以基准表为底生成骨架：保留标题/表头格式，删数据行，按 names 顺序写表头+`&=DT.列名` 标记。"""
    import copy as _copy
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from excel_parser import IntelligentExcelParser

    # 1) 解析定位表头行
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(
        base_path, read_formulas=False, calculate_formulas=False,
        active_sheet_only=True, best_region_only=True,
    )
    if not results or not results[0].regions:
        raise ValueError("基准表无法解析出表头区域")
    region = results[0].regions[0]
    sheet_name = getattr(results[0], "sheet_name", None)
    hr = region.head_row_end or region.head_row_start or 1   # 表头行(1-indexed)
    ds = region.data_row_start or (hr + 1)                    # 首个数据行(1-indexed)

    # 2) openpyxl 打开（保留格式），定位到解析所用的 sheet
    wb = load_workbook(base_path)
    ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active

    # 3) 快照：表头行各列样式/列宽 + 首个数据行各列样式（按列名 & 按位置）。
    #    标记行(数据行)必须用【基准数据行】的格式，否则会串入无关格式——
    #    例如把数值列显示成日期(1900/1/5)。
    base_by_name = {}
    base_pos = {}
    data_fmt_by_name = {}   # 列名 -> 数据行单元格 _style（标记行套用，保证数值列不被显示成日期）
    maxc = ws.max_column or len(names)
    ds_valid = ds if (ds and ds <= ws.max_row) else None
    for c in range(1, maxc + 1):
        L = get_column_letter(c)
        cell = ws.cell(hr, c)
        style = _copy.copy(cell._style)
        width = ws.column_dimensions[L].width if L in ws.column_dimensions else None
        base_pos[c] = (style, width)
        v = cell.value
        nm = str(v).strip() if (v is not None and str(v).strip()) else None
        if nm:
            base_by_name[nm] = (style, width)
            if ds_valid is not None:
                data_fmt_by_name[nm] = _copy.copy(ws.cell(ds_valid, c)._style)

    # 4) 删除表头以下所有行（数据/汇总/页脚），只留标题+表头
    if ws.max_row > hr:
        ws.delete_rows(hr + 1, ws.max_row - hr)
    marker_row = hr + 1

    # 5) 按选中列顺序重写表头 + 写 &=DT.列名 标记；表头复用基准表头格式，
    #    标记(数据)行复用基准【数据行】格式（数值列不会被显示成日期）。
    N = len(names)
    for j, name in enumerate(names, start=1):
        L = get_column_letter(j)
        try:
            hcell = ws.cell(hr, j)
            hcell.value = name
            st = base_by_name.get(name) or base_pos.get(j)
            if st:
                try:
                    hcell._style = _copy.copy(st[0])
                except Exception:
                    pass
                if st[1] is not None:
                    ws.column_dimensions[L].width = st[1]
            mcell = ws.cell(marker_row, j)
            mcell.value = f"&=DT.{name}"
            # 标记行格式：仅取基准表【同名列】的数据行格式；匹配不到就强制 General。
            # 不按位置回退——重排后位置格式不可靠，易把数值列套上日期格式(1900/1/5)。
            dstyle = data_fmt_by_name.get(name)
            if dstyle is not None:
                try:
                    mcell._style = _copy.copy(dstyle)
                except Exception:
                    mcell.number_format = "General"
            else:
                mcell.number_format = "General"
        except Exception:
            pass
    # 清掉超出选中列数的多余表头单元格
    for c in range(N + 1, maxc + 1):
        try:
            ws.cell(hr, c).value = None
        except Exception:
            pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


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

    from ..utils.merge_engine import compute_header_fingerprint, guess_key_column
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
                # 上传规范化：误设日期格式的数字单元格重置为常规（避免被当日期读错）
                # 解析走子进程 + to_thread：不冻结事件循环（多客户并发时其他请求照常响应）
                try:
                    from ..utils.source_normalizer import normalize_misformatted_dates
                    await asyncio.to_thread(normalize_misformatted_dates, str(dest))
                except Exception as _ne:
                    logger.warning(f"[merge] 日期格式规范化失败（继续）: {_ne}")
                info = await asyncio.to_thread(_parse_file_to_df, str(dest))
                cols = info["columns"]
                fp = compute_header_fingerprint(cols)
                _sk = guess_key_column(cols, info.get("df")) or (cols[0] if cols else "")
                files_meta.append({"name": name, "sheet": info["sheet"], "columns": cols,
                                   "fingerprint": fp, "suggested_key": _sk})
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
            info = await asyncio.to_thread(_parse_file_to_df, str(fp_path))
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
            date_key_mode=req.date_key_mode,
        )

        out_path = sdir / "合并结果.xlsx"

        # 决定输出方式：套用的方案带模版文件 → Aspose SmartMarker 填充（带格式/公式）；
        # 否则 → 原 openpyxl 纯数据输出（行为不变）。
        tpl_abs = None
        if req.template_id:
            from ..database.models import MergeTemplate as _MT
            owner = _tpl_owner(current_user)
            _db = SessionLocal()
            try:
                _row = _db.query(_MT).filter_by(id=req.template_id, tenant_id=owner).first()
                _tf = (_row.config or {}).get("template_file") if _row else None
                if _tf:
                    _cand = (PROJECT_ROOT / _tf).resolve()
                    if _cand.exists():
                        tpl_abs = _cand
            finally:
                _db.close()

        if tpl_abs:
            df = pd.DataFrame(result["rows"], columns=result["columns"])
            # 追加列字母别名 A/B/C…（按结果列顺序），使模版可用 &=DT.A 或 &=DT.列名
            try:
                from openpyxl.utils import get_column_letter
                for _i, _col in enumerate(result["columns"]):
                    _L = get_column_letter(_i + 1)
                    if _L not in df.columns:
                        df[_L] = df[_col].values
            except Exception:
                pass
            from ..utils.aspose_helper import generate_from_template
            await asyncio.to_thread(
                generate_from_template, str(out_path), str(tpl_abs), {"DT": df}, mode="fill")
        else:
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


# ==================== 多表合并：命名模版（保存/复用，按用户私有） ====================

def _tpl_owner(current_user) -> str:
    """模版归属：当前登录用户。每个用户只看/管自己创建的模版。"""
    uid = getattr(current_user, "id", None)
    return f"user:{uid}" if uid is not None else "user:anon"


class MergeTemplateSaveRequest(BaseModel):
    name: str
    config: dict          # {key_field, merge_mode, normalize_keys, result_columns:[{name, source_cols:[...]}]}


@router.post("/merge/template/save")
async def merge_template_save(
    name: str = Form(...),
    config: str = Form(...),                       # JSON 字符串
    template: Optional[UploadFile] = File(None),   # 可选：带 &=DT.列名 标记的模版文件
    current_user=Depends(get_current_user),
):
    """保存/覆盖当前用户的命名合并方案（同名 upsert）。

    config 按列名存，便于下月复用。可选上传一个 Excel 模版文件（含 `&=DT.列名` 标记 + 用户
    自定义格式/公式）：文件存后台，路径记入 config.template_file，套用方案执行时按模版填充。
    不带文件时保留该方案已有的模版（若之前传过）。
    """
    name = (name or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="模版名称不能为空")
    try:
        cfg = json.loads(config) if isinstance(config, str) else (config or {})
        if not isinstance(cfg, dict):
            raise ValueError("config 不是对象")
    except Exception:
        raise HTTPException(status_code=400, detail="config 不是合法 JSON")

    if template is not None and Path(template.filename or "").suffix.lower() not in EXCEL_EXTS:
        raise HTTPException(status_code=400, detail="模版文件必须是 Excel(.xlsx/.xls/.xlsm)")

    from ..database.connection import SessionLocal
    from ..database.models import MergeTemplate
    owner = _tpl_owner(current_user)
    db = SessionLocal()
    try:
        row = db.query(MergeTemplate).filter_by(tenant_id=owner, name=name).first()
        if row is None:
            row = MergeTemplate(tenant_id=owner, name=name, config=cfg)
            db.add(row)
            db.commit()          # 先拿到自增 id 作为模版文件名
            db.refresh(row)
        # 保留旧模版引用（本次未重传时不丢）
        prev_cfg = row.config or {}
        if not template:
            if prev_cfg.get("template_file"):
                cfg["template_file"] = prev_cfg["template_file"]
                cfg["template_original_name"] = prev_cfg.get("template_original_name")
        else:
            tpl_dir = MERGE_TPL_ROOT / _safe_seg(owner)
            tpl_dir.mkdir(parents=True, exist_ok=True)
            ext = Path(template.filename or "tpl.xlsx").suffix.lower() or ".xlsx"
            tpl_path = tpl_dir / f"{row.id}{ext}"
            tpl_path.write_bytes(await template.read())
            cfg["template_file"] = str(tpl_path.relative_to(PROJECT_ROOT)).replace("\\", "/")
            cfg["template_original_name"] = template.filename
        row.config = cfg
        db.commit()
        return {"ok": True, "id": row.id, "name": name, "has_template": bool(cfg.get("template_file"))}
    except HTTPException:
        db.rollback()
        raise
    except Exception as e:
        db.rollback()
        logger.exception("merge/template/save 异常")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        db.close()


@router.get("/merge/templates")
async def merge_templates_list(current_user=Depends(get_current_user)):
    """列出当前用户自己的命名合并模版（含 config 供前端套用）。"""
    from ..database.connection import SessionLocal
    from ..database.models import MergeTemplate
    owner = _tpl_owner(current_user)
    db = SessionLocal()
    try:
        rows = (db.query(MergeTemplate)
                  .filter_by(tenant_id=owner)
                  .order_by(MergeTemplate.updated_at.desc()).all())
        return [{"id": r.id, "name": r.name, "config": r.config,
                 "has_template": bool((r.config or {}).get("template_file")),
                 "template_name": (r.config or {}).get("template_original_name"),
                 "updated_at": r.updated_at.isoformat() if r.updated_at else None} for r in rows]
    finally:
        db.close()


@router.delete("/merge/template/{tpl_id}")
async def merge_template_delete(tpl_id: int, current_user=Depends(get_current_user)):
    """删除当前用户自己的模版（只能删自己的）。"""
    from ..database.connection import SessionLocal
    from ..database.models import MergeTemplate
    owner = _tpl_owner(current_user)
    db = SessionLocal()
    try:
        row = db.query(MergeTemplate).filter_by(id=tpl_id, tenant_id=owner).first()
        if row:
            # 清理该方案关联的模版文件（忽略失败）
            _tf = (row.config or {}).get("template_file")
            if _tf:
                try:
                    (PROJECT_ROOT / _tf).unlink(missing_ok=True)
                except Exception:
                    pass
            db.delete(row)
            db.commit()
            return {"ok": True}
        raise HTTPException(status_code=404, detail="模版不存在或无权删除")
    finally:
        db.close()


@router.get("/merge/template/{tpl_id}/download")
async def merge_template_download(tpl_id: int, current_user=Depends(get_current_user)):
    """下载某方案已保存的模版文件（供再次编辑）。"""
    from ..database.connection import SessionLocal
    from ..database.models import MergeTemplate
    owner = _tpl_owner(current_user)
    db = SessionLocal()
    try:
        row = db.query(MergeTemplate).filter_by(id=tpl_id, tenant_id=owner).first()
        if not row:
            raise HTTPException(status_code=404, detail="方案不存在或无权访问")
        _tf = (row.config or {}).get("template_file")
        if not _tf or not (PROJECT_ROOT / _tf).exists():
            raise HTTPException(status_code=404, detail="该方案未保存模版文件")
        data = (PROJECT_ROOT / _tf).read_bytes()
        fname = (row.config or {}).get("template_original_name") or f"{row.name}_模版.xlsx"
    finally:
        db.close()
    buf = io.BytesIO(data)
    buf.seek(0)
    from urllib.parse import quote
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}"},
    )


# ==================== 多表整合对比（主表 A + 对照表 B）====================

def _integrate_session_dir(session_id: str) -> Path:
    safe = "".join(ch for ch in str(session_id) if ch.isalnum() or ch in "-_")
    return INTEGRATE_SESSION_ROOT / safe


def _parse_file_full_impl(path: str):
    """（子进程执行体）解析单文件激活页最优区域，额外返回区域坐标（供原地回填主表定位）。

    返回 {sheet, columns, df, head_data{表头->列字母}, data_row_start, data_row_end}（1-based 行）。
    """
    from excel_parser import IntelligentExcelParser
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(
        path, read_formulas=False, calculate_formulas=True,
        active_sheet_only=True, best_region_only=True,
    )
    if not results or not results[0].regions:
        return {"sheet": "", "columns": [], "df": pd.DataFrame(),
                "head_data": {}, "data_row_start": 0, "data_row_end": 0}
    sd = results[0]
    region = sd.regions[0]
    head = region.head_data or {}                       # {表头: 列字母}
    letter_to_header = {v: k for k, v in head.items()}
    df = pd.DataFrame(region.data or [])
    if not df.empty:
        df = df.rename(columns=letter_to_header)
        ordered = [h for h in head.keys() if h in df.columns]
        if ordered:
            df = df[ordered]
    return {
        "sheet": sd.sheet_name,
        "columns": list(head.keys()),
        "df": df,
        "head_data": dict(head),
        "data_row_start": int(getattr(region, "data_row_start", 0) or 0),
        "data_row_end": int(getattr(region, "data_row_end", 0) or 0),
    }


def _parse_file_full(path: str):
    """解析单文件激活页最优区域，额外返回区域坐标（供原地回填主表定位）。

    在【独立子进程】执行（同 _parse_file_to_df 的防护说明）；失败返回空结构
    （与 parse_excel_file 失败时的旧行为一致），后续列校验会给出明确报错。
    """
    from backend.utils.subprocess_runner import run_in_subprocess, default_max_memory_mb, default_timeout

    empty = {"sheet": "", "columns": [], "df": pd.DataFrame(),
             "head_data": {}, "data_row_start": 0, "data_row_end": 0}
    r = run_in_subprocess(
        "backend.api.tools:_parse_file_full_impl", (str(path),),
        timeout=default_timeout("parse"), max_memory_mb=default_max_memory_mb(),
    )
    if r.success:
        return r.result
    reason = "超时" if r.timed_out else ("内存超限" if r.killed_by_memory else r.error)
    logger.error(f"[parse] 子进程解析失败（{reason}）: {path}")
    return empty


@router.post("/integrate/analyze")
async def integrate_analyze(
    files: List[UploadFile] = File(...),
    tenant_id: Optional[str] = Form(None),
    current_user=Depends(get_current_user),
):
    """多表整合对比第一步：上传【≥2 个】文件（1 主表 + 至少 1 对照表）→ 解析激活页
    → 返回各文件列（带来源）+ 列头指纹 + 猜键/猜姓名/猜身份证列。不在此处调用 AI。
    主表由用户在后续步骤从上传文件里选定（不在此处固定）。
    """
    if not files or len(files) < 2:
        raise HTTPException(status_code=400, detail="请至少上传 2 个文件（1 主表 + 至少 1 对照表）")

    from ..utils.merge_engine import compute_header_fingerprint, guess_key_column
    from ..utils.integrate_engine import guess_name_column, guess_id_column

    session_id = uuid.uuid4().hex
    sdir = _integrate_session_dir(session_id)
    sdir.mkdir(parents=True, exist_ok=True)

    files_meta = []
    try:
        for uf in files:
            name = uf.filename or "unnamed.xlsx"
            if Path(name).suffix.lower() not in EXCEL_EXTS:
                raise HTTPException(status_code=400, detail=f"{name}: 不支持的扩展名")
            dest = sdir / name
            dest.write_bytes(await uf.read())
            # 上传规范化：误设日期格式的数字单元格重置为常规（避免被当日期读错）
            # Aspose 操作全部走独立子进程（防 GIL 冻结主进程），失败不阻断
            try:
                from backend.utils.subprocess_runner import (
                    run_in_subprocess, default_max_memory_mb, default_timeout,
                )
                await asyncio.to_thread(
                    run_in_subprocess,
                    "backend.utils.source_normalizer:normalize_misformatted_dates",
                    (str(dest),),
                    timeout=default_timeout("parse"),
                    max_memory_mb=default_max_memory_mb(),
                )
            except Exception as _ne:
                logger.warning(f"[integrate] 日期格式规范化失败（继续）: {_ne}")
            info = await asyncio.to_thread(_parse_file_to_df, str(dest))
            cols = info["columns"]
            df = info.get("df")
            files_meta.append({
                "name": name,
                "sheet": info["sheet"],
                "columns": cols,
                "fingerprint": compute_header_fingerprint(cols),
                "suggested_key": guess_key_column(cols, df) or (cols[0] if cols else ""),
                "suggested_name_col": guess_name_column(cols) or "",
                "suggested_id_col": guess_id_column(cols) or "",
            })

        (sdir / "_meta.json").write_text(
            json.dumps({"files": files_meta}, ensure_ascii=False), encoding="utf-8"
        )
        matched_schemes = _match_integrate_schemes(current_user, files_meta)
        return {"session_id": session_id, "files": files_meta, "matched_schemes": matched_schemes}
    except HTTPException:
        shutil.rmtree(sdir, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(sdir, ignore_errors=True)
        logger.exception("integrate/analyze 异常")
        raise HTTPException(status_code=500, detail=str(e))


class IntegrateMatchRequest(BaseModel):
    session_id: str
    main_file: str                       # 主表 A
    source_cols: List[dict]              # [{"file","col"}] 对照表候选列（用户勾选范围）
    use_ai: bool = False
    ai_provider: Optional[str] = "deepseek"


@router.post("/integrate/match")
async def integrate_match(req: IntegrateMatchRequest, current_user=Depends(get_current_user)):
    """把对照表候选列匹配到主表 A 的列（覆盖/对比列对建议）：先精确同名（免 AI），
    use_ai 时再用 AI 补差异命名。返回 pairs=[{a_col,source_file,source_col,auto,confidence}]。
    """
    from ..utils.merge_engine import _norm_header, _ai_map_to_canonical

    sdir = _integrate_session_dir(req.session_id)
    meta_path = sdir / "_meta.json"
    if not meta_path.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    cols_by_file = {f["name"]: f.get("columns", []) for f in meta.get("files", [])}
    a_cols = cols_by_file.get(req.main_file, [])
    if not a_cols:
        raise HTTPException(status_code=400, detail="主表列为空或主表选择有误")

    a_norm = {_norm_header(c): c for c in a_cols}
    pairs = []
    matched_src = set()
    # 1) 精确同名（免 AI）
    for s in req.source_cols or []:
        f, c = s.get("file"), s.get("col")
        if not f or not c or f == req.main_file:
            continue
        hit = a_norm.get(_norm_header(c))
        if hit:
            pairs.append({"a_col": hit, "source_file": f, "source_col": c, "auto": True, "confidence": 1.0})
            matched_src.add((f, c))

    # 2) AI 补差异命名（对未精确命中的候选列）
    ai_suggestions = []
    if req.use_ai:
        remaining = [(s.get("file"), s.get("col")) for s in (req.source_cols or [])
                     if (s.get("file"), s.get("col")) not in matched_src
                     and s.get("file") != req.main_file and s.get("col")]
        if remaining:
            try:
                ai_suggestions = _ai_map_to_canonical(a_cols, remaining, req.ai_provider or "deepseek")
                for it in ai_suggestions:
                    pairs.append({
                        "a_col": it["suggest_group"], "source_file": it["file"],
                        "source_col": it["col"], "auto": False,
                        "confidence": it.get("confidence", 0.0),
                    })
            except Exception as e:
                logger.warning(f"[integrate] AI 匹配失败（忽略）: {e}")

    return {"pairs": pairs, "ai_suggestions": ai_suggestions}


class IntegrateExecuteRequest(BaseModel):
    session_id: str
    main_file: str                       # 主表 A（上传文件二选一/多选一）
    key_map: dict                        # {file: 关联键列名}（含主表与各对照表）
    overwrite_pairs: List[dict]          # [{"a_col","source_file","source_col"}] 有序=优先级
    compare_pairs: List[dict] = []       # [{"a_col","source_file","source_col"}]
    output_mode: int = 1                 # 1 只更新主表；2 主表 + 差异 sheet
    name_col: Optional[str] = None       # 主表姓名列（差异 sheet）
    id_col: Optional[str] = None         # 主表身份证列（差异 sheet）
    diff_order: str = "id_name"          # id_name | name_id
    normalize_keys: bool = True
    date_key_mode: str = "yearmonthday"  # 日期关联键归一 off|yearmonthday|yearmonth|month|day（默认按年月日）


def _validate_integrate_columns(parsed, key_map, overwrite_pairs, compare_pairs, main_file):
    """执行前预检：方案引用的列在对应上传文件里是否真实存在。

    缺列历史上是静默失败的（键列缺失→空索引→0 命中；源列缺失→取值返回 None→跳过），
    用户只会看到"0 行"或结果不对。这里把每个引用列逐一校验，报错精确到 文件+sheet+列。
    Returns: 错误消息列表（空则通过）。
    """
    from ..utils.integrate_engine import _expr_columns

    errs = []

    def _sheet(f):
        return (parsed.get(f) or {}).get("sheet") or ""

    def _cols(f):
        _df = (parsed.get(f) or {}).get("df")
        return list(_df.columns) if _df is not None else []

    # 1) 关联键列：主表与各对照表都必须存在
    for f, kcol in (key_map or {}).items():
        if not kcol or f not in parsed:
            continue
        if kcol not in _cols(f):
            errs.append(f"文件【{f}】的 sheet【{_sheet(f)}】的关联键列 '{kcol}' 不存在")

    # 2) 覆盖/对比对的 A 列（主表）与源列（对照表，支持公式里的列名）
    main_cols, main_sheet = _cols(main_file), _sheet(main_file)
    for pair in list(overwrite_pairs or []) + list(compare_pairs or []):
        ac, f = pair.get("a_col"), pair.get("source_file")
        expr = pair.get("source_expr") or pair.get("source_col")
        if ac and ac not in main_cols:
            errs.append(f"文件【{main_file}】的 sheet【{main_sheet}】的列 '{ac}' 不存在（覆盖/对比目标列）")
        if not (f and f != main_file and f in parsed and expr):
            continue
        fcols = _cols(f)
        expr = str(expr)
        # 已知列：本表裸列 + 所有对照表的 `文件名.列名` 跨表引用（与 eval_source_expr_cross 语法一致）
        refs = _expr_columns(expr, fcols)
        if expr in fcols:
            refs.append(expr)
        cross_refs = []
        for _fn in parsed.keys():
            if _fn == main_file:
                continue
            _fcols = _cols(_fn)
            cross_refs += [f"{_fn}.{_c}" for _c in _fcols]
        refs += [c for c in cross_refs if c in expr]
        # 把已知列抠掉后，剩下的非数字/运算符 token 即"引用但不存在"的列名。
        # 必须按长度降序抠：`B.xlsx.基本工资` 里的 `基本工资` 会被裸列先抠掉导致剩 `B.xlsx.` 误报。
        subst = expr
        for c in sorted(refs, key=len, reverse=True):
            subst = subst.replace(c, " ")
        unresolved = [t for t in re.split(r"[+\*/()\s]+", subst)
                      if t and not re.fullmatch(r"[0-9eE.+\-*/()]*", t)]
        if unresolved:
            errs.append(f"文件【{f}】的 sheet【{_sheet(f)}】的列 '{'、'.join(unresolved)}' 不存在（覆盖/对比源列）")
    return errs


@router.post("/integrate/execute")
async def integrate_execute(req: IntegrateExecuteRequest, current_user=Depends(get_current_user)):
    """按覆盖/对比配置回填主表并输出：原地覆盖主表激活页覆盖列（只写值、保全其余 sheet/公式），
    输出方式2 追加差异 sheet。返回更新后的主表 xlsx。
    """
    from ..utils.integrate_engine import build_source_indexes, compute_diffs
    from ..utils.integrate_writer import apply_integration, append_diff_sheet

    sdir = _integrate_session_dir(req.session_id)
    if not sdir.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")

    main_path = sdir / req.main_file
    if not main_path.exists():
        raise HTTPException(status_code=400, detail=f"主表文件不存在: {req.main_file}")

    try:
        # 解析：主表要区域坐标，对照表只要 df
        # 解析走子进程 + to_thread：不冻结事件循环（多客户并发时其他请求照常响应）
        main_info = await asyncio.to_thread(_parse_file_full, str(main_path))
        parsed = {req.main_file: {"df": main_info["df"], "sheet": main_info["sheet"]}}
        for fp_path in sorted(sdir.iterdir()):
            if fp_path.suffix.lower() not in EXCEL_EXTS or fp_path.name == req.main_file:
                continue
            _info = await asyncio.to_thread(_parse_file_to_df, str(fp_path))
            parsed[fp_path.name] = {"df": _info["df"], "sheet": _info["sheet"]}

        # 执行前预检：方案引用的列必须真实存在，报错精确到 文件+sheet+列（缺列历史上是静默的）
        errs = _validate_integrate_columns(
            parsed, req.key_map, req.overwrite_pairs, req.compare_pairs, req.main_file)
        if errs:
            raise HTTPException(status_code=400, detail="；".join(errs))

        n_sources = len([f for f in parsed if f != req.main_file])
        if n_sources < 1:
            raise HTTPException(status_code=400, detail="至少需要 1 张对照表")

        source_indexes = build_source_indexes(
            parsed, req.key_map, req.main_file, normalize_keys=req.normalize_keys,
            date_key_mode=req.date_key_mode,
        )

        # 先执行覆盖回填并生成结果文件（含公式重算）。差异对比放到生成之后，
        # 基于【最终生成文件】的实时值比对，避免用主表覆盖前的缓存旧值。
        out_path = sdir / f"整合结果_{req.main_file}"
        stat = await asyncio.to_thread(apply_integration,
            main_path=str(main_path), out_path=str(out_path),
            sheet_name=main_info["sheet"], head_data=main_info["head_data"],
            a_key_col=req.key_map.get(req.main_file),
            data_row_start=main_info["data_row_start"], data_row_end=main_info["data_row_end"],
            overwrite_pairs=req.overwrite_pairs, source_indexes=source_indexes,
            normalize_keys=req.normalize_keys,
            diff_rows=None, diff_order=req.diff_order,
            date_key_mode=req.date_key_mode,
        )

        # 对比差异（仅输出方式2且配了对比列时）：重新解析生成文件取重算后的实时值再比对。
        # 覆盖列联动的公式（如按被覆盖列计算的合计）此时已重算，比对结果才是最终文件的真实差异。
        if req.output_mode == 2 and req.compare_pairs:
            final_info = await asyncio.to_thread(_parse_file_full, str(out_path))
            diff_rows = compute_diffs(
                final_info["df"], source_indexes,
                a_key_col=req.key_map.get(req.main_file),
                compare_pairs=req.compare_pairs,
                a_name_col=req.name_col, a_id_col=req.id_col,
                normalize_keys=req.normalize_keys,
                label_source=(n_sources > 1),
                date_key_mode=req.date_key_mode,
            )
            stat["diff_rows"] = await asyncio.to_thread(
                append_diff_sheet, str(out_path), diff_rows, req.diff_order)

        data = out_path.read_bytes()
        buf = io.BytesIO(data)
        buf.seek(0)
        # 不删会话：支持"先下载看看，再改配置/重新生成/保存方案"等后续操作
        try:
            out_path.unlink(missing_ok=True)   # 仅清理本次生成的临时结果文件，保留上传的源文件与 _meta
        except Exception:
            pass
        from urllib.parse import quote
        fname = f"整合结果_{req.main_file}"
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(fname)}",
            "X-Integrate-Matched": str(stat.get("matched_rows", 0)),
            "X-Integrate-Cells": str(stat.get("overwritten_cells", 0)),
            "X-Integrate-Diffs": str(stat.get("diff_rows", 0)),
        }
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("integrate/execute 异常")
        raise HTTPException(status_code=500, detail=str(e))


# ==================== 多表整合对比：命名方案（按列头指纹联动）====================

def _is_admin(user) -> bool:
    return bool(getattr(user, "role", None) and getattr(user.role, "name", None) == "admin")


def _user_org_tree_ids(db, current_user) -> List[int]:
    """当前用户所属组织 + 所有子组织 ID。无组织返回空。"""
    from ..auth.dependencies import _get_org_and_children_ids
    oid = getattr(current_user, "org_id", None)
    if not oid:
        return []
    return _get_org_and_children_ids(db, oid)


def _visible_integrate_schemes(db, current_user):
    """当前用户可见的整合方案行：
    - admin：全部；
    - 其他：本组织+子组织(org_id) ∪ 本人创建(created_by) ∪ 旧私有方案(tenant_id==user:{id})。
    """
    from ..database.models import IntegrateTemplate
    from sqlalchemy import or_
    q = db.query(IntegrateTemplate)
    if _is_admin(current_user):
        return q.all()
    uid = getattr(current_user, "id", None)
    org_ids = _user_org_tree_ids(db, current_user)
    conds = []
    if org_ids:
        conds.append(IntegrateTemplate.org_id.in_(org_ids))
    if uid is not None:
        conds.append(IntegrateTemplate.created_by == uid)
        conds.append(IntegrateTemplate.tenant_id == f"user:{uid}")
    if not conds:
        return []
    return q.filter(or_(*conds)).all()


def _can_edit_scheme(row, current_user) -> bool:
    """仅创建人 + 管理员可改/删。"""
    if _is_admin(current_user):
        return True
    uid = getattr(current_user, "id", None)
    if uid is not None and row.created_by == uid:
        return True
    # 兼容旧私有方案（created_by 为空、tenant_id==user:{id}）
    if row.created_by is None and row.tenant_id == f"user:{uid}":
        return True
    return False


_PHANTOM_COL = re.compile(r"^Column_[A-Z]+$")


def _real_header_set(cols):
    """归一化的真实列头集合：过滤 excel_parser 生成的空表头幻影列(Column_X)。

    幻影列是"有格式但表头为空"的列占位（excel_parser.py 3069 行），同一张表在不同月份
    因末尾多/少一个空白样式列会导致指纹不一致，方案应用被误判为"表头结构不一致"。
    匹配角色时把幻影列忽略，仅按真实列头集合比对。
    """
    from ..utils.merge_engine import _norm_header
    return {_norm_header(c) for c in (cols or []) if not _PHANTOM_COL.match(str(c))}


def _file_matches_role(f, expected_fp, expected_cols) -> bool:
    """该文件是否可作为方案角色的候选：精确指纹，或忽略幻影列后真实列头集合一致。"""
    if f.get("fingerprint") == expected_fp:
        return True
    exp_set = _real_header_set(expected_cols)
    if exp_set:
        return _real_header_set(f.get("columns", [])) == exp_set
    return False


def _pick_file_for_role(used: set, expected_fp, expected_cols, files_meta):
    """为方案角色挑上传文件：优先精确指纹；否则按忽略幻影列后的列集合一致兜底。

    Returns: 命中文件 dict（未命中 None），不修改 used。
    """
    for f in files_meta:
        if f["name"] not in used and _file_matches_role(f, expected_fp, expected_cols):
            return f
    return None


def _match_integrate_schemes(current_user, files_meta: List[dict]) -> List[dict]:
    """按列头指纹把【当前用户可见】方案匹配到本次上传的文件，解析出角色(主表/各对照表)。

    Returns: [{id, name, main_file, fp_to_file{fp:file}, ambiguous, config}]
    - 仅返回能完整匹配的方案（每个角色至少一个候选文件）。
    - ambiguous: [{"label","fp","candidates":[文件名...],"saved_file"}] —— 存在同结构表竞争
      同一角色（候选集合与其它角色相交）时的歧义角色，由前端弹匹配框让操作人员手动确认
      对应关系；无歧义为空列表。
    """
    from ..database.connection import SessionLocal

    out: List[dict] = []
    db = SessionLocal()
    try:
        rows = _visible_integrate_schemes(db, current_user)
        for row in rows:
            cfg = row.config or {}
            main_fp = cfg.get("main_fp")
            source_fps = cfg.get("source_fps", [])
            if not main_fp:
                continue
            cols_by_fp = cfg.get("cols_by_fp", {})
            roles = [("主表", main_fp)] + [(f"对照表{i + 1}", fp) for i, fp in enumerate(source_fps)]

            # 1) 每个角色的全部候选文件（不排除占用），先保证方案完整
            role_cands = []
            incomplete = False
            for label, fp in roles:
                cands = [f["name"] for f in files_meta
                         if _file_matches_role(f, fp, cols_by_fp.get(fp) or [])]
                if not cands:
                    incomplete = True
                    break
                role_cands.append((label, fp, cands))
            if incomplete:
                continue

            # 2) 默认分配：顺序取首个未占用（保持原有行为）。
            #    role_files 直接按角色顺序收集（同指纹角色 fp_to_file 会互相覆盖，不能反查）
            used = set()
            fp_to_file = {}
            role_files = []
            ok = True
            for label, fp, cands in role_cands:
                hit = next((c for c in cands if c not in used), None)
                if not hit:
                    ok = False
                    break
                used.add(hit)
                fp_to_file[fp] = hit
                role_files.append(hit)
            if not ok:
                continue

            # 3) 歧义检测：同结构表竞争（候选集合相交）或单角色多候选（>1 个同结构文件可选）→
            #    无法自动确定对应关系，交由前端弹匹配框人工确认。
            #    角色用【角色索引】标识（同结构角色指纹相同，fp 不能作唯一键）。
            files_by_fp = cfg.get("files_by_fp") or {}
            roles_cfg = cfg.get("roles") or []
            ambiguous = []
            for i, (label, fp, cands) in enumerate(role_cands):
                set_i = set(cands)
                involved = len(cands) > 1 or any(
                    set_i & set(rc[2]) for j, rc in enumerate(role_cands) if j != i)
                if involved:
                    saved = (roles_cfg[i].get("file", "") if roles_cfg and i < len(roles_cfg)
                             else files_by_fp.get(fp, ""))
                    ambiguous.append({
                        "label": label,
                        "fp": fp,
                        "role_index": i,
                        "candidates": sorted(cands),
                        "saved_file": saved,
                    })

            out.append({"id": row.id, "name": row.name,
                        "main_file": fp_to_file.get(main_fp, ""),
                        "fp_to_file": fp_to_file,
                        "role_files": role_files,
                        "ambiguous": ambiguous,
                        "config": cfg})
    finally:
        db.close()
    return out


class IntegrateSchemeSaveRequest(BaseModel):
    session_id: str
    name: str
    main_file: str
    key_map: dict                        # {file: 关联键列名}
    overwrite_pairs: List[dict]          # [{"a_col","source_file","source_col"}]
    compare_pairs: List[dict] = []
    name_col: Optional[str] = None
    id_col: Optional[str] = None
    diff_order: str = "id_name"
    output_mode: int = 1
    normalize_keys: bool = True
    date_key_mode: str = "yearmonthday"
    scheme_id: Optional[int] = None      # 传则为"修改已有方案"，不传为"新建"


@router.post("/integrate/scheme/save")
async def integrate_scheme_save(req: IntegrateSchemeSaveRequest, current_user=Depends(get_current_user)):
    """保存整合对比方案。新建需 create 权限、方案名在本组织内唯一；修改需 edit 权限且仅创建人/管理员。
    按列头指纹存角色，跨月复用；同时存各角色列名(cols_by_fp)供应用校验时列级 diff。"""
    from ..auth.dependencies import has_permission
    name = (req.name or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="方案名称不能为空")

    sdir = _integrate_session_dir(req.session_id)
    meta_path = sdir / "_meta.json"
    if not meta_path.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    fp_by_file = {f["name"]: f["fingerprint"] for f in meta.get("files", [])}
    cols_by_file = {f["name"]: f.get("columns", []) for f in meta.get("files", [])}
    if req.main_file not in fp_by_file:
        raise HTTPException(status_code=400, detail="主表选择有误")

    main_fp = fp_by_file[req.main_file]
    # 参与映射的对照文件（覆盖对+对比对里出现的 source_file）
    src_files = []
    for p in (req.overwrite_pairs or []) + (req.compare_pairs or []):
        f = p.get("source_file")
        if f and f != req.main_file and f in fp_by_file and f not in src_files:
            src_files.append(f)
    # 配了关联键的非主表上传文件也纳入方案角色：否则"上传 N 个文件、另存为新方案"仍只有
    # 参与映射的 M 个角色，套用时上传 N 个文件会报"方案需要 M 张表，实际上传 N 张"。
    # （key_map 覆盖所有上传文件，含主表；主表已单独处理，这里只补对照表）
    for f, k in (req.key_map or {}).items():
        if f and f != req.main_file and f in fp_by_file and k and f not in src_files:
            src_files.append(f)
    source_fps = [fp_by_file[f] for f in src_files]

    # 角色有序数组（0=主表, 1..=对照表）：同结构对照表指纹相同，fp 无法作唯一键，
    # 必须用角色索引区分"第几个对照表"，供套用时按角色顺序取文件。
    roles = [{"fp": main_fp, "file": req.main_file}]
    for f in src_files:
        roles.append({"fp": fp_by_file[f], "file": f})
    role_of_file = {r["file"]: i for i, r in enumerate(roles)}

    def _to_fp_pairs(pairs):
        out = []
        for p in pairs or []:
            f = p.get("source_file")
            if f in fp_by_file:
                _expr = p.get("source_expr") or p.get("source_col")
                out.append({"a_col": p.get("a_col"),
                            "source_fp": fp_by_file[f],
                            "source_role": role_of_file.get(f),
                            "source_expr": _expr, "source_col": _expr})
        return out

    # 各角色列名（供应用校验时列级 diff）：主表 + 各对照表
    cols_by_fp = {main_fp: cols_by_file.get(req.main_file, [])}
    for f in src_files:
        cols_by_fp[fp_by_file[f]] = cols_by_file.get(f, [])

    # 各角色保存时的文件名（供套用出现同结构表歧义时提示"保存时：本月.xlsx"参考）
    files_by_fp = {main_fp: req.main_file}
    for f in src_files:
        files_by_fp[fp_by_file[f]] = f

    config = {
        "main_fp": main_fp,
        "source_fps": source_fps,
        "roles": roles,
        "cols_by_fp": cols_by_fp,
        "files_by_fp": files_by_fp,
        "key_map_by_fp": {fp_by_file[f]: k for f, k in (req.key_map or {}).items() if f in fp_by_file},
        "key_map_by_role": {role_of_file[f]: k for f, k in (req.key_map or {}).items() if f in role_of_file},
        "overwrite_pairs": _to_fp_pairs(req.overwrite_pairs),
        "compare_pairs": _to_fp_pairs(req.compare_pairs),
        "name_col": req.name_col, "id_col": req.id_col,
        "diff_order": req.diff_order, "output_mode": req.output_mode,
        "normalize_keys": req.normalize_keys,
        "date_key_mode": req.date_key_mode,
    }

    from ..database.connection import SessionLocal
    from ..database.models import IntegrateTemplate
    uid = getattr(current_user, "id", None)
    org_id = getattr(current_user, "org_id", None)
    # 有组织归属到 org:{org_id}（方案名按组织唯一），否则退回旧私有 user:{uid}
    tenant_id = f"org:{org_id}" if org_id else f"user:{uid}"
    db = SessionLocal()
    try:
        if req.scheme_id is not None:
            # 修改已有方案：需 edit 权限 + 仅创建人/管理员（同组织其他人只能"另存为"新建）
            row = db.query(IntegrateTemplate).filter_by(id=req.scheme_id).first()
            if not row:
                raise HTTPException(status_code=404, detail="方案不存在")
            if not has_permission(current_user, "tools.data_integrate.edit"):
                raise HTTPException(status_code=403, detail="缺少权限: 修改方案")
            if not _can_edit_scheme(row, current_user):
                raise HTTPException(status_code=403, detail="只有创建人或管理员可以保存修改到该方案，其他人请使用「另存为」")
            row.name = name
            row.config = config
            row.updated_by = uid
        else:
            # 新建/另存为：需 create 权限 + 组织内方案名唯一（重名提醒）
            if not has_permission(current_user, "tools.data_integrate.create"):
                raise HTTPException(status_code=403, detail="缺少权限: 新增方案")
            dup = db.query(IntegrateTemplate).filter_by(tenant_id=tenant_id, name=name).first()
            if dup:
                raise HTTPException(status_code=400, detail=f"同组织内已存在同名方案「{name}」，请换个名称")
            row = IntegrateTemplate(tenant_id=tenant_id, name=name, config=config,
                                    org_id=org_id, created_by=uid, updated_by=uid)
            db.add(row)
        db.commit()
        db.refresh(row)
        return {"ok": True, "id": row.id, "name": name}
    except HTTPException:
        db.rollback()
        raise
    except Exception as e:
        db.rollback()
        logger.exception("integrate/scheme/save 异常")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        db.close()


@router.get("/integrate/schemes")
async def integrate_schemes_list(current_user=Depends(get_current_user)):
    """列出当前用户可见（同组织+子组织，或本人创建）的整合对比方案。"""
    from ..database.connection import SessionLocal
    db = SessionLocal()
    try:
        rows = _visible_integrate_schemes(db, current_user)
        rows = sorted(rows, key=lambda r: r.updated_at or r.created_at, reverse=True)
        out = []
        for r in rows:
            cfg = r.config or {}
            n_files = 1 + len(cfg.get("source_fps", []))   # 主表 + 对照表数
            creator_name = ""
            try:
                creator_name = r.creator.display_name if r.creator else ""
            except Exception:
                creator_name = ""
            updater_name = ""
            try:
                updater_name = r.updater.display_name if r.updater else ""
            except Exception:
                updater_name = ""
            out.append({
                "id": r.id, "name": r.name, "config": cfg,
                "file_count": n_files,
                "creator_name": creator_name,
                "updated_by_name": updater_name,
                "created_at": r.created_at.isoformat() if r.created_at else None,
                "updated_at": r.updated_at.isoformat() if r.updated_at else None,
                # can_edit: 能否保存修改覆盖原方案/删除（仅创建人+管理员）
                "can_edit": _can_edit_scheme(r, current_user),
                # can_modify: 能否进入修改配置（同组织可见即可；非创建人只能另存为）
                "can_modify": True,
            })
        return out
    finally:
        db.close()


@router.delete("/integrate/scheme/{scheme_id}")
async def integrate_scheme_delete(scheme_id: int, current_user=Depends(get_current_user)):
    """删除整合对比方案：需 delete 权限，且仅创建人/管理员可删。"""
    from ..auth.dependencies import has_permission
    from ..database.connection import SessionLocal
    from ..database.models import IntegrateTemplate
    db = SessionLocal()
    try:
        row = db.query(IntegrateTemplate).filter_by(id=scheme_id).first()
        if not row:
            raise HTTPException(status_code=404, detail="方案不存在")
        if not has_permission(current_user, "tools.data_integrate.delete"):
            raise HTTPException(status_code=403, detail="缺少权限: 删除方案")
        if not _can_edit_scheme(row, current_user):
            raise HTTPException(status_code=403, detail="只有创建人或管理员可以删除该方案")
        db.delete(row)
        db.commit()
        return {"ok": True}
    except HTTPException:
        db.rollback()
        raise
    finally:
        db.close()


class IntegrateApplyValidateRequest(BaseModel):
    session_id: str
    scheme_id: int


@router.post("/integrate/scheme/apply-validate")
async def integrate_scheme_apply_validate(req: IntegrateApplyValidateRequest, current_user=Depends(get_current_user)):
    """应用方案前校验：上传文件数 + 表头结构是否与原方案一致（硬阻断，列出不一致理由）。
    通过后前端用 analyze 返回的 matched_schemes 里对应方案（含 fp_to_file）套用并执行。"""
    from ..auth.dependencies import has_permission
    from ..utils.merge_engine import _norm_header
    from ..database.connection import SessionLocal

    if not has_permission(current_user, "tools.data_integrate.apply"):
        raise HTTPException(status_code=403, detail="缺少权限: 应用方案")

    sdir = _integrate_session_dir(req.session_id)
    meta_path = sdir / "_meta.json"
    if not meta_path.exists():
        raise HTTPException(status_code=400, detail="会话已过期，请重新上传分析")
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    uploaded = meta.get("files", [])   # [{name, columns, fingerprint}]

    db = SessionLocal()
    try:
        # 只允许校验用户可见的方案
        visible_ids = {r.id: r for r in _visible_integrate_schemes(db, current_user)}
        row = visible_ids.get(req.scheme_id)
        if not row:
            raise HTTPException(status_code=404, detail="方案不存在或无权访问")
        cfg = row.config or {}
    finally:
        db.close()

    main_fp = cfg.get("main_fp")
    source_fps = cfg.get("source_fps", [])
    cols_by_fp = cfg.get("cols_by_fp", {})
    files_by_fp = cfg.get("files_by_fp") or {}
    roles_cfg = cfg.get("roles") or []
    expected = [("主表", main_fp)] + [(f"对照表{i + 1}", fp) for i, fp in enumerate(source_fps)]

    def _saved_file(i: int, fp: str) -> str:
        """该角色的保存时文件名（roles 按索引取，兼容旧方案回退 files_by_fp）"""
        if roles_cfg and i < len(roles_cfg):
            return str(roles_cfg[i].get("file") or "")
        return str(files_by_fp.get(fp) or "")

    def _label_txt(i: int, label: str, fp: str) -> str:
        sf = _saved_file(i, fp)
        return f"{label}（方案中表名：{sf}）" if sf else label

    reasons: List[str] = []
    # 1) 文件数校验（列出方案需要的具体表名，方便用户对照）
    if len(uploaded) != len(expected):
        names = "；".join(_label_txt(i, label, fp) for i, (label, fp) in enumerate(expected))
        reasons.append(f"方案需要 {len(expected)} 张表，实际上传 {len(uploaded)} 张。方案表清单：{names}")

    # 2) 表头结构校验：逐个期望角色按"精确指纹→忽略幻影列后列集合一致"找上传文件（消耗式），
    #    找不到则对最接近文件做列级 diff（同样忽略幻影列，避免空表头样式列干扰）。
    used = set()
    for i, (label, fp) in enumerate(expected):
        exp_cols = cols_by_fp.get(fp) or []
        hit = _pick_file_for_role(used, fp, exp_cols, uploaded)
        if hit:
            used.add(hit["name"])
            continue
        # 在未占用文件里按列名重合度找最接近的
        best, best_score = None, -1.0
        exp_norm = _real_header_set(exp_cols)
        for f in uploaded:
            if f["name"] in used:
                continue
            up_norm = _real_header_set(f.get("columns", []))
            inter = len(exp_norm & up_norm)
            union = len(exp_norm | up_norm) or 1
            score = inter / union
            if score > best_score:
                best_score, best = score, f
        if best is not None and exp_norm:
            up_norm = _real_header_set(best.get("columns", []))
            missing = [c for c in exp_cols if not _PHANTOM_COL.match(str(c)) and _norm_header(c) not in up_norm]
            extra = [c for c in best.get("columns", []) if not _PHANTOM_COL.match(str(c)) and _norm_header(c) not in exp_norm]
            detail = []
            if extra:
                detail.append("多列 " + "、".join(map(str, extra)))
            if missing:
                detail.append("缺列 " + "、".join(map(str, missing)))
            sheet = best.get("sheet") or "?"
            tail = (f"；最接近的是文件【{best['name']}】的 sheet【{sheet}】：" + "；".join(detail)) if detail else ""
            reasons.append(f"缺少表头结构为【{_label_txt(i, label, fp)}】的表" + tail)
        else:
            reasons.append(f"缺少表头结构为【{_label_txt(i, label, fp)}】的表")

    return {"ok": len(reasons) == 0, "reasons": reasons, "scheme_id": req.scheme_id, "scheme_name": row.name}


# ==================== SOP 维护 ====================

SOP_ROOT = PROJECT_ROOT / "tenants" / "__tools_sop__"
_sop_executor = ThreadPoolExecutor(max_workers=2)   # AI 后台分析线程池


class SopAiVerdict(BaseModel):
    """AI 结构化判定结果"""
    passed: bool
    score: int = 0
    summary: str = ""
    issues: List[str] = []
    suggestions: List[str] = []
    details: dict = {}


class SopCreateRequest(BaseModel):
    customer_name: str
    description: str = ""


class SopReviewRequest(BaseModel):
    round_id: int
    verdict: str          # completed / rejected
    comment: str = ""


def _sop_round_dir(entry_id: int, round_no: int) -> Path:
    return SOP_ROOT / "entries" / str(entry_id) / f"round_{round_no}"


def _sop_latest_round(entry: SopEntry, db) -> Optional[SopRound]:
    if entry.latest_round_id:
        return db.query(SopRound).filter_by(id=entry.latest_round_id).first()
    return (
        db.query(SopRound).filter(SopRound.sop_id == entry.id)
        .order_by(SopRound.round_no.desc()).first()
    )


def _sop_round_to_out(r: SopRound) -> dict:
    return {
        "id": r.id,
        "round_no": r.round_no,
        "status": r.status,
        "source_file_name": r.source_file_name,
        "source_file_names": r.source_file_names or ([r.source_file_name] if r.source_file_name else []),
        "result_file_name": r.result_file_name,
        "rule_file_name": r.rule_file_name,
        "ai_analysis": r.ai_analysis,
        "ai_comment": r.ai_comment,
        "ai_provider": r.ai_provider,
        "error_message": r.error_message,
        "review_status": r.review_status,
        "review_comment": r.review_comment,
        "reviewer_name": r.reviewer.display_name if r.reviewer else "",
        "reviewed_at": r.reviewed_at.isoformat() if r.reviewed_at else None,
        "started_at": r.started_at.isoformat() if r.started_at else None,
        "finished_at": r.finished_at.isoformat() if r.finished_at else None,
    }


def _sop_entry_to_out(entry: SopEntry, db) -> dict:
    latest = _sop_latest_round(entry, db)
    return {
        "id": entry.id,
        "customer_name": entry.customer_name,
        "description": entry.description,
        "status": entry.status,
        "created_by_name": entry.creator.display_name if entry.creator else "",
        "created_at": entry.created_at.isoformat() if entry.created_at else None,
        "updated_at": entry.updated_at.isoformat() if entry.updated_at else None,
        "round_no": latest.round_no if latest else 0,
        "ai_comment": (latest.ai_comment or "") if latest else "",
        "latest_round_id": entry.latest_round_id,
    }


# ---------- SOP 条目 ----------

@router.get("/sop/entries")
async def sop_entries_list(
    keyword: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    current_user=Depends(require_permission("tools.sop")),
):
    """SOP 条目列表：客户名称模糊搜索 + 状态过滤。"""
    db = SessionLocal()
    try:
        q = db.query(SopEntry)
        if keyword:
            q = q.filter(SopEntry.customer_name.like(f"%{keyword}%"))
        if status:
            q = q.filter(SopEntry.status == status)
        entries = q.order_by(SopEntry.updated_at.desc()).all()
        return {"items": [_sop_entry_to_out(e, db) for e in entries]}
    finally:
        db.close()


@router.post("/sop/entries")
async def sop_entries_create(req: SopCreateRequest, current_user=Depends(require_permission("tools.sop.create"))):
    """新建 SOP 条目（客户名称 + 描述），状态 draft。"""
    name = (req.customer_name or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="客户名称不能为空")
    db = SessionLocal()
    try:
        entry = SopEntry(
            customer_name=name, description=req.description or "",
            status="draft", created_by=current_user.id,
        )
        db.add(entry)
        db.commit()
        db.refresh(entry)
        return _sop_entry_to_out(entry, db)
    finally:
        db.close()


@router.get("/sop/entries/{entry_id}")
async def sop_entries_detail(entry_id: int, current_user=Depends(require_permission("tools.sop"))):
    """SOP 详情：含全部轮次历史。"""
    db = SessionLocal()
    try:
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if not entry:
            raise HTTPException(status_code=404, detail="SOP条目不存在")
        out = _sop_entry_to_out(entry, db)
        out["rounds"] = [_sop_round_to_out(r) for r in sorted(entry.rounds, key=lambda x: x.round_no)]
        return out
    finally:
        db.close()


@router.delete("/sop/entries/{entry_id}")
async def sop_entries_delete(entry_id: int, current_user=Depends(require_permission("tools.sop.manage"))):
    """删除 SOP 条目 + 轮次记录 + 物理文件。"""
    db = SessionLocal()
    try:
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if not entry:
            raise HTTPException(status_code=404, detail="SOP条目不存在")
        entry_dir = SOP_ROOT / "entries" / str(entry_id)
        db.delete(entry)   # 关系 cascade 级联删除 rounds
        db.commit()
        if entry_dir.exists():
            shutil.rmtree(entry_dir, ignore_errors=True)
        return {"ok": True}
    finally:
        db.close()


@router.post("/sop/entries/{entry_id}/upload")
async def sop_entries_upload(
    entry_id: int,
    source_files: List[UploadFile] = File(...),
    result_file: UploadFile = File(...),
    rule_file: UploadFile = File(None),
    current_user=Depends(require_permission("tools.sop.create")),
):
    """上传源(可多个)/结果/规则文件 → 建新轮次 → 后台启动 AI 分析。"""
    import asyncio

    db = SessionLocal()
    try:
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if not entry:
            raise HTTPException(status_code=404, detail="SOP条目不存在")
        if entry.status == "ai_analyzing":
            raise HTTPException(status_code=400, detail="正在AI分析中，请稍候再操作")

        max_no = db.query(SopRound).filter(SopRound.sop_id == entry_id).count()
        round_no = max_no + 1
        rdir = _sop_round_dir(entry_id, round_no)
        rdir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        def _save(uf: UploadFile, prefix: str, idx: int = 0):
            orig = Path(uf.filename or "").name or "file"
            dest = rdir / f"{prefix}_{ts}_{idx}_{orig}"
            dest.write_bytes(uf.file.read())
            return str(dest), orig

        if not source_files:
            raise HTTPException(status_code=400, detail="请至少上传一个源文件")
        src_items = [_save(f, "source", i) for i, f in enumerate(source_files)]
        src_paths = [p for p, _ in src_items]
        src_names = [n for _, n in src_items]
        res_path, res_name = _save(result_file, "result")
        rule_path = rule_name = None
        if rule_file and rule_file.filename:
            rule_path, rule_name = _save(rule_file, "rule")

        round = SopRound(
            sop_id=entry_id, round_no=round_no, status="ai_analyzing",
            source_file_path=src_paths[0], source_file_name=src_names[0],
            source_file_paths=src_paths, source_file_names=src_names,
            result_file_path=res_path, result_file_name=res_name,
            rule_file_path=rule_path, rule_file_name=rule_name,
        )
        db.add(round)
        db.flush()
        entry.status = "ai_analyzing"
        entry.latest_round_id = round.id
        db.commit()
        db.refresh(round)

        # 后台 AI 分析（纯 AI 调用，无子进程，规避 Windows SelectorEventLoop 坑）
        loop = asyncio.get_event_loop()
        loop.run_in_executor(_sop_executor, _run_sop_ai_analysis, round.id, entry_id)

        return {
            "round_id": round.id, "entry_id": entry_id,
            "round_no": round.round_no, "status": entry.status,
        }
    except HTTPException:
        db.rollback()
        raise
    except Exception as e:
        db.rollback()
        logger.exception("[sop] 上传失败")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        db.close()


@router.get("/sop/rounds/{round_id}")
async def sop_round_poll(round_id: int, current_user=Depends(require_permission("tools.sop"))):
    """轮询本轮分析状态（前端 2~3s 调用直到终态）。"""
    db = SessionLocal()
    try:
        r = db.query(SopRound).filter_by(id=round_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="轮次不存在")
        return _sop_round_to_out(r)
    finally:
        db.close()


@router.get("/sop/rounds/{round_id}/download")
async def sop_round_download(
    round_id: int,
    kind: str = Query("source"),
    idx: int = Query(0),
    current_user=Depends(require_permission("tools.sop")),
):
    """下载某轮某类文件：kind = source / result / rule，source 可多个用 idx 指定。"""
    from urllib.parse import quote
    db = SessionLocal()
    try:
        r = db.query(SopRound).filter_by(id=round_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="轮次不存在")
        if kind == "source":
            paths = r.source_file_paths or ([r.source_file_path] if r.source_file_path else [])
            names = r.source_file_names or ([r.source_file_name] if r.source_file_name else [])
            if idx >= len(paths):
                raise HTTPException(status_code=404, detail="源文件不存在")
            path = paths[idx]
            name = names[idx] if idx < len(names) else os.path.basename(paths[idx])
        else:
            path = {"result": r.result_file_path, "rule": r.rule_file_path}.get(kind)
            name = {"result": r.result_file_name, "rule": r.rule_file_name}.get(kind)
            if kind == "rule" and not path:
                raise HTTPException(status_code=404, detail="本轮未上传规则文件")
        if not path or not os.path.exists(path):
            raise HTTPException(status_code=404, detail="文件不存在")
        return FileResponse(
            path, filename=name,
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(name)}"},
        )
    finally:
        db.close()


@router.post("/sop/entries/{entry_id}/review")
async def sop_entries_review(entry_id: int, req: SopReviewRequest, current_user=Depends(require_permission("tools.sop.review"))):
    """人工审核：标记已完成(completed) 或 打回重写完善(rejected)。"""
    if req.verdict not in ("completed", "rejected"):
        raise HTTPException(status_code=400, detail="verdict 必须是 completed 或 rejected")
    db = SessionLocal()
    try:
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if not entry:
            raise HTTPException(status_code=404, detail="SOP条目不存在")
        if entry.status != "ai_passed":
            raise HTTPException(status_code=400, detail="当前状态不允许人工审核（仅待人工审核状态可操作）")
        r = db.query(SopRound).filter_by(id=req.round_id).first()
        if not r or r.sop_id != entry_id:
            raise HTTPException(status_code=404, detail="轮次不存在")
        if entry.latest_round_id != r.id:
            raise HTTPException(status_code=400, detail="只能审核最新一轮")
        r.review_status = req.verdict
        r.review_comment = req.comment or ""
        r.reviewed_by = current_user.id
        r.reviewed_at = datetime.utcnow()
        r.status = req.verdict
        entry.status = req.verdict
        db.commit()
        return _sop_round_to_out(r)
    finally:
        db.close()


# ---------- SOP 规则文件（系统管理页维护）----------

@router.get("/sop/rules")
async def sop_rules_list(current_user=Depends(require_permission("tools.sop.manage"))):
    """规则文件列表：全局优先，其次按客户。"""
    db = SessionLocal()
    try:
        rows = db.query(SopRuleFile).all()
        items = [{
            "id": r.id, "scope": r.scope, "customer_name": r.customer_name,
            "name": r.name, "description": r.description, "file_name": r.file_name,
            "is_active": r.is_active,
            "created_by_name": r.creator.display_name if r.creator else "",
            "updated_at": r.updated_at.isoformat() if r.updated_at else None,
        } for r in rows]
        items.sort(key=lambda x: (x["scope"] != "global", x["id"]))
        return {"items": items}
    finally:
        db.close()


@router.post("/sop/rules")
async def sop_rules_create(
    file: UploadFile = File(...),
    scope: str = Form("global"),
    customer_name: Optional[str] = Form(None),
    name: str = Form(""),
    description: str = Form(""),
    current_user=Depends(require_permission("tools.sop.manage")),
):
    """上传规则文件：scope=global 全局大规则 / scope=customer 按客户专属规则。

    同 scope（+customer）已有启用规则 → 先停用，保证每次只有一条生效。
    """
    scope = (scope or "global").strip()
    if scope not in ("global", "customer"):
        raise HTTPException(status_code=400, detail="scope 必须是 global 或 customer")
    cust = (customer_name or "").strip()
    if scope == "customer" and not cust:
        raise HTTPException(status_code=400, detail="按客户规则必须填写客户名称")
    if scope == "global":
        cust = None

    db = SessionLocal()
    try:
        # 同 scope（+customer）已有 active 规则 → 停用
        q = db.query(SopRuleFile).filter(SopRuleFile.scope == scope, SopRuleFile.is_active == True)
        if cust:
            q = q.filter(SopRuleFile.customer_name == cust)
        else:
            q = q.filter(SopRuleFile.customer_name.is_(None))
        for old in q.all():
            old.is_active = False

        rules_dir = SOP_ROOT / "rules"
        rules_dir.mkdir(parents=True, exist_ok=True)
        orig = Path(file.filename or "").name or "rule"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = rules_dir / f"rule_{ts}_{orig}"
        dest.write_bytes(file.file.read())

        row = SopRuleFile(
            scope=scope, customer_name=cust, name=name or "", description=description,
            file_path=str(dest), file_name=orig, is_active=True, created_by=current_user.id,
        )
        db.add(row)
        db.commit()
        db.refresh(row)
        return {"id": row.id, "scope": row.scope, "customer_name": row.customer_name, "is_active": row.is_active}
    finally:
        db.close()


@router.post("/sop/rules/{rule_id}/toggle")
async def sop_rules_toggle(rule_id: int, current_user=Depends(require_permission("tools.sop.manage"))):
    """启用/停用切换：启用时停用同 scope（+customer）其它生效规则。"""
    db = SessionLocal()
    try:
        r = db.query(SopRuleFile).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        if not r.is_active:
            q = db.query(SopRuleFile).filter(
                SopRuleFile.scope == r.scope, SopRuleFile.is_active == True, SopRuleFile.id != r.id,
            )
            if r.customer_name:
                q = q.filter(SopRuleFile.customer_name == r.customer_name)
            else:
                q = q.filter(SopRuleFile.customer_name.is_(None))
            for old in q.all():
                old.is_active = False
        r.is_active = not r.is_active
        db.commit()
        return {"id": r.id, "is_active": r.is_active}
    finally:
        db.close()


@router.delete("/sop/rules/{rule_id}")
async def sop_rules_delete(rule_id: int, current_user=Depends(require_permission("tools.sop.manage"))):
    """删除规则文件 + 物理文件。"""
    db = SessionLocal()
    try:
        r = db.query(SopRuleFile).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        path = r.file_path
        db.delete(r)
        db.commit()
        try:
            if path and os.path.exists(path):
                os.unlink(path)
        except Exception:
            logger.warning("[sop] 删除规则文件失败: %s", path)
        return {"ok": True}
    finally:
        db.close()


@router.get("/sop/rules/{rule_id}/download")
async def sop_rules_download(rule_id: int, current_user=Depends(require_permission("tools.sop.manage"))):
    """下载规则文件。"""
    from urllib.parse import quote
    db = SessionLocal()
    try:
        r = db.query(SopRuleFile).filter_by(id=rule_id).first()
        if not r:
            raise HTTPException(status_code=404, detail="规则文件不存在")
        if not os.path.exists(r.file_path):
            raise HTTPException(status_code=404, detail="文件不存在")
        return FileResponse(
            r.file_path, filename=r.file_name,
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(r.file_name)}"},
        )
    finally:
        db.close()


# ---------- SOP AI 分析 ----------

SOP_SYSTEM_PROMPT = """你是企业 SOP 合规审核专家。请严格依据给定的【大规则】，逐项核对用户提交的【源文件】、【结果文件】与【本轮规则文件】是否满足规则要求。
要求：
1. 只输出一个 JSON 对象，不要任何解释或 Markdown 代码块标记。
2. JSON 结构：
{
  "passed": true 或 false,
  "score": 0~100,
  "summary": "总体评价（中文，简洁）",
  "issues": ["不符合项1", "..."],
  "suggestions": ["完善建议1"],
  "details": {}
}"""


def _sop_read_file(path: str, max_chars: int = 6000) -> str:
    """读取文件内容供 AI 分析：Excel → 表头(含列字母映射) + 前 10 行数据；其它格式 → 文档解析。失败返回错误片段，不中断。"""
    if not path or not os.path.exists(path):
        return "（文件缺失）"
    ext = Path(path).suffix.lower()
    try:
        if ext in (".xlsx", ".xls", ".xlsm"):
            # 延迟导入，避免模块顶层加载 Aspose/.NET 产生副作用
            if str(PROJECT_ROOT) not in sys.path:
                sys.path.insert(0, str(PROJECT_ROOT))
            from excel_parser import IntelligentExcelParser
            sheets = IntelligentExcelParser().parse_excel_file(
                path, read_formulas=False, calculate_formulas=True,
                active_sheet_only=False, best_region_only=False,
            )
            parts = []
            for sd in sheets:
                parts.append(f"== Sheet {sd.sheet_name} ==")
                for region in (sd.regions or []):
                    # head_data 结构为 {表头名: 列字母}，反转为 {列字母: 表头名} 用于数据行渲染
                    letter_to_name = {str(v).upper(): k for k, v in (region.head_data or {}).items()}
                    if region.head_data:
                        parts.append("表头: " + " | ".join(f"{name}({col})" for name, col in region.head_data.items()))
                    for i, row in enumerate((region.data or [])[:10]):
                        cells = " | ".join(f"{letter_to_name.get(str(k).upper(), k)}: {v}" for k, v in row.items())
                        parts.append(f"第{i + 1}行: {cells}")
            text = "\n".join(parts)
        else:
            text = DocumentParser().parse_document(path)
        if not text:
            return "（空内容）"
        if len(text) > max_chars:
            text = text[:max_chars] + "\n...[内容过长已截断]"
        return text
    except Exception as e:
        return f"[文件解析失败: {e}]"


def _build_sop_prompt(entry: SopEntry, admin_rule_text: str, source_text: str, result_text: str, rule_text: Optional[str]) -> str:
    lines = [
        f"【客户名称】{entry.customer_name}",
        f"【SOP 描述】{entry.description or '（无）'}",
        "",
        "========== 大规则 ==========",
        admin_rule_text or "（无）",
        "",
        "========== 源文件 ==========",
        source_text or "（无）",
        "",
        "========== 结果文件 ==========",
        result_text or "（无）",
    ]
    if rule_text:
        lines += ["", "========== 本轮规则文件 ==========", rule_text]
    return "\n".join(lines)


def _run_sop_ai_analysis(round_id: int, entry_id: int):
    """后台 AI 分析：选规则 → 读文件 → 构建 prompt → 调 AI → 解析判定 → 更新状态。"""
    from ..ai_engine.ai_provider import AIProviderFactory
    db = SessionLocal()
    try:
        r = db.query(SopRound).filter_by(id=round_id).first()
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if not r or not entry:
            logger.warning("[sop] AI分析: 记录不存在 round=%s entry=%s", round_id, entry_id)
            return

        # 1) 选规则：客户专属优先，无则全局
        rule = (
            db.query(SopRuleFile)
            .filter(SopRuleFile.scope == "customer",
                    SopRuleFile.customer_name == entry.customer_name,
                    SopRuleFile.is_active == True)
            .first()
        )
        if not rule:
            rule = db.query(SopRuleFile).filter(SopRuleFile.scope == "global", SopRuleFile.is_active == True).first()
        r.rule_used_id = rule.id if rule else None

        # 2) 读文件
        src_paths = r.source_file_paths or ([r.source_file_path] if r.source_file_path else [])
        src_parts = []
        for i, p in enumerate(src_paths):
            src_parts.append(f"--- 源文件 {i + 1}: {os.path.basename(p)} ---\n{_sop_read_file(p)}")
        source_text = "\n\n".join(src_parts) if src_parts else "（无源文件）"
        result_text = _sop_read_file(r.result_file_path)
        rule_text = _sop_read_file(r.rule_file_path) if r.rule_file_path else None
        admin_rule_text = _sop_read_file(rule.file_path) if rule else "（无后台规则）"

        # 3) 构建 prompt 并落库（排查用）
        r.prompt_text = _build_sop_prompt(entry, admin_rule_text, source_text, result_text, rule_text)
        db.commit()

        # 4) 调 AI
        provider = AIProviderFactory.create_with_fallback()
        resp = provider.chat([
            {"role": "system", "content": SOP_SYSTEM_PROMPT},
            {"role": "user", "content": r.prompt_text},
        ]) or ""
        r.ai_response = resp

        # 5) 解析 JSON + 校验
        m = re.search(r"\{.*\}", resp, re.DOTALL)
        if not m:
            raise ValueError("AI 返回中未找到 JSON")
        verdict = SopAiVerdict.model_validate(json.loads(m.group(0)))
        r.ai_analysis = verdict.model_dump()
        r.ai_comment = verdict.summary
        r.ai_provider = getattr(provider, "model", None) or "unknown"
        r.finished_at = datetime.utcnow()

        # 6) 更新状态
        if verdict.passed:
            r.status = "ai_passed"
            entry.status = "ai_passed"
        else:
            r.status = "ai_failed"
            entry.status = "ai_failed"
        entry.latest_round_id = r.id
        db.commit()
        logger.info("[sop] AI分析完成 round=%s passed=%s", round_id, verdict.passed)
    except Exception as e:
        db.rollback()
        logger.exception("[sop] AI分析失败 round=%s", round_id)
        r = db.query(SopRound).filter_by(id=round_id).first()
        if r:
            r.status = "failed"
            r.error_message = str(e)
            r.finished_at = datetime.utcnow()
        entry = db.query(SopEntry).filter_by(id=entry_id).first()
        if entry:
            entry.status = "ai_failed"   # 技术失败也回退到"有问题"，允许重传
            if r:
                entry.latest_round_id = r.id
        db.commit()
    finally:
        db.close()
