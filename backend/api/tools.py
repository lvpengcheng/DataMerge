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
                try:
                    from ..utils.source_normalizer import normalize_misformatted_dates
                    normalize_misformatted_dates(str(dest))
                except Exception as _ne:
                    logger.warning(f"[merge] 日期格式规范化失败（继续）: {_ne}")
                info = _parse_file_to_df(str(dest))
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
            generate_from_template(str(out_path), str(tpl_abs), {"DT": df}, mode="fill")
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
