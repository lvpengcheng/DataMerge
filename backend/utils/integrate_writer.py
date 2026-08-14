"""多表整合对比：Aspose 原地回填主表 + 可选追加差异 sheet。

设计要点：
- 只用 Aspose 打开主表 A、原地改激活页的覆盖列单元格、保存——绝不走 openpyxl/DataFrame
  往返（会丢公式/坏格式，触发数字+日期格式的 1900 闰年错位）。
- 逐行按【绝对行号 data_row_start..data_row_end】读 A 的关联键单元格（不依赖 parser 的
  region.data 位置——它会跳过空行/汇总行导致位置错位）；汇总/空行的键匹配不到源自然跳过。
- 覆盖只写【值】：命中源则 PutValue 覆盖（A 原有公式被替换为值），未命中/空值保留 A 原值。
- 长数字文本（身份证/卡号等）设 '@' 文本格式写入，避免科学计数/丢精度。
"""

import logging
from typing import Dict, List, Any, Optional

import aspose_init  # noqa: F401 — 确保 Aspose 已初始化
from Aspose.Cells import Workbook, SaveFormat  # type: ignore

from openpyxl.utils import column_index_from_string

from .merge_engine import normalize_key
from .integrate_engine import resolve_overwrites
from .source_sheet_writer import is_long_digit_text

logger = logging.getLogger(__name__)


def _col_idx(letter: str) -> int:
    """列字母(A/B/AA…) → 0-based 列索引。"""
    return column_index_from_string(str(letter).strip().upper()) - 1


def _read_cell_str(cell) -> str:
    """读单元格显示/字符串值（用于关联键）：优先 StringValue，回退 Value。"""
    try:
        sv = cell.StringValue
        if sv is not None:
            return str(sv)
    except Exception:
        pass
    v = cell.Value
    return "" if v is None else str(v)


def _put_value(cell, value):
    """写值到单元格（只写值，替换任何原公式）。长数字文本设 '@' 文本格式防科学计数。"""
    if value is None:
        return
    # 长数字文本（身份证/银行卡/手机号）→ 文本格式
    s = str(value)
    if is_long_digit_text(s) or (isinstance(value, str) and value.isdigit() and len(value) >= 12):
        try:
            style = cell.GetStyle()
            style.Custom = "@"
            cell.SetStyle(style)
            cell.PutValue(s, False)   # isConverted=False → 保持文本，不转数字
            return
        except Exception:
            pass
    try:
        cell.PutValue(value)
    except Exception:
        try:
            cell.PutValue(str(value))
        except Exception:
            logger.warning(f"[integrate] 写值失败，跳过: {value!r}")


def _apply_integration_impl(
    main_path: str,
    out_path: str,
    sheet_name: Optional[str],
    head_data: Dict[str, str],       # {A 列名 -> 列字母}
    a_key_col: str,
    data_row_start: int,             # 1-based 绝对行
    data_row_end: int,               # 1-based 绝对行
    overwrite_pairs: List[dict],
    source_indexes: Dict[str, Dict[str, dict]],
    normalize_keys: bool = True,
    diff_rows: Optional[List[dict]] = None,
    diff_order: str = "id_name",
    date_key_mode: str = "off",
) -> Dict[str, Any]:
    """（子进程执行体）原地回填主表并（可选）追加差异 sheet，另存到 out_path。

    Returns: {"overwritten_cells": int, "matched_rows": int, "diff_rows": int}
    """
    aspose_init.ensure_license()
    wb = Workbook(main_path)

    # 定位激活页所用的 sheet
    ws = None
    if sheet_name:
        try:
            ws = wb.Worksheets[sheet_name]
        except Exception:
            ws = None
    if ws is None:
        ws = wb.Worksheets[wb.Worksheets.ActiveSheetIndex]

    cells = ws.Cells

    # 关联键列索引 + 覆盖目标列索引
    if a_key_col not in head_data:
        raise ValueError(f"主表关联键列 '{a_key_col}' 不在表头中")
    key_ci = _col_idx(head_data[a_key_col])
    a_col_idx = {ac: _col_idx(head_data[ac]) for ac in {p.get("a_col") for p in (overwrite_pairs or [])} if ac in head_data}

    overwritten = 0
    matched = 0
    ds = max(1, int(data_row_start or 1))
    de = int(data_row_end or ds - 1)
    for row1 in range(ds, de + 1):
        r0 = row1 - 1
        raw_key = _read_cell_str(cells[r0, key_ci])
        k = normalize_key(raw_key, normalize_keys, date_key_mode)
        if k == "":
            continue
        vals = resolve_overwrites(k, overwrite_pairs, source_indexes)
        if not vals:
            continue
        matched += 1
        for ac, v in vals.items():
            ci = a_col_idx.get(ac)
            if ci is None:
                continue
            _put_value(cells[r0, ci], v)
            overwritten += 1

    # 追加差异 sheet（输出方式2）
    n_diff = 0
    if diff_rows:
        n_diff = _append_diff_sheet(wb, diff_rows, diff_order)

    # 写入值中可能含公式（源为计算结果文件时），Save 前先算一遍并缓存，
    # 否则打开 Excel 时公式不自动计算、需手动按回车才出值。
    try:
        wb.CalculateFormula()
    except Exception as _ce:
        logger.warning(f"[integrate] CalculateFormula 跳过（不阻断）: {_ce}")
    # 兜底：Aspose 算不了的复杂公式（INDIRECT/数组公式等）缓存值缺失，Excel 打开时
    # 不显示值、需点进单元格按回车才重算。保存时写 calcPr fullCalcOnLoad=1，
    # Excel 打开文件时自动全量重算所有公式，彻底消除"回车才生效"。
    try:
        wb.Settings.ForceFullCalculate = True
    except Exception as _ffc:
        logger.warning(f"[integrate] ForceFullCalculate 设置失败（忽略）: {_ffc}")
    wb.Save(out_path, SaveFormat.Xlsx)
    logger.info(f"[integrate] 回填完成: 命中 {matched} 行, 覆盖 {overwritten} 格, 差异 {n_diff} 行 → {out_path}")
    return {"overwritten_cells": overwritten, "matched_rows": matched, "diff_rows": n_diff}


def apply_integration(
    main_path: str,
    out_path: str,
    sheet_name: Optional[str],
    head_data: Dict[str, str],       # {A 列名 -> 列字母}
    a_key_col: str,
    data_row_start: int,             # 1-based 绝对行
    data_row_end: int,               # 1-based 绝对行
    overwrite_pairs: List[dict],
    source_indexes: Dict[str, Dict[str, dict]],
    normalize_keys: bool = True,
    diff_rows: Optional[List[dict]] = None,
    diff_order: str = "id_name",
    date_key_mode: str = "off",
) -> Dict[str, Any]:
    """原地回填主表并（可选）追加差异 sheet，另存到 out_path（在【独立子进程】执行）。

    防护背景: Aspose 打开主表 + CalculateFormula + Save 在特定文件（公式密集/超大）上
    会长时间计算、内存暴涨 → VM 假死。子进程超时（600s）/超内存强杀，主进程安全；
    失败时 raise ValueError，由端点转 4xx/5xx 明确报错。

    Returns: {"overwritten_cells": int, "matched_rows": int, "diff_rows": int}
    """
    from backend.utils.subprocess_runner import run_in_subprocess, default_max_memory_mb, default_timeout

    r = run_in_subprocess(
        "backend.utils.integrate_writer:_apply_integration_impl",
        (str(main_path), str(out_path), sheet_name, head_data, a_key_col,
         data_row_start, data_row_end, overwrite_pairs, source_indexes),
        kwargs={
            "normalize_keys": normalize_keys,
            "diff_rows": diff_rows,
            "diff_order": diff_order,
            "date_key_mode": date_key_mode,
        },
        timeout=default_timeout("write"), max_memory_mb=default_max_memory_mb(),
    )
    if r.success:
        return r.result
    reason = "超时（600s）" if r.timed_out else ("内存超限" if r.killed_by_memory else r.error)
    raise ValueError(f"整合回填失败（{reason}）: {main_path}")


def _append_diff_sheet_impl(path: str, diff_rows: Optional[List[dict]], diff_order: str = "id_name") -> int:
    """（子进程执行体）往【已生成文件】末尾追加差异 sheet 并原地另存。"""
    if not diff_rows:
        return 0
    aspose_init.ensure_license()
    wb = Workbook(path)
    n = _append_diff_sheet(wb, diff_rows, diff_order)
    # 与 _apply_integration_impl 一致：Excel 打开时全量重算（Aspose 算不了的复杂公式兜底）
    try:
        wb.Settings.ForceFullCalculate = True
    except Exception:
        pass
    wb.Save(path, SaveFormat.Xlsx)
    logger.info(f"[integrate] 追加差异 sheet: {n} 行 → {path}")
    return n


def append_diff_sheet(path: str, diff_rows: Optional[List[dict]], diff_order: str = "id_name") -> int:
    """往【已生成文件】末尾追加差异 sheet 并原地另存（在【独立子进程】执行）。

    差异值须由调用方基于最终生成文件（覆盖回填 + 公式重算后）重新解析比对得出，
    不能用主表覆盖前的缓存旧值。diff_rows 为空则不动文件、返回 0。
    失败时 raise ValueError（打开的就是整合结果=主表文件，特定文件需防护）。
    """
    from backend.utils.subprocess_runner import run_in_subprocess, default_max_memory_mb, default_timeout

    if not diff_rows:
        return 0
    r = run_in_subprocess(
        "backend.utils.integrate_writer:_append_diff_sheet_impl",
        (str(path), diff_rows, diff_order),
        timeout=default_timeout("write"), max_memory_mb=default_max_memory_mb(),
    )
    if r.success:
        return r.result
    reason = "超时" if r.timed_out else ("内存超限" if r.killed_by_memory else r.error)
    raise ValueError(f"追加差异 sheet 失败（{reason}）: {path}")


def _append_diff_sheet(wb, diff_rows: List[dict], diff_order: str) -> int:
    """在工作簿末尾追加差异 sheet：仅 3 列（姓名/身份证/差异类型），顺序按 diff_order。"""
    # 前两列顺序：id_name → 身份证,姓名 ; name_id → 姓名,身份证
    if diff_order == "name_id":
        headers = ["姓名", "身份证", "差异类型"]
        keys = ["姓名", "身份证", "差异类型"]
    else:
        headers = ["身份证", "姓名", "差异类型"]
        keys = ["身份证", "姓名", "差异类型"]

    # 唯一 sheet 名
    base_name = "差异"
    name = base_name
    existing = {wb.Worksheets[i].Name for i in range(wb.Worksheets.Count)}
    _n = 2
    while name in existing:
        name = f"{base_name}{_n}"
        _n += 1

    idx = wb.Worksheets.Add(name)
    ws = wb.Worksheets[idx] if isinstance(idx, int) else idx
    cells = ws.Cells

    for c, h in enumerate(headers):
        cell = cells[0, c]
        cell.PutValue(h)
        try:
            style = cell.GetStyle()
            style.Font.IsBold = True
            cell.SetStyle(style)
        except Exception:
            pass

    for r, row in enumerate(diff_rows, start=1):
        for c, key in enumerate(keys):
            v = row.get(key)
            if v is None or (isinstance(v, str) and v == ""):
                continue
            cell = cells[r, c]
            # 身份证列强制文本，避免长数字科学计数
            if key == "身份证":
                try:
                    st = cell.GetStyle(); st.Custom = "@"; cell.SetStyle(st)
                    cell.PutValue(str(v), False)
                    continue
                except Exception:
                    pass
            _put_cell_plain(cell, v)
    return len(diff_rows)


def _put_cell_plain(cell, value):
    try:
        cell.PutValue(value)
    except Exception:
        try:
            cell.PutValue(str(value))
        except Exception:
            pass
