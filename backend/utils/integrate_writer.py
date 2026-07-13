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
) -> Dict[str, Any]:
    """原地回填主表并（可选）追加差异 sheet，另存到 out_path。

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
        k = normalize_key(raw_key, normalize_keys)
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

    wb.Save(out_path, SaveFormat.Xlsx)
    logger.info(f"[integrate] 回填完成: 命中 {matched} 行, 覆盖 {overwritten} 格, 差异 {n_diff} 行 → {out_path}")
    return {"overwritten_cells": overwritten, "matched_rows": matched, "diff_rows": n_diff}


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
