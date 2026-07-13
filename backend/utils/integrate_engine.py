"""多表整合对比核心引擎（纯函数，无 Web/DB 依赖，便于单测）。

与「多表合并」(merge_engine) 的区别：本功能 1 张主表 A（模板）+ N 张对照表，
按关联键把对照表选定列的【值】回填到 A 对应列（覆盖，多源按优先级取首个非空），
并可对比选定列产出差异。输出保留 A 的整个工作簿（全部 sheet/公式/格式），
只原地覆盖激活页的覆盖列。

本模块只负责【算】：建对照表键索引、按键解析覆盖值、对比产出差异；
真正的原地写 A（Aspose）在 integrate_writer 完成。
"""

import logging
from typing import Dict, List, Any, Optional

from .merge_engine import normalize_key, _norm_header, _norm_val, _is_empty

logger = logging.getLogger(__name__)


# 姓名 / 身份证 列名关键词（用于差异 sheet 定位；可被前端人工覆盖）
_NAME_HINTS = ["姓名", "员工姓名", "人员姓名", "name", "中文姓名"]
_ID_HINTS = ["身份证号码", "身份证号", "身份证", "证件号码", "证件号", "idcard", "id card"]


def _hint_column(columns, hints) -> Optional[str]:
    """按关键词优先级在列名里找第一个命中列（归一化后子串匹配）。找不到返回 None。"""
    cols = list(columns or [])
    if not cols:
        return None
    norm = {c: _norm_header(c) for c in cols}
    for h in hints:
        hh = _norm_header(h)
        for c in cols:
            if hh and hh in norm[c]:
                return c
    return None


def guess_name_column(columns) -> Optional[str]:
    """猜「姓名」列。"""
    return _hint_column(columns, _NAME_HINTS)


def guess_id_column(columns) -> Optional[str]:
    """猜「身份证」列。"""
    return _hint_column(columns, _ID_HINTS)


# ==================== 对照表键索引 ====================

def build_key_index(df, key_col: str, normalize_keys: bool = True) -> Dict[str, dict]:
    """把一张对照表建成 {归一化键 -> 行dict}（同键重复取首条）。"""
    idx: Dict[str, dict] = {}
    if df is None or key_col is None or key_col not in getattr(df, "columns", []):
        return idx
    for _, row in df.iterrows():
        k = normalize_key(row.get(key_col), normalize_keys)
        if k == "" or k in idx:
            continue
        idx[k] = row.to_dict()
    return idx


def build_source_indexes(parsed: Dict[str, dict], key_map: Dict[str, str],
                         main_file: str, normalize_keys: bool = True) -> Dict[str, Dict[str, dict]]:
    """为除主表外的每张对照表建键索引。

    Args:
        parsed: {file: {"df": DataFrame, ...}}
        key_map: {file: 关联键列名}（含主表与各对照表）
    Returns:
        {file: {归一化键 -> 行dict}}  仅对照表。
    """
    out: Dict[str, Dict[str, dict]] = {}
    for f, fd in parsed.items():
        if f == main_file:
            continue
        out[f] = build_key_index(fd.get("df"), key_map.get(f), normalize_keys)
    return out


# ==================== 覆盖：按键解析覆盖值 ====================

def resolve_overwrites(key: str,
                       overwrite_pairs: List[dict],
                       source_indexes: Dict[str, Dict[str, dict]]) -> Dict[str, Any]:
    """给定主表某行归一化键，按覆盖对（有序=优先级）解析出各 A 列要写入的值。

    Args:
        overwrite_pairs: [{"a_col","source_file","source_col"}]  有序，靠前优先。
        source_indexes:  {file: {key -> row}}
    Returns:
        {a_col: value}  仅含解析到【非空】值的列（未命中/空的列不写，保留 A 原值）。
        同一 a_col 被多对映射时，取首个非空源值。
    """
    out: Dict[str, Any] = {}
    for pair in overwrite_pairs or []:
        ac = pair.get("a_col")
        if not ac or ac in out:
            continue  # 已被更高优先级填过
        f, sc = pair.get("source_file"), pair.get("source_col")
        row = source_indexes.get(f, {}).get(key)
        if not row:
            continue
        v = row.get(sc)
        if not _is_empty(v):
            out[ac] = v
    return out


# ==================== 对比产出差异 ====================

def compute_diffs(
    a_df,
    source_indexes: Dict[str, Dict[str, dict]],
    a_key_col: str,
    compare_pairs: List[dict],
    a_name_col: Optional[str],
    a_id_col: Optional[str],
    normalize_keys: bool = True,
    label_source: bool = False,
) -> List[dict]:
    """按键 join，逐对比列对比 A/B 值，产出差异行。

    Args:
        compare_pairs: [{"a_col","source_file","source_col"}]
        label_source: True 时在差异类型里标出对照文件名（多对照表时区分）。
    Returns:
        [{"姓名","身份证","差异类型"}]  仅含有差异的键。
        差异类型：以【对比字段的 A 列名】为字段名，旧=A 值、新=对照值，`字段名: 旧→新`，多项分号拼接。
    """
    diffs: List[dict] = []
    if a_df is None or not compare_pairs:
        return diffs
    for _, row in a_df.iterrows():
        k = normalize_key(row.get(a_key_col), normalize_keys)
        if k == "":
            continue
        parts = []
        for pair in compare_pairs:
            ac, f, sc = pair.get("a_col"), pair.get("source_file"), pair.get("source_col")
            b_row = source_indexes.get(f, {}).get(k)
            if not b_row:
                continue  # 该对照表无此键 → 该对比对跳过
            av, bv = row.get(ac), b_row.get(sc)
            if _norm_val(av) != _norm_val(bv):
                old = "" if _is_empty(av) else str(av).strip()
                new = "" if _is_empty(bv) else str(bv).strip()
                prefix = f"[{f}] " if label_source else ""
                parts.append(f"{prefix}{ac}: {old}→{new}")
        if parts:
            diffs.append({
                "姓名": row.get(a_name_col) if a_name_col else "",
                "身份证": row.get(a_id_col) if a_id_col else "",
                "差异类型": "; ".join(parts),
            })
    return diffs
