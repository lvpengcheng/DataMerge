"""多表整合对比核心引擎（纯函数，无 Web/DB 依赖，便于单测）。

与「多表合并」(merge_engine) 的区别：本功能 1 张主表 A（模板）+ N 张对照表，
按关联键把对照表选定列的【值】回填到 A 对应列（覆盖，多源按优先级取首个非空），
并可对比选定列产出差异。输出保留 A 的整个工作簿（全部 sheet/公式/格式），
只原地覆盖激活页的覆盖列。

本模块只负责【算】：建对照表键索引、按键解析覆盖值、对比产出差异；
真正的原地写 A（Aspose）在 integrate_writer 完成。
"""

import logging
import ast as _ast
import re as _re
import operator as _operator
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

def build_key_index(df, key_col: str, normalize_keys: bool = True,
                    date_key_mode: str = "off") -> Dict[str, List[dict]]:
    """把一张对照表建成 {归一化键 -> [行dict, ...]}。

    同一主键的多行**全部保留**（供跨行汇总），不再只取首条。
    date_key_mode 对能识别成日期的键列生效（off/yearmonth/month/day），
    使 datetime / "2026-02" / "2月26日" 等归到同一粒度后可跨表对齐。
    """
    idx: Dict[str, List[dict]] = {}
    if df is None or key_col is None or key_col not in getattr(df, "columns", []):
        return idx
    for _, row in df.iterrows():
        k = normalize_key(row.get(key_col), normalize_keys, date_key_mode)
        if k == "":
            continue
        idx.setdefault(k, []).append(row.to_dict())
    return idx


def build_source_indexes(parsed: Dict[str, dict], key_map: Dict[str, str],
                         main_file: str, normalize_keys: bool = True,
                         date_key_mode: str = "off") -> Dict[str, dict]:
    """为除主表外的每张对照表建键索引。

    Args:
        parsed: {file: {"df": DataFrame, ...}}
        key_map: {file: 关联键列名}（含主表与各对照表）
        date_key_mode: 日期主键归一 off/yearmonth/month/day（与主表侧同参，保证两边同粒度）
    Returns:
        {file: {"cols": [列名...], "rows": {归一化键 -> [行dict,...]}}}  仅对照表。
        （结构较旧版多了一层：rows 存该键的**全部行**，cols 供公式解析列名。）
    """
    out: Dict[str, dict] = {}
    for f, fd in parsed.items():
        if f == main_file:
            continue
        df = fd.get("df")
        cols = list(df.columns) if df is not None else []
        out[f] = {"cols": cols, "rows": build_key_index(df, key_map.get(f), normalize_keys, date_key_mode)}
    return out


# ==================== 覆盖/对比取值：多行汇总 + 四则运算公式 ====================

_ARITH_BINOPS = {_ast.Add: _operator.add, _ast.Sub: _operator.sub,
                 _ast.Mult: _operator.mul, _ast.Div: _operator.truediv}


def _to_num(v) -> Optional[float]:
    """把值转 float；空/非数值返回 None。容忍千分位逗号、百分号、货币符号、全/半角空格。"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            import math
            return None if math.isnan(v) or math.isinf(v) else float(v)
        except Exception:
            return float(v)
    s = str(v).strip().replace(",", "").replace("，", "").replace(" ", "").replace("　", "")
    if s == "":
        return None
    pct = s.endswith("%")
    if pct:
        s = s[:-1]
    s = s.lstrip("¥$￥")
    try:
        n = float(s)
        return n / 100.0 if pct else n
    except Exception:
        return None


def _safe_arith_eval(expr: str) -> Optional[float]:
    """只对 +-*/、一元正负、数字与括号求值；出现任何其它节点/名字/函数调用 → 返回 None。"""
    try:
        node = _ast.parse(expr, mode="eval").body
    except Exception:
        return None

    def ev(n):
        if isinstance(n, _ast.BinOp) and type(n.op) in _ARITH_BINOPS:
            l, r = ev(n.left), ev(n.right)
            if l is None or r is None:
                return None
            if isinstance(n.op, _ast.Div) and r == 0:
                return None
            return _ARITH_BINOPS[type(n.op)](l, r)
        if isinstance(n, _ast.UnaryOp) and isinstance(n.op, (_ast.UAdd, _ast.USub)):
            v = ev(n.operand)
            if v is None:
                return None
            return v if isinstance(n.op, _ast.UAdd) else -v
        if isinstance(n, _ast.Constant) and isinstance(n.value, (int, float)) and not isinstance(n.value, bool):
            return float(n.value)
        if isinstance(n, getattr(_ast, "Num", ())):   # py<3.8 兼容
            return float(n.n)
        return None

    return ev(node)


def _cols_by_len_desc(cols: List[str]) -> List[str]:
    """列名按长度降序（长的先匹配，避免"工资"误配"基本工资"的子串）。"""
    return sorted([c for c in (cols or []) if c], key=len, reverse=True)


def _expr_columns(expr: str, cols: List[str]) -> List[str]:
    """expr 里引用到的对照表列名（长度降序，供替换/求和）。"""
    return [c for c in _cols_by_len_desc(cols) if c in expr]


def _expr_has_operator(expr: str, cols: List[str]) -> bool:
    """expr 去掉所有列名后是否还含 +-*/（判"纯单列引用"还是"公式"）。

    先扣掉列名再看运算符：列名本身可能含 '-'（如"太保填写-姓名"），不能误当运算符。
    """
    rest = expr
    for c in _cols_by_len_desc(cols):
        rest = rest.replace(c, " ")
    return any(op in rest for op in "+-*/")


def eval_source_expr(expr, rows: List[dict], cols: List[str]):
    """按【各列先跨行求和，再代入公式】算一个覆盖/对比值。

    - expr 是纯单列名：数值列→跨行求和；含非数值→取首个非空原值（保住姓名/备注等文本列）。
    - expr 是四则运算公式：每个被引用列跨行求和(非数值按 0)，代入 +-*/ 求值，返回数值。
    空/无行/公式非法/引用列不存在 → 返回 None（调用方据此保留原值/跳过）。
    """
    expr = str(expr or "").strip()
    if not expr or not rows:
        return None
    is_formula = _expr_has_operator(expr, cols)

    if not is_formula:
        col = expr if expr in cols else None
        if col is None:
            _refs = _expr_columns(expr, cols)
            col = _refs[0] if _refs else expr
        nums, first_non_empty, all_num = [], None, True
        for row in rows:
            v = row.get(col)
            if _is_empty(v):
                continue
            if first_non_empty is None:
                first_non_empty = v
            n = _to_num(v)
            if n is None:
                all_num = False
            else:
                nums.append(n)
        if first_non_empty is None:
            return None
        if all_num and nums:
            return round(sum(nums), 6)
        return first_non_empty   # 含文本 → 取首个非空，不汇总

    ref_cols = _expr_columns(expr, cols)
    if not ref_cols:
        return None
    subst = expr
    for c in ref_cols:
        s = 0.0
        for row in rows:
            n = _to_num(row.get(c))
            if n is not None:
                s += n
        subst = subst.replace(c, f"({s})")
    # 替换后应只剩数字/运算符/括号/小数点/空白；含其它字符 → 判为非法，返回 None
    if not _re.fullmatch(r"[0-9eE.+\-*/()\s]*", subst):
        return None
    val = _safe_arith_eval(subst)
    return None if val is None else round(val, 6)


def resolve_overwrites(key: str,
                       overwrite_pairs: List[dict],
                       source_indexes: Dict[str, dict]) -> Dict[str, Any]:
    """给定主表某行归一化键，按覆盖对（有序=优先级）解析出各 A 列要写入的值。

    Args:
        overwrite_pairs: [{"a_col","source_file","source_expr"|"source_col"}]  有序，靠前优先。
        source_indexes:  {file: {"cols":[...], "rows": {key -> [行,...]}}}
    Returns:
        {a_col: value}  仅含解析到【非空】值的列（未命中/空的列不写，保留 A 原值）。
        同一 a_col 被多对映射时，取首个非空源值。source_expr 支持"各列跨行求和后代入 +-*/"。
    """
    out: Dict[str, Any] = {}
    for pair in overwrite_pairs or []:
        ac = pair.get("a_col")
        if not ac or ac in out:
            continue  # 已被更高优先级填过
        f = pair.get("source_file")
        expr = pair.get("source_expr") or pair.get("source_col")
        entry = source_indexes.get(f) or {}
        rows = (entry.get("rows") or {}).get(key)
        if not rows:
            continue
        v = eval_source_expr(expr, rows, entry.get("cols") or [])
        if not _is_empty(v):
            out[ac] = v
    return out


# ==================== 对比产出差异 ====================

def _cmp2(v) -> str:
    """对比/展示归一化：数值四舍五入到 2 位小数（如 60.744→"60.74"、5000→"5000.00"），
    非数值回退 _norm_val（去空格 + 整数归一）。对比与展示用同一值，保证"所见即所比"。
    """
    if _is_empty(v):
        return ""
    n = _to_num(v)
    if n is None:
        return _norm_val(v)
    s = f"{n:.2f}"
    return "0.00" if s == "-0.00" else s


def compute_diffs(
    a_df,
    source_indexes: Dict[str, Dict[str, dict]],
    a_key_col: str,
    compare_pairs: List[dict],
    a_name_col: Optional[str],
    a_id_col: Optional[str],
    normalize_keys: bool = True,
    label_source: bool = False,
    date_key_mode: str = "off",
) -> List[dict]:
    """按键 join，逐对比列对比 A/B 值，产出差异行。

    Args:
        compare_pairs: [{"a_col","source_file","source_col"}]
        label_source: True 时在差异类型里标出对照文件名（多对照表时区分）。
        date_key_mode: 日期主键归一 off/yearmonth/month/day（与对照表侧同参）。
    Returns:
        [{"姓名","身份证","差异类型"}]  仅含有差异的键。
        差异类型：以【对比字段的 A 列名】为字段名，旧=A 值、新=对照值，`字段名: 旧→新`，多项分号拼接。
    """
    diffs: List[dict] = []
    if a_df is None or not compare_pairs:
        return diffs
    for _, row in a_df.iterrows():
        k = normalize_key(row.get(a_key_col), normalize_keys, date_key_mode)
        if k == "":
            continue
        parts = []
        for pair in compare_pairs:
            ac = pair.get("a_col")
            f = pair.get("source_file")
            expr = pair.get("source_expr") or pair.get("source_col")
            entry = source_indexes.get(f) or {}
            b_rows = (entry.get("rows") or {}).get(k)
            if not b_rows:
                continue  # 该对照表无此键 → 该对比对跳过
            av = row.get(ac)
            bv = eval_source_expr(expr, b_rows, entry.get("cols") or [])   # 各列跨行求和后代入公式
            old = _cmp2(av)   # 数值统一到 2 位小数后再比对/展示
            new = _cmp2(bv)
            if old != new:
                prefix = f"[{f}] " if label_source else ""
                parts.append(f"{prefix}{ac}: {old}→{new}")
        if parts:
            diffs.append({
                "姓名": row.get(a_name_col) if a_name_col else "",
                "身份证": row.get(a_id_col) if a_id_col else "",
                "差异类型": "; ".join(parts),
            })
    return diffs
