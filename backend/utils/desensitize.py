"""数据脱敏工具：智能组表功能专用。

源文件样例（前 3 行）会随结构 json 一起传给 AI 帮助理解数据结构，
其中可能包含姓名、身份证、公司名等敏感信息，必须在传给 AI 前脱敏。

规则：按列名命中内置敏感词表 → 该列样例数据保格式脱敏
（长度与数字/字母位不变，AI 依然能看懂格式）。
只脱敏传给 AI 的内容；沙箱执行使用全量真实数据（数据全程本地）。
"""

import re
import random
import unicodedata

# 敏感词表：列名包含任一关键词即视为敏感列
SENSITIVE_KEYWORDS = [
    "公司", "企业", "单位名称", "姓名", "身份证", "证件号", "证件号码",
    "手机", "电话", "手机号", "手机号码", "电话号码",
    "邮箱", "邮件", "email", "E-mail", "地址", "住址", "居住地",
    "账号", "银行账号", "银行卡", "卡号", "开户行",
    "工资", "薪资", "薪金", "薪酬", "实发", "应发", "社保", "公积金",
    "公积金基数", "社保基数", "养老金", "保险基数",
]

# 姓名模式：连续 2-6 个汉字
_RE_NAME = re.compile(r"^[一-龥]{2,6}$")
# 身份证/长数字串：13-18 位数字
_RE_ID = re.compile(r"^\d{13,18}$")
# 手机号：1 开头 11 位数字
_RE_MOBILE = re.compile(r"^1\d{10}$")
# 电话号码：含 - 分隔的 7-12 位数字，或 3-4 位区号开头
_RE_PHONE = re.compile(r"^(\d{3,4}-)?\d{7,8}$")
# 邮箱
_RE_EMAIL = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
# 数字（含小数）
_RE_NUMBER = re.compile(r"^-?\d+(\.\d+)?$")
# 中文地址：含"省/市/区/县/镇/村/路/街/号"等
_RE_ADDRESS = re.compile(r".*[一-龥]*(省|市|区|县|镇|乡|村|路|街|巷|号|大厦|小区).*")


def _col_is_sensitive(col_name: str) -> bool:
    """按列名判断是否敏感列（命中任一关键词）。"""
    if not col_name:
        return False
    low = col_name.lower()
    for kw in SENSITIVE_KEYWORDS:
        if kw.lower() in low:
            return True
    return False


def _mask_keep_tail(s: str, keep_head: int, keep_tail: int = 0) -> str:
    """保长度脱敏：保留头部 keep_head 字符，保留尾部 keep_tail 字符，中间替为 *。"""
    s = str(s)
    n = len(s)
    if n <= keep_head + keep_tail:
        # 太短时全星（保一个字符位）
        return "*" * max(n, 1)
    return s[:keep_head] + "*" * (n - keep_head - keep_tail) + (s[-keep_tail:] if keep_tail else "")


def _desensitize_value(value, col_name: str) -> str:
    """对单个值按列类型保格式脱敏。返回脱敏后的字符串（数字脱敏后仍为数字字符串）。"""
    if value is None:
        return value
    s = str(value).strip()
    if not s:
        return value

    col = (col_name or "").lower()

    # 姓名类列
    if "姓名" in col or "名字" in col:
        if _RE_NAME.match(s):
            return _mask_keep_tail(s, 1)  # 张三 → 张*, 欧阳娜娜 → 欧***
    # 身份证类
    if any(k in col for k in ("身份证", "证件号", "证件号码")):
        if _RE_ID.match(s):
            return _mask_keep_tail(s, 3, 4)  # 110101********1234
        return _mask_keep_tail(s, 1, 1)
    # 手机号
    if any(k in col for k in ("手机",)):
        if _RE_MOBILE.match(s):
            return _mask_keep_tail(s, 3, 4)  # 138****1234
    # 电话号码
    if any(k in col for k in ("电话",)):
        if _RE_PHONE.match(s) or _RE_MOBILE.match(s):
            return _mask_keep_tail(s, 3, 4)
    # 邮箱
    if any(k in col for k in ("邮箱", "邮件", "email")):
        if _RE_EMAIL.match(s):
            local, _, domain = s.partition("@")
            if len(local) > 1:
                return local[0] + "*" * max(len(local) - 1, 1) + "@" + domain
            return "*" + "@" + domain
    # 地址
    if any(k in col for k in ("地址", "住址")):
        if len(s) > 4:
            return _mask_keep_tail(s, 3)  # 保留前3字
        return _mask_keep_tail(s, 1)
    # 公司/企业/单位
    if any(k in col for k in ("公司", "企业", "单位")):
        if len(s) > 4:
            return _mask_keep_tail(s, 2)  # 保留前2字
        return _mask_keep_tail(s, 1)
    # 账号/卡号
    if any(k in col for k in ("账号", "银行卡", "卡号", "开户行")):
        if _RE_NUMBER.match(s):
            return _mask_keep_tail(s, 4, 4)  # 6222********1234
        return _mask_keep_tail(s, 2, 2)
    # 金额/工资类：数字列保格式（同长度随机数字，保留小数位）
    if any(k in col for k in ("工资", "薪资", "薪金", "薪酬", "实发", "应发", "社保", "公积金", "养老金", "基数")):
        if _RE_NUMBER.match(s):
            neg = s.startswith("-")
            body = s.lstrip("-")
            if "." in body:
                int_part, frac_part = body.split(".")
                new_int = "".join(str(random.randint(0, 9)) for _ in int_part)
                new_frac = "".join(str(random.randint(0, 9)) for _ in frac_part)
                out = f"{new_int}.{new_frac}"
            else:
                out = "".join(str(random.randint(0, 9)) for _ in body)
            return ("-" if neg else "") + out
        # 非数字文本（如"按规定缴纳"）
        if len(s) > 2:
            return _mask_keep_tail(s, 1)
        return "*" * len(s)

    # 兜底：敏感列但值类型未识别 → 保长度掩码
    return _mask_keep_tail(s, 1, 1)


def desensitize_row(headers, row_values):
    """对一行数据按表头列名脱敏。

    Args:
        headers: 列名列表（与 row_values 对应，按 excel_parser 的列字母排序）
        row_values: 该行各列的值（dict {列字母: 值} 或 list）
    Returns:
        脱敏后的值（与原结构一致）
    """
    if isinstance(row_values, dict):
        out = {}
        for col_letter, v in row_values.items():
            col_name = headers.get(col_letter, "") if isinstance(headers, dict) else ""
            if col_name and _col_is_sensitive(col_name):
                out[col_letter] = _desensitize_value(v, col_name)
            else:
                out[col_letter] = v
        return out

    if isinstance(row_values, (list, tuple)):
        out = list(row_values)
        for i, v in enumerate(out):
            col_name = headers[i] if i < len(headers) else ""
            if col_name and _col_is_sensitive(col_name):
                out[i] = _desensitize_value(v, col_name)
        return out

    return row_values


def desensitize_region(head_data: dict, data_rows: list) -> list:
    """对区域数据行批量脱敏（excel_parser 区域结构专用）。

    Args:
        head_data: 区域表头 {列名: 列字母}（注意方向：列名→列字母）
        data_rows: 数据行列表，每行 {列字母: 值}
    Returns:
        脱敏后的数据行列表（浅拷贝，不修改原数据）
    """
    # 建 列字母→列名 映射（敏感列集合）
    sensitive_cols = set()
    for col_name, col_letter in (head_data or {}).items():
        if _col_is_sensitive(col_name):
            sensitive_cols.add(col_letter)

    if not sensitive_cols:
        return data_rows

    out = []
    for row in data_rows:
        new_row = dict(row)
        for col_letter in sensitive_cols:
            if col_letter in new_row:
                col_name = next((n for n, l in (head_data or {}).items() if l == col_letter), "")
                new_row[col_letter] = _desensitize_value(new_row[col_letter], col_name)
        out.append(new_row)
    return out


def build_structure_json(source_structure: dict) -> dict:
    """从源解析结构构建传给 AI 的脱敏结构 json。

    source_structure 结构：{文件名: {sheet名: {head_data, data(前N行样例), ...}}}
    返回同样结构但样例已脱敏的 dict（仅拷贝样例层，不修改原始结构）。
    """
    out = {}
    for fname, sheets in (source_structure or {}).items():
        out[fname] = {}
        for sheet_name, info in (sheets or {}).items():
            head_data = info.get("head_data") or {}
            data = info.get("data") or []
            masked = desensitize_region(head_data, data)
            out[fname][sheet_name] = {
                "columns": info.get("columns"),
                "column_letters": info.get("column_letters"),
                "formula_columns": info.get("formula_columns"),
                "column_formats": info.get("column_formats"),
                "sample_rows": masked,
            }
    return out
