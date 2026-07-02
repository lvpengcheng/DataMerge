"""源_ 工作表写入保真 helper（供 formula / template 两种模式的生成脚本复用）。

集中三个纯函数，避免在两份生成代码模板串里各维护一份并踩两套转义规则：
- is_date_keyword_column：按列名关键词判定是否日期列（与 formula 侧 _is_date_column 一致）。
- dt_to_excel_serial：datetime → Excel 序列号（非日期列却被套了日期格式时，逆转回底层数值）。
- is_long_digit_text：≥12 位纯数字串（身份证/银行卡/手机）判定，写入时应保持文本格式。

设计背景：Excel 里"日期 = 数字 + 日期格式"，Aspose 对"真日期"和"套了 yyyy-mm-dd
格式的普通数字"都报 IsDateTime，无法凭类型区分。故由列名关键词决定哪些列当日期；
非日期关键词列即使套了日期格式，也读其底层数值（用 dt_to_excel_serial 逆转）。
"""

from datetime import datetime

# 排除误匹配（如"工作日数"、"节日"等）
_DATE_EXCLUDE_KEYWORDS = ['日数', '日常', '日志', '日报', '日均', '节日', '假日', '工日', 'update', 'today']
_DATE_KEYWORDS = ['日期', 'date', '入职日', '离职日', '生效日', '截止日', '转正日', '生日',
                  '出生日', '开始日', '结束日', '签订日', '到期日', '发放日', '申请日',
                  '创建时间', '更新时间', '时间戳', 'datetime', 'timestamp']

# Excel 日期纪元：以 1899-12-30 为 day0，(dt-epoch).days 对 ≥1900-03-01 的日期与
# Excel 序列号完全一致（-1 天偏移恰好抵消 Excel 虚构的 1900-02-29 假闰年）。
_EXCEL_EPOCH = datetime(1899, 12, 30)


def is_date_keyword_column(col_name) -> bool:
    """仅按列名关键词判断是否为日期列（不做数据内容探测）。"""
    name = str(col_name).lower().strip()
    for exc in _DATE_EXCLUDE_KEYWORDS:
        if exc in name:
            return False
    for kw in _DATE_KEYWORDS:
        if kw in name:
            return True
    return False


def dt_to_excel_serial(v):
    """datetime → Excel 序列号（float）。非 datetime 原样返回。

    用于非日期列里被套了日期格式的单元格：解析层已把它读成 datetime，
    这里逆转回底层数值（序列号），写入 源_ sheet 时按常规数字呈现。
    """
    if isinstance(v, datetime):
        d = v - _EXCEL_EPOCH
        return d.days + d.seconds / 86400.0
    return v


def is_long_digit_text(v) -> bool:
    """≥12 位纯数字串（身份证/银行卡/手机号等），写入时应设文本格式避免科学计数/丢精度。"""
    return isinstance(v, str) and v.isdigit() and len(v) >= 12
