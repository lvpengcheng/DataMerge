"""源_ 工作表写入保真 helper（供 formula / template 两种模式的生成脚本复用）。

集中三个纯函数，避免在两份生成代码模板串里各维护一份并踩两套转义规则：
- is_date_keyword_column：按列名关键词判定是否日期列（与 formula 侧 _is_date_column 一致）。
- dt_to_excel_serial：datetime → Excel 序列号（非日期列却被套了日期格式时，逆转回底层数值）。
- is_long_digit_text：≥12 位纯数字串（身份证/银行卡/手机）判定，写入时应保持文本格式。

设计背景：Excel 里"日期 = 数字 + 日期格式"，Aspose 对"真日期"和"套了 yyyy-mm-dd
格式的普通数字"都报 IsDateTime，无法凭类型区分。故由列名关键词决定哪些列当日期；
非日期关键词列即使套了日期格式，也读其底层数值（用 dt_to_excel_serial 逆转）。
"""

from datetime import date, datetime, timedelta
from numbers import Real
import math
import re

# 排除误匹配（如"工作日数"、"节日"等）；"工作/加班/出勤时间"等是时长(数字)不是日期
_DATE_EXCLUDE_KEYWORDS = ['日数', '日常', '日志', '日报', '日均', '节日', '假日', '工日',
                          '工作时间', '加班时间', '出勤时间', '休息时间', '时间段',
                          'update', 'today']
_DATE_KEYWORDS = ['日期', 'date', '入职日', '离职日', '生效日', '截止日', '转正日', '生日',
                  '出生日', '开始日', '结束日', '签订日', '到期日', '发放日', '申请日',
                  # "时间"系列（事件时点=日期；时长类已在排除表里剔除）
                  '入职时间', '离职时间', '转正时间', '生效时间', '截止时间', '签订时间',
                  '到期时间', '发放时间', '申请时间', '出生时间', '开始时间', '结束时间',
                  '登记日', '录用日', '解除日', '终止日',
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


def coerce_source_date(v, epoch=_EXCEL_EPOCH):
    """把日期关键词列的值尽力还原成真 datetime；空值一律返回 None（保持空单元格）。

    用于模板模式写 源_ sheet：源里日期列可能存成文本（"2020-06-22"）或裸 Excel 序列号
    （42430），pandas 读成 str/int/float，不会被当日期。这里对日期关键词列统一还原：
      - 各种空（None/NaN/NaT/空串/纯空格）→ None（→ 空单元格，绝不写成 1899-12-30 或 NaT）
      - 已是 datetime → 原样返回
      - 文本 → pd.to_datetime 解析；解析不出则**保留原文本**（不丢数据、不误判）
      - 数字 → 按 Excel 纪元还原成日期；<=0 视为空（0=1899-12-30 基本是占位空）

    注意：pd.NaT 是 datetime 子类，isinstance(NaT, datetime) 为 True，必须先拦掉。
    """
    import pandas as pd
    if v is None or v is pd.NaT or v is pd.NA:
        return None
    if isinstance(v, Real) and not math.isfinite(v):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, datetime):           # 已是日期（含 pd.Timestamp 子类）→ 原样
        return v
    if isinstance(v, date):
        return datetime.combine(v, datetime.min.time())
    # YYYYMMDD 是业务日期，不能当作两千万天的 Excel 序列号。
    raw = str(v).strip()
    if re.fullmatch(r"\d{8}(?:\.0+)?", raw):
        try:
            return datetime.strptime(raw.split('.')[0], '%Y%m%d')
        except ValueError:
            return v
    if isinstance(v, str):
        s = v.strip()
        if not s:                         # 空串/纯空格 → 空
            return None
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", s):
            # 年份文本含义不明确，保留；五位数字文本可明确按序列号处理。
            if len(s.split('.')[0].lstrip('+-')) >= 5:
                return coerce_source_date(float(s), epoch=epoch)
            return v
        s = s.replace('年', '-').replace('月', '-').replace('日', '')
        try:
            ts = pd.to_datetime(s, errors="coerce")
        except Exception:
            return v
        return ts.to_pydatetime() if pd.notna(ts) else v   # 解析不出 → 保留原文本
    if isinstance(v, Real):
        if v <= 0:                        # 0/负数 → 空（避免 0→1899-12-30）
            return None
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(float(v), epoch=epoch)
        except Exception:
            return v
    return v


def write_source_dataframe(ws, df, column_schemas=None, column_formats=None, header_fill=None):
    """一次遍历写源表；固定列策略，避免反复 ws['A'] 扫描整个工作表计算边界。"""
    import pandas as pd
    from openpyxl.styles import Font
    from backend.utils.data_helpers import _schema_text_value

    schemas, formats = column_schemas or {}, column_formats or {}
    policies = []
    for name in df.columns:
        schema = schemas.get(name) or {}
        kind = schema.get('field_type')
        is_date = kind in ('date', 'datetime') or (not kind and is_date_keyword_column(name))
        fmt = schema.get('number_format') or formats.get(name) or ('yyyy-mm-dd' if is_date else 'General')
        policies.append((kind, is_date, fmt))
    ws.append(list(df.columns))
    bold = Font(bold=True)
    for cell in ws[1]:
        cell.number_format = 'General'
        if header_fill is not None:
            cell.fill = header_fill
            cell.font = bold
    for ri, row in enumerate(df.itertuples(index=False, name=None), 2):
        for ci, (value, (kind, is_date, fmt)) in enumerate(zip(row, policies), 1):
            if value is None or value is pd.NA or pd.isna(value):
                value = None
            elif is_date:
                value = coerce_source_date(value)
            elif kind == 'text':
                value = _schema_text_value(value)
            elif isinstance(value, datetime):
                value = dt_to_excel_serial(value)
            cell = ws.cell(ri, ci, value)
            cell.number_format = fmt
            # 源数据的文本（包括 = 开头的说明）不得变成新增的可执行公式。
            if isinstance(value, str):
                cell.data_type = 's'
                if kind == 'text' or is_long_digit_text(value):
                    cell.number_format = '@'


def is_long_digit_text(v) -> bool:
    """≥12 位纯数字串（身份证/银行卡/手机号等），写入时应设文本格式避免科学计数/丢精度。"""
    return isinstance(v, str) and v.isdigit() and len(v) >= 12
