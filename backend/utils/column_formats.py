"""共享列格式策略：保留明确格式，仅在日期格式缺失时补默认值。"""
from openpyxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_REVERSE, is_date_format


def resolve_column_format(schema=None, source_format=None):
    schema = schema or {}
    code = schema.get("number_format")
    if not code and schema.get("is_builtin_format"):
        code = BUILTIN_FORMATS.get(schema.get("number_format_id"))
    kind = schema.get("field_type")
    if kind in ("date", "datetime", "time"):
        for candidate in (code, source_format):
            if candidate and is_date_format(candidate):
                return candidate
        return {"date": "yyyy-mm-dd", "datetime": "yyyy-mm-dd hh:mm:ss",
                "time": "h:mm:ss"}[kind]
    return code or source_format or "General"


def _builtin_format_id(schema, code):
    schema = schema or {}
    number_id = schema.get("number_format_id")
    # 只有格式码仍与解析定义一致时才能复用编号；自定义编号不能跨工作簿复制。
    if (schema.get("is_builtin_format")
            and isinstance(number_id, int) and 0 <= number_id < 164
            and code == schema.get("number_format")):
        return number_id
    # 旧 schema 没有编号元数据：仅完全匹配内置格式码时恢复编号。
    # 新 schema 明确标为自定义时，不凭格式字符串覆盖该定义。
    if "is_builtin_format" not in schema and "number_format_id" not in schema:
        return BUILTIN_FORMATS_REVERSE.get(code)
    return None


def apply_aspose_column_format(style, schema):
    code = resolve_column_format(schema)
    number_id = _builtin_format_id(schema, code)
    if number_id is not None:
        style.Number = number_id
    else:
        style.Custom = code


def apply_openpyxl_column_format(cell, schema, source_format=None):
    code = resolve_column_format(schema, source_format)
    cell.number_format = code
    number_id = _builtin_format_id(schema, code)
    if number_id is not None:
        # openpyxl 没有设置内置编号的公开接口；地区专用编号不能仅靠格式字符串还原。
        cell._style.numFmtId = number_id
