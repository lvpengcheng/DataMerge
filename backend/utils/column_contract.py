"""统一列契约：输出工作簿"值类型 + 显示格式"的唯一裁决点与唯一写入点。

以前格式由 6 个后处理函数各自判断、按执行顺序覆盖（主键转文本、模板格式还原、
源_ 默认样式兜底、expected schema、汇总区重刷、脚本显式格式），谁最后跑谁说了算。
现在它们都退化为"规则来源层"。源数据/模板填充/原表复制入口复用 canon_value，
最终由 apply_column_contract 在重算前裁决，避免各阶段采用不同的主键算法。

每个单元格分别裁决两件事，按优先级取第一个给出答案的层：
  - kind（值类型，决定 VLOOKUP 能否匹配）：key / text / date / datetime / number / raw
  - format（number_format，决定显示）

层次（高 → 低）：
  1. explicit  脚本里显式写的单元格 number_format（仅显示；不豁免主键值归一）
  2. spec      脚本 _COLUMN_SPECS 常量（智训"执行修改"里用户指定的列类型/格式）
  3. identity  明确的主键/ID 默认规范文本；含义不明确的 code/编号 必须显式 type=key
  4. template  本次实际使用的模板类型/格式（主键由统一文本策略裁决）
  5. schema    训练推断的 expected_structure column_schemas（模板未覆盖的位置兜底）
  6. inherited 源_ sheet 继承模板非 General 默认样式的单元格 → General/日期
  都没有 → 保持原样。v2 管线也保护主键，但不使用训练推断兜底。
  显式 type=key 表示规范文本，type=raw 表示保留类型及当前格式。
  normalize_match_key 为历史辅助键方案保留；新生成脚本直接查询统一为文本的主键。
  本层不能恢复脚本已转换的源类型，也不能恢复 Excel 已丢失的长编号精度。

_COLUMN_SPECS 写法（sheet 名用 "*" 表示所有含该列的 sheet）：
    _COLUMN_SPECS = {
        "源_考勤": {"考勤日期": {"type": "date", "format": "yyyy-mm-dd"},
                    "工号": "key"},
        "*": {"身份证号": {"type": "key", "format": "@"}},
        "工资表": {"入职日期": {"type": "date", "parse": "%d/%m/%Y"}},
    }
canon_value 是唯一的值转换函数，数据注入沙箱前的归一也应调用它，保证两端同规则。
"""
import ast
import logging
import math
import os
import re
from decimal import Decimal, InvalidOperation
from numbers import Integral, Real
from collections import Counter
from datetime import date, datetime

logger = logging.getLogger(__name__)

from backend.utils.column_spec_schema import KINDS, decode_column_specs, literal_column_specs
_HEADER_SCAN_ROWS = 30
_EXCEL_EPOCH = datetime(1899, 12, 30)
_DATE1904_OFFSET = 1462
_DEFAULT_DATE_FORMATS = {"date": "yyyy-mm-dd", "datetime": "yyyy-mm-dd hh:mm:ss"}
# 不可见字符删除、全角空格/不换行空格转半角空格、全角 ASCII 转半角
_TEXT_TABLE = {0x200b: None, 0x200c: None, 0x200d: None, 0xfeff: None,
               0x00a0: " ", 0x3000: " ",
               **{c: c - 0xFEE0 for c in range(0xFF01, 0xFF5F)}}


def column_contract_enabled() -> bool:
    """默认启用统一列契约；仅显式 legacy 使用旧后处理链。"""
    return os.getenv("COLUMN_CONTRACT_MODE", "unified").strip().lower() != "legacy"


# ---------------------------------------------------------------- 规则读取

def read_column_specs(script_code):
    """用 AST 读取脚本顶层 _COLUMN_SPECS 字面量（不执行代码）；返回规范化后的 dict。"""
    if not script_code:
        return {}
    try:
        tree = ast.parse(script_code)
    except SyntaxError:
        return {}
    specs = None
    for node in tree.body:
        targets = node.targets if isinstance(node, ast.Assign) else ([node.target] if isinstance(node, ast.AnnAssign) else [])
        if any(isinstance(t, ast.Name) and t.id == "_COLUMN_SPECS" for t in targets):
            try:
                specs = literal_column_specs(node.value)
            except (ValueError, TypeError) as e:
                raise ValueError(f'_COLUMN_SPECS 声明非法，未执行代码: {e}') from e
    return normalize_column_specs(specs)


def normalize_column_specs(specs):
    """所有入口共用同一声明解析，再构造运行时的列名匹配索引。"""
    canonical = decode_column_specs(specs, allow_internal=True)
    result = {}
    for sheet, columns in canonical.items():
        for column, spec in columns.items():
            if isinstance(spec, str):
                spec = {'type': spec}
            result.setdefault(sheet, {})[_header_key(column)] = {
                'type': spec.get('type'), 'format': spec.get('format'),
                'parse': spec.get('parse'), 'name': str(column).strip()}
    return result


# ---------------------------------------------------------------- 值转换

def _clean_text(s):
    return s.translate(_TEXT_TABLE).strip()


def _canon_key(v):
    if v is None or isinstance(v, (bool, datetime, date)):
        return v
    if isinstance(v, Integral):
        return str(int(v))
    if isinstance(v, Real):
        if not math.isfinite(v):
            return v
        return str(int(v)) if float(v).is_integer() else repr(float(v))
    if isinstance(v, str):
        return _clean_text(v) or v
    return v


def normalize_match_key(value):
    """独立生成匹配键，不改原值；空值不匹配，文本前导零保持有意义。

    123 / 123.0 / '123' 匹配；'00123' 不与 '123' 合并。
    调用方仍需检查同一匹配键的多条记录，不能默认覆盖重复记录。
    """
    if value is None or isinstance(value, (bool, date, datetime)):
        return None
    if isinstance(value, Real) and not math.isfinite(value):
        return None
    result = _canon_key(value)
    return result if isinstance(result, str) and result.strip() else None


def _is_identifier_column(header):
    # 只识别明确标识字段。不能因为含 code / 编号 / id 就改变业务数值列。
    text = re.sub(r'[（(][^()（）]*[)）]', '', str(header or ''))
    name = re.sub(r'[\s_-]+', '', text).casefold()
    if name in {'psid', 'id', 'employeeid', 'employeecode', 'employeenumber',
                'staffid', 'staffcode', 'staffnumber', 'nationalid',
                'identitynumber', 'idnumber', 'passportnumber', 'bankaccount',
                'bankaccountnumber', 'cardnumber', 'taxid', 'payrollid',
                '工号', '员工工号', '外服工号', '员工编号', '员工号', '职工号',
                '人员编号', '工作证号', '员工id', '身份证', '身份证号', '身份证号码',
                '证件号', '证件号码', '银行卡', '银行卡号', '银行账号', '银行账户',
                '社保号', '社保账号', '公积金号', '公积金账号', '手机号', '手机号码',
                '联系电话', '税号', '纳税人识别号'}:
        return True
    return False


def _sheet_specs(specs, sheet):
    aliases = (str(sheet).removeprefix('源_'), str(sheet), '源_' + str(sheet).removeprefix('源_'))
    declared = dict(specs.get('*', {}))
    for alias in dict.fromkeys(aliases):
        declared.update(specs.get(alias, {}))
    return declared


def _column_rules(columns, specs, sheet):
    """所有入口共用：精确声明覆盖默认主键规则；format-only 不取消默认键类型。"""
    rules = {}
    declared = _sheet_specs(specs, sheet)
    for column in columns:
        key = _header_key(column)
        explicit = declared.get(key, {})
        identifier = _is_identifier_column(column)
        if not identifier and not explicit:
            continue
        rule = {'type': 'key' if identifier else None, 'format': '@' if identifier else None,
                'parse': None, 'name': str(column)}
        rule.update({k: v for k, v in explicit.items() if v is not None})
        if not explicit.get('format'):
            rule['format'] = '@' if rule.get('type') in ('key', 'text') else None
        if rule:
            rules[key] = rule
    return rules


def normalize_source_keys(source_data, specs=None):
    """统一内存主键，复制发生改动的 DataFrame，不污染预加载缓存或源文件。"""
    if not isinstance(source_data, dict):
        return source_data
    specs = normalize_column_specs(specs or {})
    effective = {}
    groups = {}
    for sheet, entry in source_data.items():
        if isinstance(entry, dict) and entry.get('df') is not None:
            group = groups.setdefault(id(entry), {'names': [], 'declared': {}})
            group['names'].append(str(sheet))
            for name, rule in _sheet_specs(specs, sheet).items():
                previous = group['declared'].get(name, {})
                if previous.get('type') and rule.get('type') and previous['type'] != rule['type']:
                    raise ValueError(f'同一源表的别名对 {name} 声明了冲突类型')
                group['declared'][name] = {**previous, **{k: v for k, v in rule.items() if v is not None}}
    for sheet, entry in source_data.items():
        df = entry.get('df') if isinstance(entry, dict) else None
        if df is not None:
            effective[str(sheet)] = _column_rules(df.columns, {'*': groups[id(entry)]['declared']}, sheet)
            for column in df.columns:
                rule = effective[str(sheet)].get(_header_key(column), {})
                if rule.get('type') == 'key':
                    risky = sum(isinstance(v, Real) and not isinstance(v, bool)
                                and math.isfinite(v) and abs(v) >= 1e15 for v in df[column])
                    if risky:
                        logger.warning('[列契约] %s.%s 有 %s 个超过 15 位的数值主键，可能已丢失精度；请核对原始文本',
                                       sheet, column, risky)
    return apply_specs_to_source_data(source_data, effective)


def normalize_workbook_keys(wb, specs=None):
    """openpyxl 填充前后统一模板/结果主键；保留公式及字体、边框、行高。"""
    specs = normalize_column_specs(specs or {})
    changed = 0
    for ws in wb.worksheets:
        headers = {}
        # 多行表头也支持；每列以首个明确标识表头或显式声明表头为准。
        wanted = set(specs.get('*', {}))
        for alias in (ws.title, ws.title.removeprefix('源_'), '源_' + ws.title):
            wanted.update(specs.get(alias, {}))
        for row in ws.iter_rows(max_row=min(ws.max_row, _HEADER_SCAN_ROWS)):
            for cell in row:
                key = _header_key(cell.value)
                if cell.column not in headers and (_is_identifier_column(cell.value) or key in wanted):
                    headers[cell.column] = (cell.row, str(cell.value))
        rules = _column_rules([name for _, name in headers.values()], specs, ws.title)
        for column, (header, name) in headers.items():
            rule = rules.get(_header_key(name), {})
            if rule.get('type') != 'key':
                continue
            for row in range(header + 1, ws.max_row + 1):
                cell = ws.cell(row, column)
                if cell.value is None or cell.data_type == 'f':
                    continue
                value = canon_value(cell.value, 'key')
                if type(value) is not type(cell.value) or value != cell.value:
                    cell.value = value
                    changed += 1
                if isinstance(value, str):
                    cell.data_type = 's'
                    cell.number_format = rule.get('format') or '@'
    return changed


def normalize_frame_tables(tables, specs=None):
    """模块化模式的 {sheet: DataFrame} 适配，复用相同值转换与缓存保护。"""
    import pandas as pd
    if not isinstance(tables, dict) or any(not isinstance(df, pd.DataFrame) for df in tables.values()):
        raise ValueError('表数据必须为 {sheet: DataFrame}')
    entries = {name: {'df': df} for name, df in tables.items()}
    return {name: entry['df'] for name, entry in normalize_source_keys(entries, specs).items()}


def normalize_modular_source_keys(store, specs=None):
    """模块化加载接口 files/filename/sheets 适配，不修改加载缓存。"""
    if not isinstance(store, dict) or not isinstance(store.get('files'), dict):
        raise ValueError('模块化源数据必须包含 files 字典')
    files = {}
    specs = normalize_column_specs(specs or {})
    from pathlib import Path
    for filename, entry in store['files'].items():
        sheets = entry.get('sheets') if isinstance(entry, dict) else None
        if not isinstance(sheets, dict):
            raise ValueError('模块化源文件必须包含 sheets 字典')
        effective = {}
        for name in sheets:
            rules = _sheet_specs(specs, name)
            qualified = str(Path(filename).stem) + '_' + str(name)
            # 文件限定声明比同名 sheet 声明更精确；避免同名源表串用声明。
            rules = {**rules, **specs.get(qualified, {}), **specs.get('源_' + qualified, {})}
            effective[str(name)] = rules
        files[filename] = {**entry, 'sheets': normalize_frame_tables(sheets, effective)}
    return {**store, 'files': files}


def normalize_native_keys(wb, specs=None):
    """原表整表复制后也要处理主键，在保存/公式重算前调用；不改原始文件。"""
    from Aspose.Cells import CellValueType
    specs = normalize_column_specs(specs or {})
    changed = 0
    for si in range(wb.Worksheets.Count):
        ws = wb.Worksheets[si]
        cells = ws.Cells
        names = {}
        wanted = set(specs.get('*', {}))
        for alias in (str(ws.Name), str(ws.Name).removeprefix('源_'), '源_' + str(ws.Name)):
            wanted.update(specs.get(alias, {}))
        for row in range(min(_HEADER_SCAN_ROWS, cells.MaxDataRow + 1)):
            for col in range(cells.MaxDataColumn + 1):
                name = cells[row, col].Value
                if col not in names and (_is_identifier_column(name) or _header_key(name) in wanted):
                    names[col] = (row, str(name))
        rules = _column_rules([name for _, name in names.values()], specs, str(ws.Name))
        for col, (header, name) in names.items():
            rule = rules.get(_header_key(name), {})
            if rule.get('type') != 'key':
                continue
            risky = 0
            for row in range(header + 1, cells.MaxDataRow + 1):
                cell = cells[row, col]
                if cell.IsFormula or cell.Type == CellValueType.IsNull:
                    continue
                original, is_date, convertible = _native_value(cell, CellValueType)
                if not convertible or is_date:
                    continue
                if isinstance(original, Real) and math.isfinite(original) and abs(original) >= 1e15:
                    risky += 1
                value = canon_value(original, 'key')
                if type(value) is not type(original) or value != original:
                    cell.PutValue(value)
                    changed += 1
                style = cell.GetStyle()
                fmt = _fmt_from_code(rule.get('format') or '@')
                if not _style_matches(style, fmt):
                    _set_fmt(style, fmt)
                    cell.SetStyle(style)
            if risky:
                logger.warning('[列契约] %s.%s 有 %s 个超过 15 位的数值主键，可能已丢失精度；请核对原始文本',
                               ws.Name, name, risky)
    return changed


def _canon_text(v):
    from backend.utils.data_helpers import _schema_text_value
    if isinstance(v, (datetime, date)):
        return v
    converted = _schema_text_value(v)
    return v if converted is None else converted


def _to_number(v):
    if not isinstance(v, str):
        return v
    s = _clean_text(v).replace(",", "").replace("，", "")
    if not s:
        return v
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]
    percent = s.endswith("%")
    if not percent and re.fullmatch(r'[+-]?[0-9]{16,}', s):
        return v
    try:
        exact = Decimal(s.rstrip("%"))
    except InvalidOperation:
        return v
    if not exact.is_finite():
        return v
    # Excel 数值仅可靠保留 15 位有效数字。长文本整数不得先转 float 再判断，
    # 包括用科学计数法表示的长编号；原字符串保留，避免不可逆的尾数损失。
    if not percent and exact == exact.to_integral_value() and exact.copy_abs() >= Decimal('1e15'):
        return v
    n = float(exact)
    if not math.isfinite(n):
        return v
    n = n / 100 if percent else n
    return -n if negative else n


def _dt_serial(dt, keep_time, date1904):
    if dt.tzinfo is not None:
        dt = dt.replace(tzinfo=None)
    delta = dt - _EXCEL_EPOCH
    serial = delta.days + ((delta.seconds + delta.microseconds / 1e6) / 86400.0 if keep_time else 0)
    return float(serial - _DATE1904_OFFSET if date1904 else serial)


def _to_serial(v, keep_time, parse, date1904):
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        if not math.isfinite(v) or v <= 0:
            return v
        if float(v).is_integer() and 19000101 <= v <= 99991231:
            try:  # YYYYMMDD 是业务日期，不是两千万天的序列号
                return _dt_serial(datetime.strptime(str(int(v)), "%Y%m%d"), keep_time, date1904)
            except ValueError:
                pass
        # 数值已是本工作簿日期系统的序列号，不再做 1904 偏移
        return float(v) if keep_time else float(math.floor(v))
    if isinstance(v, datetime):
        if v != v:  # NaT
            return v
        return _dt_serial(v, keep_time, date1904)
    if isinstance(v, date):
        return _dt_serial(datetime.combine(v, datetime.min.time()), keep_time, date1904)
    if isinstance(v, str):
        s = _clean_text(v)
        if not s:
            return v
        if parse:
            try:
                return _dt_serial(datetime.strptime(s, parse), keep_time, date1904)
            except ValueError:
                return v
        from backend.utils.source_sheet_writer import coerce_source_date
        parsed = coerce_source_date(s)
        if isinstance(parsed, datetime) and parsed == parsed:
            return _dt_serial(parsed, keep_time, date1904)
    return v


def canon_value(value, kind, parse=None, date1904=False):
    """唯一的值转换函数。转不了的值原样返回（不变成空，也不抛错）。

    key      → 规范文本：去不可见字符、全角转半角、整数去 .0（110.0 → "110"）
    text     → 文本（与 _schema_text_value 同规则）
    date     → 日期序列号，截掉时间部分（45296.354 → 45296.0）
    datetime → 日期时间序列号，保留时间
    number   → 文本数字转数值（"1,234.50" / "12%" / "(100)"）；长整数文本保留以免丢精度
    raw/None → 不转换
    """
    if kind == "key":
        return _canon_key(value)
    if kind == "text":
        return _canon_text(value)
    if kind in ("date", "datetime"):
        return _to_serial(value, kind == "datetime", parse, date1904)
    if kind == "number":
        return _to_number(value)
    return value


# ---------------------------------------------------------------- 格式工具

def _header_key(text):
    return re.sub(r"\s+", "", str(text or ""))


def _fmt_from_code(code, schema=None):
    """格式码 → (number_id, custom)；内置格式用编号，避免 Aspose 里写空 Custom 打回 General。"""
    if not code:
        return None
    code = str(code)
    if code.strip().lower() == "general":
        return (0, None)
    from openpyxl.styles.numbers import BUILTIN_FORMATS_REVERSE
    from backend.utils.column_formats import _builtin_format_id
    number_id = _builtin_format_id(schema, code) if schema is not None else BUILTIN_FORMATS_REVERSE.get(code)
    return (number_id, None) if number_id is not None else (None, code)


def _fmt_code(fmt):
    if not fmt:
        return ""
    if fmt[1]:
        return fmt[1]
    from openpyxl.styles.numbers import BUILTIN_FORMATS
    return BUILTIN_FORMATS.get(fmt[0] or 0, "General")


def _is_date_fmt(fmt):
    from openpyxl.styles.numbers import is_date_format
    code = _fmt_code(fmt)
    return bool(code) and is_date_format(code)


def _style_code(style):
    custom = str(style.Custom or "")
    return custom or _fmt_code((style.Number or 0, None))


def _style_matches(style, fmt):
    # 按格式码比较：Custom="General" 与内置 0、Custom="0.00" 与内置 2 视为相同
    return _style_code(style).strip().lower() == _fmt_code(fmt).strip().lower()


def _set_fmt(style, fmt):
    number, custom = fmt
    if custom:
        style.Custom = custom
    else:
        style.Number = number


def _locate_headers(cells, wanted):
    """在前 30 行按去空白后的文本找表头；返回 {col: (header_row, key)}，同名取首个。"""
    found, seen = {}, set()
    if not wanted:
        return found
    max_row = min(_HEADER_SCAN_ROWS, cells.MaxDataRow + 1)
    for ri in range(max_row):
        for ci in range(cells.MaxDataColumn + 1):
            value = cells[ri, ci].Value
            if value is None:
                continue
            key = _header_key(value)
            if key in wanted and key not in seen and ci not in found:
                found[ci] = (ri, key)
                seen.add(key)
    return found


# ---------------------------------------------------------------- 规则来源层
# 每层是 layer(cell, row, col, style, has_value) -> (kind, parse, fmt) 或 None。

def _spec_layer(cells, sheet_specs):
    columns = _locate_headers(cells, set(sheet_specs))
    if not columns:
        return None, []
    from backend.utils.output_postprocess import _scan_summary_rows
    first = min(hr for hr, _ in columns.values()) + 1
    summary = set(_scan_summary_rows(cells, cells.MaxDataRow, cells.MaxDataColumn, first))
    end = min(summary) if len(summary) == 1 else cells.MaxDataRow + 1
    rules = {ci: (hr, sheet_specs[key]) for ci, (hr, key) in columns.items()}

    def layer(cell, row, col, style, has_value):
        hit = rules.get(col)
        if not hit or row <= hit[0] or row >= end or not has_value or row in summary:
            return None
        spec = hit[1]
        fmt = _fmt_from_code(spec["format"])
        if fmt is None and spec['type'] in ('key', 'text'):
            fmt = (None, '@')
        if spec['type'] == 'raw' and fmt is None:
            fmt = (None, str(style.Custom)) if style.Custom else (style.Number or 0, None)
        return spec["type"], spec["parse"], fmt
    return layer, [spec for _, spec in rules.values()]


_SCHEMA_KINDS = {"text": "text", "date": "datetime", "datetime": "datetime",
                 "integer": "number", "decimal": "number"}


def _schema_layer(cells, schemas, with_template):
    """兼容旧行为：有模板时只给日期列定格式（到首个汇总行为止）；
    无模板时按 schema 转值并定格式。schema 的 date 不截时间（旧行为如此）。"""
    from backend.utils.column_formats import resolve_column_format, _builtin_format_id
    if with_template:
        schemas = {k: v for k, v in schemas.items()
                   if (v or {}).get("field_type") in ("date", "datetime", "time")}
    wanted = {_header_key(name): schema or {} for name, schema in schemas.items()}
    columns = _locate_headers(cells, set(wanted))
    if not columns:
        return None
    rules = {ci: (hr, wanted[key]) for ci, (hr, key) in columns.items()}
    end = cells.MaxDataRow + 1
    if with_template:
        from backend.utils.output_postprocess import _scan_summary_rows
        summary = _scan_summary_rows(cells, cells.MaxDataRow, cells.MaxDataColumn,
                                     min(hr for hr, _ in rules.values()) + 1)
        end = min(summary) if summary else end

    def layer(cell, row, col, style, has_value):
        hit = rules.get(col)
        if not hit or row <= hit[0] or row >= end or not has_value:
            return None
        schema = hit[1]
        if with_template:
            code = resolve_column_format(schema)
            return None, None, _fmt_from_code(code, schema)
        current = str(style.Custom or style.CultureCustom or "General")
        effective = dict(schema)
        effective["number_format"] = resolve_column_format(schema, current)
        if effective["number_format"] != schema.get("number_format"):
            effective["is_builtin_format"] = False
        code = resolve_column_format(effective)
        number_id = _builtin_format_id(effective, code)
        fmt = (number_id, None) if number_id is not None else _fmt_from_code(code, effective)
        return _SCHEMA_KINDS.get(schema.get("field_type")), None, fmt
    return layer


def _template_layer(cells, tcells):
    """本次模板同坐标规则；新增数据行继承模板数据区的列样板，模板外列留给兜底层。"""
    from backend.utils.output_postprocess import _shifted_summary_tail, _template_format_row
    from backend.utils.output_postprocess import _scan_summary_rows
    shifted = _shifted_summary_tail(cells, tcells)
    last_col = tcells.MaxDataColumn
    template_max_row = tcells.MaxDataRow
    formula_cells = {}
    policies = {}  # 模板不在本轮改动；同一模板坐标只读取、解析一次格式。
    summaries = _scan_summary_rows(tcells, template_max_row, last_col, 1)
    data_end = min(summaries) if summaries else template_max_row + 1
    samples = {}
    it = tcells.GetEnumerator()
    while it.MoveNext():
        tcell = it.Current
        if tcell.IsFormula:
            formula_cells.setdefault((tcell.Column, str(tcell.Formula).strip()), []).append(tcell)
        elif 0 < tcell.Row < data_end and not isinstance(tcell.Value, str):
            previous = samples.get(tcell.Column)
            if previous is None or previous.Row < tcell.Row:
                samples[tcell.Column] = tcell

    def layer(cell, row, col, style, has_value):
        if last_col is None or col > last_col:
            return None
        trow = _template_format_row(row, shifted, template_max_row)
        if trow is None or trow > template_max_row:
            # 不把数据样板刷到页脚/汇总区；仅用于新增数据行。
            if (shifted and row >= shifted[0]) or (not shifted and summaries):
                return None
            tcell = samples.get(col)
            if tcell is None and data_end > 1:
                # General 空样板格可能没有写入 XML，但仍是模板列的有效默认格式。
                tcell = tcells[data_end - 1, col]
            if tcell is None:
                return None
        else:
            tcell = tcells[trow, col]
        if cell.IsFormula and (not shifted or row < shifted[0]):
            candidates = formula_cells.get((col, str(cell.Formula).strip()), [])
            if len(candidates) == 1:
                tcell = candidates[0]
        coordinate = (tcell.Row, tcell.Column)
        if coordinate in policies:
            return policies[coordinate]
        tstyle = tcell.GetStyle()
        custom = str(tstyle.Custom or "")
        fmt = (None, custom) if custom else (tstyle.Number or 0, None)
        # 本次模板的 General 也是明确选择，不能让训练时的 date/schema 再把它覆盖。
        # 日期样式只控制显示，保留真实时间值；只有显式 type=date 才截掉时间。
        code = _fmt_code(fmt)
        # 四段格式是 正数;负数;零;文本。会计/货币格式的最后一段常含 @，
        # 它不表示数值应该转为文本。先移除字面量/占位间距，再只分析数值段。
        plain = re.sub(r'"[^"]*"|\\.|_.|\*.|\[[^\]]*\]', '', code)
        numeric_sections = ';'.join(plain.split(';')[:3])
        kind = ('datetime' if _is_date_fmt(fmt)
                else 'number' if re.search(r'[0#?]', numeric_sections)
                else 'raw' if re.search(r'general', numeric_sections, re.I)
                else 'text' if '@' in numeric_sections else 'raw')
        policies[coordinate] = (kind, None, fmt)
        return policies[coordinate]
    return layer


def _key_layer(cells):
    """默认主键文本策略，与 source/openpyxl 入口使用相同 canon_value。"""
    key_cols = {}
    for row in range(min(_HEADER_SCAN_ROWS, cells.MaxDataRow + 1)):
        for ci in range(cells.MaxDataColumn + 1):
            if ci not in key_cols and _is_identifier_column(cells[row, ci].Value):
                key_cols[ci] = row
    if not key_cols:
        return None

    def layer(cell, row, col, style, has_value):
        if col not in key_cols or row <= key_cols[col] or not has_value or cell.IsFormula:
            return None
        value = cell.Value
        if (isinstance(value, (int, float)) and not isinstance(value, bool)
              and math.isfinite(value) and abs(value) >= 1e15):
            layer.precision_risks += 1
        return "key", None, (None, "@")
    layer.precision_risks = 0
    return layer


def _inherited_layer(cells, default_code):
    from backend.utils.source_sheet_writer import is_date_keyword_column
    target = re.sub(r"\s+", "", default_code).lower()
    date_cols = set()
    for ci in range(cells.MaxDataColumn + 1):
        header = cells[0, ci].Value
        if header is not None and is_date_keyword_column(header):
            date_cols.add(ci)

    def layer(cell, row, col, style, has_value):
        custom = str(style.Custom or "")
        explicit = (style.Number or 0) != 0 or bool(custom.strip())
        if explicit and re.sub(r"\s+", "", custom).lower() != target:
            return None
        return None, None, ((None, "yyyy-mm-dd") if row > 0 and col in date_cols else (0, None))
    return layer


# ---------------------------------------------------------------- 源数据注入前归一

def _df_value(value, kind, parse):
    """DataFrame 侧的转换：与 canon_value 同规则，但日期保持 datetime（不转序列号），
    避免脚本里对日期列做 .dt / pd.to_datetime 的旧代码失效。"""
    import pandas as pd
    if value is None or value is pd.NaT or (isinstance(value, float) and math.isnan(value)):
        return value
    if kind in ("date", "datetime"):
        serial = canon_value(value, kind, parse)
        if isinstance(serial, float) and not isinstance(value, bool):
            from openpyxl.utils.datetime import from_excel
            try:
                return from_excel(serial)
            except (ValueError, OverflowError):
                return value
        return value
    return canon_value(value, kind, parse)


def apply_specs_to_source_data(source_data, specs):
    """按 _COLUMN_SPECS 的类型归一源数据 DataFrame；只动声明过类型的列。

    不修改传入对象：预加载缓存会跨训练轮次复用，声明改了之后不能残留上一轮的转换。
    有改动的条目复制 DataFrame，其余条目原样共享。
    源数据 key 是"文件名_sheet"，声明里的 sheet 名可以写 key 本身、"源_"+key 或 "*"。
    """
    if not specs or not isinstance(source_data, dict):
        return source_data
    result, copies, changed = {}, {}, 0
    for key, entry in source_data.items():
        if id(entry) in copies:  # 同一条目被多个 key 引用时保持共享
            result[key] = copies[id(entry)]
            continue
        df = entry.get("df") if isinstance(entry, dict) else None
        sheet_specs = {**specs.get("*", {}), **specs.get(str(key), {}), **specs.get(f"源_{key}", {})}
        targets = [] if df is None else [
            (column, sheet_specs[_header_key(column)]) for column in df.columns
            if sheet_specs.get(_header_key(column), {}).get("type") not in (None, "raw")]
        if targets:
            df = df.copy()
            for column, spec in targets:
                kind, parse = spec["type"], spec["parse"]
                df[column] = df[column].astype(object).map(lambda v: _df_value(v, kind, parse))
            entry = {**entry, "df": df}
            changed += len(targets)
        copies[id(source_data[key])] = entry
        result[key] = entry
    if changed:
        logger.info(f"[列契约] 源数据注入前按脚本声明归一了 {changed} 列")
    return result


# ---------------------------------------------------------------- 唯一写入点

def _native_value(cell, value_types):
    """Aspose 单元格 → (Python 值, 是否日期型, 可否转换)。日期按序列号取，不经 .NET DateTime。"""
    kind = cell.Type
    if kind == value_types.IsNull:
        return None, False, False
    if kind == value_types.IsNumeric:
        return float(cell.DoubleValue), False, True
    if kind == value_types.IsDateTime:
        return float(cell.DoubleValue), True, True
    if kind == value_types.IsString:
        return str(cell.Value), False, True
    return None, False, False


def apply_column_contract(wb, output_path=None, script_code=None, template_path=None,
                          expected_structure=None, explicit_formats=None, legacy_layers=True):
    """按列契约逐格裁决值类型和格式，一次写入已打开的 Aspose 工作簿（不保存、不重算）。

    legacy_layers=False（v2 管线）时使用 explicit + spec + identity + 本次模板。
    返回 {"values": n, "formats": n, "by_layer": {...}, "failed": n}。
    """
    import json
    from Aspose.Cells import CellValueType
    from backend.utils.output_postprocess import _open_workbook, _template_default_format

    specs = read_column_specs(script_code)
    structure = expected_structure
    if isinstance(structure, str):
        try:
            structure = json.loads(structure)
        except ValueError:
            structure = None
    structure_sheets = (structure or {}).get("sheets", {}) if isinstance(structure, dict) else {}
    template_ok = bool(template_path and os.path.exists(template_path))
    date1904 = bool(getattr(wb.Settings, "Date1904", False))

    default_code = None
    if legacy_layers:
        source = template_path if template_ok else output_path
        default = _template_default_format(source) if source else None
        default_code = default[0] if default else None

    tpl = _open_workbook(template_path) if template_ok else None
    tpl_sheets = ({str(tpl.Worksheets[i].Name): tpl.Worksheets[i] for i in range(tpl.Worksheets.Count)}
                  if tpl is not None else {})
    stats = Counter()
    failed = 0
    try:
        for si in range(wb.Worksheets.Count):
            ws = wb.Worksheets[si]
            name = str(ws.Name)
            cells = ws.Cells
            if cells.MaxDataRow < 0 or cells.MaxDataColumn < 0:
                continue
            identity_layer = _key_layer(cells)
            sheet_specs = _sheet_specs(specs, name)
            if not (legacy_layers or sheet_specs or name in tpl_sheets
                    or identity_layer is not None):
                continue
            explicit = set((explicit_formats or {}).get(name, {}))
            layers = []
            spec_layer, used_specs = _spec_layer(cells, sheet_specs) if sheet_specs else (None, [])
            for spec in used_specs:
                logger.info(f"[列契约] {name}.{spec['name']}: 类型={spec['type'] or '-'} "
                            f"格式={spec['format'] or '-'} 解析={spec['parse'] or '-'}（来源: 脚本声明）")
            layers.append(("spec", spec_layer))
            layers.append(("identity", identity_layer))
            tws = tpl_sheets.get(name)
            if tws is not None:
                layers.append(("template", _template_layer(cells, tws.Cells)))
            if legacy_layers:
                schemas = (structure_sheets.get(name) or {}).get("column_schemas") or {}
                if schemas and (tws is not None or not template_ok):
                    layers.append(("schema", _schema_layer(cells, schemas, template_ok)))
                if default_code and name.startswith("源_"):
                    layers.append(("inherited", _inherited_layer(cells, default_code)))
            layers = [(label, fn) for label, fn in layers if fn is not None]
            if not layers:
                continue

            changes = []
            it = cells.GetEnumerator()
            while it.MoveNext():
                cell = it.Current
                explicit_cell = str(cell.Name) in explicit
                row, col = cell.Row, cell.Column
                is_formula = cell.IsFormula
                has_value = is_formula or cell.Type != CellValueType.IsNull
                style = cell.GetStyle()
                kind = parse = fmt = None
                kind_src = fmt_src = None
                for label, layer in layers:
                    decision = layer(cell, row, col, style, has_value)
                    if not decision:
                        continue
                    if kind_src is None and decision[0] is not None:
                        kind, parse, kind_src = decision[0], decision[1], label
                    if fmt_src is None and decision[2] is not None:
                        fmt, fmt_src = decision[2], label
                    if kind_src and fmt_src:
                        break
                # 显式显示格式不能让主键跳过值归一，否则查找列与查找值仍类型不一致。
                # 其他显式格式沿用旧优先级，避免改变日期/金额的既有行为。
                if explicit_cell:
                    if kind != 'key':
                        continue
                    fmt = None
                # 用户声明了日期类型但没给日期格式时，不让下层的常规格式把序列号露出来
                if kind_src == "spec" and kind in _DEFAULT_DATE_FORMATS and (fmt is None or not _is_date_fmt(fmt)):
                    fmt, fmt_src = (None, _DEFAULT_DATE_FORMATS[kind]), "spec"
                new_value = None
                if kind not in (None, "raw") and not is_formula:
                    value, is_date, convertible = _native_value(cell, CellValueType)
                    original_value = value
                    if is_date and kind == 'text' and kind_src == 'template':
                        # 切换为文本模板时保留当前可读日期，不能把日期序列号当文本写出。
                        value, is_date = str(cell.StringValue), False
                    if convertible and not (is_date and kind in ("key", "text")):
                        converted = canon_value(value, kind, parse, date1904)
                        if type(converted) is not type(original_value) or converted != original_value:
                            new_value = converted
                if new_value is not None or (fmt is not None and not _style_matches(style, fmt)):
                    changes.append((cell, new_value, fmt, kind_src, fmt_src))

            for cell, new_value, fmt, kind_src, fmt_src in changes:
                if new_value is not None:
                    cell.PutValue(new_value)
                    stats[f"{kind_src}:值"] += 1
                if fmt is not None:
                    style = cell.GetStyle()
                    if not _style_matches(style, fmt):
                        _set_fmt(style, fmt)
                        cell.SetStyle(style)
                        if not _style_matches(cell.GetStyle(), fmt):
                            failed += 1
                        stats[f"{fmt_src}:格式"] += 1
            if identity_layer is not None and identity_layer.precision_risks:
                count = identity_layer.precision_risks
                stats['identity:precision_risk'] += count
                logger.warning(
                    "[列契约] %s 有 %s 个主键/ID 已以超过 15 位的数值存储，"
                    "可能已丢失精度；转成文本不能恢复尾数，请核对原始文本编号", name, count)

        # 第 1 层：脚本显式单元格格式，复用原有实现（上面的循环已跳过这些单元格）
        from backend.utils.script_formats import apply_explicit_formats
        apply_explicit_formats(wb, explicit_formats or {})
        stats["explicit:格式"] += sum(len(v) for v in (explicit_formats or {}).values())
    finally:
        if tpl is not None:
            tpl.Dispose()  # 后处理会话里借出的工作簿 Dispose 是空操作，由会话统一释放

    values = sum(n for k, n in stats.items() if k.endswith(":值"))
    formats = sum(n for k, n in stats.items() if k.endswith(":格式"))
    if values or formats:
        logger.info(f"[列契约] 改值 {values} 格、改格式 {formats} 格，按来源: {dict(stats)}")
    if failed:
        logger.warning(f"[列契约] {failed} 个单元格格式写入后校验不一致，请检查格式码是否有效")
    return {"values": values, "formats": formats, "by_layer": dict(stats), "failed": failed,
            "precision_risks": stats['identity:precision_risk']}
