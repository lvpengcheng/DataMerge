"""统一列契约：输出工作簿"值类型 + 显示格式"的唯一裁决点与唯一写入点。

以前格式由 6 个后处理函数各自判断、按执行顺序覆盖（主键转文本、模板格式还原、
源_ 默认样式兜底、expected schema、汇总区重刷、脚本显式格式），谁最后跑谁说了算。
现在它们都退化为"规则来源层"，只提供规则，不写单元格；由 apply_column_contract
逐格裁决后一次写入。

每个单元格分别裁决两件事，按优先级取第一个给出答案的层：
  - kind（值类型，决定 VLOOKUP 能否匹配）：key / text / date / datetime / number / raw
  - format（number_format，决定显示）

层次（高 → 低）：
  1. explicit  脚本里显式写的单元格 number_format（值不动）
  2. spec      脚本 _COLUMN_SPECS 常量（智训"执行修改"里用户指定的列类型/格式）
  3. template  本次实际使用的模板类型/格式（汇总区位移时按位移后的行号）
  4. schema    训练推断的 expected_structure column_schemas（模板未覆盖的位置兜底）
  5. key       列名命中主键关键词 → 规范文本（兼容旧 normalize_key_columns_to_text）
  6. inherited 源_ sheet 继承模板非 General 默认样式的单元格 → General/日期
  都没有 → 保持原样。v2 管线只启用 explicit + spec + 本次模板，不使用训练推断兜底。

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
from collections import Counter
from datetime import date, datetime

logger = logging.getLogger(__name__)

KINDS = ("key", "text", "date", "datetime", "number", "raw")
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
                specs = ast.literal_eval(node.value)
            except (ValueError, TypeError) as e:
                logger.warning(f"[列契约] _COLUMN_SPECS 不是纯字面量，忽略: {e}")
    return normalize_column_specs(specs)


def normalize_column_specs(specs):
    """{sheet: {列: "kind" | {"type", "format", "parse"}}} → 统一成 dict 形式，丢弃无效项。"""
    result = {}
    if not isinstance(specs, dict):
        return result
    for sheet, columns in specs.items():
        if not isinstance(columns, dict):
            continue
        for column, spec in columns.items():
            if isinstance(spec, str):
                spec = {"type": spec}
            if not isinstance(spec, dict):
                continue
            kind = spec.get("type")
            if kind is not None and kind not in KINDS:
                logger.warning(f"[列契约] {sheet}.{column} 的类型 {kind!r} 不支持，仅保留格式")
                kind = None
            fmt = spec.get("format") or None
            if kind is None and not fmt:
                continue
            result.setdefault(str(sheet).strip(), {})[_header_key(column)] = {
                "type": kind, "format": fmt, "parse": spec.get("parse") or None,
                "name": str(column).strip()}
    return result


# ---------------------------------------------------------------- 值转换

def _clean_text(s):
    return s.translate(_TEXT_TABLE).strip()


def _canon_key(v):
    if v is None or isinstance(v, (bool, datetime, date)):
        return v
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if not math.isfinite(v):
            return v
        return str(int(v)) if v.is_integer() else format(v, ".15g")
    if isinstance(v, str):
        return _clean_text(v) or v
    return v


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
    try:
        n = float(s.rstrip("%"))
    except ValueError:
        return v
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
    number   → 文本数字转数值（"1,234.50" / "12%" / "(100)"）
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
    rules = {ci: (hr, sheet_specs[key]) for ci, (hr, key) in columns.items()}

    def layer(cell, row, col, style, has_value):
        hit = rules.get(col)
        if not hit or row <= hit[0] or not has_value or row in summary:
            return None
        spec = hit[1]
        return spec["type"], spec["parse"], _fmt_from_code(spec["format"])
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
        tstyle = tcell.GetStyle()
        custom = str(tstyle.Custom or "")
        fmt = (None, custom) if custom else (tstyle.Number or 0, None)
        # 本次模板的 General 也是明确选择，不能让训练时的 date/schema 再把它覆盖。
        # 日期样式只控制显示，保留真实时间值；只有显式 type=date 才截掉时间。
        code = _fmt_code(fmt)
        plain = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', '', code)
        kind = ('datetime' if _is_date_fmt(fmt) else 'text' if '@' in plain
                else 'number' if re.search(r'[0#?]', plain) else 'raw')
        return kind, None, fmt
    return layer


def _key_layer(cells):
    from backend.utils.output_postprocess import _is_key_column
    key_cols = {ci for ci in range(cells.MaxDataColumn + 1) if _is_key_column(cells[0, ci].Value)}
    if not key_cols:
        return None

    def layer(cell, row, col, style, has_value):
        if col not in key_cols or row < 1 or not has_value or cell.IsFormula:
            return None
        return "key", None, (None, "@")
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

    legacy_layers=False（v2 管线）时仅使用 explicit + spec + 本次模板。
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
            if not (legacy_layers or specs.get("*") or specs.get(name) or name in tpl_sheets):
                continue
            explicit = set((explicit_formats or {}).get(name, {}))
            layers = []
            sheet_specs = {**specs.get("*", {}), **specs.get(name, {})}
            spec_layer, used_specs = _spec_layer(cells, sheet_specs) if sheet_specs else (None, [])
            for spec in used_specs:
                logger.info(f"[列契约] {name}.{spec['name']}: 类型={spec['type'] or '-'} "
                            f"格式={spec['format'] or '-'} 解析={spec['parse'] or '-'}（来源: 脚本声明）")
            layers.append(("spec", spec_layer))
            tws = tpl_sheets.get(name)
            if tws is not None:
                layers.append(("template", _template_layer(cells, tws.Cells)))
            if legacy_layers:
                schemas = (structure_sheets.get(name) or {}).get("column_schemas") or {}
                if schemas and (tws is not None or not template_ok):
                    layers.append(("schema", _schema_layer(cells, schemas, template_ok)))
                layers.append(("key", _key_layer(cells)))
                if default_code and name.startswith("源_"):
                    layers.append(("inherited", _inherited_layer(cells, default_code)))
            layers = [(label, fn) for label, fn in layers if fn is not None]
            if not layers:
                continue

            changes = []
            it = cells.GetEnumerator()
            while it.MoveNext():
                cell = it.Current
                if str(cell.Name) in explicit:
                    continue
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
    return {"values": values, "formats": formats, "by_layer": dict(stats), "failed": failed}
