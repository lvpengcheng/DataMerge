"""One non-executing column-declaration decoder for generation, editing and runtime."""
import ast
import re

KINDS = ('key', 'text', 'date', 'datetime', 'number', 'raw')
TYPE_ALIASES = {
    'str': 'text', 'string': 'text', '文本': 'text', '字符串': 'text',
    'int': 'number', 'integer': 'number', 'float': 'number', 'double': 'number',
    'decimal': 'number', 'numeric': 'number', '数值': 'number', '数字': 'number',
    '日期': 'date', '日期时间': 'datetime', 'timestamp': 'datetime',
    'identifier': 'key', '主键': 'key', '原样': 'raw',
}
RULE_FIELDS = {'type', 'format', 'parse'}
FIELD_ALIASES = {'number_format': 'format', 'field_type': 'type', 'dtype': 'type'}
SHEET_METADATA = {'source', 'source_file', 'file', 'file_name', 'sheet', 'sheet_name',
                  'description', 'role', 'kind'}

DECLARATION_CONTRACT = '''
_COLUMN_SPECS 唯一标准结构是 {sheet名: {真实列名: {"type": 类型, "format": Excel格式码, "parse": 可选输入日期格式}}}。
示例：_COLUMN_SPECS = {"源_员工档案": {"工号": {"type": "key"}}, "结果": {"工资月份": {"type": "date", "format": "yyyy-m-d"}}}
type 只能为 key/text/date/datetime/number/raw；source/target 是来源角色，绝不是 type。
不要把 source/file/columns 等结构元信息混入列字典。多文件分别声明，不能按文件名再包装子脚本的列声明。
'''


def literal_column_specs(node):
    """Read literals only and reject duplicate dictionary keys before Python loses them."""
    for item in ast.walk(node):
        if not isinstance(item, ast.Dict):
            continue
        seen = set()
        for key in item.keys:
            value = ast.literal_eval(key)
            if value in seen:
                raise ValueError(f'_COLUMN_SPECS 含重复字典键 {value!r}，不能依赖后值覆盖前值')
            seen.add(value)
    return ast.literal_eval(node)


def header_key(value):
    return re.sub(r'\s+', '', str(value or ''))


def _error(path, message):
    raise ValueError(f'列声明 {path}: {message}')


def _kind(value, path):
    if not isinstance(value, str):
        _error(path, f'type 必须是字符串，可用 {"/".join(KINDS)}')
    result = value.strip().lower()
    result = TYPE_ALIASES.get(result, result)
    if result not in KINDS:
        _error(path, f'type {value!r} 不支持；可用 {"/".join(KINDS)}。source/target 属于来源元信息，不是值类型')
    return result


def _looks_rule(value):
    if isinstance(value, str):
        return value.strip().lower() in {*KINDS, *TYPE_ALIASES}
    policy = RULE_FIELDS | set(FIELD_ALIASES)
    return (isinstance(value, dict) and bool(set(value) & policy)
            and set(value) <= policy | {'name', 'source', 'source_file', 'description', 'role'}
            and all(item is None or isinstance(item, str) for key, item in value.items() if key in policy))


def _rule(value, path, column, allow_internal):
    if isinstance(value, str):
        return _kind(value, path)
    if not isinstance(value, dict) or not value:
        _error(path, '声明必须是类型字符串或含 type/format 的字典')
    result = {}
    internal = allow_internal and 'name' in value and header_key(value['name']) == header_key(column)
    for key, item in value.items():
        if key == 'name' and isinstance(item, str) and header_key(item) == header_key(column):
            continue
        if key in {'source', 'source_file', 'description', 'role'} and set(value) & (RULE_FIELDS | set(FIELD_ALIASES)):
            continue
        field = FIELD_ALIASES.get(key, key)
        if field not in RULE_FIELDS:
            _error(path, f'未知规则字段 {key!r}；仅支持 type/format/parse')
        if item is None and internal:
            continue
        if field == 'type':
            item = _kind(item, path)
        elif not isinstance(item, str) or not item.strip():
            _error(path, f'{field} 必须是非空字符串')
        if field in result and result[field] != item:
            _error(path, f'{field} 存在冲突声明')
        result[field] = item
    if not (result.get('type') or result.get('format')):
        _error(path, '缺少 type 或 format')
    return result


def decode_column_specs(value, workbook_name=None, allow_internal=False):
    """Normalize only structurally unambiguous forms; never infer business columns."""
    result = {}

    def add(sheet, column, rule):
        if not isinstance(sheet, str) or not sheet.strip():
            _error('_COLUMN_SPECS', '列声明需要非空 sheet 名及列字典；例如 {"*": {"入职日期": {"type": "date"}}}')
        if not isinstance(column, str) or not header_key(column):
            _error(sheet, '列名必须是非空字符串')
        sheet = sheet.strip()
        normalized = _rule(rule, f'{sheet}.{column}', column, allow_internal)
        columns = result.setdefault(sheet, {})
        name = next((key for key in columns if header_key(key) == header_key(column)), column)
        if name in columns:
            old = columns[name]
            old = {'type': old} if isinstance(old, str) else old
            new = {'type': normalized} if isinstance(normalized, str) else normalized
            if any(key in old and old[key] != item for key, item in new.items()):
                _error(f'{sheet}.{column}', '存在冲突的重复声明')
            normalized = {**old, **new}
        columns[name] = normalized

    def records(items, default_sheet=None):
        for index, record in enumerate(items, 1):
            if not isinstance(record, dict):
                _error(f'记录[{index}]', '需要含 sheet、column 的字典')
            row = dict(record)
            sheet = row.pop('sheet', row.pop('sheet_name', default_sheet))
            if 'columns' in row:
                columns = row.pop('columns')
                unknown = set(row) - SHEET_METADATA
                if unknown:
                    _error(f'记录[{index}]', f'未知 sheet 元信息 {sorted(unknown)}')
                sheet_columns(sheet, columns)
                continue
            column = row.pop('column', row.pop('column_name', row.pop('name', None)))
            add(sheet, column, row)

    def sheet_columns(sheet, columns):
        if not isinstance(sheet, str) or not sheet.strip():
            _error('_COLUMN_SPECS', '列声明需要非空 sheet 名及列字典；例如 {"*": {"入职日期": {"type": "date"}}}')
        if isinstance(columns, list):
            records(columns, sheet)
            return
        if not isinstance(columns, dict):
            _error(sheet, '需要列字典或列记录列表')
        if 'columns' in columns and not _looks_rule(columns['columns']):
            unknown = set(columns) - SHEET_METADATA - {'columns', 'type'}
            if unknown:
                _error(sheet, f'未知 sheet 元信息 {sorted(unknown)}')
            sheet_columns(sheet, columns['columns'])
            return
        usable = {key: item for key, item in columns.items()
                  if not (key in SHEET_METADATA and isinstance(item, str) and not _looks_rule(item))
                  and not (key == 'type' and item in ('source', 'target', 'output')
                           and isinstance(columns.get('source'), str) and not _looks_rule(columns['source']))}
        if columns and not usable:
            _error(sheet, '只有来源元信息，没有列规则；请使用 {"columns": {真实列名: 声明}}')
        result.setdefault(sheet.strip(), {})
        for column, rule in usable.items():
            add(sheet, column, rule)

    def root(specs):
        if specs is None:
            return
        if isinstance(specs, list):
            records(specs)
            return
        if not isinstance(specs, dict):
            _error('_COLUMN_SPECS', '必须是字典或含 sheet/column 的列表')
        if set(specs) == {'_COLUMN_SPECS'}:
            root(specs['_COLUMN_SPECS'])
            return
        if isinstance(specs.get('sheet', specs.get('sheet_name')), str) and 'columns' in specs:
            records([specs])
            return
        files = [name for name, item in specs.items() if isinstance(name, str)
                 and name.lower().endswith(('.xlsx', '.xlsm', '.xls'))
                 and (isinstance(item, list) or isinstance(item, dict) and any(not _looks_rule(v) for v in item.values()))]
        if files:
            if not workbook_name or set(files) != set(specs) or workbook_name not in specs:
                _error('_COLUMN_SPECS', f'无法对应当前交付文件 {workbook_name!r}；请返回当前文件的 sheet 字典')
            root(specs[workbook_name])
            return
        if specs and all(_looks_rule(item) for item in specs.values()):
            sheet_columns('*', specs)
            return
        for sheet, columns in specs.items():
            sheet_columns(sheet, columns)

    root(value)
    return result
