"""Split an integrated workbook by row groups without rebuilding its sheets."""
import ast
import hashlib
import math
import operator
import re
import zipfile
from pathlib import Path


FUNCTIONS = {'IF', 'AND', 'OR', 'NOT', 'CONTAINS', 'STARTSWITH', 'ENDSWITH'}
MAX_GROUPS = 200


def parse_expression(expression, columns):
    """Replace exact column tokens outside quoted literals and validate a small AST."""
    text = str(expression or '').strip().lstrip('=').strip()
    if not text or len(text) > 4000:
        raise ValueError('拆分表达式不能为空且不能超过4000个字符')
    refs = {}
    for i, col in enumerate(columns):
        refs[f'[{col}]'] = f'__col_{i}'
        if str(col).isidentifier() and str(col).upper() not in FUNCTIONS:
            refs[str(col)] = f'__col_{i}'
    pattern = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\'\\]|\\.)*\'|' +
                         '|'.join(re.escape(s) for s in sorted(refs, key=len, reverse=True)))
    used = set()
    def replace(match):
        token = match.group()
        if token in refs:
            used.add(int(refs[token][6:]))
            return refs[token]
        # Excel escapes a literal quote using two adjacent double quotes.
        return repr(token[1:-1].replace('""', '"')) if token.startswith('"') else token
    text = pattern.sub(replace, text)
    # Preserve comparison characters inside strings.
    chunks = re.split(r"('(?:[^'\\]|\\.)*'|\"(?:[^\"\\]|\\.)*\")", text)
    for i in range(0, len(chunks), 2):
        part = chunks[i].replace('<>', '!=')
        part = re.sub(r'(?<![<>=!])=(?!=)', '==', part)
        for name in FUNCTIONS:
            part = re.sub(rf'\b{name}\s*(?=\()', name, part, flags=re.I)
        chunks[i] = part
    try:
        root = ast.parse(''.join(chunks), mode='eval').body
    except (SyntaxError, RecursionError) as exc:
        raise ValueError('拆分表达式语法不正确，请使用 IF(工作地="上海","上海","其他") 等写法') from exc
    allowed = (ast.Expression, ast.Constant, ast.Name, ast.Load, ast.Call, ast.Compare,
               ast.Eq, ast.NotEq, ast.Gt, ast.GtE, ast.Lt, ast.LtE,
               ast.BinOp, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.UnaryOp, ast.USub, ast.UAdd)
    nodes = list(ast.walk(root))
    if len(nodes) > 500:
        raise ValueError('拆分表达式过于复杂，请缩短条件')
    for node in nodes:
        if not isinstance(node, allowed):
            raise ValueError('拆分表达式含不支持的语法')
        if isinstance(node, ast.Name) and node.id not in FUNCTIONS and not re.fullmatch(r'__col_\d+', node.id):
            raise ValueError(f'拆分表达式中不存在的列或函数：{node.id}')
        if isinstance(node, ast.Constant) and isinstance(node.value, (float, int)) and not isinstance(node.value, bool):
            if abs(node.value) > 1e15 or not math.isfinite(node.value):
                raise ValueError('拆分表达式中的数字过大')
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in FUNCTIONS or node.keywords:
                raise ValueError('拆分表达式含不支持的函数')
            n = len(node.args)
            if ((node.func.id == 'IF' and n != 3) or (node.func.id == 'NOT' and n != 1)
                    or (node.func.id in {'AND', 'OR'} and not n)
                    or (node.func.id in {'CONTAINS', 'STARTSWITH', 'ENDSWITH'} and n != 2)):
                raise ValueError(f'{node.func.id} 的参数数量不正确')
    if not used:
        raise ValueError('拆分表达式至少需要引用一个模板字段')
    return root, [list(columns)[i] for i in sorted(used)]


def evaluate(root, values):
    def ev(n):
        if isinstance(n, ast.Constant):
            return n.value
        if isinstance(n, ast.Name):
            return values[n.id]
        if isinstance(n, ast.Compare):
            left = ev(n.left)
            ops = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Gt: operator.gt,
                   ast.GtE: operator.ge, ast.Lt: operator.lt, ast.LtE: operator.le}
            for op, rhs in zip(n.ops, n.comparators):
                right = ev(rhs)
                if not ops[type(op)](left, right):
                    return False
                left = right
            return True
        if isinstance(n, ast.BinOp):
            ops = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv}
            left, right = ev(n.left), ev(n.right)
            if not isinstance(left, (int, float)) or not isinstance(right, (int, float)):
                raise ValueError('四则运算只能用于数值字段，文本请使用条件函数')
            return ops[type(n.op)](left, right)
        if isinstance(n, ast.UnaryOp):
            return -ev(n.operand) if isinstance(n.op, ast.USub) else +ev(n.operand)
        if isinstance(n, ast.Call):
            name = n.func.id
            if name == 'IF':
                return ev(n.args[1] if ev(n.args[0]) else n.args[2])
            if name == 'AND':
                return all(ev(a) for a in n.args)
            if name == 'OR':
                return any(ev(a) for a in n.args)
            if name == 'NOT':
                return not ev(n.args[0])
            a, b = (str(ev(arg) or '') for arg in n.args)
            return b in a if name == 'CONTAINS' else (a.startswith(b) if name == 'STARTSWITH' else a.endswith(b))
        raise ValueError('无法计算拆分表达式')
    return ev(root)


def validate_split_config(config, columns=None):
    if not isinstance(config, dict) or config.get('mode') not in {'value', 'expression'}:
        raise ValueError('请选择按字段值或按表达式拆分')
    if 'include_source_keys' in config and not isinstance(config['include_source_keys'], bool):
        raise ValueError('纳入数据表新增记录必须为开关值')
    if config['mode'] == 'value':
        col = config.get('column')
        if not isinstance(col, str) or not col.strip() or (columns is not None and col not in columns):
            raise ValueError('请选择有效的拆分字段')
    else:
        if columns is not None:
            parse_expression(config.get('expression'), columns)
        elif not isinstance(config.get('expression'), str) or not 0 < len(config['expression']) <= 4000:
            raise ValueError('请填写有效的拆分表达式')
    start, end = config.get('data_row_start'), config.get('data_row_end')
    if start is not None and (type(start) is not int or not 1 <= start <= 1048576):
        raise ValueError('数据起始行必须是正整数')
    if end is not None and (type(end) is not int or not 1 <= end <= 1048576):
        raise ValueError('数据结束行必须是正整数')
    if start is not None and end is not None and end < start:
        raise ValueError('数据结束行不能小于起始行')
    return dict(config)


def group_rows(rows, columns, config):
    """rows = [(absolute 1-based row number, {column: cached value}), ...]."""
    validate_split_config(config, columns)
    root = parse_expression(config['expression'], columns)[0] if config['mode'] == 'expression' else None
    groups = {}
    for row, data in rows:
        try:
            label = evaluate(root, {f'__col_{i}': data.get(c, '') for i, c in enumerate(columns)}) if root else data.get(config['column'])
        except (TypeError, ValueError, KeyError, ArithmeticError) as exc:
            raise ValueError(f'第{row}行拆分表达式计算失败：{exc}') from exc
        if label is None or (isinstance(label, float) and math.isnan(label)) or str(label).strip() == '':
            label = '空白'
        elif isinstance(label, bool):
            label = '满足条件' if label else '不满足条件'
        elif isinstance(label, float) and label.is_integer():
            label = str(int(label))
        else:
            label = str(label).strip()
        groups.setdefault(label, []).append(row)
        if len(groups) > MAX_GROUPS:
            raise ValueError(f'拆分结果超过{MAX_GROUPS}组，请缩小数据范围或调整表达式')
    return groups


def deletion_ranges(start, end, keep):
    """Return contiguous zero-based row ranges, bottom first for stable deletion."""
    keep = set(keep)
    ranges = []
    for row in range(start, end + 1):
        if row in keep:
            continue
        if ranges and ranges[-1][0] + ranges[-1][1] == row - 1:
            ranges[-1][1] += 1
        else:
            ranges.append([row - 1, 1])
    return list(reversed(ranges))


def output_name(label, used, main_name=None):
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', label).strip(' .') or '空白'
    safe = safe.encode('utf-8')[:140].decode('utf-8', 'ignore')
    if main_name:
        stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', Path(main_name).stem).strip(' .') or '主表'
        stem = stem.encode('utf-8')[:90].decode('utf-8', 'ignore')
        base = safe + '_' + stem
    else:
        base = '拆表_' + safe
    name = base + '.xlsx'
    if name.casefold() in used:
        name = base + '_' + hashlib.sha256(label.encode()).hexdigest()[:10] + '.xlsx'
    used.add(name.casefold())
    return name


def split_workbook(path, output_dir, sheet_name, head_data, start, end, config, main_name=None):
    """Open the original workbook for each group; delete only unwanted data rows.

    Aspose's updateReference overload updates formulas on all sheets, merged cells,
    named ranges and footer positions while preserving the surrounding workbook.
    """
    import aspose_init
    from Aspose.Cells import Workbook, SaveFormat
    from openpyxl.utils import column_index_from_string
    aspose_init.ensure_license()
    columns = list(head_data)
    validate_split_config(config, columns)
    wb = Workbook(str(path))
    try:
        ws = wb.Worksheets[sheet_name]
        if start < 1 or end < start or end > int(ws.Cells.MaxDataRow) + 1:
            raise ValueError('数据行范围无效，请确认明细数据起止行，排除表头和尾部汇总')
        rows = []
        for row in range(start, end + 1):
            data = {}
            for col, letter in head_data.items():
                value = ws.Cells[row - 1, column_index_from_string(letter) - 1].Value
                # pythonnet numeric/string objects become standard values for comparisons.
                if value is not None and type(value).__module__ not in {'builtins', 'datetime'}:
                    try:
                        value = float(value) if str(value.GetType().FullName).startswith('System.') and str(value.GetType().Name) not in {'String', 'DateTime', 'Boolean'} else str(value)
                    except Exception:
                        value = str(value)
                data[col] = value
            if any(value is not None and str(value).strip() != '' for value in data.values()):
                rows.append((row, data))
        groups = group_rows(rows, columns, config)
    finally:
        wb.Dispose()
    if not groups:
        raise ValueError('数据范围内没有可拆分的记录')
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    used, outputs = set(), []
    for label, keep in groups.items():
        wb = Workbook(str(path))
        try:
            cells = wb.Worksheets[sheet_name].Cells
            for first, count in deletion_ranges(start, end, keep):
                cells.DeleteRows(first, count, True)
            # Never flatten formulas; update cached values after reference adjustment.
            wb.CalculateFormula()
            wb.Settings.ForceFullCalculate = True
            name = output_name(label, used, main_name or Path(path).name)
            wb.Save(str(output_dir / name), SaveFormat.Xlsx)
            outputs.append({'name': name, 'group': label, 'rows': len(keep)})
        finally:
            wb.Dispose()
    archive = output_dir / '拆表结果.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as zf:
        for item in outputs:
            zf.write(output_dir / item['name'], item['name'])
    return {'outputs': outputs, 'group_count': len(outputs), 'split_rows': len(rows), 'archive': str(archive)}
