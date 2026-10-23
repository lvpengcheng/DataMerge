"""Explicit key-set selection shared by integration and splitting; no eval."""
import ast
import copy

from .integrate_engine import normalize_key
from .table_split import parse_expression, evaluate


def scope_tree(config):
    if not config or config.get('mode', 'all') == 'all':
        return None
    mode = config.get('mode')
    if mode == 'source':
        return {'file': config.get('source_file'), 'condition': config.get('filter_expression') or ''}
    if mode != 'expression':
        raise ValueError('数据范围请选择全部并集、指定表或表达式')
    text = str(config.get('expression') or '').strip().lstrip('=').strip()
    if not text or len(text) > 4000:
        raise ValueError('数据范围表达式不能为空且不能超过4000字')
    try:
        root = ast.parse(text, mode='eval').body
    except (SyntaxError, RecursionError) as exc:
        raise ValueError('数据范围表达式语法错误，例如 INTERSECT("B.xlsx","C.xlsx")') from exc
    if len(list(ast.walk(root))) > 200:
        raise ValueError('数据范围表达式过于复杂')

    def parse(node):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            return {'file': node.value, 'condition': ''}
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name) or node.keywords:
            raise ValueError('数据范围仅支持表名及 UNION、INTERSECT、EXCEPT、FILTER')
        op = node.func.id.upper()
        if op == 'FILTER':
            if len(node.args) != 2 or not all(isinstance(n, ast.Constant) and isinstance(n.value, str) for n in node.args):
                raise ValueError('FILTER用法：FILTER("B.xlsx", "[状态]=\'在职\'")')
            if not node.args[1].value.strip():
                raise ValueError('FILTER筛选条件不能为空')
            return {'file': node.args[0].value, 'condition': node.args[1].value}
        if op not in {'UNION', 'INTERSECT', 'EXCEPT'} or len(node.args) < 2 or (op == 'EXCEPT' and len(node.args) != 2):
            raise ValueError('UNION/INTERSECT至少两个参数，EXCEPT必须两个参数')
        return {'op': op, 'items': [parse(n) for n in node.args]}
    return parse(root)


def scope_leaves(tree):
    if tree is None:
        return []
    if 'file' in tree:
        return [tree]
    return [leaf for child in tree['items'] for leaf in scope_leaves(child)]


def validate_data_scope(config, columns, key_map=None):
    if config is not None and not isinstance(config, dict):
        raise ValueError('数据范围设置无效')
    tree = scope_tree(config)
    for leaf in scope_leaves(tree):
        file = leaf['file']
        if not isinstance(file, str) or file not in columns:
            raise ValueError(f'数据范围引用了不存在的表：{file}')
        if key_map is not None and key_map.get(file) not in columns[file]:
            raise ValueError(f'数据范围所用表「{file}」必须设置有效关联键')
        if leaf['condition']:
            parse_expression(leaf['condition'], columns[file])
    return tree


def select_scope_keys(config, parsed, key_map, normalize_keys=True, date_key_mode='off'):
    columns = {file: list(info['df'].columns) for file, info in parsed.items()}
    tree = validate_data_scope(config, columns, key_map)
    if tree is None:
        return None

    def run(node):
        if 'file' in node:
            file = node['file']
            df = parsed[file]['df']
            cols = columns[file]
            condition = parse_expression(node['condition'], cols)[0] if node['condition'] else None
            selected = {}
            for row_number, values in enumerate(df.itertuples(index=False, name=None), 1):
                row = dict(zip(cols, values))
                if condition is not None:
                    try:
                        value = evaluate(condition, {f'__col_{i}': row.get(col) for i, col in enumerate(cols)})
                    except (ValueError, TypeError, KeyError, ArithmeticError) as exc:
                        raise ValueError(f'数据范围：{file}第{row_number}条数据筛选计算失败') from exc
                    if not isinstance(value, bool):
                        raise ValueError('数据范围筛选条件必须返回满足/不满足，请使用比较或AND、OR等条件')
                    if not value:
                        continue
                value = row.get(key_map[file])
                key = normalize_key(value, normalize_keys, date_key_mode)
                if key and key not in selected:
                    selected[key] = {'normalized': key, 'value': value}
            return selected
        sets = [run(child) for child in node['items']]
        result = dict(sets[0])
        for other in sets[1:]:
            if node['op'] == 'UNION':
                for key, value in other.items():
                    result.setdefault(key, value)
            elif node['op'] == 'INTERSECT':
                result = {key: value for key, value in result.items() if key in other}
            else:
                result = {key: value for key, value in result.items() if key not in other}
        return result
    return list(run(tree).values())


def translate_data_scope(config, file_map, column_maps, replace_tokens):
    """Translate only table references and condition columns, never literal text."""
    result = copy.deepcopy(config)
    tree = scope_tree(result)
    if tree is None:
        return result
    for leaf in scope_leaves(tree):
        old_file = leaf['file']
        mapping = column_maps.get(old_file, {})
        replacements = {**mapping, **{f'[{a}]': f'[{b}]' for a, b in mapping.items()}}
        leaf['condition'] = replace_tokens(leaf['condition'], replacements)
        leaf['file'] = file_map.get(old_file, old_file)
    if result['mode'] == 'source':
        result.update(source_file=tree['file'], filter_expression=tree['condition'])
    else:
        import json
        def render(node):
            if 'file' in node:
                file = json.dumps(node['file'], ensure_ascii=False)
                return f'FILTER({file},{json.dumps(node["condition"], ensure_ascii=False)})' if node['condition'] else file
            return node['op'] + '(' + ','.join(render(n) for n in node['items']) + ')'
        result['expression'] = render(tree)
    return result
