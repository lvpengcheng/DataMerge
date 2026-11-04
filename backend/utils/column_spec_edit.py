"""Persist AI column declarations without executing generated Python."""
import ast
import copy
import pprint
import re

from backend.utils.column_spec_schema import (KINDS, header_key as _header_key,
    decode_column_specs, literal_column_specs, DECLARATION_CONTRACT)


COLUMN_SPEC_EDIT_CONTRACT = """
列类型/显示格式修改必须持久化为模块级 _COLUMN_SPECS 纯字面量字典，由宿主写入，
无需用户手写。响应 JSON 除 edits 外增加 column_specs（只放本轮明确要求的列声明）：
{"edits": [], "column_specs": {"*": {"工号": {"type": "key", "format": "@"},
"入职日期": {"type": "date", "format": "yyyy-mm-dd"}}}}
没有代码逻辑改动时 edits 可为空；没有列契约改动时 column_specs 为 {}。
sheet 明确时使用准确 sheet 名，未限定 sheet 时用 "*"；列名必须使用真实表头，忽略空格匹配。
支持 type: key/text/date/datetime/number/raw；format 为 Excel 格式码，parse 为输入日期解析格式。
标识符/工号保留前导零用 key；日期去时间用 date，保留时间用 datetime；金额/百分比用 number。
仅调整显示时可只给 format，不要擅自改变值类型。不要把 Python strftime 格式写入 format。
不要在 edits 中删除或修改已有 _COLUMN_SPECS；宿主会按列合并 column_specs 并保留其他声明。
用户要求列格式时必须返回非空 column_specs，不能仅写 number_format 或靠文字说明。
已有显式 number_format 优先级高于列声明：本轮涉及的列如有冲突，需在 edits 中移除或同步这些
显式赋值，避免覆盖新声明；不得改动其他列。legacy 模式还需同步显式格式以兼容旧后处理。
多文件模式每次只编辑当前交付文件：column_specs 直接按 sheet 分组，不要再套文件名一层。
""" + DECLARATION_CONTRACT


def _validate_specs(specs):
    if not isinstance(specs, dict):
        raise ValueError('_COLUMN_SPECS 必须是字典')
    normalized = decode_column_specs(specs)
    if normalized != specs:
        raise ValueError('_COLUMN_SPECS 需要标准的 sheet/列字典；可通过统一解析器规范化来源元信息或类型别名')
    return specs


def normalize_spec_updates(updates, workbook_name=None):
    return decode_column_specs(updates, workbook_name=workbook_name)


def detach_generated_specs(code):
    """合并分步生成的模块级列声明，避免函数抽取把它们当游离代码删掉。"""
    tree = ast.parse(code)
    merged, declarations = {}, []
    for node in tree.body:
        targets = node.targets if isinstance(node, ast.Assign) else (
            [node.target] if isinstance(node, ast.AnnAssign) else [])
        if not any(isinstance(target, ast.Name) and target.id == '_COLUMN_SPECS' for target in targets):
            continue
        if len(targets) != 1:
            raise ValueError('_COLUMN_SPECS 必须独立赋值')
        values = normalize_spec_updates(literal_column_specs(node.value))
        for sheet, columns in values.items():
            for column, rule in columns.items():
                target = merged.setdefault(sheet, {})
                name = next((n for n in target if _header_key(n) == _header_key(column)), column)
                prior = target.get(name, {})
                prior = {'type': prior} if isinstance(prior, str) else prior
                rule = {'type': rule} if isinstance(rule, str) else rule
                if prior.get('type') and rule.get('type') and prior['type'] != rule['type']:
                    raise ValueError(f'分步生成的 {sheet}.{column} 类型声明冲突')
                target[name] = {**prior, **rule}
        declarations.append(node)
        if node.col_offset or any(other is not node and other.lineno <= node.end_lineno
                                  and other.end_lineno >= node.lineno for other in tree.body):
            raise ValueError('_COLUMN_SPECS 声明必须独占语句行')
    lines = code.splitlines(keepends=True)
    for node in reversed(declarations):
        lines[node.lineno - 1:node.end_lineno] = []
    prefix = '_COLUMN_SPECS = ' + pprint.pformat(merged, sort_dicts=False) + '\n' if declarations else ''
    return prefix, ''.join(lines)


def _root(node):
    while isinstance(node, (ast.Subscript, ast.Attribute)):
        node = node.value
    return isinstance(node, ast.Name) and node.id == '_COLUMN_SPECS'


def inspect_column_specs(code, allow_legacy=False):
    """Strict AST check: one module literal, no later reassignment/mutation."""
    tree = ast.parse(code)
    declarations = []
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(_root(t) for t in node.targets):
            if len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
                raise ValueError('_COLUMN_SPECS 必须独立赋值')
            declarations.append(node)
        elif isinstance(node, ast.AnnAssign) and _root(node.target):
            declarations.append(node)
    if len(declarations) > 1:
        raise ValueError('_COLUMN_SPECS 不能重复声明')
    declaration = declarations[0] if declarations else None
    allowed = declaration.targets[0] if isinstance(declaration, ast.Assign) else (
        declaration.target if declaration else None)
    for node in ast.walk(tree):
        if isinstance(node, (ast.Name, ast.Subscript, ast.Attribute)) and isinstance(node.ctx, (ast.Store, ast.Del)):
            if _root(node) and node is not allowed:
                raise ValueError('_COLUMN_SPECS 不允许在运行时重写或删除')
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and _root(node.func):
            if node.func.attr in {'update', 'clear', 'pop', 'popitem', 'setdefault', '__setitem__', '__delitem__'}:
                raise ValueError('_COLUMN_SPECS 不允许在运行时修改')
    if declaration is None:
        return {}, None, tree
    try:
        specs = literal_column_specs(declaration.value)
    except (ValueError, TypeError) as exc:
        raise ValueError(f'_COLUMN_SPECS 必须是纯字面量且不能有重复键: {exc}') from exc
    if allow_legacy:
        specs = decode_column_specs(specs)
    return _validate_specs(specs), declaration, tree


def canonicalize_column_specs(code):
    """Normalize known literal declaration shapes without executing code or changing rules."""
    specs, declaration, _ = inspect_column_specs(code, allow_legacy=True)
    if declaration is None or ast.literal_eval(declaration.value) == specs:
        return code
    value = declaration.value
    lines = code.splitlines(keepends=True)
    first, last = value.lineno - 1, value.end_lineno - 1
    prefix = lines[first].encode('utf-8')[:value.col_offset].decode('utf-8')
    suffix = lines[last].encode('utf-8')[value.end_col_offset:].decode('utf-8')
    lines[first:last + 1] = [prefix + pprint.pformat(specs, sort_dicts=False, width=100) + suffix]
    result = ''.join(lines)
    assert inspect_column_specs(result)[0] == specs
    return result


def prepare_column_contract_script(code):
    """Shared boundary for generated/imported single and multi-workbook scripts."""
    from backend.utils.multi_workbook import script_targets, script_wrapper, build_script
    targets = script_targets(code)
    if not targets:
        return canonicalize_column_specs(code)
    changed = False
    for target in targets.values():
        normalized = prepare_column_contract_script(target['code'])
        changed |= normalized != target['code']
        target['code'] = normalized
    return build_script(targets, wrapper=script_wrapper(code)) if changed else code


def requires_column_specs(feedback):
    """Catch common format/type requests so a legacy number_format edit cannot pass silently."""
    feedback = re.sub(r'(?:保持|保留)?(?:原有|原|现有|其他|其它|其余)?(?:的)?(?:列)?格式(?:保持)?不变', '', feedback or '')
    return bool(re.search(
        r'格式|前导零|全角|半角|不可见字符|去掉时间|截掉时间|保留时间|'
        r'(?:保留|显示).{0,8}(?:位小数|年月日)|'
        r'(?:改成|改为|转成|转为|设为|设置为|按).{0,16}(?:文本|日期|数值|数字|百分比|小数|整数)',
        feedback or '', re.S))


def merge_column_spec_edit(original, edited, updates, workbook_name=None):
    """Reject lost declarations, then write validated, field-wise merged literals."""
    old, old_declaration, _ = inspect_column_specs(original, allow_legacy=True)
    legacy = old_declaration is not None and ast.literal_eval(old_declaration.value) != old
    edited = canonicalize_column_specs(edited)
    current, declaration, tree = inspect_column_specs(edited)
    updates = normalize_spec_updates(updates, workbook_name)
    if current != old:
        # A broken legacy list must be repairable. Require every old field to
        # survive; normal valid declarations still require structured updates.
        preserved = legacy
        for sheet, columns in old.items():
            current_columns = current.get(sheet, {})
            for column, rule in columns.items():
                name = next((key for key in current_columns if _header_key(key) == _header_key(column)), column)
                candidate = current_columns.get(name, {})
                candidate = {'type': candidate} if isinstance(candidate, str) else candidate
                rule = {'type': rule} if isinstance(rule, str) else rule
                preserved = preserved and all(candidate.get(key) == value for key, value in rule.items())
        if not preserved:
            raise ValueError('已有 _COLUMN_SPECS 被删除或改写；请保留原声明，通过 column_specs 提交变更')
    if not updates:
        return edited
    merged = copy.deepcopy(current)
    for sheet, columns in updates.items():
        target = merged.setdefault(sheet, {})
        for column, spec in columns.items():
            column = next((name for name in target if _header_key(name) == _header_key(column)), column)
            prior = target.get(column, {})
            prior = {'type': prior} if isinstance(prior, str) else prior
            change = {'type': spec} if isinstance(spec, str) else spec
            target[column] = {**prior, **change}
    literal = '_COLUMN_SPECS = ' + pprint.pformat(merged, sort_dicts=False, width=100) + '\n'
    lines = edited.splitlines(keepends=True)
    if declaration:
        first, last = declaration.lineno - 1, declaration.end_lineno - 1
        prefix = lines[first].encode('utf-8')[:declaration.col_offset].decode('utf-8')
        suffix = lines[last].encode('utf-8')[declaration.end_col_offset:].decode('utf-8')
        lines[first:last + 1] = [prefix + literal.rstrip('\n') + suffix]
    else:
        # Preserve module docstring, encoding/shebang comments and future imports.
        line = 0
        for index, node in enumerate(tree.body):
            if (index == 0 and isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant)
                    and isinstance(node.value.value, str)) or (
                    isinstance(node, ast.ImportFrom) and node.module == '__future__'):
                line = node.end_lineno
            else:
                if not line:
                    line = node.lineno - 1
                break
        if line and not lines[line - 1].endswith('\n'):
            lines[line - 1] += '\n'
        lines[line:line] = [literal]
    result = ''.join(lines)
    actual, _, _ = inspect_column_specs(result)
    if actual != merged:
        raise ValueError('_COLUMN_SPECS AST 写入校验失败')
    compile(result, '<column_specs>', 'exec')
    return result
