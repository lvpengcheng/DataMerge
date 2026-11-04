"""Scope-aware checks for generated Python; never execute code to inspect it."""
import ast
import re
from backend.utils.column_spec_schema import DECLARATION_CONTRACT


CODE_CONTRACT = DECLARATION_CONTRACT + """
代码完整性要求：辅助函数的定义、import、调用和参数必须一起提供；修改变量名时同步修改
所有作用域的引用。每行记录依赖的变量须在该行循环内、进入条件分支之前初始化；
嵌套辅助函数优先显式传入记录和行号，避免读取未赋值闭包变量。不能通过无依据地补 None/0
掩盖缺失定义。字段名含空格、换行、括号、中文时使用经列映射确认的字符串下标，
不要把字段名直接转为 Python 变量名或属性。每行都需要的写值/写公式逻辑必须覆盖首行；
只针对第二行后的条件块不得包住其他字段处理。不要保留含 #REF! 的模板公式作为正确结果。
检查辅助函数和第三方对象的方法确实存在；不能虚构 API。检查零行、单行和多行路径。
已有模块级 _COLUMN_SPECS 列契约必须保留，不能删除、动态重写或在函数内遮蔽；
列类型/格式要求须持久化为模块级 _COLUMN_SPECS 纯字面量声明。

主键统一规则：宿主固定函数负责把源数据、模板及结果里的匹配键统一为文本。
所有参与 VLOOKUP/XLOOKUP/MATCH/merge/join 的身份标识列（包括两端不同列名）须在
_COLUMN_SPECS 中明确声明 type="key"，使用实际 sheet 名和实际表头；code/编号 本身
不能说明它是主键，只有根据业务规则或字段映射确定用于匹配时才声明。
不要自行编写主键转型逻辑、不要用 int/float/VALUE/双负号把主键转回数字，不要为类型
匹配额外创建隐藏辅助列。保留前导零，身份证等长编号禁止经过浮点数。
需要在函数内归一新建 DataFrame 或 Workbook 时调用
backend.utils.column_contract.normalize_source_keys / normalize_workbook_keys，参数为数据
及 globals().get('_COLUMN_SPECS')；固定骨架在填充前后和最终计算前也会执行同一规则。
只有用户明确要求保留原类型时声明 type="raw"，涉及的查询两端也必须同步兼容该选择。
主键列写入实际文本值，不写自引用查找公式；缺失匹配应保留报错，不能填 0 或空白掩盖。
主键是文本：人数按非空或业务要求去重计数（如 COUNTA），不能用 COUNT/求和统计工号；
排序、补齐长度、大小写及去前导零必须有明确业务规则，不得擅自改变标识值。
"""


def _guarded_optional_calls(tree):
    """Names called only after an explicit globals existence/callability guard."""
    def globals_call(node):
        return (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                and node.func.id == 'globals' and not node.args and not node.keywords)

    allowed = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.If):
            continue
        test, name = node.test, None
        if (isinstance(test, ast.Call) and isinstance(test.func, ast.Name)
                and test.func.id == 'callable' and len(test.args) == 1):
            get = test.args[0]
            if (isinstance(get, ast.Call) and isinstance(get.func, ast.Attribute)
                    and get.func.attr == 'get' and globals_call(get.func.value)
                    and len(get.args) == 1 and isinstance(get.args[0], ast.Constant)):
                name = get.args[0].value
        elif (isinstance(test, ast.Compare) and len(test.ops) == 1
              and isinstance(test.ops[0], ast.In) and isinstance(test.left, ast.Constant)
              and globals_call(test.comparators[0])):
            name = test.left.value
        if not isinstance(name, str):
            continue
        pending = list(node.body)
        while pending:
            child = pending.pop()
            if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda)):
                continue
            if isinstance(child, ast.Call) and isinstance(child.func, ast.Name) and child.func.id == name:
                allowed.add((child.func.lineno, name))
            pending.extend(ast.iter_child_nodes(child))
    return allowed


def check_generated_code(code, available_names=()):
    from pyflakes.checker import Checker
    from pyflakes.messages import UndefinedName, UndefinedLocal
    tree = ast.parse(code)
    # Each visible workbook section executes with isolated globals. Validate its
    # module constants independently; declarations in sibling sections are not rewrites.
    from backend.utils.multi_workbook import script_targets, script_wrapper, _target_assignment, _entry_block
    if _target_assignment(tree) is not None:
        targets = script_targets(code)
        issues = check_generated_code(script_wrapper(code), {*available_names, 'MULTI_WORKBOOK_SCRIPTS'})
        metadata = ast.literal_eval(_target_assignment(tree).value)
        for name, target in targets.items():
            entry = metadata[name].get('entry')
            block = next((node for node in tree.body if entry and _entry_block(node, entry)), None)
            future_count = len(metadata[name].get('future_imports') or [])
            offset = block.lineno - future_count if block else 0
            for issue in check_generated_code(target['code'], available_names):
                issues.append({**issue, 'line': issue['line'] + offset,
                               'message': f"[{name}] {issue['message']}"})
        return issues
    checker = Checker(tree, filename='<generated>', builtins=set(available_names))
    guarded = _guarded_optional_calls(tree)
    issues = [
        {'line': item.lineno, 'kind': type(item).__name__,
         'message': item.message % item.message_args}
        for item in checker.messages
        if isinstance(item, (UndefinedName, UndefinedLocal))
        and not (isinstance(item, UndefinedName) and (item.lineno, item.message_args[0]) in guarded)
    ]
    from backend.utils.column_spec_edit import inspect_column_specs
    try:
        inspect_column_specs(code, allow_legacy=True)
    except ValueError as exc:
        issues.append({'line': 1, 'kind': 'ColumnSpecContract', 'message': str(exc)})
    # 仅约束新骨架，避免把历史脚本直接判为不可执行；新代码有匹配却未声明键时进入修复。
    versioned = any(isinstance(node, ast.Assign)
                    and any(isinstance(t, ast.Name) and t.id == 'KEY_POLICY_VERSION' for t in node.targets)
                    and isinstance(node.value, ast.Constant) and node.value.value == 1
                    for node in tree.body)
    lookup = any(isinstance(node, ast.Constant) and isinstance(node.value, str)
                 and re.search(r'\b(?:VLOOKUP|XLOOKUP|MATCH)\s*\(', node.value, re.I)
                 or isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                 and node.func.attr == 'merge' for node in ast.walk(tree))
    if versioned and lookup:
        from backend.utils.column_contract import read_column_specs
        specs = read_column_specs(code)
        if not any(rule.get('type') in ('key', 'raw') for columns in specs.values() for rule in columns.values()):
            issues.append({'line': 1, 'kind': 'KeyPolicyContract',
                           'message': '匹配公式/merge 必须用 _COLUMN_SPECS 声明两端主键 type=key；明确保留类型时用 raw'})
    return issues


def describe_issues(issues):
    return '\n'.join(f"line {x['line']}: {x['message']}" for x in issues[:30])
