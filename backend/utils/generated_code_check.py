"""Scope-aware checks for generated Python; never execute code to inspect it."""
import ast


CODE_CONTRACT = """
代码完整性要求：辅助函数的定义、import、调用和参数必须一起提供；修改变量名时同步修改
所有作用域的引用。每行记录依赖的变量须在该行循环内、进入条件分支之前初始化；
嵌套辅助函数优先显式传入记录和行号，避免读取未赋值闭包变量。不能通过无依据地补 None/0
掩盖缺失定义。字段名含空格、换行、括号、中文时使用经列映射确认的字符串下标，
不要把字段名直接转为 Python 变量名或属性。每行都需要的写值/写公式逻辑必须覆盖首行；
只针对第二行后的条件块不得包住其他字段处理。不要保留含 #REF! 的模板公式作为正确结果。
检查辅助函数和第三方对象的方法确实存在；不能虚构 API。检查零行、单行和多行路径。
"""


def check_generated_code(code, available_names=()):
    from pyflakes.checker import Checker
    from pyflakes.messages import UndefinedName, UndefinedLocal
    tree = ast.parse(code)
    checker = Checker(tree, filename='<generated>', builtins=set(available_names))
    return [
        {'line': item.lineno, 'kind': type(item).__name__,
         'message': item.message % item.message_args}
        for item in checker.messages
        if isinstance(item, (UndefinedName, UndefinedLocal))
    ]


def describe_issues(issues):
    return '\n'.join(f"line {x['line']}: {x['message']}" for x in issues[:30])
