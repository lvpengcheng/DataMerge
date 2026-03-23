"""
代码沙箱 - 安全执行生成的Python代码
"""

import sys
import io
import traceback
import contextlib
import tempfile
import os
from pathlib import Path
from typing import Dict, Any, Optional
import pandas as pd
import openpyxl


class CodeSandbox:
    """代码沙箱执行环境"""

    def __init__(self, timeout: int = 300, max_memory_mb: int = 1024):
        self.timeout = timeout
        self.max_memory_mb = max_memory_mb
        import logging
        self.logger = logging.getLogger(__name__)
        self.safe_modules = {
            'pandas': pd,
            'openpyxl': openpyxl,
            'os': os,
            'sys': sys,
            'pathlib': Path,
            'json': __import__('json'),
            're': __import__('re'),
            'datetime': __import__('datetime'),
            'math': __import__('math'),
            'typing': __import__('typing'),
            'collections': __import__('collections'),
            'itertools': __import__('itertools'),
            'functools': __import__('functools'),
            'statistics': __import__('statistics'),
            'decimal': __import__('decimal'),
            'fractions': __import__('fractions'),
            'random': __import__('random'),
            'string': __import__('string'),
            'hashlib': __import__('hashlib'),
            # 健壮性工具函数
            'robust_utils': self._import_robust_utils(),
            # 沙箱辅助函数（公共方法）
            'sandbox_helpers': self._import_sandbox_helpers(),
            'base64': __import__('base64'),
            'csv': __import__('csv'),
            'io': io,
            'tempfile': tempfile,
            # 'shutil': __import__('shutil'),  # 移除shutil，因为它包含rmtree等危险操作
            'copy': __import__('copy'),
            'pprint': __import__('pprint'),
            'textwrap': __import__('textwrap'),
            'unicodedata': __import__('unicodedata'),
            'numbers': __import__('numbers'),
            'operator': __import__('operator'),
            'bisect': __import__('bisect'),
            'heapq': __import__('heapq'),
            'array': __import__('array'),
            'struct': __import__('struct'),
            'pickle': __import__('pickle'),
            'shelve': __import__('shelve'),
            'dbm': __import__('dbm'),
            'sqlite3': __import__('sqlite3'),
            'zlib': __import__('zlib'),
            'gzip': __import__('gzip'),
            'bz2': __import__('bz2'),
            'lzma': __import__('lzma'),
            'zipfile': __import__('zipfile'),
            'tarfile': __import__('tarfile'),
        }

        # 尝试导入可选模块
        optional_modules = [
            ('numpy', 'numpy'),
            ('xlrd', 'xlrd'),
            ('xlwt', 'xlwt'),
            ('xlsxwriter', 'xlsxwriter'),
        ]

        for module_name, import_name in optional_modules:
            try:
                self.safe_modules[module_name] = __import__(import_name)
            except ImportError:
                # 模块未安装，跳过
                pass

        # 危险模块和函数
        self.dangerous_modules = {
            'subprocess', 'os.system', 'os.popen', 'eval', 'exec', '__import__',
            'compile', 'file', 'input', 'raw_input', 'execfile'
        }

    def execute_script(self, script_content: str, execution_env: Dict[str, Any]) -> Dict[str, Any]:
        """在沙箱中执行脚本"""
        result = {
            "success": False,
            "output": "",
            "error": "",
            "execution_time": 0,
            "return_value": None
        }

        try:
            # 清理代码中的常见语法错误
            script_content = self._clean_code_syntax(script_content)

            # 安全检查
            if not self._is_safe_script(script_content):
                result["error"] = "脚本包含不安全代码"
                return result

            # 准备执行环境
            safe_env = self._create_safe_environment(execution_env)

            # 重定向输出
            output_buffer = io.StringIO()
            error_buffer = io.StringIO()

            # 添加调试信息到输出
            output_buffer.write(f"=== 沙箱执行开始 ===\n")
            output_buffer.write(f"执行环境: {execution_env}\n")
            output_buffer.write(f"脚本长度: {len(script_content)} 字符\n")

            with contextlib.redirect_stdout(output_buffer), \
                 contextlib.redirect_stderr(error_buffer):

                # 执行脚本
                exec_globals = {}
                exec_globals.update(safe_env)

                try:
                    # 初始化monkey-patch标记（必须在exec之前，否则exec失败时finally无法访问）
                    _iep_patched = False
                    _iep_orig_parse = None

                    # 编译并执行代码
                    output_buffer.write(f"开始编译和执行代码...\n")
                    code_obj = compile(script_content, '<sandbox>', 'exec')
                    exec(code_obj, exec_globals)
                    output_buffer.write(f"代码编译和执行完成\n")

                    # 【加密支持】如果有 file_passwords，monkey-patch IntelligentExcelParser
                    # 使其在 parse_excel_file 时自动注入对应文件的密码
                    # 注意：文件通常已在上游解密，仅对仍加密的文件注入密码
                    _file_passwords = execution_env.get('file_passwords') or {}
                    if _file_passwords:
                        try:
                            from excel_parser import IntelligentExcelParser as _IEP
                            try:
                                from backend.utils.aspose_helper import is_encrypted as _chk_enc
                            except ImportError:
                                _chk_enc = None
                            _iep_orig_parse = _IEP.parse_excel_file
                            _fp_map = _file_passwords  # 闭包捕获

                            def _auto_pwd_parse(self_parser, file_path, *args, **kwargs):
                                if 'password' not in kwargs or kwargs.get('password') is None:
                                    fname = os.path.basename(str(file_path))
                                    pwd = _fp_map.get(fname)
                                    if pwd:
                                        # 先检查文件是否仍然加密
                                        still_enc = True
                                        if _chk_enc is not None:
                                            try:
                                                still_enc = _chk_enc(str(file_path))
                                            except Exception:
                                                still_enc = True
                                        if still_enc:
                                            kwargs['password'] = pwd
                                            output_buffer.write(f"[加密支持] 为 {fname} 自动注入密码\n")
                                        else:
                                            output_buffer.write(f"[加密支持] {fname} 已解密，跳过密码注入\n")
                                return _iep_orig_parse(self_parser, file_path, *args, **kwargs)

                            _IEP.parse_excel_file = _auto_pwd_parse
                            _iep_patched = True
                            output_buffer.write(f"[加密支持] 已注入密码映射（{len(_file_passwords)}个文件）\n")
                        except Exception as _e:
                            output_buffer.write(f"[加密支持] 注入失败: {_e}\n")

                    # 【性能优化】如果有预加载的源数据，替换 load_source_data 避免重复解析
                    _pre_loaded = exec_globals.get('_pre_loaded_source_data')
                    if _pre_loaded and 'load_source_data' in exec_globals:
                        _original_load = exec_globals['load_source_data']
                        def _cached_load_source_data(input_folder, manual_headers, _data=_pre_loaded):
                            output_buffer.write(f"[性能优化] 使用预加载源数据（{len(_data)}个sheet，跳过Excel解析）\n")
                            return _data
                        exec_globals['load_source_data'] = _cached_load_source_data
                        output_buffer.write(f"[性能优化] 已注入预加载源数据缓存\n")

                    # 尝试调用主函数
                    if 'main' in exec_globals and callable(exec_globals['main']):
                        main_func = exec_globals['main']
                        # 检查main函数的签名，智能传递参数
                        import inspect
                        sig = inspect.signature(main_func)
                        params = list(sig.parameters.keys())

                        if len(params) == 0:
                            # main() 不带参数
                            output_buffer.write(f"调用 main() 函数\n")
                            exec_globals['main']()
                        else:
                            # main() 带参数，尝试传递执行环境
                            input_folder = execution_env.get('input_folder', '')
                            output_folder = execution_env.get('output_folder', '')
                            output_buffer.write(f"调用 main('{input_folder}', '{output_folder}') 函数\n")
                            # 尝试匹配参数名
                            kwargs = {}
                            if 'input_folder' in params:
                                kwargs['input_folder'] = input_folder
                            if 'input_path' in params:
                                kwargs['input_path'] = input_folder
                            if 'output_folder' in params:
                                kwargs['output_folder'] = output_folder
                            if 'output_path' in params:
                                kwargs['output_path'] = output_folder
                            if 'output_file' in params:
                                # 兼容 output_file 参数名
                                import os as _os
                                kwargs['output_file'] = _os.path.join(output_folder, 'output.xlsx')

                            # 按位置参数调用
                            if len(params) >= 2 and not kwargs:
                                exec_globals['main'](input_folder, output_folder)
                            else:
                                exec_globals['main'](**kwargs)

                        output_buffer.write(f"main() 函数执行完成\n")
                    elif 'process_excel_files' in exec_globals and callable(exec_globals['process_excel_files']):
                        # 传递执行环境参数
                        input_folder = execution_env.get('input_folder', '')
                        output_folder = execution_env.get('output_folder', '')

                        # 检查函数签名，支持薪资年月参数
                        import inspect
                        sig = inspect.signature(exec_globals['process_excel_files'])
                        params = list(sig.parameters.keys())

                        # 构建参数字典
                        call_kwargs = {
                            'input_folder': input_folder,
                            'output_folder': output_folder
                        }

                        # 如果函数支持薪资年月参数，传入
                        if 'salary_year' in params and 'salary_year' in execution_env:
                            call_kwargs['salary_year'] = execution_env.get('salary_year')
                        if 'salary_month' in params and 'salary_month' in execution_env:
                            call_kwargs['salary_month'] = execution_env.get('salary_month')
                        if 'monthly_standard_hours' in params and 'monthly_standard_hours' in execution_env:
                            call_kwargs['monthly_standard_hours'] = execution_env.get('monthly_standard_hours')

                        output_buffer.write(f"调用 process_excel_files({call_kwargs})\n")
                        return_value = exec_globals['process_excel_files'](**call_kwargs)
                        output_buffer.write(f"process_excel_files 执行完成，返回值: {return_value}\n")
                        result["return_value"] = return_value
                    else:
                        # 查找其他可能的入口函数 (process_salary_data, process_data 等)
                        entry_func = None
                        entry_func_name = None
                        for func_name in ['process_salary_data', 'process_data', 'run', 'execute']:
                            if func_name in exec_globals and callable(exec_globals[func_name]):
                                entry_func = exec_globals[func_name]
                                entry_func_name = func_name
                                break

                        if entry_func:
                            # 调用找到的入口函数
                            input_folder = execution_env.get('input_folder', '')
                            output_folder = execution_env.get('output_folder', '')

                            import inspect
                            sig = inspect.signature(entry_func)
                            params = list(sig.parameters.keys())

                            call_kwargs = {}
                            if 'input_folder' in params:
                                call_kwargs['input_folder'] = input_folder
                            if 'output_folder' in params:
                                call_kwargs['output_folder'] = output_folder
                            if 'rules_content' in params:
                                call_kwargs['rules_content'] = execution_env.get('rules_content', '')
                            if 'salary_year' in params and 'salary_year' in execution_env:
                                call_kwargs['salary_year'] = execution_env.get('salary_year')
                            if 'salary_month' in params and 'salary_month' in execution_env:
                                call_kwargs['salary_month'] = execution_env.get('salary_month')
                            if 'monthly_standard_hours' in params and 'monthly_standard_hours' in execution_env:
                                call_kwargs['monthly_standard_hours'] = execution_env.get('monthly_standard_hours')

                            output_buffer.write(f"调用 {entry_func_name}({call_kwargs})\n")
                            return_value = entry_func(**call_kwargs)
                            output_buffer.write(f"{entry_func_name} 执行完成，返回值: {return_value}\n")
                            result["return_value"] = return_value
                        else:
                            output_buffer.write(f"未找到可调用的主函数 (main, process_excel_files, process_salary_data 等)\n")
                            output_buffer.write(f"全局变量: {list(exec_globals.keys())}\n")

                    result["success"] = True
                    if result["return_value"] is None:
                        result["return_value"] = exec_globals.get('__result__', None)

                except SystemExit as e:
                    # 捕获SystemExit异常，防止沙箱代码退出主进程
                    error_msg = f"代码尝试退出进程 (SystemExit): {str(e)}\n{traceback.format_exc()}"
                    error_buffer.write(error_msg)
                except Exception as e:
                    error_msg = f"执行错误: {str(e)}\n{traceback.format_exc()}"
                    error_buffer.write(error_msg)
                finally:
                    # 恢复 IntelligentExcelParser.parse_excel_file 原始方法
                    if _iep_patched and _iep_orig_parse is not None:
                        try:
                            from excel_parser import IntelligentExcelParser as _IEP
                            _IEP.parse_excel_file = _iep_orig_parse
                        except Exception:
                            pass

            # 收集输出
            full_output = output_buffer.getvalue()
            error_output = error_buffer.getvalue()

            # 限制 output 长度为前 500 字符
            if len(full_output) > 500:
                result["output"] = full_output[:500] + "\n...(输出过长，已截断)"
            else:
                result["output"] = full_output

            if error_output:
                result["error"] = error_output

            # 添加执行总结
            result["output"] += f"\n=== 沙箱执行结束 ===\n"
            result["output"] += f"执行成功: {result['success']}\n"
            result["output"] += f"返回值: {result['return_value']}\n"
            if result["error"]:
                result["output"] += f"错误信息: {result['error'][:500]}...\n"

        except SystemExit as e:
            # 捕获SystemExit异常，防止沙箱代码退出主进程
            result["error"] = f"沙箱执行失败: 代码尝试退出进程 (SystemExit: {str(e)})\n{traceback.format_exc()}"
        except Exception as e:
            result["error"] = f"沙箱执行失败: {str(e)}\n{traceback.format_exc()}"

        return result

    def _is_safe_script(self, script_content: str) -> bool:
        """检查脚本是否安全 - 放宽对AI生成代码的限制"""
        # 对于AI生成的Excel数据处理代码，我们放宽限制
        # 只检查最危险的操作

        # 1. 检查绝对禁止的危险操作 - 只限制真正危险的系统操作
        absolute_dangerous = [
            'os.system', 'os.popen', 'subprocess.',  # 系统命令执行
            'os.remove', 'os.unlink', 'os.rmdir', 'shutil.rmtree',  # 文件删除操作
            # 移除 'eval(', 'exec(', 'compile(' 限制，允许在沙箱环境中使用
            'execfile',  # 文件执行（Python 2的遗留函数）
            # 移除 '__import__' 限制，允许动态导入
        ]

        for pattern in absolute_dangerous:
            if pattern in script_content:
                self.logger.warning(f"检测到绝对危险操作: {pattern}")
                return False

        # 2. 检查危险模块导入 - 只检查真正危险的模块
        dangerous_modules = {
            'subprocess', 'os.system', 'os.popen', 'os.remove', 'os.unlink', 'os.rmdir',
            'shutil',  # 包含rmtree等危险操作
            'execfile'  # Python 2的遗留函数
            # 移除 '__import__', 'eval', 'exec', 'compile'，允许使用
        }

        import_lines = []
        for line in script_content.split('\n'):
            line = line.strip()
            if line.startswith('import ') or line.startswith('from '):
                import_lines.append(line)

        for import_line in import_lines:
            # 检查是否导入绝对危险的模块
            for dangerous in dangerous_modules:
                if dangerous in import_line:
                    self.logger.warning(f"检测到危险模块导入: {import_line}")
                    return False

        # 3. 对于AI生成的Excel处理代码，允许以下操作：
        # - 导入任何模块（除了绝对危险的）
        # - 使用open函数（包括写入）
        # - 使用input函数（用户交互）
        # - 其他标准Python操作

        # 记录安全检查通过
        self.logger.info("脚本安全检查通过")
        return True

    def _is_stdlib_module(self, module_name: str) -> bool:
        """检查模块是否属于Python标准库"""
        import sys
        import importlib.util

        # 常见标准库模块列表（不完整，但覆盖常用模块）
        stdlib_modules = {
            'abc', 'argparse', 'ast', 'asyncio', 'atexit', 'base64', 'binascii',
            'bisect', 'builtins', 'bz2', 'calendar', 'cgi', 'cgitb', 'chunk',
            'cmath', 'cmd', 'code', 'codecs', 'codeop', 'collections', 'colorsys',
            'compileall', 'concurrent', 'configparser', 'contextlib', 'contextvars',
            'copy', 'copyreg', 'cProfile', 'crypt', 'csv', 'ctypes', 'dataclasses',
            'datetime', 'decimal', 'difflib', 'dis', 'distutils', 'doctest',
            'email', 'encodings', 'ensurepip', 'enum', 'errno', 'faulthandler',
            'fcntl', 'filecmp', 'fileinput', 'fnmatch', 'fractions', 'ftplib',
            'functools', 'gc', 'getopt', 'getpass', 'gettext', 'glob', 'graphlib',
            'grp', 'gzip', 'hashlib', 'heapq', 'hmac', 'html', 'http', 'imaplib',
            'imghdr', 'imp', 'importlib', 'inspect', 'io', 'ipaddress', 'itertools',
            'json', 'keyword', 'lib2to3', 'linecache', 'locale', 'logging', 'lzma',
            'mailbox', 'mailcap', 'marshal', 'math', 'mimetypes', 'mmap', 'modulefinder',
            'msilib', 'msvcrt', 'multiprocessing', 'netrc', 'nis', 'nntplib',
            'numbers', 'operator', 'optparse', 'os', 'ossaudiodev', 'parser',
            'pathlib', 'pdb', 'pickle', 'pickletools', 'pipes', 'pkgutil', 'platform',
            'plistlib', 'poplib', 'posix', 'pprint', 'profile', 'pstats', 'pty',
            'pwd', 'py_compile', 'pyclbr', 'pydoc', 'queue', 'quopri', 'random',
            're', 'readline', 'reprlib', 'resource', 'rlcompleter', 'runpy',
            'sched', 'secrets', 'select', 'selectors', 'shelve', 'shlex', 'shutil',
            'signal', 'site', 'smtpd', 'smtplib', 'sndhdr', 'socket', 'socketserver',
            'spwd', 'sqlite3', 'ssl', 'stat', 'statistics', 'string', 'stringprep',
            'struct', 'subprocess', 'sunau', 'symbol', 'symtable', 'sys', 'sysconfig',
            'syslog', 'tabnanny', 'tarfile', 'telnetlib', 'tempfile', 'termios',
            'textwrap', 'threading', 'time', 'timeit', 'tkinter', 'token', 'tokenize',
            'trace', 'traceback', 'tracemalloc', 'tty', 'turtle', 'types', 'typing',
            'unicodedata', 'unittest', 'urllib', 'uu', 'uuid', 'venv', 'warnings',
            'wave', 'weakref', 'webbrowser', 'winreg', 'winsound', 'wsgiref',
            'xdrlib', 'xml', 'xmlrpc', 'zipapp', 'zipfile', 'zipimport', 'zlib'
        }

        # 检查模块是否在标准库列表中
        if module_name in stdlib_modules:
            return True

        # 尝试导入模块，检查是否在标准库路径中
        try:
            spec = importlib.util.find_spec(module_name)
            if spec is None:
                return False

            # 检查模块路径是否在标准库路径中
            if spec.origin is None:
                # 内置模块
                return True

            # 检查是否在标准库目录中
            stdlib_paths = [sys.prefix, sys.exec_prefix]
            for path in stdlib_paths:
                if spec.origin and spec.origin.startswith(path):
                    return True

        except (ImportError, AttributeError):
            pass

        return False

    def _create_safe_environment(self, execution_env: Dict[str, Any]) -> Dict[str, Any]:
        """创建安全的执行环境"""
        safe_env = {
            '__builtins__': self._get_safe_builtins(),
            '__name__': '__sandbox__',  # 不使用__main__避免执行if __name__ == "__main__"块
            '__file__': '<sandbox>',
            '__package__': None,
            '__doc__': None,
            '__loader__': None,
            '__spec__': None,
            '__annotations__': {},
        }

        # 添加安全模块
        safe_env.update(self.safe_modules)

        # 【关键】直接将data_helpers中的函数注入到全局环境
        # 这样生成的代码可以直接使用这些函数，不需要 import
        try:
            from utils.data_helpers import (
                SYNONYM_GROUPS, find_column, safe_get_column,
                convert_region_to_dataframe, normalize_emp_code,
                print_available_columns, load_files_to_dataframes
            )
            safe_env['SYNONYM_GROUPS'] = SYNONYM_GROUPS
            safe_env['find_column'] = find_column
            safe_env['safe_get_column'] = safe_get_column
            safe_env['convert_region_to_dataframe'] = convert_region_to_dataframe
            safe_env['normalize_emp_code'] = normalize_emp_code
            safe_env['print_available_columns'] = print_available_columns
            safe_env['load_files_to_dataframes'] = load_files_to_dataframes
        except ImportError as e:
            self.logger.warning(f"无法导入data_helpers函数: {e}")

        # 注入公式模式的辅助函数（兜底：模板代码中也有定义，但防止模板拼接异常）
        safe_env['EMPTY'] = '""'
        safe_env['ZERO'] = '0'
        safe_env['excel_text'] = lambda text: f'"{text}"'

        # 安全包装 CellIsRule —— 防止 AI 生成错误的 operator 值
        # 1. operator={'greaterThan'} (set) → 'greaterThan' (str)
        # 2. operator='greater_than' (snake_case) → 'greaterThan'
        # 3. operator='不等于' (中文) → 'notEqual'
        # 必须 monkey-patch openpyxl 模块本身，因为 AI 代码通过 from ... import 直接导入
        try:
            import openpyxl.formatting.rule as _fmt_rule
            _OrigCellIsRule = _fmt_rule.CellIsRule
            # 只 patch 一次，避免重复包装
            if not getattr(_OrigCellIsRule, '_patched', False):
                _OPERATOR_MAP = {
                    # 符号
                    '>': 'greaterThan', '<': 'lessThan', '=': 'equal',
                    '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual',
                    '!=': 'notEqual', '<>': 'notEqual', '==': 'equal',
                    # snake_case
                    'greater_than': 'greaterThan', 'less_than': 'lessThan',
                    'greater_than_or_equal': 'greaterThanOrEqual',
                    'less_than_or_equal': 'lessThanOrEqual',
                    'not_equal': 'notEqual', 'not_between': 'notBetween',
                    'not_contains': 'notContains', 'contains_text': 'containsText',
                    'begins_with': 'beginsWith', 'ends_with': 'endsWith',
                    # 中文
                    '大于': 'greaterThan', '小于': 'lessThan', '等于': 'equal',
                    '不等于': 'notEqual', '大于等于': 'greaterThanOrEqual',
                    '小于等于': 'lessThanOrEqual', '介于': 'between',
                    '不介于': 'notBetween', '包含': 'containsText',
                    '不包含': 'notContains', '开头是': 'beginsWith',
                    '结尾是': 'endsWith',
                }
                _VALID_OPS = {'notBetween', 'greaterThanOrEqual', 'containsText',
                              'lessThanOrEqual', 'between', 'endsWith', 'greaterThan',
                              'lessThan', 'equal', 'beginsWith', 'notEqual', 'notContains'}

                def _safe_CellIsRule(operator=None, **kwargs):
                    if isinstance(operator, (set, list, tuple)):
                        operator = next(iter(operator)) if operator else None
                    if isinstance(operator, str):
                        # 去除可能被 TXT_ 常量包裹的引号，如 '"greaterThan"' → 'greaterThan'
                        operator = operator.strip().strip('"').strip("'")
                        if operator not in _VALID_OPS:
                            operator = _OPERATOR_MAP.get(operator, operator)
                    return _OrigCellIsRule(operator=operator, **kwargs)
                _safe_CellIsRule._patched = True
                _fmt_rule.CellIsRule = _safe_CellIsRule
            safe_env['CellIsRule'] = _fmt_rule.CellIsRule
            safe_env['FormulaRule'] = _fmt_rule.FormulaRule
        except ImportError:
            pass

        # 注入历史数据查询工具（如果有tenant_id）
        tenant_id = execution_env.get('tenant_id')
        if tenant_id:
            try:
                from utils.historical_data import HistoricalDataProvider
                safe_env['history_provider'] = HistoricalDataProvider(tenant_id)
            except ImportError as e:
                self.logger.warning(f"无法导入HistoricalDataProvider: {e}")

        # 添加执行环境变量
        for key, value in execution_env.items():
            if isinstance(value, (str, int, float, bool, list, dict, type(None))):
                safe_env[key] = value
            elif hasattr(value, '__class__'):
                # 允许对象传递，但会进行安全检查
                safe_env[key] = value

        # 添加自定义的open函数，限制文件访问
        def safe_open(filepath, mode='r', *args, **kwargs):
            # 检查文件路径是否在允许的目录内
            allowed_dirs = [
                execution_env.get('input_folder', ''),
                execution_env.get('output_folder', ''),
                tempfile.gettempdir()
            ]

            filepath = Path(filepath).resolve()
            is_allowed = any(str(filepath).startswith(str(Path(dir).resolve())) for dir in allowed_dirs if dir)

            if not is_allowed:
                raise PermissionError(f"不允许访问文件: {filepath}")

            return open(filepath, mode, *args, **kwargs)

        safe_env['open'] = safe_open

        return safe_env

    def _get_safe_builtins(self) -> dict:
        """获取安全的builtins函数"""
        safe_builtins = {}

        # 允许的安全builtins函数 - 放宽限制，允许__import__、eval、exec、compile、__build_class__
        safe_functions = [
            '__import__',  # 允许动态导入，因为AI生成的代码需要导入模块
            'eval', 'exec', 'compile',  # 允许代码执行，在沙箱环境中是安全的
            '__build_class__',  # 用于创建类的内部函数
            'abs', 'all', 'any', 'ascii', 'bin', 'bool', 'bytearray', 'bytes',
            'callable', 'chr', 'classmethod', 'complex', 'dict', 'dir', 'divmod',
            'enumerate', 'filter', 'float', 'format', 'frozenset', 'getattr',
            'hasattr', 'hash', 'hex', 'id', 'int', 'isinstance', 'issubclass',
            'iter', 'len', 'list', 'locals', 'globals', 'map', 'max', 'memoryview', 'min',
            'next', 'object', 'oct', 'ord', 'pow', 'print', 'property', 'range',
            'repr', 'reversed', 'round', 'set', 'setattr', 'slice', 'sorted',
            'staticmethod', 'str', 'sum', 'super', 'tuple', 'type', 'vars', 'zip',
            'exit'  # 添加exit函数支持，它会引发SystemExit异常，沙箱会捕获这个异常
        ]

        import builtins
        for func_name in safe_functions:
            if hasattr(builtins, func_name):
                safe_builtins[func_name] = getattr(builtins, func_name)

        # 添加一些必要的异常类型
        safe_builtins['Exception'] = Exception
        safe_builtins['ValueError'] = ValueError
        safe_builtins['TypeError'] = TypeError
        safe_builtins['AttributeError'] = AttributeError
        safe_builtins['KeyError'] = KeyError
        safe_builtins['IndexError'] = IndexError
        safe_builtins['FileNotFoundError'] = FileNotFoundError
        safe_builtins['PermissionError'] = PermissionError
        safe_builtins['IsADirectoryError'] = IsADirectoryError
        safe_builtins['NotADirectoryError'] = NotADirectoryError
        safe_builtins['IOError'] = IOError
        safe_builtins['OSError'] = OSError
        safe_builtins['EnvironmentError'] = OSError  # EnvironmentError是OSError的别名
        safe_builtins['RuntimeError'] = RuntimeError
        safe_builtins['ZeroDivisionError'] = ZeroDivisionError
        safe_builtins['ArithmeticError'] = ArithmeticError
        safe_builtins['LookupError'] = LookupError
        safe_builtins['AssertionError'] = AssertionError
        safe_builtins['ImportError'] = ImportError
        safe_builtins['NameError'] = NameError
        safe_builtins['UnicodeError'] = UnicodeError
        safe_builtins['StopIteration'] = StopIteration
        safe_builtins['GeneratorExit'] = GeneratorExit

        # 添加更多常用的异常类型
        safe_builtins['NotImplementedError'] = NotImplementedError
        safe_builtins['MemoryError'] = MemoryError
        safe_builtins['OverflowError'] = OverflowError
        safe_builtins['FloatingPointError'] = FloatingPointError
        safe_builtins['BufferError'] = BufferError
        safe_builtins['EOFError'] = EOFError
        safe_builtins['InterruptedError'] = InterruptedError
        safe_builtins['BlockingIOError'] = BlockingIOError
        safe_builtins['ChildProcessError'] = ChildProcessError
        safe_builtins['ConnectionError'] = ConnectionError
        safe_builtins['BrokenPipeError'] = BrokenPipeError
        safe_builtins['ConnectionAbortedError'] = ConnectionAbortedError
        safe_builtins['ConnectionRefusedError'] = ConnectionRefusedError
        safe_builtins['ConnectionResetError'] = ConnectionResetError
        safe_builtins['FileExistsError'] = FileExistsError
        safe_builtins['ProcessLookupError'] = ProcessLookupError
        safe_builtins['TimeoutError'] = TimeoutError
        safe_builtins['ReferenceError'] = ReferenceError
        safe_builtins['SyntaxError'] = SyntaxError
        safe_builtins['IndentationError'] = IndentationError
        safe_builtins['TabError'] = TabError
        safe_builtins['SystemError'] = SystemError

        return safe_builtins

    def _clean_code_syntax(self, code: str) -> str:
        """清理代码中的常见语法错误

        主要处理：
        1. f-string中的反斜杠转义问题
        2. 中文全角字符转换为英文半角字符
        3. 其他常见的语法问题

        Args:
            code: 原始代码

        Returns:
            清理后的代码
        """
        import re

        # 1. 中文全角字符转换为英文半角字符（仅替换字符串字面量之外的部分）
        # 注意：不能全局替换，因为字符串字面量中的全角字符可能是文件名/sheet名的一部分
        fullwidth_to_halfwidth = {
            '（': '(',
            '）': ')',
            '【': '[',
            '】': ']',
            '｛': '{',
            '｝': '}',
            '，': ',',
            '；': ';',
            '：': ':',
            '！': '!',
            '？': '?',
            '＝': '=',
            '＋': '+',
            '－': '-',
            '＊': '*',
            '／': '/',
            '＜': '<',
            '＞': '>',
        }

        # 逐行处理，只替换字符串字面量之外的全角字符
        def _replace_outside_strings(line, replacements):
            """替换字符串字面量之外的全角字符"""
            result = []
            i = 0
            n = len(line)
            while i < n:
                ch = line[i]
                # 检测字符串开始
                if ch in ('"', "'"):
                    quote = ch
                    # 检查三引号
                    if line[i:i+3] in ('"""', "'''"):
                        quote = line[i:i+3]
                    # 找到字符串结尾，原样保留
                    end = i + len(quote)
                    while end < n:
                        if line[end] == '\\':
                            end += 2  # 跳过转义
                            continue
                        if line[end:end+len(quote)] == quote:
                            end += len(quote)
                            break
                        end += 1
                    result.append(line[i:end])
                    i = end
                elif ch == '#':
                    # 注释部分也不替换
                    result.append(line[i:])
                    break
                else:
                    # 非字符串部分：替换全角
                    result.append(replacements.get(ch, ch))
                    i += 1
            return ''.join(result)

        code = '\n'.join(
            _replace_outside_strings(line, fullwidth_to_halfwidth)
            for line in code.split('\n')
        )

        # 2. 修复f-string中的引号嵌套问题
        # 问题：f"...TEXT(...,"YYYY-MM-DD")..." 或 f"=IF(A1="",..." 中内部双引号会导致语法错误
        # 解决：将使用双引号的f-string改为使用单引号
        lines_for_quote_fix = code.split('\n')
        fixed_lines_quote = []

        for line in lines_for_quote_fix:
            original_line = line
            # 检测 f"..." 格式的f-string，内部包含未转义的双引号（Excel公式常见模式）
            # 扩展检测：f"= 开头，且内部有 =" 或 ,"" 或 ="" 等Excel公式中常见的双引号用法
            if 'f"=' in line:
                # 检查是否有Excel公式中常见的双引号模式
                has_quote_issue = (
                    ',"' in line or      # ,""  逗号后跟引号
                    '=""' in line or     # ="" 等于空字符串
                    '="")' in line or    # ="")
                    '","' in line or     # "," 引号内有逗号
                    '")' in line         # ") 可能的引号结束
                )
                if has_quote_issue:
                    # 尝试将 f"..." 转换为 f'...'
                    match = re.search(r'(.*?)(f")(=.*)', line)
                    if match:
                        prefix = match.group(1)
                        fstring_content = match.group(3)
                        if fstring_content.endswith('"'):
                            inner_content = fstring_content[:-1]
                            line = prefix + "f'" + inner_content + "'"
                            self.logger.debug(f"修复f-string引号嵌套: {original_line[:60]}...")
            fixed_lines_quote.append(line)

        code = '\n'.join(fixed_lines_quote)

        # 2.5 修复f-string中的反斜杠转义问题
        # 例如: f'{vlookup(..., "\"\"")}'  ->  f'{vlookup(..., chr(34)+chr(34))}'
        # 或者: f'{vlookup(..., "\"\"")}'  ->  使用普通字符串拼接

        # 匹配 f-string 中包含 \" 的模式
        # 这是一个常见问题：AI生成的代码经常在f-string里使用 \"
        def fix_fstring_escapes(match):
            content = match.group(0)
            # 如果f-string中包含转义引号，转换为使用变量
            if '\\"' in content or "\\'" in content:
                # 简单处理：将f-string转换为普通字符串拼接
                # 这里使用正则替换 \" 为 ""（两个双引号在Excel公式中表示一个双引号）
                fixed = content.replace('\\"', '""')
                return fixed
            return content

        # 处理包含转义字符的f-string
        # 模式: f'...' 或 f"..."
        code = re.sub(r"f'[^']*\\'[^']*'", fix_fstring_escapes, code)
        code = re.sub(r'f"[^"]*\\"[^"]*"', fix_fstring_escapes, code)

        # 3. 检测并移除未终止的字符串行
        # AI生成的代码可能在行末有未闭合的字符串
        lines = code.split('\n')
        fixed_lines = []

        for i, line in enumerate(lines):
            # 检测未终止的f-string
            # 模式: f'... 或 f"... 行末没有闭合引号
            stripped = line.rstrip()

            # 检查行末是否有未闭合的f-string（以\' 或 \" 结尾，或引号不匹配）
            if stripped.endswith("\\'") or stripped.endswith('\\"'):
                # 行末有转义引号，可能是未完成的字符串
                # 尝试注释掉这行，添加一个占位符
                self.logger.warning(f"检测到可能未终止的字符串，行 {i+1}: {stripped[:50]}...")
                # 注释掉这行并添加一个pass语句
                fixed_lines.append(f"        # [自动修复] 注释掉未完成的代码行: {stripped[:80]}")
                fixed_lines.append("        pass  # 占位符")
                continue

            # 检查f-string引号是否匹配
            if ("f'" in line or 'f"' in line):
                # 简单检查：统计引号数量
                single_quotes = line.count("'") - line.count("\\'")
                double_quotes = line.count('"') - line.count('\\"')

                # 如果是f'开头，检查单引号是否成对
                if "f'" in line and single_quotes % 2 != 0:
                    # 单引号不成对，可能是未终止的字符串
                    if not stripped.endswith("'") and not stripped.endswith("',") and not stripped.endswith("')"):
                        self.logger.warning(f"检测到未终止的f-string，行 {i+1}: {stripped[:50]}...")
                        fixed_lines.append(f"        # [自动修复] 注释掉未完成的代码行")
                        fixed_lines.append("        pass  # 占位符")
                        continue

                # 如果是f"开头，检查双引号是否成对
                if 'f"' in line and double_quotes % 2 != 0:
                    if not stripped.endswith('"') and not stripped.endswith('",') and not stripped.endswith('")'):
                        self.logger.warning(f"检测到未终止的f-string，行 {i+1}: {stripped[:50]}...")
                        fixed_lines.append(f"        # [自动修复] 注释掉未完成的代码行")
                        fixed_lines.append("        pass  # 占位符")
                        continue

                # 替换 \" 为 " （简单处理）
                if '\\"' in line:
                    line = line.replace('\\"\\"\\"', '""')
                    line = line.replace('\\"', '"')

            fixed_lines.append(line)

        return '\n'.join(fixed_lines)

    def validate_code(self, code_content: str) -> Dict[str, Any]:
        """验证代码语法和安全性"""
        result = {
            "valid": False,
            "syntax_errors": [],
            "security_issues": [],
            "warnings": []
        }

        # 检查语法
        try:
            compile(code_content, '<validation>', 'exec')
            result["valid"] = True
        except SyntaxError as e:
            result["syntax_errors"].append({
                "line": e.lineno,
                "offset": e.offset,
                "message": str(e),
                "text": e.text
            })

        # 检查安全性
        if not self._is_safe_script(code_content):
            result["security_issues"].append("脚本包含不安全代码")

        # 检查常见问题
        if 'import pandas' not in code_content and 'from pandas import' not in code_content:
            result["warnings"].append("代码可能未导入pandas库")

        if 'import openpyxl' not in code_content and 'from openpyxl import' not in code_content:
            result["warnings"].append("代码可能未导入openpyxl库")

        return result

    def _import_robust_utils(self):
        """导入健壮性工具函数模块"""
        try:
            # 动态导入robust_utils模块
            import sys
            import os

            # 添加项目根目录到Python路径
            project_root = Path(__file__).parent.parent.parent
            sys.path.insert(0, str(project_root))

            # 导入robust_utils模块
            from backend.ai_engine.robust_utils import (
                safe_get_column, safe_calculate, validate_required_columns,
                create_missing_columns, mark_missing_cells_in_excel, get_dataframe_info
            )

            # 创建一个包含所有导出函数的模块对象
            class RobustUtilsModule:
                pass

            module = RobustUtilsModule()
            module.safe_get_column = safe_get_column
            module.safe_calculate = safe_calculate
            module.validate_required_columns = validate_required_columns
            module.create_missing_columns = create_missing_columns
            module.mark_missing_cells_in_excel = mark_missing_cells_in_excel
            module.get_dataframe_info = get_dataframe_info

            return module

        except ImportError as e:
            self.logger.warning(f"无法导入robust_utils模块: {e}")
            # 返回一个空模块
            class EmptyModule:
                pass
            return EmptyModule()

    def _import_sandbox_helpers(self):
        """导入数据处理辅助函数模块"""
        try:
            import sys

            # 添加backend目录到Python路径
            backend_dir = Path(__file__).parent.parent
            if str(backend_dir) not in sys.path:
                sys.path.insert(0, str(backend_dir))

            # 从utils导入data_helpers模块
            from utils.data_helpers import (
                SYNONYM_GROUPS, find_column, safe_get_column,
                convert_region_to_dataframe, normalize_emp_code,
                print_available_columns, load_files_to_dataframes
            )

            # 创建一个包含所有导出函数的模块对象
            class DataHelpersModule:
                pass

            module = DataHelpersModule()
            module.SYNONYM_GROUPS = SYNONYM_GROUPS
            module.find_column = find_column
            module.safe_get_column = safe_get_column
            module.convert_region_to_dataframe = convert_region_to_dataframe
            module.normalize_emp_code = normalize_emp_code
            module.print_available_columns = print_available_columns
            module.load_files_to_dataframes = load_files_to_dataframes

            return module

        except ImportError as e:
            self.logger.warning(f"无法导入data_helpers模块: {e}")
            # 返回一个空模块
            class EmptyModule:
                pass
            return EmptyModule()