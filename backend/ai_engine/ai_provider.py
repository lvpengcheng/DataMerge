"""
AI提供者模块 - 支持多种AI服务
"""

import os
import json
import re
import logging
from abc import ABC, abstractmethod
from typing import Dict, List, Any, Optional
import requests as http_requests
from backend.utils.indentation_fixer import IndentationFixer

logger = logging.getLogger(__name__)


class BaseAIProvider(ABC):
    """AI提供者基类"""

    # 主动分段输出标记：AI在每段末尾输出此标记，收到"继续"后输出下一段
    CONTINUATION_MARKER = "# ---NEXT_PART---"

    def __init__(self):
        self.last_raw_response = None
        self.last_extracted_code = None
        self._session = http_requests.Session()
        self._default_timeout = 300  # 5分钟，代码生成可能较慢
        self._indent_fixer = IndentationFixer()

    def _request(self, method: str, url: str, headers: dict, json_body: dict,
                 timeout: int = None, stream: bool = False):
        """统一 HTTP 请求，带错误处理"""
        import logging
        logger = logging.getLogger(__name__)
        timeout = timeout or self._default_timeout
        try:
            resp = self._session.request(
                method=method,
                url=url,
                headers=headers,
                json=json_body,
                timeout=timeout,
                stream=stream
            )
            resp.raise_for_status()
            return resp
        except http_requests.exceptions.Timeout:
            raise TimeoutError(f"API请求超时({timeout}s): {url}")
        except http_requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response else "unknown"
            body = e.response.text[:500] if e.response else ""
            if status == 429:
                raise RuntimeError(f"触发限流(429): {body}")
            logger.error(f"HTTP {status} 错误: {body}")
            raise RuntimeError(f"HTTP {status} 错误: {body}")
        except http_requests.exceptions.ConnectionError as e:
            raise ConnectionError(f"连接失败: {url} - {e}")

    @staticmethod
    def _parse_sse_lines(response):
        """解析 OpenAI 格式的 SSE 流，yield 每条 data 字符串"""
        response.encoding = "utf-8"
        for line_raw in response.iter_lines(decode_unicode=True):
            if not line_raw:
                continue
            line = line_raw.strip()
            if line.startswith("data: "):
                yield line[6:]

    @staticmethod
    def _parse_anthropic_sse(response):
        """解析 Anthropic 格式的 SSE 流，yield (event_type, data_dict) 元组"""
        response.encoding = "utf-8"
        event_type = None
        for line_raw in response.iter_lines(decode_unicode=True):
            if not line_raw:
                event_type = None
                continue
            line = line_raw.strip()
            if line.startswith("event: "):
                event_type = line[7:]
            elif line.startswith("data: "):
                data = json.loads(line[6:])
                yield (event_type, data)

    @abstractmethod
    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码"""
        pass

    @abstractmethod
    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口"""
        pass

    @abstractmethod
    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        pass

    def chat_stream(self, messages: List[Dict[str, str]], chunk_callback: callable = None, **kwargs) -> str:
        """流式对话接口 - 逐块输出并返回完整内容

        Args:
            messages: OpenAI 格式消息列表 [{"role":"system/user/assistant","content":"..."}]
            chunk_callback: 每个文本块的回调函数
        Returns:
            完整的回复文本
        """
        # 默认回退到非流式 chat
        result = self.chat(messages, **kwargs)
        if chunk_callback:
            chunk_callback(result)
        return result

    def _is_code_complete(self, code: str) -> bool:
        """检测代码是否完整（不仅依赖API的stop信号，还从代码本身判断）

        检测维度：
        1. AST语法解析是否通过
        2. 括号/引号是否闭合
        3. 是否以不完整的语句结尾（如 def/if/for 后无body）
        4. 代码块标记是否闭合（```python 无 ```）

        Returns:
            True=代码完整, False=代码被截断需要续写
        """
        import ast
        import logging
        logger = logging.getLogger(__name__)

        if not code or len(code.strip()) < 50:
            return True  # 太短的代码不做判断

        code = code.strip()

        # 1. AST语法解析
        try:
            ast.parse(code)
        except SyntaxError as e:
            logger.info(f"代码完整性检测: AST解析失败 - {e}")
            return False

        # 2. 括号闭合检测
        bracket_pairs = {'(': ')', '[': ']', '{': '}'}
        stack = []
        in_string = False
        string_char = None
        i = 0
        while i < len(code):
            ch = code[i]
            # 处理三引号字符串
            if not in_string and i + 2 < len(code) and code[i:i+3] in ('"""', "'''"):
                in_string = True
                string_char = code[i:i+3]
                i += 3
                continue
            if in_string and len(string_char) == 3 and i + 2 < len(code) and code[i:i+3] == string_char:
                in_string = False
                string_char = None
                i += 3
                continue
            # 处理单引号/双引号字符串
            if not in_string and ch in ('"', "'") and (i == 0 or code[i-1] != '\\'):
                in_string = True
                string_char = ch
                i += 1
                continue
            if in_string and len(string_char) == 1 and ch == string_char and (i == 0 or code[i-1] != '\\'):
                in_string = False
                string_char = None
                i += 1
                continue
            # 处理注释
            if not in_string and ch == '#':
                # 跳到行尾
                while i < len(code) and code[i] != '\n':
                    i += 1
                continue
            if not in_string:
                if ch in bracket_pairs:
                    stack.append(bracket_pairs[ch])
                elif ch in bracket_pairs.values():
                    if stack and stack[-1] == ch:
                        stack.pop()
                    # 不匹配的右括号不算截断，可能是语法错误
            i += 1

        if in_string:
            logger.info("代码完整性检测: 字符串未闭合")
            return False
        if stack:
            logger.info(f"代码完整性检测: 括号未闭合，缺少: {stack}")
            return False

        # 3. 检查最后非空行是否是不完整的语句
        lines = code.split('\n')
        last_meaningful_line = ""
        for line in reversed(lines):
            stripped = line.strip()
            if stripped and not stripped.startswith('#'):
                last_meaningful_line = stripped
                break

        incomplete_endings = [
            # 冒号结尾但下面没有body（AST已经检查过了，这里做双重保险）
            # 赋值号结尾
            '=', '+=', '-=', '*=', '/=',
            # 运算符结尾
            '+', '-', '*', '/', '%', '|', '&', '^',
            # 逗号结尾（参数列表未完成）
            ',',
            # 反斜杠续行
            '\\',
        ]

        for ending in incomplete_endings:
            if last_meaningful_line.endswith(ending):
                logger.info(f"代码完整性检测: 最后一行以 '{ending}' 结尾，代码可能不完整")
                return False

        return True

    def _build_continuation_prompt(self, original_prompt: str, incomplete_code: str) -> str:
        """构建续写提示词 — 精确指定截断位置，减少重复生成

        Args:
            original_prompt: 原始提示词
            incomplete_code: 当前未完成的代码

        Returns:
            续写用的提示词
        """
        # 取最后20行作为上下文锚点
        lines = incomplete_code.strip().split('\n')
        total_lines = len(lines)
        tail_lines = lines[-20:] if total_lines > 20 else lines
        tail_snippet = '\n'.join(tail_lines)

        # 检测最后一行的缩进级别，帮助AI保持正确缩进
        last_indent = ""
        for line in reversed(lines):
            if line.strip():
                last_indent = line[:len(line) - len(line.lstrip())]
                break
        indent_hint = f"当前缩进级别为 {len(last_indent)} 个空格。" if last_indent else "当前在顶级作用域。"

        return (
            f"以下是原始需求：\n"
            f"---原始需求开始---\n{original_prompt}\n---原始需求结束---\n\n"
            f"你之前根据上述需求已经生成了 {total_lines} 行代码，但在中途被截断了。\n"
            f"已生成代码的最后部分如下：\n"
            f"```python\n{tail_snippet}\n```\n"
            f"（以上是第 {total_lines - len(tail_lines) + 1} 行到第 {total_lines} 行）\n\n"
            f"{indent_hint}\n\n"
            f"请严格从第 {total_lines + 1} 行开始，继续完成剩余的代码逻辑。\n"
            f"重要要求：\n"
            f"1. 绝对不要重复上面已有的任何代码，包括import、函数定义、变量赋值等\n"
            f"2. 直接从截断处接续，确保和上面的代码无缝拼接\n"
            f"3. 如果截断发生在函数内部，直接续写函数体的剩余部分，保持相同的缩进级别\n"
            f"4. 只返回续写的Python代码，不要包含解释文本\n"
            f"5. 续写代码的第一行缩进必须与截断处的上下文一致"
        )

    def _build_inline_continuation_msg(self, current_response: str) -> str:
        """构建Phase1续写消息（finish_reason=length / stop_reason=max_tokens 时使用）

        包含最后几行代码作为锚点，让AI明确知道从哪里接续，减少重复生成。
        """
        lines = current_response.strip().split('\n')
        tail = lines[-10:] if len(lines) > 10 else lines
        tail_snippet = '\n'.join(tail)
        return (
            f"代码在以下位置被截断，请严格从截断处的下一行继续：\n"
            f"```\n{tail_snippet}\n```\n"
            f"要求：\n"
            f"1. 直接从截断处的下一行开始输出，保持正确的缩进层级\n"
            f"2. 绝对不要重复上面已有的任何代码，包括import、函数定义、变量赋值\n"
            f"3. 如果截断发生在函数/循环/条件内部，续写时保持相同的缩进级别\n"
            f"4. 只返回续写部分的Python代码，不要包含解释文本"
        )

    def _find_overlap_point(self, existing_code: str, new_code: str) -> int:
        """在新代码中找到与已有代码重叠的结束位置（基于序列匹配）

        使用已有代码末尾的连续行与新代码开头的连续行做序列比对，
        找到最长的重叠前缀，返回新代码中非重叠部分的起始行索引。

        Args:
            existing_code: 已有的代码
            new_code: 续传返回的新代码

        Returns:
            新代码中非重叠部分的起始行索引（对应raw lines）
        """
        existing_raw = existing_code.strip().split('\n')
        new_raw = new_code.strip().split('\n')

        if not existing_raw or not new_raw:
            return 0

        # strip每行用于比较，但索引对应raw lines
        tail_size = min(50, len(existing_raw))
        existing_tail = [l.strip() for l in existing_raw[-tail_size:]]
        new_stripped = [l.strip() for l in new_raw]

        # 策略1: 寻找最长的连续序列匹配
        # 检查 existing_tail 的某个后缀是否等于 new_stripped 的等长前缀
        best_overlap = 0
        for start in range(len(existing_tail)):
            suffix = existing_tail[start:]
            match_len = 0
            for j in range(min(len(suffix), len(new_stripped))):
                if suffix[j] == new_stripped[j]:
                    match_len = j + 1
                else:
                    break
            # 至少3行连续匹配才算有效重叠，防止单行巧合
            if match_len >= 3 and match_len > best_overlap:
                best_overlap = match_len

        if best_overlap > 0:
            return best_overlap

        # 策略2: 检查新代码开头是否有零散的重复行（import/def等）
        existing_tail_set = set(existing_tail)
        overlap = 0
        for i, line in enumerate(new_stripped):
            if not line or line.startswith('#'):
                overlap = i + 1
                continue
            if line in existing_tail_set:
                overlap = i + 1
            else:
                break

        return overlap

    def _merge_continuation_code(self, existing_code: str, new_code: str) -> str:
        """合并续传代码 — 智能去重，精确拼接

        1. 检测新代码与已有代码的重叠部分并跳过
        2. 去除重复的import语句
        3. 修正缩进对齐

        Args:
            existing_code: 已有的代码
            new_code: 续传返回的新代码

        Returns:
            合并后的代码
        """
        import logging
        logger = logging.getLogger(__name__)

        if not existing_code:
            return new_code
        if not new_code:
            return existing_code

        existing_lines = existing_code.strip().split('\n')
        new_raw_lines = new_code.strip().split('\n')

        # 1. 找到重叠结束点，跳过新代码中与已有代码重复的部分
        overlap_point = self._find_overlap_point(existing_code, new_code)
        if overlap_point > 0:
            logger.info(f"续传去重 - 跳过新代码前 {overlap_point} 行重叠内容")
        new_lines = new_raw_lines[overlap_point:]

        if not new_lines:
            logger.info("续传去重 - 新代码全部与已有代码重叠，无需合并")
            return existing_code

        # 2. 收集已有代码的 import 语句，去除新代码中重复的 import
        existing_imports = set()
        for line in existing_lines:
            stripped = line.strip()
            if stripped.startswith('import ') or stripped.startswith('from '):
                existing_imports.add(stripped)

        # 收集已有代码中的函数签名，用于检测重复函数定义
        existing_func_sigs = set()
        for line in existing_lines:
            stripped = line.strip()
            if stripped.startswith('def '):
                # 提取函数名
                func_sig = stripped.split('(')[0] if '(' in stripped else stripped
                existing_func_sigs.add(func_sig)

        filtered_lines = []
        skip_function_body = False
        for line in new_lines:
            stripped = line.strip()

            # 跳过重复的import
            if stripped.startswith('import ') or stripped.startswith('from '):
                if stripped in existing_imports:
                    logger.info(f"续传去重 - 跳过重复import: {stripped}")
                    continue

            # 检测重复的函数定义
            if stripped.startswith('def '):
                func_sig = stripped.split('(')[0] if '(' in stripped else stripped
                if func_sig in existing_func_sigs:
                    logger.info(f"续传去重 - 跳过重复函数定义: {func_sig}")
                    skip_function_body = True
                    continue
                else:
                    skip_function_body = False

            # 如果正在跳过重复函数的body
            if skip_function_body:
                if stripped and not stripped.startswith('#') and not line.startswith(' ') and not line.startswith('\t'):
                    # 遇到顶级代码，停止跳过
                    skip_function_body = False
                else:
                    continue

            filtered_lines.append(line)

        if not filtered_lines:
            logger.info("续传去重 - 过滤后无新代码")
            return existing_code

        # 3. 检测已有代码末尾的缩进级别
        # 跳过return语句，找到实际代码体的缩进（return通常缩进更浅）
        target_indent = ""
        for line in reversed(existing_lines):
            stripped = line.strip()
            if stripped and not stripped.startswith('return '):
                target_indent = line[:len(line) - len(line.lstrip())]
                break

        # 检测新代码的基准缩进
        new_base_indent = ""
        for line in filtered_lines:
            if line.strip():
                new_base_indent = line[:len(line) - len(line.lstrip())]
                break

        # 判断是否需要重新缩进（新代码是顶级代码则不需要）
        # 如果新代码的第一个非空行是顶级的（缩进为0），且已有代码末尾也是顶级的，直接拼接
        new_first_is_toplevel = (new_base_indent == "")
        existing_last_is_toplevel = (target_indent == "")

        if new_first_is_toplevel and existing_last_is_toplevel:
            # 都是顶级代码，直接拼接
            merged = existing_code.rstrip() + '\n\n' + '\n'.join(filtered_lines)
        else:
            # 需要对齐缩进
            reindented = []
            for line in filtered_lines:
                if not line.strip():
                    reindented.append('')
                    continue
                current_indent = line[:len(line) - len(line.lstrip())]
                content = line.lstrip()
                if len(new_base_indent) > 0 and current_indent.startswith(new_base_indent):
                    extra_indent = current_indent[len(new_base_indent):]
                    reindented.append(target_indent + extra_indent + content)
                else:
                    reindented.append(target_indent + content)
            merged = existing_code.rstrip() + '\n\n' + '\n'.join(reindented)

        logger.info(f"代码合并完成 (已有: {len(existing_code)}, 新增原始: {len(new_code)}, 去重后新增: {len(chr(10).join(filtered_lines))}, 合并后: {len(merged)})")
        return merged

    def extract_python_code(self, ai_response: str) -> str:
        """
        从AI响应中提取Python代码

        AI生成的响应可能包含：
        1. 纯代码
        2. 代码块（```python ... ```）
        3. 解释 + 代码
        4. 其他格式

        Args:
            ai_response: AI返回的完整响应文本

        Returns:
            提取出的纯Python代码
        """
        if not ai_response:
            return ""

        # 尝试匹配代码块
        code_block_pattern = r'```(?:python)?\s*(.*?)\s*```'
        code_blocks = re.findall(code_block_pattern, ai_response, re.DOTALL)

        if code_blocks:
            # 如果有多个代码块，取第一个（通常是最主要的）
            return code_blocks[0].strip()

        # 如果没有代码块，尝试提取看起来像Python代码的部分
        # 查找以 import 或 def 或 class 开头的行
        lines = ai_response.split('\n')
        code_lines = []
        in_code_section = False

        for line in lines:
            stripped = line.strip()

            # 检测代码开始
            if (stripped.startswith('import ') or
                stripped.startswith('from ') or
                stripped.startswith('def ') or
                stripped.startswith('class ') or
                stripped.startswith('async def ') or
                stripped.startswith('@') or
                stripped.startswith('# ') or
                re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*\s*=', stripped) or
                re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*\s*\(', stripped)):
                in_code_section = True

            if in_code_section:
                # 跳过空行和纯注释行（除非在代码段中）
                if stripped and not stripped.startswith('#'):
                    code_lines.append(line)
                elif stripped.startswith('#') and code_lines:
                    # 保留代码中的注释
                    code_lines.append(line)

        if code_lines:
            return '\n'.join(code_lines).strip()

        # 如果以上方法都没提取到，返回原始文本（可能是纯代码）
        return ai_response.strip()

    def validate_and_fix_code_format(self, code: str) -> str:
        """
        验证和修复Python代码格式

        Args:
            code: 原始Python代码

        Returns:
            修复后的Python代码
        """
        if not code:
            return code

        import ast
        import logging
        logger = logging.getLogger(__name__)

        # 首先修复明显的错误路径（AI可能从代码库中学到的错误内容）
        code = self._fix_invalid_paths(code)

        # 修复f-string引号嵌套冲突
        code = self._fix_fstring_quotes(code)

        # 尝试解析代码，如果成功则不需要修复
        try:
            ast.parse(code)
            logger.info("代码语法验证通过，无需修复")
            return code
        except SyntaxError as e:
            logger.warning(f"检测到语法错误: {e}，尝试修复缩进")

        # 尝试修复缩进（使用统一的缩进修复器）
        fixed_code = self._indent_fixer.fix_general(code)

        # 再次验证修复后的代码
        try:
            ast.parse(fixed_code)
            logger.info("缩进修复成功，代码语法验证通过")
            return fixed_code
        except SyntaxError as e:
            logger.warning(f"通用缩进修复后仍有语法错误: {e}，尝试沙箱级管线修复")

        # 用更强的沙箱级管线修复（基于原始代码，避免 fix_general 引入的破坏）
        fixed_code = self._indent_fixer.fix_sandbox_pipeline(code)
        try:
            ast.parse(fixed_code)
            logger.info("沙箱级管线缩进修复成功")
            return fixed_code
        except SyntaxError as e:
            logger.warning(f"沙箱级管线修复后仍有语法错误: {e}，尝试截断修复")

        # 如果仍然失败，尝试检测并修复截断问题
        fixed_code = self._detect_and_fix_truncation(fixed_code)

        # 确保代码以换行符结束
        if fixed_code and not fixed_code.endswith('\n'):
            fixed_code += '\n'

        return fixed_code

    def _fix_invalid_paths(self, code: str) -> str:
        """
        修复AI生成代码中的无效路径

        AI有时会从代码库中学到一些无关的字符串并错误地当作路径使用，
        这个方法检测并替换这些明显的错误。

        Args:
            code: 原始代码

        Returns:
            修复后的代码
        """
        import re
        import logging
        logger = logging.getLogger(__name__)

        # 定义需要替换的错误路径模式
        # 格式: (错误模式, 正确替换)
        invalid_path_patterns = [
            # uvicorn启动参数被误用为路径
            (r"['\"]backend\.app\.main:app['\"]", "input_folder"),
            (r"os\.listdir\(['\"]backend\.app\.main:app['\"]\)", "os.listdir(input_folder)"),
            (r"os\.path\.join\(['\"]backend\.app\.main:app['\"]", "os.path.join(input_folder"),
            # 其他常见的错误路径
            (r"['\"]\.\/input['\"]", "input_folder"),
            (r"os\.listdir\(['\"]\.\/input['\"]\)", "os.listdir(input_folder)"),
            (r"os\.listdir\(['\"]input['\"]\)", "os.listdir(input_folder)"),
            (r"os\.listdir\(['\"]\.['\"]\)", "os.listdir(input_folder)"),
        ]

        fixed_code = code
        for pattern, replacement in invalid_path_patterns:
            if re.search(pattern, fixed_code):
                logger.warning(f"检测到无效路径模式: {pattern}，替换为: {replacement}")
                fixed_code = re.sub(pattern, replacement, fixed_code)

        return fixed_code

    def _fix_fstring_quotes(self, code: str) -> str:
        """修复f-string中引号嵌套冲突

        AI常见错误模式：
        1. f'...'{var}'...' — 单引号f-string内部用单引号 → 改外层为双引号
        2. f"..."-"...TEXT(...,"00")" — 双引号f-string内部用双引号 → 改外层为单引号
        3. 不含{变量}的f-string → 去掉f前缀，改为普通字符串
        """
        import ast
        import logging
        logger = logging.getLogger(__name__)

        lines = code.split('\n')
        fixed_lines = []
        fix_count = 0

        for line in lines:
            stripped = line.strip()

            # 只处理包含f-string的行
            if "f'" not in stripped and 'f"' not in stripped:
                fixed_lines.append(line)
                continue

            # 先检查这行是否有语法错误
            try:
                ast.parse(stripped)
                fixed_lines.append(line)
                continue
            except SyntaxError:
                pass

            indent = line[:len(line) - len(line.lstrip())]
            fixed = stripped

            # 策略1: f"..." 内含双引号 → 改为 f'...'
            # 例: f"=参数!$B$2&"-"&TEXT(参数!$B$3,"00")"
            # →   f'=参数!$B$2&"-"&TEXT(参数!$B$3,"00")'
            if 'f"' in fixed:
                fixed = self._try_swap_to_single_quote_fstring(fixed)
                try:
                    ast.parse(fixed)
                    fix_count += 1
                    fixed_lines.append(indent + fixed)
                    continue
                except SyntaxError:
                    fixed = stripped  # 回退

            # 策略2: f'...' 内含单引号 → 改为 f"..."
            # 例: f'=VLOOKUP(A2,'{var}'!$A:$B,2,FALSE)'
            # →   f"=VLOOKUP(A2,'{var}'!$A:$B,2,FALSE)"
            if "f'" in fixed:
                fixed = self._try_swap_to_double_quote_fstring(fixed)
                try:
                    ast.parse(fixed)
                    fix_count += 1
                    fixed_lines.append(indent + fixed)
                    continue
                except SyntaxError:
                    fixed = stripped  # 回退

            # 策略3: f"..."内含双引号且swap失败（因为也含单引号）→ 转义内部双引号
            # 例: f"=IFERROR(VLOOKUP(K{r},'Sheet1'!$A:$J,8,FALSE),"")"
            # →   f"=IFERROR(VLOOKUP(K{r},'Sheet1'!$A:$J,8,FALSE),\"\")"
            if 'f"' in stripped:
                fixed = self._try_escape_inner_double_quotes(stripped)
                try:
                    ast.parse(fixed)
                    fix_count += 1
                    fixed_lines.append(indent + fixed)
                    continue
                except SyntaxError:
                    fixed = stripped  # 回退

            # 策略4: 如果f-string里没有{变量}，去掉f前缀
            fixed = re.sub(r'\bf(["\'])', r'\1', stripped)
            if fixed != stripped:
                try:
                    ast.parse(fixed)
                    fix_count += 1
                    fixed_lines.append(indent + fixed)
                    continue
                except SyntaxError:
                    pass

            # 都修不了，保留原样
            fixed_lines.append(line)

        if fix_count > 0:
            logger.info(f"修复了 {fix_count} 处f-string引号冲突")

        return '\n'.join(fixed_lines)

    def _try_swap_to_single_quote_fstring(self, line: str) -> str:
        """将 f"..." 改为 f'...'，用于内部含双引号的情况。
        支持一行多个f-string，从右到左逐个尝试。"""
        import ast

        # 找到所有 f" 的位置
        positions = []
        idx = 0
        while True:
            pos = line.find('f"', idx)
            if pos < 0:
                break
            positions.append(pos)
            idx = pos + 2

        if not positions:
            return line

        # 从右到左尝试修复每个 f-string
        for pos in reversed(positions):
            search_start = pos + 2
            close_pos = line.rfind('"', search_start)
            if close_pos <= search_start:
                continue
            content = line[pos + 2:close_pos]
            if "'" in content and "{" not in content:
                continue  # 内部有单引号且无变量，换过去也会冲突
            candidate = line[:pos] + "f'" + content + "'" + line[close_pos + 1:]
            try:
                ast.parse(candidate)
                return candidate
            except SyntaxError:
                continue

        return line

    def _try_swap_to_double_quote_fstring(self, line: str) -> str:
        """将 f'...' 改为 f"..."，用于内部含单引号的情况。
        支持一行多个f-string，从右到左逐个尝试。"""
        import ast

        # 找到所有 f' 的位置
        positions = []
        idx = 0
        while True:
            pos = line.find("f'", idx)
            if pos < 0:
                break
            positions.append(pos)
            idx = pos + 2

        if not positions:
            return line

        # 从右到左尝试修复每个 f-string
        for pos in reversed(positions):
            search_start = pos + 2
            close_pos = line.rfind("'", search_start)
            if close_pos <= search_start:
                continue
            content = line[pos + 2:close_pos]
            escaped_content = content.replace('"', '\\"')
            candidate = line[:pos] + 'f"' + escaped_content + '"' + line[close_pos + 1:]
            try:
                ast.parse(candidate)
                return candidate
            except SyntaxError:
                continue

        return line

    def _try_escape_inner_double_quotes(self, line: str) -> str:
        """保持f"..."外层双引号，转义内部的双引号为\\"

        用于内部同时含单引号和双引号的情况（swap策略都失败时）。
        例: f"=IFERROR(VLOOKUP(K{r},'Sheet1'!$A:$J,8,FALSE),"")"
        →   f"=IFERROR(VLOOKUP(K{r},'Sheet1'!$A:$J,8,FALSE),\\"\\")\"
        """
        import ast

        # 找到所有 f" 的位置
        positions = []
        idx = 0
        while True:
            pos = line.find('f"', idx)
            if pos < 0:
                break
            positions.append(pos)
            idx = pos + 2

        if not positions:
            return line

        # 从右到左尝试修复每个 f-string
        for pos in reversed(positions):
            search_start = pos + 2
            # 找到这个f-string真正的结束引号比较复杂，
            # 因为内部的""会干扰。用贪心策略：取到行尾最后一个"
            close_pos = line.rfind('"', search_start)
            if close_pos <= search_start:
                continue

            content = line[pos + 2:close_pos]
            # 把内部所有未转义的双引号转义
            escaped_content = content.replace('\\"', '\x00').replace('"', '\\"').replace('\x00', '\\"')
            candidate = line[:pos] + 'f"' + escaped_content + '"' + line[close_pos + 1:]
            try:
                ast.parse(candidate)
                return candidate
            except SyntaxError:
                continue

        return line

    def _detect_and_fix_truncation(self, code: str) -> str:
        """
        检测并修复代码截断问题

        Args:
            code: 原始代码

        Returns:
            修复后的代码
        """
        if not code:
            return code

        lines = code.split('\n')
        if not lines:
            return code

        import re
        import ast

        # 常见的截断模式检测
        truncation_patterns = [
            # 未闭合的字符串
            (r'\"\s*$', '双引号未闭合'),  # 双引号未闭合
            (r"\'\s*$", '单引号未闭合'),  # 单引号未闭合
            (r'\"\"\"\s*$', '三双引号未闭合'),  # 三双引号未闭合
            (r"\'\'\'\s*$", "三单引号未闭合"),  # 三单引号未闭合

            # 未闭合的括号
            (r'\(\s*$', '圆括号未闭合'),  # 圆括号未闭合
            (r'\[\s*$', '方括号未闭合'),  # 方括号未闭合
            (r'\{\s*$', '花括号未闭合'),  # 花括号未闭合

            # 未完成的语句
            (r'print\(\s*$', 'print语句未完成'),  # print语句未完成
            (r'def\s+\w+\s*\(\s*$', '函数定义未完成'),  # 函数定义未完成
            (r'class\s+\w+\s*\(\s*$', '类定义未完成'),  # 类定义未完成
            (r'if\s+.*:\s*$', 'if语句未完成'),  # if语句未完成
            (r'for\s+.*:\s*$', 'for循环未完成'),  # for循环未完成
            (r'while\s+.*:\s*$', 'while循环未完成'),  # while循环未完成
            (r'try:\s*$', 'try语句未完成'),  # try语句未完成
            (r'except\s+.*:\s*$', 'except语句未完成'),  # except语句未完成
        ]

        last_line = lines[-1].strip()

        # 检查最后一行是否有明显的截断迹象
        truncation_detected = False
        for pattern, issue_type in truncation_patterns:
            if re.search(pattern, last_line):
                truncation_detected = True
                # 移除有问题的最后一行
                lines = lines[:-1]
                # 添加一个简单的结束标记
                lines.append(f'# [检测到截断: {issue_type}，已自动修复]')
                lines.append('pass')
                break

        # 如果没有检测到明显的截断模式，检查语法
        if not truncation_detected and last_line and not last_line.startswith('#') and not last_line.startswith('"""') and not last_line.startswith("'''"):
            # 尝试解析整个代码，检查语法
            try:
                ast.parse(code)
            except SyntaxError:
                # 语法错误，可能是截断
                # 尝试从后往前移除行，直到找到有效的语法
                for i in range(1, min(10, len(lines))):
                    test_lines = lines[:-i]
                    if test_lines:
                        test_code = '\n'.join(test_lines)
                        try:
                            ast.parse(test_code)
                            # 找到有效的语法点
                            lines = test_lines
                            lines.append('# [检测到语法错误，可能是截断，已移除无效代码]')
                            lines.append('pass')
                            truncation_detected = True
                            break
                        except SyntaxError:
                            continue

        # 如果检测到截断，记录日志
        if truncation_detected:
            import logging
            logger = logging.getLogger(__name__)
            joined_lines = '\n'.join(lines)
            logger.warning(f"检测到代码截断，已自动修复。原始代码长度: {len(code)}，修复后长度: {len(joined_lines)}")

        return '\n'.join(lines)


class OpenAIProvider(BaseAIProvider):
    """OpenAI提供者"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        self.api_key = config.get("api_key", os.getenv("OPENAI_API_KEY"))
        self.base_url = config.get("base_url", os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")).rstrip("/")
        self.model = config.get("model", "gpt-4")
        self.max_tokens = config.get("max_tokens", int(os.getenv("OPENAI_MAX_TOKENS", "8000")))
        self.timeout = int(config.get("timeout", os.getenv("OPENAI_TIMEOUT", "300")))

        from openai import OpenAI
        self._client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
            timeout=self.timeout,
        )

    def _openai_chat(self, messages, temperature=0.1, max_tokens=None, **kwargs):
        """非流式调用 OpenAI SDK，返回 (content, finish_reason)"""
        filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("stream",)}
        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens or self.max_tokens,
                **filtered_kwargs,
            )
            # 调试：检查响应类型
            if isinstance(response, str):
                logger.error(f"API返回了字符串而不是对象: {response[:500]}")
                raise ValueError(f"API返回格式错误，期望对象但得到字符串")

            choice = response.choices[0]
            content = choice.message.content or ""
            finish_reason = choice.finish_reason
            return content, finish_reason
        except Exception as e:
            logger.error(f"_openai_chat 调用失败: {e}, response type: {type(response) if 'response' in locals() else 'N/A'}")
            raise

    def _openai_chat_stream(self, messages, temperature=0.1, max_tokens=None, **kwargs):
        """流式调用 OpenAI SDK，yield (content_chunk, finish_reason)"""
        filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("stream",)}
        stream = self._client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens or self.max_tokens,
            stream=True,
            **filtered_kwargs,
        )
        finish_reason = None
        for chunk in stream:
            if chunk.choices:
                delta = chunk.choices[0].delta
                content = delta.content if delta else None
                fr = chunk.choices[0].finish_reason
                if fr:
                    finish_reason = fr
                if content:
                    yield content, finish_reason
        # 最后 yield 一个空内容带 finish_reason，确保调用方能拿到
        yield "", finish_reason

    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码（支持自动续写：finish_reason截断时循环续写，完成后再检查代码完整性）"""
        import logging
        logger = logging.getLogger(__name__)

        MAX_ITERATIONS = 10
        COMPLETENESS_RETRIES = 3
        SYSTEM_PROMPT = (
            "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
            "包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。"
            "特别注意根据业务场景选择合适的主键进行数据关联和计算。"
            "只返回Python代码，不要包含解释或其他文本。\n\n"
            "⚠️ 缩进纪律（必须严格遵守）：\n"
            "1. break退出for循环后，下一行代码必须回退到for语句的缩进级别\n"
            "2. if/elif/else块结束后，后续独立代码必须与if同级，禁止嵌套在else内部\n"
            "3. 各步骤注释（# === N. ===）必须全部在同一缩进级别\n\n"
            "重要：如果代码较长（超过150行），请主动分段输出。"
            "每段在逻辑完整的位置断开（如函数定义之间），"
            f"段末单独输出一行 {self.CONTINUATION_MARKER} 作为标记。"
            "收到'继续'后输出下一段。最后一段不需要标记。"
        )

        try:
            logger.info(f"OpenAIProvider: 开始生成代码，提示词长度: {len(prompt)}, max_tokens: {self.max_tokens}")

            messages = [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ]

            full_raw_response = ""

            # 阶段1: 基于 finish_reason 截断或主动分段标记的续写
            for iteration in range(MAX_ITERATIONS):
                logger.info(f"OpenAIProvider: 第 {iteration+1} 次调用API，模型: {self.model}")
                current_response, finish_reason = self._openai_chat(messages, max_tokens=self.max_tokens, **kwargs)
                full_raw_response += current_response

                logger.info(f"OpenAIProvider: 迭代 {iteration+1}: 本次长度={len(current_response)}, finish_reason={finish_reason}, 累计={len(full_raw_response)}")

                if len(full_raw_response) <= 1000:
                    logger.info(f"OpenAIProvider: 当前累计响应: {full_raw_response}")
                else:
                    logger.info(f"OpenAIProvider: 累计响应前500字符: {full_raw_response[:500]}")
                    logger.info(f"OpenAIProvider: 累计响应后500字符: {full_raw_response[-500:]}")

                # 检测主动分段标记
                has_marker = self.CONTINUATION_MARKER in current_response
                if has_marker:
                    full_raw_response = full_raw_response.replace(self.CONTINUATION_MARKER, "")
                    logger.info(f"OpenAIProvider: 检测到分段标记，发送'继续'获取下一段")
                    messages = messages + [
                        {"role": "assistant", "content": current_response},
                        {"role": "user", "content": "继续"}
                    ]
                    continue

                if finish_reason != "length":
                    break

                logger.info(f"OpenAIProvider: finish_reason=length，触发续写（已生成 {len(full_raw_response)} 字符）")
                messages = messages + [
                    {"role": "assistant", "content": current_response},
                    {"role": "user", "content": self._build_inline_continuation_msg(current_response)}
                ]
            else:
                logger.warning(f"OpenAIProvider: 达到最大截断续写次数 {MAX_ITERATIONS}，强制停止")

            # 阶段2: 提取代码后检查完整性，不完整则续写
            extracted_code = self.extract_python_code(full_raw_response)
            logger.info(f"OpenAIProvider: 提取的代码长度: {len(extracted_code)}")

            for retry in range(COMPLETENESS_RETRIES):
                if self._is_code_complete(extracted_code):
                    logger.info(f"OpenAIProvider: 代码完整性检查通过")
                    break

                logger.info(f"OpenAIProvider: 代码不完整，第 {retry+1} 次完整性续写")
                continuation_prompt = self._build_continuation_prompt(prompt, extracted_code)
                new_raw, _ = self._openai_chat(
                    [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": continuation_prompt}
                    ],
                    max_tokens=self.max_tokens, **kwargs
                )
                new_code = self.extract_python_code(new_raw)
                extracted_code = self._merge_continuation_code(extracted_code, new_code)
                logger.info(f"OpenAIProvider: 完整性续写合并后代码长度: {len(extracted_code)}")
            else:
                logger.warning(f"OpenAIProvider: 达到最大完整性续写次数 {COMPLETENESS_RETRIES}，使用当前代码")

            # 验证和修复代码格式
            fixed_code = self.validate_and_fix_code_format(extracted_code)
            if fixed_code != extracted_code:
                logger.info(f"OpenAIProvider: 代码格式已修复 (原始长度: {len(extracted_code)}, 修复后: {len(fixed_code)})")

            self.last_raw_response = full_raw_response
            self.last_extracted_code = fixed_code

            return fixed_code

        except Exception as e:
            logger.error(f"OpenAI API调用失败: {e}", exc_info=True)

            fallback_code = '''import pandas as pd
import os

def process_excel_files(input_folder: str, output_file: str):
    """处理Excel文件"""
    print(f"处理文件夹: {input_folder}")
    print(f"输出文件: {output_file}")
    return True

if __name__ == "__main__":
    process_excel_files("input", "output.xlsx")'''

            logger.info("OpenAIProvider: 返回备用代码")
            return self.extract_python_code(fallback_code)

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口"""
        content, _ = self._openai_chat(messages, **kwargs)
        return content

    def chat_stream(self, messages: List[Dict[str, str]], chunk_callback: callable = None, **kwargs) -> str:
        """流式对话接口"""
        full_content = ""
        for text_chunk, _fr in self._openai_chat_stream(messages, **kwargs):
            if text_chunk:
                full_content += text_chunk
                if chunk_callback:
                    chunk_callback(text_chunk)
        return full_content

    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        messages = [
            {"role": "user", "content": prompt}
        ]
        return self.chat(messages, **kwargs)

    def stream_generate_code(self, prompt: str, chunk_callback: callable = None, **kwargs):
        """流式生成代码（支持自动续写）

        Args:
            prompt: 提示词
            chunk_callback: 每个chunk的回调函数，接收chunk字符串
            **kwargs: 其他参数

        Yields:
            代码片段
        """
        import logging
        logger = logging.getLogger(__name__)

        MAX_ITERATIONS = 10
        COMPLETENESS_RETRIES = 3
        SYSTEM_PROMPT = (
            "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
            "包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。"
            "特别注意根据业务场景选择合适的主键进行数据关联和计算。"
            "只返回Python代码，不要包含解释或其他文本。\n\n"
            "⚠️ 缩进纪律（必须严格遵守）：\n"
            "1. break退出for循环后，下一行代码必须回退到for语句的缩进级别\n"
            "2. if/elif/else块结束后，后续独立代码必须与if同级，禁止嵌套在else内部\n"
            "3. 各步骤注释（# === N. ===）必须全部在同一缩进级别\n\n"
            "重要：如果代码较长（超过150行），请主动分段输出。"
            "每段在逻辑完整的位置断开（如函数定义之间），"
            f"段末单独输出一行 {self.CONTINUATION_MARKER} 作为标记。"
            "收到'继续'后输出下一段。最后一段不需要标记。"
        )

        try:
            logger.info(f"OpenAIProvider: 开始流式生成代码，提示词长度: {len(prompt)}")

            messages = [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ]

            full_raw_response = ""

            # 阶段1: 基于 finish_reason 截断或主动分段标记的续写
            for iteration in range(MAX_ITERATIONS):
                current_response = ""
                finish_reason = None

                logger.info(f"OpenAIProvider: 流式第 {iteration+1} 次调用API，模型: {self.model}")

                for content_chunk, fr in self._openai_chat_stream(messages, max_tokens=self.max_tokens, **kwargs):
                    if content_chunk:
                        current_response += content_chunk
                        full_raw_response += content_chunk
                        if chunk_callback:
                            chunk_callback(content_chunk)
                        yield content_chunk
                    if fr:
                        finish_reason = fr

                logger.info(f"OpenAIProvider: 流式迭代 {iteration+1}: 本次长度={len(current_response)}, finish_reason={finish_reason}, 累计={len(full_raw_response)}")

                # 检测主动分段标记
                has_marker = self.CONTINUATION_MARKER in current_response
                if has_marker:
                    full_raw_response = full_raw_response.replace(self.CONTINUATION_MARKER, "")
                    logger.info(f"OpenAIProvider: 流式检测到分段标记，发送'继续'获取下一段")
                    messages = messages + [
                        {"role": "assistant", "content": current_response},
                        {"role": "user", "content": "继续"}
                    ]
                    continue

                if finish_reason != "length":
                    break

                logger.info(f"OpenAIProvider: 流式finish_reason=length，触发续写")
                messages = messages + [
                    {"role": "assistant", "content": current_response},
                    {"role": "user", "content": self._build_inline_continuation_msg(current_response)}
                ]
            else:
                logger.warning(f"OpenAIProvider: 流式达到最大截断续写次数 {MAX_ITERATIONS}，强制停止")

            # 阶段2: 提取代码后检查完整性，不完整则续写
            extracted_code = self.extract_python_code(full_raw_response)

            for retry in range(COMPLETENESS_RETRIES):
                if self._is_code_complete(extracted_code):
                    logger.info(f"OpenAIProvider: 流式代码完整性检查通过")
                    break

                logger.info(f"OpenAIProvider: 流式代码不完整，第 {retry+1} 次完整性续写")
                continuation_prompt = self._build_continuation_prompt(prompt, extracted_code)
                new_raw, _ = self._openai_chat(
                    [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": continuation_prompt}
                    ],
                    max_tokens=self.max_tokens, **kwargs
                )
                # 流式输出续写内容
                if chunk_callback:
                    chunk_callback(new_raw)
                yield new_raw
                new_code = self.extract_python_code(new_raw)
                extracted_code = self._merge_continuation_code(extracted_code, new_code)
                logger.info(f"OpenAIProvider: 流式完整性续写合并后代码长度: {len(extracted_code)}")
            else:
                logger.warning(f"OpenAIProvider: 流式达到最大完整性续写次数 {COMPLETENESS_RETRIES}，使用当前代码")

            logger.info(f"OpenAIProvider: 流式全部完成，总长度: {len(full_raw_response)}")

            # 保存原始响应并提取代码
            self.last_raw_response = full_raw_response
            fixed_code = self.validate_and_fix_code_format(extracted_code)
            self.last_extracted_code = fixed_code

        except Exception as e:
            logger.error(f"OpenAI 流式API调用失败: {e}", exc_info=True)
            raise

    def generate_code_with_stream(self, prompt: str, chunk_callback: callable = None, **kwargs) -> str:
        """流式生成代码并返回完整结果

        Args:
            prompt: 提示词
            chunk_callback: 每个chunk的回调函数
            **kwargs: 其他参数

        Returns:
            生成的完整代码
        """
        # 消费完整个流式生成器（内部已做续写和统一提取）
        for _chunk in self.stream_generate_code(prompt, chunk_callback, **kwargs):
            pass

        return self.last_extracted_code or ""


class ClaudeProvider(BaseAIProvider):
    """Claude提供者"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        self.api_key = config.get("api_key", os.getenv("ANTHROPIC_API_KEY"))
        self.base_url = config.get("base_url", os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com")).rstrip("/")
        self.model = config.get("model", "claude-3-sonnet-20240229")
        self.max_tokens = config.get("max_tokens", int(os.getenv("ANTHROPIC_MAX_TOKENS", "80000")))
        self.timeout = config.get("timeout", int(os.getenv("ANTHROPIC_TIMEOUT", "600")))

        import anthropic
        self._client = anthropic.Anthropic(
            api_key=self.api_key,
            base_url=self.base_url,
            timeout=self.timeout,
            default_headers={"anthropic-beta": "context-1m-2025-08-07"},
        )

    def _claude_chat(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
        """非流式调用 Anthropic SDK，返回 (content, stop_reason)

        Args:
            system_prompt: 系统提示词（字符串或数组）
            messages: 消息列表
            max_tokens: 最大token数
            temperature: 温度参数
            use_cache: 是否启用提示词缓存（默认True）
            **kwargs: 其他参数
        """
        filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

        # 如果启用缓存，将system_prompt转换为数组格式并添加cache_control
        if use_cache and isinstance(system_prompt, str) and system_prompt.strip():
            system_prompt = [
                {
                    "type": "text",
                    "text": system_prompt,
                    "cache_control": {"type": "ephemeral"}
                }
            ]

        # 缓存首条 user message（通常包含规则+源结构+期望结构，20K-120K字符）
        # 续写/纠正时该消息会被重复发送，缓存命中可节省 90% 的输入 token 费用
        if use_cache and messages and messages[0].get("role") == "user":
            first_content = messages[0].get("content")
            if isinstance(first_content, str) and len(first_content) > 1024:
                messages = list(messages)
                messages[0] = {
                    "role": "user",
                    "content": [{"type": "text", "text": first_content, "cache_control": {"type": "ephemeral"}}]
                }

        response = self._client.messages.create(
            model=self.model,
            max_tokens=max_tokens or max(self.max_tokens, 64000),
            temperature=temperature,
            system=system_prompt,
            messages=messages,
            **filtered_kwargs,
        )
        content = response.content[0].text
        stop_reason = response.stop_reason
        return content, stop_reason

    def _claude_chat_stream(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
        """流式调用 Anthropic SDK，yield (text_chunk, stop_reason)

        Args:
            system_prompt: 系统提示词（字符串或数组）
            messages: 消息列表
            max_tokens: 最大token数
            temperature: 温度参数
            use_cache: 是否启用提示词缓存（默认True）
            **kwargs: 其他参数
        """
        filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

        # 如果启用缓存，将system_prompt转换为数组格式并添加cache_control
        if use_cache and isinstance(system_prompt, str) and system_prompt.strip():
            system_prompt = [
                {
                    "type": "text",
                    "text": system_prompt,
                    "cache_control": {"type": "ephemeral"}
                }
            ]

        # 缓存首条 user message（通常包含规则+源结构+期望结构，20K-120K字符）
        # 续写/纠正时该消息会被重复发送，缓存命中可节省 90% 的输入 token 费用
        if use_cache and messages and messages[0].get("role") == "user":
            first_content = messages[0].get("content")
            if isinstance(first_content, str) and len(first_content) > 1024:
                messages = list(messages)
                messages[0] = {
                    "role": "user",
                    "content": [{"type": "text", "text": first_content, "cache_control": {"type": "ephemeral"}}]
                }

        stop_reason = None
        with self._client.messages.stream(
            model=self.model,
            max_tokens=max_tokens or max(self.max_tokens, 64000),
            temperature=temperature,
            system=system_prompt,
            messages=messages,
            **filtered_kwargs,
        ) as stream:
            for text in stream.text_stream:
                if text:
                    yield text, stop_reason
            final_message = stream.get_final_message()
            stop_reason = final_message.stop_reason
        # 最后 yield 确保调用方拿到 stop_reason
        yield "", stop_reason

    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码（流式接收避免超时，max_tokens截断时循环续写，完成后再检查代码完整性）"""
        import logging
        import ast
        logger = logging.getLogger(__name__)

        MAX_ITERATIONS = 10
        COMPLETENESS_RETRIES = 3
        effective_max_tokens = max(self.max_tokens, 64000)

        try:
            logger.info(f"ClaudeProvider: 开始生成代码，提示词长度: {len(prompt)}, max_tokens: {effective_max_tokens}")

            system_prompt = (
                "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
                "包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。"
                "特别注意根据业务场景选择合适的主键进行数据关联和计算。"
                "只返回Python代码，不要包含解释或其他文本。\n\n"
                "⚠️ 缩进纪律（必须严格遵守）：\n"
                "1. break退出for循环后，下一行代码必须回退到for语句的缩进级别\n"
                "2. if/elif/else块结束后，后续独立代码必须与if同级，禁止嵌套在else内部\n"
                "3. 各步骤注释（# === N. ===）必须全部在同一缩进级别\n\n"
                "重要：如果代码较长（超过150行），请主动分段输出。"
                "每段在逻辑完整的位置断开（如函数定义之间），"
                f"段末单独输出一行 {self.CONTINUATION_MARKER} 作为标记。"
                "收到'继续'后输出下一段。最后一段不需要标记。"
            )

            messages = [{"role": "user", "content": prompt}]
            full_raw_response = ""

            # 阶段1: 基于 stop_reason 截断或主动分段标记的续写
            for iteration in range(MAX_ITERATIONS):
                current_chunk = ""
                stop_reason = None

                try:
                    for text_chunk, sr in self._claude_chat_stream(system_prompt, messages, max_tokens=effective_max_tokens, **kwargs):
                        if text_chunk:
                            current_chunk += text_chunk
                        if sr:
                            stop_reason = sr
                except Exception as stream_err:
                    logger.warning(f"ClaudeProvider: 流式API失败({stream_err})，回退到普通API")
                    current_chunk, stop_reason = self._claude_chat(system_prompt, messages, max_tokens=effective_max_tokens, **kwargs)

                full_raw_response += current_chunk
                logger.info(f"ClaudeProvider: 迭代 {iteration+1}: 本次长度={len(current_chunk)}, stop_reason={stop_reason}, 累计={len(full_raw_response)}")

                # 检测主动分段标记
                has_marker = self.CONTINUATION_MARKER in current_chunk
                if has_marker:
                    full_raw_response = full_raw_response.replace(self.CONTINUATION_MARKER, "")
                    logger.info(f"ClaudeProvider: 检测到分段标记，发送'继续'获取下一段")
                    messages = messages + [
                        {"role": "assistant", "content": current_chunk},
                        {"role": "user", "content": "继续"}
                    ]
                    continue

                if stop_reason != "max_tokens":
                    break

                logger.info(f"ClaudeProvider: stop_reason=max_tokens，触发续写（已生成 {len(full_raw_response)} 字符）")
                messages = messages + [
                    {"role": "assistant", "content": current_chunk},
                    {"role": "user", "content": self._build_inline_continuation_msg(current_chunk)}
                ]
            else:
                logger.warning(f"ClaudeProvider: 达到最大截断续写次数 {MAX_ITERATIONS}，强制停止")

            logger.info(f"ClaudeProvider: 最终原始响应总长度: {len(full_raw_response)}")

            if len(full_raw_response) <= 1000:
                logger.info(f"ClaudeProvider: 完整API响应: {full_raw_response}")
            else:
                logger.info(f"ClaudeProvider: API响应前500字符: {full_raw_response[:500]}")
                logger.info(f"ClaudeProvider: API响应后500字符: {full_raw_response[-500:]}")

            # 阶段2: 提取代码后检查完整性，不完整则续写
            extracted_code = self.extract_python_code(full_raw_response)
            logger.info(f"ClaudeProvider: 提取的代码长度: {len(extracted_code)}")

            for retry in range(COMPLETENESS_RETRIES):
                if self._is_code_complete(extracted_code):
                    logger.info(f"ClaudeProvider: 代码完整性检查通过")
                    break

                logger.info(f"ClaudeProvider: 代码不完整，第 {retry+1} 次完整性续写")
                continuation_prompt = self._build_continuation_prompt(prompt, extracted_code)
                try:
                    new_raw = ""
                    for text_chunk, _ in self._claude_chat_stream(system_prompt, [{"role": "user", "content": continuation_prompt}], max_tokens=effective_max_tokens):
                        if text_chunk:
                            new_raw += text_chunk
                except Exception:
                    new_raw, _ = self._claude_chat(system_prompt, [{"role": "user", "content": continuation_prompt}], max_tokens=effective_max_tokens)

                new_code = self.extract_python_code(new_raw)
                extracted_code = self._merge_continuation_code(extracted_code, new_code)
                logger.info(f"ClaudeProvider: 完整性续写合并后代码长度: {len(extracted_code)}")
            else:
                logger.warning(f"ClaudeProvider: 达到最大完整性续写次数 {COMPLETENESS_RETRIES}，使用当前代码")

            fixed_code = self.validate_and_fix_code_format(extracted_code)
            if fixed_code != extracted_code:
                logger.info(f"ClaudeProvider: 代码格式已修复 (原始: {len(extracted_code)}, 修复后: {len(fixed_code)})")

            try:
                ast.parse(fixed_code)
                logger.info("ClaudeProvider: 最终代码语法校验通过")
            except SyntaxError as e:
                logger.warning(f"ClaudeProvider: 最终代码存在语法错误: {e}")

            self.last_raw_response = full_raw_response
            self.last_extracted_code = fixed_code
            return fixed_code

        except Exception as e:
            logger.error(f"Claude API调用失败: {e}", exc_info=True)
            fallback_code = '''import pandas as pd
import os

def process_excel_files(input_folder: str, output_file: str):
    """处理Excel文件"""
    print(f"处理文件夹: {input_folder}")
    print(f"输出文件: {output_file}")
    return True

if __name__ == "__main__":
    process_excel_files("input", "output.xlsx")'''
            logger.info("ClaudeProvider: 返回备用代码")
            return self.extract_python_code(fallback_code)

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口"""
        anthropic_messages = []
        system_message = ""
        for msg in messages:
            if msg["role"] == "system":
                system_message = msg["content"]
            else:
                anthropic_messages.append({"role": msg["role"], "content": msg["content"]})

        content, _ = self._claude_chat(
            system_message or "",
            anthropic_messages,
            max_tokens=self.max_tokens,
            **{k: v for k, v in kwargs.items() if k not in ("model", "messages", "max_tokens", "system")}
        )
        return content

    def chat_stream(self, messages: List[Dict[str, str]], chunk_callback: callable = None, **kwargs) -> str:
        """流式对话接口"""
        anthropic_messages = []
        system_message = ""
        for msg in messages:
            if msg["role"] == "system":
                system_message = msg["content"]
            else:
                anthropic_messages.append({"role": msg["role"], "content": msg["content"]})

        filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("model", "messages", "max_tokens", "system")}
        full_content = ""
        for text_chunk, _sr in self._claude_chat_stream(
            system_message or "", anthropic_messages,
            max_tokens=self.max_tokens, **filtered_kwargs
        ):
            if text_chunk:
                full_content += text_chunk
                if chunk_callback:
                    chunk_callback(text_chunk)
        return full_content

    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        messages = [
            {"role": "user", "content": prompt}
        ]
        return self.chat(messages, **kwargs)

    def stream_generate_code(self, prompt: str, chunk_callback: callable = None, **kwargs):
        """流式生成代码（流式接收避免超时，max_tokens截断或代码不完整时对话续写，最后统一提取）

        Args:
            prompt: 提示词
            chunk_callback: 每个chunk的回调函数
            **kwargs: 其他参数

        Yields:
            原始文本片段
        """
        import logging
        logger = logging.getLogger(__name__)

        MAX_ITERATIONS = 10
        COMPLETENESS_RETRIES = 3
        effective_max_tokens = max(self.max_tokens, 64000)

        try:
            logger.info(f"ClaudeProvider: 开始流式生成代码，提示词长度: {len(prompt)}, max_tokens: {effective_max_tokens}")

            system_prompt = (
                "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
                "包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。"
                "特别注意根据业务场景选择合适的主键进行数据关联和计算。"
                "只返回Python代码，不要包含解释或其他文本。\n\n"
                "⚠️ 缩进纪律（必须严格遵守）：\n"
                "1. break退出for循环后，下一行代码必须回退到for语句的缩进级别\n"
                "2. if/elif/else块结束后，后续独立代码必须与if同级，禁止嵌套在else内部\n"
                "3. 各步骤注释（# === N. ===）必须全部在同一缩进级别\n\n"
                "重要：如果代码较长（超过150行），请主动分段输出。"
                "每段在逻辑完整的位置断开（如函数定义之间），"
                f"段末单独输出一行 {self.CONTINUATION_MARKER} 作为标记。"
                "收到'继续'后输出下一段。最后一段不需要标记。"
            )

            messages = [{"role": "user", "content": prompt}]
            full_raw_response = ""

            # 阶段1: 基于 stop_reason 截断或主动分段标记的续写
            for iteration in range(MAX_ITERATIONS):
                current_chunk = ""
                stop_reason = None

                for text_chunk, sr in self._claude_chat_stream(system_prompt, messages, max_tokens=effective_max_tokens, **kwargs):
                    if text_chunk:
                        current_chunk += text_chunk
                        full_raw_response += text_chunk
                        if chunk_callback:
                            chunk_callback(text_chunk)
                        yield text_chunk
                    if sr:
                        stop_reason = sr

                logger.info(f"ClaudeProvider: 流式迭代 {iteration+1}: 本次长度={len(current_chunk)}, stop_reason={stop_reason}, 累计={len(full_raw_response)}")

                # 检测主动分段标记
                has_marker = self.CONTINUATION_MARKER in current_chunk
                if has_marker:
                    full_raw_response = full_raw_response.replace(self.CONTINUATION_MARKER, "")
                    logger.info(f"ClaudeProvider: 流式检测到分段标记，发送'继续'获取下一段")
                    messages = messages + [
                        {"role": "assistant", "content": current_chunk},
                        {"role": "user", "content": "继续"}
                    ]
                    continue

                if stop_reason != "max_tokens":
                    break

                logger.info(f"ClaudeProvider: 流式stop_reason=max_tokens，触发续写")
                messages = messages + [
                    {"role": "assistant", "content": current_chunk},
                    {"role": "user", "content": self._build_inline_continuation_msg(current_chunk)}
                ]
            else:
                logger.warning(f"ClaudeProvider: 流式达到最大截断续写次数 {MAX_ITERATIONS}，强制停止")

            # 阶段2: 提取代码后检查完整性，不完整则续写
            extracted_code = self.extract_python_code(full_raw_response)

            for retry in range(COMPLETENESS_RETRIES):
                if self._is_code_complete(extracted_code):
                    logger.info(f"ClaudeProvider: 流式代码完整性检查通过")
                    break

                logger.info(f"ClaudeProvider: 流式代码不完整，第 {retry+1} 次完整性续写")
                continuation_prompt = self._build_continuation_prompt(prompt, extracted_code)
                new_raw = ""
                for text_chunk, _ in self._claude_chat_stream(system_prompt, [{"role": "user", "content": continuation_prompt}], max_tokens=effective_max_tokens):
                    if text_chunk:
                        new_raw += text_chunk
                        if chunk_callback:
                            chunk_callback(text_chunk)
                        yield text_chunk

                new_code = self.extract_python_code(new_raw)
                extracted_code = self._merge_continuation_code(extracted_code, new_code)
                logger.info(f"ClaudeProvider: 流式完整性续写合并后代码长度: {len(extracted_code)}")
            else:
                logger.warning(f"ClaudeProvider: 流式达到最大完整性续写次数 {COMPLETENESS_RETRIES}，使用当前代码")

            logger.info(f"ClaudeProvider: 流式全部完成，总长度: {len(full_raw_response)}")

            # 最后统一提取代码
            self.last_raw_response = full_raw_response
            fixed_code = self.validate_and_fix_code_format(extracted_code)
            self.last_extracted_code = fixed_code

        except Exception as e:
            logger.error(f"Claude 流式API调用失败: {e}", exc_info=True)
            raise

    def generate_code_with_stream(self, prompt: str, chunk_callback: callable = None, **kwargs) -> str:
        """流式生成代码并返回完整结果

        Args:
            prompt: 提示词
            chunk_callback: 每个chunk的回调函数
            **kwargs: 其他参数

        Returns:
            生成的完整代码
        """
        # 消费完整个流式生成器（内部已做对话续传和统一提取）
        for _chunk in self.stream_generate_code(prompt, chunk_callback, **kwargs):
            pass

        # stream_generate_code 结束后已提取并修复代码，直接返回
        return self.last_extracted_code or ""


class DeepSeekProvider(OpenAIProvider):
    """DeepSeek提供者（OpenAI兼容API，继承OpenAIProvider）"""

    def __init__(self, config: Dict[str, Any]):
        # 不调用 OpenAIProvider.__init__，直接初始化基类和自身属性
        BaseAIProvider.__init__(self)
        self.api_key = config.get("api_key", os.getenv("DEEPSEEK_API_KEY"))
        self.base_url = config.get("base_url", os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")).rstrip("/")
        self.model = config.get("model", "deepseek-chat")
        self.max_tokens = config.get("max_tokens", int(os.getenv("DEEPSEEK_MAX_TOKENS", "80000")))
        self.timeout = int(config.get("timeout", os.getenv("DEEPSEEK_TIMEOUT", "300")))

        # 创建 OpenAI 客户端（DeepSeek 兼容 OpenAI API）
        from openai import OpenAI
        self._client = OpenAI(
            api_key=self.api_key,
            base_url=self.base_url,
            timeout=self.timeout
        )


class OllamaProvider(BaseAIProvider):
    """Ollama提供者（本地AI）"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        import os
        self.base_url = config.get("base_url", os.getenv("OLLAMA_BASE_URL", "http://localhost:11434"))
        self.model = config.get("model", os.getenv("OLLAMA_MODEL", "llama2"))
        self.timeout = int(config.get("timeout", os.getenv("OLLAMA_TIMEOUT", "120")))

    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码"""
        try:
            messages = [
                {"role": "system", "content": "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。特别注意根据业务场景选择合适的主键进行数据关联和计算。只返回Python代码，不要包含解释或其他文本。\n\n⚠️ 缩进纪律：break退出for循环后必须回退到for的缩进级别；if/else块结束后后续代码必须与if同级；各步骤注释（# === N. ===）必须在同一缩进级别。"},
                {"role": "user", "content": prompt}
            ]

            response = http_requests.post(
                f"{self.base_url}/api/chat",
                json={
                    "model": self.model,
                    "messages": messages,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "num_predict": 4000
                    }
                },
                timeout=self.timeout
            )

            if response.status_code == 200:
                raw_response = response.json()["message"]["content"]
                # 提取纯Python代码
                extracted_code = self.extract_python_code(raw_response)
                # 验证和修复代码格式
                fixed_code = self.validate_and_fix_code_format(extracted_code)
                return fixed_code
            else:
                return f"Ollama API错误: {response.status_code} - {response.text}"

        except Exception as e:
            return f"Ollama请求失败: {str(e)}"

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口"""
        try:
            response = http_requests.post(
                f"{self.base_url}/api/chat",
                json={
                    "model": self.model,
                    "messages": messages,
                    "stream": False
                },
                timeout=self.timeout
            )

            if response.status_code == 200:
                return response.json()["message"]["content"]
            else:
                return f"Ollama API错误: {response.status_code} - {response.text}"

        except Exception as e:
            return f"Ollama请求失败: {str(e)}"

    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        messages = [
            {"role": "user", "content": prompt}
        ]
        return self.chat(messages, **kwargs)

    def _fallback_code(self, prompt: str) -> str:
        """备用代码生成"""
        return '''import pandas as pd
import os

def process_excel_files(input_folder: str, output_file: str):
    """处理Excel文件"""
    # 这里应该根据实际规则实现
    print(f"处理文件夹: {input_folder}")
    print(f"输出文件: {output_file}")
    return True

if __name__ == "__main__":
    process_excel_files("input", "output.xlsx")'''


class LocalAIProvider(BaseAIProvider):
    """本地AI提供者（模拟）"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__()
        self.model = config.get("model", "local")

    def generate_code(self, prompt: str, **kwargs) -> str:
        """生成代码（模拟）"""
        # 这里可以集成本地模型如Ollama等
        # 暂时返回示例代码
        raw_code = '''import pandas as pd
import os

def process_excel_files(input_folder: str, output_file: str):
    """处理Excel文件"""
    # 这里应该根据实际规则实现
    print(f"处理文件夹: {input_folder}")
    print(f"输出文件: {output_file}")
    return True

if __name__ == "__main__":
    process_excel_files("input", "output.xlsx")'''
        # 提取纯Python代码（虽然已经是纯代码，但保持一致性）
        extracted_code = self.extract_python_code(raw_code)
        # 验证和修复代码格式
        fixed_code = self.validate_and_fix_code_format(extracted_code)
        return fixed_code

    def chat(self, messages: List[Dict[str, str]], **kwargs) -> str:
        """对话接口（模拟）"""
        return "这是本地AI的回复，请配置实际的本地AI服务。"

    def _openai_chat(self, messages, temperature=0.1, max_tokens=None, **kwargs):
        """兼容 OpenAI 接口的非流式调用，返回 (content, finish_reason)"""
        content = self.chat(messages, **kwargs)
        return content, "stop"

    def _openai_chat_stream(self, messages, temperature=0.1, max_tokens=None, **kwargs):
        """兼容 OpenAI 接口的流式调用，yield (content_chunk, finish_reason)"""
        content = self.chat(messages, **kwargs)
        yield content, "stop"

    def generate_completion(self, prompt: str, **kwargs) -> str:
        """生成完成文本（单轮对话）"""
        messages = [
            {"role": "user", "content": prompt}
        ]
        return self.chat(messages, **kwargs)


class AIProviderFactory:
    """AI提供者工厂"""

    # 提供者映射表
    PROVIDERS = {
        "openai": OpenAIProvider,
        "claude": ClaudeProvider,
        "deepseek": DeepSeekProvider,
        "ollama": OllamaProvider,
        "local": LocalAIProvider
    }

    @staticmethod
    def get_default_config() -> Dict[str, Any]:
        """获取默认配置"""
        import os
        from dotenv import load_dotenv

        # 加载环境变量
        load_dotenv()

        # 获取主提供者
        provider_type = os.getenv("AI_PROVIDER", "openai").lower()

        # 根据提供者类型构建配置
        config = {
            "provider_type": provider_type,
            "max_retries": int(os.getenv("AI_MAX_RETRIES", "3")),
            "retry_delay": int(os.getenv("AI_RETRY_DELAY", "2")),
            "rate_limit": int(os.getenv("AI_RATE_LIMIT", "10"))
        }

        # 添加提供者特定配置
        if provider_type == "openai":
            config.update({
                "api_key": os.getenv("OPENAI_API_KEY"),
                "base_url": os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
                "model": os.getenv("OPENAI_MODEL", "gpt-4-turbo-preview"),
                "timeout": int(os.getenv("OPENAI_TIMEOUT", "60")),
                "max_tokens": int(os.getenv("OPENAI_MAX_TOKENS", "8000"))
            })
        elif provider_type == "claude":
            config.update({
                "api_key": os.getenv("ANTHROPIC_API_KEY"),
                "base_url": os.getenv("ANTHROPIC_BASE_URL"),
                "model": os.getenv("ANTHROPIC_MODEL", "claude-3-sonnet-20240229"),
                "max_tokens": int(os.getenv("ANTHROPIC_MAX_TOKENS", "8000")),
                "temperature": float(os.getenv("ANTHROPIC_TEMPERATURE", "0.1"))
            })
        elif provider_type == "deepseek":
            config.update({
                "api_key": os.getenv("DEEPSEEK_API_KEY"),
                "base_url": os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
                "model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
                "timeout": int(os.getenv("DEEPSEEK_TIMEOUT", "60")),
                "max_tokens": int(os.getenv("DEEPSEEK_MAX_TOKENS", "8000"))
            })
        elif provider_type == "ollama":
            config.update({
                "base_url": os.getenv("OLLAMA_BASE_URL", "http://localhost:11434"),
                "model": os.getenv("OLLAMA_MODEL", "llama2"),
                "timeout": int(os.getenv("OLLAMA_TIMEOUT", "120"))
            })
        elif provider_type == "local":
            config.update({
                "model": os.getenv("LOCAL_MODEL", "local")
            })

        return config

    @staticmethod
    def create_provider(provider_type: str = None, config: Dict[str, Any] = None) -> BaseAIProvider:
        """创建AI提供者

        Args:
            provider_type: 提供者类型，如果为None则使用配置中的默认值
            config: 配置字典，如果为None则使用默认配置
        """
        # 如果指定了 provider_type 但没有提供 config，根据 provider_type 获取配置
        if provider_type is not None and config is None:
            # 临时修改环境变量来获取正确的配置
            import os
            original_provider = os.getenv("AI_PROVIDER")
            os.environ["AI_PROVIDER"] = provider_type
            config = AIProviderFactory.get_default_config()
            # 恢复原始环境变量
            if original_provider:
                os.environ["AI_PROVIDER"] = original_provider
            else:
                os.environ.pop("AI_PROVIDER", None)
        elif config is None:
            # 如果没有提供配置，使用默认配置
            config = AIProviderFactory.get_default_config()

        # 如果没有指定提供者类型，使用配置中的默认值
        if provider_type is None:
            provider_type = config.get("provider_type", "openai")

        # 检查提供者类型是否支持
        if provider_type not in AIProviderFactory.PROVIDERS:
            raise ValueError(f"不支持的AI提供者类型: {provider_type}")

        # 创建提供者实例
        provider_class = AIProviderFactory.PROVIDERS[provider_type]
        return provider_class(config)

    @staticmethod
    def create_with_fallback() -> BaseAIProvider:
        """创建AI提供者，支持备用提供者"""
        import os
        from dotenv import load_dotenv

        load_dotenv()

        # 获取主提供者和备用提供者
        main_provider = os.getenv("AI_PROVIDER", "openai")
        fallback_provider = os.getenv("AI_FALLBACK_PROVIDER", "local")

        try:
            # 尝试创建主提供者
            return AIProviderFactory.create_provider(main_provider)
        except Exception as e:
            print(f"主AI提供者 {main_provider} 创建失败: {e}")
            print(f"尝试使用备用提供者: {fallback_provider}")

            try:
                # 尝试创建备用提供者
                return AIProviderFactory.create_provider(fallback_provider)
            except Exception as e2:
                print(f"备用AI提供者 {fallback_provider} 也创建失败: {e2}")
                raise RuntimeError("所有AI提供者都创建失败")