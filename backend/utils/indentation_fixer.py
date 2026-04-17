"""
统一的Python代码缩进修复器

将分散在 code_sandbox.py、ai_provider.py、formula_code_generator.py 中的
所有缩进修复逻辑集中到一处，方便维护和复用。

使用方式：
    from backend.utils.indentation_fixer import IndentationFixer
    fixer = IndentationFixer()
    fixed_code = fixer.fix_sandbox_pipeline(code)   # 沙箱执行前的修复
    fixed_code = fixer.fix_formula_pipeline(code)    # 公式代码生成后的修复
    fixed_code = fixer.fix_general(code)             # 通用修复（ai_provider用）
"""

import re
import logging
from typing import List, Optional, Set

logger = logging.getLogger(__name__)


class IndentationFixer:
    """统一的缩进修复器，整合所有缩进修复逻辑"""

    # ================================================================
    #  管线入口（Pipeline Entry Points）
    # ================================================================

    def fix_sandbox_pipeline(self, code: str) -> str:
        """沙箱执行前的缩进修复管线

        替代 code_sandbox._clean_code_syntax 中的缩进修复调用。
        三层防御：
          L0: Tab→Space 归一化
          L1: 自定义修复器（最多5轮）
          L2: autopep8 工业级格式化（最终防线）
        """
        # --- L0: Tab→Space 归一化 ---
        code = code.replace('\t', '    ')

        # --- L1: 自定义修复器 ---
        for _round in range(5):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError:
                pass

            code = self.fix_indent_drift_after_break(code)
            code = self.fix_block_keyword_alignment(code)
            code = self.fix_cascading_section_nesting(code)
            code = self.fix_unindent_mismatch(code)
            code = self.fix_unexpected_indent(code)
            code = self.fix_expected_indented_block(code)
            code = self.fix_orphaned_break_continue(code)

        # --- L2: autopep8 最终防线 ---
        try:
            compile(code, '<check>', 'exec')
            return code
        except SyntaxError:
            code = self._fix_with_autopep8(code)

        return code

    def fix_formula_pipeline(self, code: str) -> str:
        """公式代码生成后的缩进修复管线

        替代 formula_code_generator 中多处调用的 2 步修复。
        按顺序执行：Tab归一化 → AST检查 → for循环体脱离 → 列级联缩进 → autopep8兜底
        """
        code = code.replace('\t', '    ')
        # AST 前置检查：如果代码已经合法，跳过所有修复
        try:
            compile(code, '<check>', 'exec')
            return code
        except SyntaxError:
            pass
        code = self.fix_for_loop_body_indentation(code)
        code = self.fix_cascading_column_indentation(code)
        # 如果仍有编译错误，用 autopep8 兜底
        try:
            compile(code, '<check>', 'exec')
        except SyntaxError:
            code = self._fix_with_autopep8(code)
        return code

    def fix_general(self, code: str) -> str:
        """通用缩进修复（基于缩进栈的归一化）

        替代 ai_provider._fix_python_indentation。
        适用于一般性的缩进混乱（混用空格数、冒号后没缩进等）。
        """
        code = code.replace('\t', '    ')
        return self.fix_python_indentation(code)

    # ================================================================
    #  通用 Python 缩进修复方法
    # ================================================================

    def fix_indent_drift_after_break(self, code: str, max_passes: int = 5) -> str:
        """修复break/continue后的缩进漂移

        AI常见错误：在break/continue后继续以相同缩进写代码，
        而实际上应该回退到上层块的级别。使用自顶向下的作用域栈跟踪。
        """
        for pass_num in range(max_passes):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError:
                pass

            lines = code.split('\n')
            loop_stack = []   # 当前活跃的for/while循环缩进级别
            block_stack = []  # 当前活跃的块头语句缩进级别
            fixed_any = False
            i = 0

            while i < len(lines):
                stripped = lines[i].strip()
                if not stripped:
                    i += 1
                    continue

                indent = len(lines[i]) - len(lines[i].lstrip())

                # 弹出已退出的作用域
                while loop_stack and indent <= loop_stack[-1]:
                    loop_stack.pop()
                while block_stack and indent <= block_stack[-1]:
                    block_stack.pop()

                # 跟踪块头语句
                if stripped.rstrip().endswith(':') and not stripped.startswith('#'):
                    if stripped.startswith('for ') or stripped.startswith('while '):
                        loop_stack.append(indent)
                        block_stack.append(indent)
                    elif any(stripped.startswith(kw) for kw in (
                        'if ', 'elif ', 'else:', 'try:', 'except ', 'except:',
                        'finally:', 'with ', 'async for ', 'async with ',
                    )):
                        block_stack.append(indent)

                # 检测 break / continue
                is_break = is_continue = False
                if stripped:
                    word = stripped.split()[0]
                    if word == 'break':
                        rest = stripped[5:]
                        if not rest or rest.lstrip().startswith('#'):
                            is_break = True
                    elif word == 'continue':
                        rest = stripped[8:]
                        if not rest or rest.lstrip().startswith('#'):
                            is_continue = True

                if not (is_break or is_continue):
                    i += 1
                    continue

                # 计算目标缩进
                target_indent = None
                if is_break and loop_stack:
                    target_indent = loop_stack[-1]
                elif is_continue and loop_stack:
                    target_indent = loop_stack[-1] + 4
                elif is_continue and block_stack:
                    target_indent = block_stack[-1]

                if target_indent is None:
                    i += 1
                    continue

                # 查找break/continue后第一个非空行
                k = i + 1
                while k < len(lines) and not lines[k].strip():
                    k += 1

                if k >= len(lines):
                    i += 1
                    continue

                next_indent = len(lines[k]) - len(lines[k].lstrip())

                # 漂移判定：下一行缩进 >= break/continue缩进 → 没有正确回退
                if next_indent >= indent:
                    dedent_amount = next_indent - target_indent
                    if dedent_amount > 0:
                        for m in range(k, len(lines)):
                            mline = lines[m]
                            if not mline.strip():
                                continue
                            mindent = len(mline) - len(mline.lstrip())
                            if mindent < indent:
                                break
                            new_indent = max(0, mindent - dedent_amount)
                            lines[m] = ' ' * new_indent + mline.lstrip()
                        fixed_any = True

                i += 1

            if not fixed_any:
                break

            code = '\n'.join(lines)

        return code

    def fix_block_keyword_alignment(self, code: str) -> str:
        """修复except/finally/else/elif与对应块头缩进不一致

        常见模式：AI把except写在try体内级别，而非与try同级。
        只在编译失败时触发。
        """
        try:
            compile(code, '<check>', 'exec')
            return code
        except SyntaxError:
            pass

        lines = code.split('\n')

        block_openers = {
            'except': ('try',),
            'finally': ('try',),
            'elif': ('if', 'elif'),
            'else': ('if', 'elif', 'for', 'while', 'try', 'except'),
        }

        fixed_any = False
        i = 0

        while i < len(lines):
            stripped = lines[i].strip()
            if not stripped:
                i += 1
                continue

            indent = len(lines[i]) - len(lines[i].lstrip())

            # 判断是否为续行关键字
            keyword = None
            for kw in block_openers:
                if stripped == kw + ':' or stripped.startswith(kw + ' ') or stripped.startswith(kw + '('):
                    keyword = kw
                    break

            if keyword is None:
                i += 1
                continue

            openers = block_openers[keyword]

            # 向上查找匹配的块头
            opener_indent = None
            for j in range(i - 1, -1, -1):
                pstripped = lines[j].strip()
                if not pstripped:
                    continue
                pindent = len(lines[j]) - len(lines[j].lstrip())

                if pindent > indent:
                    continue

                for opener in openers:
                    if (pstripped == opener + ':' or pstripped.startswith(opener + ' ')
                            or pstripped.startswith(opener + ':')):
                        opener_indent = pindent
                        break
                if opener_indent is not None:
                    break

                if pstripped.startswith('def ') or pstripped.startswith('class '):
                    break

            if opener_indent is not None and opener_indent != indent:
                diff = indent - opener_indent
                lines[i] = ' ' * opener_indent + stripped
                for m in range(i + 1, len(lines)):
                    mline = lines[m]
                    if not mline.strip():
                        continue
                    mindent = len(mline) - len(mline.lstrip())
                    if mindent <= indent:
                        break
                    lines[m] = ' ' * max(0, mindent - diff) + mline.lstrip()
                fixed_any = True

            i += 1

        if fixed_any:
            return '\n'.join(lines)
        return code

    def fix_cascading_section_nesting(self, code: str, max_nesting: int = 20) -> str:
        """修复步骤注释级联嵌套（# === N. === 缩进递增）

        AI把每个处理步骤写在上一个步骤的if块内部，
        导致缩进越来越深。仅在编译失败且嵌套深度超过阈值时触发。
        """
        try:
            compile(code, '<check>', 'exec')
            return code
        except SyntaxError:
            pass

        lines = code.split('\n')

        max_indent = max((len(l) - len(l.lstrip()) for l in lines if l.strip()), default=0)
        if max_indent < max_nesting * 4:
            return code

        section_pattern = re.compile(r'^#\s*={2,}\s*\d+[\.\s]')

        # 按函数分组处理
        func_ranges = []
        for i, line in enumerate(lines):
            stripped = line.strip()
            indent = len(line) - len(line.lstrip()) if stripped else 0
            if stripped.startswith('def ') and stripped.endswith(':') and indent == 0:
                func_ranges.append([i, len(lines)])
                if len(func_ranges) > 1:
                    func_ranges[-2][1] = i

        if not func_ranges:
            func_ranges = [(0, len(lines))]

        fixed_any = False
        for func_start, func_end in func_ranges:
            markers = []
            for i in range(func_start, func_end):
                stripped = lines[i].strip()
                if section_pattern.match(stripped):
                    indent = len(lines[i]) - len(lines[i].lstrip())
                    markers.append((i, indent))

            if len(markers) < 3:
                continue

            base_indent = markers[0][1]
            has_cascade = any(m[1] > base_indent for m in markers[1:])
            if not has_cascade:
                continue

            logger.info(
                f"检测到步骤级联嵌套：{len(markers)}个步骤标记，"
                f"缩进范围 {min(m[1] for m in markers)}-{max(m[1] for m in markers)}，"
                f"统一修正为 {base_indent}"
            )

            for idx, (marker_line, marker_indent) in enumerate(markers):
                if marker_indent == base_indent:
                    continue

                if idx + 1 < len(markers):
                    block_end = markers[idx + 1][0]
                else:
                    block_end = func_end
                    for j in range(marker_line + 1, func_end):
                        s = lines[j].strip()
                        if s.startswith('return ') or s == 'return':
                            indent_j = len(lines[j]) - len(lines[j].lstrip())
                            if indent_j <= base_indent:
                                block_end = j
                                break

                indent_diff = marker_indent - base_indent

                for j in range(marker_line, block_end):
                    line = lines[j]
                    if not line.strip():
                        continue
                    current_indent = len(line) - len(line.lstrip())
                    new_indent = max(0, current_indent - indent_diff)
                    lines[j] = ' ' * new_indent + line.lstrip()

                fixed_any = True

        if not fixed_any:
            return code

        result = '\n'.join(lines)

        # 修复后可能暴露新问题，再跑一遍
        try:
            compile(result, '<check>', 'exec')
            return result
        except SyntaxError:
            result = self.fix_indent_drift_after_break(result)
            result = self.fix_block_keyword_alignment(result)
            return result

    def fix_unindent_mismatch(self, code: str, max_passes: int = 10) -> str:
        """修复 'unindent does not match any outer indentation level' 错误

        根据 SyntaxError 定位到出错行，向上收集合法的缩进级别栈，
        将出错行对齐到最近的合法缩进级别，并调整同缩进块的后续行。
        """
        for _pass in range(max_passes):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError as e:
                if 'unindent does not match' not in str(e):
                    return code
                if e.lineno is None:
                    return code

                lines = code.split('\n')
                err_line_idx = e.lineno - 1
                if err_line_idx < 0 or err_line_idx >= len(lines):
                    return code

                err_line = lines[err_line_idx]
                if not err_line.strip():
                    return code

                err_indent = len(err_line) - len(err_line.lstrip())

                # 向上收集合法的缩进级别（跳过注释行，遇到同级或更浅的def/class停止）
                valid_indents: Set[int] = {0}
                for j in range(err_line_idx - 1, -1, -1):
                    ln = lines[j]
                    s = ln.strip()
                    if not s:
                        continue
                    if s.startswith('#'):
                        continue  # 注释行不建立缩进级别
                    ind = len(ln) - len(ln.lstrip())
                    # 遇到顶层 def/class，停止向上扫描（跨函数边界）
                    if ind == 0 and (s.startswith('def ') or s.startswith('class ')
                                     or s.startswith('async def ')):
                        valid_indents.add(ind)
                        if s.endswith(':'):
                            valid_indents.add(ind + 4)
                        break
                    valid_indents.add(ind)
                    if s.endswith(':') and not s.startswith('#'):
                        valid_indents.add(ind + 4)

                if err_indent in valid_indents:
                    return code

                # 找最近的合法缩进（优先向左对齐）
                candidates_left = [v for v in valid_indents if v < err_indent]
                if candidates_left:
                    target_indent = max(candidates_left)
                else:
                    target_indent = min(valid_indents, key=lambda v: abs(v - err_indent))

                diff = err_indent - target_indent
                if diff == 0:
                    return code

                logger.info(
                    f"[unindent修复] 行{e.lineno}: 缩进{err_indent}→{target_indent} (差{diff})")

                for m in range(err_line_idx, len(lines)):
                    ln = lines[m]
                    if not ln.strip():
                        continue
                    mindent = len(ln) - len(ln.lstrip())
                    if m > err_line_idx and mindent < err_indent:
                        break
                    new_indent = max(0, mindent - diff)
                    lines[m] = ' ' * new_indent + ln.lstrip()

                code = '\n'.join(lines)

        return code

    def fix_expected_indented_block(self, code: str, max_passes: int = 10) -> str:
        """修复 'expected an indented block after ...' 错误

        AI常见错误：def/if/for/while/try等块头语句后面的代码体
        没有正确缩进（与块头同级或更浅），导致编译报错。
        策略：定位块头行，把连续缩进不足的行整体加缩进到块头+4。
        """
        for _pass in range(max_passes):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError as e:
                if 'expected an indented block' not in str(e):
                    return code
                if e.lineno is None:
                    return code

                lines = code.split('\n')
                err_idx = e.lineno - 1  # 0-based
                # 从报错行的前一行开始向上查找块头语句
                block_head_idx = None
                block_head_indent = None
                for j in range(err_idx - 1, -1, -1):
                    ln = lines[j].strip()
                    if not ln or ln.startswith('#'):
                        continue
                    if ln.endswith(':'):
                        block_head_idx = j
                        block_head_indent = len(lines[j]) - len(lines[j].lstrip())
                        break
                    # 如果遇到非块头的实际代码行，停止
                    break

                if block_head_idx is None:
                    return code

                expected_body_indent = block_head_indent + 4

                # 找到块头后第一个非空行
                body_start = None
                for j in range(block_head_idx + 1, len(lines)):
                    if lines[j].strip():
                        body_start = j
                        break

                if body_start is None:
                    return code

                body_indent = len(lines[body_start]) - len(lines[body_start].lstrip())

                if body_indent >= expected_body_indent:
                    return code

                indent_add = expected_body_indent - body_indent

                logger.info(
                    f"[expected-indent修复] 块头行{block_head_idx + 1}，"
                    f"体行{body_start + 1}起: 缩进{body_indent}→{expected_body_indent} (加{indent_add})")

                # 把所有与 body_indent 同级（或更深）的连续行整体加缩进
                # 关键改进：不在遇到第二个同级行时 break，而是把整段 body 都修
                for m in range(body_start, len(lines)):
                    ln = lines[m]
                    if not ln.strip():
                        continue
                    mindent = len(ln) - len(ln.lstrip())
                    # 遇到比块头更浅的缩进，说明已经离开这个块了
                    if mindent < block_head_indent:
                        break
                    # 缩进不足的行（在 body_indent 同级或介于 body_indent 和 expected 之间）统一加缩进
                    if mindent < expected_body_indent:
                        lines[m] = ' ' * (mindent + indent_add) + ln.lstrip()

                code = '\n'.join(lines)

        return code

    def fix_unexpected_indent(self, code: str, max_passes: int = 10) -> str:
        """修复 'unexpected indent' 错误

        AI常见错误：前一行不是块头语句(不以 : 结尾)，但下一行突然增加了缩进。
        与 fix_unindent_mismatch 不同，这里严格检查局部上下文（紧邻前一行），
        而不是全局缩进级别栈，因为 8 空格可能在全局合法但在局部不合法。
        策略：找到紧邻前一非空行的缩进，如果前一行不以 : 结尾，则把多出的缩进减回去。
        """
        for _pass in range(max_passes):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError as e:
                if 'unexpected indent' not in str(e):
                    return code
                if e.lineno is None:
                    return code

                lines = code.split('\n')
                err_idx = e.lineno - 1
                if err_idx < 0 or err_idx >= len(lines):
                    return code

                err_line = lines[err_idx]
                if not err_line.strip():
                    return code

                err_indent = len(err_line) - len(err_line.lstrip())

                # 找紧邻的前一个非空、非注释行
                prev_idx = None
                prev_indent = 0
                prev_is_block_head = False
                for j in range(err_idx - 1, -1, -1):
                    s = lines[j].strip()
                    if not s:
                        continue
                    if s.startswith('#'):
                        continue
                    prev_idx = j
                    prev_indent = len(lines[j]) - len(lines[j].lstrip())
                    prev_is_block_head = s.endswith(':')
                    break

                if prev_idx is None:
                    return code

                # 前一行是块头→允许缩进到 prev_indent+4；否则最大允许 prev_indent
                if prev_is_block_head:
                    max_allowed = prev_indent + 4
                else:
                    max_allowed = prev_indent

                over = err_indent - max_allowed
                if over <= 0:
                    return code

                logger.info(
                    f"[unexpected-indent修复] 行{e.lineno}: "
                    f"缩进{err_indent}→{max_allowed} (减{over}，前行{prev_idx+1}缩进{prev_indent}"
                    f"{'，是块头' if prev_is_block_head else '，非块头'})")

                # 从出错行开始，把同级及更深缩进的整个块减少 over 个空格
                for m in range(err_idx, len(lines)):
                    ln = lines[m]
                    if not ln.strip():
                        continue
                    mindent = len(ln) - len(ln.lstrip())
                    # 遇到缩进 < err_indent 的行（不含注释），说明整块结束
                    if m > err_idx and mindent < err_indent:
                        break
                    new_indent = max(0, mindent - over)
                    lines[m] = ' ' * new_indent + ln.lstrip()

                code = '\n'.join(lines)

        return code

    def fix_orphaned_break_continue(self, code: str, max_passes: int = 20) -> str:
        """修复 'break/continue outside loop' 错误

        级联缩进修复后，原本在循环内的 break/continue 可能被拉到循环外面。
        策略：注释掉孤立的 break/continue 语句。
        """
        for _pass in range(max_passes):
            try:
                compile(code, '<check>', 'exec')
                return code
            except SyntaxError as e:
                err_msg = str(e)
                if "'break' outside loop" not in err_msg and "'continue' not properly in loop" not in err_msg:
                    return code
                if e.lineno is None:
                    return code

                lines = code.split('\n')
                idx = e.lineno - 1
                if idx < 0 or idx >= len(lines):
                    return code

                stripped = lines[idx].strip()
                if stripped in ('break', 'continue') or stripped.startswith('break ') or stripped.startswith('continue '):
                    indent = len(lines[idx]) - len(lines[idx].lstrip())
                    lines[idx] = ' ' * indent + '# [auto-removed] ' + stripped
                    logger.info(f"[orphaned-break修复] 行{e.lineno}: 注释掉循环外的 {stripped}")
                    code = '\n'.join(lines)
                else:
                    return code

        return code

    def fix_python_indentation(self, code: str) -> str:
        """通用缩进修复（基于缩进栈归一化）

        处理：缩进不一致（混用空格数）、冒号后没缩进、代码块缩进层级错误。
        使用缩进栈跟踪每个级别，规范化为4的倍数。
        """
        lines = code.split('\n')
        fixed_lines = []

        indent_stack = [0]
        expected_indent = 0

        dedent_keywords = ['return ', 'return', 'break', 'continue', 'pass', 'raise ']

        for line in lines:
            stripped = line.strip()

            if not stripped:
                fixed_lines.append('')
                continue

            if stripped.startswith('#'):
                fixed_lines.append(' ' * expected_indent + stripped)
                continue

            original_indent = len(line) - len(line.lstrip())

            is_continuation = any(stripped.startswith(kw) for kw in ['else:', 'elif ', 'except', 'finally:'])

            if is_continuation and len(indent_stack) > 1:
                current_indent = indent_stack[-2] if len(indent_stack) > 1 else 0
                fixed_lines.append(' ' * current_indent + stripped)
                expected_indent = current_indent + 4
                continue

            if original_indent < expected_indent and expected_indent > 0:
                if any(stripped.startswith(kw) for kw in ['def ', 'class ', 'async def ']):
                    if original_indent == 0:
                        current_indent = 0
                        indent_stack = [0]
                        expected_indent = 0
                    else:
                        current_indent = expected_indent
                else:
                    current_indent = expected_indent
            else:
                current_indent = (original_indent // 4) * 4

            while indent_stack and current_indent < indent_stack[-1]:
                indent_stack.pop()

            if current_indent > indent_stack[-1]:
                indent_stack.append(current_indent)

            fixed_lines.append(' ' * current_indent + stripped)

            if stripped.endswith(':') and not stripped.startswith('#'):
                expected_indent = current_indent + 4
            elif any(stripped.startswith(kw) or stripped == kw.strip() for kw in dedent_keywords):
                expected_indent = current_indent
            else:
                expected_indent = current_indent

        result = '\n'.join(fixed_lines)
        logger.info(f"缩进修复完成，原始行数: {len(lines)}，修复后行数: {len(fixed_lines)}")
        return result

    # ================================================================
    #  领域特定修复（DataMerge公式代码）
    # ================================================================

    def fix_for_loop_body_indentation(self, code: str) -> str:
        """修复for循环体缩进脱离

        AI常见错误：for i in range(n_rows): 后的列赋值代码
        缩进从8空格降回4空格，导致代码跑到循环外面。
        """
        lines = code.split('\n')

        # 找到 for i in range(n_rows):
        for_line_idx = None
        for_indent = None
        for idx, line in enumerate(lines):
            stripped = line.strip()
            if re.match(r'for\s+i\s+in\s+range\s*\(\s*n_rows\s*\)', stripped):
                for_line_idx = idx
                for_indent = len(line) - len(line.lstrip())
                break

        if for_line_idx is None:
            return code

        expected_body_indent = for_indent + 4
        col_marker_pattern = re.compile(r'#\s*\S*列\s*\(\s*\d+\s*\)')

        # 找到for之后第一个非空行
        first_code_after_for = None
        for idx in range(for_line_idx + 1, len(lines)):
            if lines[idx].strip():
                first_code_after_for = idx
                break

        if first_code_after_for is None:
            return code

        first_indent = len(lines[first_code_after_for]) - len(lines[first_code_after_for].lstrip())

        drop_start = None
        if first_indent == expected_body_indent:
            # Case A: 循环开头缩进正确，后面掉回去了
            for idx in range(first_code_after_for + 1, len(lines)):
                stripped = lines[idx].strip()
                if not stripped:
                    continue
                current_indent = len(lines[idx]) - len(lines[idx].lstrip())
                if current_indent == for_indent and (
                    col_marker_pattern.search(stripped) or
                    'main_df.iloc[i]' in stripped or
                    'ws.cell(row=r' in stripped or
                    '.iloc[i]' in stripped
                ):
                    drop_start = idx
                    break
        elif first_indent == for_indent:
            # Case B: 整个循环体都在for_indent级别
            stripped = lines[first_code_after_for].strip()
            if any(var in stripped for var in ['= i +', 'iloc[i]', 'row=r', '列(']):
                drop_start = first_code_after_for
        else:
            return code

        if drop_start is None:
            return code

        # 找到循环体的结束位置
        loop_end = len(lines)
        for idx in range(drop_start, len(lines)):
            stripped = lines[idx].strip()
            if not stripped:
                continue
            current_indent = len(lines[idx]) - len(lines[idx].lstrip())
            if stripped.startswith('def ') and current_indent <= for_indent:
                loop_end = idx
                break
            if stripped.startswith('return ') and current_indent == for_indent:
                loop_end = idx
                break
            if current_indent == 0 and stripped.startswith('#') and '===' in stripped:
                loop_end = idx
                break

        # 修复：给drop_start到loop_end之间的所有行加4空格
        indent_add = ' ' * 4
        fixed_lines = lines[:drop_start]
        for idx in range(drop_start, loop_end):
            line = lines[idx]
            if not line.strip():
                fixed_lines.append('')
            else:
                fixed_lines.append(indent_add + line)
        fixed_lines.extend(lines[loop_end:])

        logger.info(
            f"修复for循环体缩进：第{drop_start + 1}-{loop_end}行，"
            f"共{loop_end - drop_start}行从{for_indent}空格修正为{expected_body_indent}空格"
        )

        return '\n'.join(fixed_lines)

    def fix_cascading_column_indentation(self, code: str) -> str:
        """修复列标记级联缩进（# XX列(N): 缩进递增）

        AI把下一列的代码嵌套在上一列的if/else分支内，
        导致缩进越来越深。将所有列块拉平到第一个列标记的缩进。
        """
        lines = code.split('\n')
        col_marker_pattern = re.compile(r'^#\s*\S*列\s*\(\s*\d+\s*\)\s*[:：]')

        col_markers = []
        for i, line in enumerate(lines):
            stripped = line.strip()
            if col_marker_pattern.match(stripped):
                indent = len(line) - len(line.lstrip())
                col_markers.append((i, indent))

        if len(col_markers) < 2:
            return code

        base_indent = col_markers[0][1]
        has_cascade = any(m[1] != base_indent for m in col_markers)

        if not has_cascade:
            return code

        logger.info(f"检测到级联缩进问题：{len(col_markers)}个列标记，"
                    f"缩进范围 {min(m[1] for m in col_markers)}-{max(m[1] for m in col_markers)} 空格，"
                    f"统一修正为 {base_indent} 空格")

        fixed_lines = lines[:col_markers[0][0]]

        code_end = len(lines)
        for j in range(col_markers[-1][0] + 1, len(lines)):
            if lines[j].strip().startswith('return '):
                code_end = j
                break

        for idx, (marker_line, marker_indent) in enumerate(col_markers):
            if idx + 1 < len(col_markers):
                block_end = col_markers[idx + 1][0]
            else:
                block_end = code_end

            indent_diff = marker_indent - base_indent

            for j in range(marker_line, block_end):
                line = lines[j]
                if not line.strip():
                    fixed_lines.append('')
                    continue
                current_indent = len(line) - len(line.lstrip())
                new_indent = max(0, current_indent - indent_diff)
                fixed_lines.append(' ' * new_indent + line.lstrip())

        for j in range(code_end, len(lines)):
            fixed_lines.append(lines[j])

        return '\n'.join(fixed_lines)

    def normalize_column_indentation(self, lines: List[str], target_base: int = 8) -> List[str]:
        """将列代码块的缩进归一化到指定基准

        优先使用列标记注释的缩进作为基准。

        Args:
            lines: 代码行列表
            target_base: 目标基准缩进（默认8空格，即for循环体内）
        """
        if not lines:
            return lines

        col_marker_pattern = re.compile(r'^\s*#\s*\S*列\s*\(\s*\d+\s*\)\s*[:：]')
        actual_base = None
        for line in lines:
            if line.strip() and col_marker_pattern.match(line.strip()):
                actual_base = len(line) - len(line.lstrip())
                break

        if actual_base is None:
            for line in lines:
                if line.strip():
                    actual_base = len(line) - len(line.lstrip())
                    break
            if actual_base is None:
                return lines

        if actual_base == target_base:
            return lines

        indent_diff = actual_base - target_base
        result = []
        for line in lines:
            if not line.strip():
                result.append('')
                continue
            current_indent = len(line) - len(line.lstrip())
            new_indent = max(0, current_indent - indent_diff)
            result.append(' ' * new_indent + line.lstrip())

        return result

    # ================================================================
    #  autopep8 工业级格式化（最终防线）
    # ================================================================

    def _fix_with_autopep8(self, code: str) -> str:
        """使用 autopep8 修复缩进和格式问题

        作为自定义修复器的最终防线：当自定义规则无法修复时，
        调用 autopep8 进行工业级格式化。
        """
        try:
            import autopep8
            fixed = autopep8.fix_code(code, options={
                'aggressive': 2,           # 激进模式，修复更多问题
                'max_line_length': 300,     # 不要折行，AI生成的代码行长无所谓
                'select': [
                    'E1',                   # 缩进类错误（E101, E111-E131）
                    'W1',                   # 缩进类警告（W191 tab→space）
                ],
            })
            try:
                compile(fixed, '<check>', 'exec')
                logger.info("[autopep8] 成功修复缩进问题")
                return fixed
            except SyntaxError:
                # autopep8 只修 select 的问题后仍然失败，再尝试全量修复
                fixed_full = autopep8.fix_code(code, options={
                    'aggressive': 2,
                    'max_line_length': 300,
                })
                try:
                    compile(fixed_full, '<check>', 'exec')
                    logger.info("[autopep8] 全量修复成功")
                    return fixed_full
                except SyntaxError as e:
                    logger.warning(f"[autopep8] 全量修复后仍有语法错误: {e}")
                    return code
        except ImportError:
            logger.warning("[autopep8] 未安装 autopep8，跳过自动格式化")
            return code
        except Exception as e:
            logger.warning(f"[autopep8] 修复异常: {e}")
            return code
