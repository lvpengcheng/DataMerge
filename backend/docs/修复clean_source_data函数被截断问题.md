# 修复clean_source_data函数被截断的问题

## 问题描述

用户发现生成的代码中：
1. `clean_source_data` 函数只有函数定义，没有函数体
2. 函数体的注释缩进错误（8个空格而不是4个）

```python
def clean_source_data(source_data):

        # 规则1: 员工信息表内的部门为市场部的不参与计算
        # ← 缩进8，错误！应该是4


def fill_result_sheets(wb, source_sheets, ...):
    ...
```

## 问题追踪

### 1. AI响应检查 ✅

AI的原始响应是正确的：
```python
def clean_source_data(source_data):
    """清洗源数据，应用过滤规则"""
    cleaned = {}
    for key, info in source_data.items():
        df = info["df"].copy()
        columns = info["columns"]

        # 规则1: 员工信息表内的部门为市场部的不参与计算
        if key == "01_员工信息_员工信息修改后":
            if "部门" in df.columns:
                df = df[df["部门"] != "市场部"]

        cleaned[key] = {"df": df, "columns": columns}
    return cleaned
```

### 2. 代码提取检查 ✅

`_extract_fill_result_sheets_function` 方法提取的代码也是正确的，缩进完整。

### 3. 代码清理检查 ❌ **问题所在**

`_clean_before_function_def` 方法在清理时出现问题：

**旧逻辑（错误）：**
```python
# 找到 def fill_result_sheets 的位置
func_start = 17  # 假设在第17行

# 遍历第1-16行，逐行判断是否保留
for i in range(func_start):
    stripped = lines[i].strip()
    # 只保留：空行、注释、import、def、class、常量
    if (not stripped
        or stripped.startswith('#')
        or stripped.startswith('import ')
        or stripped.startswith('def ')  # ← 只保留函数定义行！
        or ...):
        clean_prefix.append(lines[i])
    else:
        # 删除其他代码（如 cleaned = {}、for key...、return cleaned）
        logger.info(f"清理游离代码: {stripped}")
```

**问题：**
- 第1行：`def clean_source_data(source_data):` - 保留（匹配 `def `）
- 第2行：`"""清洗源数据..."""` - **删除**（不匹配任何规则）
- 第3行：`cleaned = {}` - **删除**（不匹配任何规则）
- 第4行：`for key, info...` - **删除**（不匹配任何规则）
- ...
- 第8行：`# 规则1: ...` - 保留（匹配 `#`）
- ...

结果：只保留了函数定义行和注释行，函数体被全部删除！

## 解决方案

修改 `_clean_before_function_def` 方法，识别完整的函数块：

```python
def _clean_before_function_def(self, code: str) -> str:
    """清理函数定义之前的垃圾代码

    注意：现在代码中可能包含clean_source_data函数，需要保留完整的函数块。
    """
    lines = code.split('\n')

    # 找到 def fill_result_sheets 的位置
    func_start = -1
    for i, line in enumerate(lines):
        if line.strip().startswith('def fill_result_sheet'):
            func_start = i
            break

    if func_start <= 0:
        return code

    # 保留函数定义之前的合法代码
    clean_prefix = []
    i = 0
    while i < func_start:
        stripped = lines[i].strip()

        # 保留：空行、注释、import
        if (not stripped
            or stripped.startswith('#')
            or stripped.startswith('import ')
            or stripped.startswith('from ')):
            clean_prefix.append(lines[i])
            i += 1
            continue

        # 保留：常量赋值
        if (re.match(r'^[A-Z_][A-Z_0-9]*\s*=', stripped)
            or re.match(r'^TXT_\w+\s*=', stripped)):
            clean_prefix.append(lines[i])
            i += 1
            continue

        # 保留：完整的函数或类定义（包括函数体）← 关键修改
        if stripped.startswith('def ') or stripped.startswith('class '):
            # 找到函数/类的结束位置
            func_end = i + 1
            base_indent = len(lines[i]) - len(lines[i].lstrip())

            while func_end < func_start:
                line = lines[func_end]
                # 如果是空行，继续
                if not line.strip():
                    func_end += 1
                    continue
                # 如果缩进大于函数定义行，说明还在函数体内
                current_indent = len(line) - len(line.lstrip())
                if current_indent > base_indent:
                    func_end += 1
                    continue
                # 如果缩进等于或小于函数定义行，说明函数结束
                break

            # 保留整个函数块
            for j in range(i, func_end):
                clean_prefix.append(lines[j])
            i = func_end
            continue

        # 跳过游离代码
        logger.info(f"清理函数定义前的游离代码: 行{i+1}: {stripped[:60]}")
        i += 1

    return '\n'.join(clean_prefix + lines[func_start:])
```

**关键改进：**
1. 使用 `while` 循环代替 `for` 循环，可以跳过多行
2. 当遇到 `def ` 或 `class ` 时，查找整个函数/类块的结束位置
3. 通过缩进判断函数体是否结束：
   - 缩进 > 函数定义行 → 还在函数体内
   - 缩进 ≤ 函数定义行 → 函数结束
4. 保留整个函数块（从 `def` 到函数结束）

## 验证结果

```
清理前行数: 126
清理后行数: 126
删除了 0 行  ← 完美！没有删除任何行

clean_source_data函数位置: 第302行

函数及其前后5行:
行 302 (缩进 0): def clean_source_data(source_data):
行 303 (缩进 4):     """清洗源数据，应用过滤规则"""
行 304 (缩进 4):     cleaned = {}
行 305 (缩进 4):     for key, info in source_data.items():
行 306 (缩进 8):         df = info["df"].copy()
行 307 (缩进 8):         columns = info["columns"]
行 308 (缩进 0):
行 309 (缩进 8):         # 规则1: 员工信息表内的部门为市场部的不参与计算
行 310 (缩进 8):         if key == "01_员工信息_员工信息修改后":
行 311 (缩进12):             if "部门" in df.columns:
行 312 (缩进16):                 df = df[df["部门"] != "市场部"]
行 313 (缩进 0):
行 314 (缩进 8):         cleaned[key] = {"df": df, "columns": columns}
行 315 (缩进 4):     return cleaned
```

所有缩进都正确了！✅

## 相关修复

这是数据清洗功能的第三个修复：

1. **规则提取器修复** - 支持所有中文数字格式
2. **多步分析模式提示词修复** - 添加数据清洗规则和函数要求
3. **代码清理方法修复** - 保留完整的函数块 ← 本次修复

## 总结

✅ 修复了 `_clean_before_function_def` 方法的逻辑缺陷
✅ 现在能正确保留完整的 `clean_source_data` 函数
✅ 函数体不会被误删除
✅ 缩进完全正确

现在可以重新训练rex104，验证完整的数据清洗功能。
