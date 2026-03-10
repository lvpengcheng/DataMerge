# 修复只有表头的Excel导致脚本执行KeyError的问题

## 问题描述

用户反馈：在进行计算时，如果一个表只有列头（没有数据行），脚本在执行过程中会尝试访问一个名为 `'9年终奖表_Sheet1'` 的工作表时发生 KeyError（键错误），表明该工作表在数据源中不存在或名称不匹配，导致程序中断。

## 问题分析

### 问题链路

1. **Excel解析阶段**：
   - Excel只有表头没有数据
   - `excel_parser` 正常解析，返回一个有表头但data为空列表的region

2. **DataFrame转换阶段**：
   - `convert_region_to_dataframe(region)` 被调用
   - 原有逻辑：如果 `region.data` 为空，返回 `pd.DataFrame()`（没有列名的空DataFrame）

3. **数据加载阶段**：
   - `load_source_data` 函数检查 `if df.empty: continue`
   - 因为DataFrame为空，这个sheet被跳过，不添加到 `source_data` 字典

4. **脚本执行阶段**：
   - AI生成的公式引用这个sheet（如 `'9年终奖表_Sheet1'`）
   - 尝试从 `source_data` 字典中获取这个sheet
   - **KeyError**：字典中不存在这个键

### 根本原因

**原有逻辑的问题**：

```python
# convert_region_to_dataframe (错误)
def convert_region_to_dataframe(region) -> pd.DataFrame:
    if not region.data:
        return pd.DataFrame()  # ❌ 返回没有列名的空DataFrame
    # ...

# load_source_data (错误)
df = convert_region_to_dataframe(region)
if df.empty:
    continue  # ❌ 跳过只有表头的sheet
```

**问题**：
- 只有表头的sheet被完全忽略
- AI生成的公式可能引用这个sheet
- 执行时找不到这个sheet，报KeyError

## 解决方案

### 1. 修复 `convert_region_to_dataframe` 函数

在三个文件中修复：
- `ai_engine/formula_code_generator.py`
- `ai_engine/modular_code_generator.py`
- `ai_engine/excel_formula_builder.py`

**修改前**：
```python
def convert_region_to_dataframe(region) -> pd.DataFrame:
    """将ExcelRegion转换为DataFrame"""
    if not region.data:
        return pd.DataFrame()  # ❌ 没有列名
    col_letter_to_name = {v: k for k, v in region.head_data.items()}
    converted_data = []
    for row in region.data:
        new_row = {}
        for col_letter, value in row.items():
            col_name = col_letter_to_name.get(col_letter, col_letter)
            new_row[col_name] = value
        converted_data.append(new_row)
    columns = list(region.head_data.keys())
    return pd.DataFrame(converted_data, columns=columns)
```

**修改后**：
```python
def convert_region_to_dataframe(region) -> pd.DataFrame:
    """将ExcelRegion转换为DataFrame

    即使没有数据行，也会返回带列名的空DataFrame，
    这样可以避免在引用只有表头的sheet时出现KeyError
    """
    # 获取列名映射
    col_letter_to_name = {v: k for k, v in region.head_data.items()}
    columns = list(region.head_data.keys())

    # ✅ 如果没有数据，返回带列名的空DataFrame
    if not region.data:
        return pd.DataFrame(columns=columns)

    # 转换数据行
    converted_data = []
    for row in region.data:
        new_row = {}
        for col_letter, value in row.items():
            col_name = col_letter_to_name.get(col_letter, col_letter)
            new_row[col_name] = value
        converted_data.append(new_row)

    return pd.DataFrame(converted_data, columns=columns)
```

### 2. 修复 `load_source_data` 函数

在 `ai_engine/formula_code_generator.py` 中：

**修改前**：
```python
for sheet_data in results:
    for region in sheet_data.regions:
        df = convert_region_to_dataframe(region)
        if df.empty:
            continue  # ❌ 跳过只有表头的sheet

        sheet_name = f"{file_base}_{sheet_data.sheet_name}"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]

        source_data[sheet_name] = {
            "df": df,
            "columns": list(df.columns)
        }
        print(f"加载源数据: {sheet_name}, 列: {list(df.columns)}, 行数: {len(df)}")
```

**修改后**：
```python
for sheet_data in results:
    for region in sheet_data.regions:
        df = convert_region_to_dataframe(region)
        # ✅ 即使DataFrame为空（只有表头没有数据），也要添加到source_data
        # 这样可以避免公式引用只有表头的sheet时出现KeyError
        # 只有当DataFrame连列名都没有时才跳过
        if df.empty and len(df.columns) == 0:
            continue

        sheet_name = f"{file_base}_{sheet_data.sheet_name}"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]

        source_data[sheet_name] = {
            "df": df,
            "columns": list(df.columns)
        }
        if len(df) > 0:
            print(f"加载源数据: {sheet_name}, 列: {list(df.columns)}, 行数: {len(df)}")
        else:
            print(f"加载源数据: {sheet_name}, 列: {list(df.columns)}, 行数: 0 (只有表头)")
```

### 3. 修复 `excel_formula_builder.py` 中的相同逻辑

在 `ai_engine/excel_formula_builder.py` 的 `load_source_data` 方法中应用相同的修复。

## 修复效果

### 修改前

```python
# Excel: 9年终奖表.xlsx (只有表头，没有数据)
# | 工号 | 姓名 | 年终奖 |

# 解析结果
df = convert_region_to_dataframe(region)
# df = pd.DataFrame()  # 空DataFrame，没有列名

# 加载数据
if df.empty:
    continue  # 跳过这个sheet

# source_data = {}  # 不包含 '9年终奖表_Sheet1'

# 执行脚本
result_df = source_data['9年终奖表_Sheet1']  # ❌ KeyError!
```

### 修改后

```python
# Excel: 9年终奖表.xlsx (只有表头，没有数据)
# | 工号 | 姓名 | 年终奖 |

# 解析结果
df = convert_region_to_dataframe(region)
# df = pd.DataFrame(columns=['工号', '姓名', '年终奖'])  # ✅ 有列名的空DataFrame

# 加载数据
if df.empty and len(df.columns) == 0:
    continue  # 只有完全空的DataFrame才跳过
# ✅ 这个sheet会被添加到source_data

# source_data = {
#     '9年终奖表_Sheet1': {
#         'df': pd.DataFrame(columns=['工号', '姓名', '年终奖']),
#         'columns': ['工号', '姓名', '年终奖']
#     }
# }

# 执行脚本
result_df = source_data['9年终奖表_Sheet1']  # ✅ 成功获取，不会报错
# 公式引用这个sheet时，会返回空值（因为没有数据行）
```

## 数据结构对比

### 只有表头的DataFrame

```python
# 修改前
df = pd.DataFrame()
print(df.empty)      # True
print(df.columns)    # Index([], dtype='object')
print(len(df))       # 0

# 修改后
df = pd.DataFrame(columns=['工号', '姓名', '年终奖'])
print(df.empty)      # True (仍然是空的)
print(df.columns)    # Index(['工号', '姓名', '年终奖'], dtype='object')
print(len(df))       # 0
```

### 判断逻辑

```python
# 修改前：跳过所有空DataFrame
if df.empty:
    continue  # ❌ 会跳过只有表头的sheet

# 修改后：只跳过连列名都没有的DataFrame
if df.empty and len(df.columns) == 0:
    continue  # ✅ 只跳过完全空的DataFrame
```

## 使用场景

### 场景1：模板文件（只有表头）

```
9年终奖表.xlsx:
| 工号 | 姓名 | 年终奖 |
(没有数据行)
```

**修改前**：
- 这个sheet被跳过
- 公式引用时报KeyError

**修改后**：
- 这个sheet被加载（有列名的空DataFrame）
- 公式引用时返回空值（不报错）

### 场景2：部分文件有数据，部分只有表头

```
员工信息.xlsx:
| 工号 | 姓名 | 部门 |
| 001  | 张三 | 技术 |
| 002  | 李四 | 销售 |

9年终奖表.xlsx:
| 工号 | 姓名 | 年终奖 |
(没有数据行)
```

**修改前**：
- 员工信息被加载
- 9年终奖表被跳过
- 公式引用9年终奖表时报KeyError

**修改后**：
- 员工信息被加载（2行数据）
- 9年终奖表被加载（0行数据，但有列名）
- 公式引用9年终奖表时返回空值（不报错）

### 场景3：VLOOKUP引用只有表头的sheet

```python
# AI生成的公式
formula = "=VLOOKUP(A2,'9年终奖表_Sheet1'!A:C,3,0)"

# 修改前
# KeyError: '9年终奖表_Sheet1'

# 修改后
# 公式正常写入Excel
# 执行时返回 #N/A（找不到匹配值，因为源表没有数据）
```

## 日志输出

### 修改前

```
加载源数据: 员工信息_Sheet1, 列: ['工号', '姓名', '部门'], 行数: 100
# 9年终奖表被跳过，没有日志
```

### 修改后

```
加载源数据: 员工信息_Sheet1, 列: ['工号', '姓名', '部门'], 行数: 100
加载源数据: 9年终奖表_Sheet1, 列: ['工号', '姓名', '年终奖'], 行数: 0 (只有表头)
```

## 兼容性

### 向后兼容

✅ 完全向后兼容：
- 有数据的sheet：行为不变
- 只有表头的sheet：现在会被加载（之前被跳过）
- 完全空的sheet（连表头都没有）：仍然被跳过

### 对现有脚本的影响

✅ 不影响现有正常运行的脚本：
- 如果所有源文件都有数据，行为完全不变
- 如果某些源文件只有表头，现在不会报错了

## 修改文件

1. `ai_engine/formula_code_generator.py`
   - `convert_region_to_dataframe` 函数（第1607-1630行）
   - `load_source_data` 函数（第1664-1681行）

2. `ai_engine/modular_code_generator.py`
   - `convert_region_to_dataframe` 函数（第647-673行）

3. `ai_engine/excel_formula_builder.py`
   - `_convert_region_to_dataframe` 方法（第125-147行）
   - `load_source_data` 方法（第90-114行）

## 测试建议

### 1. 测试只有表头的Excel

创建一个只有表头的Excel文件：
```
9年终奖表.xlsx:
| 工号 | 姓名 | 年终奖 |
```

运行训练，验证：
- 不会报KeyError
- 日志显示：`加载源数据: 9年终奖表_Sheet1, 列: ['工号', '姓名', '年终奖'], 行数: 0 (只有表头)`

### 2. 测试公式引用只有表头的sheet

创建规则引用只有表头的sheet：
```
结果列：年终奖
公式：=VLOOKUP(工号,'9年终奖表_Sheet1'!A:C,3,0)
```

运行训练，验证：
- 脚本正常生成
- 脚本正常执行
- 结果列显示 #N/A（因为源表没有数据）

### 3. 测试混合场景

创建多个源文件：
- 员工信息.xlsx（有数据）
- 9年终奖表.xlsx（只有表头）

运行训练，验证：
- 两个sheet都被加载
- 脚本正常执行
- 不会报KeyError

## 总结

✅ 修复了只有表头的Excel导致脚本执行KeyError的问题
✅ 即使sheet只有表头没有数据，也会被加载到source_data
✅ 公式引用只有表头的sheet时不会报错（返回空值）
✅ 完全向后兼容，不影响现有脚本
✅ 添加了清晰的日志，区分有数据和只有表头的情况

现在系统可以正确处理只有表头的Excel文件，不会因为KeyError导致脚本执行失败！
