# 修复达美乐70脚本的KeyError问题

## 问题

用户执行脚本时报错：
```
KeyError: '9年终奖表_Sheet1'
```

原因：`9年终奖表.xlsx` 只有表头没有数据，旧脚本会跳过这个sheet，导致公式引用时找不到。

## 修复内容

已手动修复脚本：`tenants/达美乐70/scripts/script_ab547adb18cc.py`

### 1. 修复 `convert_region_to_dataframe` 函数（第62-85行）

**修改前**：
```python
def convert_region_to_dataframe(region) -> pd.DataFrame:
    """将ExcelRegion转换为DataFrame"""
    if not region.data:
        return pd.DataFrame()  # ❌ 没有列名
    # ...
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
    # ...
```

### 2. 修复 `load_source_data` 函数（第109-126行）

**修改前**：
```python
for sheet_data in results:
    for region in sheet_data.regions:
        df = convert_region_to_dataframe(region)
        if df.empty:
            continue  # ❌ 跳过只有表头的sheet
        # ...
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

## 修复效果

### 修改前

```
加载源数据: 员工信息_Sheet1, 列: ['工号', '姓名'], 行数: 100
# 9年终奖表被跳过

执行脚本...
KeyError: '9年终奖表_Sheet1'  # ❌ 报错
```

### 修改后

```
加载源数据: 员工信息_Sheet1, 列: ['工号', '姓名'], 行数: 100
加载源数据: 9年终奖表_Sheet1, 列: ['工号', '年终奖'], 行数: 0 (只有表头)

执行脚本...
成功！  # ✅ 不再报错
```

## 使用说明

1. **脚本已修复**，可以直接使用
2. 重新执行计算任务，应该不会再报 KeyError
3. 如果 `9年终奖表` 只有表头没有数据，公式会返回 #N/A（找不到匹配值）

## 注意事项

- 这是临时修复，只修复了这一个脚本
- 如果重新训练，会生成新的脚本（已包含修复）
- 其他租户的旧脚本如果遇到同样问题，也需要类似修复

## 相关文档

- `backend/docs/修复只有表头的Excel导致KeyError的问题.md` - 详细的问题分析和解决方案
- `backend/docs/优化Excel解析器处理只有表头的情况.md` - Excel解析器的优化

## 总结

✅ 修复了 `script_ab547adb18cc.py` 中的两处问题
✅ 现在可以正确处理只有表头的Excel文件
✅ 不会再报 KeyError
✅ 添加了清晰的日志，显示"只有表头"的情况

用户现在可以重新执行计算任务了！
