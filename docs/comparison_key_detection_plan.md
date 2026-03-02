# Excel 对比主键智能检测优化方案

## 问题分析

当前 `compare_excel_files()` 硬编码使用"工号"和"中文姓名"作为主键，导致：
1. 非人力资源场景（如订单、库存、财务）无法对比
2. 缺少工号或姓名列时对比失败
3. 无法适应不同业务场景的主键需求

## 优化方案

### 方案 1：智能主键检测（推荐）

#### 1.1 主键候选列识别规则

按优先级检测以下类型的列作为主键候选：

**高优先级（唯一性强）：**
- 包含关键词：`["ID", "id", "编号", "工号", "订单号", "单号", "流水号", "序号", "code", "number"]`
- 数据类型：整数或字符串
- 唯一性：值的唯一率 > 95%

**中优先级（辅助标识）：**
- 包含关键词：`["姓名", "名称", "name", "title", "日期", "date", "时间", "time"]`
- 数据类型：字符串或日期
- 唯一性：值的唯一率 > 50%

**低优先级（组合键）：**
- 前 3 列（通常是标识列）
- 数据类型：非数值型

#### 1.2 主键选择策略

```python
def detect_primary_keys(df: pd.DataFrame) -> List[str]:
    """智能检测主键列

    策略：
    1. 优先选择单列唯一性 > 95% 的列
    2. 如果没有，选择 2-3 列组合使唯一性 > 95%
    3. 兜底：使用前 2 列
    """
    candidates = []

    # 1. 检测高优先级列
    for col in df.columns:
        if _is_high_priority_key(col, df[col]):
            uniqueness = df[col].nunique() / len(df)
            candidates.append((col, uniqueness, 'high'))

    # 2. 单列唯一性检查
    for col, uniqueness, priority in candidates:
        if uniqueness > 0.95:
            return [col]  # 单列足够

    # 3. 组合键检查
    if len(candidates) >= 2:
        # 尝试前 2 个候选列组合
        combo_key = candidates[0][0] + "_" + candidates[1][0]
        combo_uniqueness = df[[candidates[0][0], candidates[1][0]]].drop_duplicates().shape[0] / len(df)
        if combo_uniqueness > 0.95:
            return [candidates[0][0], candidates[1][0]]

    # 4. 兜底：使用前 2 列
    return list(df.columns[:2])
```

#### 1.3 实现步骤

**Step 1: 添加主键检测函数**
```python
# backend/utils/excel_comparator.py

def _is_high_priority_key(col_name: str, col_data: pd.Series) -> bool:
    """判断列是否为高优先级主键候选"""
    key_keywords = ["ID", "id", "编号", "工号", "订单号", "单号", "流水号", "code", "number"]
    return any(kw in col_name for kw in key_keywords)

def _calculate_uniqueness(df: pd.DataFrame, cols: List[str]) -> float:
    """计算列组合的唯一性"""
    if len(cols) == 1:
        return df[cols[0]].nunique() / len(df)
    else:
        return df[cols].drop_duplicates().shape[0] / len(df)

def detect_primary_keys(df: pd.DataFrame) -> List[str]:
    """智能检测主键列"""
    # 实现上述策略
    pass
```

**Step 2: 修改 compare_excel_files 函数签名**
```python
def compare_excel_files(
    result_file: str,
    expected_file: str,
    output_file: Optional[str] = None,
    primary_keys: Optional[List[str]] = None,  # 新增：允许手动指定
    auto_detect_keys: bool = True  # 新增：是否自动检测
) -> Dict[str, Any]:
    """对比两个Excel文件的差异

    Args:
        primary_keys: 手动指定主键列名列表，如 ["工号", "姓名"] 或 ["订单号"]
        auto_detect_keys: 如果 primary_keys 为 None，是否自动检测主键
    """
```

**Step 3: 主键检测和匹配逻辑**
```python
# 1. 确定主键
if primary_keys is None and auto_detect_keys:
    # 自动检测（两个文件都检测，取交集）
    expected_keys = detect_primary_keys(expected_df)
    result_keys = detect_primary_keys(result_df)
    primary_keys = list(set(expected_keys) & set(result_keys))
    if not primary_keys:
        # 兜底：使用第一列
        primary_keys = [expected_df.columns[0]]
    logger.info(f"自动检测主键: {primary_keys}")
elif primary_keys is None:
    # 使用默认主键（向后兼容）
    primary_keys = ["工号", "中文姓名"]
    logger.info(f"使用默认主键: {primary_keys}")

# 2. 标准化主键值
for key_col in primary_keys:
    if key_col in expected_df.columns:
        expected_df[f"标准化_{key_col}"] = expected_df[key_col].apply(_standardize_key_value)
    if key_col in result_df.columns:
        result_df[f"标准化_{key_col}"] = result_df[key_col].apply(_standardize_key_value)

# 3. 创建复合匹配键
standardized_keys = [f"标准化_{k}" for k in primary_keys]
expected_df["匹配键"] = expected_df[standardized_keys].astype(str).agg("_".join, axis=1)
result_df["匹配键"] = result_df[standardized_keys].astype(str).agg("_".join, axis=1)

# 4. 合并对比
merged_df = pd.merge(
    expected_df, result_df,
    on="匹配键",
    how="outer",
    suffixes=("_expected", "_result"),
    indicator=True
)
```

**Step 4: 标准化函数**
```python
def _standardize_key_value(value) -> str:
    """标准化主键值（去空格、统一大小写、补零等）"""
    if pd.isna(value):
        return ""

    value_str = str(value).strip()

    # 如果是数字字符串，补零到 8 位
    if value_str.isdigit():
        return value_str.zfill(8)

    # 其他情况：去空格、转小写
    return value_str.lower().replace(" ", "")
```

---

### 方案 2：配置化主键（备选）

在训练时让用户指定主键，保存到 `script_info.json`：

```json
{
  "script_id": "script_xxx",
  "comparison_keys": ["订单号", "日期"],
  "key_standardization": {
    "订单号": "zfill_8",
    "日期": "date_format"
  }
}
```

调用对比时从 `script_info` 读取：
```python
script_info = storage_manager.get_active_script(tenant_id)
primary_keys = script_info.get("comparison_keys", ["工号", "中文姓名"])
compare_excel_files(result_file, expected_file, primary_keys=primary_keys)
```

---

## 实施建议

**阶段 1（快速修复）：**
1. 添加 `primary_keys` 参数，允许手动指定
2. 默认值保持 `["工号", "中文姓名"]`，向后兼容
3. 如果指定列不存在，回退到第一列

**阶段 2（智能检测）：**
1. 实现 `detect_primary_keys()` 自动检测
2. 添加 `auto_detect_keys=True` 参数
3. 日志输出检测到的主键，便于调试

**阶段 3（配置化）：**
1. 训练时保存主键配置到 `script_info`
2. 对比时自动读取配置

---

## 测试场景

1. **人力资源场景**：工号 + 姓名（当前场景）
2. **订单场景**：订单号（单列主键）
3. **库存场景**：仓库编号 + SKU（组合键）
4. **财务场景**：日期 + 科目代码（组合键）
5. **无明显主键**：使用前 2 列兜底

---

## 预期效果

- 支持任意业务场景的 Excel 对比
- 自动适应不同的主键结构
- 向后兼容现有人力资源场景
- 提升训练成功率（不再因主键问题失败）
