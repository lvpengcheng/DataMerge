# 报表模板：文件级拆分 + Sheet 模式

## 概述

在现有报表模板系统（fill/block/zip 三种模式）基础上，增加两项能力：

1. **文件级拆分（`split_by`）**：按数据中的某一列值（如部门、公司名）将数据拆分到不同文件，每个文件内再按现有模式填充
2. **Sheet 模式**：新增第四种报表模式，按 `group_by` 分组后，每组生成一个独立 sheet

## 数据模型变更

`Template` 表新增字段：

```python
split_by = Column(String(100), default="")   # 文件级拆分字段（如：部门）
```

`report_mode` 枚举值从 `fill / block / zip` 扩展为 `fill / block / zip / sheet`。

两层分组关系：
- `split_by`：第一层，决定生成多少个文件（按此列值拆分）
- `group_by`：第二层，决定每个文件内部的数据排列方式

## 模式组合矩阵

| split_by | 模式 | 每个文件内效果 | 最终输出 |
|----------|------|--------------|---------|
| 无 | fill | 整表填充 | 1 xlsx |
| 无 | block | 每组一块纵向合并 | 1 xlsx |
| 无 | sheet | 每组一个 sheet | 1 xlsx |
| 无 | zip | 每组一个文件 | 1 zip |
| 有 | fill | 该拆分组全量填入 | 1 zip（每个拆分值一个 xlsx） |
| 有 | block | 拆分组内再按 group_by 分块 | 1 zip |
| 有 | sheet | 拆分组内每个 group_by 值一个 sheet | 1 zip |
| 有 | zip | 禁用组合，`split_by` 覆盖 zip 的 `group_by` | 1 zip（等同于按 split_by 的 zip） |

当 `split_by` 存在时，输出一律为 zip 文件，每个拆分值对应一个 xlsx。

## 核心实现

### 1. `_generate_sheet()` — 新增 sheet 模式

位置：`backend/utils/aspose_helper.py`

逻辑：
1. 提取主数据源，按 `group_by` 分组
2. 以模板的第一个 sheet 为模板
3. 对每个分组：复制模板 sheet → SmartMarker 填充该组数据 → 设置 sheet 名
4. 删除原始模板 sheet
5. 保存输出

Sheet 命名规则：
- 取 `group_by` 列的分组值作为 sheet 名
- 截断至 31 字符（Excel 限制）
- 替换非法字符 `\ / ? * [ ] :` 为 `_`
- 重名时追加序号 `_2`, `_3`

### 2. `_generate_with_split()` — 文件级拆分包装器

位置：`backend/utils/aspose_helper.py`

逻辑：
1. 按 `split_by` 列对数据做第一层分组
2. 对每个拆分组，调用对应模式的生成函数（fill/block/sheet）
3. 每个拆分组生成一个临时 xlsx
4. 所有 xlsx 打包进 zip
5. 文件命名使用 `file_name_rule` 模板，`split_by` 的值自动注入为可用变量

### 3. `generate_from_template()` 入口修改

在现有入口函数中增加判断：
- 如果 `split_by` 存在且有效 → 走 `_generate_with_split()` 包装器
- 如果 `split_by` 为空 → 走现有逻辑（fill/block/zip/sheet）

### 4. split_by + zip 的处理

当 `split_by` 有值且模式为 `zip` 时，让 `split_by` 覆盖 `group_by` 的语义：
- 等同于按 `split_by` 字段做 zip 模式
- 避免两层都拆文件的混乱

## API 变更

### 模板 CRUD

`backend/admin/router.py` 中的创建/更新模板接口增加 `split_by` 参数：

```python
split_by: str = Form("")
```

响应中也返回 `split_by` 字段。

### 报表生成接口

`generate_from_template()` 调用处透传 `split_by` 参数。

## 前端变更

### 模板管理弹窗（新建/编辑）

1. 报表模式 select 增加选项：`sheet — 分组多Sheet（每组一个Sheet）`
2. 新增 `split_by` 输入框，位于报表模式之后
   - 标签：`文件拆分字段`
   - placeholder：`如：部门（留空则不拆分文件）`
   - 提示文字：`按此列值将数据拆分到不同文件中，拆分后自动打包为 zip`
   - 所有模式下均可见（不受模式切换控制）
3. 当选择 `sheet` 模式时，显示 `group_by` 字段（与 block 类似），隐藏 `skip_rows`

### 模板列表

报表模式列增加 `sheet` 标签样式。若有 `split_by` 配置，在模式标签旁显示拆分字段。

## 数据库迁移

`backend/database/init_db.py` 中增加 `split_by` 列的迁移逻辑，与现有 `report_mode` 等字段的迁移方式一致（ALTER TABLE ADD COLUMN IF NOT EXISTS）。

## 约束与限制

- Sheet 名最长 31 字符，自动截断
- Sheet 名不能含 `\ / ? * [ ] :`，自动替换为 `_`
- 性能：分组数过多时（如 >100 个 sheet），生成和打开都会较慢，暂不做限制，日志中记录分组数
