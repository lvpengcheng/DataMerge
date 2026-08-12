# 智能组表（Smart Assemble）设计文档

日期：2026-08-12
状态：已确认（rex 逐节批准）

## 1. 需求背景

新增"智能组表"功能：用户上传源文件（多文件、全部可见 sheet）和模板文件（只读激活 sheet），
结合组表规则文件，AI 分析源/模板结构并生成填充代码，把源数据组合填充进模板的数据行，
输出保留模板样式的新 Excel（原版带 `源_` sheet 与公式 + 纯值版双版本）。

## 2. 总体架构与数据流

```
┌─ 管理页签(admin.html) ──────────────────────────────┐
│  管理组表规则文件: 上传(文字说明+可选结构化示例)      │
│  列表/下载/删除 + 匹配知识库查看/删除/恢复            │
│  全局+按租户两级, 权限 tools.assemble.manage          │
└──────────────────────┬───────────────────────────────┘
                       │
┌─ 智能小工具/智能组表页签(tools.html) ─────────────────┴─────────┐
│ 左侧: 规则选择 + 源文件(多) + 模板文件 + 强制重新匹配☑ + 开始    │
│ 右侧: 分析/执行日志(SSE) + 结果下载(原版+纯值版) + 结果确认条    │
└──────────────────────┬──────────────────────────────────────────┘
                       │ 上传
┌──────────────────────▼──────────────────────────────────────────┐
│ 1. 解析: excel_parser 读源(全部可见sheet, 数据>3行取前3行样例,    │
│          <3行取全部; 样例脱敏) + 读模板(激活sheet)                │
│          → 生成"结构签名"                                         │
│ 2. 查代码存档: 签名命中 且 未勾强制 → 跳过AI直接执行              │
│ 3. 未命中: 查字段知识库 → 已覆盖字段直接采用(日志标注)            │
│          → 未覆盖字段 AI 分析(右侧日志展示) → 结构化映射清单      │
│          → 回写知识库 → 生成 fill 骨架代码 → 存档 → 执行          │
│ 4. 执行: 子进程 + CodeSandbox → 原版(带源_sheet+公式) + 纯值版    │
│ 5. 历史: 任务记录 + 结果文件落盘                                  │
└──────────────────────────────────────────────────────────────────┘
```

## 3. 数据模型（新表）

### assemble_rules（组表规则文件元信息）
- `id` PK 自增
- `name` String(200) 规则名称
- `scope` String(20) —— `global`（全局）/ `tenant`（按租户）
- `tenant_id` String(50) nullable（scope=tenant 时有效）
- `file_names` JSON 规则文件原始文件名列表
- `description` Text nullable 说明
- `uploader_id` FK users.id
- `created_at` / `updated_at`

文件实体：全局 → `global_assets/assemble_rules/`，按租户 → `tenants/{tenant_id}/assemble_rules/`。
文件命名 `{rule_id}_{原始文件名}` 防重。

### assemble_tasks（任务历史）
- `id` PK 自增
- `tenant_id` String(50)
- `user_id` FK users.id
- `rule_id` FK assemble_rules.id nullable
- `signature` String(64) 本次任务总签名（规则哈希+源结构签名+模板结构签名）
- `source_signature` / `template_signature` String(64) 分签名（调试/知识库锚定用）
- `code_path` String(300) nullable 执行代码路径（存档或新生成）
- `used_mapping_ids` JSON 命中的知识库条目 id 列表（结果反馈时精确停用）
- `status` String(20) —— pending/analyzing/generating/executing/completed/error
- `ai_provider` String(50) nullable
- `matched_from_cache` Boolean 是否命中代码存档
- `output_files` JSON nullable 结果文件名列表
- `error` Text nullable
- `created_at` / `completed_at` DateTime nullable

### assemble_field_mappings（字段级匹配知识库）
- `id` PK 自增
- `source_column` String(100) 源列名
- `target_column` String(100) 模板列名
- `template_signature` String(64) 该映射生效的模板结构签名（锚点）
- `hit_count` Integer 默认 1 命中/确认次数
- `source_task_id` FK assemble_tasks.id nullable 首次来源任务
- `status` String(20) 默认 `active`；`review_needed`（被反馈标错，查询时不再自动采用）
- `created_at` / `updated_at`

## 4. 组表规则文件管理（管理页签）

位置：admin.html 新增 tab「智能组表规则」，UI 参考现有「SOP规则」页签。

- 顶部：`+ 上传规则文件` 按钮 + 说明文字（"全局大规则：所有租户可用；按客户规则：仅该客户优先使用"）
- 表格列：规则名称 / 类型 / 适用租户 / 文件数 / 上传人 / 更新时间 / 操作（下载、删除）
- 上传弹窗：规则名称、范围（全局/按租户+租户选择）、文件（多文件，md/txt/pdf/docx）、说明
- 权限：tab 与按钮 `data-perm="tools.assemble.manage"`
- 知识库子页：`assemble_field_mappings` 列表（源列/模板列/模板指纹/命中次数/状态），
  操作：删除、停用/恢复启用（review_needed 红色标识）

API（backend/api/assemble_rules.py，prefix `/api/assemble/rules`）：
- `GET /` 列表（管理端全部，含租户过滤参数）
- `POST /` 上传（multipart：name/scope/tenant_id/description/files）
- `GET /{id}/download` 下载规则文件（单文件按 file_name 参数）
- `DELETE /{id}` 删除（连文件）
- `GET /mappings` 知识库列表（管理端）
- `DELETE /mappings/{id}` 删除条目
- `PUT /mappings/{id}` 停用/启用/恢复

## 5. 智能组表执行页签

位置：tools.html 新增 tab「智能组表」（data-tab="smart-assemble"）。

### 左侧配置区
- 规则文件下拉（全局+本租户，GET /api/assemble/rules 过滤可用的）
- 源文件上传（multiple .xlsx/.xls）+ 文件列表
- 模板文件上传（单个）
- ☑ 强制重新匹配（默认不勾）
- `开始组表` 按钮（POST /api/assemble/submit）

### 右侧过程区
- 状态文本 + 进度条 + 深色日志区
- SSE 事件：`status` / `log` / `analyzing` / `thinking`（AI 推理灰色流，deepseek 思考透传）/
  `mapping`（知识库命中行）/ `complete` / `error`
- 日志展示脱敏后的源结构摘要
- 完成卡片：结果文件名 + 下载按钮（原版 + 纯值版）+ **结果确认条（✅正确/❌有误）**
- ❌ 有误 → POST /api/assemble/tasks/{id}/feedback → 停用 used_mapping_ids 条目 → 提示勾选强制重新匹配重跑

## 6. 解析与结构签名

### 源文件解析
- `parse_excel_file(fp, max_data_rows=3, read_formulas=True, calculate_formulas=False)` 全部可见 sheet
- 数据行 ≥3 取前 3 行、<3 取全部行作为样例
- 样例经过**脱敏**后才进结构 json / prompt / 日志
- 结构 json：`{文件: {sheet: {表头列名列表, 列字母, 样例数据(脱敏), 公式列, 列格式}}}`

### 模板文件解析
- `parse_excel_file(tpl, max_data_rows=5, read_formulas=True, skip_hidden_sheets=False, active_sheet_only=True)`
- 产出区域结构（表头行/数据起止行/列名/列字母/公式列/列格式）→ 固化 `_COL_MAP`

### 结构签名
- 模板签名：激活 sheet 表头列名集合 + 词表（参考 target_sheet_resolver `_key_signature`），
  **不包含 sheet 名**（模板按月改名不影响命中）
- 源签名：全部源文件的列名集合（按文件排序后的稳定序列）
- 总签名：`sha256(规则文件内容哈希 + 源签名 + 模板签名)`

## 7. 代码生成器与填充逻辑

新生成器 `backend/ai_engine/assemble_code_generator.py`（改造自 TemplateCodeGenerator，不动原类）：

- 模板解析限定激活 sheet；源解析保留全部可见 sheet 且样例脱敏
- Prompt 注入：规则文字 + 结构化示例（源列→目标列）+ 源/模板结构 json +
  **知识库已命中映射清单**（AI 不再处理已覆盖字段）
- AI 分析阶段**先输出结构化映射清单 JSON**（schema 校验：未覆盖源列→模板列 + 逻辑说明），
  一份作为 fill 生成依据，一份回写知识库
- 骨架固化：`_COL_MAP` / `_SOURCE_MAP` / `TEMPLATE_HASH` / `CLEANING_SPEC`
- `main()` 链路（复用模板模式成熟代码）：
  copy 模板 → openpyxl 保留格式 → 追加 `源_` sheet → `fill_template(wb, source_data, ...)` →
  **数据行不够时 Aspose 复制最后数据行样式向下扩展** → 格式快照刷回 → 保存
- 修复管线：`validate_and_fix_code_format` + compile 校验；失败自动带错误重试 1 次

### 结果双版本
- 原版：模板副本 + 填充内容 + `源_` sheet + 公式（复用智算 output_postprocess 后处理链）
- 纯值版：值复制版本（复用 values_only 逻辑）

## 8. 代码存档与匹配知识库（双层复用）

### 代码存档
- 路径：`tenants/{tenant_id}/assemble_scripts/{签名}.py`
- 命中：总签名一致 且 未勾强制 → 直接执行存档代码（日志标注"命中代码存档"）
- 强制：跳过存档，AI 重新生成并覆盖存档

### 字段级知识库
- 查询顺序：
  1. 精确锚定：当前模板签名 ∈ 条目锚点（template_signature 相等）→ 直接采用（日志 `[知识库] 源列 → 模板列`）
  2. 同名列兜底：源列名 = 模板列名（模板列集合内）→ 确定性采用
  3. 歧义候选：同一模板下同一源列名存在多条不同目标映射 → 不自动用，交 AI
- 写入：AI 输出映射清单后，未覆盖字段的映射回写（template_signature=当前模板签名）
- 任务记录 `used_mapping_ids`（自动采用 + AI 新增的条目都记录）

### 防污染（结果确认反馈环）
- 任务完成卡片：`✅ 结果正确` → 命中条目 hit_count+1；`❌ 结果有误` →
  该任务 `used_mapping_ids` 条目全部标 `review_needed` → 查询时不再自动采用（交 AI）→
  前端提示建议勾选"强制重新匹配"重跑
- review_needed 条目可在管理页签恢复启用或删除

## 9. 数据脱敏

新工具 `backend/utils/desensitize.py`：

- 内置敏感词表：公司、姓名、身份证、手机、电话、邮箱、地址、账号、银行卡、工资、薪资、社保、公积金
- 按**列名**包含关键字 → 该列样例数据脱敏，**保格式**（长度与数字/字母位不变）：
  - 身份证/长数字串：保留前 3 后 4，中间 `*`（如 `110101********1234`）
  - 手机/电话：`138****1234`
  - 姓名：保留首字，其余 `*`（`张三`→`张*`，`欧阳娜娜`→`欧**`）
  - 公司名：保留前 2 字 + `**`
  - 邮箱：`a***@xx.com`
  - 金额/数值：同长度随机数字（保留小数位）
- **只脱敏传给 AI 的内容**（结构 json 样例、prompt、日志摘要）；
  沙箱执行使用全量真实数据（数据全程本地，不出服务器）

## 10. 执行链路与错误处理

- 上传 → 临时目录（tempfile，复用智算模式）→ 加密检测（复用 /api/files/check-encrypted + 密码弹窗）
- 执行走**子进程**（`subprocess.Popen([python, -m, backend.assemble.assemble_worker, params])`，
  参考 compute_worker：Aspose/.NET 持 GIL 会冻结事件循环导致 SSE 卡死），进度事件 `@@EVT@@` 前缀回推
- `CodeSandbox.execute()` 执行，超时 360s
- 结果落盘：`tenants/{tenant_id}/assemble_results/{task_id}/`（原版+纯值版），
  下载复用 `/api/download-*` 的 FileResponse 模式
- AI 生成失败 → 带错误反馈自动重试 1 次；执行失败 → error 事件 + 日志 + 重试按钮
- 源文件无表头 → excel_parser 现有兜底 + 日志提示；模板只有表头无数据行 → 明确提示
- 加密文件 → 密码弹窗；规则文件被删/改 → 签名不命中自动重新生成

## 11. 权限与多租户

- 管理：`tools.assemble.manage`；执行页签可见：`tools.assemble`
- 规则可见性：全局规则所有租户可见；按租户规则仅该租户可见（API 按当前用户租户过滤）
- 代码存档、结果、任务历史均按 tenant_id 隔离

## 12. 边界与默认值

- 源数据行 < 3：样例取全部行
- 同名列兜底时模板列集合用 _COL_MAP 的真实列（剔除 Column_* 幻影列，参考既有教训）
- 结果文件命名：`{规则名}_{时间戳}_原版.xlsx` / `_纯值.xlsx`
- 前端静态资源改完必须 bump `?v=` 版本号（既有教训）
