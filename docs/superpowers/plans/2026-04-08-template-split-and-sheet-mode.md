# 报表模板：文件级拆分 + Sheet 模式 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为报表模板系统增加 sheet 模式（每组一个 sheet）和 split_by 文件级拆分（按列值拆分成多个文件）。

**Architecture:** 在现有 fill/block/zip 三模式基础上新增 sheet 模式，并增加 split_by 字段实现两层分组。改动从数据库模型层 → 核心生成层 → API 层 → 前端 UI 四层推进，每层独立可验证。

**Tech Stack:** Python/FastAPI, SQLAlchemy, Aspose.Cells (.NET via pythonnet), vanilla JS

---

### Task 1: 数据库模型 — 新增 split_by 字段

**Files:**
- Modify: `backend/database/models.py:95-98`
- Modify: `backend/database/init_db.py:119-127`

- [ ] **Step 1: 在 Template 模型中添加 split_by 字段**

在 `backend/database/models.py` 的 `Template` 类中，在 `name_field` 行（第98行）之后添加：

```python
split_by = Column(String(100), default="")                         # 文件级拆分字段（如：部门）
```

修改后该区域代码为：
```python
    report_mode = Column(String(20), default="fill")                     # fill / block / zip / sheet
    group_by = Column(String(100), default="")                           # block/zip/sheet 分组字段名
    skip_rows = Column(Integer, default=1)                               # block 块间空行数
    name_field = Column(String(100), default="")                         # zip 文件命名字段
    split_by = Column(String(100), default="")                           # 文件级拆分字段（如：部门）
    show_empty_period = Column(Boolean, default=True)                    # 多月合并时是否显示空月份
```

- [ ] **Step 2: 在迁移列表中添加 split_by**

在 `backend/database/init_db.py` 的 `migrations` 列表中，在 `("templates", "show_empty_period", ...)` 行之后添加：

```python
        ("templates", "split_by", "VARCHAR(100) DEFAULT ''"),
```

- [ ] **Step 3: Commit**

```bash
git add backend/database/models.py backend/database/init_db.py
git commit -m "feat: add split_by field to Template model and migration"
```

---

### Task 2: 核心生成层 — 新增 sheet 模式

**Files:**
- Modify: `backend/utils/aspose_helper.py:700-751` (入口函数)
- Modify: `backend/utils/aspose_helper.py:886` (在 block 模式之后插入 sheet 模式函数)

- [ ] **Step 1: 新增 `_sanitize_sheet_name` 辅助函数**

在 `backend/utils/aspose_helper.py` 的 `_extract_group_vars` 函数之后（约第1014行后），添加：

```python
def _sanitize_sheet_name(name: str, existing_names: set) -> str:
    """清洗 sheet 名称：截断至31字符，替换非法字符，处理重名。"""
    # 替换非法字符
    clean = re.sub(r'[\\/:*?\[\]]', '_', str(name).strip())
    # 截断至 31 字符（Excel 限制）
    if len(clean) > 31:
        clean = clean[:31]
    # 空名称兜底
    if not clean:
        clean = "Sheet"
    # 处理重名
    base = clean
    counter = 2
    while clean in existing_names:
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        clean = base[:max_base_len] + suffix
        counter += 1
    existing_names.add(clean)
    return clean
```

- [ ] **Step 2: 新增 `_generate_sheet` 函数**

在 `_generate_block` 函数之后（约第886行后，`# ── zip 模式` 注释之前），添加 sheet 模式的实现：

```python
# ── sheet 模式 ────────────────────────────────────────

def _generate_sheet(
    output_path: str, template_path: str, data: Dict,
    group_by: str = "",
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
) -> str:
    """按 group_by 分组，每组生成一个独立 sheet，sheet 名取自分组值。"""
    ds_name, full_df, vars_data = _extract_datasource(data)

    # 模糊匹配 group_by 列名
    if group_by and group_by not in full_df.columns:
        matched = _fuzzy_match_column(group_by, full_df.columns)
        if matched:
            logger.info(f"[sheet] group_by 模糊匹配: '{group_by}' -> '{matched}'")
            group_by = matched

    if not group_by or group_by not in full_df.columns:
        logger.warning(f"[sheet] group_by='{group_by}' 不在列 {list(full_df.columns)} 中，回退到 fill 模式")
        return _generate_fill(output_path, template_path, data, password, watermark_text)

    groups = full_df.groupby(group_by, sort=False)
    logger.info(f"[报表生成] sheet 模式: {len(groups)} 组, group_by={group_by}")

    # 用第一组填充到模板的第一个 sheet，后续组复制 sheet 后填充
    result_wb = None
    sheet_names = set()

    for group_idx, (group_key, group_df) in enumerate(groups):
        group_df = group_df.reset_index(drop=True)
        group_vars = _extract_group_vars(group_df, vars_data)
        group_data = {ds_name: group_df, **group_vars}

        # SmartMarker 填充该组
        filled_wb = _smartmarker_fill(template_path, group_data)

        sheet_name = _sanitize_sheet_name(str(group_key), sheet_names)

        if result_wb is None:
            # 第一组：直接用填好的 workbook，重命名第一个 sheet
            result_wb = filled_wb
            result_wb.Worksheets[0].Name = sheet_name
        else:
            # 后续组：从填好的 workbook 复制第一个 sheet 到结果 workbook
            result_wb.Worksheets.AddCopy(result_wb.Worksheets.Count - 1)
            new_ws = result_wb.Worksheets[result_wb.Worksheets.Count - 1]

            # 将填好的数据复制过来
            filled_ws = filled_wb.Worksheets[0]
            new_ws.Copy(filled_ws)
            new_ws.Name = sheet_name

        logger.info(f"[sheet] 组 {group_idx+1}/{len(groups)}: sheet='{sheet_name}', {len(group_df)} 行数据")

    return _finalize_workbook(result_wb, output_path, password, watermark_text)
```

- [ ] **Step 3: 在 `generate_from_template` 入口中添加 sheet 模式路由**

修改 `backend/utils/aspose_helper.py` 的 `generate_from_template` 函数（第701行起），在函数签名中添加 `split_by` 参数，并在路由中增加 sheet 分支：

函数签名改为：
```python
def generate_from_template(
    output_path: str,
    template_path: str,
    data: Dict,
    password: Optional[str] = None,
    watermark_text: Optional[str] = None,
    mode: str = "fill",
    group_by: str = "",
    skip_rows: int = 1,
    name_field: str = "",
    show_empty_period: bool = True,
    split_by: str = "",
) -> str:
```

函数体改为（替换原有的 if/elif/else 分支，即第733-751行）：
```python
    # 如果有 split_by，走文件级拆分包装器
    if split_by:
        return _generate_with_split(
            output_path, template_path, data,
            split_by=split_by, mode=mode,
            group_by=group_by, skip_rows=skip_rows,
            name_field=name_field,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )

    if mode == "block":
        return _generate_block(
            output_path, template_path, data,
            group_by=group_by, skip_rows=skip_rows,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )
    elif mode == "zip":
        return _generate_zip(
            output_path, template_path, data,
            group_by=group_by, name_field=name_field,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )
    elif mode == "sheet":
        return _generate_sheet(
            output_path, template_path, data,
            group_by=group_by,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )
    else:
        return _generate_fill(
            output_path, template_path, data,
            password=password, watermark_text=watermark_text,
        )
```

- [ ] **Step 4: Commit**

```bash
git add backend/utils/aspose_helper.py
git commit -m "feat: add sheet mode and split_by parameter to template generation"
```

---

### Task 3: 核心生成层 — 新增 split_by 文件级拆分

**Files:**
- Modify: `backend/utils/aspose_helper.py` (在 `_generate_sheet` 之后，`_extract_datasource` 之前插入)

- [ ] **Step 1: 新增 `_generate_with_split` 函数**

在 `_generate_sheet` 函数之后、`# ── 公共工具` 注释之前插入：

```python
# ── split_by 文件级拆分 ────────────────────────────────

def _generate_with_split(
    output_path: str, template_path: str, data: Dict,
    split_by: str = "", mode: str = "fill",
    group_by: str = "", skip_rows: int = 1,
    name_field: str = "",
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
) -> str:
    """按 split_by 字段拆分数据到多个文件，每个文件内按 mode 模式生成，打包为 zip。"""
    ds_name, full_df, vars_data = _extract_datasource(data)

    # 模糊匹配 split_by 列名
    if split_by not in full_df.columns:
        matched = _fuzzy_match_column(split_by, full_df.columns)
        if matched:
            logger.info(f"[split] split_by 模糊匹配: '{split_by}' -> '{matched}'")
            split_by = matched

    if split_by not in full_df.columns:
        logger.warning(f"[split] split_by='{split_by}' 不在列 {list(full_df.columns)} 中，忽略拆分，走普通模式")
        # 回退到无 split_by 的普通模式
        if mode == "sheet":
            return _generate_sheet(output_path, template_path, data, group_by=group_by,
                                   password=password, watermark_text=watermark_text,
                                   show_empty_period=show_empty_period)
        elif mode == "block":
            return _generate_block(output_path, template_path, data, group_by=group_by,
                                   skip_rows=skip_rows, password=password,
                                   watermark_text=watermark_text, show_empty_period=show_empty_period)
        elif mode == "zip":
            return _generate_zip(output_path, template_path, data, group_by=group_by,
                                 name_field=name_field, password=password,
                                 watermark_text=watermark_text, show_empty_period=show_empty_period)
        else:
            return _generate_fill(output_path, template_path, data, password=password,
                                  watermark_text=watermark_text)

    # split_by + zip 时，split_by 覆盖 zip 的 group_by 语义
    if mode == "zip":
        logger.info(f"[split] split_by + zip 模式，split_by 覆盖 group_by，等同于按 '{split_by}' 做 zip")
        return _generate_zip(output_path, template_path, data,
                             group_by=split_by, name_field=name_field,
                             password=password, watermark_text=watermark_text,
                             show_empty_period=show_empty_period)

    # 确保输出路径是 .zip
    if not output_path.endswith(".zip"):
        output_path = os.path.splitext(output_path)[0] + ".zip"

    split_groups = full_df.groupby(split_by, sort=False)
    logger.info(f"[报表生成] split 模式: {len(split_groups)} 个文件, split_by={split_by}, 内部模式={mode}")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for split_idx, (split_key, split_df) in enumerate(split_groups):
            split_df = split_df.reset_index(drop=True)
            # 构建该拆分组的 data dict
            split_vars = _extract_group_vars(split_df, vars_data)
            split_data = {ds_name: split_df, **split_vars}

            # 生成临时文件
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp_path = tmp.name

            try:
                if mode == "sheet":
                    _generate_sheet(tmp_path, template_path, split_data,
                                    group_by=group_by, password=password,
                                    watermark_text=watermark_text,
                                    show_empty_period=show_empty_period)
                elif mode == "block":
                    _generate_block(tmp_path, template_path, split_data,
                                    group_by=group_by, skip_rows=skip_rows,
                                    password=password, watermark_text=watermark_text,
                                    show_empty_period=show_empty_period)
                else:
                    # fill 模式
                    _generate_fill(tmp_path, template_path, split_data,
                                   password=password, watermark_text=watermark_text)

                # 文件命名：用 split_key 值
                file_label = re.sub(r'[\\/:*?"<>|]', '_', str(split_key).strip())
                if not file_label:
                    file_label = f"group_{split_idx + 1}"
                inner_name = f"{file_label}.xlsx"

                zf.write(tmp_path, inner_name)
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

            logger.info(f"[split] 文件 {split_idx+1}/{len(split_groups)}: {inner_name}, {len(split_df)} 行, 内部模式={mode}")

    logger.info(f"[报表生成] split 完成: {output_path}")
    return output_path
```

- [ ] **Step 2: Commit**

```bash
git add backend/utils/aspose_helper.py
git commit -m "feat: add split_by file-level splitting wrapper"
```

---

### Task 4: API 层 — 透传 split_by 字段

**Files:**
- Modify: `backend/admin/router.py:436-456` (_build_template_resp)
- Modify: `backend/admin/router.py:479-518` (create_template)
- Modify: `backend/admin/router.py:534-602` (update_template)
- Modify: `backend/admin/router.py:903-960` (generate_report)

- [ ] **Step 1: 在 `_build_template_resp` 中返回 split_by**

在 `backend/admin/router.py` 第449行 `"name_field"` 之后添加一行：

```python
        "split_by": getattr(t, "split_by", "") or "",
```

- [ ] **Step 2: 在 `create_template` 接口中接收 split_by 参数**

在 `backend/admin/router.py` 的 `create_template` 函数签名中，`show_empty_period` 参数之前添加：

```python
    split_by: str = Form(""),
```

在函数体中创建 `Template` 对象的地方（约第499行 `tpl = Template(...)` 内），添加：

```python
        split_by=split_by,
```

- [ ] **Step 3: 在 `update_template` 接口中接收并更新 split_by**

在 `backend/admin/router.py` 的 `update_template` 函数签名中，`show_empty_period` 参数之前添加：

```python
    split_by: Optional[str] = Form(None),
```

在函数体中更新字段的区域（约第597行 `show_empty_period` 处理之前），添加：

```python
    if 'split_by' in form:
        tpl.split_by = str(form.get('split_by', ''))
    elif split_by is not None:
        tpl.split_by = split_by
```

- [ ] **Step 4: 在 generate_report 中提取并传递 split_by**

在 `backend/admin/router.py` 第908行（`show_empty = getattr(...)` 之后），添加：

```python
    split_by_field = getattr(tpl, "split_by", "") or ""
```

修改第910-933行的校验逻辑，将 sheet 模式纳入需要 group_by 的校验：

```python
    # zip/block/sheet 模式前置校验：group_by 不能为空，且必须在数据列中
    if report_mode in ("zip", "block", "sheet"):
        if not group_by_field:
            raise HTTPException(
                status_code=400,
                detail=f"报表模式为 {report_mode}，但模版未配置分组字段(group_by)，请在模版设置中指定分组列名",
            )
        available_cols = list(dataset.columns)
        matched_col = None
        target = group_by_field.strip().lower()
        for col in available_cols:
            if str(col).strip().lower() == target:
                matched_col = col
                break
        if not matched_col:
            raise HTTPException(
                status_code=400,
                detail=f"分组字段 '{group_by_field}' 不在数据列中，可用列: {available_cols[:30]}",
            )
        if matched_col != group_by_field:
            logger.info(f"group_by 模糊匹配: '{group_by_field}' -> '{matched_col}'")
            group_by_field = matched_col
```

在第935-940行的输出扩展名判断中，增加 split_by 的处理：

```python
    # 有 split_by 或 zip 模式时输出 .zip，其余输出原始扩展名
    if report_mode == "zip" or split_by_field:
        output_ext = ".zip"
        output_name_final = os.path.splitext(output_name)[0] + output_ext
    else:
        output_name_final = output_name
```

在调用 `generate_from_template` 时（第950-960行）添加 `split_by` 参数：

```python
        actual_output_path = aspose_helper.generate_from_template(
            output_path=output_path,
            template_path=tpl.file_path,
            data=template_data,
            password=password,
            mode=report_mode,
            group_by=group_by_field,
            skip_rows=skip_rows_val,
            name_field=name_field_val,
            show_empty_period=show_empty,
            split_by=split_by_field,
        )
```

在日志行（第947行）更新为：

```python
    logger.info(f"报表模式: {report_mode}, group_by={group_by_field}, split_by={split_by_field}, skip_rows={skip_rows_val}")
```

- [ ] **Step 5: Commit**

```bash
git add backend/admin/router.py
git commit -m "feat: pass split_by through template CRUD and report generation API"
```

---

### Task 5: 前端 — 模板管理 UI 更新

**Files:**
- Modify: `frontend/static/js/admin.js:642-663` (renderTemplates)
- Modify: `frontend/static/js/admin.js:665-737` (showCreateTemplate)
- Modify: `frontend/static/js/admin.js:741-816` (showEditTemplate)
- Modify: `frontend/static/js/admin.js:825-833` (_toggleModeFields)

- [ ] **Step 1: 更新 `_toggleModeFields` 显示逻辑**

替换 `frontend/static/js/admin.js` 第825-833行的 `_toggleModeFields` 函数为：

```javascript
    _toggleModeFields() {
        const mode = document.getElementById('m-tpl-report-mode')?.value || 'fill';
        const fields = document.getElementById('m-tpl-mode-fields');
        const skipGroup = document.getElementById('m-tpl-skip-rows-group');
        const nameGroup = document.getElementById('m-tpl-name-field-group');
        if (fields) fields.style.display = (mode === 'block' || mode === 'zip' || mode === 'sheet') ? 'block' : 'none';
        if (skipGroup) skipGroup.style.display = mode === 'block' ? 'block' : 'none';
        if (nameGroup) nameGroup.style.display = mode === 'zip' ? 'block' : 'none';
    },
```

变更：`fields` 的显示条件增加了 `mode === 'sheet'`。

- [ ] **Step 2: 更新 `showCreateTemplate` 弹窗**

在 `frontend/static/js/admin.js` 的 `showCreateTemplate` 函数中做三处修改：

**2a.** 在报表模式 select（第695行附近）中增加 sheet 选项：

将：
```javascript
                        <option value="zip">zip — 分组打包（每组一文件）</option>
```
改为：
```javascript
                        <option value="zip">zip — 分组打包（每组一文件）</option>
                        <option value="sheet">sheet — 分组多Sheet（每组一个Sheet）</option>
```

**2b.** 在 `m-tpl-mode-fields` div 之前（即第700行 `</select>` 和 `</div>` 之后），添加 split_by 输入框：

```javascript
                <div class="form-group"><label>文件拆分字段</label>
                    <input id="m-tpl-split-by" placeholder="如：部门（留空则不拆分文件）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">按此列值将数据拆分到不同文件中，拆分后自动打包为 zip</small>
                </div>
```

**2c.** 在提交回调的 `FormData` 构建中（约第731行），添加：

```javascript
            fd.append('split_by', document.getElementById('m-tpl-split-by')?.value || '');
```

- [ ] **Step 3: 更新 `showEditTemplate` 弹窗**

在 `frontend/static/js/admin.js` 的 `showEditTemplate` 函数中做三处修改：

**3a.** 在报表模式 select（第774-779行）中增加 sheet 选项：

将：
```javascript
                        <option value="zip" ${t.report_mode==='zip'?'selected':''}>zip — 分组打包（每组一文件）</option>
```
改为：
```javascript
                        <option value="zip" ${t.report_mode==='zip'?'selected':''}>zip — 分组打包（每组一文件）</option>
                        <option value="sheet" ${t.report_mode==='sheet'?'selected':''}>sheet — 分组多Sheet（每组一个Sheet）</option>
```

**3b.** 在报表模式 `</select></div>` 之后、`m-tpl-mode-fields` div 之前，添加 split_by 输入框：

```javascript
                <div class="form-group"><label>文件拆分字段</label>
                    <input id="m-tpl-split-by" value="${t.split_by || ''}" placeholder="如：部门（留空则不拆分文件）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">按此列值将数据拆分到不同文件中，拆分后自动打包为 zip</small>
                </div>
```

**3c.** 更新 `m-tpl-mode-fields` div 的初始 display 条件（第781行）：

将：
```javascript
                <div id="m-tpl-mode-fields" style="display:${(t.report_mode==='block'||t.report_mode==='zip')?'block':'none'};">
```
改为：
```javascript
                <div id="m-tpl-mode-fields" style="display:${(t.report_mode==='block'||t.report_mode==='zip'||t.report_mode==='sheet')?'block':'none'};">
```

**3d.** 在提交回调的 `FormData` 构建中（约第811行），添加：

```javascript
            fd.append('split_by', document.getElementById('m-tpl-split-by')?.value || '');
```

- [ ] **Step 4: 更新 `renderTemplates` 列表显示**

修改 `frontend/static/js/admin.js` 第655行的模式标签渲染，添加 sheet 模式标签和 split_by 显示：

将：
```javascript
            <td>${t.report_mode === 'block' ? '<span class="tag" style="background:#fff3e0;color:#e65100">block</span>' : t.report_mode === 'zip' ? '<span class="tag" style="background:#e8eaf6;color:#283593">zip</span>' : 'fill'}${t.group_by ? ' <small>(' + t.group_by + ')</small>' : ''}</td>
```
改为：
```javascript
            <td>${t.report_mode === 'block' ? '<span class="tag" style="background:#fff3e0;color:#e65100">block</span>' : t.report_mode === 'zip' ? '<span class="tag" style="background:#e8eaf6;color:#283593">zip</span>' : t.report_mode === 'sheet' ? '<span class="tag" style="background:#e8f5e9;color:#2e7d32">sheet</span>' : 'fill'}${t.group_by ? ' <small>(' + t.group_by + ')</small>' : ''}${t.split_by ? ' <small style="color:#1565c0;">[拆分:' + t.split_by + ']</small>' : ''}</td>
```

- [ ] **Step 5: Commit**

```bash
git add frontend/static/js/admin.js
git commit -m "feat: add sheet mode option and split_by field to template management UI"
```

---

### Task 6: 手动验证

- [ ] **Step 1: 启动服务并验证数据库迁移**

```bash
python run.py --start
```

检查日志中是否有 `split_by` 列的迁移成功信息。

- [ ] **Step 2: 验证前端 UI**

1. 打开管理后台 → 模版管理
2. 点击「新建模版」，确认看到：
   - 报表模式下拉框有 4 个选项：fill / block / zip / sheet
   - 「文件拆分字段」输入框始终可见
   - 选择 sheet 时，分组字段、显示空月份可见，块间空行数隐藏
3. 点击已有模版的「编辑」，确认 split_by 字段能正确回显

- [ ] **Step 3: 验证 sheet 模式生成**

1. 创建一个模版，模式选 sheet，分组字段设为数据中有的列名（如"部门"）
2. 执行报表生成，确认生成的 xlsx 中每个分组值各占一个 sheet

- [ ] **Step 4: 验证 split_by 拆分生成**

1. 编辑模版，填写文件拆分字段（如"公司名"），模式选 fill 或 block
2. 执行报表生成，确认输出为 zip 文件，解压后每个公司名各一个 xlsx
