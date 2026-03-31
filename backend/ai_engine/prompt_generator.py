"""
提示词生成器 - 为AI生成训练和修正提示词（精简版）
"""

import json
import logging
import re
from typing import Dict, List, Any, Optional
from pathlib import Path
from .rule_extractor import RuleExtractor


class PromptGenerator:
    """提示词生成器"""

    # ============ 通用组件说明（只定义一次）============
    EXCEL_PARSER_INTERFACE = """## IntelligentExcelParser 接口
使用 `from excel_parser import IntelligentExcelParser` 解析Excel。

**方法签名**:
```python
parse_excel_file(file_path, max_data_rows=None, skip_rows=0, manual_headers=None, headers_only=False, active_sheet_only=False) -> List[SheetData]
```

**参数说明**:
- `file_path`: Excel文件路径
- `max_data_rows`: 每个区域最多读取的数据行数，None表示读取全部
- `skip_rows`: 从文件开头跳过的行数
- `manual_headers`: 手动指定的表头范围（通常从全局变量获取）
- `headers_only`: 是否只读取表头，不读取数据行（用于快速匹配）
- `active_sheet_only`: 是否只加载当前激活的Sheet，默认False

**返回值**: `List[SheetData]`，每个SheetData包含:
- `sheet_name`: str - Sheet名称
- `regions`: List[ExcelRegion] - 数据区域列表

**ExcelRegion结构**:
- `head_data`: Dict[str, str] - 表头名到列字母映射，如 {"姓名": "A", "工资": "B"}
- `data`: List[Dict[str, Any]] - 数据行，格式 {列字母: 值}
- `formula`: Dict[str, str] - 公式映射

**使用示例**:
```python
parser = IntelligentExcelParser()
results = parser.parse_excel_file(file_path, manual_headers=manual_headers)
for sheet_data in results:
    for region in sheet_data.regions:
        # 转换为DataFrame
        df = convert_region_to_dataframe(region)
```"""

    GLOBAL_VARS_DESC = """## 可用全局变量（直接使用，勿用os.environ）
- input_folder, output_folder: 路径字符串
- source_files: 源文件名列表
- manual_headers: 手动表头规则
- salary_year, salary_month, monthly_standard_hours: 薪资参数（可选）"""

    CORE_RULES = """## 核心规则
1. **路径拼接**: 必须用 os.path.join(input_folder/output_folder, filename)
2. **列访问**: 用 safe_get_column(df, "列名", 默认值)，禁止直接 df["列名"]
3. **DataFrame初始化**: 用 base_df = 源表.copy()，禁止空DataFrame后赋值
4. **apply用法**: df["列"].apply(lambda x: ...)，x是单值，非Series
5. **列名匹配**: 源文件列名可能与规则不一致，需建立语义映射
6. **ROUND操作**: 仅规则明确要求时才添加
7. **动态查找源数据**: 使用 find_source_sheet(source_sheets, target_columns=[...], sheet_name_hint="...") 查找源数据sheet，不要硬编码文件名。例如：
   ```python
   # 错误示例（硬编码文件名）：
   df_main = source_sheets['1薪资备份表_Sheet1']["df"]

   # 正确示例（动态查找）：
   key_salary = find_source_sheet(source_sheets, target_columns=["姓名", "部门", "基本工资"], sheet_name_hint="薪资")
   df_main = source_sheets[key_salary]["df"]
   ```
   这样可以避免计算时因文件名不同导致KeyError。"""

    # ============ 公式模式规则（集中定义，只维护一处）============
    FORMULA_RULES = {
        "char_encoding": {
            "compact": "【规则1】字符规范：必须英文半角()[]\"'，禁止中文全角",
            "detailed": "【规则1】字符规范\n- ✅ 必须：英文半角 ()[]\"'\n- ❌ 禁止：中文全角 （）【】\u201c\u201d\u2018\u2019\n- 最优先检查，每次输出前检查",
        },
        "cross_table_lookup": {
            "compact": "【规则2】跨表取数：L1同源列直接复制（不查列号），L2跨表列才用VLOOKUP/INDEX+MATCH/SUMPRODUCT",
            "detailed": (
                "【规则2】跨表取数（按层级处理）\n"
                "- **L1同源列**（与主键在同一源表）：用main_df.iloc[i].get('列名','')直接复制\n"
                "  ❌ 禁止对同源列使用VLOOKUP\n"
                "  ❌ 禁止对同源列定义col_xxx或调用get_vlookup_col_num()\n"
                "- **L2跨表列**（其他源表的数据）：必须使用Excel公式跨表取数：\n"
                "  - VLOOKUP：单条件精确匹配（默认首选），列号用get_vlookup_col_num()计算\n"
                "    格式：=IFERROR(VLOOKUP(查找值,'表'!$主键列:$末列,列号,FALSE),默认值)\n"
                "  - XLOOKUP：反向查找、自定义默认值\n"
                "    格式：=XLOOKUP(查找值,'表'!查找列,'表'!返回列,默认值,0)\n"
                "  - INDEX+MATCH：多条件匹配、左向查找\n"
                "    格式：=IFERROR(INDEX('表'!返回列,MATCH(查找值,'表'!查找列,0)),默认值)\n"
                "  - FILTER：一对多匹配+聚合\n"
                "    格式：=SUM(FILTER('表'!金额列,'表'!工号列=A2))\n"
                "  - SUMPRODUCT：多条件求和/计数\n"
                "    格式：=SUMPRODUCT(('表'!$A:$A=A2)*('表'!$B:$B=\"条件\")*('表'!$D:$D))\n"
                "- 参考规则文档中「列处理分层」章节判断每列属于L1还是L2\n"
                "- 所有跨表公式必须用IFERROR包裹"
            ),
        },
        "date_conversion": {
            "compact": "【规则3】日期必转换 - DATEVALUE()，AND()/OR()不短路需IFERROR包裹",
            "detailed": (
                "【规则3】日期必转换\n"
                "- 日期参与计算前用 `IFERROR(DATEVALUE(x),x)` 确保是数值\n"
                "- 检查清单：日期比较、日期相减、日期筛选\n"
                "- ⚠️ AND()/OR()不短路：空单元格调MONTH()/YEAR()会报#VALUE!，用IFERROR包裹"
            ),
        },
        "fstring_rules": {
            "compact": "【规则4】f-string规则：公式含双引号时外层用单引号，文本比较用TXT_xxx常量，空字符串一律用{EMPTY}",
            "detailed": (
                "【规则4】f-string引号规则\n"
                "- 外层双引号时，sheet名单引号直接写\n"
                "- 公式含双引号（如TEXT格式）时，外层用单引号：f'=TEXT(A{r},\"YYYY-MM\")'\n"
                "- Excel空字符串 → 一律用{EMPTY}代替，禁止直接写\"\"。包括：\n"
                "  - 等于空字符串：={EMPTY}  （不要写=\"\"）\n"
                "  - 不等于空字符串：<>{EMPTY}  （不要写<>\"\"）\n"
                "  - 逗号后空字符串：,{EMPTY})  （不要写,\"\")）\n"
                "- Excel文本比较 → 用excel_text()预定义TXT_xxx常量\n"
                "- 禁止在f-string内部调用函数（常量必须在循环外预定义）"
            ),
        },
        "no_import": {
            "compact": "【规则5】禁止在函数内部import模块",
            "detailed": "【规则5】禁止在函数内部导入模块\n- 所有需要的模块已在文件顶部导入\n- 禁止在fill_result_sheets或clean_source_data内部写import语句",
        },
        "employee_id": {
            "compact": "【规则6】工号一般是数字格式，不需要TEXT转换",
            "detailed": "【规则6】工号类型\n- 一般为数字格式，不需要TEXT转换\n- 仅当工号包含字母或特殊字符时才用TEXT转换",
        },
        "comment_format": {
            "compact": "【规则7】注释规范：每列一行 `# X列(N): 简要说明`",
            "detailed": "【规则7】注释规范\n- 每列仅一行注释：# X列(N): 列名 - 说明\n- 无分隔线，列之间不加空行",
        },
        "conditional_format": {
            "compact": "【规则8】条件格式：规则要求标红/高亮时用CellIsRule/FormulaRule实现",
            "detailed": (
                "【规则8】条件格式\n"
                "- 规则要求标红/高亮/颜色标记时，必须用CellIsRule/FormulaRule实现\n"
                "- 放在公式填充（for循环）之后\n"
                "- CellIsRule(operator=\"greaterThan\", formula=[\"20\"], fill=red_fill)\n"
                "- FormulaRule(formula=['条件公式'], fill=fill)"
            ),
        },
    }

    # 修正专用额外规则
    CORRECTION_EXTRA_RULES = {
        "type_safety": (
            "【额外规则】clean_source_data类型安全\n"
            "- groupby聚合时，必须用agg字典分别指定数值列sum、文本列first，"
            "禁止对整个DataFrame统一sum（含字母的工号列如'YN00002'会报错）\n"
            "- 编号补位（zfill）时，直接用str(x).strip().zfill(N)，禁止先转float/int再转str"
        ),
    }

    # 修正专用执行清单
    CORRECTION_CHECKLIST = (
        "## 执行前检查清单\n"
        "1. ✅ L1同源列（主表列）是否都用main_df.iloc[i].get()直接赋值？（禁止用VLOOKUP/INDEX+MATCH）\n"
        "2. ✅ 所有字符是否为英文半角？\n"
        "3. ✅ L2跨表列的VLOOKUP列号是否用get_vlookup_col_num()计算？\n"
        "4. ✅ 日期是否用DATEVALUE()转换？（仅L2跨表日期需要，L1主表日期直接赋值）\n"
        "5. ✅ f-string引号是否正确？\n"
        "6. ✅ 每行代码是否完整闭合？\n"
        "7. ✅ 条件格式是否用CellIsRule/FormulaRule实现？\n"
        "8. ✅ 是否所有括号都正确匹配？"
    )

    # 修正专用快速参考表
    CORRECTION_QUICK_REF = (
        "## 公式速查表\n"
        "| 场景 | 公式模板 |\n"
        "|------|----------|\n"
        "| L1主表列（直接复制） | ws.cell(row=r, column=N, value=main_df.iloc[i].get('列名','')) |\n"
        "| L2单条件查找 | =IFERROR(VLOOKUP(A{r},'表'!$A:$Z,列号,FALSE),0) |\n"
        "| L2左向查找 | =IFERROR(INDEX('表'!$返回列,MATCH(A{r},'表'!$查找列,0)),0) |\n"
        "| 多条件求和 | =SUMPRODUCT(('表'!$A$2:$A$N=A{r})*('表'!$D$2:$D$N)) |\n"
        "| 多条件计数 | =SUMPRODUCT(('表'!$A$2:$A$N=A{r})*1) |\n"
        "| 文本日期转换 | =IFERROR(DATEVALUE(x),x) |\n"
        "| 历史累计 | =IFERROR(SUMIFS(历史数据!$E:$E,历史数据!$A:$A,\"<\"&月份,历史数据!$B:$B,A{r}),0) |\n"
        "| 条件判断 | =IF(条件,真值,假值) |\n"
        "| 文本比较 | 用TXT_xxx常量，如=IF(A{r}={TXT_YES},1,0) |"
    )

    # 修正专用禁止操作
    CORRECTION_FORBIDDEN = (
        "## 禁止操作\n"
        "1. 禁止简化任何计算列为0或空字符串\n"
        "2. 禁止删除或注释掉已有的列代码\n"
        "3. 禁止修改未在差异列表中的列\n"
        "4. 禁止在结果报表内使用SUMIF等汇总函数替代源数据的逐行公式"
    )

    # 黄金样例（不含f-string转义，直接使用{r}）
    GOLDEN_EXAMPLE = r"""# ==================== 黄金样例（必须严格参照此模式编写）====================
以下是一个完整的fill_result_sheets代码示例，展示了8种常见列类型的正确写法。
你的代码必须完全遵循此模式，包括变量定义位置、缩进层级、注释格式、f-string写法。

```python
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    # === 1. 主表选择（根据规则中「列处理分层」指定的主表定位）===
    # ⚠️ 禁止用 max(len(df)) 选主表！行数最多的表不一定是主表
    # 优先级：规则中「列处理分层→主表」指定的表 > find_source_sheet按主键列+特征列定位
    main_key = find_source_sheet(source_sheets, target_columns=["工号", "姓名"], sheet_name_hint="工资")
    main_df = source_sheets[main_key]['df']
    n_rows = len(main_df)

    # === 2. 动态查找源表（禁止硬编码key，必须用find_source_sheet）===
    key_attend = find_source_sheet(source_sheets, target_columns=["出勤天数", "加班时数"], sheet_name_hint="考勤")
    key_roster = find_source_sheet(source_sheets, target_columns=["用工类型", "入职日期"], sheet_name_hint="花名册")
    sn_main = source_sheets[main_key]['ws'].title
    sn_attend = source_sheets[key_attend]['ws'].title
    sn_roster = source_sheets[key_roster]['ws'].title

    # === 3. 预计算VLOOKUP列号（在for循环外，用get_vlookup_col_num）===
    col_hire_date = get_vlookup_col_num("F", "A")      # 入职日期在F列，查找范围从A列开始
    col_attend_days = get_vlookup_col_num("E", "D")    # 出勤天数在E列，查找范围从D列(工号)开始

    # === 4. 预定义文本常量（Excel文本比较用，避免f-string引号冲突）===
    TXT_FULLTIME = excel_text('全职')

    # === 5. 创建结果sheet并写表头 ===
    ws = wb.create_sheet("结果")
    headers = ["工号", "姓名", "部门", "入职日期", "出勤天数",
               "应发工资", "岗位津贴", "累计工资", "加班费", "用工类型"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    # === 6. 逐行填充（每列平级排列，缩进统一8空格）===
    for i in range(n_rows):
        r = i + 2

        # A列(1): 工号 - [L0主键] 主表直接复制
        write_cell(ws, r, 1, main_df.iloc[i].get('工号', ''))

        # B列(2): 姓名 - [L1同源] 主表直接复制
        write_cell(ws, r, 2, main_df.iloc[i].get('姓名', ''))

        # C列(3): 部门 - [L1同源] 主表直接复制（主表内已有的列全部直接复制，不用VLOOKUP）
        write_cell(ws, r, 3, main_df.iloc[i].get('部门', ''))

        # ⚠️ 以上A/B/C列以及主表中的其他所有列（包括日期列、金额列）
        #    都用 write_cell(ws, r, 列号, main_df.iloc[i].get('源列名', '')) 写入。
        #    write_cell 会自动处理空值和日期格式（datetime自动设置yyyy/mm/dd）。
        #    只有下面来自非主表（花名册、考勤等）的列才用VLOOKUP/INDEX+MATCH。

        # D列(4): 入职日期 - [L2跨表] VLOOKUP（花名册，非主表）+ 日期格式转换
        ws.cell(row=r, column=4).value = f"=IFERROR(IFERROR(DATEVALUE(VLOOKUP(A{r},'{sn_roster}'!$A:$J,{col_hire_date},FALSE)),VLOOKUP(A{r},'{sn_roster}'!$A:$J,{col_hire_date},FALSE)),{EMPTY})"
        ws.cell(row=r, column=4).number_format = 'yyyy/mm/dd'

        # E列(5): 出勤天数 - [L2跨表] VLOOKUP（考勤表，非主表，主键非A列，范围起始为D列）
        ws.cell(row=r, column=5).value = f"=IFERROR(VLOOKUP(A{r},'{sn_attend}'!$D:$M,{col_attend_days},FALSE),0)"

        # F列(6): 应发工资 - [L3计算] 公式计算（引用参数sheet + 本表E列）
        ws.cell(row=r, column=6).value = f"=IFERROR(E{r}*参数!$B$4,0)"

        # G列(7): 岗位津贴 - [L3计算] 文本条件判断（用TXT_xxx常量 + 引用中间项J列）
        ws.cell(row=r, column=7).value = f"=IF(J{r}={TXT_FULLTIME},1000,500)"

        # H列(8): 累计工资 - [L3计算] SUMIFS历史累计（引用历史数据sheet）
        ws.cell(row=r, column=8).value = f"=IFERROR(SUMIFS(历史数据!$F:$F,历史数据!$A:$A,\"<\"&参数!$B$3,历史数据!$B:$B,A{r})+F{r},F{r})"

        # I列(9): 加班费 - [L3计算] SUMPRODUCT多条件加权计算（匹配工号，对加班时数×时薪求和）
        ws.cell(row=r, column=9).value = f"=IFERROR(SUMPRODUCT(('{sn_attend}'!$A$2:$A$1000=A{r})*('{sn_attend}'!$H$2:$H$1000)*('{sn_attend}'!$I$2:$I$1000)),0)"

        # J列(10): 用工类型（中间项，淡蓝色背景#DCE6F1）- [L2跨表] INDEX+MATCH左查找（花名册D列查工号，返回左侧B列）
        ws.cell(row=r, column=10).value = f"=IFERROR(INDEX('{sn_roster}'!$B:$B,MATCH(A{r},'{sn_roster}'!$D:$D,0)),{EMPTY})"

    # === 7. 中间项标记淡蓝色背景 ===
    if n_rows > 0:
        light_blue = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        for row in range(2, n_rows + 2):
            ws.cell(row=row, column=10).fill = light_blue

    # === 8. 条件格式（如规则要求标红/高亮等）===
    if n_rows > 0:
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        ws.conditional_formatting.add(f"E2:E{n_rows+1}", CellIsRule(operator="greaterThan", formula=["20"], fill=red_fill))

    # 共处理了 10 列（预期 10 列）
```

### 从黄金样例中提取的关键模式
1. **动态查找源表**：必须用find_source_sheet()查找，禁止硬编码source_sheets的key
2. **变量定义位置**：sn_xxx、col_xxx、TXT_xxx 全部在for循环外定义
3. **注释格式**：每列仅一行 `# X列(N): 说明`，无分隔线
4. **列平级排列**：每列代码缩进统一8空格，互不嵌套
5. **f-string规则**：
   - 外层双引号时，sheet名单引号直接写（参照D、E列）
   - 公式含双引号（如TEXT格式）时，外层用单引号：`f'=TEXT(A{r},"YYYY-MM")'`
   - Excel空字符串 → 一律用{EMPTY}代替（包括 ={EMPTY}、<>{EMPTY}、,{EMPTY})）
   - Excel文本比较 → 用excel_text('xxx')预定义TXT_xxx常量（参照G列）
   - 禁止在f-string内部调用函数（常量必须在循环外预定义）
6. **VLOOKUP列号**：用get_vlookup_col_num()计算，禁止硬编码数字
7. **INDEX+MATCH左查找**：当目标列在主键列左侧时使用（参照J列），VLOOKUP无法实现左查找
8. **SUMPRODUCT多条件计算**：多条件用乘法连接 `(条件1)*(条件2)*值列`，注意用有界范围$A$2:$A$1000而非整列（参照I列）
9. **主表数据必须用write_cell写入（L1同源）**：主表（由规则「列处理分层→主表」指定，用find_source_sheet定位，禁止用max(len(df))——行数最多的不一定是主表）中的所有列——无论是文本、日期还是金额——全部用write_cell(ws, r, 列号, main_df.iloc[i].get('源列名',''))写入。write_cell会自动处理空值和日期格式（datetime类型自动设置yyyy/mm/dd）。只有来自其他源表的列才用VLOOKUP/INDEX+MATCH（参照A、B、C列 vs D、E列的区别）
   - ⚠️ 判断标准：参考规则文档中「列处理分层」的L1（同源直接复制）和L2（跨表查找）
   - ⚠️ L1同源列不需要定义col_xxx变量，不需要调用get_vlookup_col_num()
   - ⚠️ 即使源列名与目标列名不同（如源"入职日期（卡中心）"→目标"聘用日期"），只要在主表中就直接复制
10. **中间项处理**：追加在最后一列之后，用PatternFill填充淡蓝色背景#DCE6F1（参照J列）
11. **条件格式**：规则要求标红/高亮时，在for循环结束后用CellIsRule/FormulaRule实现，不能只留注释
12. **日期格式转换**：VLOOKUP取回的日期可能是文本或数字，用 `IFERROR(DATEVALUE(x),x)` 转为真实日期值，再用 `number_format = 'yyyy/mm/dd'` 格式化显示（参照D列）。注意：AND()/OR()不短路，空单元格使用MONTH()/YEAR()时需IFERROR包裹"""

    # 补充规则（历史数据、工号、日期、完整性）—— __TOTAL_COLUMNS__ 在使用时替换
    SUPPLEMENTARY_RULES = (
        "# ==================== 补充规则 ====================\n\n"
        "## 主表列直接赋值（L1同源列，最高优先级）\n"
        "- 主表由规则「列处理分层→主表」指定，用find_source_sheet定位（禁止max(len(df))，行数最多的不一定是主表）\n"
        "- 如果规则未指定主表，则根据主键列名和特征列用find_source_sheet推断\n"
        "- 主表的所有列，用 main_df.iloc[i].get('源列名', '') 直接赋值\n"
        "- 包括日期列、金额列、文本列——只要在主表中，一律直接赋值\n"
        "- 主表日期列：赋值后加 ws.cell(row=r, column=N).number_format = 'yyyy/mm/dd'\n"
        "- ❌ 禁止对主表列使用VLOOKUP/INDEX+MATCH/XLOOKUP\n"
        "- ❌ 禁止对主表列定义col_xxx或调用get_vlookup_col_num()\n"
        "- 源列名与目标列名不同时，get()中写源列名\n\n"
        "## 历史数据引用（仅规则涉及'累计''历史'时使用）\n"
        "- 系统自动创建\"历史数据\"sheet，第1列为薪资月份，其余列与结果sheet列名一致\n"
        "- 累计示例：`=SUMIFS(历史数据!$E:$E, 历史数据!$A:$A, \"<\"&salary_month, 历史数据!$B:$B, B{r})`\n"
        "- 需用IFERROR包裹（第1个月历史数据可能为空）\n"
        "- 规则中未涉及历史数据时不要引用此sheet\n\n"
        "## 工号类型\n"
        "- 一般为数字格式，不需要TEXT转换\n"
        "- 仅当工号包含字母或特殊字符时才用TEXT转换\n\n"
        "## 日期格式转换\n"
        "- **L1主表日期列**：直接用main_df.iloc[i].get()赋值，然后加number_format='yyyy/mm/dd'即可，不需要DATEVALUE\n"
        "- **L2跨表日期列**：VLOOKUP/INDEX取回的日期往往是文本或数字序列号，必须转为真实日期值：\n"
        "  - 公式模式：`=IFERROR(DATEVALUE(VLOOKUP结果), VLOOKUP结果)` — 保留真实日期值（非文本）\n"
        "  - DATEVALUE将文本日期转为Excel日期数字，若已是数字则IFERROR兜底\n"
        "  - 格式化显示：公式赋值后紧跟 `ws.cell(row=r, column=N).number_format = 'yyyy/mm/dd'`\n"
        "- 日期参与减法运算前，也需要先用 `IFERROR(DATEVALUE(x),x)` 确保是数值\n"
        "- 禁止用TEXT()将日期转为文本，否则日期无法参与运算\n"
        "- ⚠️ AND()/OR()不短路：空单元格调MONTH()/YEAR()仍报#VALUE!，需用IFERROR包裹：`IFERROR(MONTH(A2),0)`\n\n"
        "## 完整性要求\n"
        "- 必须为全部 __TOTAL_COLUMNS__ 列生成处理逻辑\n"
        "- 在代码最后添加注释：# 共处理了 X 列（预期 __TOTAL_COLUMNS__ 列）\n"
        "- 每行代码必须完整闭合，不允许跨行写赋值语句\n"
        "- 禁止\"暂时跳过\"、\"简化为0\"、提前结束\n"
        "- 禁止在函数内部import已导入的模块\n"
        "- 全部使用英文半角字符 ()[]\"'，禁止中文全角"
    )

    # 函数签名模板 —— __TOTAL_COLUMNS__ 在使用时替换
    FUNCTION_SIGNATURES = r"""# ==================== 函数签名 ====================

## 1. 数据清洗函数（如果有清洗规则）
def clean_source_data(source_data):
    \"\"\"应用数据清洗规则：过滤无效数据 + 按主键汇总多行为单行

    代码结构要求（必须严格遵守）：
    - 必须用 for key, val in source_data.items(): 遍历所有表
    - 在循环内用 if/elif 判断不同的 key 做对应清洗
    - 不需要清洗的表也必须原样拷贝到 cleaned 中
    - return cleaned 必须在函数最外层（与for同级），禁止嵌套在if内部

    重要汇总规则（单人多行→单人单行）：
    - groupby的key只能是清洗规则中指定的汇总主键（如工号），除主键外的其他字段都不能作为groupby key
    - 数值型列（包括用文本表示的数值，但不包括日期）用sum聚合
    - 文本列、日期列用first取值
    - **禁止使用pivot_table**，不允许改变原表的列结构（表头必须和清洗前完全一致）
    - 汇总后必须更新columns：cleaned[key]["columns"] = list(df.columns)
    - ⚠️ groupby聚合时，必须用agg字典分别指定数值列sum、文本列first，禁止对整个DataFrame统一sum（含字母的工号列如'YN00002'会报错）
    - ⚠️ 编号补位（zfill）时，禁止先转float/int再转str，直接用str(x).strip().zfill(N)
    \"\"\"
    cleaned = {}
    for key, val in source_data.items():
        df = val["df"].copy()
        columns = val["columns"]
        # 根据key判断做对应清洗...
        cleaned[key] = {"df": df, "columns": list(df.columns)}
    return cleaned  # 必须在for循环之后、函数最外层返回

## 2. 结果填充函数
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    \"\"\"创建结果sheet并填充数据和公式\"\"\"
    pass

请严格参照黄金样例的代码模式，生成完整代码：
1. 如果有数据清洗规则，先生成clean_source_data函数
2. 然后生成fill_result_sheets函数，覆盖全部 __TOTAL_COLUMNS__ 列"""

    # 验证检查清单（multi-step step4专用）
    VERIFICATION_CHECKLIST = """## 检查清单
### 1. 语法检查
- [ ] 所有括号是否正确闭合（圆括号、方括号、花括号）
- [ ] f-string引号是否正确（统一 f"..." + \\' + \\"）
- [ ] f-string中是否有裸""（必须用{EMPTY}代替，包括<>""、=""、,"" 等所有场景）
- [ ] 每行代码是否完整，没有截断，没有跨行赋值
- [ ] 缩进是否一致（for循环内8空格，列之间平级，无级联嵌套）

### 2. VLOOKUP验证
- [ ] 每个VLOOKUP取的字段是否正确（字段名和源表对应）
- [ ] range_start是否等于该表主键所在列（不是固定$A）
- [ ] col_index = 目标列绝对位置 - 主键列绝对位置 + 1，计算是否正确
- [ ] range_end是否覆盖到目标列

### 3. 列完整性
- [ ] 是否覆盖了全部 __TOTAL_COLUMNS__ 列
- [ ] 基础列是否直接赋值，计算列是否有公式

### 4. 注释规范
- [ ] 每列注释是否只有一行：`# 列名(列号): 简要说明`
- [ ] 是否有多余的分隔线（# ------、# ===等）→ 必须删除

### 5. 其他
- [ ] sheet名是否用sn_变量，无硬编码
- [ ] 无中文全角字符（括号、引号）
- [ ] 无函数内import"""

    def __init__(self):
        self.templates = self._load_templates()
        self.max_structure_length = 20000000
        self.logger = logging.getLogger(__name__)
        self.rule_extractor = RuleExtractor()

    # ============ 集中规则辅助方法 ============

    def _build_formula_rules(self, detail: str = "compact", include_extra: bool = False) -> str:
        """组装公式模式规则段落

        Args:
            detail: "compact"（1行摘要）或 "detailed"（完整含示例）
            include_extra: 是否包含修正专用额外规则
        """
        lines = ["# ==================== 核心规则（违反=失败）===================="]
        for rule in self.FORMULA_RULES.values():
            lines.append(rule[detail])
        if include_extra:
            for rule in self.CORRECTION_EXTRA_RULES.values():
                lines.append(rule)
        return "\n\n".join(lines)

    def _build_golden_and_rules(self, total_columns: int, escape_braces: bool = False) -> str:
        """构建黄金样例 + 补充规则 + 函数签名

        Args:
            total_columns: 总列数，替换 __TOTAL_COLUMNS__
            escape_braces: 是否将{/}转义为{{/}}（用于f-string上下文）
        """
        parts = [
            self.GOLDEN_EXAMPLE,
            self.SUPPLEMENTARY_RULES.replace("__TOTAL_COLUMNS__", str(total_columns)),
            self.FUNCTION_SIGNATURES.replace("__TOTAL_COLUMNS__", str(total_columns)),
        ]
        result = "\n\n".join(parts)
        if escape_braces:
            # 将 {xxx} 转为 {{xxx}}，但保留已转义的 {{xxx}}
            # 先保护已有的 {{ }}
            result = result.replace("{{", "\x00\x00").replace("}}", "\x01\x01")
            result = result.replace("{", "{{").replace("}", "}}")
            result = result.replace("\x00\x00", "{{").replace("\x01\x01", "}}")
        return result

    def _extract_cleaning_and_format_rules(self, rules_content: str) -> dict:
        """统一提取数据清洗/警告/条件格式/精度规则（消除4处重复逻辑）

        Returns:
            {"data_cleaning": str, "warning": str, "conditional_format": str, "precision": str}
        """
        extracted = self.rule_extractor.extract_rules(rules_content)
        result = {"data_cleaning": "", "warning": "", "conditional_format": "", "precision": ""}

        if extracted["data_cleaning_rules"]:
            text = "\n## 数据清洗规则（在clean_source_data中应用）\n"
            text += "⚠️ 在将源数据写入Excel之前，必须先应用以下清洗规则过滤数据：\n\n"
            for i, rule in enumerate(extracted["data_cleaning_rules"], 1):
                if rule['original_text'] and not rule['original_text'].startswith('--'):
                    text += f"{i}. {rule['original_text']}\n"
            text += (
                "\n**实现方式**：生成clean_source_data函数，对每个DataFrame应用清洗逻辑后返回清洗后的数据。\n"
                "**汇总规则**：如果某个表存在单人多行数据需要汇总，groupby的key只能是清洗规则中指定的汇总主键，"
                "除主键外的其他字段都不能作为groupby key。数值型列（包括用文本表示的数值，但不包括日期）用sum聚合，"
                "文本列和日期列用first取值。禁止使用pivot_table，不允许改变原表的列结构（表头必须和清洗前完全一致）。\n"
            )
            result["data_cleaning"] = text

        if extracted.get("warning_rules"):
            text = "\n## 警告信息规则\n"
            for i, rule in enumerate(extracted["warning_rules"], 1):
                text += f"{i}. {rule['original_text']}\n"
            result["warning"] = text

        if extracted.get("conditional_format_rules"):
            text = "\n## 条件格式规则\n"
            text += "在填充公式后，需要对以下情况应用条件格式（使用CellIsRule/FormulaRule）：\n\n"
            for i, rule in enumerate(extracted["conditional_format_rules"], 1):
                text += f"{i}. {rule['original_text']}\n"
            result["conditional_format"] = text

        if extracted.get("precision_rules"):
            text = "\n## 数值精度规则\n"
            text += "在生成公式时，必须按以下精度要求使用ROUND函数：\n\n"
            for i, rule in enumerate(extracted["precision_rules"], 1):
                text += f"{i}. {rule['original_text']}\n"
            result["precision"] = text

        return result

    def _build_expected_sheets_info(self, expected_structure: dict, rules_content: str) -> tuple:
        """构建预期输出Sheet信息和总列数（消除3处重复逻辑）

        Returns:
            (expected_sheets_info: str, total_columns: int)
        """
        total_columns = 0
        expected_sheets_info = ""

        if isinstance(expected_structure, dict) and "sheets" in expected_structure:
            sheets = expected_structure.get("sheets", {})
            if sheets:
                expected_sheets_info = "\n## 预期输出Sheet列表\n"
                for sheet_name, sheet_info in sheets.items():
                    headers = sheet_info.get("headers", {})
                    col_count = len(headers)
                    total_columns += col_count
                    expected_sheets_info += f"- **{sheet_name}** ({col_count}列): {list(headers.keys())}\n"

        # 从规则中提取中间计算项
        intermediate_items = self._extract_intermediate_items_from_rules(rules_content)
        if intermediate_items:
            total_columns += len(intermediate_items)
            intermediate_names = [item.split(':')[1] if ':' in item else item for item in intermediate_items]
            expected_sheets_info += f"\n## 中间计算项（追加在最后一列之后，淡蓝色背景 #DCE6F1）\n"
            expected_sheets_info += f"- 共 {len(intermediate_items)} 个中间项: {intermediate_names}\n"
            for item in intermediate_items:
                expected_sheets_info += f"  - {item}\n"
            expected_sheets_info += f"\n⚠️ 中间项必须作为新增列追加在结果sheet最后，使用淡蓝色背景(#DCE6F1)标识\n"

        return expected_sheets_info, total_columns

    def _extract_intermediate_items_from_rules(self, rules_content: str) -> List[str]:
        """从规则文本中提取中间计算项的列名

        匹配规则中定义的中间项，如:
        ### DS列: 当月标准工时（不含法定）（中间项，淡蓝色背景 #DCE6F1）
        ### DU列: 加班工时（中间项，淡蓝色背景 #DCE6F1）
        也匹配:
        - 类型: 中间计算项
        """
        intermediate_items = []
        if not rules_content:
            return intermediate_items

        seen = set()
        lines = rules_content.split('\n')

        for idx, line in enumerate(lines):
            stripped = line.strip()
            # 模式1: ### XX列: 名称（中间项...）
            match = re.match(r'^#{1,4}\s+([A-Z]{1,3})列[:：]\s*(.+?)(?:（中间项|（中间计算项)', stripped)
            if match:
                col_letter = match.group(1)
                col_name = match.group(2).strip()
                key = f"{col_letter}列:{col_name}"
                if key not in seen:
                    intermediate_items.append(key)
                    seen.add(key)
                continue
            # 模式2: ### XX列: 名称  同行或后续几行包含 "中间" 关键词
            match = re.match(r'^#{1,4}\s+([A-Z]{1,3})列[:：]\s*(.+)', stripped)
            if match:
                col_letter = match.group(1)
                col_name = match.group(2).strip()
                # 检查同行或后续3行是否有"中间项"/"中间计算项"标注
                nearby = stripped
                for j in range(1, min(4, len(lines) - idx)):
                    nearby += " " + lines[idx + j].strip()
                if '中间项' in nearby or '中间计算项' in nearby:
                    col_name = re.sub(r'[（(]中间项.*?[）)]', '', col_name).strip()
                    col_name = re.sub(r'[（(]淡蓝色.*?[）)]', '', col_name).strip()
                    key = f"{col_letter}列:{col_name}"
                    if key not in seen:
                        intermediate_items.append(key)
                        seen.add(key)

        return intermediate_items

    def _parse_layer_info_from_rules(self, rules_content: str) -> dict:
        """从规则文本中解析「列处理分层」信息

        解析规则文档中的 ## 列处理分层 章节，提取主键、各层级列名。
        如果规则中没有该章节（旧规则），返回空dict，调用方回退到默认行为。

        Returns:
            {
                "primary_key": "工号",
                "primary_source": "考勤表",
                "L1": ["姓名", "部门", ...],
                "L2": ["社保个人", ...],
                "L3": ["应发合计", ...],
                "L4": ["实发工资", ...],
                "column_layer": {"工号": "L0", "姓名": "L1", ...}
            }
            或空dict（无分层信息时）
        """
        if not rules_content or "## 列处理分层" not in rules_content:
            return {}

        result = {
            "primary_key": "",
            "primary_source": "",
            "main_table": "",
            "L1": [], "L2": [], "L3": [], "L4": [],
            "column_layer": {},
        }

        # 截取 ## 列处理分层 到下一个 ## 之间的内容
        start = rules_content.index("## 列处理分层")
        rest = rules_content[start + len("## 列处理分层"):]
        # 找下一个 ## 标题
        next_section = re.search(r'\n## [^#]', rest)
        section_text = rest[:next_section.start()] if next_section else rest

        # 解析主键
        pk_match = re.search(r'###\s*主键[:：]\s*(.+)', section_text)
        if pk_match:
            result["primary_key"] = pk_match.group(1).strip()
            result["column_layer"][result["primary_key"]] = "L0"

        # 解析主键来源表
        src_match = re.search(r'###\s*主键来源表[:：]\s*(.+)', section_text)
        if src_match:
            result["primary_source"] = src_match.group(1).strip()

        # 解析主表（如果规则指定了，优先使用；否则回退到主键来源表）
        main_match = re.search(r'###\s*主表[:：]\s*(.+?)(?:\n|（|$)', section_text)
        if main_match:
            result["main_table"] = main_match.group(1).strip()
        elif result["primary_source"]:
            result["main_table"] = result["primary_source"]

        # 解析各层级
        for layer in ["L1", "L2", "L3", "L4"]:
            # 匹配 ### L1-xxx 或 ### L2-xxx 等标题后的列表
            pattern = rf'###\s*{layer}[-\s].*?\n((?:[-\s]*.*\n)*?)(?=###|\Z)'
            layer_match = re.search(pattern, section_text)
            if layer_match:
                block = layer_match.group(1)
                for line in block.split('\n'):
                    line = line.strip()
                    if line.startswith('-'):
                        # 提取列名（取括号前的部分）
                        col_text = line.lstrip('- ').strip()
                        # "姓名（考勤表.姓名）" → "姓名"
                        col_name = re.split(r'[（(]', col_text)[0].strip()
                        if col_name:
                            result[layer].append(col_name)
                            result["column_layer"][col_name] = layer

        return result

    def _load_templates(self) -> Dict[str, str]:
        """加载提示词模板"""
        return {
            "training": """你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景。

## 行业背景
人力资源薪资计算项目，涉及薪资、考勤、奖金、社保、税务数据处理。

{excel_parser_interface}

## 主键选择原则
-**不同场景下主键选择不同，不同表之间进行关联主键的选取也不同，所以在开始之前，要分析好针对不同的表关联数据，要使用那个数据做主键**
- HR场景: 优先用雇员工号，身份证号，员工编号等唯一标识字段
- 一般场景: 根据数据特点选择唯一标识字段
- 处理主键缺失或重复情况，验证唯一性

## 输入文件结构
{source_structure}

## 预期输出结构
{expected_structure}

## 数据处理规则
{rules_content}

{data_cleaning_rules}

{warning_rules}

## 手动表头规则
{manual_headers}

## 列名匹配说明
源文件列名可能与规则描述不一致，需：
1. 建立语义映射（如"员工编码"="工号"="员工编号"）
2. 访问前验证列存在，不存在则用0填充并记录警告

{global_vars}

{core_rules}

## 入口函数要求
定义无参 `main()` 函数作为入口，内部使用全局变量。
禁止 `if __name__ == "__main__":` 块。

## 代码要求
- 使用IntelligentExcelParser读取Excel
- 详细错误处理，缺失列设为0并记录警告
- 验证员工编号唯一性
- 输出格式必须与预期结构一致
- 禁止简化任何计算逻辑
- **数据清洗**: 在复制基础数据时，必须应用数据清洗规则过滤不符合条件的数据
- **警告收集**: 创建warnings列表收集所有警告信息，在main()函数最后返回 {"success": True, "warnings": warnings}

请生成完整Python代码。""",

            "correction": """你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景，需修正以下代码：

{excel_parser_interface}

## 主键处理
工号标准化：转换为8位字符串，不足前面补0。
用标准化工号+姓名作为联合匹配键。

## 原始代码
{original_code}

## 问题描述
{error_description}

## 差异分析
{comparison_result}

## 输入文件结构
{source_structure}

## 预期输出结构
{expected_structure}

## 数据处理规则
{rules_content}

{data_cleaning_rules}

{warning_rules}

## 手动表头规则
{manual_headers}

{global_vars}

{core_rules}

## 修正要求
1. 保持整体结构，使用IntelligentExcelParser
2. 缺失列设为0，记录警告
3. 实现列名映射处理不一致问题
4. 确保输出与预期完全一致
5. **严禁简化处理**：禁止用 `= 0 # 简化` 跳过任何计算
6. 主键字段是不是和其他表内的类型不一致，导致需要转换
7. **数据清洗**: 在复制基础数据时，必须应用数据清洗规则过滤不符合条件的数据
8. **警告收集**: 创建warnings列表收集所有警告信息，在main()函数最后返回 {"success": True, "warnings": warnings}

请提供修正后的完整代码。""",

            "validation": """请分析以下Python代码的潜在问题：

## 代码内容
{code_content}

## 检查要点
1. 语法和逻辑错误（薪资计算、税务处理）
2. 列名匹配问题（是否有映射机制、是否验证存在）
3. 错误处理（缺失列、数据完整性）
4. 性能和安全风险

请提供详细检查报告。"""
        }

    def _compress_structure(self, structure: Dict[str, Any], max_length: int = 30000) -> str:
        """压缩数据结构，只保留表头信息"""
        simplified = self._extract_headers_only(structure)
        text_output = self._structure_to_text(simplified)

        if len(text_output) <= max_length:
            return text_output

        self.logger.info(f"数据结构过长 ({len(text_output)} 字符)，进行压缩...")

        if isinstance(simplified, dict) and "files" in simplified:
            lines = [f"共 {len(simplified.get('files', {}))} 个文件:"]
            for file_name, file_data in list(simplified.get("files", {}).items())[:5]:
                lines.append(f"- {file_name}")
                if isinstance(file_data, dict) and "sheets" in file_data:
                    for sheet_name, sheet_info in file_data["sheets"].items():
                        headers = self._get_headers_from_sheet(sheet_info)
                        if headers:
                            lines.append(f"  {sheet_name}: {', '.join(headers[:10])}")
                            if len(headers) > 10:
                                lines.append(f"    ...还有 {len(headers)-10} 列")
            return '\n'.join(lines)

        if len(text_output) > max_length:
            return text_output[:max_length-50] + '\n...(内容已截断)'

        return text_output

    def _structure_to_text(self, structure: Dict[str, Any]) -> str:
        """将结构转换为简洁的文本格式"""
        lines = []

        if "files" in structure:
            for file_name, file_data in structure.get("files", {}).items():
                lines.append(f"文件: {file_name}")
                if isinstance(file_data, dict) and "sheets" in file_data:
                    for sheet_name, sheet_info in file_data["sheets"].items():
                        headers = self._get_headers_from_sheet(sheet_info)
                        row_count = sheet_info.get("data_row_count", sheet_info.get("row_count", "?"))
                        if headers:
                            lines.append(f"  Sheet[{sheet_name}] ({row_count}行): {', '.join(headers)}")

        elif "sheets" in structure:
            for sheet_name, sheet_info in structure.get("sheets", {}).items():
                headers = self._get_headers_from_sheet(sheet_info)
                row_count = sheet_info.get("data_row_count", sheet_info.get("row_count", "?"))
                if headers:
                    lines.append(f"Sheet[{sheet_name}] ({row_count}行): {', '.join(headers)}")

        if structure.get("file_name"):
            lines.insert(0, f"文件名: {structure['file_name']}")

        return '\n'.join(lines)

    def _get_headers_from_sheet(self, sheet_info: Dict[str, Any]) -> List[str]:
        """从Sheet信息中提取列名列表"""
        if not isinstance(sheet_info, dict):
            return []

        if "headers" in sheet_info:
            return sheet_info["headers"] if isinstance(sheet_info["headers"], list) else list(sheet_info["headers"].keys())
        elif "head_data" in sheet_info:
            return list(sheet_info["head_data"].keys())
        elif "regions" in sheet_info and isinstance(sheet_info["regions"], list):
            for region in sheet_info["regions"]:
                if isinstance(region, dict) and "head_data" in region:
                    return list(region["head_data"].keys())
        return []

    def _extract_headers_only(self, structure: Dict[str, Any]) -> Dict[str, Any]:
        """从数据结构中只提取表头信息"""
        if not isinstance(structure, dict):
            return structure

        simplified = {}

        if "files" in structure:
            simplified["files"] = {}
            for file_name, file_data in structure.get("files", {}).items():
                simplified["files"][file_name] = self._extract_file_headers(file_data)

        for key in ["file_name", "total_sheets", "total_regions", "total_files"]:
            if key in structure:
                simplified[key] = structure[key]

        if "sheets" in structure:
            simplified["sheets"] = {}
            for sheet_name, sheet_data in structure.get("sheets", {}).items():
                simplified["sheets"][sheet_name] = self._extract_sheet_headers(sheet_data)

        if not simplified:
            return structure

        return simplified

    def _extract_file_headers(self, file_data: Dict[str, Any]) -> Dict[str, Any]:
        """从文件数据中提取表头信息"""
        if not isinstance(file_data, dict):
            return file_data

        simplified = {}
        if "sheets" in file_data:
            simplified["sheets"] = {}
            for sheet_name, sheet_data in file_data.get("sheets", {}).items():
                simplified["sheets"][sheet_name] = self._extract_sheet_headers(sheet_data)

        return simplified

    def _extract_sheet_headers(self, sheet_data: Dict[str, Any]) -> Dict[str, Any]:
        """从Sheet数据中提取表头信息"""
        if not isinstance(sheet_data, dict):
            return sheet_data

        simplified = {}

        if "headers" in sheet_data:
            simplified["headers"] = sheet_data["headers"]
        if "head_data" in sheet_data:
            simplified["head_data"] = sheet_data["head_data"]

        for key in ["head_row_start", "head_row_end", "data_row_start", "data_row_end", "row_count"]:
            if key in sheet_data:
                simplified[key] = sheet_data[key]

        if "regions" in sheet_data:
            regions = sheet_data.get("regions", [])
            if isinstance(regions, list):
                simplified["regions"] = []
                for region in regions:
                    if not isinstance(region, dict):
                        continue
                    simplified_region = {}
                    if "head_data" in region:
                        simplified_region["head_data"] = region["head_data"]
                    if "headers" in region:
                        simplified_region["headers"] = region["headers"]
                    for key in ["head_row_start", "head_row_end", "data_row_start", "data_row_end"]:
                        if key in region:
                            simplified_region[key] = region[key]
                    if "data" in region and isinstance(region["data"], list):
                        simplified_region["data_row_count"] = len(region["data"])
                    if simplified_region:
                        simplified["regions"].append(simplified_region)
            elif isinstance(regions, int):
                simplified["regions_count"] = regions

        if not simplified:
            for key, value in sheet_data.items():
                if key not in ["data", "formula", "formulas"]:
                    if isinstance(value, list) and len(value) > 10:
                        simplified[f"{key}_count"] = len(value)
                    else:
                        simplified[key] = value

        return simplified

    def _remove_empty_lines(self, text: str) -> str:
        """去除文本中的空行"""
        if not text:
            return text
        lines = text.split('\n')
        non_empty_lines = [line for line in lines if line.strip()]
        return '\n'.join(non_empty_lines)

    def _compress_rules(self, text: str, max_length: int = 80000) -> str:
        """压缩规则文本，减少token消耗

        处理：
        1. 去除每行首尾空格
        2. 合并连续空格为单个空格
        3. 去除空行和纯装饰行（分隔线、代码块标记）
        4. 去除markdown装饰符号但保留内容层级
        5. 截断到max_length
        """
        import re
        if not text:
            return text

        lines = text.split('\n')
        result = []
        for line in lines:
            # 去除首尾空格
            stripped = line.strip()
            # 跳过空行
            if not stripped:
                continue
            # 跳过纯装饰行：分隔线、代码块标记
            if re.match(r'^[-=_*~`]{3,}$', stripped):
                continue
            if stripped in ('```', '```python', '```text'):
                continue
            # 合并行内连续空格（保留缩进结构的语义）
            stripped = re.sub(r'  +', ' ', stripped)
            # 简化markdown标题：### 标题 → 【标题】（减少#字符）
            header_match = re.match(r'^#{1,6}\s+(.+)$', stripped)
            if header_match:
                stripped = f"【{header_match.group(1)}】"
            result.append(stripped)

        compressed = '\n'.join(result)
        if len(compressed) > max_length:
            compressed = compressed[:max_length]
        return compressed

    def _optimize_prompt(self, prompt: str, target_max_length: int = 35000) -> str:
        """检查提示词长度（不压缩）"""
        original_length = len(prompt)

        if original_length > target_max_length:
            self.logger.warning(f"提示词长度: {original_length} 字符，超过建议长度 {target_max_length}")
        else:
            self.logger.info(f"提示词长度: {original_length} 字符")

        return prompt

    def generate_training_prompt(
        self,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        rules_content: str,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """生成训练提示词"""
        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)

        manual_headers_str = json.dumps(manual_headers, ensure_ascii=False, indent=2) if manual_headers else "无"
        if len(manual_headers_str) > 1000:
            manual_headers_str = self._compress_structure(manual_headers, max_length=30000)

        # 提取数据清洗规则、警告规则和条件格式规则
        extracted_rules = self.rule_extractor.extract_rules(rules_content)
        data_cleaning_rules_text = ""
        warning_rules_text = ""

        if extracted_rules["data_cleaning_rules"] or extracted_rules["warning_rules"] or extracted_rules["conditional_format_rules"]:
            formatted_rules = self.rule_extractor.format_rules_for_prompt(extracted_rules)
            # 分离数据清洗规则和警告规则
            if "## 数据清洗规则" in formatted_rules:
                parts = formatted_rules.split("## 警告信息规则")
                data_cleaning_rules_text = parts[0]
                if len(parts) > 1:
                    warning_rules_text = "## 警告信息规则" + parts[1]
            else:
                warning_rules_text = formatted_rules

        template = self.templates["training"]
        replacements = {
            "{excel_parser_interface}": self.EXCEL_PARSER_INTERFACE,
            "{global_vars}": self.GLOBAL_VARS_DESC,
            "{core_rules}": self.CORE_RULES,
            "{source_structure}": compressed_source,
            "{expected_structure}": compressed_expected,
            "{rules_content}": rules_content,
            "{data_cleaning_rules}": data_cleaning_rules_text,
            "{warning_rules}": warning_rules_text,
            "{manual_headers}": manual_headers_str
        }

        result = template
        for placeholder, value in replacements.items():
            result = result.replace(placeholder, value)

        result = self._optimize_prompt(result, target_max_length=50000)
        self.logger.info(f"生成的提示词长度: {len(result)} 字符")
        return result

    def generate_correction_prompt(
        self,
        original_code: str,
        error_description: str,
        comparison_result: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        rules_content: str,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """生成修正提示词"""
        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)

        manual_headers_str = json.dumps(manual_headers, ensure_ascii=False, indent=2) if manual_headers else "无"
        if len(manual_headers_str) > 800:
            manual_headers_str = self._compress_structure(manual_headers, max_length=30000)

        # 提取数据清洗规则、警告规则和条件格式规则
        extracted_rules = self.rule_extractor.extract_rules(rules_content)
        data_cleaning_rules_text = ""
        warning_rules_text = ""

        if extracted_rules["data_cleaning_rules"] or extracted_rules["warning_rules"] or extracted_rules["conditional_format_rules"]:
            formatted_rules = self.rule_extractor.format_rules_for_prompt(extracted_rules)
            # 分离数据清洗规则和警告规则
            if "## 数据清洗规则" in formatted_rules:
                parts = formatted_rules.split("## 警告信息规则")
                data_cleaning_rules_text = parts[0]
                if len(parts) > 1:
                    warning_rules_text = "## 警告信息规则" + parts[1]
            else:
                warning_rules_text = formatted_rules

        template = self.templates["correction"]
        replacements = {
            "{excel_parser_interface}": self.EXCEL_PARSER_INTERFACE,
            "{global_vars}": self.GLOBAL_VARS_DESC,
            "{core_rules}": self.CORE_RULES,
            "{original_code}": original_code,
            "{error_description}": error_description,
            "{comparison_result}": comparison_result,
            "{source_structure}": compressed_source,
            "{expected_structure}": compressed_expected,
            "{rules_content}": rules_content,
            "{data_cleaning_rules}": data_cleaning_rules_text,
            "{warning_rules}": warning_rules_text,
            "{manual_headers}": manual_headers_str
        }

        result = template
        for placeholder, value in replacements.items():
            result = result.replace(placeholder, value)

        result = self._optimize_prompt(result, target_max_length=25000)
        self.logger.info(f"生成的修正提示词长度: {len(result)} 字符")
        return result

    def generate_validation_prompt(self, code_content: str) -> str:
        """生成验证提示词"""
        template = self.templates["validation"]
        return template.replace("{code_content}", code_content)

    def generate_training_prompt_with_ai_rules(
        self,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        ai_rules: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """使用AI生成的规则生成训练提示词"""
        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)

        manual_headers_str = json.dumps(manual_headers, ensure_ascii=False, indent=2) if manual_headers else "无"

        # 从AI规则中提取关键信息
        mapping_rules = ai_rules.get('column_mapping', {})
        calculation_rules = ai_rules.get('calculation_rules', {})
        processing_steps = ai_rules.get('processing_steps', [])
        summary = ai_rules.get('summary', 'AI生成的规则')

        if len(processing_steps) > 10:
            processing_steps = processing_steps[:10] + [f"... (共 {len(ai_rules.get('processing_steps', []))} 个步骤)"]

        structured_rules = f"""## AI分析结果
{summary}

## 映射规则
{json.dumps(mapping_rules, ensure_ascii=False, indent=2)[:3000]}

## 计算规则
{json.dumps(calculation_rules, ensure_ascii=False, indent=2)[:3000]}

## 处理步骤
{chr(10).join(f"- {step}" for step in processing_steps)}"""

        template = self.templates["training"]
        replacements = {
            "{excel_parser_interface}": self.EXCEL_PARSER_INTERFACE,
            "{global_vars}": self.GLOBAL_VARS_DESC,
            "{core_rules}": self.CORE_RULES,
            "{source_structure}": compressed_source,
            "{expected_structure}": compressed_expected,
            "{rules_content}": structured_rules,
            "{manual_headers}": manual_headers_str
        }

        result = template
        for placeholder, value in replacements.items():
            result = result.replace(placeholder, value)

        result = self._optimize_prompt(result, target_max_length=25000)
        self.logger.info(f"生成的提示词长度: {len(result)} 字符")
        return result

    def generate_column_adjustment_prompt(
        self,
        fill_function: str,
        target_columns: list,
        adjustment_request: str,
        source_structure: dict,
        expected_structure: dict,
        rules_content: str,
        manual_headers: dict = None
    ) -> str:
        """生成单列修正提示词 - AI只返回需要修改的列代码片段

        Args:
            fill_function: 当前 fill_result_sheets 函数代码（可能包含clean_source_data）
            target_columns: 用户指定要修改的列名列表
            adjustment_request: 用户修改说明
            source_structure: 源数据结构
            expected_structure: 预期输出结构
            rules_content: 原始计算规则
            manual_headers: 手动表头映射
        """
        compressed_source = self._compress_structure(source_structure, max_length=20000)
        compressed_expected = self._compress_structure(expected_structure, max_length=15000)
        manual_headers_str = json.dumps(manual_headers, ensure_ascii=False, indent=2) if manual_headers else "无"

        columns_str = "、".join(target_columns)

        # 分离 clean_source_data 和 fill_result_sheets
        only_fill_function = fill_function
        if "def clean_source_data" in fill_function:
            fill_start = fill_function.find("def fill_result_sheets")
            if fill_start == -1:
                fill_start = fill_function.find("def fill_result_sheet")
            if fill_start > 0:
                only_fill_function = fill_function[fill_start:].strip()

        # 提取目标列当前的代码，帮助 AI 理解上下文
        from .formula_code_generator import FormulaCodeGenerator
        current_columns_code = ""
        for col_name in target_columns:
            block, _, _ = FormulaCodeGenerator.extract_column_block(only_fill_function, col_name)
            if block:
                current_columns_code += f"\n### 当前 {col_name} 的代码：\n```python\n{block.strip()}\n```\n"
            else:
                current_columns_code += f"\n### {col_name}：未找到现有代码块（可能是新列）\n"

        prompt = f"""你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师。

# 任务：单列精准修正

用户要求修改以下列：**{columns_str}**
修改要求：{adjustment_request}

## 当前完整 fill_result_sheets 函数（只读参考，不要全部返回）
```python
{only_fill_function}
```

## 目标列当前代码
{current_columns_code}

## 输入文件结构（跨表取数参考）
{compressed_source}

## 预期输出结构
{compressed_expected}

## 计算规则
{rules_content[:70000]}

## 手动表头映射
{manual_headers_str}

{self.CORE_RULES}

# 输出格式要求（严格遵守！）

请分析用户的修改要求，判断实际需要修改哪些列（可能比用户指定的多，比如修改了基本工资的取数方式，依赖它的应发工资等也需要联动修改）。

按以下固定格式返回，**不要返回完整函数**，只返回需要修改的列代码片段：

```
### MODIFIED_COLUMNS: 列名1, 列名2, ...

### COLUMN: X列(N): 列名
        # X列(N): 列名 - 说明
        ws.cell(row=r, column=N).value = ...
### END_COLUMN

### COLUMN: Y列(M): 列名
        # Y列(M): 列名 - 说明
        ws.cell(row=r, column=M).value = ...
### END_COLUMN

### PRE_LOOP_CODE
    col_new_var = get_vlookup_col_num("X", "A")
### END_PRE_LOOP_CODE
```

## 格式说明
1. `### MODIFIED_COLUMNS:` 列出所有实际修改的列名（逗号分隔）
2. 每个 `### COLUMN:` 到 `### END_COLUMN` 之间是一列的完整代码块
3. 代码块必须保持原始缩进（8个空格，即for循环内的缩进）
4. 列注释格式必须是：`# X列(N): 列名 - 说明`，与原代码保持一致
5. `### PRE_LOOP_CODE` 到 `### END_PRE_LOOP_CODE` 之间放需要新增的循环外变量定义（如新的VLOOKUP列号变量），如果不需要新增则省略此段
6. **不要**返回未修改的列
7. **不要**返回完整的 fill_result_sheets 函数
8. **不要**修改 clean_source_data 或警告规则逻辑"""

        self.logger.info(f"生成单列修正提示词（结构化输出模式），目标列: {columns_str}，长度: {len(prompt)} 字符")
        return prompt

    @staticmethod
    def parse_column_adjustment_response(ai_response: str) -> dict:
        """解析 AI 返回的结构化列修正响应

        支持多种AI输出风格：
        - 标准格式（### COLUMN: ... ### END_COLUMN）
        - 包裹在markdown代码块中（```...```）
        - AI可能使用不同的空白或大小写

        Returns:
            {
                "modified_columns": ["列名1", "列名2"],
                "column_blocks": {"列名1": "代码块", "列名2": "代码块"},
                "pre_loop_code": "新增的循环外代码" 或 None
            }
        """
        result = {
            "modified_columns": [],
            "column_blocks": {},
            "pre_loop_code": None
        }

        # 预处理：去掉markdown代码块标记
        cleaned = ai_response
        # 去掉 ```python ... ``` 和 ``` ... ``` 包裹
        cleaned = re.sub(r'```(?:python|py)?\s*\n', '', cleaned)
        cleaned = re.sub(r'\n```\s*', '\n', cleaned)

        # 提取修改列列表
        mod_match = re.search(r'###\s*MODIFIED_COLUMNS[：:]\s*(.+)', cleaned)
        if mod_match:
            result["modified_columns"] = [c.strip() for c in mod_match.group(1).split(",") if c.strip()]

        # 提取每列代码块 - 标准格式
        column_pattern = re.compile(
            r'###\s*COLUMN[：:]\s*([A-Z]{1,3})列[（\(](\d+)[）\)][：:]\s*(.+?)\s*\n(.*?)###\s*END_COLUMN',
            re.DOTALL
        )
        for match in column_pattern.finditer(cleaned):
            col_letter = match.group(1)
            col_num = match.group(2)
            col_name = match.group(3).strip()
            code_block = match.group(4)

            # 清理代码块：去掉首尾空行，但保留缩进
            lines = code_block.split('\n')
            while lines and not lines[0].strip():
                lines.pop(0)
            while lines and not lines[-1].strip():
                lines.pop()
            code_block = '\n'.join(lines)

            if code_block.strip():
                result["column_blocks"][col_name] = code_block

        # 如果标准格式没匹配到，尝试备用格式：
        # AI可能直接返回列注释+代码，没有 ### COLUMN 包裹
        if not result["column_blocks"]:
            # 查找所有列注释块: # X列(N): 列名 - 说明
            fallback_pattern = re.compile(
                r'([ \t]*# ([A-Z]{1,3})列[（\(](\d+)[）\)][：:]\s*(.+?)(?:\s*-\s*.+?)?\n'
                r'(?:[ \t]+.*\n)*)',
                re.MULTILINE
            )
            for match in fallback_pattern.finditer(cleaned):
                full_block = match.group(0)
                col_name = match.group(4).strip()

                # 清理
                lines = full_block.split('\n')
                while lines and not lines[-1].strip():
                    lines.pop()
                full_block = '\n'.join(lines)

                if full_block.strip():
                    result["column_blocks"][col_name] = full_block

        # 提取循环外新增代码
        pre_loop_match = re.search(
            r'###\s*PRE_LOOP_CODE\s*\n(.*?)###\s*END_PRE_LOOP_CODE',
            cleaned,
            re.DOTALL
        )
        if pre_loop_match:
            pre_code = pre_loop_match.group(1).strip()
            if pre_code:
                result["pre_loop_code"] = pre_code

        # 如果 MODIFIED_COLUMNS 为空但有 column_blocks，从 blocks 补全
        if not result["modified_columns"] and result["column_blocks"]:
            result["modified_columns"] = list(result["column_blocks"].keys())

        return result

    def generate_correction_prompt_with_ai_rules(
        self,
        original_code: str,
        error_description: str,
        comparison_result: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        ai_rules: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """使用AI生成的规则生成修正提示词"""
        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)

        manual_headers_str = json.dumps(manual_headers, ensure_ascii=False, indent=2) if manual_headers else "无"

        mapping_rules = ai_rules.get('column_mapping', {})
        calculation_rules = ai_rules.get('calculation_rules', {})
        summary = ai_rules.get('summary', 'AI生成的规则')

        structured_rules = f"""## AI分析结果
{summary}

## 映射规则
{json.dumps(mapping_rules, ensure_ascii=False, indent=2)[:2000]}

## 计算规则
{json.dumps(calculation_rules, ensure_ascii=False, indent=2)[:2000]}"""

        template = self.templates["correction"]
        replacements = {
            "{excel_parser_interface}": self.EXCEL_PARSER_INTERFACE,
            "{global_vars}": self.GLOBAL_VARS_DESC,
            "{core_rules}": self.CORE_RULES,
            "{original_code}": original_code,
            "{error_description}": error_description,
            "{comparison_result}": comparison_result,
            "{source_structure}": compressed_source,
            "{expected_structure}": compressed_expected,
            "{rules_content}": structured_rules,
            "{manual_headers}": manual_headers_str
        }

        result = template
        for placeholder, value in replacements.items():
            result = result.replace(placeholder, value)

        result = self._optimize_prompt(result, target_max_length=25000)
        self.logger.info(f"生成的修正提示词长度: {len(result)} 字符")
        return result

    def extract_rules_from_files(self, rule_files: List[str]) -> str:
        """从规则文件中提取内容"""
        rules_content = []

        for rule_file in rule_files:
            try:
                from .document_parser import get_document_parser
                parser = get_document_parser()
                content = parser.parse_document(rule_file)
                rules_content.append(f"=== 规则文件: {Path(rule_file).name} ===\n{content}\n")
            except Exception as e:
                rules_content.append(f"=== 规则文件: {Path(rule_file).name} (读取失败: {str(e)}) ===\n")

        return "\n".join(rules_content)

    def format_comparison_result(self, actual_data: Dict[str, Any], expected_data: Dict[str, Any]) -> str:
        """格式化对比结果"""
        result = []

        actual_sheets = set(actual_data.get("sheets", {}).keys())
        expected_sheets = set(expected_data.get("sheets", {}).keys())

        if actual_sheets != expected_sheets:
            result.append(f"Sheet不一致: 实际={sorted(actual_sheets)}, 预期={sorted(expected_sheets)}")

        for sheet_name in actual_sheets.intersection(expected_sheets):
            actual_headers = actual_data["sheets"][sheet_name].get("headers", {})
            expected_headers = expected_data["sheets"][sheet_name].get("headers", {})

            actual_header_names = set(actual_headers.keys())
            expected_header_names = set(expected_headers.keys())

            if actual_header_names != expected_header_names:
                result.append(f"Sheet '{sheet_name}' 表头不一致")

            actual_rows = len(actual_data["sheets"][sheet_name].get("data", []))
            expected_rows = len(expected_data["sheets"][sheet_name].get("data", []))

            if actual_rows != expected_rows:
                result.append(f"Sheet '{sheet_name}' 行数不一致: 实际={actual_rows}, 预期={expected_rows}")

        max_diff_record = self._find_max_diff_record_by_primary_key(actual_data, expected_data)
        if max_diff_record:
            result.append(f"\n差异最多的记录 (共{max_diff_record['diff_count']}处差异):")
            for diff in max_diff_record['diffs'][:10]:
                result.append(f"  [{diff['field']}]: 实际={diff['actual']}, 预期={diff['expected']}")

        return "\n".join(result) if result else "所有检查项都通过！"

    def _find_max_diff_record_by_primary_key(self, actual_data: Dict[str, Any], expected_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """基于主键匹配找出差异最多的记录"""
        max_diff_record = None
        max_diff_count = 0

        actual_sheets = set(actual_data.get("sheets", {}).keys())
        expected_sheets = set(expected_data.get("sheets", {}).keys())

        for sheet_name in actual_sheets.intersection(expected_sheets):
            actual_sheet = actual_data["sheets"][sheet_name]
            expected_sheet = expected_data["sheets"][sheet_name]

            actual_rows = actual_sheet.get("data", [])
            expected_rows = expected_sheet.get("data", [])
            expected_headers = expected_sheet.get("headers", {})
            expected_col_to_name = {v: k for k, v in expected_headers.items()}

            primary_key_name = self._detect_primary_key(expected_sheet)
            primary_key_col = expected_headers.get(primary_key_name) if primary_key_name else None

            actual_by_pk = {}
            if primary_key_col:
                for row in actual_rows:
                    pk_value = row.get(primary_key_col)
                    if pk_value is not None:
                        actual_by_pk[pk_value] = row

            for i, expected_row in enumerate(expected_rows):
                actual_row = None
                pk_value = None

                if primary_key_col:
                    pk_value = expected_row.get(primary_key_col)
                    actual_row = actual_by_pk.get(pk_value)

                if actual_row is None and i < len(actual_rows):
                    actual_row = actual_rows[i]

                if actual_row is None:
                    continue

                diffs = []
                for col_letter, expected_value in expected_row.items():
                    field_name = expected_col_to_name.get(col_letter, col_letter)
                    actual_value = actual_row.get(col_letter)

                    if actual_value != expected_value:
                        if self._values_approximately_equal(actual_value, expected_value):
                            continue
                        diffs.append({'field': field_name, 'actual': actual_value, 'expected': expected_value})

                if len(diffs) > max_diff_count:
                    max_diff_count = len(diffs)
                    max_diff_record = {
                        'sheet_name': sheet_name,
                        'diff_count': len(diffs),
                        'diffs': diffs,
                        'primary_key': primary_key_name,
                        'pk_value': pk_value
                    }

        return max_diff_record

    def _detect_primary_key(self, sheet_data: Dict[str, Any]) -> Optional[str]:
        """检测主键列"""
        headers = sheet_data.get("headers", {})
        header_names = list(headers.keys()) if isinstance(headers, dict) else (headers if isinstance(headers, list) else [])

        primary_key_candidates = ['工号', '员工编号', '雇员工号', '编号', 'ID', 'id', '序号', '姓名', '雇员姓名']
        for candidate in primary_key_candidates:
            if candidate in header_names:
                return candidate
        return None

    def _values_approximately_equal(self, val1: Any, val2: Any, tolerance: float = 0.01) -> bool:
        """检查两个值是否近似相等"""
        try:
            if val1 is None or val2 is None:
                return val1 == val2
            return abs(float(val1) - float(val2)) < tolerance
        except (ValueError, TypeError):
            return False

    # ============ 批量模块化提示词 ============

    def generate_batch_modular_prompt(
        self,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        manual_headers: Optional[Dict[str, Any]] = None,
        modules: List[Dict[str, Any]] = None,
        salary_year: Optional[int] = None,
        salary_month: Optional[int] = None,
        monthly_standard_hours: Optional[float] = None
    ) -> str:
        """生成批量模块化提示词 - 精简版"""

        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)
        rules = self._compress_rules(rules_content, max_length=80000)
        manual_headers_json = json.dumps(manual_headers or {}, ensure_ascii=False)

        salary_params = ""
        if salary_year: salary_params += f"salary_year = {salary_year}\n"
        if salary_month: salary_params += f"salary_month = {salary_month}\n"
        if monthly_standard_hours: salary_params += f"monthly_standard_hours = {monthly_standard_hours}\n"

        return f"""你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景。根据规则生成数据处理代码。

{self.EXCEL_PARSER_INTERFACE}

## 必须生成的6个函数
1. load_all_data(input_folder) - 数据加载
2. generate_field_mapping(data_store, rules_content) - 字段映射
3. create_output_template(mapping, expected_structure) - 输出模板
4. generate_formulas(mapping, rules_content) - 公式生成
5. fill_data(data_store, template, mapping, formulas) - 数据填充【核心】
6. save_excel_with_details(...) - Excel保存

## 核心约束
1. 路径: os.path.join(input_folder/output_folder, filename)
2. fill_data返回4元组: (result, column_sources, column_formulas, intermediate_columns)
3. 每列设置column_sources["列名"]="来源说明"
4. 计算列设置column_formulas["列名"]="={{列A}}+{{列B}}"
5. 用safe_get_column(df, "列名", 默认值)访问列
6. 用源表初始化: base_df = 源表.copy()
7. 完整实现每个计算，禁止`= 0 # 简化`
8. 必须有完整main()函数

## 常见错误
❌ 日薪 = 基本工资 / 21.75 → 变量未定义
✓ base_df["日薪"] = safe_get_column(base_df, "基本工资", 0) / 21.75

❌ df[["列名"]].apply(lambda x: ...) → x是Series
✓ df["列名"].apply(lambda x: ...) → x是单值

## 规则内容
{rules}

## 源文件结构
{compressed_source}

## 预期输出结构
{compressed_expected}

## 全局变量
input_folder, output_folder, manual_headers: {manual_headers_json}
{salary_params}

请输出完整可执行的Python代码，包含所有import和6个函数定义。"""

    # ============ Excel公式模式提示词 ============

    def generate_formula_mode_prompt(
        self,
        source_structure: str,
        expected_structure: Dict[str, Any],
        rules_content: str,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """生成Excel公式模式的提示词 - 精简版"""
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)
        rules = self._compress_rules(rules_content, max_length=80000)
        _ = manual_headers

        # 使用集中式辅助方法提取规则和统计列数
        rules_info = self._extract_cleaning_and_format_rules(rules_content)
        expected_sheets_info, total_columns = self._build_expected_sheets_info(expected_structure, rules_content)
        golden_and_rules = self._build_golden_and_rules(total_columns)

        template = """你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师。

## 任务概览
本次共 __TOTAL_COLUMNS__ 列，必须全部生成，不允许省略。
1. 生成 clean_source_data 函数：应用数据清洗规则过滤源数据
2. 生成 fill_result_sheets 函数：创建结果sheet，填充数据和Excel公式

## 执行流程
1. 源数据已加载到 source_data 字典（DataFrame格式）
2. 你生成clean_source_data函数，应用清洗规则过滤数据
3. 清洗后的数据写入Excel供公式引用
4. 参数sheet已创建（参数!$B$2=年份, $B$3=月份, $B$4=月标准工时）
5. 你生成fill_result_sheets函数，创建结果sheet并填充公式

## 已有变量
- wb: openpyxl Workbook
- source_data: {"文件名_sheet名": {"df": DataFrame, "columns": [列名]}} （清洗前）
- source_sheets: {"文件名_sheet名": {"df": DataFrame, "ws": worksheet}} （清洗后，写入Excel）
- 已导入模块：os, pandas(pd), openpyxl(Workbook, Comment, PatternFill, Font, CellIsRule, FormulaRule, get_column_letter, column_index_from_string)
- 已定义常量：EMPTY = Excel空字符串""
- 已定义函数：excel_text('文本') = Excel文本值"文本"
- 已定义函数：get_vlookup_col_num(target_col, range_start_col) -> int
- 已定义函数：find_source_sheet(source_sheets, target_columns=[...], sheet_name_hint="...") -> key

__SOURCE_STRUCTURE__
__EXPECTED_SHEETS_INFO__

⚠️ 如果规则文档中包含「列处理分层」信息，请严格按照层级处理：L1同源列用main_df.iloc[i].get()直接复制（不查列号），L2跨表列才使用VLOOKUP查找。
⚠️ 主表选择：优先使用规则中「列处理分层→主表」指定的表，用find_source_sheet()定位。禁止用max(len(df))——行数最多的不一定是主表！如果规则未指定主表，根据主键列名和特征列推断。

## 预期输出结构
__COMPRESSED_EXPECTED__

__DATA_CLEANING_RULES__

__CONDITIONAL_FORMAT_RULES__

__PRECISION_RULES__

## 计算规则
__RULES__

"""


        return (template
                .replace('__TOTAL_COLUMNS__', str(total_columns))
                .replace('__SOURCE_STRUCTURE__', source_structure)
                .replace('__EXPECTED_SHEETS_INFO__', expected_sheets_info)
                .replace('__COMPRESSED_EXPECTED__', compressed_expected)
                .replace('__DATA_CLEANING_RULES__', rules_info["data_cleaning"])
                .replace('__CONDITIONAL_FORMAT_RULES__', rules_info["conditional_format"])
                .replace('__PRECISION_RULES__', rules_info["precision"])
                .replace('__RULES__', rules)
                + "\n" + golden_and_rules)

    def generate_formula_batch_prompt(
        self,
        batch_index: int,
        total_batches: int,
        batch_columns: List[Dict[str, str]],
        all_columns_overview: str,
        source_structure: str,
        rules_content: str,
        existing_code: str = None,
        first_batch_context: str = None,
    ) -> str:
        """生成分批模式的提示词 — 每批生成独立函数

        新策略：
        - 第1批：生成主函数 fill_result_sheets（含表头、for循环、前N列逻辑）
        - 第2~N批：生成独立的 fill_columns_batch_N 函数
        - 主函数的for循环内会调用各批次函数

        Args:
            batch_index: 当前批次索引（从0开始）
            total_batches: 总批次数
            batch_columns: 当前批次的列信息
            all_columns_overview: 所有列的概览
            source_structure: 源数据结构描述
            rules_content: 与当前批次列相关的规则
            existing_code: 前面批次已生成的代码（第一批为None）
            first_batch_context: 第一批代码中的关键变量上下文

        Returns:
            提示词字符串
        """
        batch_col_list = "\n".join([
            f"  - {c['col_letter']}列: {c['col_name']}（Sheet: {c['sheet']}）"
            for c in batch_columns
        ])

        # 生成后续批次函数调用列表（供第一批使用）
        batch_calls = ""
        if total_batches > 1:
            calls = []
            for i in range(1, total_batches):
                calls.append(f"        fill_columns_batch_{i + 1}(ws, r, source_sheets)")
            batch_calls = "\n".join(calls)

        formula_rules_detailed = self._build_formula_rules("detailed")
        formula_rules_compact = self._build_formula_rules("compact")

        if batch_index == 0:
            return f"""你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师。

## 任务说明
由于列数较多，分{total_batches}批生成。本次生成主函数 fill_result_sheets + 第1批（{len(batch_columns)}列）的逻辑。
后续批次会生成独立的 fill_columns_batch_2, fill_columns_batch_3... 函数，在主函数的for循环内调用。

## 执行流程（固定代码已处理）
1. ✅ 源数据已加载到 source_sheets 字典
2. ✅ 源数据已写入Excel（供公式引用）
3. ✅ 参数sheet已创建（参数!$B$2=年份, $B$3=月份, $B$4=月标准工时）
4. 🎯 你的任务：创建结果sheet，填充数据和Excel公式
5. ✅ 生成代码时要注意引用列的位置，不能假设列位置，必须用get_vlookup_col_num()函数计算列号
6. ✅ 结果sheet的表头必须完全匹配预期结构，列顺
7. ✅ 注意生成代码的完整性，一次性生成完整的代码
8. ✅ 参与计算的日期必须用DATEVALUE()转换
9. ✅ 规则中指定的特殊规则必须严格遵守，禁止任何形式的简化处理

## 已有变量
- wb: openpyxl Workbook
- source_sheets: {{"文件名_sheet名": {{"df": DataFrame, "ws": worksheet}}}}

{source_structure}

## 全部列概览
{all_columns_overview}

## 本批次需要实现的列（第1批，共{len(batch_columns)}列）
{batch_col_list}

## 本批次相关的计算规则
{rules_content}

⚠️ 如果规则中包含「列处理分层」，L1同源列直接用main_df.iloc[i].get()复制（不查列号），L2跨表列才用VLOOKUP。

{formula_rules_detailed}

# ==================== 代码结构要求 ====================
请生成以下结构的代码：

```python
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    # 1. 主表选择（用find_source_sheet定位，禁止max(len(df))）
    # 2. 源表sheet标题映射（所有源表的key和ws_title变量）
    # 3. 创建结果sheet
    # 4. 写【全部列】的表头（不只是本批次，是所有列的表头）
    # 5. for循环逐行填充：
    for i in range(n_rows):
        r = i + 2
        # 本批次的列逻辑（第1批）
        ...
        # 调用后续批次函数
{batch_calls}
    # 6. 条件格式（如果规则要求标红/高亮等，在循环结束后用CellIsRule/FormulaRule实现）
```

重要：
- 源表变量（如 att_ws_title, memo_ws_title 等）必须定义在for循环外面
- for循环内调用 fill_columns_batch_2(ws, r, source_sheets) 等后续函数
- 表头必须包含全部列，不只是本批次的列"""

        else:
            # 后续批次：生成独立函数
            return f"""你是专业Python程序员，需要生成一个独立的列填充函数。

## 任务说明
这是第{batch_index + 1}批（共{total_batches}批）。请生成一个独立函数 `fill_columns_batch_{batch_index + 1}`。
该函数会在主函数的for循环内被调用，每次处理一行。

## 第1批代码中的关键变量（你的函数需要通过source_sheets参数获取这些信息）
{first_batch_context}

## 函数签名（必须严格遵守）
```python
def fill_columns_batch_{batch_index + 1}(ws, r, source_sheets):
    \"\"\"填充第{batch_index + 1}批列（行号r）\"\"\"
    # 从source_sheets获取需要的sheet标题
    # 然后用ws.cell(row=r, column=列号, value=公式) 填充每列
```

## 本批次需要实现的列（第{batch_index + 1}批，共{len(batch_columns)}列）
{batch_col_list}

## 本批次相关的计算规则
{rules_content}

## 源数据结构
{source_structure}

⚠️ 如果规则中包含「列处理分层」，L1同源列直接用main_df.iloc[i].get()复制（不查列号），L2跨表列才用VLOOKUP。

{formula_rules_compact}

## 要求
1. 生成完整的 fill_columns_batch_{batch_index + 1} 函数定义
2. 函数内部先从source_sheets获取需要的ws_title变量
3. 用 ws.cell(row=r, column=列号, value=...) 填充每列
4. 本批次的每一列都必须实现，不允许遗漏
5. 列号必须正确：按照预期输出结构中的列顺序"""

    # ============ 5步模块化提示词（保留接口，简化实现）============

    def generate_modular_step_prompt(
        self,
        step_number: int,
        step_name: str,
        rules_content: str,
        source_structure: Dict[str, Any],
        expected_structure: Dict[str, Any],
        previous_modules: List[Dict[str, str]] = None,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> str:
        """生成5步模块化中每一步的提示词"""
        compressed_source = self._compress_structure(source_structure, max_length=30000)
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)
        rules = self._compress_rules(rules_content, max_length=80000)

        step_prompts = {
            1: f"""生成数据加载模块 load_all_data(input_folder) -> Dict[str, Any]

{self.EXCEL_PARSER_INTERFACE}

要求：
1. 使用IntelligentExcelParser加载Excel
2. 返回 {{"files": {{filename: {{"sheets": {{sheet_name: DataFrame}}}}}}, "structure": ...}}
3. 使用传入的input_folder参数，禁止硬编码路径

源文件结构：
{compressed_source}""",

            2: f"""生成映射关系模块 generate_field_mapping(data_store, rules_content) -> Dict[str, Any]

返回：
- direct_mapping: 直接复制的字段映射
- calculated_fields: 需要计算的字段列表

规则内容：
{rules}

源文件结构：{compressed_source}
预期输出：{compressed_expected}""",

            3: f"""生成模板模块 create_output_template(mapping, expected_structure) -> Dict[str, pd.DataFrame]

返回：{{sheet_name: empty_dataframe_with_headers}}
确保列顺序与预期一致。

预期输出结构：
{compressed_expected}""",

            4: f"""生成公式模块 generate_formulas(mapping, rules_content) -> List[Dict[str, Any]]

返回计算任务列表，每个包含：
- target_column: 目标列名
- formula_type: 公式类型
- source_columns: 依赖的源列
- formula_func: 计算函数
- priority: 优先级

任务按依赖顺序排列。

规则内容：
{rules}""",

            5: f"""生成数据填充模块 fill_data(data_store, template, mapping, formulas) -> Dict[str, pd.DataFrame]

处理流程：
1. 确定数据行数
2. 填充直接映射字段
3. 按顺序执行公式计算
4. 处理异常和空值

规则内容：
{rules}

源文件：{compressed_source}
预期输出：{compressed_expected}"""
        }

        return step_prompts.get(step_number, f"无效步骤: {step_number}")

    # ============ 生成+验证模式提示词 ============

    def generate_multi_step_prompts(
        self,
        source_structure: str,
        expected_structure: Dict[str, Any],
        rules_content: str,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, str]:
        """生成+验证模式的提示词

        同一对话2轮完成：
        Step 3: 生成代码（包含完整上下文）
        Step 4: 验证并修正代码

        Returns:
            {"system": 系统提示词, "step3": 生成代码, "step4": 验证, "total_columns": 总列数}
        """
        compressed_expected = self._compress_structure(expected_structure, max_length=30000)
        rules = self._compress_rules(rules_content, max_length=80000)

        # 统一提取规则
        extracted = self._extract_cleaning_and_format_rules(rules_content)

        # 统一构建预期Sheet信息
        expected_sheets_info, total_columns = self._build_expected_sheets_info(expected_structure, rules_content)

        # 黄金样例 + 补充规则 + 函数签名（在f-string中需要转义花括号）
        golden_and_rules = self._build_golden_and_rules(total_columns, escape_braces=True)

        system_prompt = (
            "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
            "包括人力资源、财务、供应链等不同业务场景。"
            "你同时也是一个EXCEL公式大师，熟悉VLOOKUP、IF、SUMIF等公式的使用。"
        )

        # Step 3: 生成代码（包含完整上下文）
        step3 = f"""你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师。

## 任务概览
本次共 {total_columns} 列，必须全部生成，不允许省略。
1. 生成 clean_source_data 函数：应用数据清洗规则过滤源数据
2. 生成 fill_result_sheets 函数：创建结果sheet，填充数据和Excel公式

## 执行流程
1. 源数据已加载到 source_data 字典（DataFrame格式）
2. 你生成clean_source_data函数，应用清洗规则过滤数据
3. 清洗后的数据写入Excel供公式引用
4. 参数sheet已创建（参数!$B$2=年份, $B$3=月份, $B$4=月标准工时）
5. 你生成fill_result_sheets函数，创建结果sheet并填充公式

## 已有变量
- wb: openpyxl Workbook
- source_data: {{"文件名_sheet名": {{"df": DataFrame, "columns": [列名]}}}} （清洗前）
- source_sheets: {{"文件名_sheet名": {{"df": DataFrame, "ws": worksheet}}}} （清洗后，写入Excel）
- 已导入模块：os, pandas(pd), openpyxl(Workbook, Comment, PatternFill, Font, CellIsRule, FormulaRule, get_column_letter, column_index_from_string)
- 已定义常量：EMPTY = Excel空字符串""
- 已定义函数：excel_text('文本') = Excel文本值"文本"
- 已定义函数：get_vlookup_col_num(target_col, range_start_col) -> int
- 已定义函数：find_source_sheet(source_sheets, target_columns=[...], sheet_name_hint="...") -> key

{source_structure}
{expected_sheets_info}

## 预期输出结构
{compressed_expected}
{extracted['data_cleaning']}
{extracted['conditional_format']}
{extracted['precision']}

## 计算规则
{rules}

{golden_and_rules}"""

        # Step 4: 验证修正（占位符__GENERATED_CODE__在使用时替换）
        verification = self.VERIFICATION_CHECKLIST.replace("__TOTAL_COLUMNS__", str(total_columns))
        step4 = f"""## 任务：验证并修正生成的代码

请逐项检查以下代码，找出问题并输出修正后的完整代码。

## 生成的代码
```python
__GENERATED_CODE__
```

{verification}

## 输出要求
1. 先输出发现的问题列表
2. 然后输出修正后的完整fill_result_sheets函数代码
3. 如果没有问题，输出"无需修正"和原始代码

⚠️ 必须输出完整的函数代码，不能只输出片段。"""

        return {
            "system": system_prompt,
            "step3": step3,
            "step4": step4,
            "total_columns": total_columns
        }
