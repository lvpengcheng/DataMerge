"""智能组表代码生成器 — 模板模式生成器的子类化改造。

与 TemplateCodeGenerator 的差异：
1. 模板只解析激活 sheet（active_sheet_only=True）
2. 源解析保留全部可见 sheet，样例取前 3 行（数据行 <3 取全部）并脱敏
3. Prompt 额外注入：脱敏后的源样例数据、知识库已命中映射、数据区扩展说明

复用父类的骨架拼接 / fill_template 抽取 / 修复管线。
"""

import os
import re
import logging
from typing import Dict, List, Any, Optional, Tuple

from .template_code_generator import TemplateCodeGenerator

logger = logging.getLogger(__name__)


class AssembleCodeGenerator(TemplateCodeGenerator):
    """智能组表：AI 写填充逻辑，骨架处理模板复制/源sheet追加/样式保持。"""

    def __init__(self, ai_provider=None, training_logger=None,
                 file_passwords: Optional[Dict] = None):
        super().__init__(ai_provider=ai_provider, training_logger=training_logger)
        self.file_passwords = file_passwords or {}
        self._pre_mapped: Dict[str, str] = {}      # 知识库已命中映射 {源列: 模板列}
        self._tpl_data_rows: int = 0               # 模板数据区现有数据行数

    # ==================== 主入口（附加参数） ====================

    def generate_code(
        self,
        input_folder: str,
        rules_content: str,
        template_path: str,
        manual_headers: Dict = None,
        stream_callback: callable = None,
        thinking_callback: callable = None,
        multi_sheet_source: bool = True,
        use_history: bool = False,
        target_sheets: Optional[List[str]] = None,
        expected_structure: Optional[Dict] = None,
        pre_mapped: Optional[Dict[str, str]] = None,
        tpl_data_rows: int = 0,
    ) -> Tuple[str, str]:
        """生成智能组表填充代码。

        Args:
            pre_mapped: 知识库已命中映射 {源列名: 模板列名}，AI 直接采用不再分析
            tpl_data_rows: 模板数据区现有数据行数（行数不足时引导 AI 复制样式扩展）
        """
        self._pre_mapped = pre_mapped or {}
        self._tpl_data_rows = int(tpl_data_rows or 0)
        code, ai_response = super().generate_code(
            input_folder, rules_content, template_path, manual_headers,
            stream_callback, thinking_callback,
            multi_sheet_source=True,   # 智能组表源始终读全部可见 sheet
            use_history=use_history,
            target_sheets=target_sheets,
            expected_structure=expected_structure,
        )
        # 静态检查 1：fill_template 必须包含"值赋值"（主键列填充）。
        # 只写公式的代码会导致主键列为空、全部公式落空（真实事故）。
        assignments = re.findall(r'\.value\s*=\s*([^\n]+)', code)
        value_assigns = [a for a in assignments if not a.strip().startswith(
            ('f"=', "f'=", '"=', "'=", 'pattern'))]
        if not value_assigns:
            raise RuntimeError(
                "AI 生成的代码只有公式、没有主键列的值赋值（主键列必须从源_表填入实际值），"
                "无法填充主键，请重新生成")
        # 静态检查 2：fill_template 必须从 FIELD_MAPPING 驱动（人工修正映射 rematch 才能
        # 直接替换常量生效，无需重新 AI 生成）
        if "FIELD_MAPPING" not in code:
            raise RuntimeError(
                "AI 生成的代码没有从 FIELD_MAPPING 读取映射驱动（映射必须集中定义在 "
                "FIELD_MAPPING 常量中，fill_template 遍历它动态构造公式/取值），"
                "请重新生成")
        return code, ai_response

    # ==================== 模板解析：只读激活 sheet ====================

    def _parse_template(self, template_path: str) -> Dict[str, Any]:
        try:
            from excel_parser import IntelligentExcelParser
        except ImportError:
            from ...excel_parser import IntelligentExcelParser

        parser = IntelligentExcelParser()
        sheets_data = parser.parse_excel_file(
            template_path,
            max_data_rows=5,
            read_formulas=True,
            skip_hidden_sheets=False,
            active_sheet_only=True,     # 智能组表模板只读激活 sheet
        )

        out = {"file_name": os.path.basename(template_path), "sheets": {}}
        for sd in sheets_data:
            sheet_info = {"regions": []}
            for region in (sd.regions or []):
                head = region.head_data or {}
                cols = []
                for col_name in head.keys():
                    has_formula = self._region_col_has_formula(region, col_name)
                    cols.append({
                        "name": col_name,
                        "column_letter": self._guess_column_letter(region, col_name),
                        "has_formula_in_data": has_formula,
                    })
                # 空模板修正：模板只有表头没有数据行时，data_row_start 为 0，
                # 退回表头下一行，保证 _COL_MAP 里该 region 不被过滤（AI 才能看到模板列）。
                ds_raw = getattr(region, "data_row_start", None) or 0
                de_raw = getattr(region, "data_row_end", None) or 0
                header_end = getattr(region, "head_row_end", None) or 0
                ds = ds_raw if ds_raw > 0 else (header_end + 1)
                sheet_info["regions"].append({
                    "header_row": getattr(region, "head_row_start", None),
                    "header_row_end": header_end,
                    "data_start_row": ds,
                    "data_end_row": de_raw if de_raw >= ds_raw else None,
                    "columns": cols,
                })
            out["sheets"][sd.sheet_name] = sheet_info
        return out

    # ==================== 源解析：全部可见 sheet + 前3行脱敏样例 ====================

    def _parse_source(self, input_folder: str, manual_headers: Optional[Dict],
                      multi_sheet_source: bool) -> Dict[str, Any]:
        try:
            from excel_parser import IntelligentExcelParser
        except ImportError:
            from ...excel_parser import IntelligentExcelParser
        from backend.utils.desensitize import desensitize_region

        parser = IntelligentExcelParser()
        out = {"files": {}}
        for fname in os.listdir(input_folder):
            fp = os.path.join(input_folder, fname)
            if not os.path.isfile(fp):
                continue
            if not fname.lower().endswith((".xlsx", ".xls", ".xlsm")):
                continue
            try:
                password = self.file_passwords.get(fname) or None
                sheets_data = parser.parse_excel_file(
                    fp, max_data_rows=3, read_formulas=True,
                    active_sheet_only=False,          # 源读全部可见 sheet
                    best_region_only=True,
                    manual_headers=manual_headers,
                    password=password,
                )
                fdesc = {}
                for sd in sheets_data:
                    regions = sd.regions or []
                    if not regions:
                        fdesc[sd.sheet_name] = {"columns": []}
                        continue
                    region = regions[0]
                    head = region.head_data or {}
                    cols = list(head.keys())
                    # 脱敏样例：数据行前 3 行（数据行不足 3 行取全部）
                    sample = desensitize_region(head, (region.data or [])[:3])
                    fdesc[sd.sheet_name] = {
                        "columns": cols,
                        "head_data": head,
                        "sample_rows": sample,
                        "formula_columns": [
                            c for c in cols if self._region_col_has_formula(region, c)
                        ],
                        "column_formats": list((region.column_formats or {}).keys()),
                    }
                out["files"][fname] = fdesc
            except Exception as e:
                logger.warning(f"[assemble] 源文件 {fname} 解析失败: {e}")
                out["files"][fname] = {"error": str(e)}
        return out

    # ==================== 提取兜底：AI 直接输出顶层代码时包成 fill_template ====================

    def _extract_fill_function(self, ai_response: str) -> str:
        """父类找不到 def fill_template 时，若响应含代码块则整体包成 fill_template 函数。

        真实场景：AI 不按 prompt 输出函数定义，直接写顶层执行代码（引用 wb/_SOURCE_MAP 等
        全局），父类提取失败。这里把代码块内容缩进后包进函数体，参数与骨架调用签名一致。
        """
        text = (ai_response or "").strip()
        if not text:
            return ""

        res = super()._extract_fill_function(text)
        if res:
            return res

        # 兜底：取第一个代码块（无代码块时取整段），包成 fill_template
        blocks = re.findall(r"```(?:[a-zA-Z0-9_+-]*)\s*(.*?)\s*```", text, re.DOTALL)
        body = (blocks[0] if blocks else text).strip()
        if not body:
            return ""
        # 去掉 AI 可能输出的非代码前后缀（如"以下是代码："），取从 import/from/def/变量赋值起的内容
        lines = body.splitlines()
        start = 0
        for i, ln in enumerate(lines):
            if ln.strip() and not ln.strip().startswith(("#", "//", "以下", "这是", "代码")):
                start = i
                break
        body = "\n".join(lines[start:]).strip()
        indented = "\n".join(("    " + ln) if ln.strip() else ln for ln in body.splitlines())
        return (f"def fill_template(wb, source_data, salary_year, salary_month, "
                f"monthly_standard_hours):\n{indented}")

    # ==================== Prompt：追加脱敏样例 + 知识库映射 + 扩展说明 ====================

    def _build_prompt(
        self,
        rules_content: str,
        template_struct: Dict[str, Any],
        source_struct: Dict[str, Any],
        use_history: bool,
        reserved_names: Optional[set] = None,
    ) -> str:
        prompt = super()._build_prompt(rules_content, template_struct, source_struct,
                                       use_history, reserved_names=reserved_names)

        parts = ["\n\n# ==================== 智能组表附加信息 ===================="]

        # 0.1) 组表模式识别（决定整合策略）
        parts.append(
            "【组表模式识别】先根据源表结构与规则判断本次属于哪种模式，再按对应策略生成 fill_template：\n"
            "模式1（同构整合）：多个源表列结构相似（如上海/北京/南京等地同构数据表），"
            "需要把所有源数据按模板格式整合 → 策略：各源 sheet 数据【纵向拼接】→ 按主键"
            "（身份证/工号等）合并，同主键的【数值列加总、文本列取首值去重】→ 写入模板数据行。\n"
            "模式2（职责分工）：各源表职责不同（如上月发放/本月离职/本月入职/薪资变更等），"
            "每个表可能只取一列或几列 → 策略：按主键【关联合并】（merge/VLOOKUP），"
            "每表取对应列汇总进结果表。\n"
            "判断依据：源表列集合高度重合 → 模式1；列集合差异大、各有专责列 → 模式2。"
            "拿不准时按模式2（主键关联补字段）处理，并在 report 里说明判断。"
        )

        # 0) 输出格式硬性要求（AI 经常直接输出顶层代码导致提取失败）
        parts.append(
            "【输出格式硬性要求】你的输出必须是一个完整的 Python 函数定义，第一行必须是：\n"
            "def fill_template(wb, source_data, salary_year, salary_month, monthly_standard_hours):\n"
            "函数体内实现全部填充逻辑（可内部定义变量/import，但不要出现函数定义之外的顶层执行语句，"
            "不要输出 ```python 围栏，不要输出解释性文字）。\n\n"
            "【B1. 进度透明化】你可以在 fill_template 函数内调用 `report_progress(msg: str)` 推送进度信息，"
            "如 report_progress('开始填充基础列') 或 report_progress(f'已填充{i}/{total}行')，"
            "前端会实时显示这些信息。建议在关键步骤（解析数据/填充列块/写公式/计算）加进度埋点，"
            "让用户知道脚本在做什么，避免长时间无反馈显得卡死。"
        )

        # 0.3) 结构化字段映射清单（供错误反馈弹窗逐列复核，必须输出）
        parts.append(
            "【字段映射清单输出要求】在 `def fill_template(...)` 的**正上方、同一个代码块内**，"
            "先输出一个模块级常量 `FIELD_MAPPING`（这是唯一允许出现在函数定义之外的顶层赋值语句），"
            "格式严格如下：\n"
            "```\n"
            "FIELD_MAPPING = {\n"
            "    \"目标模板列名\": {\"source_column\": \"对应源列名\", \"source_letter\": \"源列字母\", \"confidence\": 0.9},\n"
            "    ...\n"
            "}\n"
            "```\n"
            "规则：\n"
            "1. **必须覆盖模板激活 sheet 的全部列**（_COL_MAP 里每个目标列名都要作为 key 出现）。\n"
            "2. source_column 填该目标列对应取的**源列名**（_SOURCE_MAP 里 源_ sheet 的列名，不带「源_」前缀）；"
            "source_letter 填该源列在源 sheet 里的**列字母**（如 A/B/C）。\n"
            "3. 无法确定映射的目标列也要列出：source_column 填空字符串 \"\"、source_letter 空字符串、confidence 给 0。\n"
            "4. confidence 是匹配把握 0~1：知识库命中/同名列 = 1.0，AI 语义近似匹配按把握给 0.5~0.95。\n"
            "5. 该清单只用于人工复核展示，必须与 fill_template 的实际填充逻辑保持一致。"
        )

        # 0.5) 主键列填充 + 非主键列公式硬性要求（防止只写公式不填主键 → 全部落空；
        #      也防止只填值不写公式 → 原版失去公式可追溯性）
        parts.append(
            "【主键列填充硬性要求】模板的身份主键列（列名含「外服工号/工号/员工编号/编号/Payroll ID」"
            "或「证件号码/身份证/证件号」的列）必须从源_ sheet **填入实际值**（逐行把源表对应主键列的值"
            "直接赋给模板主键列单元格，如 ws.cell(row=r, column=...).value = 源值），"
            "**绝不允许对主键列写公式**（尤其不能写引用自身行的 VLOOKUP/INDEX，会导致循环引用）。\n"
            "【非主键列必须写跨 sheet 公式】除主键列外，其他所有需要填充的列**必须写跨 sheet 公式**"
            "（VLOOKUP/INDEX/MATCH 引用主键列与 源_ sheet，如 "
            "=VLOOKUP(B2,'源_xxx'!B:C,2,FALSE)），**不要写死数值**"
            "——原版必须保留公式、可追溯，纯值版才会在后续转成值。\n"
            "【公式写法硬性要求】**不要用 IFERROR/VALUE 等包裹公式**，直接写裸公式"
            "（如 =VLOOKUP(...)、=INDEX(...)）。查不到就是 #N/A、类型不对就是 #VALUE!，"
            "错了就留错误值，不要吞掉。每个公式必须括号匹配、逗号分隔参数个数正确，"
            "IFERROR 绝不允许出现 3 个参数。\n"
            "【字段映射驱动要求】fill_template 的填充逻辑**必须从 FIELD_MAPPING 常量读取映射驱动**"
            "（遍历 FIELD_MAPPING，按 {目标列: {source_sheet, source_column, source_letter}} "
            "动态构造公式/取值），"
            "**不得把列映射硬编码在代码里**（如写死 source_col_letter('ID') 之类）——"
            "人工复核修正映射时只替换 FIELD_MAPPING 就能生效，无需重新生成代码。\n"
            "读取映射字段必须用 .get() 容错（如 info.get('source_letter')）：人工修正的映射"
            "可能缺 source_letter，缺失时按 source_sheet + source_column 从 _SOURCE_MAP 查列字母，"
            "查不到则跳过该列并 report 提示。\n"
            "【字段映射输出要求】在 def fill_template 函数定义的上方（同一代码块内），"
            "必须额外输出结构化字段映射，供人工复核：\n"
            "FIELD_MAPPING = {\n"
            "    \"目标列名\": {\"source_sheet\": \"源_表名\", \"source_column\": \"源列名\", "
            "\"source_letter\": \"源列字母\", \"confidence\": 0.95},\n"
            "    ...\n"
            "}\n"
            "覆盖所有参与填充的列（含主键列）；source_sheet 是该源列实际取数所在的 源_ sheet 名"
            "（必须与 _SOURCE_MAP 的 key 完全一致），source_column 为对应源列名，"
            "source_letter 是源_ sheet 中该列的字母，confidence 为 0~1 的匹配置信度。\n"
            "函数体内必须包含至少一条『值赋值』语句（主键列 .value = 非公式）；"
            "只有公式没有值赋值、或只有值没有公式的输出都是无效的。"
        )

        # 1) 数据行数处理指引：源样例仅前 3 行（结构示例，避免 prompt 膨胀），实际源行数
        #    可能远大于模板数据区。不依赖"源行数估计值"，无条件引导 AI 按 source_data
        #    实际行数全量循环 + 超行复制样式扩展，避免源多行时结果表下半部分丢样式。
        hint = (f"【数据行数指引】模板激活 sheet 数据区现有数据行数为 {self._tpl_data_rows}。"
                f"⚠️ 前面给出的源文件样例数据**仅前 3 行**（结构示例，不是完整数据）——"
                f"源数据实际行数可能远大于模板数据区行数。fill_template 必须按 "
                f"source_data / 源_ sheet 的**实际行数**循环填充全部数据行，"
                f"绝不能只填样例中出现的行数。"
                f"写入超出模板数据区的行时，必须复制模板最后一行数据行的样式"
                f"（openpyxl: from copy import copy; dst_cell._style = copy(src_cell._style)，"
                f"并复制行高 ws.row_dimensions）。"
                f"若模板数据区为空（只有表头），则从表头下一行开始逐行写入，"
                f"新行样式复制表头行样式或保持默认。"
                f"不要使用插入行/删除行（会破坏模板样式），直接在新行上写值并复制样式即可。")
        parts.append(hint)

        # 1.5) 单元格格式规范（硬性要求：防止数值被套日期格式 → 结果表百万级错误显示）
        parts.append(
            "【单元格格式规范（硬性要求，覆盖前文『绝不修改 number_format』约束）】"
            "⚠️ 本规范明确要求设置 number_format，优先级高于工作流约束第 5 条——"
            "写入单元格时严格遵守：\n"
            "1. 长数字列（身份证号/证件号/银行卡号/手机号/工号等 ≥12 位数字或含字母的编号）："
            "**必须按文本写入**，即先设 `cell.number_format = '@'` 再赋 `str(值)`，"
            "绝不写成数值（否则超长数字会显示成科学计数/被当日期序列号）。\n"
            "2. 金额/数量/比例等数值列：写入数值即可，**禁止设置日期格式**；"
            "若复制的模板样式恰好是日期格式且该列不是日期列，必须把该格 `number_format = 'General'`。\n"
            "3. 只有列名含「日期/时间/生日/入职」等日期关键词的列才使用日期格式（写 datetime 值并保留日期样式）。\n"
            "4. 复制模板最后一行样式到新行后，若新行单元格的值类型与样式不匹配（数值格套日期样式等），"
            "以第 1/2/3 条为准修正该格格式。\n"
            "5. **数值列优先写值**（`cell.value = 数值`），不要对数值列写 VLOOKUP/SUMIF 公式——"
            "数据量大时百万级公式会让 Excel 打开/计算极慢；只有规则明确要求计算结果的列才写公式。"
        )

        # 2) 知识库已命中映射（直接采用，不要重新分析）
        if self._pre_mapped:
            lines = ["【已有确定匹配关系（来自匹配知识库/人工复核，直接采用，不要改动）】"]
            for src, dst in self._pre_mapped.items():
                # 值格式：目标列 或 目标列@源表名（人工复核指定源表）
                _tgt, _, _sheet = str(dst).partition("@")
                if _sheet:
                    lines.append(f"- 源列「{src}」→ 模板列「{_tgt}」（源表 {_sheet}）")
                else:
                    lines.append(f"- 源列「{src}」→ 模板列「{dst}」")
            parts.append("\n".join(lines))

        # 3) 脱敏样例数据
        sample_parts = ["【源文件样例数据（前3行，已脱敏，仅用于理解数据结构与格式）】"]
        files = source_struct.get("files", {}) or {}
        for fname, fdesc in files.items():
            if not isinstance(fdesc, dict) or "error" in fdesc:
                continue
            for sn, sinfo in fdesc.items():
                if not isinstance(sinfo, dict):
                    continue
                cols = sinfo.get("columns") or []
                rows = sinfo.get("sample_rows") or []
                if not cols:
                    continue
                sample_parts.append(f"\n文件 {fname} / sheet {sn}，列: {cols}")
                if rows:
                    sample_parts.append("样例数据（脱敏后）:")
                    for r in rows:
                        sample_parts.append("  " + str(r))
                else:
                    sample_parts.append("（无数据行样例）")
                fm = sinfo.get("formula_columns") or []
                if fm:
                    sample_parts.append(f"公式列: {fm}")
        parts.append("\n".join(sample_parts))

        return prompt + "\n".join(parts)
