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
        self._src_data_rows: int = 0               # 源数据最大行数

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
        src_data_rows: int = 0,
    ) -> Tuple[str, str]:
        """生成智能组表填充代码。

        Args:
            pre_mapped: 知识库已命中映射 {源列名: 模板列名}，AI 直接采用不再分析
            tpl_data_rows: 模板数据区现有数据行数（行数不足时引导 AI 复制样式扩展）
            src_data_rows: 源数据最大行数
        """
        self._pre_mapped = pre_mapped or {}
        self._tpl_data_rows = int(tpl_data_rows or 0)
        self._src_data_rows = int(src_data_rows or 0)
        return super().generate_code(
            input_folder, rules_content, template_path, manual_headers,
            stream_callback, thinking_callback,
            multi_sheet_source=True,   # 智能组表源始终读全部可见 sheet
            use_history=use_history,
            target_sheets=target_sheets,
            expected_structure=expected_structure,
        )

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

        # 0) 输出格式硬性要求（AI 经常直接输出顶层代码导致提取失败）
        parts.append(
            "【输出格式硬性要求】你的输出必须是一个完整的 Python 函数定义，第一行必须是：\n"
            "def fill_template(wb, source_data, salary_year, salary_month, monthly_standard_hours):\n"
            "函数体内实现全部填充逻辑（可内部定义变量/import，但不要出现函数定义之外的顶层执行语句，"
            "不要输出 ```python 围栏，不要输出解释性文字）。"
        )

        # 1) 数据行数处理指引（模板模式同款：AI 在 openpyxl 内直接写入/扩展行）
        if self._src_data_rows > 0:
            hint = (f"【数据行数指引】模板激活 sheet 数据区现有数据行数为 {self._tpl_data_rows}，"
                    f"源数据最大行数为 {self._src_data_rows}。"
                    f"若源数据行数超过模板数据区行数，fill_template 写入时超出模板数据区的行"
                    f"必须复制模板最后一行数据行的样式（openpyxl: "
                    f"from copy import copy; dst_cell._style = copy(src_cell._style)，"
                    f"并复制行高 ws.row_dimensions）。"
                    f"不要使用插入行/删除行（会破坏模板样式），直接在新行上写值并复制样式即可。")
            parts.append(hint)

        # 2) 知识库已命中映射（直接采用，不要重新分析）
        if self._pre_mapped:
            lines = ["【已有确定匹配关系（来自匹配知识库，直接采用，不要改动）】"]
            for src, dst in self._pre_mapped.items():
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
