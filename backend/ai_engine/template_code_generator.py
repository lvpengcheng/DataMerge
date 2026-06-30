"""
模板填充模式代码生成器 — AI 直接写公式，骨架预追加源 sheet（前缀 `源_`）

工作流：
1. 用户上传**带公式/格式/合并单元格的目标文件**作为模板
2. 骨架预处理：
   - shutil.copy 模板 → openpyxl 加载副本（保留公式/格式/合并）
   - 把源数据 DataFrame 追加为 wb 的新 sheet，命名前缀 `源_`（避免与模板 sheet 重名）
3. AI 写 `fill_template(wb, source_sheets, ...)`：
   - 只在**规则文档要求**的列上写入；规则没提到的列绝不触碰
   - 优先写**跨 sheet 公式**（`=VLOOKUP/IF/SUMIFS('源_xxx'!...)`）
   - 跳过模板里 `has_formula=True` 的列（保留模板原公式）
4. 骨架保存 wb 到 output_folder
"""

import os
import re
import json
import logging
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

from .ai_provider import BaseAIProvider, AIProviderFactory
from backend.utils.indentation_fixer import IndentationFixer

logger = logging.getLogger(__name__)


class TemplateCodeGenerator:
    """模板填充模式代码生成器"""

    def __init__(self, ai_provider: BaseAIProvider = None, training_logger=None):
        if ai_provider is None:
            self.ai_provider = AIProviderFactory.create_with_fallback()
        else:
            self.ai_provider = ai_provider
        self.training_logger = training_logger
        self._indent_fixer = IndentationFixer()
        self.last_prompt = None

    # ==================== 主入口 ====================

    def generate_code(
        self,
        input_folder: str,
        rules_content: str,
        template_path: str,
        manual_headers: Dict = None,
        stream_callback: callable = None,
        multi_sheet_source: bool = False,
        use_history: bool = False,
        target_sheets: Optional[List[str]] = None,
    ) -> Tuple[str, str]:
        """生成模板填充模式的 Python 代码

        Args:
            input_folder: 源数据文件目录
            rules_content: 规则文本
            template_path: 模板文件持久化路径（写入生成的脚本中）
            manual_headers: 源数据手动表头配置
            stream_callback: 流式日志回调
            multi_sheet_source: 源是否多 sheet
            use_history: 是否启用历史数据
            target_sheets: 用户在前端勾选的目标 sheet 名列表；为空则用全部 sheet

        Returns:
            (完整的 Python 脚本, AI 原始响应)
        """
        def log(msg):
            logger.info(msg)
            if stream_callback:
                stream_callback(msg)

        log("=== 模板填充模式：开始生成代码 ===")

        if not template_path or not os.path.exists(template_path):
            raise ValueError(f"模板文件不存在或路径无效: {template_path}")

        # 1. 解析模板结构
        log("步骤1: 解析模板结构（sheet/列/已有公式）...")
        template_struct = self._parse_template(template_path)
        log(f"模板含 {len(template_struct.get('sheets', {}))} 个 sheet")

        # 1.1 按 target_sheets 过滤（仅保留用户勾选的 sheet）
        if target_sheets:
            _all_sheets = template_struct.get("sheets", {})
            _kept = {sn: info for sn, info in _all_sheets.items() if sn in target_sheets}
            _dropped = [sn for sn in _all_sheets.keys() if sn not in target_sheets]
            template_struct["sheets"] = _kept
            template_struct["target_sheets"] = list(target_sheets)
            log(f"按用户勾选过滤：保留 {len(_kept)} 个 sheet（丢弃 {len(_dropped)}：{_dropped[:5]}{'...' if len(_dropped) > 5 else ''}）")

        # 2. 解析源数据列结构
        log("步骤2: 解析源数据列...")
        source_struct = self._parse_source(input_folder, manual_headers, multi_sheet_source)
        log(f"源数据含 {len(source_struct.get('files', {}))} 个文件")

        # 3. 构造 AI prompt
        log("步骤3: 构造 AI prompt...")
        prompt = self._build_prompt(rules_content, template_struct, source_struct, use_history)
        self.last_prompt = prompt

        # 4. 调用 AI 生成 fill_template 函数
        log("步骤4: 调用 AI 生成 fill_template 函数...")
        ai_response = self._call_ai(prompt, stream_callback)

        # 5. 提取 fill_template 函数代码
        fill_function = self._extract_fill_function(ai_response)
        if not fill_function:
            preview = (ai_response or "")[:2000]
            log(f"[提取失败] AI 响应预览（前2000字符）：\n{preview}")
            logger.error(f"无法提取 fill_template，AI 响应长度={len(ai_response or '')}，预览：{preview[:500]}")
            raise RuntimeError("AI 响应中无法提取 fill_template 函数（已把响应预览写入训练日志，可在前端聊天流里查看）")

        # 6. 拼接完整脚本
        log("步骤5: 拼接完整脚本...")
        complete_code = self._build_complete_code(
            fill_function=fill_function,
            template_path=template_path,
            input_folder=input_folder,
            template_struct=template_struct,
            source_struct=source_struct,
            use_history=use_history,
        )

        # 7. 缩进修复
        try:
            complete_code = self._indent_fixer.fix(complete_code)
        except Exception as e:
            logger.warning(f"缩进修复异常（不阻断）: {e}")

        # 7.5 接入公式模式同款的"通用代码修复管线"（4层兜底，借用 ai_provider 实现）
        #     做：① _fix_invalid_paths ② _fix_fstring_quotes（4策略轮试 swap/escape/strip-f）
        #         ③ ast.parse → 失败则 fix_general 缩进 ④ 仍失败则 fix_sandbox_pipeline
        #         ⑤ 最后 _detect_and_fix_truncation 修截断
        try:
            if hasattr(self.ai_provider, "validate_and_fix_code_format"):
                fixed = self.ai_provider.validate_and_fix_code_format(complete_code)
                if fixed and isinstance(fixed, str):
                    complete_code = fixed
                    log("步骤7.5: 已通过 ai_provider 通用修复管线")
        except Exception as e:
            logger.warning(f"ai_provider 修复管线异常（不阻断）: {e}")

        # 8. Python 语法校验：所有修复后仍残留语法错则提前拦截
        self._validate_syntax(complete_code, log)

        log("=== 模板填充模式：代码生成完成 ===")
        return complete_code, ai_response

    def _validate_syntax(self, code: str, log) -> None:
        """compile 校验生成代码；失败时把出错行 +- 3 行预览写入训练日志，方便定位"""
        try:
            compile(code, "<generated_template_code>", "exec")
            return
        except SyntaxError as se:
            lines = code.splitlines()
            ln = se.lineno or 0
            start = max(1, ln - 3)
            end = min(len(lines), ln + 3)
            preview_lines = []
            for i in range(start, end + 1):
                marker = " >>> " if i == ln else "     "
                preview_lines.append(f"{i:>5}{marker}{lines[i - 1]}")
            preview = "\n".join(preview_lines)
            log(f"[语法错误] line {ln}, col {se.offset}: {se.msg}\n附近代码：\n{preview}")
            logger.error(f"AI 生成代码语法错误: line {ln}, msg={se.msg}")
            raise RuntimeError(
                f"AI 生成的 fill_template 代码有 Python 语法错误："
                f"line {ln}: {se.msg}（多半是 f-string 引号嵌套；请点击【执行修正】让 AI 重写）"
            ) from se

    # ==================== 模板解析 ====================

    def _parse_template(self, template_path: str) -> Dict[str, Any]:
        """解析模板：复用 IntelligentExcelParser 识别复杂表头"""
        try:
            from excel_parser import IntelligentExcelParser
        except ImportError:
            from ...excel_parser import IntelligentExcelParser  # 兜底

        parser = IntelligentExcelParser()
        sheets_data = parser.parse_excel_file(
            template_path,
            max_data_rows=5,           # 表头识别只需少量数据行
            read_formulas=True,        # 必须读公式（要识别"已有公式列"）
            skip_hidden_sheets=False,
        )

        out = {"file_name": os.path.basename(template_path), "sheets": {}}
        for sd in sheets_data:
            sheet_info = {"regions": []}
            for region in (sd.regions or []):
                head = region.head_data or {}
                cols = []
                for col_name, col_meta in head.items():
                    # col_meta 可能是 list/dict/Cell 对象，取首个数据行的 formula 标志
                    has_formula = self._region_col_has_formula(region, col_name)
                    cols.append({
                        "name": col_name,
                        "column_letter": self._guess_column_letter(region, col_name),
                        "has_formula_in_data": has_formula,
                    })
                # ⚠️ ExcelRegion 的真实字段名是 head_row_start/head_row_end/data_row_start/data_row_end
                #    （不是 header_rows/data_start_row），早期误用 getattr 取到 None → 退化成硬编码 2 → 写穿表头
                sheet_info["regions"].append({
                    "header_row": getattr(region, "head_row_start", None),
                    "header_row_end": getattr(region, "head_row_end", None),
                    "data_start_row": getattr(region, "data_row_start", None),
                    "data_end_row": getattr(region, "data_row_end", None),
                    "columns": cols,
                })
            out["sheets"][sd.sheet_name] = sheet_info
        return out

    def _region_col_has_formula(self, region, col_name: str) -> bool:
        """检查某列在数据行内是否有公式（保护性判断）"""
        try:
            data = getattr(region, "data", None) or []
            for row in data[:5]:
                if isinstance(row, dict):
                    v = row.get(col_name)
                    if isinstance(v, str) and v.startswith("="):
                        return True
        except Exception:
            pass
        return False

    def _guess_column_letter(self, region, col_name: str) -> Optional[str]:
        """从 head_data 推断列字母（如 'D'）"""
        try:
            head = region.head_data or {}
            keys = list(head.keys())
            if col_name in keys:
                idx = keys.index(col_name)
                # 简化：region 不一定从 A 列开始，但大多数情况下足够
                from openpyxl.utils import get_column_letter
                start_col = getattr(region, "start_col", 1) or 1
                return get_column_letter(start_col + idx)
        except Exception:
            pass
        return None

    # ==================== 源数据解析 ====================

    def _parse_source(self, input_folder: str, manual_headers: Optional[Dict], multi_sheet_source: bool) -> Dict[str, Any]:
        """轻量解析源数据列结构，仅供 AI 知道有哪些列可用"""
        try:
            from excel_parser import IntelligentExcelParser
        except ImportError:
            from ...excel_parser import IntelligentExcelParser

        parser = IntelligentExcelParser()
        out = {"files": {}}
        for fname in os.listdir(input_folder):
            fp = os.path.join(input_folder, fname)
            if not os.path.isfile(fp):
                continue
            if not fname.lower().endswith((".xlsx", ".xls", ".xlsm")):
                continue
            try:
                sheets_data = parser.parse_excel_file(
                    fp, max_data_rows=2, read_formulas=False,
                    active_sheet_only=not multi_sheet_source,
                    best_region_only=True,   # 取每 sheet 的最优(主数据)区域，与执行时 _load_full_source_data 一致
                    manual_headers=manual_headers,
                )
                fdesc = {}
                for sd in sheets_data:
                    # best_region_only=True 时每 sheet 只剩最优区域，regions[0] 即主数据区的列。
                    # ⚠️ 不取 region 的 data_start_row：骨架 _append_source_sheets 会把 DataFrame
                    #    重写为「row1=表头, row2+=数据」的干净 sheet，所以追加后源 sheet 的数据
                    #    恒从第 2 行开始，与原文件结构无关。data_start_row 固定为 2。
                    regions = sd.regions or []
                    cols: List[str] = []
                    if regions:
                        head = (regions[0].head_data or {})
                        cols.extend(list(head.keys()))
                    fdesc[sd.sheet_name] = {"columns": cols}
                out["files"][fname] = fdesc
            except Exception as e:
                logger.warning(f"源文件 {fname} 解析失败: {e}")
                out["files"][fname] = {"error": str(e)}
        return out

    # ==================== AI Prompt ====================

    def _build_prompt(
        self,
        rules_content: str,
        template_struct: Dict[str, Any],
        source_struct: Dict[str, Any],
        use_history: bool,
    ) -> str:
        rules_short = (rules_content or "")[:30000]

        # ⚠️ prompt 里展示的结构必须 = 运行时 _COL_MAP/_SOURCE_MAP 的字段名
        col_map_for_prompt = self._build_col_map(template_struct)
        source_map_for_prompt = self._build_source_map_with_letters(source_struct)

        return f"""你是 Excel 数据处理专家。任务：编写 Python 函数 `fill_template`，按规则**写公式**到模板的目标列。

## 工作流（骨架已为你完成的部分）
1. 模板已经被加载为 openpyxl Workbook（保留全部公式/格式/合并单元格）
2. **源数据已被骨架追加为 wb 的新 sheet**，统一加 `源_` 前缀（如 `源_考勤明细`、`源_津贴`），与模板 sheet 区分
3. 公式可用的**两类查找数据源**：
   - (a) 骨架追加的源文件 sheet（`_SOURCE_MAP`，带 `源_` 前缀）
   - (b) 模板自带的 sheet（`_COL_MAP` 里的 sheet 本身也可当查找表/数据源，用真实 sheet 名引用）
4. 你的工作：**只在规则文档要求的列**上写入公式；其他列（含纯数据源 sheet）绝不触碰

## 严格约束（违反必败）
1. **只写规则要求的列**：规则没提到的列、模板里 `has_formula=true` 的列**一律跳过**
2. **【起始行强约束 + 多区域必读】每个 sheet 的可写区域来自 `_COL_MAP[sheet]["regions"]`（列表，可能不止一个）：**
   - 每个 region 各有自己的 `data_start_row`，必须**逐区域**从各自的 `data_start_row` 写起
   - **绝不**用第一个区域的行号套到全表，**绝不**硬编码 1 / 2 / 3
   - 模板常有多行表头/标题，`data_start_row` 可能 = 4、5、甚至更大，已由 excel_parser 识别并固化在 _COL_MAP
   - 写法范本：
     ```python
     for region in _COL_MAP[sheet_name]["regions"]:
         data_start = region["data_start_row"]      # ← 必须从这取
         data_end = region.get("data_end_row")      # 可能为 None
         for i in range(row_count):
             r = data_start + i                      # ← data_start 起步
             ...
     ```
   - **错误**：`for r in range(2, n+2): ...`（硬编码起始 2，会写穿表头）
   - 读源 sheet 时用 `_SOURCE_MAP[源_xxx]["data_start_row"]`（恒为 2，骨架已把源表头重写到第 1 行）
3. **优先写公式**：用 `cell.value = "=VLOOKUP(...)"` 等跨 sheet 公式让 Excel 重算
   - 写值也行（`cell.value = 100`），但能写公式就别写值
4. **公式的查找数据源分两类，必须按类型用对 sheet 名**：
   - **(a) 骨架追加的源文件** → 用 `_SOURCE_MAP` 里**带 `源_` 前缀**的 key（如 `'源_考勤明细'`）。例：`=VLOOKUP(B5,'源_考勤明细'!A:Z,3,FALSE)`
   - **(b) 模板自带的 sheet**（已在 `_COL_MAP` 里、规则要求当数据源/查找表用）→ 用它的**真实 sheet 名，不加 `源_` 前缀**（如 `'员工基础表'`）。它本来就在 wb 里，可直接被其他 sheet 的公式引用。例：`=VLOOKUP(B5,'员工基础表'!A:Z,4,FALSE)`
   - 判断依据：sheet 名出现在 `_SOURCE_MAP` → 用 (a)；出现在 `_COL_MAP` → 用 (b) 真实名
   - 两类的 sheet 名都用 `'` 单引号包裹（Excel 规范）
   - 引用模板自带 sheet 做查找时用整列范围（`A:Z`），不受其表头行数影响
5. **绝不**修改 cell.style/fill/font/border/number_format/alignment
6. **绝不**调 Workbook() 创建新 wb；**绝不** wb.save（外层会调）
7. **绝不**硬编码 sheet 名/列字母字面量；用 `_COL_MAP` / `_SOURCE_MAP` 取
8. **【Excel 公式拼写强约束 — 必读】写 VLOOKUP/IF/SUMIFS 公式时 f-string 引号必须严格分层**：
   - 公式里 sheet 名用 `'` 单引号包裹（Excel 规范）
   - 因此外层 f-string 必须用 **双引号** `f"..."`，**不能用** `f'...'`
   - 空字符串 `""` 直接拼骨架注入的常量 `EMPTY`（=`'""'`），写法：`f"=IFERROR(VLOOKUP(...),{{EMPTY}})"`
   - **正确**：`formula = f"=IFERROR(VLOOKUP(B{{r}},'源_考勤'!A:Z,3,FALSE),{{EMPTY}})"`
   - **错误**：`formula = f'=IFERROR(VLOOKUP(...,'源_考勤'!A:Z,...),"")'`  ← 单引号嵌套立刻 SyntaxError

## 函数签名
```python
def fill_template(wb, source_data, salary_year, salary_month, monthly_hours):
    \"\"\"
    wb: openpyxl Workbook（已加载模板 + 已追加 `源_xxx` sheet）
    source_data: dict[sheet_key -> {{'df': DataFrame, 'columns': [...]}}]，原始 DataFrame（如果你要 Python 直接计算）
    salary_year/salary_month/monthly_hours: 上下文参数

    标准套路（必须严格按这个写起始行 + 逐区域）：
        from openpyxl.utils import column_index_from_string
        for sheet_name, sinfo in _COL_MAP.items():
            if sheet_name not in wb.sheetnames: continue
            ws = wb[sheet_name]

            # 取源 sheet 的真实 sheet 名 + 数据起始行（恒为 2）
            src_key = "源_考勤明细"   # 来自 _SOURCE_MAP 的 key
            if src_key not in wb.sheetnames: continue
            src_data_start = _SOURCE_MAP[src_key]["data_start_row"]   # = 2
            src_ws = wb[src_key]

            # 逐区域写：每个区域用自己的 data_start_row，绝不套用第一个区域
            for region in sinfo["regions"]:
                data_start = region["data_start_row"]   # ← 必须从这取，不要硬编码！
                row_count = src_ws.max_row - src_data_start + 1   # 源数据行数
                for i in range(row_count):
                    r = data_start + i              # ← 目标行：本区域 data_start 起步
                    src_r = src_data_start + i      # ← 源行：源 data_start(=2) 起步
                    ws.cell(row=r, column=column_index_from_string("D")).value = (
                        f"=VLOOKUP(B{{r}},'{{src_key}}'!A:Z,5,FALSE)"
                    )
    \"\"\"
    from openpyxl.utils import column_index_from_string
    ...
```

## _COL_MAP（运行时已注入；目标模板的列结构，**按区域分组**）
**结构**：`{{ sheet_name: {{ "regions": [ {{ "header_row": int, "data_start_row": int, "data_end_row": int|None, "columns": [ {{ "letter": "A", "name": "列名", "has_formula": bool }}, ... ] }}, ... ] }} }}`
- 一个 sheet 的 `regions` 可能有多个（上下堆叠的多张子表），**每个区域各自从它的 `data_start_row` 写起**
- `letter` 是目标列字母（写公式时用 `column_index_from_string(letter)` 转列号）
- `has_formula=true` 的列**一律跳过**（保留模板原公式）
- **这些 sheet 本身也可当查找数据源**：若规则要求从某模板 sheet 里查数据，直接用它的真实 sheet 名引用（如 `'员工基础表'!A:Z`），**不加 `源_` 前缀**

```json
{json.dumps(col_map_for_prompt, ensure_ascii=False, indent=2)[:8000]}
```

## _SOURCE_MAP（运行时已注入；**已带 `源_` 前缀**的源 sheet 结构）
**结构**：`{{ "<源_sheet名>": {{ "header_row": int, "data_start_row": int, "columns": [ {{ "letter": "A", "name": "列名" }}, ... ] }} }}`
- key 已经是骨架追加到 wb 后的真实 sheet 名（带 `源_` 前缀），可直接 `wb[key]` 访问
- 公式中 sheet 名也用这个 key（如 `'源_考勤明细'!A:Z`）
- `letter` 是源 sheet 中该列的列字母，写 VLOOKUP 时用得上
- `data_start_row`：源数据真实起始行；遍历源数据务必从这行开始，**不要硬编码 2**

```json
{json.dumps(source_map_for_prompt, ensure_ascii=False, indent=2)[:8000]}
```

## 规则文档（核心输入）
{rules_short}

## 输出要求
**只输出 `def fill_template(...)` 函数体**，用 ```python 代码块包裹。不要 import 模块（除了从 openpyxl.utils 临时导入）、不要 main、不要解释文字。
"""

    # ==================== AI 调用 ====================

    def _call_ai(self, prompt: str, stream_callback=None) -> str:
        messages = [
            {"role": "system", "content": "你是 Python + openpyxl + Excel 公式专家，擅长 HR 薪酬场景的模板填充。严格按用户给定的输出格式回答。"},
            {"role": "user", "content": prompt},
        ]
        try:
            # 流式：让前端实时看到 AI 生成过程
            if stream_callback and hasattr(self.ai_provider, "chat_stream"):
                from datetime import datetime as _dt
                def _on_chunk(chunk: str):
                    if not chunk:
                        return
                    ts = _dt.now().strftime("%H:%M:%S")
                    stream_callback(f"[{ts}] [CODE] {chunk}")
                resp = self.ai_provider.chat_stream(
                    messages,
                    chunk_callback=_on_chunk,
                    temperature=0.1,
                    max_tokens=8000,
                )
            else:
                resp = self.ai_provider.chat(messages, temperature=0.1, max_tokens=8000)
        except Exception as e:
            logger.error(f"AI 调用失败: {e}", exc_info=True)
            raise
        return resp or ""

    # ==================== 代码提取 ====================

    def _extract_fill_function(self, ai_response: str) -> str:
        """从 AI 响应中抽取 def fill_template 函数。

        鲁棒策略：
        1. 优先扫所有 ```python ... ``` / ``` ... ``` 代码块，取**第一个含 `def fill_template`** 的块
        2. 没有代码块时回退整段文本
        3. 全部失败返回空字符串（外层抛错并把响应预览输出到前端）
        """
        if not ai_response:
            return ""
        text = ai_response.strip()

        blocks = re.findall(r"```(?:[a-zA-Z0-9_+-]*)\s*(.*?)\s*```", text, re.DOTALL)
        for blk in blocks:
            if "def fill_template" in blk:
                return blk.strip()

        if "def fill_template" in text:
            cleaned = re.sub(r"^```[a-zA-Z0-9_+-]*\s*|\s*```\s*$", "", text, flags=re.MULTILINE)
            return cleaned.strip()

        return ""

    # ==================== 列映射构建（注入生成代码作为常量） ====================

    def _build_col_map(self, template_struct: Dict[str, Any]) -> Dict[str, Any]:
        """把模板结构精简为 {sheet_name: {"regions": [{header_row, data_start_row, data_end_row, columns:[{letter,name,has_formula}]}, ...]}}

        多区域支持：一个 sheet 内若有多个数据区域（如上下堆叠的两张表），每个区域都带自己的
        data_start_row，AI 须分别从各区域的数据起始行写公式，绝不能用第一个区域的行号套到全表。

        过滤策略：
        - data_start_row 无效（None / <=0，即只有表头没有数据的标题区域）的 region 直接丢弃
        - 一个 region 都没有的 sheet 仍保留空 regions（让 fill_template 感知 sheet 存在）
        """
        out: Dict[str, Any] = {}
        sheets = template_struct.get("sheets", {}) or {}
        for sn, sinfo in sheets.items():
            regions = sinfo.get("regions", []) or []
            region_list = []
            for r in regions:
                ds = r.get("data_start_row")
                # 只有表头没有数据的区域：data_row_start 为 0/None，跳过
                if ds is None or (isinstance(ds, int) and ds <= 0):
                    continue
                hr = r.get("header_row")
                de = r.get("data_end_row")
                cols_brief = []
                for c in r.get("columns", []) or []:
                    cols_brief.append({
                        "letter": c.get("column_letter"),
                        "name": c.get("name"),
                        "has_formula": bool(c.get("has_formula_in_data", False)),
                    })
                region_list.append({
                    "header_row": int(hr) if isinstance(hr, int) and hr > 0 else None,
                    "data_start_row": int(ds),
                    "data_end_row": int(de) if isinstance(de, int) and de > 0 else None,
                    "columns": cols_brief,
                })
            if not region_list:
                logger.warning(f"[_build_col_map] sheet '{sn}' 无有效数据区域（可能全为标题/表头），regions=[]")
            out[sn] = {"regions": region_list}
        return out

    def _build_source_map(self, source_struct: Dict[str, Any]) -> Dict[str, Any]:
        """{file_name: {sheet_name: [col_name, ...]}}（保留旧版，调试用）"""
        out: Dict[str, Any] = {}
        files = source_struct.get("files", {}) or {}
        for fname, fdesc in files.items():
            if not isinstance(fdesc, dict) or "error" in fdesc:
                continue
            entry = {}
            for sn, sinfo in fdesc.items():
                # 兼容新结构（dict）和旧结构（list）
                if isinstance(sinfo, dict):
                    entry[sn] = list(sinfo.get("columns") or [])
                elif isinstance(sinfo, list):
                    entry[sn] = list(sinfo)
            out[fname] = entry
        return out

    SOURCE_PREFIX = "源_"

    def _build_source_map_with_letters(self, source_struct: Dict[str, Any]) -> Dict[str, Any]:
        """把 source_struct 转为骨架追加 sheet 后的实际结构：
        {
          "源_<原sheet名>": {
            "header_row": 1,
            "data_start_row": 2,          # ← 恒为 2（见下方说明）
            "columns": [ {"letter": "A", "name": "工号"}, ... ]
          }
        }

        - key 已加 `源_` 前缀（与骨架 _append_source_sheets 一致），AI 可直接用作公式中的 sheet 名
        - 若多文件 sheet 同名则自动加 `_<文件 stem>` 后缀避免冲突
        - ⚠️ data_start_row 恒为 2：骨架把源 DataFrame 重写为 row1=表头/row2+=数据 的干净 sheet，
          追加后的源 sheet 与原文件的多行表头/标题行无关，数据必从第 2 行起。
        """
        from openpyxl.utils import get_column_letter
        out: Dict[str, Any] = {}
        used_names = set()
        files = source_struct.get("files", {}) or {}
        for fname, fdesc in files.items():
            if not isinstance(fdesc, dict) or "error" in fdesc:
                continue
            for sn, sinfo in fdesc.items():
                # 兼容新结构（dict）和旧结构（list）
                if isinstance(sinfo, dict):
                    cols = list(sinfo.get("columns") or [])
                else:
                    cols = list(sinfo or [])

                base_key = f"{self.SOURCE_PREFIX}{sn}"
                final_key = base_key
                if final_key in used_names:
                    stem = Path(fname).stem
                    final_key = f"{self.SOURCE_PREFIX}{stem}_{sn}"
                final_key = final_key[:31]  # Excel sheet 名上限
                used_names.add(final_key)
                out[final_key] = {
                    "header_row": 1,
                    "data_start_row": 2,
                    "columns": [
                        {"letter": get_column_letter(i + 1), "name": c}
                        for i, c in enumerate(cols)
                    ],
                }
        return out

    # ==================== 完整脚本拼接 ====================

    def _build_complete_code(
        self,
        fill_function: str,
        template_path: str,
        input_folder: str,
        template_struct: Dict[str, Any],
        source_struct: Dict[str, Any],
        use_history: bool,
    ) -> str:
        """把 AI 生成的 fill_template 包到固定的脚本骨架里"""
        import pprint as _pprint
        tpl_repr = repr(str(Path(template_path).resolve()))

        # ⚠️ 必须用 pprint.pformat 输出 Python 字面量（None/True/False），不能用 json.dumps
        col_map = self._build_col_map(template_struct)
        source_map = self._build_source_map_with_letters(source_struct)
        col_map_literal = _pprint.pformat(col_map, width=120, sort_dicts=False)
        source_map_literal = _pprint.pformat(source_map, width=120, sort_dicts=False)

        skeleton = f'''"""
DataMerge 自动生成 — 模板填充模式（AI 写公式 + 骨架预追加 `源_xxx` sheet）
- 加载训练时的模板（含公式/格式），骨架先把源数据追加为 `源_*` sheet，AI 在目标列写公式
- 不创建新 Workbook，不动单元格样式
"""
import os
import shutil
import sys
import pandas as pd
from pathlib import Path
# 顶层导入 openpyxl：智算走 importlib 执行（非沙箱），不会注入全局 openpyxl，
# AI 生成的 fill_template 直接引用 openpyxl/列字母工具时需要它在模块级可用，
# 否则会报 name 'openpyxl' is not defined（智训沙箱有注入故不报）。
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# 模板持久化路径（训练时确定，智算复用）
TEMPLATE_PATH = {tpl_repr}

# ==================== Excel 公式辅助常量（避免 f-string 引号嵌套 SyntaxError） ====================
# AI 写跨 sheet 公式时常需在 f-string 里嵌入 "" 作为 IFERROR 默认值，统一用 EMPTY 常量替代。
EMPTY = '""'   # 用法：f"=IFERROR(VLOOKUP(B{{r}},'源_考勤'!A:Z,5,FALSE),{{EMPTY}})"
SOURCE_PREFIX = "源_"
# ==================================================================================

# ==================== 列定位字典（训练时由 excel_parser 解析模板/源数据后固化） ====================
# 目标模板结构（按区域分组，多区域时每个 region 各有自己的 data_start_row）：
#   _COL_MAP[sheet_name] = {{"regions": [
#       {{"header_row": int, "data_start_row": int, "data_end_row": int|None,
#         "columns": [{{"letter": "A", "name": "工号", "has_formula": False}}, ...]}}, ...]}}
_COL_MAP = {col_map_literal}

# 源 sheet 结构（key 已带 "源_" 前缀，与骨架 _append_source_sheets 追加后的实际 sheet 名一致）：
#   骨架把源 DataFrame 重写为 row1=表头/row2+=数据，故 data_start_row 恒为 2。
#   _SOURCE_MAP["源_考勤明细"] = {{"header_row": 1, "data_start_row": 2,
#                                  "columns": [{{"letter": "A", "name": "工号"}}, ...]}}
_SOURCE_MAP = {source_map_literal}
# ==================== 列定位字典 结束 ====================

# 注入：薪资参数 / 历史数据上下文（沙箱外部 globals 注入）
salary_year = globals().get("salary_year", None)
salary_month = globals().get("salary_month", None)
monthly_standard_hours = globals().get("monthly_standard_hours", 174)
monthly_hours = monthly_standard_hours
output_folder = globals().get("output_folder", "")
input_folder = globals().get("input_folder", r"{input_folder}")


def load_source_data():
    """读取 input_folder 下所有 Excel 为 {{sheet_key: {{'df': DataFrame, 'columns': [...]}}}}

    优先复用沙箱注入的 _pre_loaded_source_data（智算时由 fast_header_matcher 预加载）。
    """
    pre = globals().get("_pre_loaded_source_data")
    if pre:
        print(f"使用预加载源数据：{{len(pre)}} 个 sheet")
        return pre

    out = {{}}
    for fname in os.listdir(input_folder):
        if not fname.lower().endswith((".xlsx", ".xls", ".xlsm")):
            continue
        if fname.startswith("~"):
            continue
        fp = os.path.join(input_folder, fname)
        try:
            xls = pd.ExcelFile(fp)
            for sn in xls.sheet_names:
                df = pd.read_excel(fp, sheet_name=sn)
                key = sn if sn not in out else f"{{Path(fname).stem}}_{{sn}}"
                out[key] = {{"df": df, "columns": list(df.columns)}}
        except Exception as e:
            print(f"[源数据加载警告] {{fname}}: {{e}}")
    print(f"加载完成：{{len(out)}} 个 sheet")
    return out


def _append_source_sheets(wb, source_data):
    """把 source_data 中的 DataFrame 追加为 wb 的新 sheet，sheet 名加 `源_` 前缀。

    - sheet 名 = SOURCE_PREFIX + source_data 的 key
    - 若与已存在 sheet 重名（极小概率），加 `_dup` 后缀；不会覆盖模板自带 sheet
    - 写表头 + 所有数据行；不动样式
    - 返回 {{原 source_key: 追加后的 sheet 名}}，便于 fill_template 用
    """
    appended_map = {{}}
    appended = []
    for sk, sv in (source_data or {{}}).items():
        df = (sv or {{}}).get("df")
        if df is None:
            continue
        target_name = f"{{SOURCE_PREFIX}}{{sk}}"[:31] if sk else f"{{SOURCE_PREFIX}}源数据"
        # 若与模板已有 sheet 同名，加后缀避免冲突
        suffix = 1
        base_name = target_name
        while target_name in wb.sheetnames:
            target_name = f"{{base_name[:28]}}_{{suffix}}"
            suffix += 1
        ws = wb.create_sheet(title=target_name)
        cols = list(df.columns)
        for ci, cname in enumerate(cols, start=1):
            ws.cell(row=1, column=ci).value = cname
        for ri, row in enumerate(df.itertuples(index=False, name=None), start=2):
            for ci, val in enumerate(row, start=1):
                if isinstance(val, float) and val != val:
                    val = None
                ws.cell(row=ri, column=ci).value = val
        appended_map[sk] = target_name
        appended.append(f"{{target_name}}({{len(df)}}行x{{len(cols)}}列)")
    if appended:
        print(f"[append_source_sheets] 已追加: {{', '.join(appended)}}")
    return appended_map


# ==================== AI 生成 ====================
{fill_function}
# ==================== AI 生成结束 ====================


def _snapshot_workbook(wb):
    snap = {{}}
    for ws in wb.worksheets:
        try:
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
        except Exception:
            max_row, max_col = 0, 0
        if max_row <= 0 or max_col <= 0:
            continue
        sheet_snap = {{}}
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                if cell.value is not None:
                    sheet_snap[(cell.row, cell.column)] = cell.value
        snap[ws.title] = sheet_snap
    return snap


def _snapshot_number_formats(wb):
    \"\"\"Snapshot every cell number_format (including empty cells), used to restore template formats after fill.\"\"\"
    snap = {{}}
    for ws in wb.worksheets:
        try:
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
        except Exception:
            max_row, max_col = 0, 0
        if max_row <= 0 or max_col <= 0:
            continue
        sheet_snap = {{}}
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                sheet_snap[(cell.row, cell.column)] = cell.number_format
        snap[ws.title] = sheet_snap
    return snap


def _restore_number_formats(wb, fmt_snapshot):
    \"\"\"Template-fill mode: strictly keep the template original cell formats.

    openpyxl auto-changes a General cell to a date format when a datetime value is
    written, turning the template's plain/numeric columns into date columns in the
    output. Restore each cell's number_format from the snapshot (date stays date,
    general stays general).\"\"\"
    restored = 0
    for ws in wb.worksheets:
        before = fmt_snapshot.get(ws.title)
        if not before:
            continue
        for (r, c), fmt in before.items():
            try:
                cell = ws.cell(row=r, column=c)
                if cell.number_format != fmt:
                    cell.number_format = fmt
                    restored += 1
            except Exception:
                continue
    if restored:
        print(f"  [fmt-guard] restored {{restored}} cells to template format")
    return restored


def _safe_repr(v):
    if v is None:
        return None
    if isinstance(v, (str, int, float, bool)):
        return v
    return str(v)[:200]


def _build_fill_report(wb, snapshot_before):
    from openpyxl.utils import get_column_letter

    filled_cells = []
    summary = {{}}

    for ws in wb.worksheets:
        before = snapshot_before.get(ws.title, {{}})
        try:
            max_row = ws.max_row or 0
            max_col = ws.max_column or 0
        except Exception:
            max_row, max_col = 0, 0
        if max_row <= 0 or max_col <= 0:
            continue
        for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                key = (cell.row, cell.column)
                old_v = before.get(key)
                new_v = cell.value
                if old_v == new_v:
                    continue
                col_letter = get_column_letter(cell.column)
                summary.setdefault(ws.title, {{}}).setdefault(col_letter, 0)
                summary[ws.title][col_letter] += 1
                if len(filled_cells) < 200:
                    filled_cells.append({{
                        "sheet": ws.title,
                        "address": f"{{col_letter}}{{cell.row}}",
                        "old": _safe_repr(old_v),
                        "new": _safe_repr(new_v),
                    }})
    return {{"filled_cells": filled_cells, "summary": summary,
            "total_changed": sum(sum(v.values()) for v in summary.values())}}


def main():
    print("=" * 60)
    print("模板填充模式 — 开始（AI 写公式 + 骨架追加 `源_xxx` sheet）")
    print("=" * 60)

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"模板文件丢失：{{TEMPLATE_PATH}}")

    # 有效模板：智算时若注入了 _template_override_path（用户上传的新模板）则优先用它；否则沿用训练时模板
    _override = globals().get("_template_override_path")
    _eff_template = _override if (_override and os.path.exists(_override)) else TEMPLATE_PATH
    if not os.path.exists(_eff_template):
        raise FileNotFoundError(f"模板文件丢失：{{_eff_template}}")
    if _override and _eff_template == _override:
        print(f"使用上传的新模板：{{_override}}")
    else:
        print(f"使用训练时模板：{{TEMPLATE_PATH}}")

    out_path = os.path.join(output_folder, Path(_eff_template).name)
    shutil.copy2(_eff_template, out_path)
    print(f"模板已复制至：{{out_path}}")

    source_data = load_source_data()

    import openpyxl
    # keep_vba 只对 .xlsm 启用：对普通 .xlsx 启用会让 openpyxl 把内容类型写成
    # macroEnabled(xlsm)，但扩展名仍是 .xlsx → Excel 报"文件损坏或扩展名无效"。
    _keep_vba = out_path.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(out_path, keep_vba=_keep_vba, data_only=False)

    # 结构提示（不拦截）：新模板缺少训练时的目标 sheet 时，相关公式会落空
    _missing = [sn for sn in _COL_MAP.keys() if sn not in wb.sheetnames]
    if _missing:
        print(f"[模板结构警告] 当前模板缺少训练时的目标 sheet: {{_missing}}，相关公式可能落空")

    print("步骤：追加源数据 sheet（前缀 `源_`，供 AI 公式引用）...")
    try:
        _append_source_sheets(wb, source_data)
    except Exception as _ap_e:
        print(f"[append_source_sheets] 异常（不阻断）：{{_ap_e}}")

    snapshot_before = _snapshot_workbook(wb)
    _fmt_snapshot = _snapshot_number_formats(wb)

    print("步骤：调用 AI 生成的 fill_template 写入规则要求的列...")
    fill_template(wb, source_data, salary_year, salary_month, monthly_standard_hours)

    # Template-fill mode: restore template original cell formats
    # (prevent openpyxl auto-changing General to date format when datetime is written)
    try:
        _restore_number_formats(wb, _fmt_snapshot)
    except Exception as _fe:
        print(f"[fmt-guard] restore format error (ignored): {{_fe}}")

    try:
        import json as _json
        from datetime import datetime as _dt
        report = _build_fill_report(wb, snapshot_before)
        report.update({{
            "mode": "template",
            "timestamp": _dt.now().isoformat(timespec="seconds"),
            "template": Path(TEMPLATE_PATH).name,
        }})
        report_path = os.path.join(output_folder, "fill_report.json")
        with open(report_path, "w", encoding="utf-8") as fp:
            _json.dump(report, fp, ensure_ascii=False, indent=2)
        print(f"填充报告：{{report_path}}（共改动 {{report['total_changed']}} 个 cell）")
        for sn, cols in report["summary"].items():
            cols_brief = ", ".join(f"{{c}}({{n}})" for c, n in cols.items())
            print(f"  · {{sn}}: {{cols_brief}}")
    except Exception as _rep_e:
        print(f"[填充报告] 生成失败（不阻断保存）: {{_rep_e}}")

    wb.save(out_path)
    print(f"保存成功：{{out_path}}")
    print("=" * 60)
    print("处理完成!")
    print("=" * 60)
    return True


if __name__ == "__main__":
    main()
'''
        return skeleton

    # ==================== 差异修正（chat 修正流程） ====================

    # 骨架里包裹 fill_template 的边界标记（与 _build_complete_code 中保持一致）
    _AI_BLOCK_BEGIN = "# ==================== AI 生成 ===================="
    _AI_BLOCK_END = "# ==================== AI 生成结束 ===================="

    def _extract_existing_fill_block(self, original_code: str) -> Tuple[str, Optional[Tuple[int, int]]]:
        """从已有完整脚本中抽取 fill_template 函数代码，并返回边界标记的位置区间。

        优先策略：用骨架的边界标记切片（最稳定）；找不到时回退正则按 def fill_template 抽函数。
        Returns:
            (fill_function_code, (begin_idx, end_idx)) — 区间为 begin_marker 之后到 end_marker 之前的字符索引；找不到时位置返回 None。
        """
        if not original_code:
            return "", None
        b = original_code.find(self._AI_BLOCK_BEGIN)
        e = original_code.find(self._AI_BLOCK_END, b + 1) if b >= 0 else -1
        if b >= 0 and e > b:
            inner_start = b + len(self._AI_BLOCK_BEGIN)
            inner = original_code[inner_start:e].strip("\n")
            return inner, (inner_start, e)

        m = re.search(r"(def\s+fill_template\s*\([^)]*\)\s*:[\s\S]+?)(?=\n[ \t]*def\s+\w|\nif\s+__name__\s*==)", original_code)
        if m:
            return m.group(1).rstrip(), None
        return "", None

    def _replace_fill_block(self, original_code: str, span: Optional[Tuple[int, int]], new_fill_function: str) -> str:
        """把新的 fill_template 写回原脚本：优先按 span 替换；否则按 def fill_template 整段替换。"""
        new_block = f"\n{new_fill_function.strip()}\n"
        if span is not None:
            inner_start, inner_end = span
            return original_code[:inner_start] + new_block + original_code[inner_end:]
        # 兜底：把 def fill_template 整段函数替换掉
        pattern = re.compile(
            r"(def\s+fill_template\s*\([^)]*\)\s*:[\s\S]+?)(?=\n[ \t]*def\s+\w|\nif\s+__name__\s*==)"
        )
        if pattern.search(original_code):
            return pattern.sub(new_block.strip() + "\n\n", original_code, count=1)
        # 实在没法替换：在 main() 之前追加（最坏兜底）
        return original_code + "\n\n" + new_block

    def _extract_map_block(self, code: str, var_name: str) -> str:
        """从生成代码中抽出 _COL_MAP / _SOURCE_MAP 的字面量字符串（仅作 prompt 上下文展示，不参与执行）。

        通过括号配对从 `var_name = {` 切到对应的闭合 `}`，避免 pprint 字面量含 `'有些 sheet'` 把正则吃坏。
        """
        if not code:
            return ""
        idx = code.find(f"{var_name} = {{")
        if idx < 0:
            idx = code.find(f"{var_name}={{")
        if idx < 0:
            return ""
        brace_start = code.find("{", idx)
        depth = 0
        in_str = False
        quote = ""
        i = brace_start
        while i < len(code):
            ch = code[i]
            if in_str:
                if ch == "\\":
                    i += 2
                    continue
                if ch == quote:
                    in_str = False
            else:
                if ch in ("'", '"'):
                    in_str = True
                    quote = ch
                elif ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        return code[brace_start:i + 1]
            i += 1
        return ""

    def generate_correction_code(
        self,
        original_code: str,
        comparison_result: str,
        rules_content: str,
        source_structure: str = "",
        stream_callback: callable = None,
        iteration_num: int = 1,
    ) -> str:
        """生成模板模式的修正代码 — 仅改写 fill_template 函数，骨架原样保留

        Args:
            original_code: 上一轮训练完整 Python 脚本（含骨架 + fill_template + _COL_MAP/_SOURCE_MAP）
            comparison_result: diff 文本 + 用户反馈拼接（由 training_chat 的修正分支构建）
            rules_content: 规则文档（核心输入）
            source_structure: 源数据结构描述（可选；模板模式主要靠 _SOURCE_MAP，留空也能跑）
            stream_callback: 流式回调（前端实时显示）
            iteration_num: 当前修正轮次（>=2 时启用"换思路"指令防止死循环）

        Returns:
            完整的修正脚本（直接可写入沙箱执行）；失败时返回 original_code 兜底
        """
        def log(msg):
            logger.info(msg)
            if stream_callback:
                stream_callback(msg)

        log("=== 模板填充模式：开始差异修正 ===")

        if not original_code:
            log("[修正失败] 原始代码为空")
            return None

        # 1. 抽出旧的 fill_template
        prev_fill, span = self._extract_existing_fill_block(original_code)
        if not prev_fill or "def fill_template" not in prev_fill:
            log("[修正失败] 无法从原始代码中抽取 fill_template 函数")
            return None
        log(f"已抽取上一轮 fill_template（{len(prev_fill)} 字符）")

        # 2. 抽出 _COL_MAP / _SOURCE_MAP 给 AI 看（让它知道有哪些列/sheet 可用）
        col_map_str = self._extract_map_block(original_code, "_COL_MAP")
        source_map_str = self._extract_map_block(original_code, "_SOURCE_MAP")

        # 3. 第 2 轮起加"换思路"指令
        is_retry = iteration_num >= 2
        if is_retry:
            prev_label = f"## 你上一轮（第 {iteration_num - 1} 轮）生成的 fill_template 函数"
            retry_directive = f"""
⚠️⚠️⚠️ **本次为第 {iteration_num} 轮修正，请严格遵守以下"换思路"原则：**

1. 上方代码 **就是你自己上一轮（第 {iteration_num - 1} 轮）生成并已验证为不正确的版本**——本次差异与上一轮高度相似（即上次的修正没起作用），说明上次的修改方向是错误的
2. **禁止重复同一种修正套路**：如果上次改的是 VLOOKUP 的列号，本次请考虑换不同的源 sheet/源列/换不同的公式（VLOOKUP→INDEX/MATCH→SUMIFS）
3. **对每一个 diff 中的字段**，先在心里回答："上一版的代码是用什么思路写的？为什么这个思路得不到正确结果？换什么思路能避开这个问题？"——再动笔改
"""
        else:
            prev_label = "## 上一轮 fill_template 函数（你需要在此基础上修改）"
            retry_directive = ""

        col_map_block = f"## 模板 _COL_MAP（运行时已注入，可直接引用）\n```python\n_COL_MAP = {col_map_str}\n```\n" if col_map_str else ""
        source_map_block = f"## 源数据 _SOURCE_MAP（运行时已注入，可直接引用）\n```python\n_SOURCE_MAP = {source_map_str}\n```\n" if source_map_str else ""

        prompt = f"""你是 Excel 数据处理专家。任务：**仅修改** `fill_template` 函数，根据用户反馈和差异修正代码。

## 严格修正原则（违反必败）
1. **只改用户反馈中提到的列**：未提到的列代码必须**原封不动**保留，一个字符都不能改
2. **保留所有变量定义、import、辅助函数**
3. **完整输出 fill_template 函数**（包括所有列的处理，未改动的列也要 1:1 保留原代码）
4. **【起始行强约束】数据写入起始行必须用 `_COL_MAP[sheet]["data_start_row"]`，绝不硬编码 1/2/3**
   - 同样：读源 sheet 必须用 `_SOURCE_MAP[源_xxx]["data_start_row"]`
5. **公式查找数据源分两类**：骨架追加的源文件用 `_SOURCE_MAP` 的 `源_xxx` key；模板自带 sheet（在 `_COL_MAP` 里）用其**真实 sheet 名、不加 `源_` 前缀**。sheet 名外用 `'` 单引号
6. **跳过模板里 `has_formula=True` 的列**（保留原公式）
7. **【f-string 引号强约束】**：外层 f-string 用 **双引号** `f"..."`，公式中 sheet 名用 `'` 单引号；空字符串拼接 `EMPTY` 常量（=`'""'`），写法：`f"=IFERROR(VLOOKUP(B{{r}},'源_考勤'!A:Z,3,FALSE),{{EMPTY}})"`
   - **错误**：`f'=IFERROR(VLOOKUP(...,'源_考勤'!A:Z,...),"")'`  ← 单引号嵌套立即 SyntaxError

{prev_label}
```python
{prev_fill}
```

## 与预期结果的差异 + 用户修正指示
{comparison_result[:30000]}
{retry_directive}
{col_map_block}{source_map_block}## 计算规则（参考）
{(rules_content or '')[:30000]}

## 输出要求
**只输出修正后的 `def fill_template(...)` 完整函数**，用 ```python 代码块包裹。
- 不要 import 模块（除从 openpyxl.utils 临时导入）
- 不要 main、不要解释文字、不要 Workbook() 创建新 wb
- 未在用户反馈中提到的列，其代码必须与上一轮完全一致
"""

        if self.training_logger and hasattr(self.training_logger, "log_full_prompt"):
            try:
                self.training_logger.log_full_prompt(prompt, "correct")
            except Exception:
                pass

        # 4. 调用 AI（流式）
        ai_response = self._call_ai(prompt, stream_callback)
        if self.training_logger and hasattr(self.training_logger, "log_full_ai_response"):
            try:
                self.training_logger.log_full_ai_response(ai_response, "correct")
            except Exception:
                pass

        # 5. 抽取新的 fill_template
        new_fill = self._extract_fill_function(ai_response)
        if not new_fill or "def fill_template" not in new_fill:
            log("[修正失败] AI 响应中未找到 fill_template，回退原代码")
            return original_code

        # 6. 应用与 generate_code 同款的修复管线
        try:
            new_fill = self._indent_fixer.fix(new_fill)
        except Exception as e:
            logger.warning(f"缩进修复异常（不阻断）: {e}")

        try:
            if hasattr(self.ai_provider, "validate_and_fix_code_format"):
                fixed = self.ai_provider.validate_and_fix_code_format(new_fill)
                if fixed and isinstance(fixed, str):
                    new_fill = fixed
        except Exception as e:
            logger.warning(f"ai_provider 修复管线异常（不阻断）: {e}")

        # 7. 拼回完整脚本
        complete_code = self._replace_fill_block(original_code, span, new_fill)

        # 8. 完整脚本再过一遍 ai_provider 修复（防 f-string 残留）
        try:
            if hasattr(self.ai_provider, "_fix_fstring_quotes"):
                _full_fixed = self.ai_provider._fix_fstring_quotes(complete_code)
                if _full_fixed and isinstance(_full_fixed, str):
                    complete_code = _full_fixed
        except Exception as e:
            logger.warning(f"完整脚本 f-string 二次修复异常（不阻断）: {e}")

        # 9. 语法校验
        try:
            self._validate_syntax(complete_code, log)
        except RuntimeError as se:
            log(f"[修正失败] 语法校验未通过：{se}")
            return None

        log(f"=== 模板填充模式：差异修正完成（complete_code={len(complete_code)} 字符） ===")
        return complete_code
