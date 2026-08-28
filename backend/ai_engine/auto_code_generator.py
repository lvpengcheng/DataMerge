"""
自动模式代码生成器 - AI 跳出目标文件束缚，自由设计输出结构

核心思路：
1. AI 完全按规则文档/对话提示词生成 Excel，不必对齐目标文件结构
2. 用 pandas 在内存中计算所有值，结果写入新 Workbook
3. 不写源数据 sheet 到结果文件（解决公式模式产物体积大、Excel 重算慢的问题）
4. 目标文件可选：提供则作为列结构软参考；不提供 AI 完全自主设计

适用场景：
- 用户有明确算法/规则但无固定模板
- 需要灵活报表（汇总、透视、自定义 sheet 结构）
- 希望产物文件小、打开快（值已固化）

vs 公式模式：产物不含源 sheet，无 Excel 实时公式（值已计算完）
vs 模板模式：从零生成 wb，不需要预存模板
"""

import os
import re
import logging
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path

from .ai_provider import BaseAIProvider, AIProviderFactory
from backend.utils.indentation_fixer import IndentationFixer

logger = logging.getLogger(__name__)


class AutoCodeGenerator:
    """自动模式代码生成器：AI 自由设计 + 纯计算输出"""

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
        expected_structure: Optional[Dict[str, Any]] = None,
        manual_headers: Dict = None,
        stream_callback: callable = None,
        thinking_callback: callable = None,
        multi_sheet_source: bool = False,
        use_history: bool = False,
        target_sheets: Optional[List[str]] = None,
    ) -> Tuple[str, str]:
        """生成自动模式 Python 代码

        Args:
            input_folder: 源数据目录
            rules_content: 规则文本（核心输入，AI 主要靠这个）
            expected_structure: 目标文件结构（可选，软参考）
            manual_headers: 源数据手动表头
            stream_callback: 流式日志回调
            thinking_callback: 推理模型思考过程逐块回调
            multi_sheet_source: 源是否多 sheet
            use_history: 是否启用历史数据
            target_sheets: 用户勾选的目标 sheet 名列表；过滤 expected_structure 软参考

        Returns:
            (完整脚本, AI 原始响应)
        """
        def log(msg):
            logger.info(msg)
            if stream_callback:
                stream_callback(msg)

        log("=== 自动模式：开始生成代码 ===")

        # 0. 按 target_sheets 过滤 expected_structure（用户勾选时仅保留这些 sheet 作为软参考）
        if target_sheets and expected_structure and isinstance(expected_structure, dict):
            _all = expected_structure.get("sheets") or {}
            if _all:
                _kept = {sn: info for sn, info in _all.items() if sn in target_sheets}
                _dropped = [sn for sn in _all.keys() if sn not in target_sheets]
                expected_structure = dict(expected_structure)
                expected_structure["sheets"] = _kept
                expected_structure["target_sheets"] = list(target_sheets)
                log(f"按用户勾选过滤 expected_structure：保留 {len(_kept)}（丢弃 {_dropped[:5]}{'...' if len(_dropped) > 5 else ''}）")

        # 1. 解析源数据列结构
        log("步骤1: 解析源数据列...")
        source_struct = self._parse_source(input_folder, manual_headers, multi_sheet_source)
        log(f"源数据含 {len(source_struct.get('files', {}))} 个文件")

        # 2. 构造 prompt（expected_structure 仅作软参考）
        log("步骤2: 构造 AI prompt...")
        prompt = self._build_prompt(rules_content, source_struct, expected_structure, use_history)
        self.last_prompt = prompt

        # 3. 调用 AI
        log("步骤3: 调用 AI 生成构建函数...")
        ai_response = self._call_ai(prompt, stream_callback, thinking_callback)

        # 4. 提取函数代码
        build_function = self._extract_build_function(ai_response)
        if not build_function:
            preview = (ai_response or "")[:2000]
            log(f"[提取失败] AI 响应预览（前2000字符）：\n{preview}")
            logger.error(f"无法提取 build_result_workbook，AI 响应长度={len(ai_response or '')}，预览：{preview[:500]}")
            raise RuntimeError("AI 响应中无法提取 build_result_workbook 函数（已把响应预览写入训练日志）")

        # 5. 拼接完整脚本
        log("步骤4: 拼接完整脚本...")
        complete_code = self._build_complete_code(
            build_function=build_function,
            input_folder=input_folder,
            use_history=use_history,
        )

        # 6. 缩进修复
        try:
            complete_code = self._indent_fixer.fix(complete_code)
        except Exception as e:
            logger.warning(f"缩进修复异常（不阻断）: {e}")

        log("=== 自动模式：代码生成完成 ===")
        return complete_code, ai_response

    # ==================== 源数据解析 ====================

    def _parse_source(self, input_folder: str, manual_headers: Optional[Dict], multi_sheet_source: bool) -> Dict[str, Any]:
        """轻量解析源数据，让 AI 知道有哪些列可用"""
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
            if fname.startswith("~"):
                continue
            try:
                sheets_data = parser.parse_excel_file(
                    fp, max_data_rows=2, read_formulas=False,
                    active_sheet_only=not multi_sheet_source,
                    best_region_only=True,   # 取每 sheet 最优(主数据)区域，避免 banner 关闭后多区域列名混在一起
                    manual_headers=manual_headers,
                )
                fdesc = {}
                for sd in sheets_data:
                    cols = []
                    for region in (sd.regions or []):
                        head = region.head_data or {}
                        cols.extend(list(head.keys()))
                    fdesc[sd.sheet_name] = cols
                out["files"][fname] = fdesc
            except Exception as e:
                logger.warning(f"源文件 {fname} 解析失败: {e}")
                out["files"][fname] = {"error": str(e)}
        return out

    # ==================== AI Prompt ====================

    def _build_prompt(
        self,
        rules_content: str,
        source_struct: Dict[str, Any],
        expected_structure: Optional[Dict[str, Any]],
        use_history: bool,
    ) -> str:
        import json
        rules_short = (rules_content or "")[:30000]

        exp_block = ""
        if expected_structure and isinstance(expected_structure, dict) and expected_structure.get("sheets"):
            exp_block = (
                "## 目标结构（软参考，仅供参考列名/sheet 命名习惯，**不要求严格对齐**）\n"
                f"```json\n{json.dumps(expected_structure, ensure_ascii=False)[:3000]}\n```\n\n"
            )

        return f"""你是 Excel 数据处理 + Python pandas 专家。任务：编写函数 `build_result_workbook`，按规则文档自由设计输出 Excel。

## 严格约束（违反必败）
1. 所有计算用 pandas 在内存中完成；**绝不**把源数据 DataFrame 整体写入结果 sheet（产物文件要小、打开要快）
2. 输出 cell 写入计算后的**值**（数字、字符串、日期），不写 Excel 公式（VLOOKUP/INDEX 等不要用）
   - 例外：简单的 SUM/AVERAGE 跨 sheet 汇总公式可以用，单元格内有自计算需求时
3. 自由设计 sheet 名称、列顺序、汇总行/列；规则里没要求的格式（颜色/字体/边框）**不要**乱加
4. **不要** wb.save（外层会调）；**不要** import 模块（外层已 import）
5. 用 source_data dict 拿数据：`source_data["sheet_key"]["df"]` 是 pandas DataFrame

## 函数签名
```python
def build_result_workbook(source_data, salary_year, salary_month, monthly_hours):
    \"\"\"
    Args:
        source_data: dict[sheet_key -> {{'df': DataFrame, 'columns': [...]}}]
        salary_year/salary_month/monthly_hours: 薪资上下文参数
    Returns:
        openpyxl Workbook 对象
    \"\"\"
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)  # 移除默认空 sheet

    # 你的代码：
    # 1) 用 pandas 算出每个目标 sheet 的 DataFrame
    # 2) wb.create_sheet(...)，把表头和值写入
    # 3) 返回 wb
    ...
    return wb
```

## 源数据结构
```json
{json.dumps(source_struct, ensure_ascii=False, indent=2)[:5000]}
```

{exp_block}## 规则文档（这是核心输入）
{rules_short}

## 输出要求
**只输出 `def build_result_workbook(...)` 函数体**，用 ```python 代码块包裹。不要 import、不要 main、不要解释文字。
"""

    # ==================== AI 调用 ====================

    def _call_ai(self, prompt: str, stream_callback=None, thinking_callback=None) -> str:
        messages = [
            {"role": "system", "content": "你是 Python + pandas + openpyxl 专家，擅长 HR/薪酬数据处理。严格按用户给定的输出格式回答。"},
            {"role": "user", "content": prompt},
        ]
        # 输出预算用 provider 配置：deepseek 推理模型的 max_tokens 是【含思考的总预算】，
        # 写死 8000 会被思考阶段吃掉大半、代码截断（Claude 思考不占输出预算所以没暴露）
        _effective_max_tokens = getattr(self.ai_provider, "max_tokens", None) or 8000
        try:
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
                    thinking_callback=thinking_callback,
                    temperature=0.1,
                    max_tokens=_effective_max_tokens,
                )
            else:
                resp = self.ai_provider.chat(messages, temperature=0.1, max_tokens=_effective_max_tokens)
        except Exception as e:
            logger.error(f"AI 调用失败: {e}", exc_info=True)
            raise
        return resp or ""

    # ==================== 代码提取 ====================

    def _extract_build_function(self, ai_response: str) -> str:
        """从 AI 响应中抽取 def build_result_workbook 函数

        鲁棒策略：
        1. 扫所有 ```...``` 代码块，取第一个含 `def build_result_workbook` 的。
        2. 无代码块时回退整段文本。
        3. 全部失败返回空字符串。
        """
        if not ai_response:
            return ""
        text = ai_response.strip()

        blocks = re.findall(r"```(?:[a-zA-Z0-9_+-]*)\s*(.*?)\s*```", text, re.DOTALL)
        for blk in blocks:
            if "def build_result_workbook" in blk:
                return blk.strip()

        if "def build_result_workbook" in text:
            cleaned = re.sub(r"^```[a-zA-Z0-9_+-]*\s*|\s*```\s*$", "", text, flags=re.MULTILINE)
            return cleaned.strip()

        return ""

    # ==================== 完整脚本拼接 ====================

    def _build_complete_code(
        self,
        build_function: str,
        input_folder: str,
        use_history: bool,
    ) -> str:
        """把 AI 生成的 build_result_workbook 包到固定的脚本骨架里"""
        skeleton = f'''"""
DataMerge 自动生成 — 自动模式（AI 自由设计 + 纯计算输出）
- 不写源数据 sheet 到结果，产物文件小、Excel 打开快
- 所有计算在 pandas 内存中完成，cell 存死值
"""
import os
import sys
import pandas as pd
from pathlib import Path
from openpyxl import Workbook

# 注入：薪资参数 / 上下文（沙箱外部 globals 注入）
salary_year = globals().get("salary_year", None)
salary_month = globals().get("salary_month", None)
monthly_standard_hours = globals().get("monthly_standard_hours", 174)
output_folder = globals().get("output_folder", "")
input_folder = globals().get("input_folder", r"{input_folder}")


# ---- 键列强制文本（与 output_postprocess 主键归一同规则）----
# 避免 pd.read_excel / 预加载把数字型工号、证件号读成 int，导致内存 join 两端类型不一致匹配失败
_KEY_KW = ("工号", "员工工号", "员工编号", "员工号", "职工号", "人员编号", "工作证号", "员工id",
           "身份证", "证件号", "身份证号", "银行卡", "卡号", "银行账号",
           "社保号", "社保账号", "公积金号", "公积金账号",
           "手机号", "联系电话", "税号", "纳税人识别号", "编号")
_KEY_EX = ("工资", "金额", "薪资", "比例", "系数", "天数", "月数", "说明", "规则", "姓名")


def _is_key_col(name):
    n = str(name or "").strip().lower()
    if not n or any(e in n for e in _KEY_EX):
        return False
    return any(k in n for k in _KEY_KW)


def _canon_key(v):
    import math
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return v
        return str(int(v)) if v == int(v) else repr(v)
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


def _normalize_key_columns(df):
    """把命中主键/ID 关键词的列统一成规范文本（110.0→"110"），与输出端主键归一同规则。"""
    try:
        for _c in list(df.columns):
            if _is_key_col(_c):
                df[_c] = df[_c].map(_canon_key)
    except Exception as _e:
        print(f"[键列归一警告] {{_e}}")
    return df


def load_source_data():
    """读取 input_folder 下所有 Excel 为 {{sheet_key: {{'df': DataFrame, 'columns': [...]}}}}

    优先复用沙箱注入的 _pre_loaded_source_data（智算时由 fast_header_matcher 预加载）。
    """
    pre = globals().get("_pre_loaded_source_data")
    if pre:
        print(f"使用预加载源数据：{{len(pre)}} 个 sheet")
        for _v in pre.values():
            if isinstance(_v, dict) and _v.get("df") is not None:
                _normalize_key_columns(_v["df"])
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
                # dtype=object is required even when the Excel cell itself is text:
                # pandas otherwise infers an all-digit text column as float64. 18-digit
                # identifiers then exceed IEEE-754 precision and their trailing digits
                # are already corrupted before key-column normalization can stringify them.
                df = pd.read_excel(fp, sheet_name=sn, dtype=object)
                _normalize_key_columns(df)
                key = sn if sn not in out else f"{{Path(fname).stem}}_{{sn}}"
                out[key] = {{"df": df, "columns": list(df.columns)}}
        except Exception as e:
            print(f"[源数据加载警告] {{fname}}: {{e}}")
    print(f"加载完成：{{len(out)}} 个 sheet")
    return out


# ==================== AI 生成 ====================
{build_function}
# ==================== AI 生成结束 ====================


def main():
    print("=" * 60)
    print("自动模式 — 开始")
    print("=" * 60)

    # 1. 加载源数据
    source_data = load_source_data()

    # 2. AI 生成的构建函数生成 wb
    print("步骤：调用 build_result_workbook ...")
    wb = build_result_workbook(source_data, salary_year, salary_month, monthly_standard_hours)

    # 3. 保存到 output_folder
    out_path = os.path.join(output_folder, "薪资汇总表.xlsx")
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
