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
6. **ROUND操作**: 仅规则明确要求时才添加"""

    def __init__(self):
        self.templates = self._load_templates()
        self.max_structure_length = 20000000
        self.logger = logging.getLogger(__name__)
        self.rule_extractor = RuleExtractor()

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

    def _compress_rules(self, text: str, max_length: int = 20000) -> str:
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
{rules_content[:10000]}

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
        rules = self._compress_rules(rules_content, max_length=15000)
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
        rules = self._compress_rules(rules_content, max_length=20000)
        _ = manual_headers

        # 提取数据清洗规则、警告规则和条件格式规则
        extracted_rules = self.rule_extractor.extract_rules(rules_content)
        data_cleaning_rules_text = ""

        if extracted_rules["data_cleaning_rules"]:
            data_cleaning_rules_text = "\n## 数据清洗规则（在write_source_sheets中应用）\n"
            data_cleaning_rules_text += "⚠️ 在将源数据写入Excel之前，必须先应用以下清洗规则过滤数据：\n\n"
            for i, rule in enumerate(extracted_rules["data_cleaning_rules"], 1):
                if rule['original_text'] and not rule['original_text'].startswith('--'):
                    data_cleaning_rules_text += f"{i}. {rule['original_text']}\n"
            data_cleaning_rules_text += "\n**实现方式**：在write_source_sheets函数中，对每个DataFrame应用清洗逻辑后再写入Excel。\n"

        conditional_format_text = ""
        if extracted_rules["conditional_format_rules"]:
            conditional_format_text = "\n## 条件格式规则\n"
            conditional_format_text += "在填充公式后，需要对以下情况应用条件格式（使用CellIsRule/FormulaRule）：\n\n"
            for i, rule in enumerate(extracted_rules["conditional_format_rules"], 1):
                conditional_format_text += f"{i}. {rule['original_text']}\n"

        precision_rules_text = ""
        if extracted_rules["precision_rules"]:
            precision_rules_text = "\n## 数值精度规则\n"
            precision_rules_text += "在生成公式时，必须按以下精度要求使用ROUND函数：\n\n"
            for i, rule in enumerate(extracted_rules["precision_rules"], 1):
                precision_rules_text += f"{i}. {rule['original_text']}\n"

        # 统计总列数
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

__SOURCE_STRUCTURE__
__EXPECTED_SHEETS_INFO__

## 预期输出结构
__COMPRESSED_EXPECTED__

__DATA_CLEANING_RULES__

__CONDITIONAL_FORMAT_RULES__

__PRECISION_RULES__

## 计算规则
__RULES__

# ==================== 黄金样例（必须严格参照此模式编写）====================
以下是一个完整的fill_result_sheets代码示例，展示了5种常见列类型的正确写法。
你的代码必须完全遵循此模式，包括变量定义位置、缩进层级、注释格式、f-string写法。

```python
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    # === 1. 主表选择（员工数最多的表）===
    main_key = max(source_sheets.keys(), key=lambda k: len(source_sheets[k]['df']))
    main_df = source_sheets[main_key]['df']
    n_rows = len(main_df)

    # === 2. 源表sheet名变量（在for循环外定义，VLOOKUP公式中使用）===
    sn_main = source_sheets[main_key]['ws'].title
    sn_attend = source_sheets['考勤表_Sheet1']['ws'].title       # 示例：考勤表
    sn_memo = source_sheets['薪资备忘录_Sheet1']['ws'].title     # 示例：备忘录

    # === 3. 预计算VLOOKUP列号（在for循环外，用get_vlookup_col_num）===
    col_dept = get_vlookup_col_num("D", "A")         # 部门在D列，主键在A列
    col_attend_days = get_vlookup_col_num("AM", "F") # 出勤天数在AM列，主键在F列

    # === 4. 创建结果sheet并写表头 ===
    ws = wb.create_sheet("结果")
    headers = ["工号", "姓名", "部门", "出勤天数", "日薪", "应发工资"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    # === 5. 逐行填充（每列平级排列，缩进统一8空格）===
    for i in range(n_rows):
        r = i + 2

        # A列(1): 工号 - 主表直接复制
        ws.cell(row=r, column=1, value=main_df.iloc[i].get('工号', ''))

        # B列(2): 姓名 - 主表直接复制
        ws.cell(row=r, column=2, value=main_df.iloc[i].get('姓名', ''))

        # C列(3): 部门 - VLOOKUP跨表取数（引用花名册）
        ws.cell(row=r, column=3).value = f"=IFERROR(VLOOKUP(A{r},\'{sn_memo}\'!$A:$J,{col_dept},FALSE),{EMPTY})"

        # D列(4): 出勤天数 - VLOOKUP跨表取数（主键非A列，注意范围起始列）
        ws.cell(row=r, column=4).value = f"=IFERROR(VLOOKUP(A{r},\'{sn_attend}\'!$F:$BG,{col_attend_days},FALSE),0)"

        # E列(5): 日薪 - 公式计算（引用参数sheet）
        ws.cell(row=r, column=5).value = f"=IFERROR(F{r}/参数!$B$4,0)"

        # F列(6): 应发工资 - 含TEXT函数时用单引号f-string
        ws.cell(row=r, column=6).value = f'=IF(D{r}>0,D{r}*E{r},0)'

    # === 6. 条件格式（如规则要求标红/高亮等，在循环结束后实现）===
    # 出勤天数>20标红
    if n_rows > 0:
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        ws.conditional_formatting.add(f"D2:D{n_rows+1}", CellIsRule(operator="greaterThan", formula=["20"], fill=red_fill))

    # 共处理了 6 列（预期 6 列）
```

### 从黄金样例中提取的关键模式
1. **变量定义位置**：sn_xxx 和 col_xxx 全部在for循环外定义
2. **注释格式**：每列仅一行 `# X列(N): 说明`，无分隔线
3. **列平级排列**：每列代码缩进统一8空格，互不嵌套
4. **f-string规则**：
   - 公式含sheet名（单引号）→ 外层双引号：`f"=VLOOKUP(A{r},\'{sn_xxx}\'!...)"`
   - 公式含双引号（TEXT等）→ 外层单引号：`f'=TEXT(A{r},"YYYY-MM-DD")'`
   - Excel空字符串 → 用{EMPTY}代替
   - Excel文本比较 → 用{excel_text('xxx')}代替
5. **VLOOKUP列号**：用get_vlookup_col_num()计算，禁止硬编码数字
6. **主表数据**：直接用main_df.iloc[i].get()复制，不用VLOOKUP
7. **日期计算**：参与运算前用DATEVALUE()转换
8. **条件格式**：规则要求标红/高亮时，在for循环结束后用CellIsRule/FormulaRule实现，不能只留注释

# ==================== 补充规则 ====================

## 历史数据引用（仅规则涉及"累计""历史"时使用）
- 系统自动创建"历史数据"sheet，第1列为薪资月份，其余列与结果sheet列名一致
- 累计示例：`=SUMIFS(历史数据!$E:$E, 历史数据!$A:$A, "<"&salary_month, 历史数据!$B:$B, B{{r}})`
- 需用IFERROR包裹（第1个月历史数据可能为空）
- 规则中未涉及历史数据时不要引用此sheet

## 工号类型
- 一般为数字格式，不需要TEXT转换
- 仅当工号包含字母或特殊字符时才用TEXT转换

## 完整性要求
- 必须为全部 __TOTAL_COLUMNS__ 列生成处理逻辑
- 在代码最后添加注释：# 共处理了 X 列（预期 __TOTAL_COLUMNS__ 列）
- 每行代码必须完整闭合，不允许跨行写赋值语句
- 禁止"暂时跳过"、"简化为0"、提前结束
- 禁止在函数内部import已导入的模块
- 全部使用英文半角字符 ()[]"'，禁止中文全角

# ==================== 函数签名 ====================

## 1. 数据清洗函数（如果有清洗规则）
def clean_source_data(source_data):
    \"\"\"应用数据清洗规则过滤源数据

    Args:
        source_data: {"文件名_sheet名": {"df": DataFrame, "columns": [列名]}}

    Returns:
        清洗后的source_data（相同格式）
    \"\"\"
    pass

## 2. 结果填充函数
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    \"\"\"创建结果sheet并填充数据和公式

    Args:
        wb: openpyxl Workbook
        source_sheets: {"文件名_sheet名": {"df": DataFrame, "ws": worksheet}}
        salary_year: 薪资年份
        salary_month: 薪资月份
        monthly_standard_hours: 月标准工时
    \"\"\"
    pass

请严格参照黄金样例的代码模式，生成完整代码：
1. 如果有数据清洗规则，先生成clean_source_data函数
2. 然后生成fill_result_sheets函数，覆盖全部 __TOTAL_COLUMNS__ 列"""

        return (template
                .replace('__TOTAL_COLUMNS__', str(total_columns))
                .replace('__SOURCE_STRUCTURE__', source_structure)
                .replace('__EXPECTED_SHEETS_INFO__', expected_sheets_info)
                .replace('__COMPRESSED_EXPECTED__', compressed_expected)
                .replace('__DATA_CLEANING_RULES__', data_cleaning_rules_text)
                .replace('__CONDITIONAL_FORMAT_RULES__', conditional_format_text)
                .replace('__PRECISION_RULES__', precision_rules_text)
                .replace('__RULES__', rules))

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

# ==================== 核心规则（违反=失败）====================
## 【规则1】字符规范
- ✅ 必须：英文半角 ()[]"'
- ❌ 禁止：中文全角 （）【】""''

## 【规则2】跨表取数方式（按优先级，禁止直接引用DataFrame）
- 主表数据：直接复制,不能用VLOOKUP
- 非主表数据：必须使用Excel公式跨表取数，根据场景选择：
  - VLOOKUP：单条件精确匹配（默认首选），列号 = 目标列位置 - 范围起始列位置 + 1
  - XLOOKUP：反向查找、自定义默认值，格式：`=XLOOKUP(查找值,'表'!查找列,'表'!返回列,默认值,0)`
  - INDEX+MATCH：多条件匹配、左向查找，格式：`=IFERROR(INDEX('表'!返回列,MATCH(查找值,'表'!查找列,0)),0)`
  - FILTER：一对多匹配+聚合，格式：`=SUM(FILTER('表'!金额列,'表'!工号列=A2))`
  - SUMPRODUCT：多条件求和/计数，格式：`=SUMPRODUCT(('表'!$A:$A=A2)*('表'!$B:$B="正常")*('表'!$D:$D))`
- 所有跨表公式必须用IFERROR包裹

## 【规则3】日期必转换 - DATEVALUE()
## 【规则4】f-string规则：公式含双引号时，外层用单引号
## 【规则5】❌ 禁止在函数内部导入模块
## 【规则6】工号一般是数字格式，不需要TEXT转换
## 【规则7】注释规范：每列一行 `# 列名(列号): 简要说明`
## 【规则8】条件格式：规则要求标红/高亮/颜色标记时，必须用CellIsRule/FormulaRule实现，放在公式填充之后

# ==================== 代码结构要求 ====================
请生成以下结构的代码：

```python
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    # 1. 主表选择
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

# ==================== 核心规则 ====================
- 跨表取数：优先VLOOKUP，也可用XLOOKUP/INDEX+MATCH/FILTER/SUMPRODUCT（按场景选择），列号 = 目标列位置 - 范围起始列 + 1
- 英文半角括号引号，禁止中文全角
- f-string含双引号时外层用单引号
- 日期参与计算前用DATEVALUE()
- 注释规范：每列一行 `# 列名(列号): 简要说明`
- 条件格式：规则要求标红/高亮时用CellIsRule/FormulaRule实现（已导入）

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
        rules = self._compress_rules(rules_content, max_length=15000)

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
        rules = self._compress_rules(rules_content, max_length=20000)

        # 提取数据清洗规则、警告规则和条件格式规则
        extracted_rules = self.rule_extractor.extract_rules(rules_content)
        data_cleaning_rules_text = ""

        if extracted_rules["data_cleaning_rules"]:
            data_cleaning_rules_text = "\n## 数据清洗规则（在clean_source_data中应用）\n"
            data_cleaning_rules_text += "⚠️ 在将源数据写入Excel之前，必须先应用以下清洗规则过滤数据：\n\n"
            for i, rule in enumerate(extracted_rules["data_cleaning_rules"], 1):
                if rule['original_text'] and not rule['original_text'].startswith('--'):
                    data_cleaning_rules_text += f"{i}. {rule['original_text']}\n"
            data_cleaning_rules_text += "\n**实现方式**：生成clean_source_data函数，对每个DataFrame应用清洗逻辑后返回清洗后的数据。\n"

        conditional_format_text = ""
        if extracted_rules["conditional_format_rules"]:
            conditional_format_text = "\n## 条件格式规则\n"
            conditional_format_text += "在填充公式后，需要对以下情况应用条件格式（使用CellIsRule/FormulaRule）：\n\n"
            for i, rule in enumerate(extracted_rules["conditional_format_rules"], 1):
                conditional_format_text += f"{i}. {rule['original_text']}\n"

        precision_rules_text = ""
        if extracted_rules["precision_rules"]:
            precision_rules_text = "\n## 数值精度规则\n"
            precision_rules_text += "在生成公式时，必须按以下精度要求使用ROUND函数：\n\n"
            for i, rule in enumerate(extracted_rules["precision_rules"], 1):
                precision_rules_text += f"{i}. {rule['original_text']}\n"

        # 统计总列数和sheet信息
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

        system_prompt = (
            "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
            "包括人力资源、财务、供应链等不同业务场景。"
            "你同时也是一个EXCEL公式大师，熟悉VLOOKUP、IF、SUMIF等公式的使用。"
        )

        # Step 3: 生成代码（包含完整上下文，不再依赖前置分析步骤）
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
- 已定义函数：get_vlookup_col_num(target_col, range_start_col) -> int

{source_structure}
{expected_sheets_info}

## 预期输出结构
{compressed_expected}
{data_cleaning_rules_text}
{conditional_format_text}
{precision_rules_text}

## 计算规则
{rules}

# ==================== 黄金样例（必须严格参照此模式编写）====================
以下是一个完整的fill_result_sheets代码示例，展示了5种常见列类型的正确写法。
你的代码必须完全遵循此模式，包括变量定义位置、缩进层级、注释格式、f-string写法。

```python
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    # === 1. 主表选择（员工数最多的表）===
    main_key = max(source_sheets.keys(), key=lambda k: len(source_sheets[k]['df']))
    main_df = source_sheets[main_key]['df']
    n_rows = len(main_df)

    # === 2. 源表sheet名变量（在for循环外定义，VLOOKUP公式中使用）===
    sn_main = source_sheets[main_key]['ws'].title
    sn_attend = source_sheets['考勤表_Sheet1']['ws'].title       # 示例：考勤表
    sn_memo = source_sheets['薪资备忘录_Sheet1']['ws'].title     # 示例：备忘录

    # === 3. 预计算VLOOKUP列号（在for循环外，用get_vlookup_col_num）===
    col_dept = get_vlookup_col_num("D", "A")         # 部门在D列，主键在A列
    col_attend_days = get_vlookup_col_num("AM", "F") # 出勤天数在AM列，主键在F列

    # === 4. 预定义文本常量（Excel文本比较用，避免f-string引号冲突）===
    TXT_FULLTIME = '"门店全职"'
    TXT_YES = '"是"'

    # === 5. 创建结果sheet并写表头 ===
    ws = wb.create_sheet("结果")
    headers = ["工号", "姓名", "部门", "出勤天数", "日薪", "应发工资"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)

    # === 6. 逐行填充（每列平级排列，缩进统一8空格）===
    for i in range(n_rows):
        r = i + 2

        # A列(1): 工号 - 主表直接复制
        ws.cell(row=r, column=1, value=main_df.iloc[i].get('工号', ''))

        # B列(2): 姓名 - 主表直接复制
        ws.cell(row=r, column=2, value=main_df.iloc[i].get('姓名', ''))

        # C列(3): 部门 - VLOOKUP跨表取数（引用花名册）
        ws.cell(row=r, column=3).value = f"=IFERROR(VLOOKUP(A{{r}},\\'{{sn_memo}}\\'!$A:$J,{{col_dept}},FALSE),{{EMPTY}})"

        # D列(4): 出勤天数 - VLOOKUP跨表取数（主键非A列，注意范围起始列）
        ws.cell(row=r, column=4).value = f"=IFERROR(VLOOKUP(A{{r}},\\'{{sn_attend}}\\'!$F:$BG,{{col_attend_days}},FALSE),0)"

        # E列(5): 日薪 - 公式计算（引用参数sheet）
        ws.cell(row=r, column=5).value = f"=IFERROR(F{{r}}/参数!$B$4,0)"

        # F列(6): 应发工资 - 含TEXT函数时用单引号f-string
        ws.cell(row=r, column=6).value = f'=IF(D{{r}}>0,D{{r}}*E{{r}},0)'

    # 共处理了 6 列（预期 6 列）
```

### 从黄金样例中提取的关键模式
1. **变量定义位置**：sn_xxx、col_xxx、TXT_xxx 全部在for循环外定义
2. **注释格式**：每列仅一行 `# X列(N): 说明`，无分隔线
3. **列平级排列**：每列代码缩进统一8空格，互不嵌套
4. **f-string规则**：
   - 公式含sheet名（单引号）→ 外层双引号：`f"=VLOOKUP(A{{r}},\\'{{sn_xxx}}\\'!...)"`
   - 公式含双引号（TEXT等）→ 外层单引号：`f'=TEXT(A{{r}},"YYYY-MM-DD")'`
   - Excel空字符串 → 用{{EMPTY}}代替
   - Excel文本比较 → 用预定义TXT_xxx变量：`f"=IF(P{{r}}={{TXT_FULLTIME}},1,0)"`
   - 禁止在f-string内部调用函数
5. **VLOOKUP列号**：用get_vlookup_col_num()计算，禁止硬编码数字
6. **主表数据**：直接用main_df.iloc[i].get()复制，不用VLOOKUP
7. **日期计算**：参与运算前用DATEVALUE()转换

# ==================== 补充规则 ====================

## 历史数据引用（仅规则涉及"累计""历史"时使用）
- 系统自动创建"历史数据"sheet，第1列为薪资月份，其余列与结果sheet列名一致
- 累计示例：`=SUMIFS(历史数据!$E:$E, 历史数据!$A:$A, "<"&salary_month, 历史数据!$B:$B, B{{r}})`
- 需用IFERROR包裹（第1个月历史数据可能为空）
- 规则中未涉及历史数据时不要引用此sheet

## 工号类型
- 一般为数字格式，不需要TEXT转换
- 仅当工号包含字母或特殊字符时才用TEXT转换

## 完整性要求
- 必须为全部 {total_columns} 列生成处理逻辑
- 在代码最后添加注释：# 共处理了 X 列（预期 {total_columns} 列）
- 每行代码必须完整闭合，不允许跨行写赋值语句
- 禁止"暂时跳过"、"简化为0"、提前结束
- 禁止在函数内部import已导入的模块
- 全部使用英文半角字符 ()[]"'，禁止中文全角

# ==================== 函数签名 ====================

## 1. 数据清洗函数（如果有清洗规则）
def clean_source_data(source_data):
    \"\"\"应用数据清洗规则过滤源数据\"\"\"
    pass

## 2. 结果填充函数
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):
    \"\"\"创建结果sheet并填充数据和公式\"\"\"
    pass

请严格参照黄金样例的代码模式，生成完整代码：
1. 如果有数据清洗规则，先生成clean_source_data函数
2. 然后生成fill_result_sheets函数，覆盖全部 {total_columns} 列"""

        # Step 4: 验证修正（占位符__GENERATED_CODE__在使用时替换）
        step4 = f"""## 任务：验证并修正生成的代码

请逐项检查以下代码，找出问题并输出修正后的完整代码。

## 生成的代码
```python
__GENERATED_CODE__
```

## 检查清单
### 1. 语法检查
- [ ] 所有括号是否正确闭合（圆括号、方括号、花括号）
- [ ] f-string引号是否正确（统一 f"..." + \\' + \\"）
- [ ] 每行代码是否完整，没有截断，没有跨行赋值
- [ ] 缩进是否一致（for循环内8空格，列之间平级，无级联嵌套）

### 2. VLOOKUP验证
- [ ] 每个VLOOKUP取的字段是否正确（字段名和源表对应）
- [ ] range_start是否等于该表主键所在列（不是固定$A）
- [ ] col_index = 目标列绝对位置 - 主键列绝对位置 + 1，计算是否正确
- [ ] range_end是否覆盖到目标列

### 3. 列完整性
- [ ] 是否覆盖了全部 {total_columns} 列
- [ ] 基础列是否直接赋值，计算列是否有公式

### 4. 注释规范
- [ ] 每列注释是否只有一行：`# 列名(列号): 简要说明`
- [ ] 是否有多余的分隔线（# ------、# ===等）→ 必须删除

### 5. 其他
- [ ] sheet名是否用sn_变量，无硬编码
- [ ] 无中文全角字符（括号、引号）
- [ ] 无函数内import

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
