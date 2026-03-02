"""
提示词生成器 - 为AI生成训练和修正提示词（精简版）
"""

import json
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path


class PromptGenerator:
    """提示词生成器"""

    # ============ 通用组件说明（只定义一次）============
    EXCEL_PARSER_INTERFACE = """## IntelligentExcelParser 接口
使用 `from excel_parser import IntelligentExcelParser` 解析Excel。
返回 `List[SheetData]`，每个SheetData包含:
- sheet_name: str
- regions: List[ExcelRegion]

ExcelRegion结构:
- head_data: Dict[str, str] - 表头名到列字母映射，如 {"姓名": "A", "工资": "B"}
- data: List[Dict[str, Any]] - 数据行，格式 {列字母: 值}
- formula: Dict[str, str] - 公式映射"""

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

    def _load_templates(self) -> Dict[str, str]:
        """加载提示词模板"""
        return {
            "training": """你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景。

## 行业背景
人力资源薪资计算项目，涉及薪资、考勤、奖金、社保、税务数据处理。

{excel_parser_interface}

## 主键选择原则
-**不同场景下主键选择不同，不同表之间进行关联主键的选取也不同，所以在开始之前，要分析好针对不同的表关联数据，要使用那个数据做主键**
- HR场景: 优先用雇员工号，其次姓名、身份证号
- 一般场景: 根据数据特点选择唯一标识字段
- 处理主键缺失或重复情况，验证唯一性

## 输入文件结构
{source_structure}

## 预期输出结构
{expected_structure}

## 数据处理规则
{rules_content}

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
6.主键字段是不是和其他表内的类型不一致，导致需要转换

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

        template = self.templates["training"]
        replacements = {
            "{excel_parser_interface}": self.EXCEL_PARSER_INTERFACE,
            "{global_vars}": self.GLOBAL_VARS_DESC,
            "{core_rules}": self.CORE_RULES,
            "{source_structure}": compressed_source,
            "{expected_structure}": compressed_expected,
            "{rules_content}": rules_content,
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
        rules = self._remove_empty_lines(rules_content[:15000])
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
        rules = self._remove_empty_lines(rules_content[:20000])
        _ = manual_headers

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

        template = """你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景，需生成fill_result_sheets函数来创建结果sheet。

⚠️⚠️⚠️ 本次任务共有 __TOTAL_COLUMNS__ 列，你必须为每一列都生成处理逻辑，在全部 __TOTAL_COLUMNS__ 列完成之前绝对不能停止生成！⚠️⚠️⚠️

## 执行流程（固定代码已处理）
1. ✅ 源数据已加载到 source_sheets 字典
2. ✅ 源数据已写入Excel（供公式引用）
3. ✅ 参数sheet已创建（参数!$B$2=年份, $B$3=月份, $B$4=月标准工时）
4. 🎯 你的任务：创建结果sheet，填充数据和Excel公式

## 已有变量
- wb: openpyxl Workbook
- source_sheets: {"文件名_sheet名": {"df": DataFrame, "ws": worksheet}}

__SOURCE_STRUCTURE__
__EXPECTED_SHEETS_INFO__

## 预期输出结构
__COMPRESSED_EXPECTED__

## 计算规则
__RULES__

# ==================== 核心规则（违反=失败）====================
## 【规则1】字符规范 - 最优先检查
- ✅ 必须：英文半角 ()[]"'
- ❌ 禁止：中文全角 （）【】""''
- 🔍 每次输出前检查：括号、引号是否全部半角

## 【规则2】VLOOKUP是唯一跨表取数方式
- 主表数据：直接复制,不能用VLOOKUP（会导致性能问题）
- 非主表数据：**必须且只能**用VLOOKUP，禁止任何其他方式

### VLOOKUP列号规则（最关键！必须使用get_vlookup_col_num函数）
- ⚠️ **禁止硬编码VLOOKUP列号数字！必须使用get_vlookup_col_num()函数计算**
- ⚠️ 在for循环之前，预先计算所有需要的VLOOKUP列号变量
- ⚠️ VLOOKUP范围的第一列必须是主键列，不一定是$A列！
- 函数签名：`get_vlookup_col_num(target_col: str, range_start_col: str) -> int`
- 用法示例：
```python
# 在for循环之前预计算列号
col_bank_account = get_vlookup_col_num("F", "A")   # 银行卡号在F列，主键在A列 → 6
col_attend_work = get_vlookup_col_num("AM", "F")   # 出勤天数在AM列，主键在F列 → 34
# 在for循环内使用变量
f"=IFERROR(VLOOKUP(K{r},\'{sn_bank}\'!$A:$J,{col_bank_account},FALSE),0)"
f"=IFERROR(VLOOKUP(K{r},\'{sn_attend}\'!$F:$BG,{col_attend_work},FALSE),0)"
```
- ❌ 禁止：`f"=VLOOKUP(K{r},\'{sn_attend}\'!$F:$BG,34,FALSE)"` ← 硬编码34
- ✅ 正确：`f"=VLOOKUP(K{r},\'{sn_attend}\'!$F:$BG,{col_attend_work},FALSE)"` ← 使用变量

### sheet名变量规则（必须遵守）
- ⚠️ 必须在for循环之前，为每个源表定义sheet名变量
- ⚠️ 所有VLOOKUP公式中必须使用变量引用sheet名，禁止硬编码
- 定义方式：从source_sheets字典中获取ws.title

## 【规则3】日期必转换
- 所有日期参与计算前必须用 `DATEVALUE()`

## 【规则4】f-string引号规则（最重要的语法规则，彻底避免冲突）
- 每行代码必须完整闭合，不允许截断，不允许跨行写一个赋值语句
- ⚠️ **已定义常量和辅助函数，必须使用：**
  - `EMPTY` = Excel空字符串""，用法：`f"=IFERROR(...,{EMPTY})"`
  - `excel_text('门店全职')` = Excel文本值"门店全职"，用法：`f"=IF(P{r}={excel_text('门店全职')},1,0)"`
- ⚠️ f-string统一写法：
  - 所有公式f-string一律用双引号 f"..."
  - Excel sheet名的单引号一律转义为 \'
  - **Excel空字符串用{EMPTY}代替，禁止写\"\"或""**
  - **Excel文本比较用{excel_text('xxx')}代替，禁止写\"xxx\"**
- ✅ 正确示例：
  - `f"=IFERROR(VLOOKUP(K{r},\'{sn_bank}\'!$A:$J,{col_num},FALSE),{EMPTY})"`
  - `f"=IF(P{r}={excel_text('门店全职')},1,0)"`
- ❌ 绝对禁止：
  - `f"=IFERROR(...,\"\")"` ← 用{EMPTY}代替
  - `f"=IF(P{r}=\"门店全职\",1,0)"` ← 用{excel_text('门店全职')}代替
  - 跨行写赋值语句

## 【规则5】模块导入规则
- ❌ 禁止在函数内部导入已在顶层导入的模块
- 已导入模块：os, pandas(pd), openpyxl(Workbook, Comment, PatternFill, Font, get_column_letter, column_index_from_string)

## 【规则5.1】历史数据查询工具（可选使用）
- 沙箱中提供了 `history_provider` 全局变量（HistoricalDataProvider实例）
- 当规则中涉及"累计"、"本年前几个月汇总"、"历史数据"等需求时可使用
- 可用方法：
  - `history_provider.get_sum(field, year, months=[1,2,3])` → 获取指定字段的汇总值
  - `history_provider.get_avg(field, year, months=[1,2,3])` → 获取平均值
  - `history_provider.get_count(field, year, months=[1,2,3])` → 获取计数
  - `history_provider.get_employee_history(emp_code, year, months, fields)` → 获取指定员工的历史数据DataFrame
  - `history_provider.get_available_months(year)` → 获取有历史数据的月份列表
  - `history_provider.load_history(year, month)` → 加载指定月份的完整DataFrame
- condition参数支持筛选：`{{"field": "部门", "op": "==", "value": "销售部"}}`
- ⚠️ 如果规则中没有涉及历史数据需求，不要使用此工具

## 【规则6】工号类型
- 一般情况下，工号都是数字格式，不需要TEXT转换
- 只有当工号包含字母或特殊字符时，才需要用TEXT转换

## 【规则7】注释规范（严格执行）
- ⚠️ 每列只允许一行注释，格式：`# 列名(列号): 简要说明`
- ❌ 禁止多行注释、分隔线（如 # ------、# ===、# **）、注释块
- ✅ 示例：`# AO列(41): 补产假工资 - VLOOKUP薪资备忘录S列`
- ❌ 禁止示例：
```
# --------------------------------------------------
# AO列(41): 补产假工资
# --------------------------------------------------
```

## 【规则8】列代码结构（最重要的结构规则）
- ⚠️ 每列的代码必须在for循环内平级排列，缩进统一为8空格
- ⚠️ 绝对禁止把下一列的代码嵌套在上一列的if/else分支内
- ⚠️ 每列独立，互不嵌套，即使有if/else判断也必须在当前列内闭合
- ❌ 错误（级联嵌套，缩进越来越深）：
```python
        # D列(4)
        if condition:
            ws.cell(row=r, column=4).value = xxx
        else:
            ws.cell(row=r, column=4, value="")
            # E列(5) ← 错误！嵌套在D列的else里
            ws.cell(row=r, column=5).value = yyy
```
- ✅ 正确（平级排列，缩进一致）：
```python
        # D列(4)
        if condition:
            ws.cell(row=r, column=4).value = xxx
        else:
            ws.cell(row=r, column=4, value="")

        # E列(5) ← 正确！与D列平级
        ws.cell(row=r, column=5).value = yyy
```


# ==================== 执行检查清单 ====================
每生成一列公式前，按顺序检查：
1. [ ] 是否跨表取数？→ 是 → 必须用VLOOKUP
2. [ ] 是否涉及日期？→ 是 → 加DATEVALUE
3. [ ] 是否有文本比较？→ 是 → 值加双引号（如"是"）
4. [ ] 是否有中文标点？→ 是 → 全部改为英文
5. [ ] f-string统一用 f"..." + \' + \" 了吗？

# ==================== 快速参考 ====================
## VLOOKUP列号计算（必须正确）
列号 = 目标列绝对位置 - 范围起始列绝对位置 + 1
⚠️ 范围起始列 = 主键所在列（不一定是A列！）

例1：主键在A列 → 范围$A:$AC，取N列(14) → 列号=14-1+1=14
例2：主键在C列 → 范围$C:$CD，取CD列(82) → 列号=82-3+1=80
例3：主键在F列 → 范围$F:$BG，取AM列(39) → 列号=39-6+1=34

## 常用模板
| 场景 | 公式模板 |
|------|---------|
| 跨表取数(主键A列) | `=IFERROR(VLOOKUP(K2,'{sn_xxx}'!$A:$Z,列号,FALSE),0)` |
| 跨表取数(主键非A列) | `=IFERROR(VLOOKUP(K2,'{sn_xxx}'!$F:$BG,列号,FALSE),0)` |
| 日期比较 | `=IF(DATEVALUE(A2)>参数!$B$5,"是","否")` |
| 表内汇总 | `=SUMIF('{sn_xxx}'!$A:$A,A2,'{sn_xxx}'!$C:$C)` |
| 参数引用 | `参数!$B$4` (绝对引用) |

## 禁止行为（立即停止）
- ❌ "暂时跳过" / "简化为0" → 必须完整实现
- ❌ 直接引用DataFrame值 → 必须改用VLOOKUP
- ❌ 跨行不闭合引号 → 每行必须独立完整
- ❌ 提前结束 → 必须实现预期输出结构中的每一列，一列都不能少

# ==================== 完整性要求（最高优先级）====================
⚠️ 本次任务共有 __TOTAL_COLUMNS__ 列，你必须为每一列都生成对应的处理逻辑，包括基础列和计算列。
⚠️ 生成过程中请自行计数，确认已处理的列数与预期 __TOTAL_COLUMNS__ 列一致后才能结束。
⚠️ 如果你已处理的列数不足 __TOTAL_COLUMNS__ 列，绝对不能停止，必须继续生成剩余列的代码。
⚠️ 绝对禁止提前结束生成！即使代码很长也必须完整输出所有列的处理逻辑，不能省略、跳过或用注释代替。
⚠️ 在代码最后添加注释：# 共处理了 X 列（预期 __TOTAL_COLUMNS__ 列），用于自我验证。

# ==================== 函数签名 ====================
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):

    主表策略：选员工数最全的表
    其他表：全部VLOOKUP
    计算列：完整实现所有逻辑

请生成完整的fill_result_sheets函数代码。必须覆盖预期输出结构中的所有列，不允许遗漏任何一列，全部列的逻辑都生成完毕后才能结束。"""

        return (template
                .replace('__TOTAL_COLUMNS__', str(total_columns))
                .replace('__SOURCE_STRUCTURE__', source_structure)
                .replace('__EXPECTED_SHEETS_INFO__', expected_sheets_info)
                .replace('__COMPRESSED_EXPECTED__', compressed_expected)
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

## 【规则2】VLOOKUP是唯一跨表取数方式
- 主表数据：直接复制,不能用VLOOKUP
- 非主表数据：**必须且只能**用VLOOKUP
- 列号 = 目标列位置 - 范围起始列位置 + 1
- 格式：`=IFERROR(VLOOKUP(主键,'xxx'!$A:$Z,列号,FALSE),0)`

## 【规则3】日期必转换 - DATEVALUE()
## 【规则4】f-string规则：公式含双引号时，外层用单引号
## 【规则5】❌ 禁止在函数内部导入模块
## 【规则6】工号一般是数字格式，不需要TEXT转换
## 【规则7】注释规范：每列一行 `# 列名(列号): 简要说明`

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
- VLOOKUP跨表取数，列号 = 目标列位置 - 范围起始列 + 1
- 英文半角括号引号，禁止中文全角
- f-string含双引号时外层用单引号
- 日期参与计算前用DATEVALUE()
- 注释规范：每列一行 `# 列名(列号): 简要说明`

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
        rules = self._remove_empty_lines(rules_content[:15000])

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
        rules = self._remove_empty_lines(rules_content[:20000])

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
        step3 = f"""你是专业Python程序员，擅长人力资源行业的薪资计算、税务处理、考勤管理，同时你也是一个EXCEL公式大师，熟悉每个公式的使用方法和使用场景，需生成fill_result_sheets函数来创建结果sheet。

⚠️⚠️⚠️ 本次任务共有 {total_columns} 列，你必须为每一列都生成处理逻辑，在全部 {total_columns} 列完成之前绝对不能停止生成！⚠️⚠️⚠️

## 执行流程（固定代码已处理）
1. ✅ 源数据已加载到 source_sheets 字典
2. ✅ 源数据已写入Excel（供公式引用）
3. ✅ 参数sheet已创建（参数!$B$2=年份, $B$3=月份, $B$4=月标准工时）
4. 🎯 你的任务：创建结果sheet，填充数据和Excel公式

## 已有变量
- wb: openpyxl Workbook
- source_sheets: {{"文件名_sheet名": {{"df": DataFrame, "ws": worksheet}}}}

{source_structure}
{expected_sheets_info}

## 预期输出结构
{compressed_expected}

## 计算规则
{rules}

# ==================== 核心规则（违反=失败）====================
## 【规则1】字符规范 - 最优先检查
- ✅ 必须：英文半角 ()[]"'
- ❌ 禁止：中文全角 （）【】""''
- 🔍 每次输出前检查：括号、引号是否全部半角

## 【规则2】VLOOKUP是唯一跨表取数方式
- 主表数据：直接复制,不能用VLOOKUP（会导致性能问题）
- 非主表数据：**必须且只能**用VLOOKUP，禁止任何其他方式

### VLOOKUP列号规则（最关键！必须使用get_vlookup_col_num函数）
- ⚠️ **禁止硬编码VLOOKUP列号数字！必须使用get_vlookup_col_num()函数计算**
- ⚠️ 在for循环之前，预先计算所有需要的VLOOKUP列号变量
- ⚠️ VLOOKUP范围的第一列必须是主键列，不一定是$A列！
- 函数签名：`get_vlookup_col_num(target_col: str, range_start_col: str) -> int`
- 用法示例：
```python
# 在for循环之前预计算列号
col_bank_account = get_vlookup_col_num("F", "A")   # 银行卡号在F列，主键在A列 → 6
col_attend_work = get_vlookup_col_num("AM", "F")   # 出勤天数在AM列，主键在F列 → 34
# 在for循环内使用变量
f"=IFERROR(VLOOKUP(K{{r}},\\'{{sn_bank}}\\'!$A:$J,{{col_bank_account}},FALSE),0)"
f"=IFERROR(VLOOKUP(K{{r}},\\'{{sn_attend}}\\'!$F:$BG,{{col_attend_work}},FALSE),0)"
```
- ❌ 禁止：`f"=VLOOKUP(K{{r}},\\'{{sn_attend}}\\'!$F:$BG,34,FALSE)"` ← 硬编码34
- ✅ 正确：`f"=VLOOKUP(K{{r}},\\'{{sn_attend}}\\'!$F:$BG,{{col_attend_work}},FALSE)"` ← 使用变量

### sheet名变量规则（必须遵守）
- ⚠️ 必须在for循环之前，为每个源表定义sheet名变量
- ⚠️ 所有VLOOKUP公式中必须使用变量引用sheet名，禁止硬编码
- 定义方式：从source_sheets字典中获取ws.title
- ❌ 禁止：f"=VLOOKUP(K{{r}},'7银行卡表_Sheet1'!$A:$J,6,FALSE)"
- ✅ 正确：f"=VLOOKUP(K{{r}},\\'{{sn_bank}}\\'!$A:$J,{{col_bank_account}},FALSE)"

## 【规则3】日期必转换
- 所有日期参与计算前必须用 `DATEVALUE()`

## 【规则4】f-string引号规则（最重要的语法规则，彻底避免冲突）
- 每行代码必须完整闭合，不允许截断，不允许跨行写一个赋值语句
- ⚠️ **已定义EMPTY常量，必须使用：**
  - `EMPTY = '""'` — Excel空字符串，用法：`f"=IFERROR(...,{{EMPTY}})"`
- ⚠️ **Excel文本比较：必须在for循环之前预定义文本常量变量，f-string里只引用变量**
  - 定义方式：`TXT_xxx = '"文本值"'`（外层单引号，内层双引号）
  - 用法：`f"=IF(P{{r}}={{TXT_xxx}},1,0)"`
  - 示例：
```python
# 在for循环之前定义文本常量
TXT_FULLTIME = '"门店全职"'
TXT_YES = '"是"'
TXT_NO = '"否"'

# 在for循环内使用
f"=IF(P{{r}}={{TXT_FULLTIME}},1,0)"
f"=IF(Q{{r}}={{TXT_YES}},100,0)"
```
- ⚠️ f-string统一写法：
  - 所有公式f-string一律用双引号 f"..."
  - Excel sheet名的单引号一律转义为 \\'
  - **Excel空字符串用{{EMPTY}}代替**
  - **Excel文本比较用预定义的TXT_变量代替**
  - **禁止在f-string内部调用任何函数（如excel_text()等）**
- ✅ 正确示例：
  - `f"=IFERROR(VLOOKUP(K{{r}},\\'{{sn_bank}}\\'!$A:$J,{{col_num}},FALSE),{{EMPTY}})"`
  - `f"=IF(P{{r}}={{TXT_FULLTIME}},1,0)"`
  - `f"=IFERROR(VLOOKUP(K{{r}},\\'{{sn_bank}}\\'!$A:$J,{{col_num}},FALSE),0)"`
- ❌ 绝对禁止：
  - `f"=IFERROR(...,\\"\\")"` ← 用{{EMPTY}}代替
  - `f"=IFERROR(...,"")"` ← 双引号冲突
  - `f"=IF(P{{r}}=\\"门店全职\\",1,0)"` ← 用TXT_变量代替
  - `f"=IF(P{{r}}={{excel_text('门店全职')}},1,0)"` ← 禁止在f-string内调用函数
  - 跨行写赋值语句（如 .value = ( 换行 f"..." 换行 )）

## 【规则5】模块导入规则
- ❌ 禁止在函数内部导入已在顶层导入的模块
- 已导入模块：os, pandas(pd), openpyxl(Workbook, Comment, PatternFill, Font, get_column_letter, column_index_from_string)

## 【规则5.1】历史数据查询工具（可选使用）
- 沙箱中提供了 `history_provider` 全局变量（HistoricalDataProvider实例）
- 当规则中涉及"累计"、"本年前几个月汇总"、"历史数据"等需求时可使用
- 可用方法：
  - `history_provider.get_sum(field, year, months=[1,2,3])` → 获取指定字段的汇总值
  - `history_provider.get_avg(field, year, months=[1,2,3])` → 获取平均值
  - `history_provider.get_count(field, year, months=[1,2,3])` → 获取计数
  - `history_provider.get_employee_history(emp_code, year, months, fields)` → 获取指定员工的历史数据DataFrame
  - `history_provider.get_available_months(year)` → 获取有历史数据的月份列表
  - `history_provider.load_history(year, month)` → 加载指定月份的完整DataFrame
- condition参数支持筛选：`{{"field": "部门", "op": "==", "value": "销售部"}}`
- ⚠️ 如果规则中没有涉及历史数据需求，不要使用此工具

## 【规则6】工号类型
- 一般情况下，工号都是数字格式，不需要TEXT转换

## 【规则7】注释规范（严格执行）
- ⚠️ 每列只允许一行注释，格式：`# 列名(列号): 简要说明`
- ❌ 禁止多行注释、分隔线（如 # ------、# ===）、注释块
- ✅ 示例：`# AO列(41): 补产假工资 - VLOOKUP薪资备忘录S列`
- ❌ 禁止示例：
```
# --------------------------------------------------
# AO列(41): 补产假工资
# --------------------------------------------------
```

## 【规则8】列代码结构（最重要的结构规则）
- ⚠️ 每列的代码必须在for循环内平级排列，缩进统一为8空格
- ⚠️ 绝对禁止把下一列的代码嵌套在上一列的if/else分支内
- ⚠️ 每列独立，互不嵌套，即使有if/else判断也必须在当前列内闭合

# ==================== 完整性要求（最高优先级）====================
⚠️ 本次任务共有 {total_columns} 列，你必须为每一列都生成对应的处理逻辑。
⚠️ 在代码最后添加注释：# 共处理了 X 列（预期 {total_columns} 列）

# ==================== 函数签名 ====================
def fill_result_sheets(wb, source_sheets, salary_year=None,
                       salary_month=None, monthly_standard_hours=174):

    主表策略：选员工数最全的表
    其他表：全部VLOOKUP
    计算列：完整实现所有逻辑

请生成完整的fill_result_sheets函数代码。必须覆盖预期输出结构中的所有列，不允许遗漏任何一列。"""

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
