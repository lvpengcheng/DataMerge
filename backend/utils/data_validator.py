"""
数据校验和自动映射模块

功能：
1. 校验上传数据的表头和数量
2. 根据规则文件中的校验规则进行校验（如有）
3. 自动根据表头映射文件名（当文件名不匹配但表头匹配时）
"""

import os
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import shutil

logger = logging.getLogger(__name__)


class DataValidator:
    """数据校验器

    用于在计算前校验上传的数据是否符合要求，
    并自动根据表头映射文件名。
    """

    def __init__(self, training_structure: Dict[str, Any] = None):
        """初始化

        Args:
            training_structure: 训练时使用的源数据结构
        """
        self.training_structure = training_structure or {}
        self.header_to_file_map = {}  # {frozenset(表头): 文件名}
        self._build_header_map()

    def _build_header_map(self):
        """根据训练结构建立表头到文件名的映射"""
        if not self.training_structure:
            return

        files = self.training_structure.get("files", {})
        for file_name, file_data in files.items():
            sheets = file_data.get("sheets", {})
            for sheet_name, sheet_info in sheets.items():
                headers = self._extract_headers(sheet_info)
                if headers:
                    # 使用表头集合作为键（忽略顺序）
                    header_key = frozenset(headers)
                    self.header_to_file_map[header_key] = {
                        "file_name": file_name,
                        "sheet_name": sheet_name,
                        "headers": headers
                    }

    def _extract_headers(self, sheet_info: Dict[str, Any]) -> List[str]:
        """从sheet信息中提取表头列表"""
        if not isinstance(sheet_info, dict):
            return []

        if "headers" in sheet_info:
            headers = sheet_info["headers"]
            if isinstance(headers, dict):
                return list(headers.keys())
            elif isinstance(headers, list):
                return headers

        if "head_data" in sheet_info:
            return list(sheet_info["head_data"].keys())

        if "regions" in sheet_info:
            regions = sheet_info.get("regions", [])
            if isinstance(regions, list) and regions:
                for region in regions:
                    if isinstance(region, dict) and "head_data" in region:
                        return list(region["head_data"].keys())

        return []

    def validate_and_map(
        self,
        input_folder: str,
        validation_rules: Dict[str, Any] = None
    ) -> Tuple[bool, str, Dict[str, str]]:
        """校验数据并返回文件映射

        Args:
            input_folder: 上传数据的文件夹路径
            validation_rules: 校验规则（从规则文件解析）

        Returns:
            (校验通过, 错误信息, 文件映射 {原文件名: 目标文件名})
        """
        from excel_parser import IntelligentExcelParser

        errors = []
        file_mapping = {}  # {原文件路径: 目标文件名}

        # 解析上传的文件
        parser = IntelligentExcelParser()
        uploaded_files = {}  # {文件名: {sheet名: 表头列表}}

        input_path = Path(input_folder)
        for file_path in input_path.glob("*.xlsx"):
            if file_path.name.startswith("~"):
                continue
            try:
                results = parser.parse_excel_file(str(file_path))
                sheets_info = {}
                for sheet_data in results:
                    headers = []
                    for region in sheet_data.regions:
                        if region.head_data:
                            headers = list(region.head_data.keys())
                            break
                    sheets_info[sheet_data.sheet_name] = headers
                uploaded_files[file_path.name] = {
                    "path": str(file_path),
                    "sheets": sheets_info
                }
            except Exception as e:
                logger.warning(f"解析文件失败 {file_path.name}: {e}")

        # 1. 校验文件数量
        if self.training_structure:
            expected_file_count = len(self.training_structure.get("files", {}))
            actual_file_count = len(uploaded_files)
            if actual_file_count < expected_file_count:
                errors.append(f"文件数量不足: 需要{expected_file_count}个，实际{actual_file_count}个")

        # 2. 根据表头匹配文件
        matched_training_files = set()

        for upload_name, upload_info in uploaded_files.items():
            matched = False
            for sheet_name, headers in upload_info["sheets"].items():
                if not headers:
                    continue
                header_key = frozenset(headers)

                # 精确匹配
                if header_key in self.header_to_file_map:
                    target_info = self.header_to_file_map[header_key]
                    file_mapping[upload_info["path"]] = target_info["file_name"]
                    matched_training_files.add(target_info["file_name"])
                    matched = True
                    logger.info(f"表头精确匹配: {upload_name} -> {target_info['file_name']}")
                    break

                # 模糊匹配（80%以上表头匹配）
                for train_headers, target_info in self.header_to_file_map.items():
                    common = header_key & train_headers
                    similarity = len(common) / max(len(header_key), len(train_headers))
                    if similarity >= 0.8:
                        file_mapping[upload_info["path"]] = target_info["file_name"]
                        matched_training_files.add(target_info["file_name"])
                        matched = True
                        logger.info(f"表头模糊匹配({similarity:.0%}): {upload_name} -> {target_info['file_name']}")
                        break

                if matched:
                    break

            if not matched:
                # 尝试按文件名匹配
                for train_file in self.training_structure.get("files", {}).keys():
                    # 去掉扩展名和数字前缀比较
                    upload_base = upload_name.replace(".xlsx", "").replace(".xls", "")
                    train_base = train_file.replace(".xlsx", "").replace(".xls", "")
                    # 移除数字前缀如 "01_"
                    import re
                    upload_clean = re.sub(r'^\d+_', '', upload_base)
                    train_clean = re.sub(r'^\d+_', '', train_base)

                    if upload_clean == train_clean or upload_base in train_base or train_base in upload_base:
                        file_mapping[upload_info["path"]] = train_file
                        matched_training_files.add(train_file)
                        matched = True
                        logger.info(f"文件名匹配: {upload_name} -> {train_file}")
                        break

                if not matched:
                    logger.warning(f"无法匹配文件: {upload_name}")

        # 3. 检查是否所有训练文件都有对应的上传文件
        if self.training_structure:
            expected_files = set(self.training_structure.get("files", {}).keys())
            missing_files = expected_files - matched_training_files
            if missing_files:
                # 尝试找出缺失文件需要的表头
                missing_details = []
                for missing_file in missing_files:
                    file_info = self.training_structure["files"].get(missing_file, {})
                    sheets = file_info.get("sheets", {})
                    for sheet_name, sheet_info in sheets.items():
                        headers = self._extract_headers(sheet_info)
                        if headers:
                            missing_details.append(f"{missing_file}(需要表头: {headers[:5]}...)")
                            break
                    else:
                        missing_details.append(missing_file)

                errors.append(f"缺少以下数据文件: {', '.join(missing_details)}")

        # 4. 应用自定义校验规则
        if validation_rules:
            rule_errors = self._apply_validation_rules(uploaded_files, validation_rules)
            errors.extend(rule_errors)

        # 返回结果
        if errors:
            return False, "\n".join(errors), file_mapping

        return True, "", file_mapping

    def _apply_validation_rules(
        self,
        uploaded_files: Dict[str, Any],
        rules: Dict[str, Any]
    ) -> List[str]:
        """应用自定义校验规则

        Args:
            uploaded_files: 上传的文件信息
            rules: 校验规则

        Returns:
            错误信息列表
        """
        errors = []

        # 规则示例结构:
        # {
        #   "required_columns": {"员工信息": ["工号", "姓名", "部门"]},
        #   "min_rows": {"员工信息": 1},
        #   "data_types": {"工资": {"基本工资": "number"}},
        #   "value_constraints": [
        #       {"column": "加班小时数", "operator": "<=", "value": 20, "message": "加班小时数不能超过20"}
        #   ]
        # }

        required_columns = rules.get("required_columns", {})
        for file_pattern, columns in required_columns.items():
            found = False
            for file_name, file_info in uploaded_files.items():
                if file_pattern in file_name:
                    found = True
                    for sheet_name, headers in file_info["sheets"].items():
                        missing_cols = set(columns) - set(headers)
                        if missing_cols:
                            errors.append(f"{file_name} 缺少必需列: {missing_cols}")
            if not found:
                errors.append(f"未找到匹配 '{file_pattern}' 的文件")

        # 数值范围校验
        value_constraints = rules.get("value_constraints", [])
        if value_constraints:
            constraint_errors = self._check_value_constraints(uploaded_files, value_constraints)
            errors.extend(constraint_errors)

        return errors

    def _check_value_constraints(
        self,
        uploaded_files: Dict[str, Any],
        constraints: List[Dict[str, Any]]
    ) -> List[str]:
        """检查数值范围约束

        Args:
            uploaded_files: 上传的文件信息
            constraints: 约束规则列表

        Returns:
            错误信息列表
        """
        import pandas as pd

        errors = []

        for constraint in constraints:
            column = constraint.get("column")
            operator = constraint.get("operator")
            value = constraint.get("value")
            message = constraint.get("message", f"{column} 不满足约束条件 {operator} {value}")

            if not column or not operator or value is None:
                continue

            # 遍历所有文件查找该列
            for file_name, file_info in uploaded_files.items():
                file_path = file_info.get("path")
                if not file_path:
                    continue

                try:
                    # 读取Excel文件
                    xl = pd.ExcelFile(file_path)
                    for sheet_name in xl.sheet_names:
                        df = pd.read_excel(xl, sheet_name=sheet_name)
                        if column in df.columns:
                            # 检查约束
                            col_data = pd.to_numeric(df[column], errors='coerce')
                            violations = []

                            if operator == ">":
                                violations = df[col_data > value]
                            elif operator == ">=":
                                violations = df[col_data >= value]
                            elif operator == "<":
                                violations = df[col_data < value]
                            elif operator == "<=":
                                violations = df[col_data <= value]
                            elif operator == "==":
                                violations = df[col_data == value]
                            elif operator == "!=":
                                violations = df[col_data != value]
                            # 反向操作符（用于"不能>"这种表述）
                            elif operator == "!>":
                                violations = df[col_data > value]
                            elif operator == "!>=":
                                violations = df[col_data >= value]
                            elif operator == "!<":
                                violations = df[col_data < value]
                            elif operator == "!<=":
                                violations = df[col_data <= value]

                            if len(violations) > 0:
                                # 获取违规行的详情
                                violation_count = len(violations)
                                sample_values = col_data[col_data.notna()].head(3).tolist()
                                errors.append(
                                    f"{file_name} - {message} (发现{violation_count}条违规数据，示例值: {sample_values})"
                                )
                    xl.close()
                except Exception as e:
                    logger.warning(f"检查约束时出错 {file_name}: {e}")

        return errors

    def prepare_input_folder(
        self,
        source_folder: str,
        target_folder: str,
        file_mapping: Dict[str, str]
    ) -> bool:
        """根据映射准备输入文件夹

        将源文件按照映射关系复制/重命名到目标文件夹。

        Args:
            source_folder: 源文件夹（上传的文件）
            target_folder: 目标文件夹（准备好的文件）
            file_mapping: 文件映射 {源路径: 目标文件名}

        Returns:
            是否成功
        """
        try:
            target_path = Path(target_folder)
            target_path.mkdir(parents=True, exist_ok=True)

            # 清空目标文件夹
            for f in target_path.glob("*.xlsx"):
                f.unlink()

            # 复制文件
            for source_path, target_name in file_mapping.items():
                source = Path(source_path)
                if source.exists():
                    target = target_path / target_name
                    shutil.copy2(source, target)
                    logger.info(f"复制文件: {source.name} -> {target_name}")

            return True
        except Exception as e:
            logger.error(f"准备输入文件夹失败: {e}")
            return False


def parse_validation_rules_from_content(rules_content: str) -> Dict[str, Any]:
    """从规则内容中解析校验规则

    查找规则文件中的校验规则定义。

    Args:
        rules_content: 规则文件内容

    Returns:
        校验规则字典
    """
    import re

    rules = {
        "required_columns": {},
        "min_rows": {},
        "data_types": {},
        "value_constraints": []
    }

    # 查找"校验规则"或"数据校验"部分
    validation_section_patterns = [
        r'#+\s*校验规则\s*\n(.*?)(?=\n#+|\Z)',
        r'#+\s*数据校验\s*\n(.*?)(?=\n#+|\Z)',
        r'【校验规则】\s*(.*?)(?=【|\Z)',
        r'【数据校验】\s*(.*?)(?=【|\Z)',
    ]

    validation_content = ""
    for pattern in validation_section_patterns:
        match = re.search(pattern, rules_content, re.DOTALL | re.IGNORECASE)
        if match:
            validation_content = match.group(1)
            break

    if not validation_content:
        # 如果没有专门的校验规则部分，尝试在整个内容中查找
        validation_content = rules_content

    # 解析必需列
    # 格式示例: "员工信息必须包含: 工号, 姓名, 部门"
    required_pattern = r'(\S+)\s*(?:必须包含|需要包含|必需列)[：:]\s*([^\n]+)'
    for match in re.finditer(required_pattern, validation_content):
        file_pattern = match.group(1)
        columns = [col.strip() for col in match.group(2).split(',')]
        rules["required_columns"][file_pattern] = columns

    # 解析最小行数
    # 格式示例: "员工信息至少1行"
    min_rows_pattern = r'(\S+)\s*(?:至少|最少)\s*(\d+)\s*行'
    for match in re.finditer(min_rows_pattern, validation_content):
        file_pattern = match.group(1)
        min_count = int(match.group(2))
        rules["min_rows"][file_pattern] = min_count

    # 解析数值约束规则
    # 格式示例:
    # - "加班小时数不能>20" / "加班小时数不能大于20"
    # - "加班小时数必须<=20" / "加班小时数必须小于等于20"
    # - "基本工资不能<0" / "基本工资不能小于0"

    # 不能/不可以 + 比较符
    constraint_patterns = [
        # "xxx不能>20" 或 "xxx不能大于20"
        (r'(\S+?)\s*(?:不能|不可以|不应该)\s*(?:>|大于)\s*(\d+(?:\.\d+)?)', '!>', '不能大于'),
        (r'(\S+?)\s*(?:不能|不可以|不应该)\s*(?:>=|大于等于)\s*(\d+(?:\.\d+)?)', '!>=', '不能大于等于'),
        (r'(\S+?)\s*(?:不能|不可以|不应该)\s*(?:<|小于)\s*(\d+(?:\.\d+)?)', '!<', '不能小于'),
        (r'(\S+?)\s*(?:不能|不可以|不应该)\s*(?:<=|小于等于)\s*(\d+(?:\.\d+)?)', '!<=', '不能小于等于'),
        (r'(\S+?)\s*(?:不能|不可以|不应该)\s*(?:==|等于)\s*(\d+(?:\.\d+)?)', '!=', '不能等于'),
        # "xxx必须<=20" 或 "xxx必须小于等于20"
        (r'(\S+?)\s*(?:必须|应该|需要)\s*(?:<=|小于等于)\s*(\d+(?:\.\d+)?)', '!>', '必须小于等于'),
        (r'(\S+?)\s*(?:必须|应该|需要)\s*(?:<|小于)\s*(\d+(?:\.\d+)?)', '!>=', '必须小于'),
        (r'(\S+?)\s*(?:必须|应该|需要)\s*(?:>=|大于等于)\s*(\d+(?:\.\d+)?)', '!<', '必须大于等于'),
        (r'(\S+?)\s*(?:必须|应该|需要)\s*(?:>|大于)\s*(\d+(?:\.\d+)?)', '!<=', '必须大于'),
    ]

    for pattern, operator, desc in constraint_patterns:
        for match in re.finditer(pattern, validation_content):
            column = match.group(1).strip()
            value = float(match.group(2))
            # 清理列名中的前导符号
            column = re.sub(r'^[-•\*\d\.]+\s*', '', column)

            rules["value_constraints"].append({
                "column": column,
                "operator": operator,
                "value": value,
                "message": f"{column}{desc}{value}"
            })

    return rules
