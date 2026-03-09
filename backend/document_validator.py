"""
文档格式验证器 - 确保上传文档与训练模版格式一致
"""

import json
import logging
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
from excel_parser import IntelligentExcelParser

logger = logging.getLogger(__name__)


class DocumentValidator:
    """文档格式验证器"""

    def __init__(self):
        self.excel_parser = IntelligentExcelParser()

    def extract_document_schema(self, parsed_data: List[Any]) -> Dict[str, Any]:
        """提取文档格式模版"""
        schema = {
            "sheets": {},
            "total_sheets": 0,
            "validation_rules": {}
        }

        for sheet_data in parsed_data:
            sheet_schema = {
                "sheet_name": sheet_data.sheet_name,
                "header_ranges": [],
                "column_count": 0,
                "headers": {},
                "data_sample_count": 0
            }

            for region in sheet_data.regions:
                # 记录表头范围
                header_range = {
                    "start_row": region.head_row_start,
                    "end_row": region.head_row_end,
                    "data_start_row": region.data_row_start
                }
                sheet_schema["header_ranges"].append(header_range)

                # 记录表头映射
                sheet_schema["headers"].update(region.head_data)

                # 更新列数（取最大值）
                column_count = len(region.head_data)
                if column_count > sheet_schema["column_count"]:
                    sheet_schema["column_count"] = column_count

                # 记录数据样本数量
                sheet_schema["data_sample_count"] += len(region.data)

            schema["sheets"][sheet_data.sheet_name] = sheet_schema

        schema["total_sheets"] = len(schema["sheets"])

        # 添加验证规则
        schema["validation_rules"] = {
            "sheet_names": list(schema["sheets"].keys()),
            "required_headers": self._extract_required_headers(schema),
            "column_counts": {name: info["column_count"] for name, info in schema["sheets"].items()}
        }

        return schema

    def validate_document(
        self, document_data: List[Any], template_schema: Dict[str, Any]
    ) -> Tuple[bool, List[str]]:
        """验证文档格式与模版的一致性

        Sheet名支持模糊匹配：如果精确匹配失败，尝试基于包含关系和相似度匹配。
        """
        errors = []

        # 1. 验证Sheet - 支持模糊匹配
        actual_sheets = {sheet.sheet_name for sheet in document_data}
        expected_sheets = set(template_schema["validation_rules"]["sheet_names"])
        logger.info(f"[调试] DocumentValidator: 实际sheets={actual_sheets}, 期望sheets={expected_sheets}")

        # 构建sheet名映射：actual_name -> expected_name
        sheet_name_map = {}
        matched_expected = set()

        # 第一轮：精确匹配
        for actual in actual_sheets:
            if actual in expected_sheets:
                sheet_name_map[actual] = actual
                matched_expected.add(actual)

        # 第二轮：模糊匹配未匹配的sheet
        unmatched_actual = actual_sheets - set(sheet_name_map.keys())
        unmatched_expected = expected_sheets - matched_expected
        logger.info(f"[调试] 精确匹配后: 未匹配actual={unmatched_actual}, 未匹配expected={unmatched_expected}")

        for expected in list(unmatched_expected):
            best_match = None
            best_score = 0
            for actual in unmatched_actual:
                score = self._sheet_name_similarity(actual, expected)
                logger.info(f"[调试] 模糊匹配: '{actual}' vs '{expected}' -> score={score}")
                if score > best_score:
                    best_score = score
                    best_match = actual
            # 相似度阈值：包含关系或50%以上相似
            if best_match and best_score >= 0.5:
                sheet_name_map[best_match] = expected
                matched_expected.add(expected)
                unmatched_actual.discard(best_match)
                logger.info(f"[调试] 模糊匹配成功: '{best_match}' -> '{expected}' (score={best_score})")

        # 检查仍未匹配的sheet
        still_missing = expected_sheets - matched_expected
        still_extra = actual_sheets - set(sheet_name_map.keys())

        if still_missing:
            errors.append(f"缺少Sheet: {', '.join(still_missing)}")
        if still_extra:
            errors.append(f"多余的Sheet: {', '.join(still_extra)}")

        # 2. 验证每个已匹配的Sheet（使用映射后的名称查找模板）
        for actual_name, expected_name in sheet_name_map.items():
            if expected_name in template_schema["sheets"]:
                sheet_errors = self._validate_sheet(
                    document_data, actual_name, template_schema["sheets"][expected_name]
                )
                errors.extend(sheet_errors)

        return len(errors) == 0, errors

    def _sheet_name_similarity(self, name1: str, name2: str) -> float:
        """计算两个sheet名的相似度

        Returns:
            0.0 ~ 1.0 的相似度分数
        """
        # 完全相同
        if name1 == name2:
            return 1.0

        # 去除空格和常见后缀数字后比较
        import re
        clean1 = re.sub(r'[\s\d_\-]+$', '', name1).strip()
        clean2 = re.sub(r'[\s\d_\-]+$', '', name2).strip()

        if clean1 == clean2:
            return 0.9

        # 包含关系
        if clean1 in clean2 or clean2 in clean1:
            return 0.8

        # 字符级相似度（Jaccard）
        set1 = set(name1)
        set2 = set(name2)
        intersection = len(set1 & set2)
        union = len(set1 | set2)
        if union == 0:
            return 0.0
        return intersection / union

    def _validate_sheet(
        self, document_data: List[Any], sheet_name: str, sheet_template: Dict[str, Any]
    ) -> List[str]:
        """验证单个Sheet

        注意：不验证表头行位置，因为IntelligentExcelParser会自动识别表头位置。
        只验证表头内容（列名是否存在）。
        """
        errors = []

        # 找到对应的Sheet数据
        sheet_data = None
        for sheet in document_data:
            if sheet.sheet_name == sheet_name:
                sheet_data = sheet
                break

        if not sheet_data:
            errors.append(f"Sheet '{sheet_name}' 数据不存在")
            return errors

        # 注意：不再验证区域数量和表头行位置
        # 因为IntelligentExcelParser会智能识别表头位置，不同文件的表头位置可能不同
        # 只要能解析出正确的表头内容即可

        # 验证表头内容（只验证列名是否存在，不验证列位置）
        all_headers = {}
        for region in sheet_data.regions:
            all_headers.update(region.head_data)

        template_headers = sheet_template.get("headers", {})

        # 检查是否包含所有必需的表头（只检查列名，不检查列位置）
        for header in template_headers.keys():
            if header not in all_headers:
                # 尝试模糊匹配（去除空格、大小写不敏感）
                header_normalized = header.strip().lower()
                found = False
                for actual_header in all_headers.keys():
                    if actual_header.strip().lower() == header_normalized:
                        found = True
                        break
                if not found:
                    errors.append(f"Sheet '{sheet_name}' 缺少表头: {header}")

        return errors

    def _extract_required_headers(self, schema: Dict[str, Any]) -> Dict[str, List[str]]:
        """提取必需的表头"""
        required_headers = {}

        for sheet_name, sheet_info in schema["sheets"].items():
            headers = list(sheet_info["headers"].keys())
            # 过滤掉自动生成的列名
            required_headers[sheet_name] = [
                h for h in headers if not h.startswith("Column_")
            ]

        return required_headers

    def validate_file(
        self, file_path: str, template_schema: Dict[str, Any], manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, List[str]]:
        """验证文件格式"""
        try:
            parsed_data = self.excel_parser.parse_excel_file(file_path, manual_headers=manual_headers, active_sheet_only=True)
            return self.validate_document(parsed_data, template_schema)
        except Exception as e:
            return False, [f"解析文件失败: {str(e)}"]

    def compare_schemas(
        self, schema1: Dict[str, Any], schema2: Dict[str, Any]
    ) -> Dict[str, Any]:
        """比较两个格式模版"""
        comparison = {
            "identical": True,
            "differences": [],
            "summary": {
                "sheet_count_diff": schema1["total_sheets"] - schema2["total_sheets"],
                "matching_sheets": [],
                "different_sheets": []
            }
        }

        # 比较Sheet
        sheets1 = set(schema1["sheets"].keys())
        sheets2 = set(schema2["sheets"].keys())

        common_sheets = sheets1.intersection(sheets2)
        only_in_1 = sheets1 - sheets2
        only_in_2 = sheets2 - sheets1

        if only_in_1:
            comparison["identical"] = False
            comparison["differences"].append(f"仅在模版1中的Sheet: {', '.join(only_in_1)}")

        if only_in_2:
            comparison["identical"] = False
            comparison["differences"].append(f"仅在模版2中的Sheet: {', '.join(only_in_2)}")

        # 比较共同的Sheet
        for sheet_name in common_sheets:
            sheet1 = schema1["sheets"][sheet_name]
            sheet2 = schema2["sheets"][sheet_name]

            sheet_comparison = self._compare_sheets(sheet1, sheet2, sheet_name)

            if sheet_comparison["identical"]:
                comparison["summary"]["matching_sheets"].append(sheet_name)
            else:
                comparison["identical"] = False
                comparison["summary"]["different_sheets"].append(sheet_name)
                comparison["differences"].extend(sheet_comparison["differences"])

        return comparison

    def _compare_sheets(
        self, sheet1: Dict[str, Any], sheet2: Dict[str, Any], sheet_name: str
    ) -> Dict[str, Any]:
        """比较两个Sheet的格式"""
        comparison = {
            "identical": True,
            "differences": []
        }

        # 比较区域数量
        if len(sheet1["header_ranges"]) != len(sheet2["header_ranges"]):
            comparison["identical"] = False
            comparison["differences"].append(
                f"Sheet '{sheet_name}' 区域数量不同: {len(sheet1['header_ranges'])} vs {len(sheet2['header_ranges'])}"
            )

        # 比较列数
        if sheet1["column_count"] != sheet2["column_count"]:
            comparison["identical"] = False
            comparison["differences"].append(
                f"Sheet '{sheet_name}' 列数不同: {sheet1['column_count']} vs {sheet2['column_count']}"
            )

        # 比较表头
        headers1 = set(sheet1["headers"].keys())
        headers2 = set(sheet2["headers"].keys())

        if headers1 != headers2:
            comparison["identical"] = False
            missing_in_2 = headers1 - headers2
            extra_in_2 = headers2 - headers1

            if missing_in_2:
                comparison["differences"].append(
                    f"Sheet '{sheet_name}' 模版2缺少表头: {', '.join(missing_in_2)}"
                )

            if extra_in_2:
                comparison["differences"].append(
                    f"Sheet '{sheet_name}' 模版2有多余表头: {', '.join(extra_in_2)}"
                )

        return comparison