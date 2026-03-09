"""
智能文件和表头匹配器 V2 - 基于数据样例的精确匹配
"""

import os
import json
import logging
from typing import Dict, List, Any, Tuple, Optional
from pathlib import Path
import pandas as pd

logger = logging.getLogger(__name__)


class SmartMatcherV2:
    """智能匹配器 V2 - 基于数据样例进行匹配"""

    def __init__(self, ai_provider=None):
        """初始化

        Args:
            ai_provider: AI提供者实例
        """
        self.ai_provider = ai_provider
        # 从环境变量读取表头相似度阈值，默认0.3
        self.header_similarity_threshold = float(
            os.environ.get("HEADER_SIMILARITY_THRESHOLD", "0.3")
        )

    def match_files_and_headers(
        self,
        training_folder: str,
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None,
        script_content: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """匹配文件和表头

        Args:
            training_folder: 训练文件夹路径（包含训练时的source文件）
            input_files: 输入文件列表
            manual_headers: 手动表头配置
            script_content: 生成的Python脚本内容，用于检测列名引用

        Returns:
            (是否成功, 错误信息, 映射关系)
        """
        # 1. 读取训练时的source文件
        training_source_folder = os.path.join(training_folder, "source")
        if not os.path.exists(training_source_folder):
            return False, f"训练文件夹不存在: {training_source_folder}", None

        training_files_data = self._read_files_with_samples(
            training_source_folder, manual_headers
        )

        # 2. 读取上传的文件
        input_files_data = self._read_files_with_samples_from_list(
            input_files, manual_headers
        )

        # 3. 对比差异
        diff_report = self._compare_files(training_files_data, input_files_data)

        # 4. 如果完全精确匹配，直接返回恒等映射
        if diff_report["is_exact_match"]:
            logger.info("文件和表头完全匹配，直接使用")
            return True, None, self._create_identity_mapping(input_files_data)

        # 5. 如果没有实质性差异（只有文件名/sheet名不同但已通过单sheet映射等确定），直接生成映射
        if not diff_report.get("differences"):
            logger.info("无实质性差异，根据已确定的映射关系直接生成映射")
            mapping = self._create_mapping_from_file_sheet_mapping(diff_report, input_files_data)
            return True, None, mapping

        # 5.5. 检查是否只有文件名/Sheet名差异，但所有表头结构完全一致
        only_name_diff = self._check_only_name_differences(diff_report)
        if only_name_diff:
            logger.info("所有表头结构完全一致，只有文件名/Sheet名不同，直接生成映射（跳过AI）")
            mapping = self._create_mapping_from_file_sheet_mapping(diff_report, input_files_data)
            return True, None, mapping

        # 6. 有差异，使用AI分析
        logger.info("文件或表头不完全匹配，使用AI进行智能匹配...")
        logger.info(f"差异报告:\n{json.dumps(diff_report, ensure_ascii=False, indent=2, default=str)}")

        if not self.ai_provider:
            return False, "文件或表头不匹配，且未配置AI进行智能匹配", None

        try:
            ai_match, error_msg, ai_mapping = self._ai_match_with_diff(
                training_files_data, input_files_data, diff_report,
                script_content=script_content
            )
            if ai_match:
                logger.info("AI匹配成功")
                # 合并file_sheet_mapping中的确定性映射（如单sheet映射），防止AI遗漏
                ai_mapping = self._merge_deterministic_mapping(diff_report, ai_mapping, input_files_data)
                return True, None, ai_mapping
            else:
                logger.error(f"AI匹配失败: {error_msg}")
                return False, error_msg, None
        except Exception as e:
            logger.error(f"AI匹配过程出错: {e}")
            return False, f"AI匹配失败: {str(e)}", None

    def _read_files_with_samples(
        self,
        folder_path: str,
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """读取文件夹中的所有Excel文件，只获取结构信息（文件名、sheet名、列名）

        Returns:
            {
                "file1.xlsx": {
                    "sheets": {
                        "Sheet1": {
                            "headers": ["列1", "列2"]
                        }
                    }
                }
            }
        """
        files_data = {}
        for file_name in os.listdir(folder_path):
            if not file_name.endswith(('.xlsx', '.xls')):
                continue

            file_path = os.path.join(folder_path, file_name)
            try:
                from excel_parser import IntelligentExcelParser

                file_manual_headers = None
                if manual_headers:
                    file_manual_headers = manual_headers.get(file_name)

                parser = IntelligentExcelParser()
                sheet_list = parser.parse_excel_file(file_path, manual_headers=file_manual_headers, headers_only=True, active_sheet_only=True)

                sheets_data = {}
                for sheet_data in sheet_list:
                    sheet_name = sheet_data.sheet_name
                    all_headers = []
                    for region in sheet_data.regions:
                        all_headers.extend(list(region.head_data.keys()))

                    if all_headers:
                        sheets_data[sheet_name] = {
                            "headers": all_headers
                        }

                if sheets_data:
                    files_data[file_name] = {"sheets": sheets_data}
            except Exception as e:
                logger.warning(f"读取文件 {file_path} 失败: {e}")
                continue

        return files_data

    def _read_files_with_samples_from_list(
        self,
        file_paths: List[str],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """从文件路径列表读取文件，只获取结构信息"""
        files_data = {}
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            try:
                from excel_parser import IntelligentExcelParser

                file_manual_headers = None
                if manual_headers:
                    file_manual_headers = manual_headers.get(file_name)

                parser = IntelligentExcelParser()
                sheet_list = parser.parse_excel_file(file_path, manual_headers=file_manual_headers, headers_only=True, active_sheet_only=True)

                sheets_data = {}
                for sheet_data in sheet_list:
                    sheet_name = sheet_data.sheet_name
                    all_headers = []
                    for region in sheet_data.regions:
                        all_headers.extend(list(region.head_data.keys()))

                    if all_headers:
                        sheets_data[sheet_name] = {
                            "headers": all_headers,
                            "file_path": file_path
                        }

                if sheets_data:
                    files_data[file_name] = {
                        "sheets": sheets_data,
                        "file_path": file_path
                    }
            except Exception as e:
                logger.warning(f"读取文件 {file_path} 失败: {e}")
                continue

        return files_data

    def _calc_header_similarity(self, headers_a: List[str], headers_b: List[str]) -> float:
        """计算两组表头的相似度（Jaccard系数）

        Args:
            headers_a: 第一组表头
            headers_b: 第二组表头

        Returns:
            0.0 ~ 1.0 的相似度
        """
        if not headers_a and not headers_b:
            return 1.0
        if not headers_a or not headers_b:
            return 0.0
        set_a = set(headers_a)
        set_b = set(headers_b)
        intersection = len(set_a & set_b)
        union = len(set_a | set_b)
        return intersection / union if union > 0 else 0.0

    def _calc_file_header_similarity(
        self, file_a_sheets: Dict[str, Any], file_b_sheets: Dict[str, Any]
    ) -> float:
        """计算两个文件的整体表头相似度

        将文件内所有sheet的表头合并后计算Jaccard相似度
        """
        all_headers_a = []
        for sheet_data in file_a_sheets.values():
            all_headers_a.extend(sheet_data.get("headers", []))
        all_headers_b = []
        for sheet_data in file_b_sheets.values():
            all_headers_b.extend(sheet_data.get("headers", []))
        return self._calc_header_similarity(all_headers_a, all_headers_b)

    def _match_by_header_similarity(
        self, unmatched_training: List[str], unmatched_input: List[str],
        training_files: Dict[str, Any], input_files: Dict[str, Any]
    ) -> Dict[str, str]:
        """基于表头结构相似度匹配未配对的文件

        Args:
            unmatched_training: 未匹配的训练文件名列表
            unmatched_input: 未匹配的输入文件名列表
            training_files: 训练文件数据
            input_files: 输入文件数据

        Returns:
            {training_file_name: input_file_name} 配对结果
        """
        if not unmatched_training or not unmatched_input:
            return {}

        threshold = self.header_similarity_threshold

        # 计算所有可能配对的相似度
        scores = []
        for t_name in unmatched_training:
            t_sheets = training_files[t_name]["sheets"]
            for i_name in unmatched_input:
                i_sheets = input_files[i_name]["sheets"]
                sim = self._calc_file_header_similarity(t_sheets, i_sheets)
                scores.append((sim, t_name, i_name))

        # 按相似度降序排列，贪心匹配
        scores.sort(key=lambda x: x[0], reverse=True)
        matched = {}
        used_input = set()
        used_training = set()

        for sim, t_name, i_name in scores:
            if sim < threshold:
                break
            if t_name in used_training or i_name in used_input:
                continue
            matched[t_name] = i_name
            used_training.add(t_name)
            used_input.add(i_name)
            logger.info(f"表头结构匹配文件: {i_name} <-> {t_name} (相似度={sim:.2f})")

        return matched

    def _match_sheets_by_header_similarity(
        self, unmatched_t_sheets: List[str], unmatched_i_sheets: List[str],
        training_sheets: Dict[str, Any], input_sheets: Dict[str, Any]
    ) -> Dict[str, str]:
        """基于表头结构相似度匹配未配对的Sheet

        Returns:
            {training_sheet_name: input_sheet_name} 配对结果
        """
        if not unmatched_t_sheets or not unmatched_i_sheets:
            return {}

        threshold = self.header_similarity_threshold

        scores = []
        for t_sheet in unmatched_t_sheets:
            t_headers = training_sheets[t_sheet].get("headers", [])
            for i_sheet in unmatched_i_sheets:
                i_headers = input_sheets[i_sheet].get("headers", [])
                sim = self._calc_header_similarity(t_headers, i_headers)
                scores.append((sim, t_sheet, i_sheet))

        scores.sort(key=lambda x: x[0], reverse=True)
        matched = {}
        used_input = set()
        used_training = set()

        for sim, t_sheet, i_sheet in scores:
            if sim < threshold:
                break
            if t_sheet in used_training or i_sheet in used_input:
                continue
            matched[t_sheet] = i_sheet
            used_training.add(t_sheet)
            used_input.add(i_sheet)
            logger.info(f"表头结构匹配Sheet: '{i_sheet}' <-> '{t_sheet}' (相似度={sim:.2f})")

        return matched

    def _compare_files(
        self,
        training_files: Dict[str, Any],
        input_files: Dict[str, Any]
    ) -> Dict[str, Any]:
        """对比训练文件和输入文件的差异

        匹配策略（按优先级）：
        1. 精确文件名/Sheet名匹配
        2. 单Sheet文件直接映射
        3. 基于表头结构相似度反向匹配（文件级和Sheet级）
        4. 剩余无法匹配的差异交给AI处理

        Returns:
            差异报告
        """
        diff_report = {
            "is_exact_match": True,
            "differences": [],
            "identical_sheets": [],
            "file_sheet_mapping": {}
        }

        # 精确匹配文件名
        matched_pairs = {}  # training_file -> input_file
        used_input = set()

        for t_name in training_files:
            if t_name in input_files:
                matched_pairs[t_name] = t_name
                used_input.add(t_name)

        unmatched_training = [t for t in training_files if t not in matched_pairs]
        unmatched_input = [i for i in input_files if i not in used_input]

        # 文件名不匹配的，先用表头结构相似度反向匹配
        if unmatched_training or unmatched_input:
            diff_report["is_exact_match"] = False

            # 基于表头结构相似度配对
            header_matched = self._match_by_header_similarity(
                unmatched_training, unmatched_input,
                training_files, input_files
            )
            for t_name, i_name in header_matched.items():
                matched_pairs[t_name] = i_name
                used_input.add(i_name)

            # 更新未匹配列表
            still_unmatched_training = [t for t in unmatched_training if t not in header_matched]
            still_unmatched_input = [i for i in unmatched_input if i not in header_matched.values()]

            if still_unmatched_training or still_unmatched_input:
                diff_report["differences"].append({
                    "type": "file_name_mismatch",
                    "training_files": still_unmatched_training,
                    "input_files": still_unmatched_input,
                    "message": f"文件名不匹配: 训练={still_unmatched_training}, 上传={still_unmatched_input}"
                })
                # 剩余数量一致时按顺序配对以便对比header
                if len(still_unmatched_training) == len(still_unmatched_input):
                    for t_name, i_name in zip(still_unmatched_training, still_unmatched_input):
                        matched_pairs[t_name] = i_name
                        logger.info(f"文件名不匹配，临时配对: {i_name} <-> {t_name}")

        # 对每对文件，对比Sheet
        for training_file_name, input_file_name in matched_pairs.items():
            training_sheets = training_files[training_file_name]["sheets"]
            input_sheets = input_files[input_file_name]["sheets"]

            sheet_pairs = {}  # training_sheet -> input_sheet
            used_input_sheets = set()

            # 单Sheet直接映射
            if len(training_sheets) == 1 and len(input_sheets) == 1:
                t_sheet = list(training_sheets.keys())[0]
                i_sheet = list(input_sheets.keys())[0]
                sheet_pairs[t_sheet] = i_sheet
                used_input_sheets.add(i_sheet)
                if t_sheet != i_sheet:
                    diff_report["is_exact_match"] = False
                    logger.info(f"单Sheet直接映射: '{i_sheet}' -> '{t_sheet}' ({input_file_name})")
            else:
                # 精确匹配Sheet名
                for t_sheet in training_sheets:
                    if t_sheet in input_sheets and t_sheet not in used_input_sheets:
                        sheet_pairs[t_sheet] = t_sheet
                        used_input_sheets.add(t_sheet)

                # Sheet名不匹配的，先用表头结构相似度反向匹配
                unmatched_t_sheets = [s for s in training_sheets if s not in sheet_pairs]
                unmatched_i_sheets = [s for s in input_sheets if s not in used_input_sheets]

                if unmatched_t_sheets or unmatched_i_sheets:
                    diff_report["is_exact_match"] = False

                    # 基于表头结构相似度配对Sheet
                    sheet_header_matched = self._match_sheets_by_header_similarity(
                        unmatched_t_sheets, unmatched_i_sheets,
                        training_sheets, input_sheets
                    )
                    for t_s, i_s in sheet_header_matched.items():
                        sheet_pairs[t_s] = i_s
                        used_input_sheets.add(i_s)

                    # 更新未匹配列表
                    still_unmatched_t = [s for s in unmatched_t_sheets if s not in sheet_header_matched]
                    still_unmatched_i = [s for s in unmatched_i_sheets if s not in sheet_header_matched.values()]

                    if still_unmatched_t or still_unmatched_i:
                        diff_report["differences"].append({
                            "type": "sheet_name_mismatch",
                            "file": training_file_name,
                            "input_file": input_file_name,
                            "training_sheets": still_unmatched_t,
                            "input_sheets": still_unmatched_i,
                            "message": f"Sheet名不匹配: 训练={still_unmatched_t}, 上传={still_unmatched_i}"
                        })
                        # 剩余数量一致时临时配对以便对比header
                        if len(still_unmatched_t) == len(still_unmatched_i):
                            for t_s, i_s in zip(still_unmatched_t, still_unmatched_i):
                                sheet_pairs[t_s] = i_s
                                logger.info(f"Sheet名不匹配，临时配对: '{i_s}' <-> '{t_s}'")

            # 记录映射关系
            file_path = input_files[input_file_name].get("file_path", "")
            diff_report["file_sheet_mapping"][input_file_name] = {
                "expected_file": training_file_name,
                "file_path": file_path,
                "sheets": {i_sheet: t_sheet for t_sheet, i_sheet in sheet_pairs.items()}
            }

            # 对每对Sheet，对比表头
            for t_sheet, i_sheet in sheet_pairs.items():
                training_sheet_data = training_sheets[t_sheet]
                input_sheet_data = input_sheets[i_sheet]

                training_headers = set(training_sheet_data["headers"])
                input_headers = set(input_sheet_data["headers"])

                common_headers = training_headers & input_headers
                missing_headers = training_headers - input_headers
                extra_headers = input_headers - training_headers

                if not missing_headers and not extra_headers:
                    diff_report["identical_sheets"].append({
                        "file": training_file_name,
                        "sheet": t_sheet,
                        "input_file": input_file_name,
                        "input_sheet": i_sheet
                    })
                    logger.info(f"Sheet表头完全匹配: {input_file_name}/{i_sheet} -> {training_file_name}/{t_sheet}")
                else:
                    diff_report["is_exact_match"] = False
                    diff_report["differences"].append({
                        "type": "header_mismatch",
                        "file": training_file_name,
                        "sheet": t_sheet,
                        "input_file": input_file_name,
                        "input_sheet": i_sheet,
                        "common_headers": list(common_headers),
                        "missing_headers": list(missing_headers),
                        "extra_headers": list(extra_headers),
                        "all_columns_changed": len(common_headers) == 0
                    })
                    logger.info(
                        f"Sheet表头有差异: {input_file_name}/{i_sheet} -> {training_file_name}/{t_sheet}, "
                        f"相同列: {len(common_headers)}, 缺少: {len(missing_headers)}, 多余: {len(extra_headers)}"
                    )

        return diff_report

    def _check_only_name_differences(self, diff_report: Dict[str, Any]) -> bool:
        """检查是否只有文件名/Sheet名差异，但所有表头结构完全一致

        Args:
            diff_report: 差异报告

        Returns:
            True: 只有名称差异，表头完全一致
            False: 存在表头差异
        """
        differences = diff_report.get("differences", [])

        # 如果没有差异，说明已经在前面处理了
        if not differences:
            return False

        # 检查所有差异是否只是文件名/Sheet名不匹配
        for diff in differences:
            diff_type = diff.get("type")
            # 如果有表头差异，返回False
            if diff_type == "header_mismatch":
                return False

        # 所有差异都是文件名/Sheet名不匹配，且已经建立了映射关系
        # 检查是否所有Sheet都已经通过file_sheet_mapping建立了映射
        file_sheet_mapping = diff_report.get("file_sheet_mapping", {})
        if not file_sheet_mapping:
            return False

        logger.info("检测到只有文件名/Sheet名差异，所有表头结构完全一致")
        return True

    def _check_columns_in_script(self, script_content: str, column_names: set) -> List[str]:
        """检查给定列名是否在脚本中被引用

        Args:
            script_content: Python脚本内容
            column_names: 需要检查的列名集合

        Returns:
            在脚本中被引用的列名列表
        """
        import re
        used_columns = []
        for col in column_names:
            escaped = re.escape(col)
            # 匹配 df['列名'], df["列名"], ['列名'], ["列名"], '列名', "列名"
            patterns = [
                rf"""\[['\"]{escaped}['\"]\]""",
                rf"""['\"]({escaped})['\"]""",
            ]
            for pattern in patterns:
                if re.search(pattern, script_content):
                    used_columns.append(col)
                    break
        return used_columns

    def _create_identity_mapping(self, input_files_data: Dict[str, Any]) -> Dict[str, Any]:
        """创建恒等映射（完全匹配时使用）"""
        mapping = {"file_mapping": {}}

        for file_name, file_data in input_files_data.items():
            file_path = file_data.get("file_path")
            if not file_path:
                continue

            sheet_mapping = {}
            header_mapping = {}

            for sheet_name, sheet_data in file_data["sheets"].items():
                sheet_mapping[sheet_name] = sheet_name
                for header in sheet_data["headers"]:
                    header_mapping[header] = header

            mapping["file_mapping"][file_path] = {
                "expected_file": file_name,
                "sheet_mapping": sheet_mapping,
                "header_mapping": header_mapping
            }

        return mapping

    def _create_mapping_from_file_sheet_mapping(
        self, diff_report: Dict[str, Any], input_files_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """根据file_sheet_mapping生成映射（单sheet映射等确定性场景，无需AI）"""
        mapping = {"file_mapping": {}}
        fsm = diff_report.get("file_sheet_mapping", {})

        for input_file_name, info in fsm.items():
            expected_file = info["expected_file"]
            file_path = info.get("file_path") or input_files_data.get(input_file_name, {}).get("file_path", "")
            if not file_path:
                continue

            # sheet映射：input_sheet -> training_sheet
            sheet_mapping = info.get("sheets", {})

            # header恒等映射
            header_mapping = {}
            file_data = input_files_data.get(input_file_name, {})
            for sheet_name, sheet_data in file_data.get("sheets", {}).items():
                for header in sheet_data.get("headers", []):
                    header_mapping[header] = header

            mapping["file_mapping"][file_path] = {
                "expected_file": expected_file,
                "sheet_mapping": sheet_mapping,
                "header_mapping": header_mapping
            }
            logger.info(f"确定性映射: {input_file_name} -> {expected_file}, sheets={sheet_mapping}")

        return mapping

    def _merge_deterministic_mapping(
        self, diff_report: Dict[str, Any], ai_mapping: Dict[str, Any],
        input_files_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """将file_sheet_mapping中的确定性映射合并到AI结果中，确保所有文件都有映射"""
        fsm = diff_report.get("file_sheet_mapping", {})
        ai_file_mapping = ai_mapping.get("file_mapping", {})

        for input_file_name, info in fsm.items():
            file_path = info.get("file_path") or input_files_data.get(input_file_name, {}).get("file_path", "")
            if not file_path:
                continue

            if file_path in ai_file_mapping:
                # AI已有此文件映射，补充缺失的sheet映射
                existing = ai_file_mapping[file_path]
                existing_sheets = existing.get("sheet_mapping", {})
                for i_sheet, t_sheet in info.get("sheets", {}).items():
                    if i_sheet not in existing_sheets:
                        existing_sheets[i_sheet] = t_sheet
                existing["sheet_mapping"] = existing_sheets
            else:
                # AI没有此文件映射，用确定性映射补充
                header_mapping = {}
                file_data = input_files_data.get(input_file_name, {})
                for sheet_name, sheet_data in file_data.get("sheets", {}).items():
                    for header in sheet_data.get("headers", []):
                        header_mapping[header] = header

                ai_file_mapping[file_path] = {
                    "expected_file": info["expected_file"],
                    "sheet_mapping": info.get("sheets", {}),
                    "header_mapping": header_mapping
                }
                logger.info(f"补充确定性映射: {input_file_name} -> {info['expected_file']}, sheets={info.get('sheets', {})}")

        ai_mapping["file_mapping"] = ai_file_mapping
        return ai_mapping

    def _ai_match_with_diff(
        self,
        training_files: Dict[str, Any],
        input_files: Dict[str, Any],
        diff_report: Dict[str, Any],
        script_content: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """使用AI基于差异报告进行匹配

        在调用AI之前，先检查是否存在所有列都变化且在脚本中被引用的情况，
        如果是则直接报错。AI只处理有差异的sheet，相同部分跳过。
        """

        # 前置检查：全列变化 + 脚本引用 → 直接报错
        if script_content:
            for diff in diff_report.get("differences", []):
                if diff.get("type") == "header_mismatch" and diff.get("all_columns_changed"):
                    file_name = diff.get("file")
                    sheet_name = diff.get("sheet")
                    missing_headers = set(diff.get("missing_headers", []))

                    used_cols = self._check_columns_in_script(script_content, missing_headers)
                    if used_cols:
                        error_msg = (
                            f"文件 {file_name} 的 Sheet '{sheet_name}' 中所有列名都已更改，"
                            f"且以下列在处理脚本中被使用: {', '.join(used_cols)}。"
                            f"无法自动匹配，请检查上传的数据源是否正确。"
                        )
                        logger.error(error_msg)
                        return False, error_msg, None

        # 构建AI提示词（只包含差异报告）
        prompt = f"""你是一个Excel文件结构匹配专家。请分析以下差异报告，判断它们是否可以匹配，并生成映射关系。

## 差异报告
{json.dumps(diff_report["differences"], ensure_ascii=False, indent=2, default=str)}

## 已确认匹配的部分（无需处理）
{json.dumps(diff_report.get("identical_sheets", []), ensure_ascii=False, indent=2)}

## 任务
1. 只需要处理差异报告中列出的不匹配项，已确认匹配的sheet无需处理
2. 重点关注：
   - 文件名是否表示相同的业务含义（如"01_员工信息"和"02_员工信息"是同一业务文件，只是编号不同）
   - Sheet名是否表示相同的业务含义（如"考勤记录1"和"考勤记录"是同一业务Sheet，只是带了数字后缀）
   - Sheet名可能带有日期月份后缀（如"考勤记录202501"对应"考勤记录"）
   - 如果一个文件只有一个Sheet，训练时也只有一个Sheet，则直接映射
   - 变化的列名是否表示相同的业务含义
   - 相同的列名无需映射
3. 如果可以匹配，生成详细的映射关系
4. 如果无法匹配，详细说明原因

## 匹配规则
- 文件名可以不完全相同，只要业务含义一致即可匹配（忽略前缀编号差异）
- Sheet名可以不完全相同，只要业务含义一致即可匹配（忽略数字后缀、日期后缀）
- 列名可以不完全相同，但业务含义应该一致
- 相同的列名直接保留，不需要映射
- 如果训练时的必需列在当前文件中找不到对应列，则匹配失败

## 输出格式
请输出JSON格式，结构如下：
```json
{{
    "success": true/false,
    "error_message": "如果失败，说明原因",
    "file_mapping": {{
        "当前文件名": {{
            "expected_file": "训练时的文件名",
            "sheet_mapping": {{
                "当前sheet名": "训练时的sheet名"
            }},
            "header_mapping": {{
                "当前列名": "训练时的列名"
            }}
        }}
    }}
}}
```

注意：
1. 只输出JSON，不要有其他内容
2. 如果无法匹配，success设为false，并在error_message中详细说明原因
3. header_mapping中只包含需要变更的列映射，相同列名不要包含
4. file_mapping的key必须是当前上传的文件名，不是训练时的文件名
"""

        try:
            # 调用AI（使用流式调用）
            messages = [{"role": "user", "content": prompt}]

            response = ""
            logger.info("开始AI智能匹配...")
            try:
                if hasattr(self.ai_provider, '_openai_chat_stream'):
                    logger.info("使用 OpenAI 流式调用")
                    for chunk, finish_reason in self.ai_provider._openai_chat_stream(messages):
                        if chunk:
                            response += chunk
                            import sys
                            sys.stdout.write(chunk)
                            sys.stdout.flush()
                elif hasattr(self.ai_provider, '_claude_chat_stream'):
                    logger.info("使用 Claude 流式调用")
                    for chunk, finish_reason in self.ai_provider._claude_chat_stream("", messages):
                        if chunk:
                            response += chunk
                            import sys
                            sys.stdout.write(chunk)
                            sys.stdout.flush()
                else:
                    logger.warning("AI provider 不支持流式调用，使用非流式")
                    response = self.ai_provider.chat(messages)
            except Exception as stream_error:
                logger.warning(f"流式调用失败，回退到非流式: {stream_error}")
                response = self.ai_provider.chat(messages)

            logger.info(f"\nAI响应长度: {len(response)} 字符")

            # 解析AI响应
            result = self._parse_ai_response(response)

            if not result:
                return False, "AI响应格式错误", None

            if not result.get("success"):
                error_msg = result.get("error_message", "AI判断无法匹配")
                return False, error_msg, None

            # 转换映射关系（将文件名映射转换为文件路径映射）
            file_mapping = result.get("file_mapping", {})
            path_mapping = {}

            for current_file_name, mapping_info in file_mapping.items():
                # 找到对应的文件路径
                file_data = input_files.get(current_file_name)
                if file_data and "file_path" in file_data:
                    path_mapping[file_data["file_path"]] = mapping_info

            final_mapping = {"file_mapping": path_mapping}
            return True, None, final_mapping

        except Exception as e:
            logger.error(f"AI匹配失败: {e}")
            return False, f"AI匹配过程出错: {str(e)}", None

    def _parse_ai_response(self, response: str) -> Optional[Dict[str, Any]]:
        """解析AI响应"""
        try:
            import re
            json_match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
            if json_match:
                json_str = json_match.group(1)
            else:
                json_str = response.strip()

            result = json.loads(json_str)
            return result
        except Exception as e:
            logger.error(f"解析AI响应失败: {e}, 响应内容: {response[:500]}")
            return None
