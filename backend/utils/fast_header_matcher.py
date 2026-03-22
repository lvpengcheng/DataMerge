"""
快速表头匹配器 - 性能优化版

优化策略：
1. 一次完整解析同时获取表头+数据（避免 headers_only + full 两次解析）
2. 多文件并行解析（ThreadPoolExecutor）
3. 缓存完整解析结果供后续注入脚本执行环境
"""

import os
import logging
import concurrent.futures
from typing import Dict, List, Any, Tuple, Optional
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)


class FastHeaderMatcher:
    """快速表头匹配器 - 性能优化版（一次解析+并行+缓存注入）"""

    def __init__(self, similarity_threshold: float = None):
        if similarity_threshold is None:
            similarity_threshold = float(os.environ.get('HEADER_MATCH_THRESHOLD', '0.85'))
        self.similarity_threshold = similarity_threshold

    @staticmethod
    def _is_valid_header(name) -> bool:
        """过滤空列头"""
        if name is None:
            return False
        s = str(name).strip()
        if not s:
            return False
        if s.startswith('Unnamed:') or s.startswith('Unnamed：'):
            return False
        return True

    # ==================== 主入口 ====================

    def match_and_prepare(
        self,
        source_structure: Dict[str, Any],
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]]]:
        """主入口：一次解析 + 并行 + 缓存

        优化后流程：
        1. 从source_structure提取训练基准
        2. 一次性完整解析所有上传文件（并行），同时提取表头
        3. 对比表头建立映射
        4. 直接从缓存取已解析数据（不再重复读取）
        """
        try:
            # 步骤1: 从source_structure提取训练基准
            logger.info("[匹配] ===== 步骤1: 提取训练基准 =====")
            train_sheets = self._build_training_sheets(source_structure)
            if not train_sheets:
                return False, "训练时的source_structure为空或格式异常", None
            for ts in train_sheets:
                logger.info(f"[匹配] 训练基准: {ts['file_name']}/{ts['sheet_name']} - {len(ts['headers'])}列")

            # 步骤2: 【优化】一次性完整解析所有文件（并行），同时提取表头
            logger.info("[匹配] ===== 步骤2: 完整解析上传文件（并行） =====")
            input_sheets, full_parsed_cache = self._parse_all_files_with_headers(
                input_files, manual_headers
            )
            if not input_sheets:
                return False, "上传的文件无法读取或为空", None
            for si in input_sheets:
                logger.info(f"[匹配] 上传文件: {si['file_name']}/{si['sheet_name']} - {len(si['headers'])}列")

            # 步骤3: 对比表头
            logger.info("[匹配] ===== 步骤3: 对比表头 =====")
            match_result = self._match_by_training_base(train_sheets, input_sheets)

            if not match_result["success"]:
                logger.error(f"[匹配] ===== 匹配失败 =====")
                return False, match_result["error"], None

            logger.info("[匹配] ===== 匹配成功 =====")

            # 步骤4: 【优化】直接从缓存取已解析数据，不再重复读取
            logger.info("[匹配] ===== 步骤4: 关联缓存数据 =====")
            file_mapping = match_result["mapping"]["file_mapping"]
            for input_file, mapping_info in file_mapping.items():
                # 所有文件都从缓存获取完整解析结果（供后续注入脚本环境）
                mapping_info["parsed_data"] = full_parsed_cache.get(
                    mapping_info["file_path"], []
                )
                if mapping_info.get("needs_rewrite"):
                    logger.info(f"[匹配] {input_file} 需要映射，使用缓存数据")
                else:
                    logger.info(f"[匹配] {input_file} 完全一致，缓存数据备用")

            return True, None, match_result["mapping"]

        except Exception as e:
            logger.error(f"[匹配] 过程出错: {e}", exc_info=True)
            return False, f"表头匹配失败: {str(e)}", None

    # ==================== 步骤1: 提取训练基准 ====================

    def _build_training_sheets(self, source_structure: Dict[str, Any]) -> List[Dict[str, Any]]:
        """从source_structure提取每个Sheet的表头"""
        result = []
        files_data = source_structure.get("files", {})

        for file_name, file_data in files_data.items():
            if "error" in file_data:
                continue
            sheets = file_data.get("sheets", {})
            for sheet_name, sheet_data in sheets.items():
                headers = sheet_data.get("headers", {})
                valid = {k: v for k, v in headers.items() if self._is_valid_header(k)}
                if valid:
                    result.append({
                        "file_name": file_name,
                        "sheet_name": sheet_name,
                        "headers": valid
                    })

        return result

    # ==================== 步骤2: 一次性完整解析所有文件（并行） ====================

    def _parse_all_files_with_headers(
        self, file_paths: List[str], manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[List[Dict[str, Any]], Dict[str, List[Any]]]:
        """一次性完整解析所有文件（并行），同时提取表头信息

        Returns:
            (header_info_list, full_parsed_cache)
            - header_info_list: 表头信息列表（用于匹配）
            - full_parsed_cache: {file_path: [SheetData, ...]} 完整解析缓存
        """
        from excel_parser import IntelligentExcelParser

        header_info_list = []
        full_parsed_cache = {}

        def _parse_one_file(file_path):
            """单文件解析（线程安全：每线程独立parser实例）"""
            file_name = os.path.basename(file_path)
            file_manual_headers = None
            if manual_headers:
                file_manual_headers = manual_headers.get(file_name)

            parser = IntelligentExcelParser()
            sheet_list = parser.parse_excel_file(
                file_path, manual_headers=file_manual_headers,
                active_sheet_only=True, best_region_only=True
            )
            return file_path, file_name, sheet_list

        # 并行解析所有文件
        max_workers = min(len(file_paths), 4)
        if max_workers <= 1:
            # 单文件直接串行，避免线程池开销
            results = [_parse_one_file(fp) for fp in file_paths]
        else:
            results = []
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(_parse_one_file, fp): fp for fp in file_paths}
                for future in concurrent.futures.as_completed(futures):
                    try:
                        results.append(future.result())
                    except Exception as e:
                        failed_path = futures[future]
                        logger.warning(f"[匹配] 并行解析文件失败: {os.path.basename(failed_path)} - {e}")

        # 整理结果
        for file_path, file_name, sheet_list in results:
            full_parsed_cache[file_path] = sheet_list

            for sheet_data in sheet_list:
                all_headers = {}
                for region in sheet_data.regions:
                    for k, v in region.head_data.items():
                        if self._is_valid_header(k):
                            all_headers[k] = v

                if all_headers:
                    header_info_list.append({
                        "file_name": file_name,
                        "file_path": file_path,
                        "sheet_name": sheet_data.sheet_name,
                        "headers": all_headers
                    })
                    logger.info(f"[匹配] 解析完成: {file_name}/{sheet_data.sheet_name} - {len(all_headers)}列")

        return header_info_list, full_parsed_cache

    # ==================== 步骤3: 对比表头 ====================

    def _match_by_training_base(
        self,
        train_sheets: List[Dict[str, Any]],
        input_sheets: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """以训练结构为基准，逐个训练Sheet在上传Sheet中找匹配"""
        used_input_indices = set()
        match_results = []
        errors = []

        for train_idx, train_sheet in enumerate(train_sheets):
            train_file = train_sheet["file_name"]
            train_sheet_name = train_sheet["sheet_name"]
            train_headers = train_sheet["headers"]
            train_col_names = set(train_headers.keys())

            logger.info(f"[匹配] 正在匹配训练Sheet: {train_file}/{train_sheet_name} ({len(train_col_names)}列)")

            best_score = 0
            best_input_idx = None
            best_col_mapping = None

            for input_idx, input_sheet in enumerate(input_sheets):
                if input_idx in used_input_indices:
                    continue

                input_headers = input_sheet["headers"]
                col_mapping, score = self._match_headers(
                    list(input_headers.keys()), list(train_headers.keys())
                )

                logger.info(f"[匹配]   vs {input_sheet['file_name']}/{input_sheet['sheet_name']}: 得分={score:.2f}")

                if score > best_score:
                    best_score = score
                    best_input_idx = input_idx
                    best_col_mapping = col_mapping

            if best_score >= self.similarity_threshold and best_input_idx is not None:
                matched_input = input_sheets[best_input_idx]
                used_input_indices.add(best_input_idx)

                logger.info(f"[匹配]   ✓ 匹配成功: {matched_input['file_name']}/{matched_input['sheet_name']} (得分={best_score:.2f})")

                needs_rewrite = not self._is_fully_identical(
                    train_sheet, matched_input, best_col_mapping
                )

                if needs_rewrite:
                    logger.info(f"[匹配]   → 需要生成映射文件")
                else:
                    logger.info(f"[匹配]   → 完全一致，直接使用")

                match_results.append({
                    "train_file": train_file,
                    "train_sheet": train_sheet_name,
                    "train_headers": train_headers,
                    "input_file": matched_input["file_name"],
                    "input_file_path": matched_input["file_path"],
                    "input_sheet": matched_input["sheet_name"],
                    "input_headers": matched_input["headers"],
                    "col_mapping": best_col_mapping,
                    "score": best_score,
                    "needs_rewrite": needs_rewrite
                })
            else:
                logger.error(f"[匹配]   ✗ 未找到匹配")
                missing_info = self._describe_missing(train_sheet, input_sheets, used_input_indices, best_score, best_input_idx)
                errors.append(missing_info)

        if errors:
            error_msg = "以下训练时的数据源在上传文件中未找到匹配:\n" + "\n".join(errors)
            return {"success": False, "error": error_msg}

        file_mapping = self._build_file_mapping(match_results)
        return {"success": True, "mapping": {"file_mapping": file_mapping}}

    def _is_fully_identical(
        self, train_sheet: Dict, input_sheet: Dict, col_mapping: Dict[str, str]
    ) -> bool:
        """判断是否完全一致"""
        if train_sheet["file_name"] != input_sheet["file_name"]:
            return False
        if train_sheet["sheet_name"] != input_sheet["sheet_name"]:
            return False
        for k, v in col_mapping.items():
            if k != v:
                return False
        train_valid = {k: v for k, v in train_sheet["headers"].items() if self._is_valid_header(k)}
        input_valid = {k: v for k, v in input_sheet["headers"].items() if self._is_valid_header(k)}
        return train_valid == input_valid

    def _describe_missing(
        self,
        train_sheet: Dict,
        input_sheets: List[Dict],
        used_indices: set,
        best_score: float,
        best_input_idx: Optional[int]
    ) -> str:
        """生成具体的缺失错误信息"""
        train_file = train_sheet["file_name"]
        train_sheet_name = train_sheet["sheet_name"]
        train_cols = [k for k in train_sheet["headers"].keys() if self._is_valid_header(k)]

        msg = f"\n  【缺失】{train_file} / {train_sheet_name} ({len(train_cols)}列)"

        if best_input_idx is not None and best_score > 0:
            closest = input_sheets[best_input_idx]
            closest_cols = set(k for k in closest["headers"].keys() if self._is_valid_header(k))
            train_col_set = set(train_cols)

            missing_cols = train_col_set - closest_cols
            extra_cols = closest_cols - train_col_set

            msg += f"\n    最接近的上传Sheet: {closest['file_name']}/{closest['sheet_name']} (匹配度={best_score:.0%})"

            if missing_cols:
                missing_list = sorted(missing_cols)
                msg += f"\n    缺少的列 ({len(missing_cols)}列): {', '.join(missing_list[:10])}"
                if len(missing_list) > 10:
                    msg += f"...等共{len(missing_list)}列"

            if extra_cols:
                extra_list = sorted(extra_cols)
                msg += f"\n    多余的列 ({len(extra_cols)}列): {', '.join(extra_list[:10])}"
                if len(extra_list) > 10:
                    msg += f"...等共{len(extra_list)}列"
        else:
            msg += f"\n    在上传的文件中完全找不到表头相似的Sheet"
            msg += f"\n    训练时的列: {', '.join(train_cols[:10])}"
            if len(train_cols) > 10:
                msg += f"...等共{len(train_cols)}列"

        return msg

    def _build_file_mapping(self, match_results: List[Dict]) -> Dict[str, Any]:
        """将匹配结果按上传文件名聚合"""
        file_mapping = {}

        for mr in match_results:
            input_file = mr["input_file"]

            if input_file not in file_mapping:
                file_mapping[input_file] = {
                    "expected_file": mr["train_file"],
                    "sheet_mapping": {},
                    "header_mapping": {},
                    "needs_rewrite": False,
                    "parsed_data": [],  # 稍后填充
                    "file_path": mr["input_file_path"]
                }

            fm = file_mapping[input_file]
            fm["sheet_mapping"][mr["input_sheet"]] = mr["train_sheet"]
            fm["header_mapping"].update(mr["col_mapping"])

            if mr["needs_rewrite"]:
                fm["needs_rewrite"] = True

        return file_mapping

    # ==================== 表头匹配算法 ====================

    def _match_headers(
        self, input_headers: List[str], train_headers: List[str]
    ) -> Tuple[Dict[str, str], float]:
        """匹配两组表头列名"""
        input_headers = [h for h in input_headers if self._is_valid_header(h)]
        train_headers = [h for h in train_headers if self._is_valid_header(h)]

        if not input_headers or not train_headers:
            return {}, 0.0

        input_set = set(input_headers)
        train_set = set(train_headers)

        if input_set == train_set:
            return {h: h for h in input_headers}, 1.0

        header_mapping = {}
        exact = input_set & train_set
        for h in exact:
            header_mapping[h] = h

        unmatched_input = [h for h in input_headers if h not in exact]
        unmatched_train = [h for h in train_headers if h not in exact]

        for inp in unmatched_input:
            best = self._find_similar_header(inp, unmatched_train)
            if best:
                header_mapping[inp] = best
                unmatched_train.remove(best)

        total = max(len(input_headers), len(train_headers))
        score = len(header_mapping) / total if total > 0 else 0.0
        return header_mapping, score

    def _find_similar_header(self, target: str, candidates: List[str]) -> Optional[str]:
        best_match = None
        best_score = 0
        for c in candidates:
            s = SequenceMatcher(None, target, c).ratio()
            if s > best_score and s >= self.similarity_threshold:
                best_score = s
                best_match = c
        return best_match

    # ==================== 生成映射文件 ====================

    @staticmethod
    def rewrite_excel(mapping_info: Dict[str, Any], output_dir: str) -> str:
        """用内存数据按映射关系生成新Excel文件（write_only优化版）"""
        import openpyxl

        expected_file = mapping_info["expected_file"]
        sheet_mapping = mapping_info.get("sheet_mapping", {})
        header_mapping = mapping_info.get("header_mapping", {})
        parsed_data = mapping_info.get("parsed_data", [])

        output_path = os.path.join(output_dir, expected_file)

        # 【性能优化】使用 write_only 模式，内存更低、写入更快
        wb = openpyxl.Workbook(write_only=True)

        for sheet_data in parsed_data:
            target_sheet_name = sheet_mapping.get(sheet_data.sheet_name, sheet_data.sheet_name)
            ws = wb.create_sheet(title=target_sheet_name)

            for region in sheet_data.regions:
                # 构建映射后的列顺序: [(映射后列名, 原始列字母), ...]
                col_order = []
                for col_name, col_letter in region.head_data.items():
                    target_name = header_mapping.get(col_name, col_name)
                    col_order.append((target_name, col_letter))

                # write_only 模式用 ws.append() 按行写入
                ws.append([name for name, _ in col_order])

                for data_row in region.data:
                    ws.append([data_row.get(col_letter) for _, col_letter in col_order])

        wb.save(output_path)
        wb.close()
        logger.info(f"[匹配] 生成映射文件(write_only): {output_path}")
        return output_path
