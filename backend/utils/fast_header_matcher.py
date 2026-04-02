"""
快速表头匹配器 - 性能优化版

优化策略：
1. headers_only 模式快速解析表头（避免全量解析50MB大文件）
2. 多文件并行解析（ThreadPoolExecutor）
3. 需要重写的文件由 rewrite_excel() 单独做全量解析
"""

import os
import logging
import concurrent.futures
from typing import Dict, List, Any, Tuple, Optional
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)


class FastHeaderMatcher:
    """快速表头匹配器 - 性能优化版（headers_only + 并行 + 按需全量解析）"""

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
        """主入口：表头匹配（headers_only + 并行）

        流程：
        1. 从source_structure提取训练基准
        2. headers_only解析所有上传文件（并行），提取表头
        3. 对比表头建立映射
        注：需要重写的文件由 rewrite_excel() 单独做全量解析
        """
        try:
            # 防御性处理：source_structure 可能是 JSON 字符串（从DB或文件读取时未反序列化）
            if isinstance(source_structure, str):
                import json
                try:
                    source_structure = json.loads(source_structure)
                except (json.JSONDecodeError, TypeError):
                    return False, "source_structure 格式异常（非有效JSON字符串）", None
            if not isinstance(source_structure, dict):
                return False, f"source_structure 类型异常: {type(source_structure).__name__}", None

            # 步骤1: 从source_structure提取训练基准
            logger.info("[匹配] ===== 步骤1: 提取训练基准 =====")
            train_sheets = self._build_training_sheets(source_structure)
            if not train_sheets:
                logger.warning("[匹配] 训练时的source_structure为空或格式异常，将尝试基于文件名兜底匹配")
            else:
                for ts in train_sheets:
                    logger.info(f"[匹配] 训练基准: {ts['file_name']}/{ts['sheet_name']} - {len(ts['headers'])}列")

            # 步骤2: 【优化】解析所有文件表头（headers_only，并行）
            logger.info("[匹配] ===== 步骤2: 解析上传文件表头（并行） =====")
            input_sheets = self._parse_all_files_with_headers(
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

    # ==================== 步骤2: 解析所有文件表头（并行） ====================

    def _parse_all_files_with_headers(
        self, file_paths: List[str], manual_headers: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """解析所有文件的表头（headers_only=True，并行）

        Returns:
            header_info_list: 表头信息列表（用于匹配）
        """
        from excel_parser import IntelligentExcelParser

        header_info_list = []

        def _parse_one_file(file_path):
            """单文件解析（线程安全：每线程独立parser实例）"""
            file_name = os.path.basename(file_path)
            file_manual_headers = None
            if manual_headers:
                file_manual_headers = manual_headers.get(file_name)

            parser = IntelligentExcelParser()
            # 【性能优化】匹配表头阶段开启 headers_only=True，避免全量解析50MB大文件
            # 这能将匹配过程从分钟级提速至秒级
            sheet_list = parser.parse_excel_file(
                file_path, manual_headers=file_manual_headers,
                active_sheet_only=True, best_region_only=True,
                headers_only=True, read_formulas=False
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

        return header_info_list

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

        # 如果没有训练基准，直接基于文件名映射
        if not train_sheets:
            logger.info("[匹配] 无训练基准，直接基于文件名进行1:1映射")
            file_mapping = {}
            for i, input_sheet in enumerate(input_sheets):
                f_name = input_sheet["file_name"]
                if f_name not in file_mapping:
                    file_mapping[f_name] = {
                        "expected_file": f_name,
                        "sheet_mapping": {},
                        "header_mapping": {},
                        "file_path": input_sheet["file_path"]
                    }
                file_mapping[f_name]["sheet_mapping"][input_sheet["sheet_name"]] = input_sheet["sheet_name"]
            return {"success": True, "mapping": {"file_mapping": file_mapping}}

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

                # 如果文件名完全匹配且分数过得去，加权
                if input_sheet['file_name'] == train_file:
                    score += 0.1

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
                # 防御性修复：expected_file 不应有双重后缀（如 .xlsx.xlsx）
                expected = mr["train_file"]
                while expected.endswith('.xlsx.xlsx'):
                    expected = expected[:-5]  # 去掉多余的 .xlsx
                while expected.endswith('.xls.xls'):
                    expected = expected[:-4]

                file_mapping[input_file] = {
                    "expected_file": expected,
                    "sheet_mapping": {},
                    "header_mapping": {},
                    "needs_rewrite": False,
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
        """按映射关系生成新Excel文件

        方案B：匹配阶段只解析表头(headers_only=True)，
        重写阶段对需要映射的文件单独做一次全量解析再写出。
        """
        import openpyxl
        from excel_parser import IntelligentExcelParser

        expected_file = mapping_info["expected_file"]
        sheet_mapping = mapping_info.get("sheet_mapping", {})
        header_mapping = mapping_info.get("header_mapping", {})
        file_path = mapping_info.get("file_path", "")

        output_path = os.path.join(output_dir, expected_file)

        # 对需要重写的文件做一次全量解析（带数据）
        logger.info(f"[匹配] 全量解析文件用于重写: {os.path.basename(file_path)}")
        parser = IntelligentExcelParser()
        parsed_data = parser.parse_excel_file(
            file_path, active_sheet_only=True, best_region_only=True,
            read_formulas=False
        )

        if not parsed_data:
            logger.error(f"[匹配] 全量解析失败，无法生成映射文件: {file_path}")
            return output_path

        # 使用 write_only 模式，内存更低、写入更快
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
        logger.info(f"[匹配] 生成映射文件(write_only): {output_path} ({len(parsed_data)}个sheet)")
        return output_path

    # ==================== 单次解析入口（性能优化版） ====================

    def match_parse_and_prepare(
        self,
        source_structure: Dict[str, Any],
        input_files: List[str],
        manual_headers: Optional[Dict[str, Any]] = None,
        output_dir: Optional[str] = None
    ) -> Tuple[bool, Optional[str], Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
        """单次解析：全量读取 + 表头匹配 + 构建预加载数据 + 写 fallback 文件

        每个上传文件仅打开 1 次（Aspose.Cells），通过内存中的纯 Python 操作
        完成表头匹配、数据映射、预加载构建。

        Returns: (success, error, file_mapping, pre_loaded_source_data)
            file_mapping: 与 match_and_prepare 返回的 mapping["file_mapping"] 格式一致
            pre_loaded_source_data: {"文件名_Sheet名": {"df": DataFrame, "columns": [列名]}}
        """
        try:
            # 防御性处理：source_structure 可能是 JSON 字符串
            if isinstance(source_structure, str):
                import json
                try:
                    source_structure = json.loads(source_structure)
                except (json.JSONDecodeError, TypeError):
                    return False, "source_structure 格式异常（非有效JSON字符串）", None, None
            if not isinstance(source_structure, dict):
                return False, f"source_structure 类型异常: {type(source_structure).__name__}", None, None

            # 步骤1: 从 source_structure 提取训练基准
            logger.info("[单次解析] ===== 步骤1: 提取训练基准 =====")
            train_sheets = self._build_training_sheets(source_structure)
            if not train_sheets:
                logger.warning("[单次解析] source_structure 为空或格式异常，将基于文件名兜底匹配")
            else:
                for ts in train_sheets:
                    logger.info(f"[单次解析] 训练基准: {ts['file_name']}/{ts['sheet_name']} - {len(ts['headers'])}列")

            # 步骤2: 全量解析所有上传文件（每文件仅 1 次 Aspose，read_formulas=False + ExportArray）
            logger.info("[单次解析] ===== 步骤2: 全量解析上传文件（并行） =====")
            input_sheets, parsed_sheets_map = self._parse_all_files_full(input_files, manual_headers)
            if not input_sheets:
                return False, "上传的文件无法读取或为空", None, None
            for si in input_sheets:
                logger.info(f"[单次解析] 解析完成: {si['file_name']}/{si['sheet_name']} - {len(si['headers'])}列")

            # 步骤3: 对比表头（纯 Python，复用已有匹配算法）
            logger.info("[单次解析] ===== 步骤3: 对比表头 =====")
            match_result = self._match_by_training_base(train_sheets, input_sheets)
            if not match_result["success"]:
                return False, match_result["error"], None, None

            file_mapping = match_result["mapping"]["file_mapping"]

            # 步骤4: 从内存构建预加载数据（纯 Python，region → DataFrame → 列重命名）
            logger.info("[单次解析] ===== 步骤4: 构建预加载数据 =====")
            pre_loaded_source_data = self._build_pre_loaded_from_memory(file_mapping, parsed_sheets_map)

            # 步骤5: 文件处理
            # 当 pre_loaded_source_data 有效时，脚本从内存加载数据，完全不读磁盘文件
            # 因此跳过所有文件 I/O（重写 / 重命名），只在预加载失败时才写 fallback 文件
            if output_dir and not pre_loaded_source_data:
                logger.info("[单次解析] ===== 步骤5: 预加载数据为空，写出 fallback 文件 =====")
                self._write_fallback_files(file_mapping, parsed_sheets_map, output_dir)
            elif pre_loaded_source_data:
                logger.info(f"[单次解析] 步骤5: 跳过所有文件 I/O（脚本从内存加载 {len(pre_loaded_source_data)} 个sheet）")

            logger.info(f"[单次解析] ===== 完成: {len(file_mapping)}个文件映射, {len(pre_loaded_source_data)}个sheet预加载 =====")
            return True, None, file_mapping, pre_loaded_source_data

        except Exception as e:
            logger.error(f"[单次解析] 过程出错: {e}", exc_info=True)
            return False, f"单次解析失败: {str(e)}", None, None

    def _parse_all_files_full(
        self, file_paths: List[str], manual_headers: Optional[Dict[str, Any]] = None
    ) -> Tuple[List[Dict[str, Any]], Dict[tuple, Any]]:
        """全量解析所有文件（含数据），返回 (header_info_list, parsed_sheets_map)

        与 _parse_all_files_with_headers 相同的表头提取逻辑，
        但不使用 headers_only，全量读取数据（read_formulas=False + ExportArray）。

        Returns:
            header_info_list: 表头信息列表（用于匹配，格式与 _parse_all_files_with_headers 一致）
            parsed_sheets_map: {(file_path, sheet_name): SheetData} 全量解析数据
        """
        from excel_parser import IntelligentExcelParser

        header_info_list = []
        parsed_sheets_map = {}

        def _parse_one_file(file_path):
            """单文件全量解析（线程安全：每线程独立 parser 实例）"""
            file_name = os.path.basename(file_path)
            file_manual_headers = None
            if manual_headers:
                file_manual_headers = manual_headers.get(file_name)

            parser = IntelligentExcelParser()
            # 全量解析（含数据），read_formulas=False → ExportArray 快速路径
            sheet_list = parser.parse_excel_file(
                file_path, manual_headers=file_manual_headers,
                active_sheet_only=True, best_region_only=True,
                read_formulas=False
            )
            return file_path, file_name, sheet_list

        # 并行解析所有文件
        max_workers = min(len(file_paths), 4)
        if max_workers <= 1:
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
                        logger.warning(f"[单次解析] 并行解析文件失败: {os.path.basename(failed_path)} - {e}")

        # 整理结果
        for file_path, file_name, sheet_list in results:
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
                    parsed_sheets_map[(file_path, sheet_data.sheet_name)] = sheet_data
                    data_rows = sum(len(r.data) for r in sheet_data.regions)
                    logger.info(f"[单次解析] 全量解析完成: {file_name}/{sheet_data.sheet_name} - {len(all_headers)}列, {data_rows}行")

        return header_info_list, parsed_sheets_map

    def _build_pre_loaded_from_memory(
        self,
        file_mapping: Dict[str, Any],
        parsed_sheets_map: Dict[tuple, Any]
    ) -> Dict[str, Any]:
        """从内存中的解析数据构建 pre_loaded_source_data

        与脚本 load_source_data 返回格式一致:
        {"文件名_Sheet名": {"df": DataFrame, "columns": [列名]}}
        """
        from backend.utils.data_helpers import convert_region_to_dataframe
        import pandas as pd

        source_data = {}

        for input_file_name, mapping_info in file_mapping.items():
            expected_file = mapping_info.get("expected_file", input_file_name)
            file_base = expected_file.replace('.xlsx', '').replace('.xls', '')
            file_path = mapping_info.get("file_path", "")
            needs_rewrite = mapping_info.get("needs_rewrite", False)
            sheet_mapping = mapping_info.get("sheet_mapping", {})
            header_mapping = mapping_info.get("header_mapping", {})

            for input_sheet, train_sheet in sheet_mapping.items():
                sheet_data = parsed_sheets_map.get((file_path, input_sheet))
                if not sheet_data:
                    logger.warning(f"[预加载] 未找到解析数据: {input_file_name}/{input_sheet}")
                    continue

                dfs = []
                first_columns = None
                for region in sheet_data.regions:
                    # needs_rewrite 时需要映射表头名（input → train）
                    if needs_rewrite and header_mapping:
                        from excel_parser import ExcelRegion
                        mapped_head = {}
                        for col_name, col_letter in region.head_data.items():
                            mapped_name = header_mapping.get(col_name, col_name)
                            mapped_head[mapped_name] = col_letter
                        mapped_region = ExcelRegion(
                            head_data=mapped_head,
                            data=region.data,
                            formula=region.formula
                        )
                        df = convert_region_to_dataframe(mapped_region)
                    else:
                        df = convert_region_to_dataframe(region)

                    if df.empty and len(df.columns) == 0:
                        continue
                    if first_columns is None:
                        first_columns = list(df.columns)
                    dfs.append(df)

                if not dfs:
                    continue

                if len(dfs) == 1:
                    merged_df = dfs[0]
                else:
                    merged_df = pd.concat(dfs, ignore_index=True)

                key = f"{file_base}_{train_sheet}"
                if len(key) > 31:
                    key = key[:31]

                source_data[key] = {
                    "df": merged_df,
                    "columns": first_columns
                }
                logger.info(f"[预加载] {key}: {len(merged_df)}行 × {len(first_columns)}列")

        return source_data

    def _write_fallback_files(
        self,
        file_mapping: Dict[str, Any],
        parsed_sheets_map: Dict[tuple, Any],
        output_dir: str
    ) -> None:
        """从内存数据写出映射后的文件到磁盘（fallback，不重新解析 Aspose）"""
        import shutil

        for input_file_name, mapping_info in file_mapping.items():
            expected_file = mapping_info.get("expected_file", input_file_name)
            needs_rewrite = mapping_info.get("needs_rewrite", False)
            file_path = mapping_info.get("file_path", "")
            sheet_mapping = mapping_info.get("sheet_mapping", {})
            header_mapping = mapping_info.get("header_mapping", {})

            if needs_rewrite:
                import openpyxl
                output_path = os.path.join(output_dir, expected_file)
                wb = openpyxl.Workbook(write_only=True)

                for input_sheet, train_sheet in sheet_mapping.items():
                    sheet_data = parsed_sheets_map.get((file_path, input_sheet))
                    if not sheet_data:
                        continue

                    ws = wb.create_sheet(title=train_sheet)

                    for region in sheet_data.regions:
                        col_order = []
                        for col_name, col_letter in region.head_data.items():
                            target_name = header_mapping.get(col_name, col_name)
                            col_order.append((target_name, col_letter))

                        # 写表头行
                        ws.append([name for name, _ in col_order])
                        # 写数据行
                        for data_row in region.data:
                            ws.append([data_row.get(col_letter) for _, col_letter in col_order])

                wb.save(output_path)
                wb.close()

                # 删除原文件（如果文件名不同）
                if input_file_name != expected_file:
                    old_path = os.path.join(output_dir, input_file_name)
                    if os.path.exists(old_path):
                        os.remove(old_path)

                logger.info(f"[单次解析] 从内存生成映射文件: {expected_file}")
            else:
                # 不需要重写，但可能需要重命名
                if input_file_name != expected_file:
                    old_path = os.path.join(output_dir, input_file_name)
                    new_path = os.path.join(output_dir, expected_file)
                    if os.path.exists(new_path):
                        os.remove(new_path)
                    if os.path.exists(old_path):
                        shutil.move(old_path, new_path)
                        logger.info(f"[单次解析] 文件重命名: {input_file_name} → {expected_file}")
