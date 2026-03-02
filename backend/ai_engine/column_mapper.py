"""
列映射模块 - 在生成代码前建立预期列名到源数据的映射关系

核心思路：
1. 精确匹配：列名完全相同 → 直接映射
2. 相似匹配：列名相似（字符串相似度）→ 自动映射
3. AI匹配：完全找不到 → 交给AI处理
"""

import logging
from typing import Dict, List, Any, Optional, Tuple
from difflib import SequenceMatcher

logger = logging.getLogger(__name__)


class ColumnMapper:
    """列映射器 - 建立预期列名到源数据的映射关系"""

    def __init__(self, similarity_threshold: float = 0.85):
        """初始化

        Args:
            similarity_threshold: 相似度阈值（0.0~1.0），高于此值认为是相似匹配
        """
        self.similarity_threshold = similarity_threshold

    def build_column_mapping(
        self,
        expected_columns: List[str],
        source_info: Dict[str, Any]
    ) -> Dict[str, Any]:
        """建立列映射关系

        Args:
            expected_columns: 预期输出的列名列表
            source_info: 源数据信息，格式：
                {
                    "sheets": {
                        "文件名_Sheet名": {
                            "columns": ["列1", "列2", ...],
                            "row_count": 100,
                            "source_file": "原始文件名.xlsx"
                        }
                    }
                }

        Returns:
            映射结果字典：
            {
                "exact_matches": {
                    "预期列名": {
                        "source_sheet": "文件名_Sheet名",
                        "source_column": "源列名",
                        "source_file": "原始文件名.xlsx",
                        "match_type": "exact"
                    }
                },
                "similar_matches": {
                    "预期列名": {
                        "source_sheet": "文件名_Sheet名",
                        "source_column": "源列名",
                        "source_file": "原始文件名.xlsx",
                        "match_type": "similar",
                        "similarity": 0.92
                    }
                },
                "unmatched": ["列名1", "列名2", ...],
                "statistics": {
                    "total": 总列数,
                    "exact": 精确匹配数,
                    "similar": 相似匹配数,
                    "unmatched": 未匹配数
                }
            }
        """
        exact_matches = {}
        similar_matches = {}
        unmatched = []

        # 构建源数据的列名索引：{列名: [(sheet名, 文件名), ...]}
        source_column_index = self._build_source_column_index(source_info)

        for expected_col in expected_columns:
            # 1. 尝试精确匹配
            if expected_col in source_column_index:
                # 如果有多个sheet包含同名列，选择第一个
                sheet_name, source_file = source_column_index[expected_col][0]
                exact_matches[expected_col] = {
                    "source_sheet": sheet_name,
                    "source_column": expected_col,
                    "source_file": source_file,
                    "match_type": "exact"
                }
                logger.info(f"精确匹配: '{expected_col}' → {sheet_name}.{expected_col}")
                continue

            # 2. 尝试相似匹配
            best_match = self._find_similar_column(expected_col, source_column_index)
            if best_match:
                source_col, similarity, sheet_name, source_file = best_match
                similar_matches[expected_col] = {
                    "source_sheet": sheet_name,
                    "source_column": source_col,
                    "source_file": source_file,
                    "match_type": "similar",
                    "similarity": similarity
                }
                logger.info(f"相似匹配: '{expected_col}' → {sheet_name}.{source_col} (相似度: {similarity:.2%})")
                continue

            # 3. 未匹配
            unmatched.append(expected_col)
            logger.warning(f"未匹配: '{expected_col}'")

        # 统计信息
        statistics = {
            "total": len(expected_columns),
            "exact": len(exact_matches),
            "similar": len(similar_matches),
            "unmatched": len(unmatched)
        }

        logger.info(f"列映射完成: 总计 {statistics['total']} 列, "
                   f"精确匹配 {statistics['exact']} 列, "
                   f"相似匹配 {statistics['similar']} 列, "
                   f"未匹配 {statistics['unmatched']} 列")

        return {
            "exact_matches": exact_matches,
            "similar_matches": similar_matches,
            "unmatched": unmatched,
            "statistics": statistics
        }

    def _build_source_column_index(
        self,
        source_info: Dict[str, Any]
    ) -> Dict[str, List[Tuple[str, str]]]:
        """构建源数据列名索引

        Args:
            source_info: 源数据信息

        Returns:
            {列名: [(sheet名, 文件名), ...]}
        """
        column_index = {}

        for sheet_name, sheet_data in source_info.get("sheets", {}).items():
            source_file = sheet_data.get("source_file", "")
            for col_name in sheet_data.get("columns", []):
                if col_name not in column_index:
                    column_index[col_name] = []
                column_index[col_name].append((sheet_name, source_file))

        return column_index

    def _find_similar_column(
        self,
        expected_col: str,
        source_column_index: Dict[str, List[Tuple[str, str]]]
    ) -> Optional[Tuple[str, float, str, str]]:
        """查找相似的列名

        Args:
            expected_col: 预期列名
            source_column_index: 源列名索引

        Returns:
            (源列名, 相似度, sheet名, 文件名) 或 None
        """
        best_match = None
        best_similarity = 0.0

        for source_col, locations in source_column_index.items():
            similarity = self._calculate_similarity(expected_col, source_col)
            if similarity > best_similarity and similarity >= self.similarity_threshold:
                best_similarity = similarity
                sheet_name, source_file = locations[0]  # 选择第一个位置
                best_match = (source_col, similarity, sheet_name, source_file)

        return best_match

    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """计算两个字符串的相似度

        Args:
            str1: 字符串1
            str2: 字符串2

        Returns:
            相似度（0.0~1.0）
        """
        # 使用SequenceMatcher计算相似度
        return SequenceMatcher(None, str1, str2).ratio()

    def format_mapping_for_prompt(
        self,
        mapping_result: Dict[str, Any]
    ) -> str:
        """格式化映射结果为提示词

        Args:
            mapping_result: build_column_mapping的返回结果

        Returns:
            格式化的映射描述
        """
        lines = []
        lines.append("## 列映射关系")
        lines.append("")

        # 精确匹配
        if mapping_result["exact_matches"]:
            lines.append("### 精确匹配（直接使用）")
            for expected_col, match_info in mapping_result["exact_matches"].items():
                lines.append(f"- `{expected_col}` → `{match_info['source_sheet']}.{match_info['source_column']}`")
            lines.append("")

        # 相似匹配
        if mapping_result["similar_matches"]:
            lines.append("### 相似匹配（建议使用）")
            for expected_col, match_info in mapping_result["similar_matches"].items():
                similarity = match_info['similarity']
                lines.append(f"- `{expected_col}` → `{match_info['source_sheet']}.{match_info['source_column']}` (相似度: {similarity:.0%})")
            lines.append("")

        # 未匹配
        if mapping_result["unmatched"]:
            lines.append("### 未匹配（需要计算或AI推断）")
            for col in mapping_result["unmatched"]:
                lines.append(f"- `{col}` - 需要根据业务规则计算或从多个源列组合")
            lines.append("")

        # 统计
        stats = mapping_result["statistics"]
        lines.append(f"**统计**: 总计 {stats['total']} 列, "
                    f"精确匹配 {stats['exact']} 列, "
                    f"相似匹配 {stats['similar']} 列, "
                    f"未匹配 {stats['unmatched']} 列")

        return "\n".join(lines)
