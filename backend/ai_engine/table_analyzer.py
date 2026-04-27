"""表结构分析器 — 主键/主表/列分层/VLOOKUP预计算/L1模板代码

在训练开始前运行一次，产出 TableAnalysisResult 供公式模式和自由模式共用。
所有推断步骤写入 detection_log 供审计。
"""

import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class ColumnMapping:
    col_name: str
    source_sheet: str
    source_col_name: str
    col_index_in_source: int  # 0-based


@dataclass
class VLookupInfo:
    target_col: str
    source_sheet: str
    key_col_letter: str
    target_col_letter: str
    col_num: int
    range_str: str
    formula_template: str


@dataclass
class TableAnalysisResult:
    primary_key: str = ""
    primary_key_source: str = ""
    main_table_type: str = "single"
    main_table_sheets: List[str] = field(default_factory=list)
    main_table_join_keys: List[str] = field(default_factory=list)
    column_layers: Dict[str, str] = field(default_factory=dict)
    l1_columns: List[ColumnMapping] = field(default_factory=list)
    l2_columns: List[ColumnMapping] = field(default_factory=list)
    l3_columns: List[str] = field(default_factory=list)
    l4_columns: List[str] = field(default_factory=list)
    vlookup_map: Dict[str, VLookupInfo] = field(default_factory=dict)
    merge_config: Optional[dict] = None
    confidence: float = 0.0
    detection_log: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def _col_index_to_letter(idx: int) -> str:
    """0-based列索引转Excel列字母 (0->A, 25->Z, 26->AA)"""
    result = ""
    idx += 1  # 转为1-based
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _normalize_col_name(name: str) -> str:
    """标准化列名用于比较"""
    if not isinstance(name, str):
        name = str(name)
    name = name.strip().lower()
    name = name.replace("\n", " ").replace("\r", "")
    name = re.sub(r'\s+', ' ', name)
    name = name.replace("（", "(").replace("）", ")")
    return name


def _col_overlap_ratio(cols_a: list, cols_b: list) -> float:
    """两组列名的重叠率 (Jaccard-like, 以较小集合为分母)"""
    set_a = {_normalize_col_name(c) for c in cols_a}
    set_b = {_normalize_col_name(c) for c in cols_b}
    if not set_a or not set_b:
        return 0.0
    intersection = set_a & set_b
    return len(intersection) / min(len(set_a), len(set_b))


# ---------------------------------------------------------------------------
# 主类
# ---------------------------------------------------------------------------

class TableAnalyzer:
    """表结构分析器

    分析源数据sheets、期望结果结构、规则文本，产出：
    - 主键及来源
    - 主表类型（单表 / 纵向合并 / 横向关联 / 复合）
    - 列分层 (L0-L4)
    - VLOOKUP预计算映射
    - L1列模板代码
    """

    COMMON_KEY_COLUMNS = [
        "工号", "员工工号", "员工编号", "编号", "人员编号", "职工号", "员工号",
        "工作证号", "empno", "emp_no", "employee_no", "employee_id",
        "staff_no", "staff_id",
        "姓名", "员工姓名", "人员姓名", "name", "employee_name",
        "身份证号", "身份证", "证件号", "证件号码", "身份证号码",
        "idcard", "id_card", "id_no",
        "手机号", "手机", "电话", "phone", "mobile",
        "产品编号", "产品代码", "商品编号", "sku",
    ]

    AUXILIARY_KEYWORDS = ["原稿", "原始", "疑问点", "说明", "模板", "目录"]

    def __init__(self):
        self._log: List[str] = []

    def _add_log(self, msg: str):
        self._log.append(msg)
        logger.info(f"[TableAnalyzer] {msg}")

    # ------------------------------------------------------------------
    # 主入口
    # ------------------------------------------------------------------

    def analyze(
        self,
        source_sheets: Dict[str, Dict[str, Any]],
        expected_structure: Dict[str, Any],
        rules_content: str = "",
    ) -> TableAnalysisResult:
        """分析源数据和期望结构，返回完整分析结果"""
        self._log = []
        self._add_log(f"开始分析: {len(source_sheets)} 个源sheet")

        expected_columns = self._extract_expected_columns(expected_structure)
        expected_row_count = self._extract_expected_row_count(expected_structure)
        self._add_log(f"期望输出: {len(expected_columns)} 列, 约 {expected_row_count} 行")

        # 1. 检测主键
        primary_key, pk_source, pk_confidence = self._detect_primary_key(
            source_sheets, expected_columns, rules_content
        )

        # 2. 确定主表
        main_type, main_sheets, join_keys, merge_config = self._determine_main_table(
            source_sheets, expected_columns, expected_row_count,
            primary_key, rules_content
        )

        # 3. 列分层
        column_layers, l1_cols, l2_cols, l3_cols, l4_cols = self._classify_columns(
            primary_key, main_sheets, source_sheets, expected_columns, rules_content
        )

        # 4. VLOOKUP预计算
        vlookup_map = self._build_vlookup_map(
            l2_cols, source_sheets, primary_key, main_sheets
        )

        result = TableAnalysisResult(
            primary_key=primary_key,
            primary_key_source=pk_source,
            main_table_type=main_type,
            main_table_sheets=main_sheets,
            main_table_join_keys=join_keys,
            column_layers=column_layers,
            l1_columns=l1_cols,
            l2_columns=l2_cols,
            l3_columns=l3_cols,
            l4_columns=l4_cols,
            vlookup_map=vlookup_map,
            merge_config=merge_config,
            confidence=pk_confidence,
            detection_log=list(self._log),
        )

        self._add_log(
            f"分析完成: 主键={primary_key}, 主表类型={main_type}, "
            f"L1={len(l1_cols)}列, L2={len(l2_cols)}列, "
            f"L3={len(l3_cols)}列, L4={len(l4_cols)}列, "
            f"VLOOKUP预计算={len(vlookup_map)}条"
        )
        result.detection_log = list(self._log)
        return result

    # ------------------------------------------------------------------
    # 1. 主键检测 (4级级联)
    # ------------------------------------------------------------------

    def _detect_primary_key(
        self,
        source_sheets: Dict[str, Dict[str, Any]],
        expected_columns: List[str],
        rules_content: str,
    ) -> Tuple[str, str, float]:
        """返回 (主键名, 来源sheet, 置信度)"""

        # Tier 1: 规则文档中的 ## 列处理分层
        layer_info = self._parse_layer_info(rules_content)
        if layer_info.get("primary_key"):
            pk = layer_info["primary_key"]
            pk_src = layer_info.get("primary_source", "")
            if pk_src:
                pk_src = self._fuzzy_match_sheet(pk_src, source_sheets) or pk_src
            else:
                pk_src = self._find_key_in_sheets(pk, source_sheets)
            self._add_log(f"Tier1 [列处理分层] 主键={pk}, 来源={pk_src}, 置信度=1.0")
            return pk, pk_src, 1.0

        # Tier 2: 规则文本中的主键声明
        from backend.utils.excel_comparator import extract_primary_keys_from_rules
        declared_keys = extract_primary_keys_from_rules(rules_content, expected_columns)
        if declared_keys:
            pk = declared_keys[0]
            pk_src = self._find_key_in_sheets(pk, source_sheets)
            self._add_log(f"Tier2 [规则声明] 主键={pk}, 来源={pk_src}, 置信度=0.9")
            return pk, pk_src, 0.9

        # Tier 3: 跨sheet启发式检测
        from backend.utils.excel_comparator import detect_primary_keys
        best_pk, best_score, best_src = "", 0, ""
        for sheet_key, sheet_info in source_sheets.items():
            df = sheet_info.get("df")
            if df is None or df.empty:
                continue
            detected = detect_primary_keys(df, max_keys=1)
            if detected:
                col = detected[0]
                non_null = df[col].dropna()
                uniqueness = non_null.nunique() / max(len(non_null), 1)
                score = uniqueness * 100
                for kw in self.COMMON_KEY_COLUMNS:
                    if kw in _normalize_col_name(col) or _normalize_col_name(col) in kw:
                        score += 50
                        break
                if score > best_score:
                    best_pk, best_score, best_src = col, score, sheet_key

        if best_pk:
            self._add_log(f"Tier3 [启发式] 主键={best_pk}, 来源={best_src}, 得分={best_score:.0f}, 置信度=0.7")
            return best_pk, best_src, 0.7

        # Tier 4: 跨sheet列名交集 + COMMON_KEY_COLUMNS
        all_col_sets = []
        for info in source_sheets.values():
            cols = info.get("columns", [])
            if cols:
                all_col_sets.append({_normalize_col_name(c) for c in cols})
        if all_col_sets:
            common = all_col_sets[0]
            for s in all_col_sets[1:]:
                common = common & s
            for kw in self.COMMON_KEY_COLUMNS:
                kw_norm = _normalize_col_name(kw)
                if kw_norm in common:
                    original = self._find_original_col_name(kw_norm, source_sheets)
                    src = self._find_key_in_sheets(original, source_sheets)
                    self._add_log(f"Tier4 [列名交集] 主键={original}, 来源={src}, 置信度=0.5")
                    return original, src, 0.5

        # 最终兜底: expected第一列
        if expected_columns:
            pk = expected_columns[0]
            pk_src = self._find_key_in_sheets(pk, source_sheets)
            self._add_log(f"Tier5 [兜底] 使用期望结果第一列作为主键: {pk}, 置信度=0.3")
            return pk, pk_src, 0.3

        self._add_log("未能检测到主键")
        return "", "", 0.0

    # ------------------------------------------------------------------
    # 2. 主表确定 (支持派生主表)
    # ------------------------------------------------------------------

    def _determine_main_table(
        self,
        source_sheets: Dict[str, Dict[str, Any]],
        expected_columns: List[str],
        expected_row_count: int,
        primary_key: str,
        rules_content: str,
    ) -> Tuple[str, List[str], List[str], Optional[dict]]:
        """返回 (类型, sheet列表, join_keys, merge_config)"""

        # Step A: 规则声明
        layer_info = self._parse_layer_info(rules_content)
        declared_main = layer_info.get("main_table", "")
        if not declared_main:
            m = re.search(r'###\s*主表[:：]\s*(.+?)(?:\n|（|$)', rules_content or "")
            if not m:
                m = re.search(r'###\s*主键来源表[:：]\s*(.+?)(?:\n|（|$)', rules_content or "")
            if m:
                declared_main = m.group(1).strip()

        if declared_main:
            matched = self._fuzzy_match_sheet(declared_main, source_sheets)
            if matched:
                coverage = self._compute_column_coverage(
                    source_sheets[matched], expected_columns
                )
                self._add_log(f"StepA [规则声明] 主表={matched}, 列覆盖率={coverage:.1%}")
                if coverage > 0.3:
                    return "single", [matched], [], None

        # Step B: 单sheet覆盖率
        candidates = []
        for sheet_key, info in source_sheets.items():
            if self._is_auxiliary_sheet(sheet_key):
                continue
            coverage = self._compute_column_coverage(info, expected_columns)
            row_count = len(info.get("df", pd.DataFrame()))
            has_pk = self._sheet_has_column(info, primary_key)
            candidates.append((sheet_key, coverage, row_count, has_pk))

        candidates.sort(key=lambda x: (x[3], x[1], x[2]), reverse=True)
        if candidates and candidates[0][1] >= 0.5:
            best = candidates[0]
            row_ratio = best[2] / max(expected_row_count, 1) if expected_row_count > 0 else 1.0
            self._add_log(
                f"StepB [覆盖率] 主表={best[0]}, 覆盖率={best[1]:.1%}, "
                f"行数={best[2]}, 含主键={best[3]}, 行数比={row_ratio:.2f}"
            )
            # 只在行数足够时接受为单表主表（行数比 >= 0.7 或无期望行数）
            if row_ratio >= 0.7 or expected_row_count == 0:
                return "single", [best[0]], [], None
            else:
                self._add_log(
                    f"StepB 行数不足(比率={row_ratio:.2f}<0.7)，继续检测派生主表"
                )

        # Step C: 纵向合并检测
        concat_groups = self._detect_vertical_concat(
            source_sheets, primary_key, expected_row_count
        )
        # Step D: 横向关联检测
        join_result = self._detect_horizontal_join(
            source_sheets, expected_columns, primary_key
        )

        # 组合判断
        if concat_groups and join_result:
            concat_sheets = concat_groups[0]
            join_sheets, join_key = join_result
            all_sheets = list(set(concat_sheets + join_sheets))
            merge_cfg = {
                "type": "composite",
                "vertical_groups": [concat_sheets],
                "horizontal_joins": [
                    {"left": "derived_main_table", "right": s, "on": join_key}
                    for s in join_sheets if s not in concat_sheets
                ],
                "output_key": "derived_main_table",
            }
            self._add_log(
                f"StepE [复合] 纵向合并={concat_sheets}, "
                f"横向关联={[s for s in join_sheets if s not in concat_sheets]}, "
                f"关联键={join_key}"
            )
            return "composite", all_sheets, [join_key], merge_cfg

        if concat_groups:
            concat_sheets = concat_groups[0]
            merge_cfg = {
                "type": "vertical_concat",
                "vertical_groups": [concat_sheets],
                "horizontal_joins": [],
                "output_key": "derived_main_table",
            }
            self._add_log(f"StepC [纵向合并] sheets={concat_sheets}")
            return "vertical_concat", concat_sheets, [], merge_cfg

        if join_result:
            join_sheets, join_key = join_result
            merge_cfg = {
                "type": "horizontal_join",
                "vertical_groups": [],
                "horizontal_joins": [
                    {"left": join_sheets[0], "right": s, "on": join_key}
                    for s in join_sheets[1:]
                ],
                "output_key": join_sheets[0],
            }
            self._add_log(f"StepD [横向关联] sheets={join_sheets}, 关联键={join_key}")
            return "horizontal_join", join_sheets, [join_key], merge_cfg

        # 兜底: 行数最多且含主键的sheet
        pk_sheets = [c for c in candidates if c[3]]
        if pk_sheets:
            best = max(pk_sheets, key=lambda x: x[2])
            self._add_log(f"StepF [兜底] 主表={best[0]}, 行数={best[2]}")
            return "single", [best[0]], [], None

        if candidates:
            best = candidates[0]
            self._add_log(f"StepG [末选] 主表={best[0]}")
            return "single", [best[0]], [], None

        self._add_log("未能确定主表")
        return "single", [], [], None

    def _detect_vertical_concat(
        self,
        source_sheets: Dict[str, Dict[str, Any]],
        primary_key: str,
        expected_row_count: int,
    ) -> List[List[str]]:
        """检测可纵向合并的sheet组"""
        sheets = [
            (k, v) for k, v in source_sheets.items()
            if not self._is_auxiliary_sheet(k)
        ]
        if len(sheets) < 2:
            return []

        groups = []
        used = set()
        for i, (k1, v1) in enumerate(sheets):
            if k1 in used:
                continue
            cols1 = v1.get("columns", [])
            if not cols1:
                continue
            group = [k1]
            for j, (k2, v2) in enumerate(sheets):
                if j <= i or k2 in used:
                    continue
                cols2 = v2.get("columns", [])
                if _col_overlap_ratio(cols1, cols2) >= 0.85:
                    group.append(k2)
            if len(group) >= 2:
                total_rows = sum(
                    len(source_sheets[k].get("df", pd.DataFrame())) for k in group
                )
                all_have_pk = all(
                    self._sheet_has_column(source_sheets[k], primary_key)
                    for k in group
                )
                if all_have_pk:
                    row_ratio = total_rows / max(expected_row_count, 1)
                    if 0.5 <= row_ratio <= 2.0 or expected_row_count == 0:
                        groups.append(group)
                        used.update(group)
                        self._add_log(
                            f"纵向合并候选: {group}, 合计行数={total_rows}, "
                            f"期望行数={expected_row_count}, 比率={row_ratio:.2f}"
                        )

        return groups

    def _detect_horizontal_join(
        self,
        source_sheets: Dict[str, Dict[str, Any]],
        expected_columns: List[str],
        primary_key: str,
    ) -> Optional[Tuple[List[str], str]]:
        """检测需要横向关联的sheets，返回 (sheet列表, join_key) 或 None"""
        pk_sheets = {}
        for sheet_key, info in source_sheets.items():
            if self._is_auxiliary_sheet(sheet_key):
                continue
            if self._sheet_has_column(info, primary_key):
                cols = {_normalize_col_name(c) for c in info.get("columns", [])}
                pk_sheets[sheet_key] = cols

        if len(pk_sheets) < 2:
            return None

        expected_norm = {_normalize_col_name(c) for c in expected_columns}

        # 贪心集合覆盖
        uncovered = set(expected_norm)
        selected = []
        remaining = dict(pk_sheets)

        while uncovered and remaining:
            best_key = max(remaining.keys(), key=lambda k: len(remaining[k] & uncovered))
            best_gain = len(remaining[best_key] & uncovered)
            if best_gain == 0:
                break
            selected.append(best_key)
            uncovered -= remaining[best_key]
            del remaining[best_key]

        if len(selected) >= 2:
            total_coverage = 1.0 - len(uncovered) / max(len(expected_norm), 1)
            self._add_log(
                f"横向关联候选: {selected}, 联合覆盖率={total_coverage:.1%}"
            )
            if total_coverage >= 0.6:
                return selected, primary_key

        return None

    # ------------------------------------------------------------------
    # 3. 列分层
    # ------------------------------------------------------------------

    def _classify_columns(
        self,
        primary_key: str,
        main_sheets: List[str],
        source_sheets: Dict[str, Dict[str, Any]],
        expected_columns: List[str],
        rules_content: str,
    ) -> Tuple[Dict[str, str], List[ColumnMapping], List[ColumnMapping], List[str], List[str]]:
        """返回 (column_layers, l1_cols, l2_cols, l3_cols, l4_cols)"""

        column_layers: Dict[str, str] = {}
        l1_cols: List[ColumnMapping] = []
        l2_cols: List[ColumnMapping] = []
        l3_cols: List[str] = []
        l4_cols: List[str] = []

        # 优先使用规则中的分层信息
        layer_info = self._parse_layer_info(rules_content)
        rules_layer_map = layer_info.get("column_layer", {})

        # 构建主表列名集合
        main_cols_map = {}  # normalized -> (sheet_key, original_name, index)
        for sheet_key in main_sheets:
            info = source_sheets.get(sheet_key, {})
            for idx, col in enumerate(info.get("columns", [])):
                norm = _normalize_col_name(col)
                if norm not in main_cols_map:
                    main_cols_map[norm] = (sheet_key, col, idx)

        # 构建所有非主表的列名集合
        other_cols_map = {}  # normalized -> (sheet_key, original_name, index)
        for sheet_key, info in source_sheets.items():
            if sheet_key in main_sheets or self._is_auxiliary_sheet(sheet_key):
                continue
            for idx, col in enumerate(info.get("columns", [])):
                norm = _normalize_col_name(col)
                if norm not in other_cols_map:
                    other_cols_map[norm] = (sheet_key, col, idx)

        for col in expected_columns:
            col_norm = _normalize_col_name(col)

            # 检查规则中是否有明确声明
            declared_layer = None
            for rule_col, rule_layer in rules_layer_map.items():
                if _normalize_col_name(rule_col) == col_norm:
                    declared_layer = rule_layer
                    break

            if declared_layer:
                column_layers[col] = declared_layer
                if declared_layer == "L0":
                    pass  # 主键，不加入l1/l2
                elif declared_layer == "L1":
                    src = main_cols_map.get(col_norm)
                    if src:
                        l1_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                elif declared_layer == "L2":
                    src = other_cols_map.get(col_norm)
                    if src:
                        l2_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                    else:
                        src = main_cols_map.get(col_norm)
                        if src:
                            l2_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                elif declared_layer in ("L3", "L4"):
                    if declared_layer == "L3":
                        l3_cols.append(col)
                    else:
                        l4_cols.append(col)
                continue

            # 自动分层
            pk_norm = _normalize_col_name(primary_key)
            if col_norm == pk_norm or col_norm in pk_norm or pk_norm in col_norm:
                column_layers[col] = "L0"
                continue

            if col_norm in main_cols_map:
                column_layers[col] = "L1"
                src = main_cols_map[col_norm]
                l1_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                continue

            if col_norm in other_cols_map:
                column_layers[col] = "L2"
                src = other_cols_map[col_norm]
                l2_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                continue

            # 模糊匹配: 子串包含
            found = False
            for norm_key, src in main_cols_map.items():
                if col_norm in norm_key or norm_key in col_norm:
                    column_layers[col] = "L1"
                    l1_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                    found = True
                    break
            if found:
                continue

            for norm_key, src in other_cols_map.items():
                if col_norm in norm_key or norm_key in col_norm:
                    column_layers[col] = "L2"
                    l2_cols.append(ColumnMapping(col, src[0], src[1], src[2]))
                    found = True
                    break
            if found:
                continue

            column_layers[col] = "L3"
            l3_cols.append(col)

        self._add_log(
            f"列分层: L0=1(主键), L1={len(l1_cols)}, L2={len(l2_cols)}, "
            f"L3={len(l3_cols)}, L4={len(l4_cols)}"
        )
        return column_layers, l1_cols, l2_cols, l3_cols, l4_cols

    # ------------------------------------------------------------------
    # 4. VLOOKUP预计算
    # ------------------------------------------------------------------

    def _build_vlookup_map(
        self,
        l2_columns: List[ColumnMapping],
        source_sheets: Dict[str, Dict[str, Any]],
        primary_key: str,
        main_sheets: List[str],
    ) -> Dict[str, VLookupInfo]:
        """为每个L2列预计算VLOOKUP参数"""
        result = {}
        for mapping in l2_columns:
            info = source_sheets.get(mapping.source_sheet, {})
            columns = info.get("columns", [])
            if not columns:
                continue

            # 找主键在源sheet中的位置
            pk_idx = self._find_column_index(columns, primary_key)
            if pk_idx < 0:
                self._add_log(
                    f"VLOOKUP警告: {mapping.source_sheet} 中未找到主键 {primary_key}"
                )
                continue

            target_idx = mapping.col_index_in_source
            col_num = target_idx - pk_idx + 1
            if col_num <= 0:
                self._add_log(
                    f"VLOOKUP警告: {mapping.col_name} 的col_num={col_num}<=0 "
                    f"(target_idx={target_idx}, pk_idx={pk_idx}), "
                    f"需要INDEX+MATCH而非VLOOKUP"
                )
                continue

            pk_letter = _col_index_to_letter(pk_idx)
            target_letter = _col_index_to_letter(target_idx)
            last_letter = _col_index_to_letter(len(columns) - 1)
            range_str = f"${pk_letter}:${last_letter}"

            formula = (
                f"=IFERROR(VLOOKUP({{key}},'{mapping.source_sheet}'!"
                f"{range_str},{col_num},FALSE),0)"
            )

            result[mapping.col_name] = VLookupInfo(
                target_col=mapping.col_name,
                source_sheet=mapping.source_sheet,
                key_col_letter=pk_letter,
                target_col_letter=target_letter,
                col_num=col_num,
                range_str=range_str,
                formula_template=formula,
            )

        self._add_log(f"VLOOKUP预计算完成: {len(result)} 个L2列")
        return result

    # ------------------------------------------------------------------
    # 5. L1模板代码生成
    # ------------------------------------------------------------------

    def generate_l1_code(
        self,
        analysis: 'TableAnalysisResult',
        expected_structure: Dict[str, Any],
    ) -> str:
        """为L1列生成模板代码（write_cell调用）

        返回可直接注入到 fill_result_sheets for循环体中的Python代码。
        """
        if not analysis.l1_columns:
            return ""

        # 从expected_structure获取目标列顺序，确定每列在结果sheet中的列号
        target_columns = self._extract_expected_columns(expected_structure)
        target_col_index = {}  # normalized -> 1-based column index
        for idx, col in enumerate(target_columns):
            target_col_index[_normalize_col_name(col)] = idx + 1

        lines = ["        # === L1 同源列（自动生成，禁止修改）==="]
        for mapping in analysis.l1_columns:
            col_norm = _normalize_col_name(mapping.col_name)
            col_idx = target_col_index.get(col_norm)
            if col_idx is None:
                continue

            src_col = mapping.source_col_name
            line = (
                f"        write_cell(ws, r, {col_idx}, "
                f"main_df.iloc[i].get('{src_col}', ''))"
            )
            lines.append(f"        # {col_idx}列: {mapping.col_name}")
            lines.append(line)

        lines.append("        # === L1 同源列结束 ===")
        code = "\n".join(lines)
        self._add_log(f"L1模板代码: {len(analysis.l1_columns)} 列")
        return code

    # ------------------------------------------------------------------
    # VLOOKUP速查表（用于prompt注入）
    # ------------------------------------------------------------------

    def generate_vlookup_table(self, analysis: 'TableAnalysisResult') -> str:
        """生成markdown格式的VLOOKUP列号速查表"""
        if not analysis.vlookup_map:
            return ""

        # 按源sheet分组
        by_sheet: Dict[str, List[VLookupInfo]] = {}
        for info in analysis.vlookup_map.values():
            by_sheet.setdefault(info.source_sheet, []).append(info)

        lines = ["## VLOOKUP列号速查表（已预计算，直接使用，禁止自行计算列号）\n"]

        for sheet, infos in by_sheet.items():
            pk_letter = infos[0].key_col_letter if infos else "A"
            lines.append(
                f"### {sheet}（主键列: {pk_letter}列, 范围: {infos[0].range_str}）"
            )
            lines.append("| 目标列名 | 源列字母 | VLOOKUP列号(第3参数) | 公式模板 |")
            lines.append("|---------|---------|---------------------|---------|")
            for info in sorted(infos, key=lambda x: x.col_num):
                lines.append(
                    f"| {info.target_col} | {info.target_col_letter} | "
                    f"{info.col_num} | `{info.formula_template}` |"
                )
            lines.append("")

        return "\n".join(lines)

    # ------------------------------------------------------------------
    # 分层信息摘要（用于prompt注入）
    # ------------------------------------------------------------------

    def generate_layer_summary(self, analysis: 'TableAnalysisResult') -> str:
        """生成分层信息摘要，供prompt使用"""
        lines = ["## 表结构分析结果（已预分析，请严格遵循）\n"]
        lines.append(f"- **主键**: {analysis.primary_key}（来源: {analysis.primary_key_source}）")
        lines.append(
            f"- **主表**: {', '.join(analysis.main_table_sheets)} "
            f"(类型: {analysis.main_table_type})"
        )

        if analysis.main_table_type != "single":
            lines.append(f"  - 合并方式: {analysis.main_table_type}")
            if analysis.main_table_join_keys:
                lines.append(f"  - 关联键: {', '.join(analysis.main_table_join_keys)}")

        l1_names = [c.col_name for c in analysis.l1_columns]
        l2_names = [c.col_name for c in analysis.l2_columns]

        lines.append(f"- **L0 主键** (1列): {analysis.primary_key}")
        lines.append(
            f"- **L1 同源列** ({len(l1_names)}列，已自动生成代码，你不需要处理): "
            f"{', '.join(l1_names[:10])}"
            + (f" ... 等共{len(l1_names)}列" if len(l1_names) > 10 else "")
        )
        lines.append(
            f"- **L2 跨表列** ({len(l2_names)}列，使用速查表中的VLOOKUP): "
            f"{', '.join(l2_names[:10])}"
            + (f" ... 等共{len(l2_names)}列" if len(l2_names) > 10 else "")
        )
        lines.append(f"- **L3 计算列** ({len(analysis.l3_columns)}列): {', '.join(analysis.l3_columns[:10])}")
        if analysis.l4_columns:
            lines.append(f"- **L4 复合列** ({len(analysis.l4_columns)}列): {', '.join(analysis.l4_columns[:10])}")

        ai_cols = 1 + len(l2_names) + len(analysis.l3_columns) + len(analysis.l4_columns)
        lines.append(f"\n**你需要生成的列数: {ai_cols}** (L0主键 + L2跨表 + L3计算 + L4复合)")
        lines.append(f"L1同源列已自动处理，不要重复生成。")
        lines.append("")
        return "\n".join(lines)

    # ------------------------------------------------------------------
    # 辅助方法
    # ------------------------------------------------------------------

    def _parse_layer_info(self, rules_content: str) -> dict:
        """解析规则中的列处理分层信息（复用prompt_generator的逻辑）"""
        if not rules_content or "## 列处理分层" not in rules_content:
            return {}

        result = {
            "primary_key": "",
            "primary_source": "",
            "main_table": "",
            "L1": [], "L2": [], "L3": [], "L4": [],
            "column_layer": {},
        }

        start = rules_content.index("## 列处理分层")
        rest = rules_content[start + len("## 列处理分层"):]
        next_section = re.search(r'\n## [^#]', rest)
        section_text = rest[:next_section.start()] if next_section else rest

        pk_match = re.search(r'###\s*主键[:：]\s*(.+)', section_text)
        if pk_match:
            result["primary_key"] = pk_match.group(1).strip()
            result["column_layer"][result["primary_key"]] = "L0"

        src_match = re.search(r'###\s*主键来源表[:：]\s*(.+)', section_text)
        if src_match:
            result["primary_source"] = src_match.group(1).strip()

        main_match = re.search(r'###\s*主表[:：]\s*(.+?)(?:\n|（|$)', section_text)
        if main_match:
            result["main_table"] = main_match.group(1).strip()
        elif result["primary_source"]:
            result["main_table"] = result["primary_source"]

        for layer in ["L1", "L2", "L3", "L4"]:
            pattern = rf'###\s*{layer}[-\s].*?\n((?:[-\s]*.*\n)*?)(?=###|\Z)'
            layer_match = re.search(pattern, section_text)
            if layer_match:
                block = layer_match.group(1)
                for line in block.split('\n'):
                    line = line.strip()
                    if line.startswith('-'):
                        col_text = line.lstrip('- ').strip()
                        col_name = re.split(r'[（(]', col_text)[0].strip()
                        if col_name:
                            result[layer].append(col_name)
                            result["column_layer"][col_name] = layer

        return result

    def _extract_expected_columns(self, expected_structure: Dict[str, Any]) -> List[str]:
        """从期望结构中提取列名列表"""
        if not expected_structure:
            return []
        sheets = expected_structure.get("sheets", [])
        if not sheets:
            return []
        first_sheet = sheets[0] if isinstance(sheets, list) else list(sheets.values())[0]
        if isinstance(first_sheet, dict):
            return first_sheet.get("columns", first_sheet.get("headers", []))
        return []

    def _extract_expected_row_count(self, expected_structure: Dict[str, Any]) -> int:
        """从期望结构中提取行数"""
        if not expected_structure:
            return 0
        sheets = expected_structure.get("sheets", [])
        if not sheets:
            return 0
        first_sheet = sheets[0] if isinstance(sheets, list) else list(sheets.values())[0]
        if isinstance(first_sheet, dict):
            return first_sheet.get("row_count", first_sheet.get("data_rows", 0))
        return 0

    def _fuzzy_match_sheet(self, name: str, source_sheets: Dict) -> Optional[str]:
        """模糊匹配sheet名称"""
        name_norm = _normalize_col_name(name)
        for key in source_sheets:
            key_norm = _normalize_col_name(key)
            if name_norm == key_norm or name_norm in key_norm or key_norm in name_norm:
                return key
        return None

    def _find_key_in_sheets(self, key_name: str, source_sheets: Dict) -> str:
        """找到包含指定主键列的第一个sheet"""
        for sheet_key, info in source_sheets.items():
            if self._sheet_has_column(info, key_name):
                return sheet_key
        return ""

    def _sheet_has_column(self, info: Dict, col_name: str) -> bool:
        """检查sheet是否包含指定列"""
        if not col_name:
            return False
        col_norm = _normalize_col_name(col_name)
        for c in info.get("columns", []):
            c_norm = _normalize_col_name(c)
            if col_norm == c_norm or col_norm in c_norm or c_norm in col_norm:
                return True
        return False

    def _compute_column_coverage(self, info: Dict, expected_columns: List[str]) -> float:
        """计算sheet列对期望列的覆盖率"""
        if not expected_columns:
            return 0.0
        sheet_cols = {_normalize_col_name(c) for c in info.get("columns", [])}
        matched = 0
        for col in expected_columns:
            col_norm = _normalize_col_name(col)
            if col_norm in sheet_cols:
                matched += 1
                continue
            for sc in sheet_cols:
                if col_norm in sc or sc in col_norm:
                    matched += 1
                    break
        return matched / len(expected_columns)

    def _is_auxiliary_sheet(self, sheet_key: str) -> bool:
        """检查是否为辅助sheet（非数据sheet）"""
        key_lower = sheet_key.lower()
        return any(kw in key_lower for kw in self.AUXILIARY_KEYWORDS)

    def _find_column_index(self, columns: List[str], col_name: str) -> int:
        """在列名列表中查找列的索引，返回0-based索引或-1"""
        col_norm = _normalize_col_name(col_name)
        for idx, c in enumerate(columns):
            c_norm = _normalize_col_name(c)
            if col_norm == c_norm:
                return idx
        for idx, c in enumerate(columns):
            c_norm = _normalize_col_name(c)
            if col_norm in c_norm or c_norm in col_norm:
                return idx
        return -1

    def _find_original_col_name(self, normalized: str, source_sheets: Dict) -> str:
        """从normalized名查找原始列名"""
        for info in source_sheets.values():
            for c in info.get("columns", []):
                if _normalize_col_name(c) == normalized:
                    return c
        return normalized
