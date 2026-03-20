"""
Excel智能解析器 - Python实现
基于C# ExcelParser的功能移植

使用方法:
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file("your_file.xlsx")
    
    for sheet_data in results:
        print(f"Sheet: {sheet_data.sheet_name}")
        for region in sheet_data.regions:
            print(f"  表头: {region.head_data}")
            print(f"  数据行数: {len(region.data)}")
"""

import re
import os
import logging
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Callable
from enum import Enum
from pathlib import Path
from datetime import datetime

# ==================== Aspose.Cells for .NET 初始化 ====================
# 通过 pythonnet 调用 Aspose.Cells.dll（.NET 版本），替代 aspose-cells-python

_libs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "libs")
os.add_dll_directory(_libs_dir)

import pythonnet
pythonnet.load("coreclr", runtime_config=os.path.join(_libs_dir, "runtimeconfig.json"))
import clr
clr.AddReference(os.path.join(_libs_dir, "SkiaSharp.dll"))
clr.AddReference(os.path.join(_libs_dir, "Aspose.Cells.dll"))
clr.AddReference("System.Text.Encoding.CodePages")
import System.Text
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

from Aspose.Cells import (
    License as _AsposeLicense,
    Workbook as _AsposeWorkbook,
    BorderType as _BorderType,
    CellBorderType as _CellBorderType,
    BackgroundType as _BackgroundType,
)

# 设置许可证（避免评估模式水印）
_lic_path = os.path.join(_libs_dir, "Aspose.Total.NET.lic")
if os.path.exists(_lic_path):
    _lic = _AsposeLicense()
    _lic.SetLicense(_lic_path)


# ==================== Aspose.Cells 适配器层 ====================
# 将 Aspose Worksheet 包装为兼容原有业务逻辑的接口（1-indexed row/col）

class _AsposeMergedRange:
    """适配 openpyxl CellRange 接口（1-indexed）"""
    __slots__ = ['min_row', 'max_row', 'min_col', 'max_col']

    def __init__(self, aspose_range):
        # .NET CellArea 是 0-indexed，转换为 1-indexed
        self.min_row = aspose_range.StartRow + 1
        self.max_row = aspose_range.EndRow + 1
        self.min_col = aspose_range.StartColumn + 1
        self.max_col = aspose_range.EndColumn + 1


class _AsposeMergedCells:
    """适配 openpyxl merged_cells.ranges 接口"""
    def __init__(self, ws):
        raw = ws.Cells.MergedCells   # .NET ArrayList of CellArea
        self.ranges = [_AsposeMergedRange(raw[i]) for i in range(raw.Count)]


class _AsposeFontProxy:
    """适配 openpyxl font 接口"""
    __slots__ = ['bold']
    def __init__(self, is_bold):
        self.bold = is_bold


class _AsposeFillProxy:
    """适配 openpyxl fill 接口（用于检测背景色）"""
    __slots__ = ['patternType', 'fgColor']
    def __init__(self, has_bg):
        self.patternType = 'solid' if has_bg else None
        self.fgColor = type('_Color', (), {'rgb': 'FF000000' if has_bg else '00000000'})()


class _AsposeBorderProxy:
    """适配 openpyxl border 接口"""
    __slots__ = ['left', 'right', 'top', 'bottom']
    def __init__(self, left, right, top, bottom):
        _S = type('_Side', (), {'style': None})
        self.left   = _S(); self.left.style   = left
        self.right  = _S(); self.right.style  = right
        self.top    = _S(); self.top.style    = top
        self.bottom = _S(); self.bottom.style = bottom


class _AsposeCell:
    """适配 openpyxl Cell 接口，封装 Aspose .NET Cell"""
    __slots__ = ['_cell', '_row', '_col', '_style', '_style_loaded',
                 'value', 'row', 'column', 'formula']

    def __init__(self, aspose_cell, row_1idx, col_1idx):
        self._cell = aspose_cell
        self._row = row_1idx
        self._col = col_1idx
        self._style = None
        self._style_loaded = False
        # 读取公式（.NET 版 Formula 为 None 或字符串）
        raw_formula = aspose_cell.Formula
        self.formula = raw_formula if raw_formula else None
        # 读取值：有公式时 value 存公式字符串（兼容原有 startswith('=') 检测）
        if self.formula:
            self.value = self.formula
        else:
            raw = aspose_cell.Value
            if raw is None or (isinstance(raw, str) and raw == ''):
                self.value = None
            else:
                self.value = raw
        self.row = row_1idx
        self.column = col_1idx

    def _load_style(self):
        if not self._style_loaded:
            self._style = self._cell.GetStyle()
            self._style_loaded = True

    @property
    def coordinate(self):
        col_letter = ''
        c = self._col
        while c > 0:
            c, r = divmod(c - 1, 26)
            col_letter = chr(65 + r) + col_letter
        return f"{col_letter}{self._row}"

    @property
    def font(self):
        self._load_style()
        return _AsposeFontProxy(self._style.Font.IsBold)

    @property
    def fill(self):
        self._load_style()
        has_bg = (self._style.Pattern == _BackgroundType.Solid and
                  self._style.ForegroundColor.ToArgb() != 0)
        return _AsposeFillProxy(has_bg)

    @property
    def border(self):
        self._load_style()
        def _has(bt):
            ls = self._style.Borders[bt].LineStyle
            return 'thin' if int(ls) != 0 else None
        return _AsposeBorderProxy(
            _has(_BorderType.LeftBorder),
            _has(_BorderType.RightBorder),
            _has(_BorderType.TopBorder),
            _has(_BorderType.BottomBorder),
        )


class _AsposeWorksheet:
    """适配 openpyxl Worksheet 接口，封装 Aspose .NET Worksheet"""

    def __init__(self, aspose_ws):
        self._ws = aspose_ws
        self.title = aspose_ws.Name
        # max_row / max_column：.NET 0-indexed → 1-indexed
        mr = aspose_ws.Cells.MaxDataRow
        mc = aspose_ws.Cells.MaxDataColumn
        self.max_row = (mr + 1) if mr >= 0 else 0
        self.max_column = (mc + 1) if mc >= 0 else 0
        self.merged_cells = _AsposeMergedCells(aspose_ws)
        self._cell_cache: Dict = {}

    def cell(self, row=None, column=None):
        """1-indexed 行列访问（兼容 openpyxl 接口）"""
        key = (row, column)
        cached = self._cell_cache.get(key)
        if cached is not None:
            return cached
        # .NET Cells 用 0-indexed 索引器
        ac = self._ws.Cells[row - 1, column - 1]
        c = _AsposeCell(ac, row, column)
        self._cell_cache[key] = c
        return c


# ==================== 数据类定义 ====================

@dataclass
class ExcelRegion:
    """Excel区域数据结构"""
    head_row_start: int = 0
    head_row_end: int = 0
    data_row_start: int = 0
    data_row_end: int = 0
    head_data: Dict[str, str] = field(default_factory=dict)
    data: List[Dict[str, Any]] = field(default_factory=list)
    formula: Dict[str, str] = field(default_factory=dict)


@dataclass
class SheetData:
    """Sheet数据结构"""
    sheet_name: str = ""
    regions: List[ExcelRegion] = field(default_factory=list)


@dataclass
class RowContext:
    """行上下文信息"""
    row_index: int = 0
    worksheet: Optional[Any] = None
    max_col: int = 0
    non_empty_count: int = 0
    text_count: int = 0
    number_count: int = 0
    has_merged_cells: bool = False
    has_special_formatting: bool = False


@dataclass
class HeaderRule:
    """表头规则定义"""
    name: str
    evaluator: Callable[[Any, RowContext], float]
    weight: float


class ValueType(Enum):
    """值类型枚举"""
    EMPTY = 0
    TEXT = 1
    NUMBER = 2
    DATE = 3
    OTHER = 4


class RowType(Enum):
    """行类型枚举"""
    HEADER = 0
    DATA = 1
    SUMMARY = 2
    UNKNOWN = 3


@dataclass
class HeaderInfo:
    """表头信息"""
    start_row: int = 0
    end_row: int = 0


@dataclass
class RowFeatures:
    """行特征分析结果"""
    row_index: int = 0
    non_empty_count: int = 0
    text_count: int = 0
    number_count: int = 0
    date_count: int = 0
    number_ratio: float = 0.0
    text_ratio: float = 0.0
    has_sequence_number: bool = False  # 是否有序号
    sequence_value: int = -1  # 序号值
    has_merged_cells: bool = False
    has_special_formatting: bool = False
    keyword_count: int = 0
    keyword_ratio: float = 0.0
    avg_text_length: float = 0.0  # 平均文本长度
    format_signature: str = ""  # 格式签名（用于比较格式一致性）


# ==================== 表头规则引擎 ====================

class HeaderRuleEngine:
    """表头识别规则引擎"""
    
    # 静态关键字集合
    HEADER_KEYWORDS = {
        "序号", "姓名", "名称", "编码", "证件号", "身份证", "账单月份", "户口性质",
        "基数", "比例", "金额", "公司", "个人", "合计", "备注", "工号", "外服编号",
        "城市", "年月", "日期", "社保", "公积金", "养老", "医疗", "失业", "工伤", "生育",
        "大病", "采暖费", "工会费", "滞纳金", "管理费", "客户", "商社", "单位",
        "员工编号", "员工性质", "外派补贴", "推荐奖金", "区域激励", "产假工资", "经济补偿",
        "税前调整", "税后调整", "年假工资", "交通补贴", "住房补贴", "奖金", "补贴", "工资",
        "employment", "insurance", "provident", "fund", "medical", "endowment",
        "unemployment", "maternity", "injury", "name", "date", "amount", "base",
        "employer", "employee", "subtotal", "total", "remarks", "id", "card",
        "ratio", "小计", "总计"
    }
    
    def __init__(self):
        self.rules: List[HeaderRule] = []
        self._merged_cell_index = None  # 【性能优化】外部注入的合并单元格索引
        self._initialize_default_rules()
    
    def _initialize_default_rules(self):
        """初始化默认规则"""
        
        # 1. 关键字匹配规则
        def keyword_rule(cell_value: Any, context: RowContext) -> float:
            if isinstance(cell_value, str):
                text_lower = cell_value.lower()
                return 1.0 if any(k.lower() in text_lower for k in self.HEADER_KEYWORDS) else 0.0
            return 0.0
        
        self.add_rule("KeywordRule", keyword_rule, 0.8)
        
        # 2. 数据类型规则（文本为主）
        def text_dominance_rule(cell_value: Any, context: RowContext) -> float:
            if context.non_empty_count == 0:
                return 0.0
            return context.text_count / context.non_empty_count
        
        self.add_rule("TextDominanceRule", text_dominance_rule, 0.7)
        
        # 3. 合并单元格规则
        def merged_cell_rule(cell_value: Any, context: RowContext) -> float:
            return 1.0 if context.has_merged_cells else 0.0
        
        self.add_rule("MergedCellRule", merged_cell_rule, 0.6)
        
        # 4. 位置规则（通常表头在前几行）
        def position_rule(cell_value: Any, context: RowContext) -> float:
            return max(0, 1.0 - (context.row_index / 20.0))
        
        self.add_rule("PositionRule", position_rule, 0.5)
        
        # 5. 格式规则（加粗、背景色等）
        def format_rule(cell_value: Any, context: RowContext) -> float:
            return 0.8 if context.has_special_formatting else 0.0
        
        self.add_rule("FormatRule", format_rule, 0.5)
    
    def add_rule(self, name: str, evaluator: Callable[[Any, RowContext], float], weight: float):
        """添加自定义规则"""
        self.rules.append(HeaderRule(name=name, evaluator=evaluator, weight=weight))
    
    def calculate_header_score(self, worksheet: Any, row: int, max_col: int) -> float:
        """计算行的表头得分"""
        context = RowContext(row_index=row, worksheet=worksheet, max_col=max_col)
        self.analyze_row_context(worksheet, row, max_col, context)
        
        total_score = 0.0
        total_weight = 0.0
        
        for rule in self.rules:
            rule_score = 0.0
            for col in range(1, max_col + 1):
                cell_value = self._get_cell_value(worksheet.cell(row, col))
                if cell_value is not None and str(cell_value).strip():
                    rule_score = max(rule_score, rule.evaluator(cell_value, context))
            
            total_score += rule_score * rule.weight
            total_weight += rule.weight
        
        return total_score / total_weight if total_weight > 0 else 0.0
    
    def analyze_row_context(self, worksheet: Any, row: int, max_col: int, context: RowContext):
        """分析行的上下文信息"""
        context.non_empty_count = 0
        context.text_count = 0
        context.number_count = 0
        context.has_merged_cells = False
        context.has_special_formatting = False
        
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row, col)
            value = self._get_cell_value(cell)
            
            if value is not None and str(value).strip():
                context.non_empty_count += 1
                
                if isinstance(value, str):
                    context.text_count += 1
                elif self._is_numeric(value):
                    context.number_count += 1
                
                if not context.has_merged_cells:
                    if self._is_merged_cell(worksheet, row, col):
                        context.has_merged_cells = True
                
                if not context.has_special_formatting:
                    try:
                        if (cell.font and cell.font.bold) or \
                           (cell.fill and cell.fill.fgColor and str(cell.fill.fgColor.rgb) != '00000000') or \
                           (cell.alignment and cell.alignment.horizontal == 'center'):
                            context.has_special_formatting = True
                    except:
                        pass
    
    @staticmethod
    def _get_cell_value(cell: Any) -> Any:
        """获取单元格值"""
        if cell.value is None:
            return None
        if isinstance(cell.value, datetime):
            return cell.value.strftime("%Y-%m-%d")
        elif isinstance(cell.value, (int, float)):
            return cell.value
        else:
            return str(cell.value)
    
    @staticmethod
    def _is_numeric(value: Any) -> bool:
        """判断是否为数值类型"""
        return isinstance(value, (int, float, complex))
    
    def _is_merged_cell(self, worksheet: Any, row: int, col: int) -> bool:
        """检查是否为合并单元格（使用索引优化）"""
        if self._merged_cell_index is not None:
            return (row, col) in self._merged_cell_index
        cell = worksheet.cell(row, col)
        for merged_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return True
        return False


# ==================== 列一致性验证器 ====================

class ColumnConsistencyValidator:
    """列一致性验证器 - 验证表头和数据的一致性"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def validate_region(self, worksheet: Any, region: 'ExcelRegion', max_col: int) -> Dict[str, Any]:
        """验证区域的列一致性

        Returns:
            {
                'score': float,  # 0-1之间的得分
                'issues': List[str],  # 发现的问题列表
                'column_scores': Dict[str, float]  # 每列的得分
            }
        """
        if not region.data or len(region.data) == 0:
            return {'score': 0.0, 'issues': ['没有数据行'], 'column_scores': {}}

        issues = []
        column_scores = {}

        # 检查每一列的类型一致性
        for header, col_letter in region.head_data.items():
            col_num = self._get_column_number(col_letter)

            # 分析表头暗示的类型
            expected_type = self._infer_type_from_header(header)

            # 分析数据列的实际类型
            actual_types = []
            for data_row in region.data[:min(20, len(region.data))]:  # 检查前20行
                value = data_row.get(col_letter)
                if value is not None and str(value).strip():
                    actual_types.append(self._get_value_type(value))

            if not actual_types:
                column_scores[header] = 0.5  # 空列给中等分数
                continue

            # 计算类型匹配度
            match_score = self._calculate_type_match(expected_type, actual_types)
            column_scores[header] = match_score

            if match_score < 0.5:
                issues.append(f"列'{header}'类型不匹配: 期望{expected_type}, 实际类型分布不符")

        # 检查数据完整性（非空比例）
        completeness_scores = []
        for data_row in region.data:
            non_empty = sum(1 for v in data_row.values() if v is not None and str(v).strip())
            completeness_scores.append(non_empty / len(region.head_data) if region.head_data else 0)

        avg_completeness = sum(completeness_scores) / len(completeness_scores) if completeness_scores else 0

        if avg_completeness < 0.3:
            issues.append(f"数据完整性低: 平均只有{avg_completeness*100:.1f}%的列有值")

        # 综合得分
        type_score = sum(column_scores.values()) / len(column_scores) if column_scores else 0
        overall_score = (type_score * 0.7 + avg_completeness * 0.3)

        return {
            'score': overall_score,
            'issues': issues,
            'column_scores': column_scores,
            'completeness': avg_completeness
        }

    def _infer_type_from_header(self, header: str) -> str:
        """从表头推断期望的数据类型"""
        header_lower = header.lower()

        # 数值类型关键字
        number_keywords = ['金额', '数量', '比例', '基数', '工资', '奖金', '补贴', '费用',
                          '合计', '小计', '总计', '天数', '次数', '人数', 'amount', 'salary',
                          'bonus', 'total', 'count', 'rate', 'ratio', 'price', 'cost']

        # 日期类型关键字
        date_keywords = ['日期', '时间', '年月', '月份', 'date', 'time', 'month', 'year']

        # 文本类型关键字
        text_keywords = ['姓名', '名称', '部门', '职位', '地址', '备注', '说明',
                        'name', 'department', 'position', 'address', 'remark', 'note']

        for keyword in number_keywords:
            if keyword in header_lower:
                return 'number'

        for keyword in date_keywords:
            if keyword in header_lower:
                return 'date'

        for keyword in text_keywords:
            if keyword in header_lower:
                return 'text'

        return 'mixed'  # 未知类型

    def _get_value_type(self, value: Any) -> str:
        """获取值的类型"""
        if isinstance(value, (int, float)):
            return 'number'
        elif isinstance(value, datetime):
            return 'date'
        elif isinstance(value, str):
            # 尝试判断是否是数字字符串
            try:
                float(value.replace(',', '').replace('%', ''))
                return 'number'
            except:
                return 'text'
        return 'other'

    def _calculate_type_match(self, expected_type: str, actual_types: List[str]) -> float:
        """计算类型匹配度"""
        if expected_type == 'mixed':
            return 0.8  # 未知类型给较高分数

        if not actual_types:
            return 0.5

        # 计算期望类型的比例
        match_count = sum(1 for t in actual_types if t == expected_type)
        match_ratio = match_count / len(actual_types)

        return match_ratio

    @staticmethod
    def _get_column_number(column_letter: str) -> int:
        """获取列编号"""
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


# ==================== 后验证器 ====================

class PostValidator:
    """后验证器 - 对识别结果进行验证和自动修正"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.consistency_validator = ColumnConsistencyValidator()

    def validate_and_fix(self, worksheet: Any, region: 'ExcelRegion',
                        max_col: int) -> tuple['ExcelRegion', List[str]]:
        """验证并尝试修正区域

        Returns:
            (修正后的region, 修正日志)
        """
        fixes = []

        # 1. 列一致性验证
        consistency_result = self.consistency_validator.validate_region(worksheet, region, max_col)

        if consistency_result['score'] < 0.5:
            self.logger.warning(f"列一致性得分低: {consistency_result['score']:.2f}")
            fixes.append(f"列一致性得分: {consistency_result['score']:.2f}")

            # 尝试修正：可能表头识别过多
            if len(region.data) < 3 and region.head_row_end > region.head_row_start:
                fixes.append("尝试减少表头行数")
                # 这里可以尝试调整表头范围，但需要重新解析

        # 2. 数据行数验证
        if len(region.data) < 2:
            fixes.append(f"数据行数过少: {len(region.data)}")

        # 3. 表头完整性验证
        empty_headers = sum(1 for h in region.head_data.keys() if not h or h.startswith('Column_'))
        if empty_headers > len(region.head_data) * 0.3:
            fixes.append(f"表头不完整: {empty_headers}/{len(region.head_data)}列缺少表头")

        return region, fixes


# ==================== 列关联性分析器 ====================

class ColumnRelationAnalyzer:
    """列关联性分析器 - 分析列之间的语义关联"""

    # 列关联规则：某些列通常一起出现
    COLUMN_RELATIONS = {
        '姓名': ['工号', '部门', '职位', '员工编号'],
        '工号': ['姓名', '部门'],
        '金额': ['日期', '类型', '备注'],
        '基数': ['比例', '金额'],
        '日期': ['金额', '类型'],
        '证件号': ['姓名', '身份证'],
        '社保': ['养老', '医疗', '失业', '工伤', '生育'],
        '公积金': ['基数', '比例', '金额'],
    }

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def analyze_header_relations(self, headers: List[str]) -> float:
        """分析表头之间的关联性

        Returns:
            关联性得分 (0-1)
        """
        if len(headers) < 2:
            return 0.5

        relation_score = 0
        total_checks = 0

        for header in headers:
            # 查找相关列
            related_keywords = self._find_related_keywords(header)
            if not related_keywords:
                continue

            # 检查是否有相关列存在
            for related in related_keywords:
                total_checks += 1
                if any(related.lower() in h.lower() for h in headers):
                    relation_score += 1

        if total_checks == 0:
            return 0.5  # 没有已知关联规则，给中等分数

        return relation_score / total_checks

    def _find_related_keywords(self, header: str) -> List[str]:
        """查找与表头相关的关键字"""
        header_lower = header.lower()
        related = []

        for key, values in self.COLUMN_RELATIONS.items():
            if key.lower() in header_lower:
                related.extend(values)

        return related


# ==================== 多候选方案评分机制 ====================

@dataclass
class BoundaryCandidate:
    """边界候选方案"""
    header_start: int
    header_end: int
    data_start: int
    method: str  # 识别方法名称
    score: float = 0.0
    details: Dict[str, Any] = field(default_factory=dict)


class BoundaryCandidateEvaluator:
    """边界候选方案评估器 - 评估多个候选边界方案并选择最优"""

    def __init__(self, row_analyzer: 'EnhancedRowAnalyzer', relation_analyzer: ColumnRelationAnalyzer):
        self.logger = logging.getLogger(__name__)
        self.row_analyzer = row_analyzer
        self.relation_analyzer = relation_analyzer

    def evaluate_candidates(self, worksheet: Any, candidates: List[BoundaryCandidate],
                          max_col: int) -> Optional[BoundaryCandidate]:
        """评估多个候选方案，返回最优方案

        评分维度：
        1. 表头质量得分 (30%)
        2. 数据区域质量得分 (30%)
        3. 边界清晰度得分 (20%)
        4. 列关联性得分 (20%)
        """
        if not candidates:
            return None

        for candidate in candidates:
            # 1. 表头质量得分
            header_score = self._evaluate_header_quality(worksheet, candidate, max_col)

            # 2. 数据区域质量得分
            data_score = self._evaluate_data_quality(worksheet, candidate, max_col)

            # 3. 边界清晰度得分
            boundary_score = self._evaluate_boundary_clarity(worksheet, candidate, max_col)

            # 4. 列关联性得分
            relation_score = self._evaluate_column_relations(worksheet, candidate, max_col)

            # 综合得分
            candidate.score = (
                header_score * 0.3 +
                data_score * 0.3 +
                boundary_score * 0.2 +
                relation_score * 0.2
            )

            candidate.details = {
                'header_score': header_score,
                'data_score': data_score,
                'boundary_score': boundary_score,
                'relation_score': relation_score
            }

            self.logger.debug(f"候选方案[{candidate.method}] 表头:{candidate.header_start}-{candidate.header_end} "
                            f"得分:{candidate.score:.3f} (表头:{header_score:.2f} 数据:{data_score:.2f} "
                            f"边界:{boundary_score:.2f} 关联:{relation_score:.2f})")

        # 返回得分最高的候选方案
        best_candidate = max(candidates, key=lambda c: c.score)
        return best_candidate

    def _evaluate_header_quality(self, worksheet: Any, candidate: BoundaryCandidate,
                                max_col: int) -> float:
        """评估表头质量"""
        score = 0.0
        header_rows = range(candidate.header_start, candidate.header_end + 1)

        # 检查表头行的特征
        for row in header_rows:
            features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)

            # 文本比例高
            if features.text_ratio > 0.7:
                score += 0.3

            # 关键字密度高
            if features.keyword_ratio > 0.3:
                score += 0.3

            # 有特殊格式
            if features.has_special_formatting:
                score += 0.2

            # 有合并单元格
            if features.has_merged_cells:
                score += 0.2

        # 归一化
        return min(1.0, score / len(header_rows))

    def _evaluate_data_quality(self, worksheet: Any, candidate: BoundaryCandidate,
                              max_col: int) -> float:
        """评估数据区域质量"""
        # 检查数据起始行后的几行
        check_rows = min(5, worksheet.max_row - candidate.data_start + 1)
        if check_rows < 1:
            return 0.0

        rows_to_check = [candidate.data_start + i for i in range(check_rows)
                        if candidate.data_start + i <= worksheet.max_row]

        if not rows_to_check:
            return 0.0

        # 检查模式一致性
        consistency = self.row_analyzer.check_row_pattern_consistency(worksheet, rows_to_check, max_col)

        # 检查是否有序号列
        seq_col = self.row_analyzer.detect_sequence_column(worksheet, candidate.data_start, check_rows, max_col)
        has_sequence = seq_col > 0

        # 检查数值密度
        features_list = [self.row_analyzer.analyze_row_features(worksheet, r, max_col) for r in rows_to_check]
        avg_number_ratio = sum(f.number_ratio for f in features_list) / len(features_list)

        score = consistency * 0.5 + (0.3 if has_sequence else 0) + min(avg_number_ratio, 0.2)
        return min(1.0, score)

    def _evaluate_boundary_clarity(self, worksheet: Any, candidate: BoundaryCandidate,
                                  max_col: int) -> float:
        """评估边界清晰度"""
        if candidate.header_end >= candidate.data_start:
            return 0.0

        # 计算表头到数据的转换得分
        transition_score = self.row_analyzer.calculate_header_to_data_transition_score(
            worksheet, candidate.header_end, candidate.data_start, max_col
        )

        return transition_score

    def _evaluate_column_relations(self, worksheet: Any, candidate: BoundaryCandidate,
                                  max_col: int) -> float:
        """评估列关联性"""
        # 提取表头文本
        headers = []
        for row in range(candidate.header_start, candidate.header_end + 1):
            for col in range(1, max_col + 1):
                cell = worksheet.cell(row, col)
                value = self._get_cell_value(cell)
                if value and str(value).strip():
                    headers.append(str(value).strip())

        if not headers:
            return 0.0

        # 使用列关联分析器
        return self.relation_analyzer.analyze_header_relations(headers)

    @staticmethod
    def _get_cell_value(cell) -> Any:
        if cell.value is None:
            return None
        if isinstance(cell.value, datetime):
            return cell.value.strftime("%Y-%m-%d")
        elif isinstance(cell.value, (int, float)):
            return cell.value
        else:
            return str(cell.value)


# ==================== 增强型行分析器 ====================

class EnhancedRowAnalyzer:
    """增强型行分析器 - 用于更准确地识别表头边界"""

    HEADER_KEYWORDS = HeaderRuleEngine.HEADER_KEYWORDS

    # 序号列的常见表头名称
    SEQUENCE_HEADER_KEYWORDS = {"序号", "序", "no", "no.", "#", "编号", "行号", "sn", "id"}

    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._merged_cell_index = None  # 【性能优化】外部注入的合并单元格索引

    def analyze_row_features(self, worksheet: Any, row: int, max_col: int) -> RowFeatures:
        """分析行的详细特征"""
        features = RowFeatures(row_index=row)

        text_lengths = []
        format_parts = []

        for col in range(1, max_col + 1):
            cell = worksheet.cell(row, col)
            value = self._get_cell_value(cell)

            if value is not None and str(value).strip():
                features.non_empty_count += 1
                str_value = str(value).strip()

                # 类型统计
                if isinstance(value, str):
                    features.text_count += 1
                    text_lengths.append(len(str_value))

                    # 关键字检测
                    if self._contains_header_keyword(str_value):
                        features.keyword_count += 1
                elif isinstance(value, (int, float)):
                    features.number_count += 1

                    # 序号检测（第一列或第二列的小整数）
                    if col <= 2 and isinstance(value, (int, float)):
                        int_val = int(value) if value == int(value) else -1
                        if 1 <= int_val <= 10000:
                            features.has_sequence_number = True
                            features.sequence_value = int_val

                # 格式签名
                format_parts.append(self._get_cell_format_signature(cell))

            # 合并单元格检测
            if not features.has_merged_cells:
                features.has_merged_cells = self._is_merged_cell(worksheet, row, col)

            # 特殊格式检测
            if not features.has_special_formatting:
                features.has_special_formatting = self._has_special_format(cell)

        # 计算比例
        if features.non_empty_count > 0:
            features.number_ratio = features.number_count / features.non_empty_count
            features.text_ratio = features.text_count / features.non_empty_count
            features.keyword_ratio = features.keyword_count / features.non_empty_count
            features.avg_text_length = sum(text_lengths) / len(text_lengths) if text_lengths else 0

        # 格式签名
        features.format_signature = "|".join(format_parts[:5])  # 取前5列的格式

        return features

    def detect_sequence_column(self, worksheet: Any, start_row: int, check_rows: int, max_col: int) -> int:
        """检测序号列的位置

        返回序号列的列号，如果没有找到返回-1
        """
        # 检查前3列
        for col in range(1, min(4, max_col + 1)):
            values = []
            for row in range(start_row, min(start_row + check_rows, worksheet.max_row + 1)):
                cell = worksheet.cell(row, col)
                value = self._get_cell_value(cell)
                if value is not None:
                    try:
                        num_val = float(value)
                        if num_val == int(num_val):
                            values.append(int(num_val))
                        else:
                            break
                    except (ValueError, TypeError):
                        break
                else:
                    break

            # 检查是否是递增序列（1,2,3... 或 从某个数开始递增）
            if len(values) >= 3:
                is_sequence = True
                for i in range(1, len(values)):
                    if values[i] != values[i-1] + 1:
                        is_sequence = False
                        break
                if is_sequence and values[0] >= 1:
                    return col

        return -1

    def analyze_number_density_change(self, worksheet: Any, row1: int, row2: int, max_col: int) -> float:
        """分析两行之间的数值密度变化

        返回值 > 0 表示row2的数值密度更高（更像数据行）
        """
        features1 = self.analyze_row_features(worksheet, row1, max_col)
        features2 = self.analyze_row_features(worksheet, row2, max_col)

        return features2.number_ratio - features1.number_ratio

    def check_row_pattern_consistency(self, worksheet: Any, rows: list, max_col: int) -> float:
        """检查多行之间的模式一致性

        返回0-1之间的值，越高表示越一致（越像数据行）
        """
        if len(rows) < 2:
            return 0.0

        features_list = [self.analyze_row_features(worksheet, r, max_col) for r in rows]

        # 计算一致性得分
        consistency_score = 0.0
        comparisons = 0

        for i in range(len(features_list) - 1):
            f1, f2 = features_list[i], features_list[i + 1]

            # 非空单元格数量相似度
            if max(f1.non_empty_count, f2.non_empty_count) > 0:
                count_sim = min(f1.non_empty_count, f2.non_empty_count) / max(f1.non_empty_count, f2.non_empty_count)
            else:
                count_sim = 1.0

            # 数值比例相似度
            ratio_sim = 1.0 - abs(f1.number_ratio - f2.number_ratio)

            # 序号连续性
            seq_score = 0.0
            if f1.has_sequence_number and f2.has_sequence_number:
                if f2.sequence_value == f1.sequence_value + 1:
                    seq_score = 1.0

            # 格式一致性
            format_sim = 1.0 if f1.format_signature == f2.format_signature else 0.5

            row_consistency = (count_sim * 0.2 + ratio_sim * 0.3 + seq_score * 0.3 + format_sim * 0.2)
            consistency_score += row_consistency
            comparisons += 1

        return consistency_score / comparisons if comparisons > 0 else 0.0

    def look_ahead_analysis(self, worksheet: Any, start_row: int, look_ahead_count: int, max_col: int) -> dict:
        """向前看分析 - 分析接下来几行的特征

        返回分析结果字典
        """
        result = {
            'is_data_region': False,
            'confidence': 0.0,
            'sequence_detected': False,
            'pattern_consistency': 0.0,
            'number_density_avg': 0.0,
            'keyword_density_avg': 0.0
        }

        rows_to_check = []
        for row in range(start_row, min(start_row + look_ahead_count, worksheet.max_row + 1)):
            if not self._is_empty_row(worksheet, row, max_col):
                rows_to_check.append(row)

        if len(rows_to_check) < 2:
            return result

        # 序号列检测
        seq_col = self.detect_sequence_column(worksheet, start_row, look_ahead_count, max_col)
        result['sequence_detected'] = seq_col > 0

        # 模式一致性
        result['pattern_consistency'] = self.check_row_pattern_consistency(worksheet, rows_to_check, max_col)

        # 数值密度和关键字密度
        features_list = [self.analyze_row_features(worksheet, r, max_col) for r in rows_to_check]
        if features_list:
            result['number_density_avg'] = sum(f.number_ratio for f in features_list) / len(features_list)
            result['keyword_density_avg'] = sum(f.keyword_ratio for f in features_list) / len(features_list)

        # 综合判断是否是数据区域
        confidence = 0.0
        if result['sequence_detected']:
            confidence += 0.4
        if result['pattern_consistency'] > 0.6:
            confidence += 0.3
        if result['number_density_avg'] > 0.3:
            confidence += 0.2
        if result['keyword_density_avg'] < 0.2:
            confidence += 0.1

        result['confidence'] = confidence
        result['is_data_region'] = confidence > 0.5

        return result

    def calculate_header_to_data_transition_score(self, worksheet: Any,
                                                   header_row: int, data_row: int, max_col: int) -> float:
        """计算从表头行到数据行的转换得分

        得分越高表示越可能是表头到数据的边界
        """
        header_features = self.analyze_row_features(worksheet, header_row, max_col)
        data_features = self.analyze_row_features(worksheet, data_row, max_col)

        score = 0.0

        # 1. 数值密度增加
        if data_features.number_ratio > header_features.number_ratio + 0.2:
            score += 0.25

        # 2. 关键字密度下降
        if header_features.keyword_ratio > data_features.keyword_ratio + 0.2:
            score += 0.25

        # 3. 数据行有序号
        if data_features.has_sequence_number and data_features.sequence_value == 1:
            score += 0.3

        # 4. 表头有特殊格式，数据行没有
        if header_features.has_special_formatting and not data_features.has_special_formatting:
            score += 0.1

        # 5. 表头有合并单元格，数据行没有
        if header_features.has_merged_cells and not data_features.has_merged_cells:
            score += 0.1

        return score

    # 辅助方法
    @staticmethod
    def _get_cell_value(cell) -> any:
        if cell.value is None:
            return None
        if isinstance(cell.value, datetime):
            return cell.value.strftime("%Y-%m-%d")
        elif isinstance(cell.value, (int, float)):
            return cell.value
        else:
            return str(cell.value)

    def _contains_header_keyword(self, text: str) -> bool:
        if not text:
            return False
        text_lower = text.lower()
        return any(keyword.lower() in text_lower for keyword in self.HEADER_KEYWORDS)

    def _is_merged_cell(self, worksheet: Any, row: int, col: int) -> bool:
        """【性能优化】优先使用索引查找"""
        if self._merged_cell_index is not None:
            return (row, col) in self._merged_cell_index
        cell = worksheet.cell(row, col)
        for merged_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return True
        return False

    @staticmethod
    def _has_special_format(cell) -> bool:
        try:
            if (cell.font and cell.font.bold) or \
               (cell.fill and cell.fill.fgColor and str(cell.fill.fgColor.rgb) not in ('00000000', 'FFFFFFFF')) or \
               (cell.alignment and cell.alignment.horizontal == 'center'):
                return True
        except:
            pass
        return False

    @staticmethod
    def _get_cell_format_signature(cell) -> str:
        """获取单元格的格式签名"""
        parts = []
        try:
            if cell.font and cell.font.bold:
                parts.append("B")
            if cell.fill and cell.fill.fgColor:
                rgb = str(cell.fill.fgColor.rgb)[:6]
                if rgb not in ('000000', 'FFFFFF'):
                    parts.append(f"F{rgb}")
            if cell.alignment:
                if cell.alignment.horizontal:
                    parts.append(f"H{cell.alignment.horizontal[0]}")
        except:
            pass
        return "".join(parts) if parts else "N"

    def _is_empty_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        for col in range(1, max_col + 1):
            value = self._get_cell_value(worksheet.cell(row, col))
            if value is not None and str(value).strip():
                return False
        return True


# ==================== Excel智能解析器 ====================

class IntelligentExcelParser:
    """Excel智能解析器（优化版）"""
    
    # 关键字列表
    HEADER_KEYWORDS = HeaderRuleEngine.HEADER_KEYWORDS
    
    TITLE_KEYWORDS = {
        "公司名称", "公司", "单位", "企业", "报表", "明细表", "汇总表", "统计表",
        "报告", "清单", "名称：", "标题", "主题", "年度", "月度", "季度", "表名",
        "制表", "编制", "审核", "日期：", "时间：", "部门：", "备注：",
        "备忘录", "工资调整", "薪资", "调整项", "说明", "通知", "表格",
        "company", "corporation", "enterprise", "report", "summary", "statement",
        "title", "subject", "annual", "monthly", "quarterly", "department", "memo"
    }

    # 说明区域关键字（用于识别说明行的开头）
    INSTRUCTION_KEYWORDS = {
        "填写说明", "说明：", "注意事项", "项目注意", "备注说明", "使用说明",
        "操作说明", "重要提示", "温馨提示", "注：", "手工填写", "下拉选项",
        "instructions", "notes", "remarks", "tips"
    }

    # 表头必填标记（这些符号在表头中表示必填项，不应该被当作说明行）
    REQUIRED_FIELD_MARKERS = {"※", "*", "★"}
    
    SUMMARY_KEYWORDS = {
        "合计", "总计", "小计", "汇总", "共计", "累计", "总和", "总额",
        "total", "subtotal", "summary", "sum", "grand total"
    }

    def __init__(self):
        self.header_rule_engine = HeaderRuleEngine()
        self.row_analyzer = EnhancedRowAnalyzer()  # 添加增强型行分析器
        self.logger = logging.getLogger(__name__)

        # 从环境变量读取配置
        import os
        self.max_columns = int(os.environ.get('EXCEL_MAX_COLUMNS', '150'))
        self.logger.info(f"Excel解析器配置: 最大列数={self.max_columns}")

        # 添加新的优化组件
        self.consistency_validator = ColumnConsistencyValidator()
        self.post_validator = PostValidator()
        self.relation_analyzer = ColumnRelationAnalyzer()
        self.boundary_evaluator = BoundaryCandidateEvaluator(self.row_analyzer, self.relation_analyzer)

        # 【性能优化】合并单元格索引缓存（按worksheet缓存）
        self._merged_cell_index = {}  # {worksheet_id: {(row, col): merged_range}}
        self._merged_cell_rows = {}   # {worksheet_id: {row: [merged_range, ...]}}
        self._current_ws_id = None

        # 【性能优化】预编译汇总行英文关键字正则
        self._summary_en_patterns = {}
        for keyword in self.SUMMARY_KEYWORDS:
            keyword_lower = keyword.lower()
            if not any(ord(c) > 127 for c in keyword):
                self._summary_en_patterns[keyword_lower] = re.compile(
                    r'\b' + re.escape(keyword_lower) + r'\b'
                )

        # 添加自定义日期格式规则
        _date_pattern = re.compile(r'(年|月|日|日期|time|date)', re.IGNORECASE)
        def custom_date_format_rule(value: Any, context: RowContext) -> float:
            if isinstance(value, str):
                return 1.0 if _date_pattern.search(value) else 0.0
            return 0.0

        self.header_rule_engine.add_rule("CustomDateFormatRule", custom_date_format_rule, 0.6)

    # ==================== Aspose.Cells 加载层 ====================
    def _build_merged_cell_index(self, worksheet):
        """为 worksheet 构建合并单元格索引，避免每次查找都遍历所有合并区域"""
        ws_id = id(worksheet)
        if ws_id == self._current_ws_id:
            return  # 已经构建过了

        self._current_ws_id = ws_id
        cell_index = {}   # (row, col) -> merged_range
        row_index = {}    # row -> [merged_range, ...]

        # _AsposeWorksheet.merged_cells.ranges 已是 _AsposeMergedRange 列表（1-indexed）
        for merged_range in worksheet.merged_cells.ranges:
            for r in range(merged_range.min_row, merged_range.max_row + 1):
                if r not in row_index:
                    row_index[r] = []
                row_index[r].append(merged_range)
                for c in range(merged_range.min_col, merged_range.max_col + 1):
                    cell_index[(r, c)] = merged_range

        self._merged_cell_index[ws_id] = cell_index
        self._merged_cell_rows[ws_id] = row_index

        # 注入索引到子组件，避免它们各自遍历
        self.row_analyzer._merged_cell_index = cell_index
        self.header_rule_engine._merged_cell_index = cell_index

    def _extract_file_manual_headers(self, file_path: str, manual_headers: Dict[str, Any]) -> Dict[str, Any]:
        """
        从manual_headers中提取当前文件的配置

        Args:
            file_path: 当前文件路径
            manual_headers: 完整的手动表头配置

        Returns:
            当前文件的manual_headers配置
        """
        if not manual_headers:
            return {}

        file_name = Path(file_path).name
        base_name = Path(file_path).stem  # 不带扩展名的文件名

        # 尝试多种key匹配
        possible_keys = [file_name, base_name]

        for key in possible_keys:
            if key in manual_headers:
                value = manual_headers[key]
                if isinstance(value, dict):
                    # 新格式: {"文件名.xlsx": {"Sheet1": [1, 1]}}
                    return value
                elif isinstance(value, list) and len(value) == 2:
                    # 旧格式: {"Sheet名称": [1, 1]}
                    # 转换为sheet配置
                    return {key: value}
                else:
                    self.logger.warning(f"manual_headers中{key}的格式无效: {value}")

        # 如果没有找到文件级别的配置，检查是否有sheet级别的配置（旧格式）
        # 旧格式的key就是sheet名
        return manual_headers
    
    def parse_excel_file(self, file_path: str, max_data_rows: int = None, skip_rows: int = 0,
                        manual_headers: Dict[str, Any] = None, headers_only: bool = False,
                        active_sheet_only: bool = False) -> List[SheetData]:
        """读取并解析Excel文件

        Args:
            file_path: Excel文件路径
            max_data_rows: 每个区域最多读取的数据行数，None表示读取全部
            skip_rows: 从文件开头跳过的行数（用于跳过非数据行）
            manual_headers: 手动指定的表头范围，支持多种格式：
                          1. 旧格式: {"Sheet名称": [start_row, end_row]}
                          2. 新格式: {"文件名.xlsx": {"Sheet1": [1, 1], "Sheet2": [3, 3]}}
            headers_only: 是否只读取表头，不读取数据行（用于快速匹配）
            active_sheet_only: 是否只加载当前激活的Sheet，默认False（加载所有Sheet）

        Returns:
            解析后的Sheet数据列表
        """
        result = []

        try:
            # 使用 Aspose.Cells for .NET 加载
            aspose_wb = _AsposeWorkbook(str(file_path))

            # 提取当前文件的 manual_headers 配置
            file_manual_headers = self._extract_file_manual_headers(file_path, manual_headers)

            # 确定要处理的 sheet 列表
            if active_sheet_only:
                active_idx = aspose_wb.Worksheets.ActiveSheetIndex
                sheets_to_parse = [aspose_wb.Worksheets[active_idx]]
            else:
                sheets_to_parse = [aspose_wb.Worksheets[i] for i in range(aspose_wb.Worksheets.Count)]

            for aspose_ws in sheets_to_parse:
                # 包装为兼容适配器
                ws = _AsposeWorksheet(aspose_ws)

                # 检查是否有手动指定的表头范围
                manual_header_range = None
                if file_manual_headers:
                    sheet_ranges = file_manual_headers.get(ws.title)
                    if isinstance(sheet_ranges, list) and len(sheet_ranges) == 2:
                        manual_header_range = sheet_ranges
                    elif isinstance(sheet_ranges, dict):
                        for key, value in sheet_ranges.items():
                            if isinstance(value, list) and len(value) == 2:
                                manual_header_range = value
                                break

                sheet_data = self._parse_sheet(ws, max_data_rows, skip_rows, manual_header_range, headers_only)
                if sheet_data and sheet_data.regions:
                    result.append(sheet_data)

        except Exception as e:
            print(f"解析Excel文件时出错: {e}")
            import traceback
            traceback.print_exc()

        return result
    
    def _parse_sheet(self, worksheet: Any, max_data_rows: int = None, skip_rows: int = 0,
                    manual_header_range: List[int] = None, headers_only: bool = False) -> Optional[SheetData]:
        """解析单个Sheet

        Args:
            worksheet: 工作表对象
            max_data_rows: 每个区域最多读取的数据行数，None表示读取全部
            skip_rows: 从文件开头跳过的行数
            manual_header_range: 手动指定的表头范围 [start_row, end_row]
            headers_only: 是否只读取表头，不读取数据行

        Returns:
            解析后的Sheet数据
        """
        sheet_data = SheetData(sheet_name=worksheet.title, regions=[])

        # 【性能优化】预建合并单元格索引
        self._build_merged_cell_index(worksheet)

        max_row = worksheet.max_row
        max_col = worksheet.max_column

        # 如果只读取表头，限制扫描行数
        if headers_only:
            max_data_rows = 0
            # 表头通常在前30行内，限制扫描范围以加速
            if max_row and max_row > 50:
                max_row = 50

        # 修正 max_col：openpyxl 可能因空格式单元格导致 max_column 虚高
        # 通过扫描前几行实际数据来确定真实的最大列数
        if max_col and max_col > self.max_columns:
            real_max_col = 1
            sample_rows = min(max_row or 0, 20)
            for r in range(1, sample_rows + 1):
                for c in range(max_col, 0, -1):
                    cell_val = worksheet.cell(row=r, column=c).value
                    if cell_val is not None:
                        real_max_col = max(real_max_col, c)
                        break
            max_col = min(real_max_col + 5, max_col)  # 留5列余量

        if max_row == 0 or max_col == 0:
            return sheet_data
        
        # 如果手动指定了表头范围，直接使用
        if manual_header_range:
            try:
                # 确保manual_header_range是有效的范围
                if isinstance(manual_header_range, (list, tuple)) and len(manual_header_range) == 2:
                    start_row, end_row = manual_header_range
                    region = self._parse_region_with_manual_header(worksheet, start_row, end_row, max_row, max_col, max_data_rows)
                    if region:
                        sheet_data.regions.append(region)
                    return sheet_data
                else:
                    self.logger.warning(f"无效的手动表头范围格式: {manual_header_range}")
            except Exception as e:
                self.logger.warning(f"解析手动表头范围失败: {e}, 范围: {manual_header_range}")
        
        # 否则使用自动解析逻辑
        # 从跳过指定行数后开始解析
        current_row = skip_rows + 1
        
        while current_row <= max_row:
            # 跳过空行和标题行
            if self._is_empty_row(worksheet, current_row, max_col) or \
               self._is_title_row(worksheet, current_row, max_col):
                current_row += 1
                continue
            
            header_info = self._analyze_header_range(worksheet, current_row, min(current_row + 20, max_row), max_col)
            
            if header_info:
                region = self._parse_region(worksheet, header_info.start_row, max_row, max_col, max_data_rows)
                
                if region:
                    sheet_data.regions.append(region)
                    
                    if region.data_row_end >= region.data_row_start:
                        current_row = region.data_row_end + 1
                    else:
                        current_row = region.head_row_end + 1
                    
                    # 跳过汇总行和空行
                    while current_row <= max_row and \
                          (self._is_empty_row(worksheet, current_row, max_col) or \
                           self._is_summary_row(worksheet, current_row, max_col)):
                        current_row += 1
                else:
                    current_row = header_info.end_row + 1
            else:
                current_row += 1
        
        return sheet_data
    
    def _analyze_header_range(self, worksheet: Any, start_row: int, max_row: int, max_col: int) -> Optional[HeaderInfo]:
        """分析表头范围（增强版 - 多候选方案评分）

        策略：生成多个候选边界方案，评分后选择最优
        1. 【优先】如果第1行有加粗格式，直接使用单行表头
        2. 方法1：反向查找（从数据行向上找表头）
        3. 方法2：正向查找（传统方法）
        4. 方法3：滑动窗口（检测特征突变点）
        5. 评分选择最优方案
        """
        # 【关键优化】优先检查第1行是否有加粗格式
        if start_row == 1 and self._check_row_has_bold(worksheet, 1, max_col):
            # 检查第1行是否是文本为主
            features = self.row_analyzer.analyze_row_features(worksheet, 1, max_col)
            if features.text_ratio > 0.7:
                # 第1行有加粗且是文本为主，检查后续行是否也是表头（多行表头）
                header_end = 1
                for next_row in range(2, min(start_row + 6, max_row + 1)):
                    if self._is_empty_row(worksheet, next_row, max_col):
                        break
                    if self._is_summary_row(worksheet, next_row, max_col):
                        break
                    next_features = self.row_analyzer.analyze_row_features(worksheet, next_row, max_col)
                    next_has_bold = self._check_row_has_bold(worksheet, next_row, max_col)
                    # 如果下一行也是加粗的文本行（关键字多或有合并单元格），扩展表头
                    if next_has_bold and next_features.text_ratio > 0.7:
                        header_end = next_row
                    else:
                        break
                self.logger.debug(f"第1行有加粗格式，表头范围: 1-{header_end}")
                return HeaderInfo(start_row=1, end_row=header_end)

        candidates = []

        # 方法1：反向查找策略
        candidate1 = self._find_boundary_by_reverse_search(worksheet, start_row, max_row, max_col)
        if candidate1:
            candidates.append(candidate1)

        # 方法2：正向查找策略（传统方法）
        candidate2 = self._find_boundary_by_forward_search(worksheet, start_row, max_row, max_col)
        if candidate2:
            candidates.append(candidate2)

        # 方法3：滑动窗口策略
        candidate3 = self._find_boundary_by_sliding_window(worksheet, start_row, max_row, max_col)
        if candidate3:
            candidates.append(candidate3)

        # 如果没有候选方案，返回None
        if not candidates:
            return None

        # 使用评估器选择最优方案
        best_candidate = self.boundary_evaluator.evaluate_candidates(worksheet, candidates, max_col)

        if best_candidate:
            self.logger.debug(f"选择最优方案: {best_candidate.method}, "
                            f"表头:{best_candidate.header_start}-{best_candidate.header_end}, "
                            f"得分:{best_candidate.score:.3f}")
            return HeaderInfo(start_row=best_candidate.header_start, end_row=best_candidate.header_end)

        return None

    def _find_boundary_by_reverse_search(self, worksheet: Any, start_row: int,
                                        max_row: int, max_col: int) -> Optional[BoundaryCandidate]:
        """方法1：反向查找策略"""
        # 步骤1：找到第一个明确的数据行
        first_data_row = self._find_first_data_row(worksheet, start_row, min(start_row + 20, max_row), max_col)

        if first_data_row is None:
            return None

        # 步骤2：从数据行向上查找表头
        header_end_row = self._find_header_by_looking_up(worksheet, first_data_row, start_row, max_col)

        if header_end_row is None:
            return None

        # 步骤3：确定表头起始行
        # 始终调用_find_header_start，允许多行表头检测（即使有加粗格式）
        header_start_row = self._find_header_start(worksheet, header_end_row, start_row, max_col)

        return BoundaryCandidate(
            header_start=header_start_row,
            header_end=header_end_row,
            data_start=first_data_row,
            method="反向查找"
        )

    def _find_boundary_by_forward_search(self, worksheet: Any, start_row: int,
                                        max_row: int, max_col: int) -> Optional[BoundaryCandidate]:
        """方法2：正向查找策略（传统方法）"""
        header_info = self._analyze_header_range_forward(worksheet, start_row, max_row, max_col)

        if header_info is None:
            return None

        return BoundaryCandidate(
            header_start=header_info.start_row,
            header_end=header_info.end_row,
            data_start=header_info.end_row + 1,
            method="正向查找"
        )

    def _find_boundary_by_sliding_window(self, worksheet: Any, start_row: int,
                                        max_row: int, max_col: int) -> Optional[BoundaryCandidate]:
        """方法3：滑动窗口策略 - 检测特征突变点"""
        window_size = 3
        best_transition_row = None
        best_transition_score = 0

        # 滑动窗口扫描，寻找最明显的表头到数据转换点
        for row in range(start_row, min(start_row + 15, max_row)):
            if self._is_empty_row(worksheet, row, max_col):
                continue

            # 检查当前行和后续行的特征变化
            if row + window_size <= max_row:
                # 分析窗口内的行特征
                window_rows = [row + i for i in range(window_size)
                              if not self._is_empty_row(worksheet, row + i, max_col)]

                if len(window_rows) < 2:
                    continue

                # 检查是否有明显的类型转换
                first_row_features = self.row_analyzer.analyze_row_features(worksheet, window_rows[0], max_col)
                rest_rows_features = [self.row_analyzer.analyze_row_features(worksheet, r, max_col)
                                     for r in window_rows[1:]]

                # 计算转换得分
                if first_row_features.text_ratio > 0.6:  # 第一行是文本为主
                    avg_number_ratio = sum(f.number_ratio for f in rest_rows_features) / len(rest_rows_features)
                    if avg_number_ratio > 0.3:  # 后续行数值为主
                        transition_score = (first_row_features.text_ratio + avg_number_ratio) / 2

                        # 检查后续行的一致性
                        consistency = self.row_analyzer.check_row_pattern_consistency(
                            worksheet, window_rows[1:], max_col
                        )
                        transition_score = transition_score * 0.6 + consistency * 0.4

                        if transition_score > best_transition_score:
                            best_transition_score = transition_score
                            best_transition_row = row

        if best_transition_row is None or best_transition_score < 0.4:
            return None

        # 确定表头起始行
        header_start = self._find_header_start(worksheet, best_transition_row, start_row, max_col)

        return BoundaryCandidate(
            header_start=header_start,
            header_end=best_transition_row,
            data_start=best_transition_row + 1,
            method="滑动窗口"
        )

    def _find_first_data_row(self, worksheet: Any, start_row: int, max_row: int, max_col: int) -> Optional[int]:
        """找到第一个明确的数据行（基于结构特征）

        策略：不依赖关键字，而是基于结构特征识别数据行
        1. 数据行通常包含数值（如员工编号、金额等）
        2. 数据行之间的结构相似（一致性高）
        3. 表头和数据行之间类型分布有明显差异
        """
        # 收集所有非空行的特征
        row_features_list = []
        for row in range(start_row, min(max_row + 1, start_row + 30)):
            if self._is_empty_row(worksheet, row, max_col):
                continue
            if self._is_title_row(worksheet, row, max_col):
                continue
            if self._is_summary_row(worksheet, row, max_col):
                continue
            features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)
            row_features_list.append((row, features))

        if len(row_features_list) < 2:
            return None

        # 策略1：找到数值比例突然增加的行（表头->数据的转换点）
        for i in range(1, len(row_features_list)):
            prev_row, prev_features = row_features_list[i - 1]
            curr_row, curr_features = row_features_list[i]

            # 检查是否是表头到数据的转换
            # 条件：前一行文本为主，当前行有较多数值
            if prev_features.text_ratio > 0.7 and curr_features.number_ratio > 0.2:
                # 额外验证：检查后续几行是否结构一致
                if i + 1 < len(row_features_list):
                    next_row, next_features = row_features_list[i + 1]
                    # 如果当前行和下一行结构相似，确认是数据行
                    if abs(curr_features.number_ratio - next_features.number_ratio) < 0.3:
                        return curr_row
                else:
                    return curr_row

        # 策略2：找到第一列是数字（员工编号）且行结构与后续行一致的行
        for i, (row, features) in enumerate(row_features_list):
            first_val = self._get_cell_value(worksheet.cell(row, 1))
            if first_val is not None:
                try:
                    num_val = float(first_val)
                    # 第一列是大数字（如员工编号26123013）
                    if num_val > 1000:
                        # 验证后续行也有类似结构
                        if i + 1 < len(row_features_list):
                            next_row, next_features = row_features_list[i + 1]
                            next_first_val = self._get_cell_value(worksheet.cell(next_row, 1))
                            if next_first_val is not None:
                                try:
                                    next_num = float(next_first_val)
                                    if next_num > 1000:
                                        return row
                                except (ValueError, TypeError):
                                    pass
                        else:
                            return row
                except (ValueError, TypeError):
                    pass

        # 策略3：找到结构高度一致的连续行（数据区域）
        for i in range(len(row_features_list) - 2):
            rows_to_check = [row_features_list[j][0] for j in range(i, min(i + 3, len(row_features_list)))]
            consistency = self.row_analyzer.check_row_pattern_consistency(worksheet, rows_to_check, max_col)
            if consistency > 0.7:
                return row_features_list[i][0]

        return None

    def _find_header_by_looking_up(self, worksheet: Any, data_row: int, min_row: int, max_col: int) -> Optional[int]:
        """从数据行向上查找表头行（基于结构特征）

        策略：不依赖关键字，基于以下结构特征识别表头：
        1. 表头行非空单元格多（完整性高）
        2. 表头行全是文本，数据行有数字
        3. 表头行的类型分布与数据行有明显差异
        4. 表头行有特殊格式（加粗、背景色等）- 权重提高
        5. 优先返回最近的高分行（不再无条件优先加粗行）
        """
        # 获取数据行的特征作为参照
        data_features = self.row_analyzer.analyze_row_features(worksheet, data_row, max_col)

        best_header_row = None
        best_score = 0
        first_bold_row = None

        # 从数据行的上一行开始向上查找
        for row in range(data_row - 1, max(min_row - 1, 0), -1):
            if self._is_empty_row(worksheet, row, max_col):
                continue

            features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)

            # 如果是标题行或说明行，跳过继续向上
            if self._is_title_row(worksheet, row, max_col) or self._is_instruction_row(worksheet, row, max_col):
                continue

            has_bold = self._check_row_has_bold(worksheet, row, max_col)
            if has_bold and first_bold_row is None:
                first_bold_row = row

            # 计算该行作为表头的得分
            header_score = 0

            # 特征1：完整性 - 非空单元格应该较多
            if features.non_empty_count >= data_features.non_empty_count * 0.8:
                header_score += 0.2

            # 特征2：文本主导 - 表头应该主要是文本
            if features.text_ratio > 0.7:
                header_score += 0.2

            # 特征3：类型差异 - 与数据行的类型分布应该不同
            type_diff = features.text_ratio - data_features.text_ratio
            if type_diff > 0.2:
                header_score += 0.15

            # 特征4：数值比例低 - 表头几乎没有数值
            if features.number_ratio < 0.2:
                header_score += 0.1

            # 特征5：格式特殊（加粗、居中等）
            if features.has_special_formatting:
                header_score += 0.25
                if not data_features.has_special_formatting:
                    header_score += 0.15

            # 特征6：没有序号
            if not features.has_sequence_number:
                header_score += 0.05

            # 特征7：关键字密度
            if features.keyword_ratio > 0.3:
                header_score += 0.1

            # 如果有加粗，提高得分
            if has_bold:
                header_score += 0.3

            # 如果这行的得分高，可能是表头
            if header_score > best_score and header_score >= 0.5:
                best_score = header_score
                best_header_row = row

            # 如果遇到数值比例很高的行，说明已经进入了另一个数据区域
            if features.number_ratio > 0.4:
                break

        # 【优化】优先返回最近数据行的高分行（best_header_row），
        # 仅当 best_header_row 为空且有加粗行时回退到 first_bold_row
        if best_header_row is not None:
            # 如果 first_bold_row 和 best_header_row 相邻（差 <=1 行），优先返回加粗行
            if first_bold_row is not None and abs(first_bold_row - best_header_row) <= 1:
                return first_bold_row
            return best_header_row

        return first_bold_row

    def _check_row_has_bold(self, worksheet: Any, row: int, max_col: int) -> bool:
        """检查行中是否有加粗的单元格

        Args:
            row: 行号
            max_col: 最大列数

        Returns:
            True 表示有加粗单元格
        """
        bold_count = 0
        total_count = 0

        for col in range(1, min(max_col + 1, 20)):  # 检查前20列
            cell = worksheet.cell(row, col)
            if cell.value is not None and str(cell.value).strip():
                total_count += 1
                try:
                    if cell.font and cell.font.bold:
                        bold_count += 1
                except:
                    pass

        # 如果超过50%的单元格加粗，认为这行有加粗格式
        if total_count > 0 and bold_count / total_count > 0.5:
            return True

        return False

    def _find_header_start(self, worksheet: Any, header_end_row: int, min_row: int, max_col: int) -> int:
        """从表头结束行向上查找表头起始行（合并格感知版）

        策略：
        1. 优先检查合并格覆盖率，高覆盖率的行直接判定为表头的一部分
        2. 检查垂直合并穿透（上方行的合并格延伸到当前表头区域）
        3. 对于无合并格的行，使用原有的文本比例/关键字检测
        4. 遇到空行、标题行、格式突变时停止
        """
        header_start = header_end_row

        # 获取表头结束行的特征作为参照
        end_row_features = self.row_analyzer.analyze_row_features(worksheet, header_end_row, max_col)

        prev_row = header_end_row

        for row in range(header_end_row - 1, max(min_row - 1, 0), -1):
            if self._is_empty_row(worksheet, row, max_col):
                break

            # 如果是标题行或说明行，表头从下一行开始
            if self._is_title_row(worksheet, row, max_col) or self._is_instruction_row(worksheet, row, max_col):
                break

            features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)

            # 检查相邻行之间的格式突变
            # 【优化】如果当前行有高合并格覆盖率，跳过格式突变检查
            # 多行合并表头中，上层行和底层行的格式（加粗/背景色）经常不同，这不应该被视为断裂
            merge_coverage = self._calc_merge_coverage(worksheet, row, max_col)
            if merge_coverage < 0.5:
                if self._has_format_break(worksheet, row, prev_row, max_col):
                    break

            # ==================== 合并格感知判断 ====================

            # 计算合并格覆盖率（水平合并 + 垂直穿透 + 普通非空格）
            merge_coverage = self._calc_merge_coverage(worksheet, row, max_col)
            has_h_merge = self._has_horizontal_merge(worksheet, row, max_col)
            v_merge_count = self._has_vertical_merge_from_above(worksheet, row, max_col)

            should_continue = False

            # 【核心优化1】高合并格覆盖率 → 直接判定为多行表头的上层行
            # 例如: Row 1 的 "养老保险[7-12]" + "医疗保险[13-20]" + 垂直合并的"序号"等 → 覆盖率 > 0.7
            if merge_coverage >= 0.65 and has_h_merge:
                should_continue = True

            # 【核心优化2】有垂直合并穿透 + 有水平合并 → 明确的多行合并表头
            # 例如: Row 1-3 的"序号"列是垂直合并的，同时 Row 1 有水平合并的大类别
            elif v_merge_count >= 2 and has_h_merge:
                should_continue = True

            # 【核心优化3】高垂直合并穿透 + 高覆盖率（无需水平合并）
            # 例如: Row 3 有19个垂直穿透列 + 覆盖率0.83，属于多行表头的中间层
            elif v_merge_count >= 3 and merge_coverage >= 0.6:
                should_continue = True

            # 【核心优化4】有水平合并 + 合并格覆盖率不太低 → 可能是中间层表头
            # 例如: Row 2 的 "公司[8-9]", "个人[10-11]" 等子类别
            elif has_h_merge and merge_coverage >= 0.4:
                # 额外检查：非空文本应该是表头关键字
                if features.non_empty_count > 0 and features.text_ratio > 0.5:
                    should_continue = True

            # ==================== 原有逻辑（无合并格的行） ====================

            if not should_continue:
                # 条件：主要是文本（放宽：考虑合并格覆盖时已经处理过高覆盖情况）
                if features.text_ratio < 0.6:
                    break

                # 条件：完整性相近
                completeness_ratio = min(features.non_empty_count, end_row_features.non_empty_count) / \
                                     max(features.non_empty_count, end_row_features.non_empty_count, 1)
                if completeness_ratio < 0.3:
                    break

                # 严格条件：无合并格时需要更强的文本/关键字/加粗证据
                if features.text_ratio > 0.9 and features.keyword_ratio > 0.5:
                    should_continue = True
                elif (features.text_ratio > 0.8 and
                      self._check_row_has_bold(worksheet, row, max_col) and
                      self._check_row_has_bold(worksheet, header_end_row, max_col)):
                    should_continue = True

            if should_continue:
                header_start = row
                prev_row = row
            else:
                break

        return header_start

    def _has_format_break(self, worksheet: Any, row1: int, row2: int, max_col: int) -> bool:
        """检测两行之间是否有明显的格式突变

        Args:
            row1: 第一行
            row2: 第二行（参考行）
            max_col: 最大列数

        Returns:
            True 表示有明显格式突变
        """
        # 统计两行的格式特征
        row1_bold_count = 0
        row1_bg_count = 0
        row1_total = 0

        row2_bold_count = 0
        row2_bg_count = 0
        row2_total = 0

        for col in range(1, min(max_col + 1, 20)):  # 检查前20列
            # 检查第一行
            cell1 = worksheet.cell(row1, col)
            if cell1.value is not None and str(cell1.value).strip():
                row1_total += 1
                try:
                    if cell1.font and cell1.font.bold:
                        row1_bold_count += 1
                    if cell1.fill and cell1.fill.fgColor and str(cell1.fill.fgColor.rgb) not in ('00000000', 'FFFFFFFF'):
                        row1_bg_count += 1
                except:
                    pass

            # 检查第二行
            cell2 = worksheet.cell(row2, col)
            if cell2.value is not None and str(cell2.value).strip():
                row2_total += 1
                try:
                    if cell2.font and cell2.font.bold:
                        row2_bold_count += 1
                    if cell2.fill and cell2.fill.fgColor and str(cell2.fill.fgColor.rgb) not in ('00000000', 'FFFFFFFF'):
                        row2_bg_count += 1
                except:
                    pass

        # 如果两行都没有有效单元格，不算格式突变
        if row1_total == 0 or row2_total == 0:
            return False

        # 计算格式比例
        row1_bold_ratio = row1_bold_count / row1_total if row1_total > 0 else 0
        row2_bold_ratio = row2_bold_count / row2_total if row2_total > 0 else 0

        row1_bg_ratio = row1_bg_count / row1_total if row1_total > 0 else 0
        row2_bg_ratio = row2_bg_count / row2_total if row2_total > 0 else 0

        # 检测格式突变：
        # 1. 加粗比例差异大于50%（如一行全加粗，另一行全不加粗）
        bold_diff = abs(row1_bold_ratio - row2_bold_ratio)
        if bold_diff > 0.5:
            return True

        # 2. 背景色比例差异大于50%
        bg_diff = abs(row1_bg_ratio - row2_bg_ratio)
        if bg_diff > 0.5:
            return True

        return False

    def _calc_merge_coverage(self, worksheet: Any, row: int, max_col: int) -> float:
        """计算某行的合并格覆盖率（包含水平合并覆盖 + 垂直合并穿透 + 普通非空格）

        对于多行合并表头的上层行：
        - 水平合并格（如"养老保险"跨6列）覆盖的列全部计入
        - 垂直合并格（如"序号"跨3行）穿透当前行的列也计入
        - 当前行有值的普通列也计入

        Returns:
            覆盖率 0.0~1.0
        """
        if max_col <= 0:
            return 0.0

        covered_cols = set()
        ws_id = id(worksheet)
        row_ranges = self._merged_cell_rows.get(ws_id, {}).get(row)

        if row_ranges:
            for mr in row_ranges:
                for c in range(mr.min_col, mr.max_col + 1):
                    if c <= max_col:
                        covered_cols.add(c)

        # 同时检查从上方穿透到当前行的垂直合并格（起始行 < 当前行）
        # 通过索引查找：如果 (row, col) 在索引中但起始行 < row，说明是穿透
        cell_index = self._merged_cell_index.get(ws_id, {})
        for col in range(1, max_col + 1):
            if col in covered_cols:
                continue
            mr = cell_index.get((row, col))
            if mr and mr.min_row < row:
                # 垂直合并穿透：上方行的合并格延伸到本行
                covered_cols.add(col)
            elif self._get_cell_value(worksheet.cell(row, col)) is not None:
                # 普通非空格
                covered_cols.add(col)

        return len(covered_cols) / max_col

    def _has_vertical_merge_from_above(self, worksheet: Any, row: int, max_col: int) -> int:
        """统计有多少列是从上方垂直合并穿透到当前行的

        Returns:
            穿透列数
        """
        ws_id = id(worksheet)
        cell_index = self._merged_cell_index.get(ws_id, {})
        count = 0
        for col in range(1, max_col + 1):
            mr = cell_index.get((row, col))
            if mr and mr.min_row < row and mr.max_row >= row:
                count += 1
        return count

    def _has_horizontal_merge(self, worksheet: Any, row: int, max_col: int) -> bool:
        """检查行中是否有水平合并单元格（跨多列）（使用索引优化）"""
        ws_id = id(worksheet)
        row_ranges = self._merged_cell_rows.get(ws_id, {}).get(row)
        if row_ranges is not None:
            for merged_range in row_ranges:
                if merged_range.max_col > merged_range.min_col:
                    return True
            return False
        # 回退
        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row:
                if merged_range.max_col > merged_range.min_col:
                    return True
        return False

    def _analyze_header_range_forward(self, worksheet: Any, start_row: int, max_row: int, max_col: int) -> Optional[HeaderInfo]:
        """原来的正向查找逻辑（作为回退方案）"""
        actual_start_row = start_row

        # 跳过空行、标题行和说明行
        while actual_start_row <= max_row:
            if self._is_empty_row(worksheet, actual_start_row, max_col):
                actual_start_row += 1
                continue

            if self._is_title_row(worksheet, actual_start_row, max_col):
                actual_start_row += 1
                continue

            if self._is_instruction_row(worksheet, actual_start_row, max_col):
                actual_start_row += 1
                continue

            # 检查当前行是否像真正的表头（包含多个表头关键字）
            row_features = self.row_analyzer.analyze_row_features(worksheet, actual_start_row, max_col)

            # 如果关键字比例很低且文本很多，可能还是说明区域
            if row_features.keyword_ratio < 0.2 and row_features.text_ratio > 0.8:
                # 再检查是否大部分是长文本（说明文字通常较长）
                if row_features.avg_text_length > 20:
                    actual_start_row += 1
                    continue

            break

        if actual_start_row > max_row:
            return None

        end_row = self._find_header_end(worksheet, actual_start_row, max_row, max_col)

        if end_row < actual_start_row:
            end_row = actual_start_row

        return HeaderInfo(start_row=actual_start_row, end_row=end_row)
    
    def _find_header_end(self, worksheet: Any, start_row: int, max_row: int, max_col: int) -> int:
        """查找表头结束位置（增强版）

        使用多种策略综合判断：
        1. 序号列检测 - 如果下一行开始出现序号1,2,3...则表头结束
        2. 多行向前看 - 分析接下来几行的模式一致性
        3. 数值密度突变 - 从文本为主突变到数字为主
        4. 格式一致性 - 数据行格式通常一致
        5. 传统规则引擎评分
        """
        current_header_end = start_row
        MAX_HEADER_ROWS = 8  # 增加表头最大行数限制

        for row in range(start_row, min(start_row + MAX_HEADER_ROWS + 1, max_row + 1)):
            if self._is_empty_row(worksheet, row, max_col):
                continue

            if self._is_title_row(worksheet, row, max_col):
                continue

            # === 策略1: 序号列检测 ===
            # 检查从当前行开始是否有序号列（1,2,3...）
            if row > start_row:
                seq_col = self.row_analyzer.detect_sequence_column(worksheet, row, 5, max_col)
                if seq_col > 0:
                    # 检查第一个值是否是1或接近1的小数字
                    first_val = worksheet.cell(row, seq_col).value
                    try:
                        if isinstance(first_val, (int, float)) and 1 <= float(first_val) <= 3:
                            # 确认这是序号列的开始
                            self.logger.debug(f"Row {row}: 检测到序号列从{first_val}开始，表头结束于{current_header_end}")
                            return current_header_end
                    except:
                        pass

            # === 策略2: 多行向前看分析 ===
            if row > start_row:
                look_ahead = self.row_analyzer.look_ahead_analysis(worksheet, row, 5, max_col)
                if look_ahead['is_data_region'] and look_ahead['confidence'] > 0.6:
                    self.logger.debug(f"Row {row}: 向前看分析确认数据区域，置信度{look_ahead['confidence']:.2f}")
                    return current_header_end

            # === 策略3: 表头到数据的转换评分 ===
            if row > start_row:
                transition_score = self.row_analyzer.calculate_header_to_data_transition_score(
                    worksheet, current_header_end, row, max_col
                )
                if transition_score > 0.5:
                    self.logger.debug(f"Row {row}: 检测到表头到数据的转换，得分{transition_score:.2f}")
                    return current_header_end

            # === 策略4: 数据类型转换检测（原有逻辑增强）===
            if row > start_row and self._has_significant_data_type_transition(worksheet, current_header_end, row, max_col):
                self.logger.debug(f"Row {row}: 检测到显著的数据类型转换")
                return current_header_end

            # === 策略5: 行特征分析 ===
            row_features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)

            # 如果当前行的数值比例很高且关键字比例很低，很可能是数据行
            if row > start_row:
                if row_features.number_ratio > 0.4 and row_features.keyword_ratio < 0.2:
                    self.logger.debug(f"Row {row}: 数值比例高({row_features.number_ratio:.2f})，关键字少，判定为数据行")
                    return current_header_end

            # === 策略6: 传统规则引擎评分 ===
            row_type = self._analyze_row_type(worksheet, row, max_col, start_row, current_header_end)

            if row_type == RowType.HEADER:
                current_header_end = row
            elif row_type in (RowType.DATA, RowType.SUMMARY):
                return current_header_end

        return current_header_end
    
    def _analyze_row_type(self, worksheet: Any, row: int, max_col: int,
                         header_start_row: int, current_header_end: int) -> RowType:
        """分析行的类型（增强版）"""
        # 如果是第一行，更倾向于判断为表头
        is_first_row = (row == header_start_row)

        # 使用增强型分析器获取行特征
        features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)

        # 如果检测到序号且值为1-3，很可能是数据行的开始
        if features.has_sequence_number and 1 <= features.sequence_value <= 3 and not is_first_row:
            return RowType.DATA

        # 规则引擎评分
        rule_engine_score = self.header_rule_engine.calculate_header_score(worksheet, row, max_col)

        # 提高表头判断阈值，避免将数据行误判为表头
        header_threshold = 0.60 if is_first_row else 0.70

        # 综合判断
        if rule_engine_score > header_threshold:
            # 额外验证：即使规则引擎分数高，如果数值比例很高也可能是数据行
            if features.number_ratio > 0.5 and features.keyword_ratio < 0.3:
                return RowType.DATA
            return RowType.HEADER

        if self._is_summary_row(worksheet, row, max_col):
            return RowType.SUMMARY

        # 基于特征的判断
        if features.non_empty_count == 0:
            return RowType.UNKNOWN

        # 更严格的表头判断条件
        if is_first_row:
            if features.keyword_ratio > 0.5 or (features.has_merged_cells and features.text_ratio > 0.7 and features.keyword_ratio > 0.3):
                return RowType.HEADER
        else:
            if features.keyword_ratio > 0.6:
                return RowType.HEADER

        # 数据行判断
        # 1. 数字占比高
        if features.number_ratio > 0.3:
            return RowType.DATA

        # 2. 有序号
        if features.has_sequence_number:
            return RowType.DATA

        # 3. 文本很多但关键字很少（如姓名、地址等）
        if features.text_ratio > 0.6 and features.keyword_ratio < 0.2:
            return RowType.DATA

        return RowType.UNKNOWN
    
    def _has_significant_data_type_transition(self, worksheet: Any, header_row: int, data_row: int, max_col: int) -> bool:
        """检查是否存在显著的数据类型转换（从表头到数据）"""
        if header_row == data_row:
            return False
        
        transition_count = 0
        valid_compare_count = 0
        
        for col in range(1, max_col + 1):
            header_value = self._get_cell_value(worksheet.cell(header_row, col))
            data_value = self._get_cell_value(worksheet.cell(data_row, col))
            
            # 两个都要有值才比较
            if not (header_value and str(header_value).strip() and data_value and str(data_value).strip()):
                continue
            
            valid_compare_count += 1
            
            header_type = self._get_value_type(header_value)
            data_type = self._get_value_type(data_value)
            
            # 表头是文本且包含关键字，数据是数字或日期
            if header_type == ValueType.TEXT and data_type in (ValueType.NUMBER, ValueType.DATE):
                if self._contains_header_keyword(str(header_value)):
                    transition_count += 1
        
        if valid_compare_count == 0:
            return False
        
        # 提高阈值，至少50%的列发生转换才认为是表头到数据的转换
        return (transition_count / valid_compare_count) > 0.5
    
    def _parse_region(self, worksheet: Any, header_start_row: int, max_row: int, max_col: int, max_data_rows: int = None) -> Optional[ExcelRegion]:
        """解析单个数据区域

        Args:
            worksheet: 工作表对象
            header_start_row: 表头起始行
            max_row: 最大行数
            max_col: 最大列数
            max_data_rows: 最多读取的数据行数，None表示读取全部

        Returns:
            解析后的区域数据
        """
        region = ExcelRegion()

        header_info = self._analyze_header_range(worksheet, header_start_row, max_row, max_col)
        if header_info is None:
            return None

        region.head_row_start = header_info.start_row
        region.head_row_end = header_info.end_row
        region.head_data = self._build_header_mapping(worksheet, header_info.start_row, header_info.end_row, max_col)

        if not region.head_data:
            self.logger.warning(f"表头解析失败：行 {header_info.start_row}-{header_info.end_row} 没有找到有效的表头")
            return None

        region.data_row_start = header_info.end_row + 1
        region.data = []
        region.formula = {}

        # 查找数据结束行
        potential_data_end_row = self._find_data_end_row(worksheet, region.data_row_start, max_row, max_col)

        if potential_data_end_row < region.data_row_start:
            # 只有表头没有数据的情况，这是正常的
            region.data_row_end = region.data_row_start - 1
            self.logger.info(f"区域只有表头没有数据：表头行 {region.head_row_start}-{region.head_row_end}，表头数量 {len(region.head_data)}")
            return region

        region.data_row_end = potential_data_end_row

        # 收集数据行（限制行数）
        collected_rows = 0
        for row in range(region.data_row_start, region.data_row_end + 1):
            # 如果设置了最大行数限制且已达到限制，则停止收集
            if max_data_rows is not None and collected_rows >= max_data_rows:
                break

            # 【性能优化】快速预检：先检查第1列是否有值，大部分数据行第1列非空
            first_cell_val = worksheet.cell(row, 1).value
            if first_cell_val is not None and str(first_cell_val).strip():
                # 第1列有值，大概率不是空行，直接快速检查汇总行（只查前几列）
                if self._is_summary_row(worksheet, row, max_col):
                    continue
            else:
                # 第1列为空，进行完整检查
                if self._is_empty_row(worksheet, row, max_col):
                    continue

            if self._is_title_row(worksheet, row, max_col):
                continue

            data_row = self._collect_row_data(worksheet, row, max_col, region.head_data, region.formula)

            if data_row and self._has_valid_data(data_row):
                region.data.append(data_row)
                collected_rows += 1

        # 后验证：验证并尝试修正区域
        validated_region, fixes = self.post_validator.validate_and_fix(worksheet, region, max_col)

        if fixes:
            self.logger.debug(f"区域后验证发现问题: {'; '.join(fixes)}")

        return validated_region
    
    def _parse_region_with_manual_header(self, worksheet: Any, header_start_row: int, header_end_row: int,
                                        max_row: int, max_col: int, max_data_rows: int = None) -> Optional[ExcelRegion]:
        """使用手动指定的表头范围解析数据区域
        
        Args:
            worksheet: 工作表对象
            header_start_row: 表头起始行
            header_end_row: 表头结束行
            max_row: 最大行数
            max_col: 最大列数
            max_data_rows: 最多读取的数据行数，None表示读取全部
        
        Returns:
            解析后的区域数据
        """
        region = ExcelRegion()
        
        region.head_row_start = header_start_row
        region.head_row_end = header_end_row
        region.head_data = self._build_header_mapping(worksheet, header_start_row, header_end_row, max_col)

        if not region.head_data:
            self.logger.warning(f"手动指定的表头解析失败：行 {header_start_row}-{header_end_row} 没有找到有效的表头")
            return None

        region.data_row_start = header_end_row + 1
        region.data = []
        region.formula = {}

        # 查找数据结束行
        potential_data_end_row = self._find_data_end_row(worksheet, region.data_row_start, max_row, max_col)

        if potential_data_end_row < region.data_row_start:
            # 只有表头没有数据的情况，这是正常的
            region.data_row_end = region.data_row_start - 1
            self.logger.info(f"手动指定的区域只有表头没有数据：表头行 {region.head_row_start}-{region.head_row_end}，表头数量 {len(region.head_data)}")
            return region
        
        region.data_row_end = potential_data_end_row
        
        # 收集数据行（限制行数）
        collected_rows = 0
        for row in range(region.data_row_start, region.data_row_end + 1):
            # 如果设置了最大行数限制且已达到限制，则停止收集
            if max_data_rows is not None and collected_rows >= max_data_rows:
                break

            # 【性能优化】快速预检
            first_cell_val = worksheet.cell(row, 1).value
            if first_cell_val is not None and str(first_cell_val).strip():
                if self._is_summary_row(worksheet, row, max_col):
                    continue
            else:
                if self._is_empty_row(worksheet, row, max_col):
                    continue

            if self._is_title_row(worksheet, row, max_col):
                continue

            data_row = self._collect_row_data(worksheet, row, max_col, region.head_data, region.formula)

            if data_row and self._has_valid_data(data_row):
                region.data.append(data_row)
                collected_rows += 1
        
        return region
    
    def _find_data_end_row(self, worksheet: Any, start_row: int, max_row: int, max_col: int) -> int:
        """查找数据结束行（性能优化版）"""
        consecutive_empty_rows = 0
        consecutive_header_like_rows = 0
        last_valid_data_row = start_row - 1
        # 【性能优化】连续确认为数据行后，降低检查频率
        confirmed_data_rows = 0
        check_interval = 1  # 初始每行检查

        for row in range(start_row, max_row + 1):
            if self._is_empty_row(worksheet, row, max_col):
                consecutive_empty_rows += 1
                # 连续3个空行，数据区域结束
                if consecutive_empty_rows >= 3:
                    break
                continue

            consecutive_empty_rows = 0

            # 遇到汇总行
            if self._is_summary_row(worksheet, row, max_col):
                if last_valid_data_row >= start_row:
                    # 已有数据行，汇总行标志数据结束
                    return last_valid_data_row
                else:
                    # 汇总行出现在数据区域最前面（如紧跟表头的合计行），跳过它
                    continue

            # 遇到标题行（且不是第一行数据），数据结束
            if self._is_title_row(worksheet, row, max_col) and row > start_row:
                return last_valid_data_row

            # 【性能优化】已确认大量数据行后，降低header_score检查频率
            # 前20行每行检查，之后每50行抽样检查一次
            if confirmed_data_rows >= 20 and (row - start_row) % 50 != 0:
                last_valid_data_row = row
                confirmed_data_rows += 1
                continue

            # 检查是否像表头（提高阈值，避免误判）
            header_score = self.header_rule_engine.calculate_header_score(worksheet, row, max_col)

            # 只有当表头得分非常高（>0.8）且连续出现2行时，才认为是新的表头区域
            if header_score > 0.8:
                consecutive_header_like_rows += 1
                if consecutive_header_like_rows >= 2 and row > start_row + 5:
                    # 额外验证：如果这些"表头样"行与前面的数据行结构相似，则仍是数据行
                    # （数据行可能因合并单元格、关键字匹配等导致header_score偏高）
                    if last_valid_data_row >= start_row:
                        data_features = self.row_analyzer.analyze_row_features(worksheet, last_valid_data_row, max_col)
                        curr_features = self.row_analyzer.analyze_row_features(worksheet, row, max_col)
                        if abs(data_features.number_ratio - curr_features.number_ratio) < 0.3 and \
                           abs(data_features.text_ratio - curr_features.text_ratio) < 0.3:
                            # 结构与数据行相似，仍视为数据行
                            consecutive_header_like_rows = 0
                            last_valid_data_row = row
                            confirmed_data_rows += 1
                            continue
                    # 至少要有5行数据后，才考虑可能是新表头
                    return last_valid_data_row
            else:
                consecutive_header_like_rows = 0
                last_valid_data_row = row
                confirmed_data_rows += 1

        return last_valid_data_row
    
    def _build_header_mapping(self, worksheet: Any, start_row: int, end_row: int, max_col: int) -> Dict[str, str]:
        """构建表头映射（增强版 - 支持多行表头、复合表头、合并单元格）

        优化策略：
        1. 首先使用结构化分析识别所有表头层级
        2. 正确处理垂直和水平合并单元格
        3. 智能合并多行表头（父级-子级关系）
        4. 去除重复值，保持层级顺序
        """
        header_mapping = {}

        # 使用新的结构化分析方法
        column_structure = self._analyze_header_structure(worksheet, start_row, end_row, max_col)

        # 备用：传统方法收集的表头（用于没有被结构化分析覆盖的情况）
        fallback_headers = {col: [] for col in range(1, max_col + 1)}

        # 逐行处理表头区域（作为备用）
        for row in range(start_row, end_row + 1):
            if self._is_title_row(worksheet, row, max_col):
                continue

            row_processed_cols = set()
            is_no_border_row = self._is_no_border_row(worksheet, row, max_col)

            for col in range(1, max_col + 1):
                if col in row_processed_cols:
                    continue

                # 检查物理合并单元格
                physical_merge = self._get_physical_merged_cell_range(worksheet, row, col)
                if physical_merge:
                    merged_value = self._get_merged_cell_value(worksheet, physical_merge)

                    if merged_value:
                        for apply_col in range(physical_merge.min_col, physical_merge.max_col + 1):
                            if apply_col <= max_col:
                                if not fallback_headers[apply_col] or fallback_headers[apply_col][-1] != merged_value:
                                    fallback_headers[apply_col].append(merged_value)
                                row_processed_cols.add(apply_col)

                    for apply_col in range(physical_merge.min_col, physical_merge.max_col + 1):
                        if apply_col <= max_col:
                            row_processed_cols.add(apply_col)
                else:
                    cell_value = self._get_cell_value(worksheet.cell(row, col))

                    if is_no_border_row:
                        if cell_value and str(cell_value).strip():
                            str_value = self._clean_header_string(str(cell_value))
                            if str_value:
                                fallback_headers[col].append(str_value)
                        row_processed_cols.add(col)
                    else:
                        if cell_value and str(cell_value).strip():
                            str_value = self._clean_header_string(str(cell_value))
                            if str_value:
                                visual_range = self._find_actual_merge_range(worksheet, row, col, max_col)
                                for apply_col in range(visual_range['start_col'], visual_range['end_col'] + 1):
                                    if apply_col <= max_col:
                                        if not fallback_headers[apply_col] or fallback_headers[apply_col][-1] != str_value:
                                            fallback_headers[apply_col].append(str_value)
                                        row_processed_cols.add(apply_col)
                        else:
                            row_processed_cols.add(col)

        # 处理没有找到表头的列
        for col in range(1, max_col + 1):
            if not column_structure[col] and (not fallback_headers[col] or all(not h for h in fallback_headers[col])):
                inherited_header = self._find_column_header_recursively(worksheet, col, start_row, end_row, max_col)
                if inherited_header:
                    fallback_headers[col] = [inherited_header]

        # 构建最终的表头映射
        used_headers = set()

        for col in range(1, max_col + 1):
            column_letter = self._get_column_letter(col)

            # 优先使用结构化分析的结果
            if column_structure[col]:
                final_header = self._build_composite_header(column_structure[col])
            else:
                # 使用备用方法的结果
                valid_headers = [h for h in fallback_headers[col] if h]
                # 去重但保持顺序
                seen = set()
                valid_headers = [h for h in valid_headers if not (h in seen or seen.add(h))]
                final_header = "-".join(valid_headers) if valid_headers else ""

            if final_header:
                # 处理重复的表头名称
                if final_header in used_headers:
                    original_header = final_header
                    final_header = f"{original_header}_{column_letter}"
                    suffix = 2
                    while final_header in used_headers:
                        final_header = f"{original_header}_{suffix}"
                        suffix += 1

                header_mapping[final_header] = column_letter
                used_headers.add(final_header)
            else:
                final_header = f"Column_{column_letter}"
                header_mapping[final_header] = column_letter

        return header_mapping
    
    def _collect_row_data(self, worksheet: Any, row: int, max_col: int,
                         head_data: Dict[str, str], formula_dict: Dict[str, str]) -> Dict[str, Any]:
        """收集行数据"""
        data_row = {}
        
        for header, column_letter in head_data.items():
            col = self._get_column_number(column_letter)
            cell = worksheet.cell(row, col)
            
            data_row[column_letter] = self._get_cell_value(cell)
            
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                cell_address = f"{column_letter}{row}"
                formula_dict[cell_address] = cell.value
        
        return data_row
    
    def _has_valid_data(self, data_row: Dict[str, Any]) -> bool:
        """检查数据行是否包含有效数据"""
        if not data_row:
            return False
        
        non_empty_count = sum(1 for v in data_row.values() if v is not None and str(v).strip())
        return non_empty_count > 0
    
    # ==================== 辅助方法 ====================
    
    def _clean_header_string(self, text: str) -> str:
        """清理表头字符串，去除前后空格、Tab、特殊字符等
        
        Args:
            text: 原始表头字符串
            
        Returns:
            清理后的表头字符串
        """
        if not text:
            return ""
        
        # 去除前后空格和Tab
        cleaned = text.strip()
        
        # 去除特殊字符（※、★、●、◆等）
        special_chars = ['※', '★', '☆', '●', '○', '◆', '◇', '■', '□', '▲', '△', '▼', '▽',
                        '◎', '◉', '⊙', '⊕', '⊗', '⊘', '⊚', '⊛', '⊜', '⊝', '⊞', '⊟',
                        '\u3000']  # \u3000 是全角空格

        for char in special_chars:
            cleaned = cleaned.replace(char, '')

        # 将制表符、换行符、回车符替换为空格
        cleaned = cleaned.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
        
        # 再次去除前后空格（因为去除特殊字符后可能产生新的空格）
        cleaned = cleaned.strip()
        
        # 将多个连续空格替换为单个空格
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        return cleaned
    
    def _is_empty_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        """判断是否为空行"""
        for col in range(1, max_col + 1):
            value = self._get_cell_value(worksheet.cell(row, col))
            if value is not None and str(value).strip():
                return False
        return True
    
    def _is_title_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        """判断是否为标题行或说明行（增强版 + 性能优化）"""
        non_empty_cells = 0
        first_value = None
        second_value = None
        has_required_marker = False  # 是否包含必填标记

        for col in range(1, max_col + 1):
            value = self._get_cell_value(worksheet.cell(row, col))
            if value is not None and str(value).strip():
                non_empty_cells += 1
                str_val = str(value).strip()

                # 检查是否包含必填标记（※、*等），如果有多个这样的单元格，很可能是表头
                for marker in self.REQUIRED_FIELD_MARKERS:
                    if marker in str_val:
                        has_required_marker = True
                        break

                if first_value is None:
                    first_value = value
                elif second_value is None:
                    second_value = value

                # 【性能优化】标题行最多1-2个非空单元格
                # 超过3个非空且无必填标记时，不可能是标题行，提前退出
                if non_empty_cells > 3 and not has_required_marker:
                    return False

        # 如果行中有多个非空单元格且包含必填标记，很可能是表头，不是标题行
        if has_required_marker and non_empty_cells >= 3:
            return False

        if not first_value or not isinstance(first_value, str):
            return False

        first_value_str = str(first_value).strip()

        # 情况1: 只有1个非空单元格，包含标题关键字或示例关键字
        if non_empty_cells == 1:
            # 添加示例、样例等垃圾数据关键字
            garbage_keywords = ["示例", "样例", "例子", "example", "sample", "demo"]
            if any(keyword.lower() in first_value_str.lower() for keyword in self.TITLE_KEYWORDS):
                return True
            # 如果只有一个单元格且内容是示例类关键字，也视为标题行（垃圾数据）
            if any(keyword in first_value_str for keyword in garbage_keywords):
                return True

        # 情况2: 第一个单元格包含说明区域关键字（但不是只有必填标记开头）
        # 排除 "※ 员工编号" 这种表头格式
        is_instruction = False
        for keyword in self.INSTRUCTION_KEYWORDS:
            if keyword.lower() in first_value_str.lower():
                is_instruction = True
                break

        if is_instruction:
            # 但如果是 "※ xxx" 格式且后面跟的是表头关键字，则不是说明行
            if first_value_str.startswith('※') or first_value_str.startswith('*'):
                rest = first_value_str.lstrip('※* ').strip()
                if any(keyword.lower() in rest.lower() for keyword in self.HEADER_KEYWORDS):
                    return False
            return True

        # 情况3: 只有1-2个单元格，第一个单元格是"填写说明："类似格式
        if non_empty_cells <= 2:
            if first_value_str.endswith('：') or first_value_str.endswith(':'):
                return True

        # 情况4: 第二个单元格以"数字、"开头（如"1、xxx说明"）
        if second_value and isinstance(second_value, str):
            second_str = str(second_value).strip()
            if re.match(r'^\d+[、.,]\s*.+', second_str):
                return True

        # 情况5: 包含长说明文字（超过50字符且只有1-2个非空单元格）
        if non_empty_cells <= 2 and len(first_value_str) > 50:
            # 但不能是表头关键字
            if not any(keyword.lower() in first_value_str.lower() for keyword in self.HEADER_KEYWORDS):
                return True

        return False

    def _is_instruction_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        """判断是否为说明/指导行

        这些行通常出现在表头之前，包含填写说明、注意事项等
        """
        non_empty_cells = 0
        values = []
        instruction_keyword_count = 0

        for col in range(1, min(10, max_col + 1)):  # 检查前10列
            value = self._get_cell_value(worksheet.cell(row, col))
            if value is not None and str(value).strip():
                non_empty_cells += 1
                str_val = str(value).strip()
                values.append(str_val)

                # 统计包含说明关键字的单元格数
                for keyword in self.INSTRUCTION_KEYWORDS:
                    if keyword.lower() in str_val.lower():
                        instruction_keyword_count += 1
                        break

        if not values:
            return False

        first_value = values[0]

        # 情况1: 第一个单元格包含说明关键字
        for keyword in self.INSTRUCTION_KEYWORDS:
            if keyword.lower() in first_value.lower():
                return True

        # 情况2: 多个单元格都是说明关键字（如"手工填写", "手工填写", "下拉选项"...）
        if non_empty_cells >= 3 and instruction_keyword_count >= non_empty_cells * 0.5:
            return True

        # 情况3: 是编号格式的说明（如"1、xxx" "2、xxx"）
        for val in values:
            if re.match(r'^\d+[、.,]\s*.{10,}', val):  # 数字+分隔符+至少10个字符的说明
                return True

        return False
    
    def _is_summary_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        """判断是否为汇总行（性能优化版 - 使用预编译正则）"""
        for col in range(1, min(5, max_col + 1)):
            value = self._get_cell_value(worksheet.cell(row, col))
            if value and isinstance(value, str):
                value_lower = value.lower().strip()

                # 排除邮箱地址（包含@符号）
                if '@' in value_lower:
                    continue

                for keyword in self.SUMMARY_KEYWORDS:
                    keyword_lower = keyword.lower()

                    # 中文关键字直接匹配
                    if any(ord(c) > 127 for c in keyword):
                        if keyword_lower in value_lower:
                            return True
                    else:
                        # 【性能优化】使用预编译正则
                        pattern = self._summary_en_patterns.get(keyword_lower)
                        if pattern and pattern.search(value_lower):
                            # 额外验证：英文关键字匹配后，检查单元格值是否确实像汇总描述
                            # 如果单元格值包含多个单词（如人名 "Sum YiShou Zhang"），
                            # 且关键字只是其中一个单词，则不是汇总行
                            words = value_lower.split()
                            if len(words) > 2:
                                # 超过2个单词，很可能是人名或其他描述性文本，不是汇总
                                continue
                            # 单元格值过长（超过20字符），关键字只占很小比例，不是汇总
                            if len(value_lower) > 20 and len(keyword_lower) / len(value_lower) < 0.3:
                                continue
                            return True
        return False
    
    def _contains_header_keyword(self, text: str) -> bool:
        """检查字符串是否包含表头关键字"""
        if not text:
            return False
        text_lower = text.lower()
        return any(keyword.lower() in text_lower for keyword in self.HEADER_KEYWORDS)
    
    def _get_cell_value(self, cell: Any) -> Any:
        """获取单元格值"""
        if cell.value is None:
            return None
        if isinstance(cell.value, datetime):
            return cell.value.strftime("%Y-%m-%d")
        elif isinstance(cell.value, (int, float)):
            return cell.value
        else:
            return str(cell.value)
    
    def _get_value_type(self, value: Any) -> ValueType:
        """获取值类型"""
        if value is None:
            return ValueType.EMPTY
        if isinstance(value, str):
            return ValueType.TEXT
        if isinstance(value, datetime):
            return ValueType.DATE
        if self._is_numeric(value):
            return ValueType.NUMBER
        return ValueType.OTHER
    
    def _is_numeric(self, value: Any) -> bool:
        """判断是否为数值类型"""
        return isinstance(value, (int, float, complex))

    def _is_merged_cell(self, worksheet: Any, row: int, col: int) -> bool:
        """检查是否为合并单元格（使用索引优化）"""
        ws_id = id(worksheet)
        cell_index = self._merged_cell_index.get(ws_id)
        if cell_index is not None:
            return (row, col) in cell_index
        # 回退到遍历方式
        cell = worksheet.cell(row, col)
        for merged_range in worksheet.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return True
        return False

    def _get_physical_merged_cell_range(self, worksheet: Any, row: int, col: int):
        """获取物理合并单元格范围（使用索引优化）"""
        try:
            ws_id = id(worksheet)
            cell_index = self._merged_cell_index.get(ws_id)
            if cell_index is not None:
                return cell_index.get((row, col))
            for merged_range in worksheet.merged_cells.ranges:
                if (row >= merged_range.min_row and row <= merged_range.max_row and
                    col >= merged_range.min_col and col <= merged_range.max_col):
                    return merged_range
        except Exception:
            pass
        return None

    def _get_merged_cell_value(self, worksheet: Any, merged_range) -> str:
        """获取合并单元格的值（从合并区域的任意有值单元格获取）"""
        for m_row in range(merged_range.min_row, merged_range.max_row + 1):
            for m_col in range(merged_range.min_col, merged_range.max_col + 1):
                value = self._get_cell_value(worksheet.cell(m_row, m_col))
                if value and str(value).strip():
                    return self._clean_header_string(str(value))
        return ""

    def _analyze_header_structure(self, worksheet: Any, start_row: int, end_row: int, max_col: int) -> Dict[int, List[Dict]]:
        """分析表头结构，识别多行表头和合并单元格的层级关系

        Returns:
            {col: [{'value': str, 'row': int, 'is_parent': bool, 'span_cols': list}]}
        """
        column_structure = {col: [] for col in range(1, max_col + 1)}

        # 收集所有合并单元格信息
        merge_info = {}  # {(row, col): merged_range}
        for merged_range in worksheet.merged_cells.ranges:
            if merged_range.min_row >= start_row and merged_range.min_row <= end_row:
                for r in range(merged_range.min_row, merged_range.max_row + 1):
                    for c in range(merged_range.min_col, merged_range.max_col + 1):
                        merge_info[(r, c)] = merged_range

        # 逐行分析
        processed_merges = set()
        for row in range(start_row, end_row + 1):
            if self._is_title_row(worksheet, row, max_col):
                continue

            for col in range(1, max_col + 1):
                merged_range = merge_info.get((row, col))

                if merged_range:
                    # 避免重复处理同一个合并区域
                    merge_key = (merged_range.min_row, merged_range.min_col,
                                merged_range.max_row, merged_range.max_col)
                    if merge_key in processed_merges:
                        continue
                    processed_merges.add(merge_key)

                    value = self._get_merged_cell_value(worksheet, merged_range)
                    if not value:
                        continue

                    # 判断是否是父级表头（跨多列）
                    is_parent = (merged_range.max_col - merged_range.min_col) > 0
                    span_cols = list(range(merged_range.min_col, merged_range.max_col + 1))

                    # 将值添加到所有覆盖的列
                    for span_col in span_cols:
                        if span_col <= max_col:
                            column_structure[span_col].append({
                                'value': value,
                                'row': row,
                                'is_parent': is_parent,
                                'span_cols': span_cols,
                                'merge_rows': merged_range.max_row - merged_range.min_row + 1
                            })
                else:
                    # 非合并单元格
                    cell_value = self._get_cell_value(worksheet.cell(row, col))
                    if cell_value and str(cell_value).strip():
                        value = self._clean_header_string(str(cell_value))
                        if value:
                            column_structure[col].append({
                                'value': value,
                                'row': row,
                                'is_parent': False,
                                'span_cols': [col],
                                'merge_rows': 1
                            })

        return column_structure

    def _build_composite_header(self, header_parts: List[Dict], separator: str = "-") -> str:
        """构建复合表头名称（多行合并表头优化版）

        智能合并策略：
        1. 按行号排序
        2. 去除重复值
        3. 父级表头放在前面
        4. 【新增】检测展平行（flattened row）：如果最后一行的值是上层行值的拼接，跳过它
           例如: ["养老保险", "公司", "缴费", "公司缴费"] → "养老保险-公司-缴费"（跳过"公司缴费"）
        """
        if not header_parts:
            return ""

        # 按行号排序
        sorted_parts = sorted(header_parts, key=lambda x: x['row'])

        if len(sorted_parts) <= 1:
            return sorted_parts[0]['value'] if sorted_parts else ""

        # 先收集去重的上层部分（排除最后一行）
        last_row = sorted_parts[-1]['row']
        upper_parts = []
        last_row_parts = []
        seen = set()

        for part in sorted_parts:
            value = part['value']
            if part['row'] == last_row:
                last_row_parts.append(value)
            else:
                if value and value not in seen:
                    seen.add(value)
                    upper_parts.append(value)

        # 如果没有上层部分，直接用最后一行
        if not upper_parts:
            final_parts = []
            for v in last_row_parts:
                if v and v not in seen:
                    seen.add(v)
                    final_parts.append(v)
            return separator.join(final_parts)

        # 检测最后一行是否是上层的展平/拼接
        # 策略：如果最后一行的值可以由上层部分中的 2+ 个值拼接组成，则跳过
        for lv in last_row_parts:
            if not lv or lv in seen:
                continue  # 完全重复，已去重

            # 检查是否是上层值的拼接（如 "公司缴费" = "公司" + "缴费"）
            if self._is_concatenation_of(lv, upper_parts):
                continue  # 跳过展平值

            # 检查是否被上层某个值完全包含（如子串）
            if any(lv in up for up in upper_parts):
                continue

            # 不是展平值，保留
            seen.add(lv)
            upper_parts.append(lv)

        return separator.join(upper_parts)

    def _is_concatenation_of(self, target: str, parts: List[str]) -> bool:
        """检查 target 是否可以由 parts 中的 2+ 个值拼接组成

        例如：
        - _is_concatenation_of("公司缴费", ["养老保险", "公司", "缴费"]) → True（"公司"+"缴费"）
        - _is_concatenation_of("养老基数", ["养老保险", "基数"]) → True（"养老"⊂"养老保险", "基数"）
        - _is_concatenation_of("序号", ["序号"]) → False（只有1个，不算拼接）
        """
        if len(parts) < 2:
            return False

        # 方法1：target 恰好是 parts 中连续子集的拼接
        for i in range(len(parts)):
            for j in range(i + 1, len(parts) + 1):
                concat = "".join(parts[i:j])
                if concat == target:
                    return True

        # 方法2：target 可以由 parts 中任意 2 个值的子串拼接
        # 例如："养老基数" = "养老"(from "养老保险") + "基数"
        for i in range(len(parts)):
            for j in range(len(parts)):
                if i == j:
                    continue
                pi, pj = parts[i], parts[j]
                # 检查 target 是否以 pi 的前缀开始，以 pj 结尾
                for prefix_len in range(1, len(pi) + 1):
                    prefix = pi[:prefix_len]
                    if target.startswith(prefix) and target[len(prefix):] == pj:
                        return True
                    if target.startswith(prefix) and target.endswith(pj):
                        middle = target[len(prefix):len(target) - len(pj)]
                        if not middle:  # 恰好是 prefix + pj
                            return True

        return False
    
    def _is_no_border_row(self, worksheet: Any, row: int, max_col: int) -> bool:
        """检查整行是否都没有边框"""
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row, col)
            if self._has_any_border(cell):
                return False
        return True
    
    def _has_any_border(self, cell) -> bool:
        """检查单元格是否有任何边框"""
        try:
            if cell.border:
                return (cell.border.left.style is not None or
                       cell.border.right.style is not None or
                       cell.border.top.style is not None or
                       cell.border.bottom.style is not None)
        except:
            pass
        return False
    
    def _find_actual_merge_range(self, worksheet: Any, row: int, center_col: int, max_col: int) -> dict:
        """基于边框找到实际的合并范围"""
        start_col = center_col
        end_col = center_col
        
        # 首先检查是否有物理合并单元格
        physical_merge = self._get_physical_merged_cell_range(worksheet, row, center_col)
        if physical_merge:
            start_col = physical_merge.min_col
            end_col = physical_merge.max_col
        else:
            # 获取当前单元格的值
            center_value = self._get_cell_value(worksheet.cell(row, center_col))
            has_center_value = center_value and str(center_value).strip()
            
            # 检查整行是否都没有边框
            is_no_border_row = self._is_no_border_row(worksheet, row, max_col)
            
            if is_no_border_row:
                # 如果整行都没有边框，每个有值的单元格独立成为一个区域
                if has_center_value:
                    # 向左查找连续的空单元格
                    for col in range(center_col - 1, 0, -1):
                        value = self._get_cell_value(worksheet.cell(row, col))
                        if value and str(value).strip():
                            break
                        start_col = col
                    
                    # 向右查找连续的空单元格
                    for col in range(center_col + 1, max_col + 1):
                        value = self._get_cell_value(worksheet.cell(row, col))
                        if value and str(value).strip():
                            break
                        end_col = col
                    
                    # 如果左右都没有扩展，则该单元格独立
                    if start_col == center_col and end_col == center_col:
                        return {'start_col': center_col, 'end_col': center_col}
                else:
                    # 当前单元格为空，找到最近的有值单元格
                    nearest_value_col = self._find_nearest_value_column(worksheet, row, center_col, max_col)
                    if nearest_value_col != -1:
                        return self._find_actual_merge_range(worksheet, row, nearest_value_col, max_col)
                    return {'start_col': center_col, 'end_col': center_col}
            else:
                # 有边框的情况，使用边框判断逻辑
                # 向左扫描
                for col in range(center_col, 0, -1):
                    cell = worksheet.cell(row, col)
                    if col == 1:
                        start_col = col
                        break
                    elif self._has_left_border(cell):
                        start_col = col
                        break
                    else:
                        prev_cell = worksheet.cell(row, col - 1)
                        if self._has_right_border(prev_cell):
                            start_col = col
                            break
                
                # 向右扫描
                for col in range(center_col, max_col + 1):
                    cell = worksheet.cell(row, col)
                    if col == max_col:
                        end_col = col
                        break
                    elif self._has_right_border(cell):
                        end_col = col
                        break
                    else:
                        next_cell = worksheet.cell(row, col + 1)
                        if self._has_left_border(next_cell):
                            end_col = col
                            break
        
        return {'start_col': start_col, 'end_col': end_col}
    
    def _find_nearest_value_column(self, worksheet: Any, row: int, col: int, max_col: int) -> int:
        """找到距离指定位置最近的有值列"""
        left_distance = float('inf')
        right_distance = float('inf')
        left_value_col = -1
        right_value_col = -1
        
        # 向左查找
        for c in range(col - 1, 0, -1):
            value = self._get_cell_value(worksheet.cell(row, c))
            if value and str(value).strip():
                left_value_col = c
                left_distance = col - c
                break
        
        # 向右查找
        for c in range(col + 1, max_col + 1):
            value = self._get_cell_value(worksheet.cell(row, c))
            if value and str(value).strip():
                right_value_col = c
                right_distance = c - col
                break
        
        # 返回最近的有值列
        if left_distance <= right_distance and left_value_col != -1:
            return left_value_col
        elif right_value_col != -1:
            return right_value_col
        
        return -1
    
    def _has_left_border(self, cell) -> bool:
        """检查单元格是否有左边框"""
        try:
            return cell.border and cell.border.left.style is not None
        except:
            return False
    
    def _has_right_border(self, cell) -> bool:
        """检查单元格是否有右边框"""
        try:
            return cell.border and cell.border.right.style is not None
        except:
            return False
    
    def _find_column_header_recursively(self, worksheet: Any, col: int,
                                       header_start_row: int, header_end_row: int, max_col: int) -> str:
        """递归查找列的表头值（增强版 - 更好地处理垂直合并单元格）"""
        header_parts = []
        processed_merge_ranges = set()  # 避免重复处理相同的合并区域

        # 从上到下遍历表头区域
        for row in range(header_start_row, header_end_row + 1):
            if self._is_title_row(worksheet, row, max_col):
                continue

            header_value = None

            # 首先检查是否有合并单元格覆盖当前列
            merged_range = self._get_physical_merged_cell_range(worksheet, row, col)
            if merged_range:
                # 创建合并区域的唯一标识
                merge_key = (merged_range.min_row, merged_range.min_col,
                            merged_range.max_row, merged_range.max_col)

                # 避免重复处理同一个合并区域
                if merge_key not in processed_merge_ranges:
                    processed_merge_ranges.add(merge_key)
                    header_value = self._get_merged_cell_value(worksheet, merged_range)

                    # 如果是垂直合并（跨多行），需要跳过后续行
                    if merged_range.max_row > row:
                        # 已经处理了这个合并区域，后续行会被跳过
                        pass
            else:
                # 没有合并单元格，直接获取单元格值
                direct_value = self._get_cell_value(worksheet.cell(row, col))
                if direct_value and str(direct_value).strip():
                    header_value = self._clean_header_string(str(direct_value))

            # 如果找到值且不重复，添加到列表
            if header_value:
                if not header_parts or header_parts[-1] != header_value:
                    header_parts.append(header_value)

        # 如果还是没找到，尝试从相邻列继承
        if not header_parts:
            header_parts = self._inherit_from_adjacent_columns(worksheet, col, header_start_row, header_end_row, max_col)

        return "-".join(header_parts) if header_parts else ""
    
    def _inherit_from_adjacent_columns(self, worksheet: Any, col: int,
                                      header_start_row: int, header_end_row: int, max_col: int) -> list:
        """从相邻列继承表头信息（增强版 - 支持左右方向继承）"""
        inherited_headers = []
        processed_merges = set()

        # 查找包含当前列的合并区域
        for row in range(header_start_row, header_end_row + 1):
            merged_range = self._get_physical_merged_cell_range(worksheet, row, col)
            if merged_range:
                merge_key = (merged_range.min_row, merged_range.min_col,
                            merged_range.max_row, merged_range.max_col)
                if merge_key not in processed_merges:
                    processed_merges.add(merge_key)
                    value = self._get_merged_cell_value(worksheet, merged_range)
                    if value and value not in inherited_headers:
                        inherited_headers.append(value)

        # 如果还是没找到，尝试从左边相邻列继承
        if not inherited_headers and col > 1:
            left_col = col - 1
            for row in range(header_start_row, header_end_row + 1):
                # 检查左边列是否有跨越到当前列的合并单元格
                merged_range = self._get_physical_merged_cell_range(worksheet, row, left_col)
                if merged_range and merged_range.max_col >= col:
                    value = self._get_merged_cell_value(worksheet, merged_range)
                    if value and value not in inherited_headers:
                        inherited_headers.append(value)

        return inherited_headers
    
    def _get_column_letter(self, column_number: int) -> str:
        """获取列字母"""
        column_letter = ""
        while column_number > 0:
            modulo = (column_number - 1) % 26
            column_letter = chr(ord('A') + modulo) + column_letter
            column_number = (column_number - modulo) // 26
        return column_letter
    
    def _get_column_number(self, column_letter: str) -> int:
        """获取列编号"""
        result = 0
        for char in column_letter:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


# ==================== 使用示例 ====================

if __name__ == "__main__":
    # 使用示例
    parser = IntelligentExcelParser()
    
    # 解析Excel文件
    results = parser.parse_excel_file("example.xlsx")
    
    # 输出结果
    for sheet_data in results:
        print(f"\n=== Sheet: {sheet_data.sheet_name} ===")
        print(f"找到 {len(sheet_data.regions)} 个数据区域\n")
        
        for i, region in enumerate(sheet_data.regions, 1):
            print(f"区域 {i}:")
            print(f"  表头行: {region.head_row_start} - {region.head_row_end}")
            print(f"  数据行: {region.data_row_start} - {region.data_row_end}")
            print(f"  表头映射: {region.head_data}")
            print(f"  数据行数: {len(region.data)}")
            
            if region.data:
                print(f"  第一行数据示例: {region.data[0]}")
            print()