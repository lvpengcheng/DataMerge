"""源文件上传规范化

针对"数值被误设成日期格式"的源文件：某些单元格底层是数字，却被设了日期/自定义日期
格式。Aspose 读取时会把它当 DateTime 返回，再经 Excel 1900 闰年 bug 转换，
数字会差一天（如 60.74 被读成 59.74，小数 .74 保留、整数差 1）。

本模块在【上传阶段】对源文件做一次规范化：把"被误设成日期格式的数字单元格"
的数字格式重置为常规（General），使后续解析读到正确数字。真实业务日期（1950 年后，
即 Excel 序列号 >= 18262）保持不变。不改动解析器读值逻辑。
"""

import logging

logger = logging.getLogger(__name__)

# Excel 序列号 18262 = 1950-01-01。小于此值的"日期"几乎不可能是真实业务日期，
# 基本是被误设成日期格式的数字（社保费、比例等）。
_REAL_DATE_SERIAL_MIN = 18262

# 列名含这些关键词 → 判定为真实日期列，即使序列号偏小也绝不改（保护早于1950的生日等）
_DATE_HEADER_KW = ('日期', '年月', '出生', '生日', '时间', 'date', 'birth', '年 月', '日 期')


def _date_protected_columns(ws, scan_rows: int = 25) -> set:
    """扫描表头区域，返回"列名含日期关键词"的列索引集合。

    这些列即便序列号偏小也不规范化，避免误改真实日期列（如 1950 年前的出生日期）。
    只看文本型单元格（表头通常是文本），按列收集，命中关键词即保护该列。
    """
    protected = set()
    try:
        from Aspose.Cells import CellValueType
        cells = ws.Cells
        max_col = cells.MaxDataColumn
        max_row = min(cells.MaxDataRow, scan_rows)
        if max_col is None or max_col < 0:
            return protected
        for col in range(0, max_col + 1):
            for row in range(0, max_row + 1):
                cell = cells.get_Item(row, col)
                try:
                    if cell.Type != CellValueType.IsString:
                        continue
                    text = str(cell.StringValue or "").strip().lower()
                except Exception:
                    continue
                if text and any(kw in text for kw in _DATE_HEADER_KW):
                    protected.add(col)
                    break
    except Exception:
        pass
    return protected


def normalize_misformatted_dates(file_path: str, out_path: str = None) -> int:
    """将源文件中"被误设成日期格式的数字单元格"格式重置为常规。

    判定（三个条件同时满足才重置）：
      1. cell.Type == IsDateTime（普通数字/文本列是 IsNumeric/IsString，永不命中）
      2. Excel 序列号 < 18262（1950-01-01，真实业务日期不会这么早）
      3. 该列列名不含日期关键词（保护真实日期列，即便有 1950 年前的日期）

    Args:
        file_path: 源文件路径
        out_path:  输出路径，None 表示原地覆盖
    Returns:
        被规范化的单元格数量（0 表示无需改动）
    """
    try:
        from Aspose.Cells import Workbook, CellValueType
        import aspose_init
        aspose_init.ensure_license()
    except Exception as e:
        logger.warning(f"[normalize] Aspose 不可用，跳过规范化: {e}")
        return 0

    try:
        wb = Workbook(file_path)
    except Exception as e:
        logger.warning(f"[normalize] 打开失败，跳过: {file_path} - {e}")
        return 0

    fixed = 0
    try:
        # 先重算公式，使"公式结果被设成日期格式"的单元格类型落定为 IsDateTime，便于识别
        try:
            wb.CalculateFormula()
        except Exception as _ce:
            logger.warning(f"[normalize] CalculateFormula 跳过: {file_path} - {_ce}")

        for i in range(wb.Worksheets.Count):
            ws = wb.Worksheets[i]
            cells = ws.Cells
            protected_cols = _date_protected_columns(ws)   # 日期列保护名单
            it = cells.GetEnumerator()
            while it.MoveNext():
                cell = it.Current
                try:
                    if cell.Type != CellValueType.IsDateTime:
                        continue
                    serial = float(cell.DoubleValue)
                except Exception:
                    continue
                # 真实业务日期（1950 年后）不动
                if serial >= _REAL_DATE_SERIAL_MIN:
                    continue
                # 列名是日期列 → 保护，不动（即便序列号偏小，可能是早期生日）
                try:
                    if cell.Column in protected_cols:
                        continue
                except Exception:
                    pass
                # 被误设成日期格式的数字 → 重置为常规格式
                try:
                    style = cell.GetStyle()
                    style.Number = 0          # 0 = General/常规
                    style.Custom = ""
                    cell.SetStyle(style)
                    fixed += 1
                except Exception:
                    continue

        if fixed > 0:
            wb.Save(out_path or file_path)
            logger.info(f"[normalize] 已规范化 {fixed} 个误设日期格式的数字单元格: {file_path}")
    except Exception as e:
        logger.warning(f"[normalize] 规范化过程异常，按原文件处理: {file_path} - {e}")
        return 0
    finally:
        try:
            wb.Dispose()
        except Exception:
            pass

    return fixed
