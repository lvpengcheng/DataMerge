"""
Aspose.Cells 工具类 (.NET via pythonnet)
集中管理：读写Excel、模板填充、格式转换（PDF）、加密/解密、样式设置等
依赖 aspose_init 模块完成全局初始化（pythonnet + 许可证）

功能列表：
  ── 读取 ──────────────────────────────────────────────
  read_excel(path, sheet, header_row, lowercase_cols, password)
      读取指定 Sheet 为 DataFrame
  read_all_sheets(path, header_row, password)
      读取所有 Sheet（跳过评估版水印 Sheet）

  ── 写入 ──────────────────────────────────────────────
  dataframe_to_excel(df, path, sheet_name, bold_header)
      DataFrame 直接写入 xlsx，支持加粗表头

  ── 格式转换 ──────────────────────────────────────────
  save_as(wb_or_path, output_path)
      Workbook 按扩展名保存（xlsx / pdf / html / csv）
  convert_to_pdf / convert_to_encrypted_pdf / convert_format
      专用格式转换
  sheet_to_image(ws_or_path, output_path)
      单个 Sheet 导出为 PNG / JPEG 图片

  ── 样式与格式 ────────────────────────────────────────
  set_header_style(ws, row, col_count, bold, bg_color_rgb)
  add_conditional_format(ws, cell_range, threshold)
  freeze_header(ws, row_count, col_count)
  enable_auto_filter(ws, header_range)
  set_column_best_fit(ws, col_indices)

  ── 图表 ──────────────────────────────────────────────
  add_chart(ws, data_range, ...)

  ── 模板与水印 ────────────────────────────────────────
  generate_from_template(output_path, template_path, data, ...)
      模板填充生成文件 — 支持 fill / block / zip 三种模式
  add_excel_watermark(wb, text)

  ── 加密 / 解密 ──────────────────────────────────────
  is_encrypted / encrypt_excel / decrypt_excel / write_protect_excel

  ── 工具 ──────────────────────────────────────────────
  change_error_message(msg)      Aspose 英文异常转中文
"""

import io
import logging
import os
import re
import tempfile
import zipfile
from typing import Optional, Union, Dict, List

import pandas as pd

import aspose_init  # noqa: F401 — 确保 Aspose 已初始化

from Aspose.Cells import (  # type: ignore
    Workbook, SaveFormat, PdfSaveOptions, LoadOptions, EncryptionType,
    FileFormatUtil, HtmlSaveOptions, BackgroundType,
    CellArea, FormatConditionType, OperatorType,
)
from Aspose.Cells.Rendering.PdfSecurity import PdfSecurityOptions  # type: ignore

logger = logging.getLogger(__name__)


def _licensed_workbook(*args, **kwargs):
    """创建 Workbook 前自动确保许可证有效，防止 .NET GC 回收许可证"""
    aspose_init.ensure_license()
    return Workbook(*args, **kwargs)


# ═══════════════════════════════════════════════════════
# 内部工具
# ═══════════════════════════════════════════════════════

def _get_color(r: int, g: int, b: int):
    """获取 .NET System.Drawing.Color 对象"""
    from System.Drawing import Color
    return Color.FromArgb(int(r), int(g), int(b))


def _ext_save_format(path: str):
    """按文件扩展名返回 Aspose SaveFormat 枚举值"""
    ext = os.path.splitext(path)[1].lower()
    return {
        ".xlsx": SaveFormat.Xlsx,
        ".xls": SaveFormat.Excel97To2003,
        ".pdf": SaveFormat.Pdf,
        ".csv": SaveFormat.Csv,
        ".html": None,
    }.get(ext, SaveFormat.Xlsx)


def _sheet_to_dataframe(
    ws,
    header_row: int = 0,
    lowercase_cols: bool = False,
) -> pd.DataFrame:
    """将 .NET Worksheet 转换为 DataFrame（内部复用）"""
    cells = ws.Cells
    max_row = cells.MaxDataRow       # 0-indexed
    max_col = cells.MaxDataColumn    # 0-indexed

    if max_row < header_row + 1 or max_col < 0:
        return pd.DataFrame()

    headers = []
    for c in range(max_col + 1):
        val = str(cells[header_row, c].StringValue or "").strip()
        headers.append(val)
    if lowercase_cols:
        headers = [h.lower() for h in headers]

    rows = []
    for r in range(header_row + 1, max_row + 1):
        row = [cells[r, c].Value for c in range(max_col + 1)]
        rows.append(row)

    return pd.DataFrame(rows, columns=headers)


_TYPE_CHANGES = {
    "string": "文本", "int": "数字", "long": "数字",
    "datetime": "日期", "date": "日期", "boolean": "真/假",
}

_ERROR_PATTERNS = [
    (
        r"The value of the cell (\w+) should not be a (\w+) value",
        "Excel表格中的单元格[{0}]不可以是{1}内容",
    ),
]


# ═══════════════════════════════════════════════════════
# 读取
# ═══════════════════════════════════════════════════════

def read_excel(
    path: str,
    sheet: Union[int, str] = 0,
    header_row: int = 0,
    lowercase_cols: bool = False,
    password: str = None,
) -> pd.DataFrame:
    """
    读取 Excel 指定 Sheet 为 DataFrame。

    Args:
        path:           Excel 文件路径
        sheet:          Sheet 索引（int）或名称（str），默认第一个
        header_row:     列名所在行号（0-indexed），默认 0
        lowercase_cols: 是否将列名转为小写
        password:       打开密码（可选）

    Returns:
        pd.DataFrame
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"未找到文件: {path}")

    try:
        if password:
            opts = LoadOptions()
            opts.Password = password
            wb = _licensed_workbook(path, opts)
        else:
            wb = _licensed_workbook(path)

        ws = wb.Worksheets[sheet]
        return _sheet_to_dataframe(ws, header_row, lowercase_cols)
    except Exception as ex:
        raise Exception(change_error_message(str(ex))) from ex


def read_all_sheets(
    path: str,
    header_row: int = 0,
    password: str = None,
) -> Dict[str, pd.DataFrame]:
    """
    读取 Excel 所有 Sheet，返回 {sheet_name: DataFrame}。
    自动跳过 Aspose 评估版附加的 Evaluation Sheet。

    Args:
        path:       Excel 文件路径
        header_row: 列名所在行号（0-indexed）
        password:   打开密码（可选）

    Returns:
        dict[str, pd.DataFrame]
    """
    if password:
        opts = LoadOptions()
        opts.Password = password
        wb = _licensed_workbook(path, opts)
    else:
        wb = _licensed_workbook(path)

    result: Dict[str, pd.DataFrame] = {}
    for i in range(wb.Worksheets.Count):
        ws = wb.Worksheets[i]
        name = ws.Name or ""
        if "Evaluation" in name:
            continue
        result[name] = _sheet_to_dataframe(ws, header_row)
    return result


def read_all_sheets_calculated(
    path: str,
    header_row: int = 0,
    password: str = None,
) -> Dict[str, pd.DataFrame]:
    """
    读取 Excel 所有 Sheet，先 CalculateFormula() 强制计算公式，
    再提取计算后的值。适用于含公式的计算结果文件。

    Args:
        path:       Excel 文件路径
        header_row: 列名所在行号（0-indexed）
        password:   打开密码（可选）

    Returns:
        dict[str, pd.DataFrame]
    """
    if password:
        opts = LoadOptions()
        opts.Password = password
        wb = _licensed_workbook(path, opts)
    else:
        wb = _licensed_workbook(path)

    wb.CalculateFormula()

    result: Dict[str, pd.DataFrame] = {}
    for i in range(wb.Worksheets.Count):
        ws = wb.Worksheets[i]
        name = ws.Name or ""
        if "Evaluation" in name:
            continue
        result[name] = _sheet_to_dataframe(ws, header_row)
    return result


# ═══════════════════════════════════════════════════════
# 写入
# ═══════════════════════════════════════════════════════

def dataframe_to_excel(
    df: pd.DataFrame,
    path: str,
    sheet_name: str = "Sheet1",
    bold_header: bool = True,
) -> str:
    """
    将 DataFrame 写入 xlsx 文件。

    Args:
        df:          数据源 DataFrame
        path:        输出 xlsx 路径
        sheet_name:  Sheet 名称
        bold_header: 是否加粗表头行

    Returns:
        实际保存路径
    """
    wb = _licensed_workbook()
    ws = wb.Worksheets[0]
    ws.Name = sheet_name

    # 写表头
    for c, col in enumerate(df.columns):
        cell = ws.Cells[0, c]
        cell.PutValue(str(col))
        if bold_header:
            style = cell.GetStyle()
            style.Font.IsBold = True
            cell.SetStyle(style)

    # 写数据
    for r_idx, row in enumerate(df.itertuples(index=False), start=1):
        for c_idx, v in enumerate(row):
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                ws.Cells[r_idx, c_idx].PutValue(v)

    wb.Save(path, SaveFormat.Xlsx)
    logger.info(f"DataFrame 写入完成: {path} ({len(df)} 行)")
    return path


# ═══════════════════════════════════════════════════════
# 格式转换
# ═══════════════════════════════════════════════════════

def flatten_formulas_to_values(src_path: str, dst_path: str = None, password: str = None) -> str:
    """另存一份"保留全部格式与数据、仅清除关联公式"的 Excel。

    用途：Sheet 拆分前的预处理。原样保留 Aspose 加载到的所有信息（单元格样式、
    合并、行高/列宽、批注、图表、表格等），仅将公式 cell 替换为计算后的字面值，
    避免拆分时跨 sheet 引用 / 同 sheet 行号偏移导致 #REF!。

    Args:
        src_path: 源 Excel 路径
        dst_path: 目标路径（不传时直接覆盖源文件）
        password: 打开密码（可选）

    Returns:
        实际保存的目标路径
    """
    if dst_path is None:
        dst_path = src_path

    if password:
        opts = LoadOptions()
        opts.Password = password
        wb = _licensed_workbook(src_path, opts)
    else:
        wb = _licensed_workbook(src_path)

    # 递归求值所有公式（含跨 sheet / 跨工作簿外部引用）
    wb.CalculateFormula()

    # 遍历每个 sheet 的真实使用范围，把公式 cell 替换为字面值
    # 保留：cell.Style、行高、列宽、合并、批注、图表等一切非公式属性
    converted = 0
    for si in range(wb.Worksheets.Count):
        ws = wb.Worksheets[si]
        cells = ws.Cells
        try:
            # 用 MaxDataRow/MaxDataColumn（仅数据范围），不用 MaxRow/MaxColumn：
            # 后者含"空白但有样式"的行列（整行/整列刷底色、边框、筛选区等），
            # 会把范围虚高到几十上百列/行，导致这里的双重遍历空转甚至内存溢出。
            # 公式一定落在数据范围内，故用 MaxData* 不会漏掉任何公式。
            max_row = cells.MaxDataRow
            max_col = cells.MaxDataColumn
        except Exception:
            continue
        if max_row is None or max_row < 0 or max_col is None or max_col < 0:
            continue
        for r in range(max_row + 1):
            for c in range(max_col + 1):
                cell = cells.GetCell(r, c)
                if cell is None or not cell.IsFormula:
                    continue
                try:
                    val = cell.Value  # 已计算的字面值
                except Exception:
                    val = None
                # 公式求值失败时（cell.Value 抛错或为 None）写入空串，避免残留破公式
                cell.PutValue(val if val is not None else "")
                converted += 1

    fmt = _ext_save_format(dst_path) or SaveFormat.Xlsx
    wb.Save(dst_path, fmt)
    return dst_path


def save_as(wb_or_path, output_path: str) -> str:
    """
    Workbook（或文件路径）按目标文件扩展名保存。
    支持：.xlsx / .xls / .pdf / .csv / .html

    Args:
        wb_or_path:  Workbook 对象 或 Excel 文件路径
        output_path: 目标文件路径（扩展名决定格式）

    Returns:
        output_path
    """
    wb = wb_or_path if isinstance(wb_or_path, Workbook) else _licensed_workbook(str(wb_or_path))
    fmt = _ext_save_format(output_path)

    if fmt is None:  # HTML 特殊处理
        hso = HtmlSaveOptions()
        hso.ExportImagesAsBase64 = True
        hso.ExportHiddenWorksheet = False
        hso.ExportActiveWorksheetOnly = True
        wb.CalculateFormula()
        wb.Save(output_path, hso)
    else:
        if fmt == SaveFormat.Pdf:
            wb.CalculateFormula()
        wb.Save(output_path, fmt)

    logger.info(f"文件保存完成: {output_path}")
    return output_path


def sheet_to_image(
    ws_or_path,
    output_path: str,
    sheet_index: int = 0,
    image_format: str = "png",
) -> str:
    """
    将 Sheet 导出为图片（PNG/JPEG）。

    Args:
        ws_or_path:    .NET Worksheet 对象 或 Excel 文件路径
        output_path:   输出图片路径
        sheet_index:   当传入文件路径时使用的 Sheet 索引
        image_format:  'png'（默认）或 'jpeg'

    Returns:
        output_path
    """
    from Aspose.Cells.Drawing import ImageType as AsposeImageType
    from Aspose.Cells.Rendering import ImageOrPrintOptions, SheetRender

    if isinstance(ws_or_path, str):
        wb = _licensed_workbook(ws_or_path)
        ws = wb.Worksheets[sheet_index]
    else:
        ws = ws_or_path

    opts = ImageOrPrintOptions()
    fmt_map = {
        "png": AsposeImageType.Png,
        "jpeg": AsposeImageType.Jpeg,
        "jpg": AsposeImageType.Jpeg,
    }
    opts.ImageType = fmt_map.get(image_format.lower(), AsposeImageType.Png)

    sr = SheetRender(ws, opts)
    sr.ToImage(0, output_path)

    logger.info(f"Sheet 导出图片完成: {output_path}")
    return output_path


def convert_to_pdf(input_path: str, output_path: str = None, active_sheet_only: bool = True) -> str:
    """Excel 转 PDF

    Args:
        active_sheet_only: True=只导出激活sheet（默认），False=导出所有sheet
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".pdf")

    wb = _licensed_workbook(input_path)

    if active_sheet_only:
        active_index = wb.Worksheets.ActiveSheetIndex
        for i in range(wb.Worksheets.Count):
            if i != active_index:
                wb.Worksheets[i].IsVisible = False

    opts = PdfSaveOptions()
    opts.CalculateFormula = True
    wb.Save(output_path, opts)

    logger.info(f"PDF转换完成: {input_path} -> {output_path} (active_only={active_sheet_only})")
    return output_path


def convert_to_encrypted_pdf(
    input_path: str,
    output_path: str = None,
    user_password: str = "123456",
    owner_password: str = "admin",
    active_sheet_only: bool = True,
) -> str:
    """Excel 转加密 PDF"""
    if not output_path:
        output_path = tempfile.mktemp(suffix=".pdf")

    wb = _licensed_workbook(input_path)

    if active_sheet_only:
        active_index = wb.Worksheets.ActiveSheetIndex
        for i in range(wb.Worksheets.Count):
            if i != active_index:
                wb.Worksheets[i].IsVisible = False

    opts = PdfSaveOptions()
    opts.CalculateFormula = True

    security = PdfSecurityOptions()
    security.UserPassword = user_password
    security.OwnerPassword = owner_password
    security.PrintPermission = True
    security.FullQualityPrintPermission = True
    opts.SecurityOptions = security

    wb.Save(output_path, opts)

    logger.info(f"加密PDF转换完成: {input_path} -> {output_path}")
    return output_path


def convert_format(
    input_path: str,
    output_format: str,
    output_path: str = None,
) -> str:
    """通用格式转换（pdf, csv, html, xlsx, xls）"""
    format_map = {
        "pdf": (SaveFormat.Pdf, ".pdf"),
        "csv": (SaveFormat.Csv, ".csv"),
        "html": (SaveFormat.Html, ".html"),
        "xlsx": (SaveFormat.Xlsx, ".xlsx"),
        "xls": (SaveFormat.Excel97To2003, ".xls"),
    }

    fmt = output_format.lower()
    if fmt not in format_map:
        raise ValueError(f"不支持的格式: {output_format}，可选: {list(format_map.keys())}")

    save_fmt, ext = format_map[fmt]
    if not output_path:
        output_path = tempfile.mktemp(suffix=ext)

    wb = _licensed_workbook(input_path)
    wb.Save(output_path, save_fmt)

    logger.info(f"格式转换完成: {input_path} -> {output_path} ({fmt})")
    return output_path


# ═══════════════════════════════════════════════════════
# 样式与格式
# ═══════════════════════════════════════════════════════

def set_header_style(
    ws,
    row: int = 0,
    col_count: int = 1,
    bold: bool = True,
    bg_color_rgb: tuple = (70, 130, 180),
) -> None:
    """
    设置指定行为表头样式：加粗、背景色、白色字体、居中对齐、底部边框。

    Args:
        ws:           .NET Worksheet 对象
        row:          行号（0-indexed）
        col_count:    列数
        bold:         是否加粗
        bg_color_rgb: 背景色 (R, G, B) 元组，默认 steel blue
    """
    from Aspose.Cells import TextAlignmentType, BorderType, CellBorderType

    bg = _get_color(*bg_color_rgb)
    white = _get_color(255, 255, 255)

    for c in range(col_count):
        cell = ws.Cells[row, c]
        style = cell.GetStyle()
        style.Font.IsBold = bold
        style.Font.Color = white
        style.ForegroundColor = bg
        style.Pattern = BackgroundType.Solid
        style.HorizontalAlignment = TextAlignmentType.Center
        style.VerticalAlignment = TextAlignmentType.Center
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin
        style.Borders[BorderType.BottomBorder].Color = white
        cell.SetStyle(style)


def add_conditional_format(
    ws,
    cell_range: str,
    threshold: float,
    highlight_color_rgb: tuple = (144, 238, 144),
) -> None:
    """
    对指定单元格区域添加条件格式：值大于 threshold 时高亮背景色。

    Args:
        ws:                  .NET Worksheet 对象
        cell_range:          如 'A2:D10'
        threshold:           触发条件的数值阈值
        highlight_color_rgb: 高亮背景色 (R, G, B)，默认 light green
    """
    r = ws.Cells.CreateRange(cell_range)
    area = CellArea()
    area.StartRow = r.FirstRow
    area.EndRow = r.FirstRow + r.RowCount - 1
    area.StartColumn = r.FirstColumn
    area.EndColumn = r.FirstColumn + r.ColumnCount - 1

    fca = ws.ConditionalFormattings
    idx = fca.Add()
    fc = fca[idx]
    fc.AddArea(area)
    cond_idx = fc.AddCondition(
        FormatConditionType.CellValue,
        OperatorType.GreaterThan,
        str(threshold),
        None,
    )
    cond = fc[cond_idx]
    style = cond.Style
    style.ForegroundColor = _get_color(*highlight_color_rgb)
    style.Pattern = BackgroundType.Solid
    cond.Style = style


def freeze_header(ws, row_count: int = 1, col_count: int = 0) -> None:
    """
    冻结前 N 行和前 M 列。

    Args:
        ws:        .NET Worksheet 对象
        row_count: 冻结行数，默认 1（冻结首行）
        col_count: 冻结列数，默认 0
    """
    ws.FreezePanes(row_count, col_count, row_count, col_count)


def enable_auto_filter(ws, header_range: str) -> None:
    """
    启用自动筛选。

    Args:
        ws:           .NET Worksheet 对象
        header_range: 表头行范围，如 'A1:E1'
    """
    ws.AutoFilter.Range = header_range


def set_column_best_fit(ws, col_indices: list = None) -> None:
    """
    自动调整列宽。如果未指定 col_indices，则自动适应所有列。

    Args:
        ws:          .NET Worksheet 对象
        col_indices: 列索引列表（0-indexed），如 [0, 1, 2]；None 则全部自适应
    """
    if col_indices:
        for c in col_indices:
            ws.AutoFitColumn(c)
    else:
        ws.AutoFitColumns()


# ═══════════════════════════════════════════════════════
# 图表
# ═══════════════════════════════════════════════════════

def add_chart(
    ws,
    data_range: str,
    upper_left_row: int,
    upper_left_col: int,
    lower_right_row: int,
    lower_right_col: int,
    chart_type: str = "column",
    title: str = "",
):
    """
    在 Sheet 中插入图表。

    Args:
        ws:               .NET Worksheet 对象
        data_range:       数据区域，如 'A1:C5'
        upper_left_row:   图表左上角行（0-indexed）
        upper_left_col:   图表左上角列
        lower_right_row:  图表右下角行
        lower_right_col:  图表右下角列
        chart_type:       图表类型: 'column'/'bar'/'line'/'pie'/'area'/'scatter'
        title:            图表标题

    Returns:
        Chart 对象
    """
    from Aspose.Cells.Charts import ChartType as CT

    type_map = {
        "column": CT.Column, "bar": CT.Bar, "line": CT.Line,
        "pie": CT.Pie, "area": CT.Area, "scatter": CT.Scatter,
    }
    ct = type_map.get(chart_type.lower(), CT.Column)

    idx = ws.Charts.Add(ct, upper_left_row, upper_left_col, lower_right_row, lower_right_col)
    chart = ws.Charts[idx]
    chart.NSeries.Add(data_range, True)
    if title:
        chart.Title.Text = title
        chart.Title.IsVisible = True
    return chart


# ═══════════════════════════════════════════════════════
# 模板与水印
# ═══════════════════════════════════════════════════════

def _fix_smart_marker_spacing(wb) -> int:
    """修复模板中 SmartMarker 标记字段名与修饰符括号之间的空格。

    例如: "&=DT.工号 (group:repeat,skip:1)" → "&=DT.工号(group:repeat,skip:1)"
    空格会导致 Aspose 无法识别修饰符，必须紧跟括号。
    返回修复的标记数量。
    """
    fixed = 0
    pattern = re.compile(r'(&=\S+)\s+(\([^)]+\))')

    for ws_idx in range(wb.Worksheets.Count):
        ws = wb.Worksheets[ws_idx]
        cells = ws.Cells
        max_row = cells.MaxDataRow
        max_col = cells.MaxDataColumn
        if max_row < 0 or max_col < 0:
            continue
        for r in range(max_row + 1):
            for c in range(max_col + 1):
                val = str(cells[r, c].StringValue or "")
                if "&=" in val and "(" in val:
                    new_val = pattern.sub(r'\1\2', val)
                    if new_val != val:
                        cells[r, c].PutValue(new_val)
                        logger.info(f"[SmartMarker] 修复标记空格: '{val}' → '{new_val}'")
                        fixed += 1
    return fixed


def _dataframe_to_datatable(df: pd.DataFrame, table_name: str):
    """将 DataFrame 转为类型明确的 .NET System.Data.DataTable（供 WorkbookDesigner SmartMarker）。

    直接按【实际值】逐列判定数值/文本并构建 typed DataTable，绕开
    Aspose ExportDataTable 的列类型推断（首格空/文本样式残留会让整列变 string 导致数值→文本）：
      - 某列非空值全可作数值(含 numpy 数字 与 普通数值字符串) → System.Double 列，
        让模板单元格自身的数字格式(如 0.00)生效，0 显示成 0.00；
      - 其余列 → System.String 列，工号/前导零/超长卡号/身份证/#N/A 原样保文本。
    """
    import numbers as _numbers
    import re as _re
    try:
        import clr
        clr.AddReference("System.Data")
    except Exception:
        pass
    from System.Data import DataTable
    from System import DBNull, Double as _Double, String as _String

    _NUM_RE = _re.compile(r'^[+-]?\d+(?:\.\d+)?$')
    # 财务表"无金额"占位符：视作空值(None)，不让它把整列数值列拖成文本列
    #  → 这样 O/Q/R… 等金额列仍以真数字写入，SmartMarker 公式可直接计算。
    _PLACEHOLDER_EMPTY = {
        '-', '−', '—', '–', '－', '/', '／', '\\', '＼',
        'n/a', '#n/a', 'na', 'null', 'none', 'nan', '无', '空', '/-',
    }

    def _isna(v):
        if v is None:
            return True
        try:
            return bool(pd.isna(v))
        except (TypeError, ValueError):
            return False

    def _as_number(v):
        """(是否可作数值, Python数值或None)。空值/占位符→(True, None) 不破坏列判定。
        numpy/py 数字直接用；数值字符串在排除前导零/超长位数(工号/卡号/身份证)后转数值。"""
        if _isna(v):
            return (True, None)
        if isinstance(v, bool):
            return (False, None)
        if isinstance(v, _numbers.Number):        # 含 numpy.int64/float64（isinstance int/float 会漏）
            f = float(v)
            return (True, int(f) if f.is_integer() else f)
        s = str(v).strip().replace(',', '').replace('，', '')
        if s == '':
            return (True, None)
        if s.lower() in _PLACEHOLDER_EMPTY:          # -、／、N/A、无… 等无金额占位符 → 空
            return (True, None)
        if not _NUM_RE.match(s):
            return (False, None)
        intpart = s.lstrip('+-').split('.')[0]
        if len(intpart) > 1 and intpart[0] == '0':   # 前导零 → 编码/工号，保文本
            return (False, None)
        if len(intpart) > 15:                        # 超长 → 卡号/身份证，保文本(丢精度)
            return (False, None)
        f = float(s)
        return (True, int(f) if f.is_integer() else f)

    n_rows = len(df)
    n_cols = len(df.columns)

    # 逐列判定 + 缓存转换值：非空值全可作数值 且 至少含一个真数值 → 数值列
    col_is_num = [False] * n_cols
    col_num_cache = [None] * n_cols
    for c in range(n_cols):
        conv, ok, has_real = [], True, False
        for r in range(n_rows):
            isnum, num = _as_number(df.iloc[r, c])
            if not isnum:
                ok = False
                break
            conv.append(num)
            if num is not None:
                has_real = True
        col_is_num[c] = ok and has_real
        if col_is_num[c]:
            col_num_cache[c] = conv

    dt = DataTable(table_name)
    for c, col_name in enumerate(df.columns):
        dt.Columns.Add(str(col_name), _Double if col_is_num[c] else _String)

    for r in range(n_rows):
        row = dt.NewRow()
        for c in range(n_cols):
            if col_is_num[c]:
                num = col_num_cache[c][r]
                row[c] = DBNull.Value if num is None else float(num)
            else:
                v = df.iloc[r, c]
                # 空值 / 纯空白串统一写 DBNull（真空白）：文本空串 "" 进入数值公式
                # (如 =O+X) 会被当文本操作数导致 #VALUE!，真空白才会被当 0 参与计算。
                if _isna(v) or (isinstance(v, str) and v.strip() == ''):
                    row[c] = DBNull.Value
                else:
                    row[c] = v if isinstance(v, str) else str(v)
        dt.Rows.Add(row)

    num_cols = [str(df.columns[i]) for i in range(n_cols) if col_is_num[i]]
    logger.info(f"[DataTable] {table_name}: {dt.Rows.Count} rows x {dt.Columns.Count} cols; 数值列={num_cols}")
    return dt


# ── sheet 名占位符 ─────────────────────────────────────

_DT_SHEET_TOKEN = re.compile(r"^DT(\d*)$", re.IGNORECASE)


def _rename_sheets_by_token(wb, source_sheet_names=None, sys_vars=None) -> None:
    """把报表模版里的占位符 sheet 名替换成实际名（就地改 wb，不返回）。

    规则（仅单文件模式 fill/block/sheet 调用）：
      - 整名精确匹配 ``^DT(\\d*)$`` → 数据源第 idx 个 sheet 名（DT=idx0、DT1=idx1…）。
        越界则保留原名并告警。
      - 名字里含 ``{`` → 用 sys_vars 的 year/month/date/tenant 做 {token} 子串替换；
        未知 {xxx} 原样保留。
      - 其余不动。
    产出名经 _sanitize_sheet_name 合法化 + 去重；与原名相同则跳过。
    """
    source_sheet_names = source_sheet_names or []
    sys_vars = sys_vars or {}
    try:
        count = wb.Worksheets.Count
    except Exception:
        return

    # 先收集当前所有 sheet 名（去重时排除自身，避免把自己算成"已占用"）
    taken = set()
    for i in range(count):
        try:
            taken.add(str(wb.Worksheets[i].Name))
        except Exception:
            pass

    for i in range(count):
        try:
            ws = wb.Worksheets[i]
            old = str(ws.Name)
        except Exception:
            continue

        new = None
        m = _DT_SHEET_TOKEN.match(old.strip())
        if m:
            idx = int(m.group(1)) if m.group(1) else 0
            if idx < len(source_sheet_names) and source_sheet_names[idx]:
                new = str(source_sheet_names[idx])
            else:
                logger.warning(
                    f"[报表·sheet名] '{old}' 需数据源第 {idx + 1} 个 sheet，"
                    f"但仅有 {len(source_sheet_names)} 个，保留原名")
                continue
        elif "{" in old:
            new = old
            for key in ("year", "month", "date", "tenant"):
                if key in sys_vars and sys_vars[key] is not None:
                    new = new.replace("{" + key + "}", str(sys_vars[key]))

        if not new or new == old:
            continue

        taken.discard(old)              # 原名腾出，供其它 sheet 去重时复用
        final = _sanitize_sheet_name(new, taken)   # 内部已把 final 加入 taken
        try:
            ws.Name = final
            logger.info(f"[报表·sheet名] '{old}' -> '{final}'")
        except Exception as _re:
            taken.discard(final)        # 改名失败，撤销占位并恢复原名
            taken.add(old)
            logger.warning(f"[报表·sheet名] '{old}' -> '{final}' 改名失败（忽略）: {_re}")


def _generate_from_template_impl(
    output_path: str,
    template_path: str,
    data: Dict,
    password: Optional[str] = None,
    watermark_text: Optional[str] = None,
    mode: str = "fill",
    group_by: str = "",
    skip_rows: int = 1,
    name_field: str = "",
    show_empty_period: bool = True,
    split_by: str = "",
    sheet_source_names: Optional[List[str]] = None,
    sheet_vars: Optional[Dict] = None,
) -> str:
    """（子进程执行体）SmartMarker 填充模板生成文件。

    使用 Aspose WorkbookDesigner（SmartMarker 引擎）填充模板生成文件。

    四种模式:
      fill  — 整个 DataFrame 一次性填入模板（默认，适合汇总表）
      block — 按 group_by 字段分组，每组独立填充模板后合并到一个文件
      zip   — 按 group_by 字段分组，每组生成独立 xlsx，打包为 zip 下载
      sheet — 按 group_by 字段分组，每组生成独立 sheet

    split_by: 文件级拆分字段，按此列值将数据拆分到多个文件，打包为 zip。
              可与 fill/block/sheet 模式组合使用。

    模板标记:
      &=DT.Column              数据集字段
      &=DT.Column(skip:1)      带修饰符
      &=$year                  单值变量

    data 示例:
      {"DT": pd.DataFrame(...), "$year": "2026", "$month": "03"}

    Returns:
      fill/block/sheet → output_path (xlsx)
      zip / 有split_by → output_path (zip)
    """
    # sheet 名占位符仅在单文件模式生效；zip/split 打包模式本期不处理
    if (split_by or mode == "zip") and (sheet_source_names or sheet_vars):
        logger.info("[报表·sheet名] sheet 名占位符在 zip/split 打包模式下暂不处理")

    # 如果有 split_by，走文件级拆分包装器
    if split_by:
        return _generate_with_split(
            output_path, template_path, data,
            split_by=split_by, mode=mode,
            group_by=group_by, skip_rows=skip_rows,
            name_field=name_field,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )

    if mode == "block":
        return _generate_block(
            output_path, template_path, data,
            group_by=group_by, skip_rows=skip_rows,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
            sheet_source_names=sheet_source_names, sheet_vars=sheet_vars,
        )
    elif mode == "zip":
        return _generate_zip(
            output_path, template_path, data,
            group_by=group_by, name_field=name_field,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
        )
    elif mode == "sheet":
        return _generate_sheet(
            output_path, template_path, data,
            group_by=group_by,
            password=password, watermark_text=watermark_text,
            show_empty_period=show_empty_period,
            sheet_source_names=sheet_source_names, sheet_vars=sheet_vars,
        )
    else:
        return _generate_fill(
            output_path, template_path, data,
            password=password, watermark_text=watermark_text,
            sheet_source_names=sheet_source_names, sheet_vars=sheet_vars,
        )


def generate_from_template(
    output_path: str,
    template_path: str,
    data: Dict,
    password: Optional[str] = None,
    watermark_text: Optional[str] = None,
    mode: str = "fill",
    group_by: str = "",
    skip_rows: int = 1,
    name_field: str = "",
    show_empty_period: bool = True,
    split_by: str = "",
    sheet_source_names: Optional[List[str]] = None,
    sheet_vars: Optional[Dict] = None,
) -> str:
    """SmartMarker 填充模板生成文件（在【独立子进程】执行）。

    在子进程内打开模板 + WorkbookDesigner.Process + CalculateFormula + Save：
    特定文件（大模板/大 DataFrame/公式密集）在主进程跑会内存暴涨，此处超时/超内存
    强杀；失败 raise RuntimeError。所有调用方（多表合并套模板、报表生成等）自动受益。

    data 示例: {"DT": pd.DataFrame(...), "$year": "2026", "$month": "03"}
    Returns: fill/block/sheet → output_path (xlsx)；zip / 有split_by → output_path (zip)
    """
    from backend.utils.subprocess_runner import run_in_subprocess, default_max_memory_mb, default_timeout

    r = run_in_subprocess(
        "backend.utils.aspose_helper:_generate_from_template_impl",
        (str(output_path), str(template_path), data),
        kwargs={
            "password": password,
            "watermark_text": watermark_text,
            "mode": mode,
            "group_by": group_by,
            "skip_rows": skip_rows,
            "name_field": name_field,
            "show_empty_period": show_empty_period,
            "split_by": split_by,
            "sheet_source_names": sheet_source_names,
            "sheet_vars": sheet_vars,
        },
        timeout=default_timeout("write"), max_memory_mb=default_max_memory_mb(),
    )
    if r.success:
        return r.result
    reason = "超时" if r.timed_out else ("内存超限" if r.killed_by_memory else r.error)
    raise RuntimeError(f"模板填充失败（{reason}）: {template_path}")


def _smartmarker_fill(template_path: str, data: Dict) -> Workbook:
    """SmartMarker 核心填充: 打开模板 → 设置数据源 → Process → 返回填好的 Workbook。
    fill / block / zip 三种模式复用此函数。
    """
    from Aspose.Cells import WorkbookDesigner

    wb = _licensed_workbook(template_path)
    _fix_smart_marker_spacing(wb)

    designer = WorkbookDesigner()
    designer.Workbook = wb
    designer.RepeatFormulasWithSubtotal = True

    for name, value in data.items():
        if name.startswith("$"):
            var_name = name[1:]
            designer.SetDataSource(var_name, str(value))
            designer.SetDataSource(name, str(value))
        else:
            df = value if isinstance(value, pd.DataFrame) else pd.DataFrame(value)
            dt = _dataframe_to_datatable(df, name)
            designer.SetDataSource(dt)
            logger.info(f"[SmartMarker] 数据源 {name}: {len(df)} 行, 列={list(df.columns)}")

    designer.Process()
    wb.CalculateFormula()
    return wb


def _finalize_workbook(
    wb, output_path: str,
    password: Optional[str] = None,
    watermark_text: Optional[str] = None,
) -> str:
    """统一收尾: 水印 → 加密 → 保存"""
    if watermark_text:
        add_excel_watermark(wb, watermark_text)
    if password:
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
        wb.Settings.Password = password
    else:
        # 主动清除模板可能继承的密码属性
        try:
            wb.Settings.Password = ""
        except Exception:
            pass
    return save_as(wb, output_path)


def append_carryover_sheets(target_path: str, source_path: str, specs: List[str],
                            password: Optional[str] = None) -> int:
    """把计算结果(source_path)中指定的 sheet 整表(值+格式)拷贝追加到报表(target_path)末尾。

    specs: 每项为结果表的 sheet 名，或 '#N'（0-based 位置：#0=第一个 sheet）。
           可加 '@M' 后缀指定插入到报表的目标位置（0-based，如 '工资明细@0' 插到最前）。
    target 若已加密需传 password 以便打开并重新保存。返回成功拷贝的 sheet 数。
    """
    if not specs:
        return 0
    if not (source_path and os.path.exists(source_path) and os.path.exists(target_path)):
        logger.warning(f"[整表搬运] 源或目标文件缺失: src={source_path}, tgt={target_path}")
        return 0

    src_wb = _licensed_workbook(source_path)
    if password:
        _lo = LoadOptions()
        _lo.Password = password
        tgt_wb = _licensed_workbook(target_path, _lo)
    else:
        tgt_wb = _licensed_workbook(target_path)

    src_sheets = src_wb.Worksheets
    src_count = src_sheets.Count
    # 源 sheet 名 → 对象（精确 + 去空格忽略大小写兜底）
    _by_exact, _by_loose = {}, {}
    for i in range(src_count):
        ws = src_sheets[i]
        _by_exact[ws.Name] = ws
        _by_loose[str(ws.Name).strip().lower()] = ws

    def _resolve(spec):
        s = (spec or "").strip()
        if not s:
            return None
        if s.startswith("#"):
            try:
                idx = int(s[1:])          # 0 基：#0=第一个源 sheet，与插入位置 @N 一致
            except ValueError:
                return None
            return src_sheets[idx] if 0 <= idx < src_count else None
        return _by_exact.get(s) or _by_loose.get(s.lower())

    def _split_pos(spec):
        """拆出可选的插入位置后缀 '@N'（0 基目标位置）。
        返回 (源spec, 目标位置或None)。'工资明细@0'→('工资明细',0)、'#2'→('#2',None)。"""
        s = (spec or "").strip()
        if "@" in s:
            left, _, right = s.rpartition("@")
            right = right.strip()
            if right.lstrip("+-").isdigit():
                return left.strip(), int(right)
        return s, None

    existing = {tgt_wb.Worksheets[i].Name for i in range(tgt_wb.Worksheets.Count)}
    copied = 0
    for spec in specs:
        src_spec, insert_pos = _split_pos(spec)
        src_ws = _resolve(src_spec)
        if src_ws is None:
            logger.warning(f"[整表搬运] 结果表未找到 sheet: {spec}")
            continue
        # 唯一 sheet 名（与报表已有 sheet 不冲突）
        base = src_ws.Name
        name, k = base, 1
        while name in existing:
            name = f"{base}_{k}"
            k += 1
        # 新增空 sheet 再整表拷贝（Add 在不同 Aspose 版本可能返回 int 或 Worksheet）
        res = tgt_wb.Worksheets.Add()
        new_ws = res if hasattr(res, "Copy") else tgt_wb.Worksheets[res]
        new_ws.Copy(src_ws)          # 值 + 样式 + 列宽 + 合并单元格 全部复制
        new_ws.Name = name
        # 指定了插入位置 → 从末尾移动到目标索引（越界则钳到合法范围）
        if insert_pos is not None:
            try:
                _cnt = tgt_wb.Worksheets.Count
                _dst = max(0, min(insert_pos, _cnt - 1))
                new_ws.MoveTo(_dst)
                logger.info(f"[整表搬运] '{name}' 插入到位置 {_dst}（请求 @{insert_pos}）")
            except Exception as _me:
                logger.warning(f"[整表搬运] '{name}' 移动到位置 {insert_pos} 失败（保留在末尾）: {_me}")
        existing.add(name)
        copied += 1
        logger.info(f"[整表搬运] 已拷贝结果 sheet '{src_ws.Name}' → 报表 '{name}'")

    if copied:
        try:
            # 强制全量重算：新增 sheet 产生的跨表引用不在原计算链(calcChain)里，
            # 普通 CalculateFormula() 会依链跳过 → 报表引用该 sheet 的公式仍是旧缓存值
            # (搬运前该 sheet 不存在时算出的坏值)，Excel 打开需手工回车才刷新。
            tgt_wb.Settings.ForceFullCalculate = True
            tgt_wb.CalculateFormula()   # 拷入的公式先算并缓存，打开无需按回车
            # 双保险：写入 fullCalcOnLoad 标记，Excel 打开时也强制重算一遍
            tgt_wb.Settings.ReCalculateOnOpen = True
        except Exception as _ce:
            logger.warning(f"[整表搬运] CalculateFormula 跳过: {_ce}")
        if password:
            tgt_wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
            tgt_wb.Settings.Password = password
        tgt_wb.Save(target_path, _ext_save_format(target_path))
    return copied


def _fuzzy_match_column(target: str, columns) -> Optional[str]:
    """模糊匹配列名：去空格、忽略大小写"""
    target_clean = target.strip().lower()
    for col in columns:
        if str(col).strip().lower() == target_clean:
            return col
    return None


# ── fill 模式 ──────────────────────────────────────────

def _generate_fill(
    output_path: str, template_path: str, data: Dict,
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    sheet_source_names: Optional[List[str]] = None, sheet_vars: Optional[Dict] = None,
) -> str:
    """整个 DataFrame 一次性填入模板"""
    logger.info(f"[报表生成] fill 模式: {template_path}")
    wb = _smartmarker_fill(template_path, data)
    _rename_sheets_by_token(wb, sheet_source_names, sheet_vars)
    return _finalize_workbook(wb, output_path, password, watermark_text)


# ── block 模式 ─────────────────────────────────────────

def _generate_block(
    output_path: str, template_path: str, data: Dict,
    group_by: str = "", skip_rows: int = 1,
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
    sheet_source_names: Optional[List[str]] = None, sheet_vars: Optional[Dict] = None,
) -> str:
    """按 group_by 分组，每组用 SmartMarker 填充模板，合并到一个文件。"""
    # 找到主数据源（非 $ 开头的第一个 DataFrame）
    ds_name, full_df, vars_data, extra_dfs = _extract_datasource(data)

    # 模糊匹配 group_by 列名
    if group_by and group_by not in full_df.columns:
        matched = _fuzzy_match_column(group_by, full_df.columns)
        if matched:
            logger.info(f"[block] group_by 模糊匹配: '{group_by}' -> '{matched}'")
            group_by = matched

    if not group_by or group_by not in full_df.columns:
        logger.warning(f"[block] group_by='{group_by}' 不在列 {list(full_df.columns)} 中，回退到 fill 模式")
        return _generate_fill(output_path, template_path, data, password, watermark_text,
                              sheet_source_names=sheet_source_names, sheet_vars=sheet_vars)

    groups = full_df.groupby(group_by, sort=False)
    logger.info(f"[报表生成] block 模式: {len(groups)} 组, group_by={group_by}, skip_rows={skip_rows}")

    # 用第一组填充，确定一个块占多少行
    result_wb = None
    result_ws = None
    current_row = 0

    for group_idx, (group_key, group_df) in enumerate(groups):
        group_df = group_df.reset_index(drop=True)
        # 从分组数据首行自动提取 $变量（覆盖全局同名变量）
        group_vars = _extract_group_vars(group_df, vars_data)
        group_data = {ds_name: group_df, **extra_dfs, **group_vars}

        # SmartMarker 填充该组
        filled_wb = _smartmarker_fill(template_path, group_data)
        filled_ws = filled_wb.Worksheets[0]
        block_rows = filled_ws.Cells.MaxDataRow + 1

        if result_wb is None:
            # 第一组: 直接用填好的 workbook 作为结果
            result_wb = filled_wb
            result_ws = result_wb.Worksheets[0]
            current_row = block_rows
        else:
            # 后续组: 空行 + 复制块
            current_row += skip_rows

            # 从填好的 sheet 复制行到结果 sheet
            result_ws.Cells.CopyRows(
                filled_ws.Cells,       # 源
                0,                     # 源起始行
                current_row,           # 目标起始行
                block_rows,            # 行数
            )
            current_row += block_rows

        logger.info(f"[block] 组 {group_idx+1}/{len(groups)}: {group_key}, {len(group_df)} 行数据, 块={block_rows} 行")

    _rename_sheets_by_token(result_wb, sheet_source_names, sheet_vars)
    return _finalize_workbook(result_wb, output_path, password, watermark_text)


# ── sheet 模式 ────────────────────────────────────────

def _generate_sheet(
    output_path: str, template_path: str, data: Dict,
    group_by: str = "",
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
    sheet_source_names: Optional[List[str]] = None, sheet_vars: Optional[Dict] = None,
) -> str:
    """按 group_by 分组，每组生成一个独立 sheet，sheet 名取自分组值。"""
    ds_name, full_df, vars_data, extra_dfs = _extract_datasource(data)

    # 模糊匹配 group_by 列名
    if group_by and group_by not in full_df.columns:
        matched = _fuzzy_match_column(group_by, full_df.columns)
        if matched:
            logger.info(f"[sheet] group_by 模糊匹配: '{group_by}' -> '{matched}'")
            group_by = matched

    if not group_by or group_by not in full_df.columns:
        logger.warning(f"[sheet] group_by='{group_by}' 不在列 {list(full_df.columns)} 中，回退到 fill 模式")
        return _generate_fill(output_path, template_path, data, password, watermark_text,
                              sheet_source_names=sheet_source_names, sheet_vars=sheet_vars)

    groups = full_df.groupby(group_by, sort=False)
    logger.info(f"[报表生成] sheet 模式: {len(groups)} 组, group_by={group_by}")

    result_wb = None
    sheet_names = set()

    for group_idx, (group_key, group_df) in enumerate(groups):
        group_df = group_df.reset_index(drop=True)
        group_vars = _extract_group_vars(group_df, vars_data)
        group_data = {ds_name: group_df, **extra_dfs, **group_vars}

        # SmartMarker 填充该组
        filled_wb = _smartmarker_fill(template_path, group_data)

        sheet_name = _sanitize_sheet_name(str(group_key), sheet_names)

        if result_wb is None:
            # 第一组：直接用填好的 workbook，重命名第一个 sheet
            result_wb = filled_wb
            result_wb.Worksheets[0].Name = sheet_name
        else:
            # 后续组：从填好的 workbook 复制 sheet 到结果 workbook
            filled_ws = filled_wb.Worksheets[0]
            result_wb.Worksheets.AddCopy(result_wb.Worksheets.Count - 1)
            new_ws = result_wb.Worksheets[result_wb.Worksheets.Count - 1]
            new_ws.Copy(filled_ws)
            new_ws.Name = sheet_name

        logger.info(f"[sheet] 组 {group_idx+1}/{len(groups)}: sheet='{sheet_name}', {len(group_df)} 行数据")

    _rename_sheets_by_token(result_wb, sheet_source_names, sheet_vars)
    return _finalize_workbook(result_wb, output_path, password, watermark_text)


# ── zip 模式 ──────────────────────────────────────────

def _generate_zip(
    output_path: str, template_path: str, data: Dict,
    group_by: str = "", name_field: str = "",
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
) -> str:
    """按 group_by 分组，每组生成独立 xlsx，打包为 zip。"""
    ds_name, full_df, vars_data, extra_dfs = _extract_datasource(data)

    # 模糊匹配 group_by 列名（去空格、忽略大小写）
    if group_by and group_by not in full_df.columns:
        matched = _fuzzy_match_column(group_by, full_df.columns)
        if matched:
            logger.info(f"[zip] group_by 模糊匹配: '{group_by}' -> '{matched}'")
            group_by = matched

    if not group_by or group_by not in full_df.columns:
        logger.warning(f"[zip] group_by='{group_by}' 不在列 {list(full_df.columns)} 中，回退到 fill 模式")
        # 回退到 fill 模式时，输出路径改为 .xlsx（否则 Aspose 保存到 .zip 扩展名会产生无效文件）
        fill_path = os.path.splitext(output_path)[0] + ".xlsx"
        return _generate_fill(fill_path, template_path, data, password, watermark_text)

    # 确保输出路径是 .zip
    if not output_path.endswith(".zip"):
        output_path = os.path.splitext(output_path)[0] + ".zip"

    groups = full_df.groupby(group_by, sort=False)
    logger.info(f"[报表生成] zip 模式: {len(groups)} 组, group_by={group_by}")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for group_idx, (group_key, group_df) in enumerate(groups):
            group_df = group_df.reset_index(drop=True)
            # 从分组数据首行自动提取 $变量（覆盖全局同名变量）
            group_vars = _extract_group_vars(group_df, vars_data)
            group_data = {ds_name: group_df, **extra_dfs, **group_vars}

            filled_wb = _smartmarker_fill(template_path, group_data)

            if watermark_text:
                add_excel_watermark(filled_wb, watermark_text)
            if password:
                filled_wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
                filled_wb.Settings.Password = password
            else:
                try:
                    filled_wb.Settings.Password = ""
                except Exception:
                    pass

            # 确定文件名
            if name_field and name_field in group_df.columns:
                file_label = str(group_df[name_field].iloc[0])
            else:
                file_label = str(group_key)
            # 清理非法文件名字符
            file_label = re.sub(r'[\\/:*?"<>|]', '_', file_label)
            inner_name = f"{file_label}.xlsx"

            # 保存到内存 → 写入 zip
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp_path = tmp.name
            try:
                filled_wb.Save(tmp_path, SaveFormat.Xlsx)
                zf.write(tmp_path, inner_name)
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

            logger.info(f"[zip] 组 {group_idx+1}/{len(groups)}: {inner_name}, {len(group_df)} 行")

    logger.info(f"[报表生成] zip 完成: {output_path}")
    return output_path


# ── split_by 文件级拆分 ────────────────────────────────

def _generate_with_split(
    output_path: str, template_path: str, data: Dict,
    split_by: str = "", mode: str = "fill",
    group_by: str = "", skip_rows: int = 1,
    name_field: str = "",
    password: Optional[str] = None, watermark_text: Optional[str] = None,
    show_empty_period: bool = True,
) -> str:
    """按 split_by 字段拆分数据到多个文件，每个文件内按 mode 模式生成，打包为 zip。"""
    ds_name, full_df, vars_data, extra_dfs = _extract_datasource(data)

    # 模糊匹配 split_by 列名
    if split_by not in full_df.columns:
        matched = _fuzzy_match_column(split_by, full_df.columns)
        if matched:
            logger.info(f"[split] split_by 模糊匹配: '{split_by}' -> '{matched}'")
            split_by = matched

    if split_by not in full_df.columns:
        logger.warning(f"[split] split_by='{split_by}' 不在列 {list(full_df.columns)} 中，忽略拆分，走普通模式")
        if mode == "sheet":
            return _generate_sheet(output_path, template_path, data, group_by=group_by,
                                   password=password, watermark_text=watermark_text,
                                   show_empty_period=show_empty_period)
        elif mode == "block":
            return _generate_block(output_path, template_path, data, group_by=group_by,
                                   skip_rows=skip_rows, password=password,
                                   watermark_text=watermark_text, show_empty_period=show_empty_period)
        elif mode == "zip":
            return _generate_zip(output_path, template_path, data, group_by=group_by,
                                 name_field=name_field, password=password,
                                 watermark_text=watermark_text, show_empty_period=show_empty_period)
        else:
            return _generate_fill(output_path, template_path, data, password=password,
                                  watermark_text=watermark_text)

    # split_by + zip 时，split_by 覆盖 zip 的 group_by 语义
    if mode == "zip":
        logger.info(f"[split] split_by + zip 模式，split_by 覆盖 group_by，等同于按 '{split_by}' 做 zip")
        return _generate_zip(output_path, template_path, data,
                             group_by=split_by, name_field=name_field,
                             password=password, watermark_text=watermark_text,
                             show_empty_period=show_empty_period)

    # 确保输出路径是 .zip
    if not output_path.endswith(".zip"):
        output_path = os.path.splitext(output_path)[0] + ".zip"

    split_groups = full_df.groupby(split_by, sort=False)
    logger.info(f"[报表生成] split 模式: {len(split_groups)} 个文件, split_by={split_by}, 内部模式={mode}")

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for split_idx, (split_key, split_df) in enumerate(split_groups):
            split_df = split_df.reset_index(drop=True)
            split_vars = _extract_group_vars(split_df, vars_data)
            split_data = {ds_name: split_df, **extra_dfs, **split_vars}

            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp_path = tmp.name

            try:
                if mode == "sheet":
                    _generate_sheet(tmp_path, template_path, split_data,
                                    group_by=group_by, password=password,
                                    watermark_text=watermark_text,
                                    show_empty_period=show_empty_period)
                elif mode == "block":
                    _generate_block(tmp_path, template_path, split_data,
                                    group_by=group_by, skip_rows=skip_rows,
                                    password=password, watermark_text=watermark_text,
                                    show_empty_period=show_empty_period)
                else:
                    _generate_fill(tmp_path, template_path, split_data,
                                   password=password, watermark_text=watermark_text)

                file_label = re.sub(r'[\\/:*?"<>|]', '_', str(split_key).strip())
                if not file_label:
                    file_label = f"group_{split_idx + 1}"
                inner_name = f"{file_label}.xlsx"

                zf.write(tmp_path, inner_name)
            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

            logger.info(f"[split] 文件 {split_idx+1}/{len(split_groups)}: {inner_name}, {len(split_df)} 行, 内部模式={mode}")

    logger.info(f"[报表生成] split 完成: {output_path}")
    return output_path


# ── 公共工具 ──────────────────────────────────────────

def _extract_datasource(data: Dict):
    """从 data dict 中分离主数据源 DataFrame、额外数据源和 $ 变量。
    Returns: (ds_name, DataFrame, vars_dict, extra_dfs)
    extra_dfs: 主数据源之外的其他 DataFrame（用于多 sheet 数据源）
    """
    ds_name = None
    full_df = None
    extra_dfs = {}
    vars_data = {}

    for name, value in data.items():
        if name.startswith("$"):
            vars_data[name] = value
        elif ds_name is None:
            ds_name = name
            full_df = value if isinstance(value, pd.DataFrame) else pd.DataFrame(value)
        else:
            extra_dfs[name] = value if isinstance(value, pd.DataFrame) else pd.DataFrame(value)

    if full_df is None:
        raise ValueError("data 中未找到 DataFrame 数据源")

    return ds_name, full_df, vars_data, extra_dfs


def _extract_group_vars(group_df: pd.DataFrame, global_vars: Dict) -> Dict:
    """从分组 DataFrame 首行提取 $变量，用于 block/zip 模式。

    逻辑：
    1. 以全局 vars_data 为基础（复制一份）
    2. 遍历 group_df 的列名，如果全局存在同名 $变量（$col_name），
       则用该组首行的实际值覆盖
    3. 这样模版中的 &=$month 等变量会自动取到当前分组对应的值

    例: 全局 $month="03"，但当前分组首行 salary_month=1 → $salary_month="1"
        全局 $tenant="XXX"，分组中无 tenant 列 → 保持 $tenant="XXX"
    """
    result = dict(global_vars)  # 复制全局变量作为基础

    if group_df.empty:
        return result

    first_row = group_df.iloc[0]
    for col in group_df.columns:
        var_key = f"${col}"
        # 将分组首行的列值作为 $变量，覆盖全局同名变量
        val = first_row[col]
        if val is not None and str(val).strip() != "":
            result[var_key] = str(val)

    return result


def _sanitize_sheet_name(name: str, existing_names: set) -> str:
    """清洗 sheet 名称：截断至31字符，替换非法字符，处理重名。"""
    # 替换非法字符
    clean = re.sub(r'[\\/:*?\[\]]', '_', str(name).strip())
    # 截断至 31 字符（Excel 限制）
    if len(clean) > 31:
        clean = clean[:31]
    # 空名称兜底
    if not clean:
        clean = "Sheet"
    # 处理重名
    base = clean
    counter = 2
    while clean in existing_names:
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        clean = base[:max_base_len] + suffix
        counter += 1
    existing_names.add(clean)
    return clean


def _fill_template_data(wb, name: str, df: pd.DataFrame) -> None:
    """将 DataFrame 填入模板中 &=name.col 占位符。

    扫描所有 sheet，找到 &=name.col 标记后，将 DataFrame 对应列数据从该行向下填充。
    如果数据行数 > 1，会自动向下插入行并填充。
    """
    marker_prefix = f"&={name}."

    for ws_idx in range(wb.Worksheets.Count):
        ws = wb.Worksheets[ws_idx]
        cells = ws.Cells
        max_row = cells.MaxDataRow
        max_col = cells.MaxDataColumn
        if max_row < 0 or max_col < 0:
            continue

        # 第一遍：收集所有标记的位置
        markers = []  # [(row, col, col_name)]
        for r in range(max_row + 1):
            for c in range(max_col + 1):
                val = str(cells[r, c].StringValue or "")
                if val.startswith(marker_prefix):
                    col_name = val[len(marker_prefix):]
                    if col_name in df.columns:
                        markers.append((r, c, col_name))

        if not markers:
            continue

        # 确定数据起始行（所有标记应在同一行）
        data_row = markers[0][0]
        n_rows = len(df)

        # 如果数据多于 1 行，先插入空行（保留格式）
        if n_rows > 1:
            cells.InsertRows(data_row + 1, n_rows - 1)

        # 第二遍：填充数据
        for marker_row, marker_col, col_name in markers:
            data_list = df[col_name].tolist()
            for i, v in enumerate(data_list):
                target_row = data_row + i
                if v is not None and not (isinstance(v, float) and pd.isna(v)):
                    cells[target_row, marker_col].PutValue(v)
                else:
                    cells[target_row, marker_col].PutValue("")


def _fill_single_value(wb, var_name: str, value) -> None:
    """将单值填入模板中 &=$Variable 占位符"""
    marker = f"&={var_name}"

    for ws_idx in range(wb.Worksheets.Count):
        ws = wb.Worksheets[ws_idx]
        cells = ws.Cells
        max_row = cells.MaxDataRow
        max_col = cells.MaxDataColumn
        if max_row < 0 or max_col < 0:
            continue
        for r in range(max_row + 1):
            for c in range(max_col + 1):
                val = str(cells[r, c].StringValue or "")
                if val == marker:
                    cells[r, c].PutValue(value)


def add_excel_watermark(wb, watermark_text: str) -> None:
    """在工作簿第一个 Sheet 添加 WordArt 文字水印"""
    from Aspose.Cells.Drawing import MsoPresetTextEffect

    ws = wb.Worksheets[0]
    ws.Shapes.AddTextEffect(
        MsoPresetTextEffect.TextEffect1,
        watermark_text,
        "Arial Black",
        60, False, True,
        1, 8, 1, 1, 130, 500,
    )


# ═══════════════════════════════════════════════════════
# 加密 / 解密
# ═══════════════════════════════════════════════════════

def is_encrypted(file_path: str) -> bool:
    """检测 Excel 文件是否有打开密码"""
    try:
        info = FileFormatUtil.DetectFileFormat(file_path)
        return bool(info.IsEncrypted)
    except Exception:
        return False


def encrypt_excel(
    input_path: str,
    output_path: str = None,
    password: str = "123456",
) -> str:
    """Excel 文件加密（设置打开密码）
    使用 Aspose.Cells 原生加密（StrongCryptographicProvider, 128位）
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    wb = _licensed_workbook(input_path)
    wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
    wb.Settings.Password = password
    wb.Save(output_path)

    logger.info(f"Excel加密完成: {input_path} -> {output_path}")
    return output_path


def decrypt_excel(
    input_path: str,
    output_path: str = None,
    password: str = "",
) -> str:
    """Excel 文件解密（移除打开密码）
    使用 Aspose.Cells 原生解密
    """
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    logger.info(f"[decrypt_excel] path={input_path}, pwd_len={len(password)}")

    try:
        load_opts = LoadOptions()
        load_opts.Password = password
        wb = _licensed_workbook(input_path, load_opts)
    except Exception as e:
        err_str = str(e)
        if 'Invalid password' in err_str:
            logger.warning(f"[decrypt_excel] 密码无效, 尝试无密码打开: {input_path}")
            try:
                wb = _licensed_workbook(input_path)
            except Exception as e2:
                logger.error(f"[decrypt_excel] 无密码打开也失败: {e2}")
                raise ValueError(
                    f"文件 '{input_path}' 加密且密码不正确"
                ) from e
        else:
            raise
    wb.Settings.Password = None
    wb.Save(output_path)

    logger.info(f"Excel解密完成: {input_path} -> {output_path}")
    return output_path


def write_protect_excel(
    input_path: str,
    output_path: str = None,
    password: str = "123456",
) -> str:
    """Excel 写保护（可打开查看，编辑需密码）"""
    if not output_path:
        output_path = tempfile.mktemp(suffix=".xlsx")

    wb = _licensed_workbook(input_path)
    wb.Settings.WriteProtection.Password = password
    wb.Settings.WriteProtection.RecommendReadOnly = True
    wb.Save(output_path)

    logger.info(f"Excel写保护完成: {input_path} -> {output_path}")
    return output_path


# ═══════════════════════════════════════════════════════
# 错误处理
# ═══════════════════════════════════════════════════════

def change_error_message(source_error_message: str) -> str:
    """将 Aspose 抛出的英文异常转为可读的中文提示。未匹配时原样返回。"""
    for pattern, template in _ERROR_PATTERNS:
        m = re.search(pattern, source_error_message, re.IGNORECASE)
        if m:
            cell_name = m.group(1)
            type_name = _TYPE_CHANGES.get(m.group(2).lower(), m.group(2))
            return template.format(cell_name, type_name)
    return source_error_message
