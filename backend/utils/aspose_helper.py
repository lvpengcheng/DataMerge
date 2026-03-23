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
  generate_from_template(output_path, template_path, data, password, watermark_text)
      模板填充生成文件（xlsx / pdf / html）
  add_excel_watermark(wb, text)

  ── 加密 / 解密 ──────────────────────────────────────
  is_encrypted / encrypt_excel / decrypt_excel / write_protect_excel

  ── 工具 ──────────────────────────────────────────────
  change_error_message(msg)      Aspose 英文异常转中文
"""

import logging
import os
import re
import tempfile
from typing import Optional, Union, Dict

import pandas as pd

import aspose_init  # noqa: F401 — 确保 Aspose 已初始化

from Aspose.Cells import (  # type: ignore
    Workbook, SaveFormat, PdfSaveOptions, LoadOptions, EncryptionType,
    FileFormatUtil, HtmlSaveOptions, BackgroundType,
    CellArea, FormatConditionType, OperatorType,
)
from Aspose.Cells.Rendering.PdfSecurity import PdfSecurityOptions  # type: ignore

logger = logging.getLogger(__name__)


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
            wb = Workbook(path, opts)
        else:
            wb = Workbook(path)

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
        wb = Workbook(path, opts)
    else:
        wb = Workbook(path)

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
    wb = Workbook()
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
    wb = wb_or_path if isinstance(wb_or_path, Workbook) else Workbook(str(wb_or_path))
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
        wb = Workbook(ws_or_path)
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

    wb = Workbook(input_path)

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

    wb = Workbook(input_path)

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

    wb = Workbook(input_path)
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

def generate_from_template(
    output_path: str,
    template_path: str,
    data: Dict,
    password: Optional[str] = None,
    watermark_text: Optional[str] = None,
) -> str:
    """
    以 Excel 模板 + 数据集生成目标文件。

    模板标记方式：
      - 数据集标记: &=DataSource.Column  （对应 data 中的 DataFrame 或 list[dict]）
      - 单值标记:   &=$Variable          （对应 data 中以 $ 开头的键值）

    data 示例:
      {
          "员工": pd.DataFrame({"姓名": [...], "薪资": [...]}),
          "$公司名称": "XX科技有限公司",
          "$日期": "2026-03-23",
      }

    模板中写 &=员工.姓名、&=员工.薪资 → 自动展开为多行
    模板中写 &=$公司名称       → 替换为单个值

    输出格式由 output_path 的扩展名决定（xlsx/pdf/html/csv）

    Args:
        output_path:    输出文件路径
        template_path:  模板 xlsx 文件路径
        data:           数据字典
        password:       可选，加密密码
        watermark_text: 可选，水印文字

    Returns:
        output_path
    """
    wb = Workbook(template_path)

    for name, value in data.items():
        if name.startswith("$"):
            _fill_single_value(wb, name, value)
        else:
            df = value if isinstance(value, pd.DataFrame) else pd.DataFrame(value)
            _fill_template_data(wb, name, df)

    wb.CalculateFormula()

    if watermark_text:
        add_excel_watermark(wb, watermark_text)

    if password:
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128)
        wb.Settings.Password = password

    return save_as(wb, output_path)


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

    wb = Workbook(input_path)
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
        wb = Workbook(input_path, load_opts)
    except Exception as e:
        err_str = str(e)
        if 'Invalid password' in err_str:
            logger.warning(f"[decrypt_excel] 密码无效, 尝试无密码打开: {input_path}")
            try:
                wb = Workbook(input_path)
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

    wb = Workbook(input_path)
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
