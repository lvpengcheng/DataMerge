"""SOP structures and three data samples per region, using the shared parser."""
from pathlib import Path
import re

from backend.utils.sop_privacy import SopPrivacy


def _visible_sheet_names(file_path):
    """Read only the inventory, including sheets with no recognized data region."""
    if Path(file_path).suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False)
        try:
            return [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
        finally:
            wb.close()
    import aspose_init
    aspose_init.ensure_license()
    from Aspose.Cells import Workbook
    wb = Workbook(str(file_path))
    try:
        return [str(wb.Worksheets[i].Name) for i in range(wb.Worksheets.Count) if wb.Worksheets[i].IsVisible]
    finally:
        wb.Dispose()


def read_sop_excel(file_path: str, max_chars: int, privacy=None) -> str:
    from excel_parser import IntelligentExcelParser

    privacy = privacy or SopPrivacy()
    visible = _visible_sheet_names(file_path)
    if not visible:
        raise ValueError("没有可读取的非隐藏 Sheet")
    sheets = IntelligentExcelParser().parse_excel_file(
        file_path, max_data_rows=3, read_formulas=True, calculate_formulas=True,
        active_sheet_only=False, best_region_only=False, skip_hidden_sheets=True,
        raise_errors=True,
    )
    by_name = {sheet.sheet_name: sheet for sheet in sheets}
    parts = ["读取范围：全部非隐藏 Sheet 的结构、表头、字段类型、样本公式；每个数据区域最多 3 行脱敏数据示例，表头不占示例行数。样本不代表全量核对。",
             "非隐藏 Sheet 清单：" + "、".join(visible)]
    for name in visible:
        parts.append(f"=== Sheet: {name} ===")
        sheet = by_name.get(name)
        if not sheet or not sheet.regions:
            parts.append("（未识别到数据区域，可能为空表或非标准结构；不能据此认定所需字段不存在。）")
            continue
        for index, region in enumerate(sheet.regions, 1):
            headers = region.head_data or {}
            fields = {column: privacy.field_kind(header) for header, column in headers.items()}
            parts.append(f"数据区域 {index}，表头行 {region.head_row_start}-{region.head_row_end}，数据区 {region.data_row_start}-{region.data_row_end}")
            parts.append("表头：" + " | ".join(f"{header}({column})" for header, column in headers.items()))
            for column, schema in (region.column_schemas or {}).items():
                parts.append(f"字段 {column}: 类型={schema.get('field_type', '')}，格式={schema.get('number_format', '')}")
            for sample_index, row in enumerate((region.data or [])[:3], 1):
                cells = []
                for column, value in row.items():
                    kind = fields.get(column)
                    rendered = privacy.alias(value, kind) if kind else value
                    if not kind:
                        privacy.learn_text(str(value))
                    cells.append(f"{column}={rendered}")
                parts.append(f"示例数据 {sample_index}: " + " | ".join(cells))
            for address, formula in (region.formula or {}).items():
                column = re.sub(r"[\d$]", "", address)
                kind = fields.get(column)
                if kind:
                    for literal in re.findall(r'"((?:[^"]|"")*)"', formula):
                        if literal:
                            privacy.alias(literal.replace('""', '"'), kind)
                parts.append(f"样本公式 {address}: {formula}")
    text = privacy.redact("\n".join(parts))
    if len(text) > max_chars:
        raise ValueError("全部 Sheet 的结构和样本合计超过审核容量，请拆分文件；本次没有截掉后续 Sheet 继续审核")
    return text
