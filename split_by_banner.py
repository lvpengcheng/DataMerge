"""按区域 banner 拆分 sheet 脚本

对 source 文件夹下每个 Excel 文件：
- 用 IntelligentExcelParser 解析每个 sheet 的所有 region
- 对于多区域 sheet，取每个 region head_row_start - 1 行的第一个非空值作为 banner
- 按 banner 分组，把同一 sheet 拆成多个子 sheet：{原sheet名}-{banner}
- 子 sheet 只含表头 + 数据行（不含 banner 行）
- 单区域 sheet / 无 banner 的 sheet 原样复制

输出：对每个源文件生成 {stem}_split.xlsx
"""

import re
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from excel_parser import IntelligentExcelParser


_INVALID_SHEET_CHARS = re.compile(r"[\\/*?\[\]:]")
_MAX_SHEET_NAME = 31


def _sanitize_sheet_name(name: str, used: set) -> str:
    """清洗 sheet 名：去除非法字符、截断到 31 字符、去重"""
    cleaned = _INVALID_SHEET_CHARS.sub("_", str(name)).strip().strip("'") or "unnamed"
    if len(cleaned) > _MAX_SHEET_NAME:
        cleaned = cleaned[:_MAX_SHEET_NAME]
    base = cleaned
    suffix = 1
    while cleaned in used:
        suffix_str = f"_{suffix}"
        cleaned = base[: _MAX_SHEET_NAME - len(suffix_str)] + suffix_str
        suffix += 1
    used.add(cleaned)
    return cleaned


def _get_banner_value(ws, head_row_start: int):
    """取 head_row_start - 1 行的第一个非空值，作为该区域的 banner"""
    banner_row = head_row_start - 1
    if banner_row < 1:
        return None
    for col in range(1, ws.max_column + 1):
        v = ws.cell(banner_row, col).value
        if v is not None and str(v).strip():
            return str(v).strip()
    return None


def _copy_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)


def _copy_region_rows(src_ws, dst_ws, src_rows: list, dst_start_row: int) -> int:
    """从 src_ws 复制 src_rows 列表中的整行到 dst_ws，从 dst_start_row 开始顺序写入。
    返回写入的行数。"""
    max_col = src_ws.max_column
    for offset, src_row in enumerate(src_rows):
        for col in range(1, max_col + 1):
            _copy_cell(src_ws.cell(src_row, col), dst_ws.cell(dst_start_row + offset, col))
    return len(src_rows)


def _copy_full_sheet(src_ws, dst_ws):
    """整 sheet 复制（值 + 格式）"""
    max_col = src_ws.max_column
    for row in range(1, src_ws.max_row + 1):
        for col in range(1, max_col + 1):
            _copy_cell(src_ws.cell(row, col), dst_ws.cell(row, col))


def _region_row_indices(region) -> list:
    """region 的 head + data 行索引（数据无效时仅返回 head 行）"""
    rows = list(range(region.head_row_start, region.head_row_end + 1))
    if region.data_row_start > 0 and region.data_row_end >= region.data_row_start:
        rows.extend(range(region.data_row_start, region.data_row_end + 1))
    return rows


def split_one_file(source_path: Path, output_path: Path):
    parser = IntelligentExcelParser()
    results = parser.parse_excel_file(str(source_path), max_data_rows=1, read_formulas=False)

    src_wb = openpyxl.load_workbook(str(source_path), data_only=False)
    dst_wb = openpyxl.Workbook()
    dst_wb.remove(dst_wb.active)
    used_names: set = set()

    parsed_sheets = {sd.sheet_name: sd for sd in results}

    for sheet_name in src_wb.sheetnames:
        src_ws = src_wb[sheet_name]
        sheet_data = parsed_sheets.get(sheet_name)
        regions = sheet_data.regions if sheet_data else []

        # 多区域且至少一个区域能取到 banner 时才执行拆分
        banner_groups: dict = {}
        no_banner_regions = []
        if len(regions) >= 2:
            for region in sorted(regions, key=lambda r: r.head_row_start):
                banner = _get_banner_value(src_ws, region.head_row_start)
                if banner:
                    banner_groups.setdefault(banner, []).append(region)
                else:
                    no_banner_regions.append(region)

        if not banner_groups:
            # 单区域或无 banner → 原样复制整 sheet
            dst_name = _sanitize_sheet_name(sheet_name, used_names)
            dst_ws = dst_wb.create_sheet(dst_name)
            _copy_full_sheet(src_ws, dst_ws)
            continue

        for banner, region_list in banner_groups.items():
            sub_name = _sanitize_sheet_name(f"{sheet_name}-{banner}", used_names)
            dst_ws = dst_wb.create_sheet(sub_name)
            cur_row = 1
            for region in region_list:
                rows = _region_row_indices(region)
                written = _copy_region_rows(src_ws, dst_ws, rows, cur_row)
                cur_row += written + 1  # 不同 region 之间空一行分隔

        if no_banner_regions:
            sub_name = _sanitize_sheet_name(f"{sheet_name}-other", used_names)
            dst_ws = dst_wb.create_sheet(sub_name)
            cur_row = 1
            for region in no_banner_regions:
                rows = _region_row_indices(region)
                written = _copy_region_rows(src_ws, dst_ws, rows, cur_row)
                cur_row += written + 1

    if not dst_wb.sheetnames:
        dst_wb.create_sheet("empty")

    dst_wb.save(str(output_path))


def main():
    script_dir = Path(__file__).parent
    source_folder = script_dir / "source"
    out_folder = script_dir / "split_output"
    out_folder.mkdir(exist_ok=True)

    excel_extensions = {".xlsx", ".xls", ".xlsm"}
    files = [f for f in source_folder.iterdir()
             if f.suffix.lower() in excel_extensions and not f.name.startswith("~$")]

    if not files:
        print(f"未找到 Excel 文件: {source_folder}")
        return

    print(f"待处理 {len(files)} 个文件")
    for i, src in enumerate(files, 1):
        out = out_folder / f"{src.stem}_split.xlsx"
        print(f"  [{i}/{len(files)}] {src.name}")
        try:
            split_one_file(src, out)
            print(f"      → {out.name}")
        except Exception as e:
            print(f"      错误: {e}")

    print("完成")


if __name__ == "__main__":
    main()
