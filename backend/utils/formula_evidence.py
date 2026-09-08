"""读取公式原文和缓存状态；公式样本独立于数据样本的行数限制。"""
from pathlib import Path
import re


def _pattern(value, row):
    from openpyxl.formula.tokenizer import Tokenizer
    from openpyxl.utils.cell import range_boundaries
    try:
        parts = []
        for token in Tokenizer(value).items:
            part = token.value
            if token.subtype == 'RANGE':
                prefix, sep, ref = part.rpartition('!')
                ref = ref if sep else part
                try:
                    range_boundaries(ref)
                    ref = re.sub(r'(\$?[A-Za-z]+)(\$?)(\d+)',
                                 lambda m: m[0] if m[2] else f'{m[1]}[r{int(m[3])-row:+d}]', ref)
                    part = prefix + sep + ref if sep else ref
                except ValueError:
                    pass
            parts.append(part)
        return ''.join(parts)
    except Exception:
        return value


def _collect_legacy_formulas(path):
    import aspose_init
    aspose_init.ensure_license()
    from Aspose.Cells import Workbook
    wb = Workbook(str(path))
    evidence = {}
    try:
        for si in range(wb.Worksheets.Count):
            ws = wb.Worksheets[si]
            seen, formulas, count = set(), {}, 0
            iterator = ws.Cells.GetEnumerator()
            while iterator.MoveNext():
                cell = iterator.Current
                if not cell.IsFormula:
                    continue
                count += 1
                value = str(cell.Formula)
                key = (cell.Column, _pattern(value, cell.Row + 1))
                if key not in seen:
                    seen.add(key)
                    formulas[str(cell.Name)] = value
            evidence[str(ws.Name)] = {'formula_count': count, 'formulas': formulas}
    finally:
        wb.Dispose()
    return evidence


def collect_formula_evidence(file_path):
    """流式扫描所有单元格，按列去重填充公式，保留不同分支及绝对引用。"""
    import openpyxl
    from openpyxl.formula.tokenizer import Tokenizer
    from openpyxl.utils.cell import range_boundaries

    if Path(file_path).suffix.lower() == '.xls':
        return _collect_legacy_formulas(file_path)
    if Path(file_path).suffix.lower() not in ('.xlsx', '.xlsm'):
        return None
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=False, keep_links=False)
    evidence = {}
    try:
        for ws in wb.worksheets:
            ws.reset_dimensions()  # Ignore inflated declared dimensions; scan actual XML rows.
            seen, formulas, count = set(), {}, 0
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type != 'f':
                        continue
                    count += 1
                    value = cell.value
                    if not isinstance(value, str):
                        value = getattr(value, 'text', None) or str(value)
                    # 仅在 RANGE token 内归一相对行号；数字常量/文本中的 A1 不得改写。
                    try:
                        tokens = Tokenizer(value).items
                        parts = []
                        for token in tokens:
                            part = token.value
                            if token.subtype == 'RANGE':
                                prefix, sep, ref = part.rpartition('!')
                                ref = ref if sep else part
                                try:
                                    range_boundaries(ref)
                                    ref = re.sub(r'(\$?[A-Za-z]+)(\$?)(\d+)',
                                                 lambda m: m[0] if m[2] else f'{m[1]}[r{int(m[3])-cell.row:+d}]', ref)
                                    part = prefix + sep + ref if sep else ref
                                except ValueError:
                                    pass
                            parts.append(part)
                        pattern = ''.join(parts)
                    except Exception:
                        pattern = value
                    key = (cell.column, pattern)
                    if key not in seen:
                        seen.add(key)
                        formulas[cell.coordinate] = value
            evidence[ws.title] = {'formula_count': count, 'formulas': formulas}
    finally:
        wb.close()
    return evidence


def format_formula_evidence(evidence):
    if evidence is None:
        return '公式范围说明：此格式只能提供解析样本中的公式，未完成全表公式扫描。\n'
    lines = ['全表公式扫描（同列相同填充逻辑去重，保留原始坐标；不依赖缓存值）：']
    for sheet, info in evidence.items():
        lines.append(f"Sheet: {sheet}，公式单元格 {info['formula_count']} 个，不同公式样式 {len(info['formulas'])} 个")
        lines.extend(f'  {address}: {formula}' for address, formula in info['formulas'].items())
    return '\n'.join(lines) + '\n'
