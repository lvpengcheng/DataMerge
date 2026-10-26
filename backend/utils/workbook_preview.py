"""Bounded worksheet-name/header preview, separate from full Excel processing.

Modern workbooks need only their ZIP/XML metadata. Do not initialize Aspose,
read styles, resolve links, calculate formulas or load a full shared-string table.
"""
import posixpath
import zipfile
import asyncio
import os
from xml.etree import ElementTree as ET

_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_METADATA_LIMIT = 4 * 1024 * 1024
_SHEET_PREVIEW_LIMIT = 128 * 1024
_TOTAL_PREVIEW_LIMIT = 2 * 1024 * 1024
_SHARED_STRING_LIMIT = 512 * 1024
_ROWS = 3
_COLS = 30
_preview_gate = None


async def preview_uploaded_workbook(path, password=None):
    """Modern metadata never waits behind formula calculation/training jobs."""
    global _preview_gate
    if await asyncio.to_thread(zipfile.is_zipfile, path):
        if _preview_gate is None:
            _preview_gate = asyncio.Semaphore(2)
        try:
            await asyncio.wait_for(_preview_gate.acquire(), timeout=8)
        except asyncio.TimeoutError as exc:
            raise TimeoutError('目标文件预览请求较多，请稍后重试') from exc
        try:
            return await asyncio.wait_for(asyncio.to_thread(peek_xlsx_sheets, path), timeout=15)
        except asyncio.TimeoutError as exc:
            raise TimeoutError('目标文件目录读取超时，请检查文件后重试') from exc
        finally:
            _preview_gate.release()
    from backend.utils.upload_stream import get_excel_work_semaphore
    from backend.utils.subprocess_runner import run_in_fresh_subprocess_async
    gate = get_excel_work_semaphore()
    try:
        await asyncio.wait_for(gate.acquire(), timeout=8)
    except asyncio.TimeoutError as exc:
        raise TimeoutError('旧版或加密文件预览正在等待 Excel 处理，请稍后重试') from exc
    try:
        result = await run_in_fresh_subprocess_async(
            'backend.utils.workbook_preview:peek_legacy_sheets', args=(path, password),
            timeout=60, max_memory_mb=int(os.getenv('TRAINING_UPLOAD_MAX_MEMORY_MB', '1536')))
        if not result.success:
            raise ValueError(result.error or '目标文件目录读取失败')
        return result.result or []
    finally:
        gate.release()


def _metadata(archive, name):
    if archive.getinfo(name).file_size > _METADATA_LIMIT:
        raise ValueError('工作簿目录过大，无法快速预览，请拆分目标文件')
    content = archive.read(name)
    if b'<!DOCTYPE' in content or b'<!ENTITY' in content:
        raise ValueError('工作簿包含不支持的 XML 声明')
    return ET.fromstring(content)


def _prefix_events(archive, name, budget):
    """Stop on a prefix, even when a producer declares a million-row dimension."""
    parser = ET.XMLPullParser(events=('end',))
    with archive.open(name) as stream:
        while budget > 0:
            data = stream.read(min(8192, budget))
            if not data:
                break
            budget -= len(data)
            if b'<!DOCTYPE' in data or b'<!ENTITY' in data:
                break
            parser.feed(data)
            yield from parser.read_events()
    # A bounded prefix may end midway through an element. Do not close/parse
    # the missing tail merely to display a best-effort header preview.


def _column_index(address):
    column = 0
    for char in address:
        if not char.isalpha():
            break
        column = column * 26 + ord(char.upper()) - ord('A') + 1
        if column > _COLS:
            return _COLS + 1
    return column


def _sheet_preview(archive, name, budget):
    rows, references = [], []
    for _, element in _prefix_events(archive, name, budget):
        if element.tag != '{' + _MAIN + '}row':
            continue
        row_index = int(element.get('r', str(len(rows) + 1)))
        if row_index < 1:
            element.clear()
            continue
        if row_index > _ROWS:
            break
        while len(rows) < row_index:
            rows.append([])
        row = rows[row_index - 1]
        next_column = 1
        for cell in element:
            if cell.tag != '{' + _MAIN + '}c':
                continue
            column = _column_index(cell.get('r', '')) or next_column
            next_column = column + 1
            if column > _COLS:
                continue
            while len(row) < column:
                row.append('')
            value = cell.find('{' + _MAIN + '}v')
            text = value.text if value is not None and value.text else ''
            if cell.get('t') == 'inlineStr':
                text = ''.join(part.text or '' for part in cell.iter('{' + _MAIN + '}t'))
            elif cell.get('t') == 's' and text:
                references.append((row, column - 1, int(text)))
                text = ''
            row[column - 1] = text.strip()[:200]
        element.clear()
        if len(rows) >= _ROWS:
            break
    return rows, references


def peek_xlsx_sheets(path):
    with zipfile.ZipFile(path) as archive:
        workbook = _metadata(archive, 'xl/workbook.xml')
        relationships = _metadata(archive, 'xl/_rels/workbook.xml.rels')
        locations = {}
        for relation in relationships:
            if relation.get('TargetMode') == 'External':
                continue
            target = relation.get('Target', '')
            locations[relation.get('Id')] = (target.lstrip('/') if target.startswith('/')
                                            else posixpath.normpath(posixpath.join('xl', target)))
        result, references = [], []
        budget = _TOTAL_PREVIEW_LIMIT
        for sheet in workbook.findall('{' + _MAIN + '}sheets/{' + _MAIN + '}sheet'):
            row = {'name': sheet.get('name', ''), 'preview': []}
            result.append(row)
            location = locations.get(sheet.get('{' + _REL + '}id'))
            if not location or not location.startswith('xl/worksheets/') or budget <= 0:
                continue
            limit = min(_SHEET_PREVIEW_LIMIT, budget)
            budget -= limit
            try:
                row['preview'], refs = _sheet_preview(archive, location, limit)
                references.extend(refs)
            except (KeyError, ET.ParseError, ValueError):
                # A malformed/oversized preview must not hide the valid sheet list.
                pass
        if references and 'xl/sharedStrings.xml' in archive.namelist():
            wanted = {index for _, _, index in references}
            values, index = {}, 0
            try:
                for _, element in _prefix_events(archive, 'xl/sharedStrings.xml', _SHARED_STRING_LIMIT):
                    if element.tag != '{' + _MAIN + '}si':
                        continue
                    if index in wanted:
                        values[index] = ''.join(part.text or '' for part in element.iter('{' + _MAIN + '}t')).strip()[:200]
                    index += 1
                    element.clear()
                    if len(values) == len(wanted):
                        break
            except (ET.ParseError, ValueError):
                pass
            for row, column, index in references:
                row[column] = values.get(index, '')
        return result


def peek_legacy_sheets(path, password=None):
    """Isolated native fallback for encrypted workbooks and binary .xls only."""
    import aspose_init
    aspose_init.ensure_license()
    from Aspose.Cells import Workbook, LoadOptions
    from backend.utils.output_postprocess import _silent_load_options
    options = _silent_load_options() or LoadOptions()
    if password:
        options.Password = password
    workbook = Workbook(str(path), options)
    try:
        result = []
        for index in range(workbook.Worksheets.Count):
            sheet = workbook.Worksheets[index]
            cells = sheet.Cells
            preview = [[str(cells[r, c].StringValue or '').strip()[:200]
                        for c in range(min(_COLS, cells.MaxDataColumn + 1))]
                       for r in range(min(_ROWS, cells.MaxDataRow + 1))]
            result.append({'name': str(sheet.Name), 'preview': preview})
        return result
    finally:
        workbook.Dispose()
