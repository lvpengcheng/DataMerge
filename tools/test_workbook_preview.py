"""Upload-preview regressions: bounded XML reads and independence from Excel jobs."""
import asyncio
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock, patch

from backend.utils import workbook_preview as preview

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def fixture(path, worksheets, shared=None):
    names = list(worksheets)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('xl/workbook.xml', f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets>' +
            ''.join(f'<sheet name="{name}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, name in enumerate(names)) +
            '</sheets></workbook>')
        archive.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            ''.join(f'<Relationship Id="rId{i+1}" Target="/xl/worksheets/sheet{i+1}.xml"/>' for i in range(len(names))) + '</Relationships>')
        for index, sheet in enumerate(worksheets.values(), 1):
            archive.writestr(f'xl/worksheets/sheet{index}.xml', sheet)
        if shared is not None:
            archive.writestr('xl/sharedStrings.xml', shared)
        # Preview must never touch workbook styles, even if malformed or huge.
        archive.writestr('xl/styles.xml', b'broken styles' * 100000)


def sheet(rows):
    return f'<worksheet xmlns="{MAIN}"><dimension ref="A1:XFD1048576"/><sheetData>{rows}</sheetData></worksheet>'


class WorkbookPreviewTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / '目标.xlsx'

    def test_sheet_names_inline_strings_and_zero_ignore_giant_dimensions(self):
        fixture(self.path, {'结果': sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>工号</t></is></c>'
            '<c r="B1"><v>0</v></c><c r="XFD1"><v>999</v></c></row>'), '辅助': sheet('')})
        result = preview.peek_xlsx_sheets(self.path)
        self.assertEqual([row['name'] for row in result], ['结果', '辅助'])
        self.assertEqual(result[0]['preview'], [['工号', '0']])
        self.assertEqual(result[1]['preview'], [])

    def test_shared_strings_rich_text_and_cached_formula_values(self):
        fixture(self.path, {'结果': sheet('<row r="1"><c r="A1" t="s"><v>1</v></c>'
            '<c r="B1"><f>1+1</f><v>2</v></c></row>')},
            f'<sst xmlns="{MAIN}"><si><t>unused</t></si><si><r><t>出勤</t></r><r><t>天数</t></r></si></sst>')
        self.assertEqual(preview.peek_xlsx_sheets(self.path)[0]['preview'], [['出勤天数', '2']])

    def test_large_shared_strings_and_large_row_never_block_sheet_listing(self):
        fixture(self.path, {'结果': sheet('<row r="1"><c r="A1" t="s"><v>999999</v></c></row>'),
                            '大标题': sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>' + 'x' * 2000000 + '</t></is></c></row>')},
            f'<sst xmlns="{MAIN}">' + '<si><t>long header text</t></si>' * 100000 + '</sst>')
        reads = {}
        real_zip = zipfile.ZipFile
        class TrackedZip(real_zip):
            def open(self, name, *args, **kwargs):
                stream = super().open(name, *args, **kwargs)
                class Reader:
                    def read(self, size=-1):
                        data = stream.read(size)
                        reads[name] = reads.get(name, 0) + len(data)
                        return data
                    def __enter__(self):
                        return self
                    def __exit__(self, *args):
                        stream.close()
                return Reader()
        with patch.object(preview.zipfile, 'ZipFile', TrackedZip):
            result = preview.peek_xlsx_sheets(self.path)
        self.assertEqual([row['name'] for row in result], ['结果', '大标题'])
        self.assertLessEqual(reads['xl/sharedStrings.xml'], preview._SHARED_STRING_LIMIT)
        self.assertLessEqual(reads['xl/worksheets/sheet2.xml'], preview._SHEET_PREVIEW_LIMIT)
        self.assertNotIn('xl/styles.xml', reads)

    def test_broken_or_late_preview_does_not_hide_valid_names(self):
        fixture(self.path, {'晚数据': sheet('<row r="500000"><c r="A500000"><v>1</v></c></row>'),
                            '预览损坏': '<not valid XML'})
        result = preview.peek_xlsx_sheets(self.path)
        self.assertEqual([row['name'] for row in result], ['晚数据', '预览损坏'])
        self.assertTrue(all(row['preview'] == [] for row in result))


class PreviewRoutingTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        preview._preview_gate = None
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / '目标.xlsx'

    async def test_xlsx_does_not_enter_heavy_excel_gate_or_start_worker(self):
        fixture(self.path, {'结果': sheet('')})
        with patch('backend.utils.upload_stream.get_excel_work_semaphore', side_effect=AssertionError('heavy Excel queue used')):
            result = await preview.preview_uploaded_workbook(self.path)
        self.assertEqual(result[0]['name'], '结果')

    async def test_legacy_fallback_has_light_entry_and_bounded_queue(self):
        self.path.write_bytes(b'not a modern zip')
        gate = SimpleNamespace(acquire=AsyncMock(return_value=True), release=Mock())
        worker = AsyncMock(return_value=SimpleNamespace(success=True, result=[{'name': '旧版表', 'preview': []}]))
        with patch('backend.utils.upload_stream.get_excel_work_semaphore', return_value=gate), \
             patch.dict(sys.modules, {'backend.utils.subprocess_runner': SimpleNamespace(run_in_fresh_subprocess_async=worker)}):
            result = await preview.preview_uploaded_workbook(str(self.path), 'secret')
        self.assertEqual(result[0]['name'], '旧版表')
        self.assertEqual(worker.call_args.args[0], 'backend.utils.workbook_preview:peek_legacy_sheets')
        self.assertEqual(worker.call_args.kwargs['timeout'], 60)
        self.assertEqual(worker.call_args.kwargs['args'], (str(self.path), 'secret'))
        gate.release.assert_called_once()

    async def test_busy_legacy_queue_fails_clearly_without_releasing_unacquired_slot(self):
        self.path.write_bytes(b'not a modern zip')
        gate = SimpleNamespace(acquire=AsyncMock(side_effect=asyncio.TimeoutError), release=Mock())
        worker = AsyncMock()
        with patch('backend.utils.upload_stream.get_excel_work_semaphore', return_value=gate), \
             patch.dict(sys.modules, {'backend.utils.subprocess_runner': SimpleNamespace(run_in_fresh_subprocess_async=worker)}):
            with self.assertRaisesRegex(TimeoutError, '稍后重试'):
                await preview.preview_uploaded_workbook(str(self.path))
        worker.assert_not_called()
        gate.release.assert_not_called()


if __name__ == '__main__':
    unittest.main()
