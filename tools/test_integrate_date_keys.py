"""Native integration path: localized dates must match the parsed date keys."""
import ast
import json
import logging
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import pandas as pd


class IntegrateDateKeyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import aspose_init
        from Aspose.Cells import Workbook, SaveFormat
        from System import DateTime
        cls.Workbook, cls.SaveFormat, cls.DateTime = Workbook, SaveFormat, DateTime
        names = {'_parse_file_full_impl', '_parse_file_to_df_impl', '_regions_to_df',
                 '_enrich_matching_context', '_validate_integrate_columns', '_integrate_execute_impl'}
        tree = ast.parse(Path('backend/api/tools.py').read_text(encoding='utf-8'))
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
        scope = dict(pd=pd, Path=Path, json=json, logger=logging.getLogger(__name__), re=__import__('re'),
                     __name__='backend.api.tools', __package__='backend.api')
        exec(compile(ast.Module(body=nodes, type_ignores=[]), '<integration routes>', 'exec'), scope)
        cls.run_integration = staticmethod(scope['_integrate_execute_impl'])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for main in (True, False):
            wb = self.Workbook()
            cells = wb.Worksheets[0].Cells
            cells['A1'].PutValue('GR加班工资')
            headers = ['ID', '日期', '姓名', '加班工资'] if main else ['日期', '姓名', '加班工资']
            for ci, header in enumerate(headers):
                cells[1, ci].PutValue(header)
            date_ci = 1 if main else 0
            for index, (day, amount) in enumerate([(6, 400), (7, 500), (15, 400), (16, 600)], 2):
                cells[index, date_ci].PutValue(self.DateTime(2026, 8, day))
                style = cells[index, date_ci].GetStyle()
                style.Custom = '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy' if main else 'yyyy-mm-dd'
                cells[index, date_ci].SetStyle(style)
                cells[index, date_ci + 1].PutValue('员工' + str(index))
                if main:
                    cells[index, 0].PutValue(str(index))
                else:
                    cells[index, 2].PutValue(amount)
            wb.Save(str(self.root / ('原表.xlsx' if main else '数据表.xlsx')), self.SaveFormat.Xlsx)
            wb.Dispose()
        (self.root / '_meta.json').write_text(json.dumps({'files': [
            {'name': name, 'sheet': 'Sheet1', 'header_manual': True} for name in ['原表.xlsx', '数据表.xlsx']],
            'header_ranges': {name: [2, 2] for name in ['原表.xlsx', '数据表.xlsx']}}), encoding='utf-8')

    def request(self, mode, overwrite=False):
        pair = {'a_col': '加班工资', 'source_file': '数据表.xlsx', 'source_col': '加班工资'}
        return dict(main_file='原表.xlsx', key_map={'原表.xlsx': '日期', '数据表.xlsx': '日期'},
                    normalize_keys=True, date_key_mode=mode, output_mode=2,
                    compare_pairs=[pair], overwrite_pairs=[pair] if overwrite else [],
                    name_col='姓名', id_col='ID')

    def test_date_keys_compare_and_fill_at_all_supported_granularities(self):
        for mode in ['off', 'yearmonthday', 'yearmonth', 'month', 'day']:
            with self.subTest(mode=mode):
                # Empty wages in the main table must yield four differences.
                stat = self.run_integration(str(self.root), self.request(mode))
                self.assertEqual(stat['diff_rows'], 4)
                # Filling follows the existing aggregation semantics for coarse keys.
                stat = self.run_integration(str(self.root), self.request(mode, overwrite=True))
                self.assertEqual(stat['matched_rows'], 4)
                self.assertEqual(stat['overwritten_cells'], 4)
                self.assertEqual(stat['diff_rows'], 0)
                wb = self.Workbook(str(self.root / '整合结果_原表.xlsx'))
                try:
                    cells = wb.Worksheets['Sheet1'].Cells
                    expected = [1900] * 4 if mode in ('yearmonth', 'month') else [400, 500, 400, 600]
                    self.assertEqual([cells[r, 3].DoubleValue for r in range(2, 6)], expected)
                    self.assertEqual(cells['B3'].DateTimeValue.Day, 6)
                    self.assertIn('F800', cells['B3'].GetStyle().Custom)
                finally:
                    wb.Dispose()


if __name__ == '__main__':
    unittest.main()
