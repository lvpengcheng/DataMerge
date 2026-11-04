"""Identifier types, long-number precision and native output policy regressions."""
import importlib.util
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from backend.utils import column_contract as contract


class IdentifierPolicyTests(unittest.TestCase):
    def test_match_keys_do_not_strip_leading_zeroes_or_round_decimals(self):
        self.assertEqual(contract.normalize_match_key(123), '123')
        self.assertEqual(contract.normalize_match_key(123.0), '123')
        self.assertEqual(contract.normalize_match_key(' １２３ '), '123')
        self.assertEqual(contract.normalize_match_key('00123'), '00123')
        value = 1.2345678901234567
        self.assertEqual(contract.normalize_match_key(value), repr(value))
        import numpy as np
        self.assertEqual(contract.normalize_match_key(np.int64(123)), '123')
        self.assertEqual(contract.normalize_match_key(np.float64(123)), '123')
        for value in (None, '', '  ', True, float('nan'), float('inf'), datetime(2026, 1, 1)):
            with self.subTest(value=value):
                self.assertIsNone(contract.normalize_match_key(value))

    def test_identifier_aliases_are_specific(self):
        for name in ('PSID', 'Employee Code', 'employee_id', '员工工号', '身份证号码'):
            self.assertTrue(contract._is_identifier_column(name), name)
        for name in ('Paid', '工资金额', 'ID说明', 'identity amount'):
            self.assertFalse(contract._is_identifier_column(name), name)

    def test_explicit_key_remains_backwards_compatible_and_raw_preserves_type(self):
        self.assertEqual(contract.canon_value(123, 'key'), '123')
        self.assertEqual(contract.canon_value('00123', 'key'), '00123')
        self.assertIsInstance(contract.canon_value(123, 'raw'), int)

    def test_long_integer_text_never_rounds_through_numeric_conversion(self):
        for text in ('110101199001011234', '00110101199001011234',
                     '00000000000000123', '1.10101199001011234E17', '1E1000000'):
            self.assertEqual(contract.canon_value(text, 'number'), text)
        self.assertEqual(contract.canon_value('123456789012345', 'number'), 123456789012345)
        self.assertEqual(contract.canon_value('1,234.50', 'number'), 1234.5)
        self.assertEqual(contract.canon_value('(100)', 'number'), -100)
        self.assertEqual(contract.canon_value('12%', 'number'), .12)

    def test_source_conversion_preserves_long_ids_and_does_not_mutate_cache(self):
        import pandas as pd
        frame = pd.DataFrame({'身份证': ['110101199001011234'], '工号': [123]})
        data = {'employees': {'df': frame}}
        specs = contract.normalize_column_specs({'*': {'身份证': 'number', '工号': 'text'}})
        result = contract.apply_specs_to_source_data(data, specs)
        self.assertEqual(result['employees']['df'].iloc[0]['身份证'], '110101199001011234')
        self.assertEqual(result['employees']['df'].iloc[0]['工号'], '123')
        self.assertEqual(frame.iloc[0]['工号'], 123)


@unittest.skipUnless(importlib.util.find_spec('pythonnet'), 'requires native Excel runtime')
class NativeIdentifierPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import aspose_init
        if not aspose_init.init_aspose():
            raise RuntimeError('native Excel runtime could not initialize')
        from Aspose.Cells import Workbook, CellValueType
        cls.Workbook = Workbook
        cls.Types = CellValueType

    def make_book(self, header='PSID', values=(123, '123', '00123', '110101199001011234')):
        book = self.Workbook()
        cells = book.Worksheets[0].Cells
        cells[0, 0].PutValue(header)
        for row, value in enumerate(values, 1):
            cells[row, 0].PutValue(value)
        self.addCleanup(book.Dispose)
        return book

    def test_default_text_policy_wins_over_template_and_stale_schema(self):
        for fmt in ('@', '0', 'General'):
            with self.subTest(fmt=fmt), tempfile.TemporaryDirectory() as directory:
                template = self.make_book()
                for row in range(1, 5):
                    cell = template.Worksheets[0].Cells[row, 0]
                    style = cell.GetStyle(); style.Custom = fmt; cell.SetStyle(style)
                path = str(Path(directory) / 'template.xlsx'); template.Save(path)
                book = self.make_book()
                structure = {'sheets': {'Sheet1': {'column_schemas': {'PSID': {'field_type': 'text'}}}}}
                result = contract.apply_column_contract(book, template_path=path, expected_structure=structure)
                cells = book.Worksheets[0].Cells
                self.assertEqual(cells[1, 0].Type, self.Types.IsString)
                self.assertEqual([cells[r, 0].StringValue for r in range(2, 5)],
                                 ['123', '00123', '110101199001011234'])
                self.assertEqual(cells[3, 0].GetStyle().Custom, '@')
                self.assertEqual(cells[4, 0].GetStyle().Custom, '@')
                self.assertEqual(result['values'], 1)
                # Check storage after the actual xlsx serialization boundary as well.
                saved = str(Path(directory) / 'result.xlsx'); book.Save(saved)
                from openpyxl import load_workbook
                exported = load_workbook(saved)
                self.assertEqual(exported.active['A2'].value, '123')
                self.assertEqual(exported.active['A4'].value, '00123')
                self.assertEqual(exported.active['A5'].value, '110101199001011234')
                exported.close()
                template.Dispose()
                book.Dispose()

    def test_explicit_type_override_still_wins(self):
        book = self.make_book(values=(123,))
        code = "_COLUMN_SPECS = {'*': {'PSID': {'type': 'text', 'format': '@'}}}"
        contract.apply_column_contract(book, script_code=code)
        self.assertEqual(book.Worksheets[0].Cells[1, 0].Type, self.Types.IsString)

    def test_numeric_identifier_precision_risks_are_reported(self):
        book = self.make_book(values=(1.1010119900101123e17,))
        with self.assertLogs(contract.logger, level='WARNING') as logs:
            result = contract.apply_column_contract(book)
        self.assertEqual(result['precision_risks'], 1)
        self.assertIn('可能已丢失精度', logs.output[0])
        self.assertEqual(book.Worksheets[0].Cells[1, 0].Type, self.Types.IsString)

    def test_currency_fix_and_date_time_behaviour_survive(self):
        with tempfile.TemporaryDirectory() as directory:
            template = self.make_book('工资金额', (123.5,))
            cell = template.Worksheets[0].Cells[1, 0]
            style = cell.GetStyle(); style.Custom = '#,##0.00;[Red](#,##0.00);"-";@'; cell.SetStyle(style)
            path = str(Path(directory) / 'template.xlsx'); template.Save(path)
            book = self.make_book('工资金额', ('123.50',))
            contract.apply_column_contract(book, template_path=path)
            self.assertEqual(book.Worksheets[0].Cells[1, 0].Type, self.Types.IsNumeric)
            self.assertEqual(book.Worksheets[0].Cells[1, 0].DoubleValue, 123.5)
        code = "_COLUMN_SPECS = {'*': {'入职日期': {'type': 'datetime'}}}"
        book = self.make_book('入职日期', ('2026-11-03 12:00:00',))
        contract.apply_column_contract(book, script_code=code)
        self.assertAlmostEqual(book.Worksheets[0].Cells[1, 0].DoubleValue % 1, .5)

    def test_formula_is_never_replaced_with_a_value(self):
        book = self.make_book(values=(123,))
        book.Worksheets[0].Cells[2, 0].Formula = '=A2'
        contract.apply_column_contract(book)
        self.assertTrue(book.Worksheets[0].Cells[2, 0].IsFormula)

    def test_v2_template_also_converts_numeric_identifiers(self):
        with tempfile.TemporaryDirectory() as directory:
            template = self.make_book(values=('123',))
            cell = template.Worksheets[0].Cells[1, 0]
            style = cell.GetStyle(); style.Custom = '@'; cell.SetStyle(style)
            path = str(Path(directory) / 'template.xlsx'); template.Save(path)
            book = self.make_book(values=(123,))
            contract.apply_column_contract(book, template_path=path, legacy_layers=False)
            self.assertEqual(book.Worksheets[0].Cells[1, 0].Type, self.Types.IsString)

    def test_v2_without_template_protects_long_text_ids(self):
        book = self.make_book(values=('00110101199001011234',))
        contract.apply_column_contract(book, legacy_layers=False)
        cell = book.Worksheets[0].Cells[1, 0]
        self.assertEqual(cell.StringValue, '00110101199001011234')
        self.assertEqual(cell.GetStyle().Custom, '@')


if __name__ == '__main__':
    unittest.main()
