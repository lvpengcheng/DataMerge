import unittest
import importlib.util
import sys
import tempfile
import types
from pathlib import Path
from unittest.mock import patch


_spec = importlib.util.spec_from_file_location(
    "output_postprocess_under_test", Path(__file__).resolve().parent / "utils" / "output_postprocess.py")
_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_module)
_shifted_summary_tail = _module._shifted_summary_tail
_template_format_row = _module._template_format_row


class _Cell:
    def __init__(self, value=None, row=0, col=0, fmt=""):
        self.Value = value
        self.Row = row
        self.Column = col
        self.IsFormula = False
        self.Formula = ""
        self.style = types.SimpleNamespace(Number=0, Custom=fmt)

    def GetStyle(self):
        return types.SimpleNamespace(Number=self.style.Number, Custom=self.style.Custom)

    def SetStyle(self, style):
        self.style = style


class _Cells:
    def __init__(self, labels):
        self._labels = labels
        self.MaxDataRow = max(labels)
        self.MaxDataColumn = 9
        self._actual = {(row, 0): _Cell(value, row, 0) for row, value in labels.items()}

    def __getitem__(self, address):
        row, col = address
        return self._actual.get((row, col), _Cell(None, row, col))

    def GetEnumerator(self):
        iterator = iter(self._actual.values())

        class Enumerator:
            Current = None

            def MoveNext(self):
                self.Current = next(iterator, None)
                return self.Current is not None

        return Enumerator()


class _Sheets(list):
    @property
    def Count(self):
        return len(self)


class _Book:
    def __init__(self, cells):
        self.Worksheets = _Sheets([types.SimpleNamespace(Name="账单明细", Cells=cells)])

    def Save(self, path):
        pass

    def Dispose(self):
        pass


class ShiftedSummaryTailTests(unittest.TestCase):
    def setUp(self):
        self.template = _Cells({25: "合计", 26: "服务费总计（小写）金额",
                                27: "服务费总计（大写）金额", 38: "合计"})

    def test_shorter_result_keeps_merged_amount_with_footer(self):
        output = _Cells({15: "合计", 16: "服务费总计（小写）金额",
                         17: "服务费总计（大写）金额", 28: "合计"})
        tail = _shifted_summary_tail(output, self.template)
        self.assertEqual(tail, (15, 25))
        self.assertEqual(_template_format_row(16, tail, self.template.MaxDataRow), 26)

    def test_longer_result_keeps_merged_amount_with_footer(self):
        output = _Cells({35: "合计", 36: "服务费总计（小写）金额",
                         37: "服务费总计（大写）金额", 48: "合计"})
        tail = _shifted_summary_tail(output, self.template)
        self.assertEqual(tail, (35, 25))
        self.assertEqual(_template_format_row(36, tail, self.template.MaxDataRow), 26)
        self.assertIsNone(_template_format_row(26, tail, self.template.MaxDataRow))

    def test_same_rows_need_no_remapping(self):
        self.assertIsNone(_shifted_summary_tail(self.template, self.template))

    def test_format_restoration_uses_moved_merged_anchor(self):
        output = _Cells({15: "合计", 16: "服务费总计（小写）金额",
                         17: "服务费总计（大写）金额", 28: "合计"})
        template_amount = _Cell(35664.75, 26, 9, r"\¥#,##0.00")
        output_amount = _Cell(35664.75, 16, 9)
        self.template._actual[(26, 9)] = template_amount
        output._actual[(16, 9)] = output_amount
        template_book, output_book = _Book(self.template), _Book(output)
        fake_aspose = types.ModuleType("Aspose")
        fake_cells = types.ModuleType("Aspose.Cells")
        fake_cells.Workbook = _Book
        fake_aspose.Cells = fake_cells
        fake_init = types.ModuleType("aspose_init")
        fake_init.ensure_license = lambda: None
        with tempfile.TemporaryDirectory() as directory:
            template_path = Path(directory) / "template.xlsx"
            output_path = Path(directory) / "output.xlsx"
            template_path.touch()
            output_path.touch()
            def open_book(path):
                return template_book if str(path) == str(template_path) else output_book
            with patch.dict(sys.modules, {"Aspose": fake_aspose,
                                          "Aspose.Cells": fake_cells,
                                          "aspose_init": fake_init}), \
                 patch.object(_module, "_open_workbook", side_effect=open_book):
                _module.restore_formats_from_template(output_path, template_path)
        self.assertEqual(output_amount.style.Custom, template_amount.style.Custom)
        self.assertEqual(output_amount.Value, 35664.75)


if __name__ == "__main__":
    unittest.main()
