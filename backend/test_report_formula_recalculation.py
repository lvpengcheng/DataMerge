"""报表最终保存前公式重算回归测试。"""

import importlib
import sys
import types
import unittest
from types import SimpleNamespace
from unittest.mock import patch


class _FakeSettings:
    ForceFullCalculate = False
    ReCalculateOnOpen = False
    Password = ""


class _FakeWorkbook:
    instances = []

    def __init__(self, path=None, _options=None):
        self.path = path
        self.Settings = _FakeSettings()
        self.calculated = 0
        self.saved = []
        self.disposed = False
        self.__class__.instances.append(self)

    def CalculateFormula(self):
        self.calculated += 1

    def Save(self, path, save_format=None):
        self.saved.append((path, save_format))

    def Dispose(self):
        self.disposed = True

    def SetEncryptionOptions(self, *_args):
        pass


def _fake_modules():
    aspose_init = types.ModuleType("aspose_init")
    aspose_init.ensure_license = lambda: None

    cells = types.ModuleType("Aspose.Cells")
    cells.Workbook = _FakeWorkbook
    cells.SaveFormat = SimpleNamespace(
        Xlsx="xlsx",
        Excel97To2003="xls",
        Pdf="pdf",
        Csv="csv",
    )
    for name in (
        "PdfSaveOptions", "LoadOptions", "EncryptionType", "FileFormatUtil",
        "HtmlSaveOptions", "BackgroundType", "CellArea", "FormatConditionType",
        "OperatorType",
    ):
        setattr(cells, name, type(name, (), {}))
    cells.EncryptionType.StrongCryptographicProvider = "strong"

    aspose = types.ModuleType("Aspose")
    aspose.Cells = cells
    rendering = types.ModuleType("Aspose.Cells.Rendering")
    pdf_security = types.ModuleType("Aspose.Cells.Rendering.PdfSecurity")
    pdf_security.PdfSecurityOptions = type("PdfSecurityOptions", (), {})
    return {
        "aspose_init": aspose_init,
        "Aspose": aspose,
        "Aspose.Cells": cells,
        "Aspose.Cells.Rendering": rendering,
        "Aspose.Cells.Rendering.PdfSecurity": pdf_security,
    }


class ReportFormulaRecalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.modules_patch = patch.dict(sys.modules, _fake_modules())
        cls.modules_patch.start()
        sys.modules.pop("backend.utils.aspose_helper", None)
        cls.helper = importlib.import_module("backend.utils.aspose_helper")

    @classmethod
    def tearDownClass(cls):
        sys.modules.pop("backend.utils.aspose_helper", None)
        cls.modules_patch.stop()

    def setUp(self):
        _FakeWorkbook.instances.clear()

    def test_finalizer_calculates_after_all_workbook_changes(self):
        wb = _FakeWorkbook()
        result = self.helper._finalize_workbook(wb, "report.xlsx")

        self.assertEqual(result, "report.xlsx")
        self.assertEqual(wb.calculated, 1)
        self.assertTrue(wb.Settings.ForceFullCalculate)
        self.assertTrue(wb.Settings.ReCalculateOnOpen)
        self.assertTrue(wb.saved)
        self.assertTrue(wb.disposed)

    def test_finalizer_can_defer_calculation_until_report_postprocessing_finishes(self):
        wb = _FakeWorkbook()
        result = self.helper._finalize_workbook(
            wb, "report.xlsx", calculate_formulas=False,
        )

        self.assertEqual(result, "report.xlsx")
        self.assertEqual(wb.calculated, 0)
        self.assertTrue(wb.saved)
        self.assertTrue(wb.disposed)

    def test_postprocess_recalculation_supports_password_protected_report(self):
        self.assertTrue(
            self.helper._recalculate_report_file_impl("report.xlsx", "secret")
        )

        wb = _FakeWorkbook.instances[-1]
        self.assertEqual(wb.calculated, 1)
        self.assertTrue(wb.Settings.ForceFullCalculate)
        self.assertTrue(wb.Settings.ReCalculateOnOpen)
        self.assertTrue(wb.saved)
        self.assertTrue(wb.disposed)


if __name__ == "__main__":
    unittest.main()
