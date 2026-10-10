"""智算结果公式重算回归测试。"""

import sys
import types
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from backend.utils.excel_comparator import (
    _aspose_calc_impl, _aspose_mark_recalculation_impl, inspect_formula_cache,
)


class _Settings:
    ForceFullCalculate = False
    ReCalculateOnOpen = False


class _FakeWorkbook:
    instances = []

    def __init__(self, path):
        self.path = path
        self.Settings = _Settings()
        self.calculated = False
        self.saved = False
        self.disposed = False
        self.__class__.instances.append(self)

    def CalculateFormula(self):
        self.calculated = True

    def Save(self, path):
        self.saved = path == self.path

    def Dispose(self):
        self.disposed = True


def _fake_aspose_modules():
    aspose_init = types.ModuleType("aspose_init")
    aspose_init.ensure_license = lambda: None
    aspose_package = types.ModuleType("Aspose")
    cells = types.ModuleType("Aspose.Cells")
    cells.Workbook = _FakeWorkbook
    aspose_package.Cells = cells
    return {
        "aspose_init": aspose_init,
        "Aspose": aspose_package,
        "Aspose.Cells": cells,
    }


class FormulaRecalculationTests(unittest.TestCase):
    def setUp(self):
        _FakeWorkbook.instances.clear()

    def test_calculation_refreshes_cache_and_sets_recalc_flags(self):
        with patch.dict(sys.modules, _fake_aspose_modules()):
            self.assertTrue(_aspose_calc_impl("result.xlsx"))

        wb = _FakeWorkbook.instances[-1]
        self.assertTrue(wb.Settings.ForceFullCalculate)
        self.assertTrue(wb.Settings.ReCalculateOnOpen)
        self.assertTrue(wb.calculated)
        self.assertTrue(wb.saved)
        self.assertTrue(wb.disposed)

    def test_marker_fallback_does_not_run_formula_engine(self):
        with patch.dict(sys.modules, _fake_aspose_modules()):
            self.assertTrue(_aspose_mark_recalculation_impl("result.xlsx"))

        wb = _FakeWorkbook.instances[-1]
        self.assertTrue(wb.Settings.ForceFullCalculate)
        self.assertTrue(wb.Settings.ReCalculateOnOpen)
        self.assertFalse(wb.calculated)
        self.assertTrue(wb.saved)

    def test_formula_cache_inspection_distinguishes_empty_error_and_bad_ref(self):
        workbook_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets><sheet name="结果" sheetId="1" r:id="rId1"/></sheets>
        </workbook>"""
        rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
        </Relationships>"""
        sheet_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1">
            <c r="A1"><f>SUM(B1:C1)</f><v>3</v></c>
            <c r="A2"><f>ROUND(B2,2)</f><v/></c>
            <c r="A3" t="e"><f>#REF!</f><v>#REF!</v></c>
            <c r="A4"><f>_xlfn.XLOOKUP(B4,'[2]源'!B:B,'[2]源'!C:C)</f><v>9</v></c>
          </row></sheetData>
        </worksheet>"""
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "formula.xlsx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("xl/workbook.xml", workbook_xml)
                archive.writestr("xl/_rels/workbook.xml.rels", rels_xml)
                archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
            report = inspect_formula_cache(path)

        self.assertEqual(report["formula_count"], 4)
        self.assertEqual(report["empty_cache_count"], 1)
        self.assertEqual(report["error_cache_count"], 1)
        self.assertEqual(report["invalid_ref_formula_count"], 1)
        self.assertEqual(report["external_formula_count"], 1)


if __name__ == "__main__":
    unittest.main()
