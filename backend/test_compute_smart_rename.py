import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from backend.utils.source_auto_filler import (
    _add_structural_fallback_recommendations,
    _find_matching_asset,
    _period_index,
    auto_fill_missing_sources,
    auto_rename_uploaded_by_combined_score,
)


class ComputeSmartRenameTests(unittest.TestCase):
    def test_parses_common_month_file_names(self):
        expected = 2026 * 12 + 7 - 1
        self.assertEqual(_period_index("202607.xlsx"), expected)
        self.assertEqual(_period_index("工资_2026-07.xls"), expected)
        self.assertEqual(_period_index("工资2026年7月.xlsm"), expected)

    def test_maps_numeric_months_to_previous_and_current_roles(self):
        structure = {
            "files": {
                "上月.xlsx": {"sheets": {}},
                "本月.xlsx": {"sheets": {}},
            }
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "202607.xlsx").touch()
            (root / "202608.xlsx").touch()

            renamed, ambiguous, _ = auto_rename_uploaded_by_combined_score(
                temp_dir, structure, salary_year=2026, salary_month=8,
            )

            self.assertEqual(ambiguous, [])
            self.assertTrue((root / "上月.xlsx").exists())
            self.assertTrue((root / "本月.xlsx").exists())
            self.assertEqual(
                {(item["from"], item["to"]) for item in renamed},
                {("202607.xlsx", "上月.xlsx"), ("202608.xlsx", "本月.xlsx")},
            )

    def test_low_confidence_still_gets_a_manual_confirmation_suggestion(self):
        rows = [
            {"uploaded": "a.xlsx", "candidates": [
                {"expected": "上月.xlsx", "score": 0.12},
                {"expected": "本月.xlsx", "score": 0.11},
            ]},
            {"uploaded": "b.xlsx", "candidates": [
                {"expected": "上月.xlsx", "score": 0.10},
                {"expected": "本月.xlsx", "score": 0.09},
            ]},
        ]

        result = _add_structural_fallback_recommendations(rows)

        self.assertEqual(result[0]["ai_recommended"], "上月.xlsx")
        self.assertEqual(result[1]["ai_recommended"], "本月.xlsx")
        self.assertEqual(result[0]["recommendation_source"], "structure_fallback")

    def test_tenant_structure_match_precedes_global_filename_match(self):
        tenant_asset = SimpleNamespace(
            file_name="tenant-base.xlsx",
            parsed_headers={"Sheet1": ["工号", "金额"]},
            sheet_summary=None,
        )
        global_asset = SimpleNamespace(
            file_name="工资.xlsx",
            parsed_headers={"Sheet1": ["无关列"]},
            sheet_summary=None,
        )

        asset, scope = _find_matching_asset(
            "工资.xlsx",
            {"sheets": {"数据": {"headers": {"工号": "A", "金额": "B"}}}},
            [tenant_asset], [global_asset],
        )

        self.assertIs(asset, tenant_asset)
        self.assertEqual(scope, "租户")

    def test_autofill_preserves_real_excel_extension(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_dir = root / "source"
            base_dir = root / "base"
            source_dir.mkdir()
            base_dir.mkdir()
            base_path = base_dir / "converted.xlsx"
            base_path.write_bytes(b"xlsx-content")
            asset = SimpleNamespace(
                file_path=str(base_path), file_name="converted.xlsx",
                parsed_headers={"Sheet1": ["工号", "金额"]}, sheet_summary=None,
                name="工资基础资料", id=7,
            )
            structure = {"files": {"工资.xls": {"sheets": {
                "数据": {"headers": {"工号": "A", "金额": "B"}},
            }}}}

            with patch("backend.utils.source_auto_filler._load_reference_assets",
                       return_value=([asset], [])):
                filled, missing = auto_fill_missing_sources(
                    str(source_dir), structure, "tenant-a", object())

            self.assertEqual(missing, [])
            self.assertEqual(filled[0]["file_name"], "工资.xls")
            self.assertEqual(filled[0]["stored_file_name"], "工资.xlsx")
            self.assertEqual((source_dir / "工资.xlsx").read_bytes(), b"xlsx-content")
            self.assertFalse((source_dir / "工资.xls").exists())


if __name__ == "__main__":
    unittest.main()
