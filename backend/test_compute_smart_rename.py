import tempfile
import unittest
from pathlib import Path

from backend.utils.source_auto_filler import (
    _add_structural_fallback_recommendations,
    _period_index,
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


if __name__ == "__main__":
    unittest.main()
