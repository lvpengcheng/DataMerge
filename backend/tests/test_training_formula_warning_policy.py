from pathlib import Path

import pandas as pd


def test_formula_quality_errors_are_warnings_after_successful_save(monkeypatch, tmp_path):
    from backend.utils import excel_comparator, output_postprocess, training_validation

    output = tmp_path / "result.xlsx"
    output.write_bytes(b"saved workbook")
    report = {
        "formula_count": 4,
        "empty_cache_count": 0,
        "error_cache_count": 2,
        "invalid_ref_formula_count": 0,
        "external_formula_count": 1,
        "error_cache_samples": [{"sheet": "Summary", "cell": "A2", "value": "#N/A"}],
    }
    monkeypatch.setattr(
        output_postprocess, "finalize_output_workbook",
        lambda *args, **kwargs: {"calculated": True, "rows_processed": 3},
    )
    monkeypatch.setattr(excel_comparator, "inspect_formula_cache", lambda *args, **kwargs: report)

    result = training_validation._finalize_training_output_impl(str(output), "", None, None)

    assert result["calculated"] is True
    assert result["formula_warning"] is True
    assert result["formula_report"] == report


def test_comparison_keeps_headers_that_collide_after_whitespace_normalization(tmp_path):
    from backend.utils.excel_comparator import compare_dataframes

    columns = ["工号", "Bonus\nTax", "Bonus Tax"]
    expected = pd.DataFrame([["001", 10, 20]], columns=columns)
    result = pd.DataFrame([["001", 10, 20]], columns=columns)

    comparison = compare_dataframes(
        result, expected, str(tmp_path / "diff.xlsx"), primary_keys=["工号"])

    assert comparison["comparison_complete"] is True
    assert comparison["success"] is True
    assert comparison["total_cells"] == 2
    assert comparison["matched_cells"] == 2
    assert Path(comparison["output_file"]).exists()
