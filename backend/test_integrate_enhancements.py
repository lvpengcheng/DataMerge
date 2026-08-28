"""多表整合对比增强功能的回归测试。"""

from backend.api.tools import (
    _match_scheme_config,
    _required_cols_by_role,
    _resolved_scheme_config,
    _suggest_column_map,
)
from backend.utils.integrate_engine import eval_source_expr_cross


def _scheme_config():
    return {
        "main_fp": "saved-main",
        "source_fps": ["saved-source"],
        "roles": [
            {"fp": "saved-main", "file": "主表.xlsx"},
            {"fp": "saved-source", "file": "5月社保.xlsx"},
        ],
        "cols_by_fp": {
            "saved-main": ["工号", "姓名", "个人养老", "已删除的无关列"],
            "saved-source": ["工号", "2026年5月社保个人养老", "已删除的无关列"],
        },
        "key_map_by_role": {"0": "工号", "1": "工号"},
        "overwrite_pairs": [{
            "a_col": "个人养老",
            "source_fp": "saved-source",
            "source_role": 1,
            "source_expr": "ROUND(2026年5月社保个人养老,2)",
            "source_col": "ROUND(2026年5月社保个人养老,2)",
        }],
        "compare_pairs": [],
        "name_col": "姓名",
        "output_mode": 1,
    }


def test_scheme_ignores_unreferenced_added_and_removed_columns():
    cfg = _scheme_config()
    required = _required_cols_by_role(cfg)
    assert "已删除的无关列" not in required[0]
    assert "已删除的无关列" not in required[1]

    files = [
        {"name": "新主表.xlsx", "fingerprint": "new-main",
         "columns": ["新增列", "工号", "姓名", "个人养老"]},
        {"name": "6月社保.xlsx", "fingerprint": "new-source",
         "columns": ["工号", "2026年6月社保个人养老", "新增列"]},
    ]
    match = _match_scheme_config(cfg, files)
    assert match is not None
    assert match["role_files"] == ["新主表.xlsx", "6月社保.xlsx"]


def test_month_header_change_produces_confirmable_mapping_and_resolved_formula():
    result = _suggest_column_map(
        ["2026年5月社保个人养老"], ["2026年6月社保个人养老"])
    assert result["missing"] == []
    assert result["suggestions"][0]["method"] == "period"

    cfg = _scheme_config()
    files = [
        {"name": "新主表.xlsx", "fingerprint": "new-main",
         "columns": ["工号", "姓名", "个人养老"]},
        {"name": "6月社保.xlsx", "fingerprint": "new-source",
         "columns": ["工号", "2026年6月社保个人养老"]},
    ]
    match = _match_scheme_config(cfg, files)
    resolved = _resolved_scheme_config(cfg, match)
    assert resolved["overwrite_pairs"][0]["source_expr"] == "ROUND(2026年6月社保个人养老,2)"


def test_excel_round_and_if_formula_subset():
    indexes = {
        "6月社保.xlsx": {
            "cols": ["养老", "调整"],
            "rows": {"001": [{"养老": 10.235, "调整": 2}, {"养老": 1, "调整": 3}]},
        }
    }
    assert eval_source_expr_cross(
        "ROUND(养老-调整,2)", "6月社保.xlsx", indexes, "001") == 6.24
    assert eval_source_expr_cross(
        "IF(养老>10,ROUND(养老,2),0)", "6月社保.xlsx", indexes, "001") == 11.24
