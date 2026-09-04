"""多表整合对比增强功能的回归测试。"""

import pandas as pd

from backend.api.tools import (
    _match_scheme_config,
    _required_cols_by_role,
    _resolved_scheme_config,
    _suggest_column_map,
)
from backend.utils.integrate_engine import (
    CLEAR_CELL,
    compute_diffs,
    eval_source_expr_cross,
    resolve_overwrites,
)


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


def test_cross_table_formula_uses_zero_for_partially_missing_keys():
    indexes = {
        "a.xlsx": {"cols": ["金额"], "rows": {"001": [{"金额": 20}]}},
        "b.xlsx": {"cols": ["金额"], "rows": {}},
        "c.xlsx": {"cols": ["金额"], "rows": {"001": [{"金额": 10}]}},
    }
    expr = "金额+b.xlsx.金额+c.xlsx.金额"
    assert eval_source_expr_cross(expr, "a.xlsx", indexes, "001") == 30
    assert resolve_overwrites("001", [{
        "a_col": "合计", "source_file": "a.xlsx", "source_expr": expr,
    }], indexes) == {"合计": 30}
    diffs = compute_diffs(
        pd.DataFrame([{"工号": "001", "合计": 0, "姓名": "测试"}]),
        indexes, "工号",
        [{"a_col": "合计", "source_file": "a.xlsx", "source_expr": expr}],
        "姓名", None,
    )
    assert len(diffs) == 1 and "30.00" in diffs[0]["差异类型"]

    # 默认来源表和 b 表都缺失，c 表存在：两个缺失项按 0，结果仍为 10。
    indexes["a.xlsx"]["rows"] = {}
    assert eval_source_expr_cross(expr, "a.xlsx", indexes, "001") == 10
    assert resolve_overwrites("001", [{
        "a_col": "合计", "source_file": "a.xlsx", "source_expr": expr,
    }], indexes) == {"合计": 10}

    # 所有引用均不存在时，不把纯 0 当成业务结果，保持为空。
    indexes["c.xlsx"]["rows"] = {}
    assert eval_source_expr_cross(expr, "a.xlsx", indexes, "001") is None
    cleared = resolve_overwrites("001", [{
        "a_col": "合计", "source_file": "a.xlsx", "source_expr": expr,
    }], indexes)
    assert cleared.get("合计") is CLEAR_CELL

    # 普通单列映射仍沿用旧语义：源行不存在时不覆盖主表旧值。
    assert resolve_overwrites("001", [{
        "a_col": "合计", "source_file": "a.xlsx", "source_expr": "金额",
    }], indexes) == {}
    assert compute_diffs(
        pd.DataFrame([{"工号": "001", "合计": 0}]), indexes, "工号",
        [{"a_col": "合计", "source_file": "a.xlsx", "source_expr": expr}],
        None, None,
    ) == []
