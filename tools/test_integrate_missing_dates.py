from datetime import datetime

import numpy as np
import pandas as pd
import pytest

from backend.utils.merge_engine import _is_empty
from backend.utils.integrate_engine import (
    CLEAR_CELL, build_key_index, eval_source_expr_cross, resolve_overwrites,
)


@pytest.mark.parametrize('value', [None, pd.NaT, pd.NA, float('nan'), np.datetime64('NaT')])
def test_scalar_missing_values_are_empty(value):
    assert _is_empty(value)


@pytest.mark.parametrize('value', [0, False, datetime(2026, 1, 1), 'NaT', 'nan'])
def test_real_values_and_literal_text_are_not_missing(value):
    assert not _is_empty(value)


def test_date_formula_clears_missing_and_unmatched_rows_but_preserves_real_dates():
    frame = pd.DataFrame({'员工号': ['empty', 'valid', 'duplicate', 'duplicate'],
                          '最早入职日期': pd.to_datetime([None, '2026-01-02', None, '2026-01-03'])})
    index = {'花名册.xlsx': {'cols': list(frame.columns), 'rows': build_key_index(frame, '员工号')}}
    formula = 'IF(((最早入职日期&"")<>""),(最早入职日期&""),"")'
    pairs = [{'a_col': '入职时间_AM', 'source_file': '花名册.xlsx', 'source_expr': formula}]
    for key in ['empty', 'unmatched']:
        assert eval_source_expr_cross(formula, '花名册.xlsx', index, key) is None
        assert resolve_overwrites(key, pairs, index)['入职时间_AM'] is CLEAR_CELL
    assert eval_source_expr_cross(formula, '花名册.xlsx', index, 'valid') == '2026-01-02 00:00:00'
    assert eval_source_expr_cross(formula, '花名册.xlsx', index, 'duplicate') == '2026-01-03 00:00:00'
