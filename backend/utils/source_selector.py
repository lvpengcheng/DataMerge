"""Deterministic source selection: ambiguity must not silently choose another table."""
import re


def find_source_sheet(source_sheets, target_columns=None, sheet_name_hint=None,
                      salary_year=None, salary_month=None):
    keys = sorted(source_sheets)
    if not keys:
        raise KeyError('没有可用源表')
    candidates = keys
    if sheet_name_hint:
        hint = str(sheet_name_hint).strip().casefold()
        exact = [k for k in keys if str(k).casefold() == hint]
        if exact:
            candidates = exact
        else:
            candidates = [k for k in keys if hint in str(k).casefold()]
            if not candidates:
                raise KeyError(f'源表提示 {sheet_name_hint!r} 无匹配，可用表: {keys}')
    if salary_year and salary_month and len(candidates) > 1:
        year, month = int(salary_year), int(salary_month)
        if not 1 <= month <= 12:
            raise ValueError('薪资月份须为 1 至 12')
        # Boundaries prevent 1月 matching 11月 and 202601 matching 2026010.
        full = re.compile(rf'(?<!\d){year}(?:[-年/]?0?{month})(?:月)?(?!\d)')
        matching = [k for k in candidates if full.search(str(k))]
        if not matching:
            month_only = re.compile(rf'(?<!\d)0?{month}月')
            matching = [k for k in candidates if month_only.search(str(k))]
        if matching:
            candidates = matching
    if target_columns:
        required = set(target_columns)
        scores = {k: len(required & set(source_sheets[k]['df'].columns)) for k in candidates}
        best = max(scores.values(), default=0)
        if best == 0:
            raise KeyError(f'源表中不存在所需列 {list(target_columns)}')
        candidates = [k for k in candidates if scores[k] == best]
    if len(candidates) != 1:
        raise ValueError(f'源表选择不唯一，请明确文件/Sheet 或关联规则: {candidates}')
    return candidates[0]
