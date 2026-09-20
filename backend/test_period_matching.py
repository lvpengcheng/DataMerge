"""月份角色确定性匹配回归测试。"""
import json

from backend.ai_engine import ai_provider
from backend.utils.ai_source_mapping import match_sources_with_ai
from backend.utils.fast_header_matcher import FastHeaderMatcher


HEADERS = {"工号": "A", "金额": "B"}
CONTEXT = {"salary_year": 2026, "salary_month": 8}


def _training():
    return [
        {"file_name": "四川外包人员工资上月.xlsx", "sheet_name": "非风控", "headers": HEADERS},
        {"file_name": "四川外包人员工资本月.xlsx", "sheet_name": "非风控", "headers": HEADERS},
    ]


def _actual():
    return [
        {"file_name": "四川外包人员工资202607.xlsx", "file_path": "/tmp/202607.xlsx",
         "sheet_name": "非风控", "headers": HEADERS},
        {"file_name": "四川外包人员工资202608.xlsx", "file_path": "/tmp/202608.xlsx",
         "sheet_name": "非风控", "headers": HEADERS},
    ]


def test_prev_and_current_month_are_deterministic():
    result = FastHeaderMatcher().match_headers_only(_training(), _actual(), None, CONTEXT)
    assert result["success"]
    mapping = result["mapping"]["file_mapping"]
    assert mapping["四川外包人员工资202607.xlsx"]["expected_file"] == "四川外包人员工资上月.xlsx"
    assert mapping["四川外包人员工资202608.xlsx"]["expected_file"] == "四川外包人员工资本月.xlsx"


def test_single_wrong_month_is_rejected():
    training = [{"file_name": "四川外包人员工资上月.xlsx", "sheet_name": "非风控", "headers": HEADERS}]
    actual = [{"file_name": "四川外包人员工资202608.xlsx", "file_path": "/tmp/202608.xlsx",
               "sheet_name": "非风控", "headers": HEADERS}]
    result = FastHeaderMatcher().match_headers_only(training, actual, None, CONTEXT)
    assert not result["success"]
    assert not (result.get("mapping") or {}).get("file_mapping")


def test_ai_wrong_month_is_dropped(monkeypatch):
    training = [{"file_name": "四川外包人员工资上月.xlsx", "sheet_name": "非风控", "headers": HEADERS}]
    actual = [{"file_name": "四川外包人员工资202608.xlsx", "file_path": "/tmp/202608.xlsx",
               "sheet_name": "非风控", "headers": HEADERS}]
    monkeypatch.setattr(ai_provider.AIProviderFactory, "create_provider", lambda _: object())
    monkeypatch.setattr(ai_provider, "chat_with_timeout", lambda *args, **kwargs: json.dumps({
        "mappings": [{"training_id": 0, "actual_id": 0,
                      "columns": {"工号": "工号", "金额": "金额"}}]
    }, ensure_ascii=False))
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, "claude", None, CONTEXT)
    assert not result["mapping"]["file_mapping"]


def test_month_only_names_are_inferred_from_salary_month():
    from backend.utils.period_matching import period_candidate_allowed, period_index

    context = {'salary_year': 2026, 'salary_month': 8}
    current_idx = 2026 * 12 + 8 - 1
    assert period_index('7月速创', current_idx) == 2026 * 12 + 7 - 1
    assert period_index('8月速创', current_idx) == 2026 * 12 + 8 - 1
    assert not period_candidate_allowed(
        {'file_name': '四川导出当月.xlsx'}, {'file_name': '7月速创.xlsx'}, context)
    assert period_candidate_allowed(
        {'file_name': '四川导出当月.xlsx'}, {'file_name': '8月速创.xlsx'}, context)
    assert period_candidate_allowed(
        {'file_name': '四川导出上月.xlsx'}, {'file_name': '7月速创.xlsx'}, context)


def test_month_only_actual_file_is_filtered_by_period_role():
    training = [{'file_name': '四川导出当月.xlsx', 'sheet_name': '第一批', 'headers': HEADERS}]
    actual = [
        {'file_name': '7月速创.xlsx', 'file_path': '/tmp/7月速创.xlsx',
         'sheet_name': '第一批', 'headers': HEADERS},
        {'file_name': '四川太保 202608 速创导出.xlsx',
         'file_path': '/tmp/202608.xlsx', 'sheet_name': '第一批', 'headers': HEADERS},
    ]
    result = FastHeaderMatcher().match_headers_only(training, actual, None, CONTEXT)
    assert result['success']
    assert result['mapping']['file_mapping']['四川太保 202608 速创导出.xlsx']['expected_file'] == '四川导出当月.xlsx'
    assert '7月速创.xlsx' not in result['mapping']['file_mapping']
