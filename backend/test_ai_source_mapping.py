import json
import pytest
from backend.utils.fast_header_matcher import FastHeaderMatcher
from backend.utils.ai_source_mapping import validate_mapping


def schemas():
    return ([{'file_name': 'old.xlsx', 'sheet_name': '工资', 'headers': {'工号': 'A', '金额': 'B'}}],
            [{'file_name': 'new.xlsx', 'file_path': '/input/new.xlsx', 'sheet_name': 'Payroll',
              'headers': {'Employee ID': 'A', 'Pay': 'B'}}])


def test_ai_fallback_reuses_parse_and_loads_validated_mapping(monkeypatch):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    matcher = FastHeaderMatcher()
    calls = []
    monkeypatch.setattr(matcher, '_build_training_sheets', lambda *_: training)
    monkeypatch.setattr(matcher, '_parse_all_files_full', lambda *a, **k: (actual, {}))
    monkeypatch.setattr(matcher, '_match_by_training_base', lambda *a: {'success': False, 'error': '名称不同'})
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda name: name)
    def reply(provider, *args, **kwargs):
        calls.append(provider)
        return json.dumps({'mappings': [{'training_id': 0, 'actual_id': 0,
                            'columns': {'Employee ID': '工号', 'Pay': '金额'}}]})
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    def load(mapping, *args, **kwargs):
        assert mapping['new.xlsx']['sheet_mapping'] == {'Payroll': '工资'}
        assert mapping['new.xlsx']['header_mapping']['Pay'] == '金额'
        return {'工资': {'validated': True}}
    monkeypatch.setattr(matcher, '_build_pre_loaded_from_memory', load)
    result = matcher.match_parse_and_prepare({'files': {}}, ['new.xlsx'], ai_provider_name='claude')
    assert result[0] and result[3]['工资']['validated']
    assert calls == ['claude']


@pytest.mark.parametrize('columns', [{}, {'Missing': '工号', 'Pay': '金额'}, {'Employee ID': '工号'}])
def test_invalid_ai_columns_are_rejected(columns):
    training, actual = schemas()
    with pytest.raises(ValueError):
        validate_mapping(FastHeaderMatcher(), training, actual,
                         {'mappings': [{'training_id': 0, 'actual_id': 0, 'columns': columns}]})


def test_reused_actual_sheet_is_rejected():
    training, actual = schemas()
    training.append(dict(training[0], sheet_name='奖金'))
    with pytest.raises(ValueError, match='重复'):
        validate_mapping(FastHeaderMatcher(), training, actual, {'mappings': [
            {'training_id': i, 'actual_id': 0, 'columns': {'Employee ID': '工号', 'Pay': '金额'}}
            for i in range(2)]})


def test_ai_only_receives_unresolved_tables_and_columns(monkeypatch):
    from backend.ai_engine import ai_provider
    from backend.utils.ai_source_mapping import match_sources_with_ai
    training = [
        {'file_name': 'a.xlsx', 'sheet_name': '已匹配', 'headers': {'ID': 'A'}},
        {'file_name': 'b.xlsx', 'sheet_name': '工资', 'headers': {'ID': 'A', '金额': 'B'}}]
    actual = [dict(training[0], file_path='/a.xlsx'),
              {'file_name': 'b.xlsx', 'sheet_name': 'Payroll', 'file_path': '/b.xlsx',
               'headers': {'ID': 'A', 'Pay': 'B'}}]
    determined = [{'train_file': 'a.xlsx', 'train_sheet': '已匹配',
                  'input_file': 'a.xlsx', 'input_sheet': '已匹配', 'col_mapping': {'ID': 'ID'}}]
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    def reply(provider, messages, **kwargs):
        prompt = messages[0]['content']
        assert '已匹配' not in prompt and 'a.xlsx' not in prompt
        assert '"missing_columns": ["金额"]' in prompt
        assert '"candidate_columns": ["Pay"]' in prompt
        return json.dumps({'mappings': [{'training_id': 1, 'actual_id': 1, 'columns': {'Pay': '金额'}}]})
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, 'claude', determined)
    mapping = result['mapping']['file_mapping']
    assert mapping['a.xlsx']['header_mapping'] == {'ID': 'ID'}
    assert mapping['b.xlsx']['header_mapping'] == {'ID': 'ID', 'Pay': '金额'}
