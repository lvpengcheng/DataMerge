"""General field matching: noisy headers, bounded AI batches and safe acceptance."""
import json

import pytest

from backend.utils.field_matching import (
    SYSTEM_PROMPT, header_context, make_fields, match_fields, parse_response, value_profile,
)


def fields(target_cols, source_cols, target_context=None, source_context=None):
    meta = [{'name': '模板.xlsx', 'columns': target_cols, 'field_context': target_context or {}},
            {'name': '来源.xlsx', 'sheet': '明细', 'columns': source_cols, 'field_context': source_context or {}}]
    return make_fields(meta, '模板.xlsx', target_cols, [{'file': '来源.xlsx', 'col': col} for col in source_cols])


class Provider:
    def __init__(self, decide):
        self.decide, self.calls = decide, []

    def chat(self, messages, **kwargs):
        payload = json.loads(messages[1]['content'])
        self.calls.append((payload, kwargs))
        return json.dumps({'matches': [self.decide(t, payload['source_fields']) for t in payload['target_fields']]})


def prediction(target, sources, status='matched', confidence=.98):
    return {'target_id': target['id'], 'status': status, 'source_ids': [sources[0]['id']],
            'confidence': confidence, 'reason': '测试字段语义对应', 'suggestion': ''}


@pytest.mark.parametrize('token', ['1723451238904', '9999999999999'])
def test_system_ids_do_not_change_matching_and_original_keys_survive(token):
    original = f'工资-应发工资-HiddenHeader:_widget_{token}-应发金额'
    meta = [{'name': '模板.xlsx', 'columns': [original], 'field_context': {original:
        header_context(original, ['工资-应发工资', f'HiddenHeader:_widget_{token}', '应发金额'])}},
        {'name': '任意客户工资.xlsx', 'columns': ['应发工资']}]
    targets, sources = make_fields(meta, '模板.xlsx', [original], [{'file': '任意客户工资.xlsx', 'col': '应发工资'}])
    result = match_fields(targets, sources)['results'][0]
    assert result['status'] == 'matched'
    assert result['a_col'] == original
    assert result['label'] == '应发金额'
    assert result['sources'][0]['source_col'] == '应发工资'


def test_other_businesses_and_english_headers_use_same_provenance_logic():
    original = 'Inventory-Available Qty-HiddenHeader:_widget_123-可用数量'
    metadata = [{'name': 'report.xlsx', 'columns': [original]},
                {'name': 'Inventory.xlsx', 'columns': ['Available Qty']}]
    ts, ss = make_fields(metadata, 'report.xlsx', [original], [{'file': 'Inventory.xlsx', 'col': 'Available Qty'}])
    assert match_fields(ts, ss)['results'][0]['status'] == 'matched'


def test_header_layers_keep_company_year_rate_and_business_notes():
    context = header_context('公司养老', ['2026年', '公司部分', '养老保险(16%)'])
    assert context['label'] == '养老保险(16%)'
    assert context['instructions'] == ['2026年', '公司部分']
    assert '16%' in context['label']


@pytest.mark.parametrize('name,layers,expected', [
    ('个人-公积金', None, ['个人-公积金']),
    ('公司-公积金', None, ['公司-公积金']),
    ('a-个人-养老', None, ['a-个人-养老']),
    ('b-公司-养老金', None, ['b-公司-养老金']),
    ('a-个人-养老', ['a', '个人', '养老'], ['a', '个人', '养老']),
    ('b-公司-养老金', ['b', '公司', '养老金'], ['b', '公司', '养老金']),
])
def test_ordinary_header_paths_reach_ai_without_discarding_parent_levels(name, layers, expected):
    context = header_context(name, layers)
    assert context['header_parts'] == expected
    assert not context['technical_tokens']
    ts, ss = fields([name], ['待判断来源'], {name: context})
    provider = Provider(prediction)
    match_fields(ts, ss, True, provider)
    submitted = provider.calls[0][0]['target_fields'][0]
    assert submitted['header']['path'] == expected
    assert submitted['header']['preceding_context_role'] == 'unclassified'
    assert 'instructions' not in submitted


@pytest.mark.parametrize('layered', [False, True])
@pytest.mark.parametrize('target,source', [
    ('个人-公积金', '公司-公积金'),
    ('a-个人-养老', 'b-公司-养老金'),
])
def test_parent_business_qualifiers_prevent_wrong_auto_match(target, source, layered):
    target_context = {target: header_context(target, target.split('-'))} if layered else None
    source_context = {source: header_context(source, source.split('-'))} if layered else None
    ts, ss = fields([target], [source], target_context, source_context)
    result = match_fields(ts, ss, True, Provider(prediction))['results'][0]
    assert result['status'] == 'needs_review'


def test_id_response_resolves_simplified_label_to_the_original_long_header():
    raw = '本表说明-HiddenHeader:_widget_444-实发金额'
    ts, ss = fields([raw], ['到账工资'])
    provider = Provider(prediction)
    result = match_fields(ts, ss, True, provider)
    assert result['results'][0]['a_col'] == raw
    assert result['results'][0]['status'] == 'matched'
    payload, kwargs = provider.calls[0]
    assert 'widget' not in json.dumps(payload, ensure_ascii=False)
    assert payload['target_fields'][0]['header']['leaf_label'] == '实发金额'
    assert kwargs['require_complete'] is True
    assert '不得默认相加' in SYSTEM_PROMPT


def test_multiple_sources_are_reviewed_instead_of_summed():
    meta = [{'name': '主表.xlsx', 'columns': ['金额']}, {'name': '本月.xlsx', 'columns': ['金额']},
            {'name': '上月.xlsx', 'columns': ['金额']}]
    ts, ss = make_fields(meta, '主表.xlsx', ['金额'], [{'file': f, 'col': '金额'} for f in ['本月.xlsx', '上月.xlsx']])
    row = match_fields(ts, ss)['results'][0]
    assert row['status'] == 'needs_review' and len(row['sources']) == 2


@pytest.mark.parametrize('target,source', [('公司社保', '个人社保'), ('累计收入', '本月收入'),
                                          ('税前金额', '税后金额'), ('应发工资', '实发工资')])
def test_conflicting_business_qualifiers_cannot_be_auto_accepted(target, source):
    ts, ss = fields([target], [source])
    result = match_fields(ts, ss, True, Provider(prediction))['results'][0]
    assert result['status'] == 'needs_review'


@pytest.mark.parametrize('status', ['formula', 'constant', 'needs_review'])
def test_calculation_and_constant_suggestions_are_never_auto_written(status):
    ts, ss = fields(['目标'], ['来源'])
    result = match_fields(ts, ss, True, Provider(lambda t, s: {**prediction(t, s, status), 'suggestion': '请确认业务规则'}))
    assert result['results'][0]['status'] == status
    assert result['results'][0]['suggestion'] == '请确认业务规则'


def test_bad_ids_missing_targets_and_invalid_scores_are_exposed():
    ts, ss = fields(['目标'], ['来源'])
    result, issues = parse_response(json.dumps({'matches': [{'target_id': 'T0001', 'status': 'matched',
        'source_ids': ['S9999'], 'confidence': 1}]}), ts, ss)
    assert not result and issues
    result, _ = parse_response(json.dumps({'matches': [{'target_id': 'T0001', 'status': 'matched',
        'source_ids': ['S0001'], 'confidence': True}]}), ts, ss)
    assert result['T0001']['confidence'] == 0
    with pytest.raises(ValueError, match='JSON_INVALID'):
        parse_response('{"matches":[', ts, ss)


def test_ai_failure_keeps_deterministic_matches_and_marks_others_review():
    ts, ss = fields(['已知', '未知目标'], ['已知', '其他来源'])
    class Failing:
        def chat(self, *args, **kwargs):
            raise ValueError('AI_RESPONSE_TRUNCATED')
    result = match_fields(ts, ss, True, Failing())
    assert [r['status'] for r in result['results']] == ['matched', 'needs_review']
    assert any('截断' in w for w in result['warnings'])


def test_every_source_after_200_is_processed_and_batches_merge():
    ts, ss = fields(['目标'], [f'field_{i}' for i in range(205)])
    def choose(target, sources):
        found = [s for s in sources if s['id'] == 'S0205']
        return prediction(target, found) if found else {'target_id': target['id'], 'status': 'unmatched', 'source_ids': []}
    provider = Provider(choose)
    result = match_fields(ts, ss, True, provider)
    assert len(provider.calls) == 3
    assert result['results'][0]['status'] == 'matched'
    assert result['results'][0]['sources'][0]['source_col'] == 'field_204'


def test_ai_contract_preserves_hierarchical_join_key_and_table_context():
    ts, ss = fields(['a-个人-养老'], ['b-公司-养老金'])
    provider = Provider(prediction)
    match_fields(ts, ss, True, provider, {'模板.xlsx': '人员-身份证号', '来源.xlsx': '证件号码'})
    payload = provider.calls[0][0]
    assert payload['schema_version'] == 2
    tables = {table['id']: table for table in payload['tables']}
    for field in payload['target_fields'] + payload['source_fields']:
        assert field['table_id'] in tables
        assert set(field) == {'id', 'table_id', 'header', 'profile'}
    assert tables[payload['source_fields'][0]['table_id']]['sheet'] == '明细'
    assert payload['join_keys'][0]['header']['path'] == ['人员-身份证号']
    assert payload['batch_context']['sources_complete'] is True


def test_source_batch_coverage_is_explicit_and_table_ids_are_stable():
    ts, ss = fields(['目标'], [f'field_{i}' for i in range(81)])
    provider = Provider(lambda t, s: {'target_id': t['id'], 'status': 'unmatched', 'source_ids': []})
    match_fields(ts, ss, True, provider)
    first, second = [call[0] for call in provider.calls]
    assert first['tables'] == second['tables']
    assert [p['batch_context']['source_count'] for p in (first, second)] == [80, 1]
    assert all(p['batch_context']['source_total'] == 81 and not p['batch_context']['sources_complete']
               for p in (first, second))


def test_every_target_after_200_is_processed():
    ts, ss = fields([f'目标{i}' for i in range(201)], ['来源'])
    provider = Provider(lambda t, s: {'target_id': t['id'], 'status': 'unmatched', 'source_ids': []})
    result = match_fields(ts, ss, True, provider)
    assert len(result['results']) == 201 and len(provider.calls) == 21
    assert not result['warnings']


def test_type_summary_has_no_actual_employee_values():
    profile = value_profile(['PRIVATE-NAME', '110101199001011234', '6222000000000000', None])
    assert profile['sampled_count'] == 4
    text = json.dumps(profile)
    assert 'PRIVATE' not in text and '6222' not in text and '110101' not in text


@pytest.mark.parametrize('failure', ['AI_RESPONSE_TRUNCATED', 'AI_FIELD_MATCH_JSON_INVALID'])
def test_incomplete_batches_retry_smaller_target_sets(failure):
    ts, ss = fields(['目标甲', '目标乙'], ['来源'])
    class RetryProvider(Provider):
        def chat(self, messages, **kwargs):
            payload = json.loads(messages[1]['content'])
            if len(payload['target_fields']) > 1:
                raise ValueError(failure)
            return super().chat(messages, **kwargs)
    report = match_fields(ts, ss, True, RetryProvider(prediction))
    assert report['ai_batches'] == 3
    assert not report['warnings']
    assert [row['status'] for row in report['results']] == ['matched', 'matched']
    assert report['trace'][0]['retry'] == 'split_targets'


def test_template_instruction_conflicting_with_label_requires_review():
    original = '工资-税前调整-HiddenHeader:_widget_123-税后调整'
    meta = [{'name': '模板.xlsx', 'columns': [original]},
            {'name': '工资.xlsx', 'columns': ['税前调整']}]
    ts, ss = make_fields(meta, '模板.xlsx', [original], [{'file': '工资.xlsx', 'col': '税前调整'}])
    row = match_fields(ts, ss)['results'][0]
    assert row['status'] == 'needs_review'
    assert '冲突' in row['reason']
