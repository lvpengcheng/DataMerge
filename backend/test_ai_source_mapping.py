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
        assert "\"global_training\"" in prompt and "已匹配" in prompt
        assert "\"global_actual\"" in prompt and "a.xlsx" in prompt
        assert '"missing_columns": ["金额"]' in prompt
        assert '"candidate_columns": ["Pay"]' in prompt
        return json.dumps({'mappings': [{'training_id': 1, 'actual_id': 1, 'columns': {'Pay': '金额'}}]})
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, 'claude', determined)
    mapping = result['mapping']['file_mapping']
    assert mapping['a.xlsx']['header_mapping'] == {'ID': 'ID'}
    assert mapping['b.xlsx']['header_mapping'] == {'ID': 'ID', 'Pay': '金额'}


@pytest.mark.parametrize('headers', [
    {'工号': 'C', '金额': 'A'},
    {' 工号 ': 'A', '金 额': 'B'},
])
def test_renamed_file_and_sheet_use_same_structure_without_ai(monkeypatch, headers):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    actual[0]['headers'] = headers
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider',
                        lambda *_: pytest.fail('same column structure must not call AI'))
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert result['success'] and result['match_method'] == 'structure'
    assert not result.get('needs_confirmation')
    assert result['mapping']['file_mapping']['new.xlsx']['sheet_mapping'] == {'Payroll': '工资'}


def test_exact_training_structure_skips_ai(monkeypatch):
    from backend.ai_engine import ai_provider
    training, _ = schemas()
    actual = [dict(training[0], file_path='/input/old.xlsx')]
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider',
                        lambda *_: pytest.fail('exact training structure must calculate directly'))
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert result['success'] and result['match_method'] == 'structure'
    assert result['structure_identical'] and not result.get('needs_confirmation')


def test_ai_popup_requires_structure_file_and_sheet_names_all_changed(monkeypatch):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    calls = []
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda *_: object())
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', lambda *args, **kwargs:
                        calls.append(1) or json.dumps({'mappings': [{
                            'training_id': 0, 'actual_id': 0,
                            'columns': {'Employee ID': '工号', 'Pay': '金额'},
                        }]}))
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert result['success'] and result['match_method'] == 'ai'
    assert result['needs_confirmation'] and calls == [1]


@pytest.mark.parametrize('keep_identity', ['file', 'sheet'])
def test_structure_change_does_not_call_ai_when_one_identity_is_unchanged(monkeypatch, keep_identity):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    if keep_identity == 'file':
        actual[0]['file_name'] = training[0]['file_name']
    else:
        actual[0]['sheet_name'] = training[0]['sheet_name']
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider',
                        lambda *_: pytest.fail('AI requires both file and Sheet names to change'))
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert not result['success'] and result.get('match_method') != 'ai'


def test_identical_schemas_with_ambiguous_names_do_not_pick_first():
    training, actual = schemas()
    actual[0]['headers'] = training[0]['headers']
    actual.append(dict(actual[0], file_name='another.xlsx', file_path='/input/another.xlsx'))
    result = FastHeaderMatcher().match_headers_only(training, actual)
    assert not result['success'] and not result['mapping']['file_mapping']


def test_structure_change_with_same_file_and_sheet_does_not_open_ai_review(monkeypatch):
    from backend.ai_engine import ai_provider
    training = [{'file_name': '工资.xlsx', 'sheet_name': 'S', 'headers': {'本月个人缴费金额': 'A'}}]
    actual = [dict(training[0], file_path='/input/工资.xlsx', headers={'本月单位缴费金额': 'A'})]
    calls = []
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', lambda *a, **k: calls.append(1) or '{"mappings":[]}')
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert not result['success'] and calls == []


def test_ai_receives_hierarchical_script_context_without_sample_literals(monkeypatch):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    def reply(provider, messages, **kwargs):
        prompt = messages[0]['content']
        assert '文件到文件' in prompt and 'Sheet 到 Sheet' in prompt
        assert '本阶段不要处理列映射' in prompt and 'salary_year' in prompt
        assert 'current_period_index' in prompt
        assert '工资模板.xlsx' in prompt and 'df' in prompt and 'Pay' in prompt
        assert 'private-row-name' not in prompt and 'private-token' not in prompt
        return json.dumps({'mappings': [{'training_id': 0, 'actual_id': 0,
            'columns': {'Employee ID': '工号', 'Pay': '金额'},
            'file_reason': '两份文件均为工资明细', 'sheet_reason': 'Payroll 对应工资明细表',
            'code_evidence': 'fill_template 从 Pay 列读取工资金额',
            'column_reasons': {'Employee ID': 'Employee ID 与工号均为员工关联键', 'Pay': 'Pay 对应工资金额'}}]})
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude', {
        'script_content': 'TOKEN = "private-token"\nTEMPLATE_NAME = "工资模板.xlsx"\ndef fill_template(df):\n    employee = "private-row-name"\n    return df["Pay"]',
        'template_name': '工资模板.xlsx', 'target_sheets': ['结果']})
    assert result['success'] and result['needs_confirmation']
    suggestion = result['ai_suggestions'][0]
    assert suggestion['reason'] == 'Employee ID 与工号均为员工关联键'
    assert 'Payroll 对应工资明细表' in suggestion['sheet_reason']
    assert 'fill_template' in suggestion['sheet_reason']
    assert suggestion['recommendation_source'] == 'ai'


def test_same_structure_resolves_stale_name_candidates_without_confirmation(monkeypatch):
    from backend.utils import compute_precheck as pre
    from backend.utils.compute_ingest import IngestMeta, resolve_with_confirmations
    training, actual = schemas()
    actual[0]['headers'] = training[0]['headers']
    meta = IngestMeta(train_sheets=training, input_sheets=actual,
        source_structure={'files': {'old.xlsx': {'sheets': {'工资': {'headers': training[0]['headers']}}}}},
        missing_files=['old.xlsx'], rename_candidates=[{'uploaded': 'new.xlsx',
            'candidates': [{'expected': 'old.xlsx', 'score': .5}]}])
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *a: None)
    result = resolve_with_confirmations(meta, skip_history_check=True)
    assert result.ok and not result.mapping_requires_confirmation
    assert not result.rename_candidates and not result.missing_files
    assert result.file_mapping['new.xlsx']['expected_file'] == 'old.xlsx'


def test_ai_recommendation_only_needs_one_confirmation(monkeypatch):
    from backend.ai_engine import ai_provider
    from backend.utils import compute_precheck as pre
    from backend.utils.compute_ingest import IngestMeta, resolve_with_confirmations
    training, actual = schemas()
    meta = IngestMeta(train_sheets=training, input_sheets=actual, ai_provider_name='claude',
        source_structure={'files': {'old.xlsx': {'sheets': {'工资': {'headers': training[0]['headers']}}}}})
    calls = []
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *a: None)
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', lambda *a, **k: calls.append(1) or json.dumps({
        'mappings': [{'training_id': 0, 'actual_id': 0, 'columns': {'Employee ID': '工号', 'Pay': '金额'}}]}))
    first = resolve_with_confirmations(meta, skip_history_check=True)
    assert not first.ok and first.mapping_requires_confirmation
    second = resolve_with_confirmations(meta, confirmed_mapping={'file_mapping': first.file_mapping}, skip_history_check=True)
    assert second.ok and not second.mapping_requires_confirmation and calls == [1]



def test_ingest_does_not_rename_or_call_ai_when_structure_is_complete(monkeypatch, tmp_path):
    from backend.utils import source_auto_filler as filler
    from backend.utils.compute_ingest import ingest_source_dir
    training, actual = schemas()
    actual[0]['headers'] = training[0]['headers']
    path = tmp_path / 'new.xlsx'
    path.touch()
    actual[0]['file_path'] = str(path)
    monkeypatch.setattr(FastHeaderMatcher, 'parse_inputs', lambda *a, **k: (actual, {}))
    monkeypatch.setattr(filler, 'auto_rename_uploaded_by_combined_score',
                        lambda **k: pytest.fail('structure must precede filename scoring'))
    def fill(**kwargs):
        assert kwargs['assume_present'] == {'old.xlsx'}
        return [], []
    monkeypatch.setattr(filler, 'auto_fill_missing_sources', fill)
    meta, _ = ingest_source_dir(str(tmp_path), {'files': {'old.xlsx': {
        'sheets': {'工资': {'headers': training[0]['headers']}}}}},
        tenant_id='test', db_session=object(), ai_provider_name='claude')
    assert not meta.rename_candidates and not meta.missing_files
    assert path.exists() and not (tmp_path / 'old.xlsx').exists()


def test_extra_template_instance_is_kept_in_resolved_workbook():
    training = [{'file_name': 'a.xlsx', 'sheet_name': '原表',
                 'headers': {'工号': 'A', '金额': 'B', '月份': 'C'}}]
    actual = [dict(training[0], file_path='/a.xlsx'),
              dict(training[0], file_path='/a.xlsx', sheet_name='新增表')]
    result = FastHeaderMatcher().match_headers_only(training, actual)
    assert result['success']
    assert result['mapping']['file_mapping']['a.xlsx']['sheet_mapping'] == {'原表': '原表', '新增表': '新增表'}


def test_ai_column_meaning_is_scoped_by_sheet():
    training = [{'file_name': 'old.xlsx', 'sheet_name': sheet, 'headers': {col: 'A'}}
                for sheet, col in [('工资', '工资金额'), ('社保', '缴费金额')]]
    actual = [{'file_name': 'new.xlsx', 'file_path': '/new.xlsx', 'sheet_name': sheet,
               'headers': {'Amount': 'A'}} for sheet in ['Pay', 'Insurance']]
    result = validate_mapping(FastHeaderMatcher(), training, actual, {'mappings': [
        {'training_id': i, 'actual_id': i, 'columns': {'Amount': col}}
        for i, col in enumerate(['工资金额', '缴费金额'])]})
    assert result['mapping']['file_mapping']['new.xlsx']['header_mapping_by_sheet'] == {
        'Pay': {'Amount': '工资金额'}, 'Insurance': {'Amount': '缴费金额'}}



def test_single_filename_candidate_does_not_rename_unrelated_upload(tmp_path, monkeypatch):
    from backend.utils.compute_ingest import ingest_source_dir
    training, actual = schemas()
    path = tmp_path / 'new.xlsx'
    path.write_bytes(b'original')
    actual[0]['file_path'] = str(path)
    monkeypatch.setattr(FastHeaderMatcher, 'parse_inputs', lambda *a, **k: (actual, {}))
    meta, _ = ingest_source_dir(str(tmp_path), {'files': {'old.xlsx': {
        'sheets': {'工资': {'headers': training[0]['headers']}}}}})
    assert path.read_bytes() == b'original' and not (tmp_path / 'old.xlsx').exists()
    assert meta.rename_candidates[0]['uploaded'] == 'new.xlsx'
    assert not meta.auto_renamed


def test_confirmed_filename_does_not_leak_virtual_names_into_ai(monkeypatch):
    from backend.ai_engine import ai_provider
    from backend.utils.compute_ingest import IngestMeta, resolve_with_confirmations
    from backend.utils import compute_precheck as pre
    training, actual = schemas()
    meta = IngestMeta(train_sheets=training, input_sheets=actual, ai_provider_name='claude',
        source_structure={'files': {'old.xlsx': {'sheets': {'工资': {'headers': training[0]['headers']}}}}})
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *a: None)
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    def reply(provider, messages, **kwargs):
        prompt = messages[0]['content']
        assert '"file": "new.xlsx"' in prompt and '"original_sheet": "Payroll"' in prompt
        return json.dumps({'mappings': [{'training_id': 0, 'actual_id': 0,
            'columns': {'Employee ID': '工号', 'Pay': '金额'}}]})
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = resolve_with_confirmations(meta, confirmed_renames={'new.xlsx': 'old.xlsx'}, skip_history_check=True)
    assert result.actual_sources[0]['original_file'] == 'new.xlsx'
    assert all(s['suggested_path'].startswith('new.xlsx > Payroll > ') for s in result.ai_suggestions)


def test_ai_partial_columns_survive_instead_of_discarding_all_recommendations(monkeypatch):
    from backend.ai_engine import ai_provider
    training, actual = schemas()
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())
    monkeypatch.setattr(ai_provider, 'chat_with_timeout', lambda *a, **k: json.dumps({
        'mappings': [{'training_id': 0, 'actual_id': 0, 'columns': {'Employee ID': '工号'}}]}))
    result = FastHeaderMatcher().match_headers_only(training, actual, 'claude')
    assert not result['success'] and result['needs_confirmation']
    assert result['mapping']['file_mapping']['new.xlsx']['header_mapping'] == {'Employee ID': '工号'}
    assert result['ai_suggestions'][0]['suggested_path'] == 'new.xlsx > Payroll > Employee ID'


def test_xls_conversion_retains_original_display_name_and_does_not_overwrite(tmp_path, monkeypatch):
    from backend.utils import compute_ingest as ingest, source_normalizer
    monkeypatch.setattr(ingest, 'detect_encrypted_files', lambda _: [])
    (tmp_path / 'legacy.xls').write_bytes(b'old')
    (tmp_path / 'same.xls').write_bytes(b'xls')
    (tmp_path / 'same.xlsx').write_bytes(b'xlsx')
    def convert(path):
        from pathlib import Path
        original = Path(path)
        target = original.with_suffix('.xlsx')
        assert not target.exists()
        target.write_bytes(original.read_bytes())
        original.unlink()
        return str(target)
    monkeypatch.setattr(source_normalizer, 'convert_xls_to_xlsx', convert)
    ingest.prepare_source_dir(str(tmp_path))
    assert json.loads((tmp_path / '_upload_names.json').read_text(encoding='utf-8')) == {'legacy.xlsx': 'legacy.xls'}
    assert (tmp_path / 'same.xls').read_bytes() == b'xls'
    assert (tmp_path / 'same.xlsx').read_bytes() == b'xlsx'


def test_template_decryption_password_retry_and_same_named_source(tmp_path, monkeypatch):
    from pathlib import Path
    from backend.utils import compute_ingest as ingest, aspose_helper, source_normalizer
    source = tmp_path / 'sources'
    source.mkdir()
    template = tmp_path / 'same.xlsx'
    uploaded = source / template.name
    for file in (template, uploaded):
        file.write_bytes(b'encrypted')
    monkeypatch.setattr(ingest, 'detect_encrypted_files', lambda files:
                        [key for path, key in files if Path(path).read_bytes() == b'encrypted'])
    monkeypatch.setattr(source_normalizer, 'convert_xls_to_xlsx', lambda path: path)
    def decrypt(path, password):
        expected = 'template-password' if Path(path) == template else 'source-password'
        if password != expected:
            raise ValueError('invalid password')
        result = Path(path).with_suffix('.plain.xlsx')
        result.write_bytes(b'PK-plaintext')
        return str(result)
    monkeypatch.setattr(aspose_helper, 'decrypt_excel', decrypt)
    passwords = {'same.xlsx': 'source-password'}
    assert ingest.prepare_source_dir(str(source), passwords, str(template))[0] == ['template::same.xlsx']
    assert template.read_bytes() == b'encrypted'
    passwords['template::same.xlsx'] = 'wrong'
    assert ingest.prepare_source_dir(str(source), passwords, str(template))[0] == ['template::same.xlsx']
    assert template.read_bytes() == b'encrypted'
    passwords['template::same.xlsx'] = 'template-password'
    assert ingest.prepare_source_dir(str(source), passwords, str(template)) == ([], str(template))
    assert template.read_bytes() == uploaded.read_bytes() == b'PK-plaintext'
    assert not list(tmp_path.rglob('*.plain.xlsx'))


def test_compute_provider_reads_current_config_over_stale_environment(tmp_path, monkeypatch):
    import ast
    import os
    from pathlib import Path
    source = Path(__file__).parent / 'app' / 'main.py'
    node = next(n for n in ast.parse(source.read_text(encoding='utf-8')).body
                if isinstance(n, ast.FunctionDef) and n.name == '_resolve_compute_ai_provider')
    node.returns = None
    (tmp_path / '.env').write_text('AI_PROVIDER=deepseek\n', encoding='utf-8')
    monkeypatch.setenv('AI_PROVIDER', 'claude')
    ns = {'Path': Path, 'os': os, '__file__': str(tmp_path / 'backend' / 'app' / 'main.py'),
          '_resolve_enabled_ai_provider': lambda name: name}
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(source), 'exec'), ns)
    assert ns['_resolve_compute_ai_provider']() == 'deepseek'
    ns['_resolve_enabled_ai_provider'] = lambda name: None
    assert ns['_resolve_compute_ai_provider']() is None


def test_two_stage_ai_mapping_uses_file_sheet_then_columns(monkeypatch):
    from backend.ai_engine import ai_provider
    from backend.utils.ai_source_mapping import match_sources_with_ai

    training, actual = schemas()
    calls = []
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())

    def reply(provider, messages, **kwargs):
        calls.append(messages[0]['content'])
        if len(calls) == 1:
            return json.dumps({'mappings': [{
                'training_id': 0,
                'actual_id': 0,
                'actual_sheet': 'Payroll',
                'file_confidence': 0.95,
                'sheet_confidence': 0.93,
                'file_reason': '两份文件均为工资明细',
                'sheet_reason': 'Payroll 对应工资表',
                'code_evidence': 'fill_template 使用 Pay 列',
            }]}, ensure_ascii=False)
        return json.dumps({'mappings': [{
            'training_id': 0,
            'actual_id': 0,
            'columns': {'Employee ID': '工号', 'Pay': '金额'},
            'column_confidence': {'Employee ID': 0.99, 'Pay': 0.92},
            'column_reasons': {'Employee ID': '工号关联键', 'Pay': '工资金额'},
        }]}, ensure_ascii=False)

    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, 'claude')
    assert len(calls) == 2
    assert '文件到文件' in calls[0] and '本阶段不要处理列映射' in calls[0]
    assert 'column_confidence' not in calls[0]
    assert 'actual_columns' in calls[1] and 'fixed_mapping' in calls[1]
    mapping = result['mapping']['file_mapping']
    assert mapping['new.xlsx']['sheet_mapping'] == {'Payroll': '工资'}
    assert mapping['new.xlsx']['header_mapping'] == {'Employee ID': '工号', 'Pay': '金额'}
    assert mapping['new.xlsx']['header_confidence_by_sheet']['Payroll']['Pay'] == 0.92


def test_two_stage_keeps_file_sheet_when_column_stage_fails(monkeypatch):
    from backend.ai_engine import ai_provider
    from backend.utils.ai_source_mapping import match_sources_with_ai

    training, actual = schemas()
    calls = []
    monkeypatch.setattr(ai_provider.AIProviderFactory, 'create_provider', lambda _: object())

    def reply(provider, messages, **kwargs):
        calls.append(messages[0]['content'])
        if len(calls) == 1:
            return json.dumps({'mappings': [{
                'training_id': 0,
                'actual_id': 0,
                'actual_sheet': 'Payroll',
                'file_confidence': 0.95,
                'sheet_confidence': 0.93,
                'file_reason': '文件角色一致',
                'sheet_reason': 'Sheet 结构一致',
                'code_evidence': '未发现直接代码依据',
            }]}, ensure_ascii=False)
        return '{}'

    monkeypatch.setattr(ai_provider, 'chat_with_timeout', reply)
    result = match_sources_with_ai(FastHeaderMatcher(), training, actual, 'claude')
    mapping = result['mapping']['file_mapping']['new.xlsx']
    assert mapping['sheet_mapping'] == {'Payroll': '工资'}
    assert mapping['header_mapping'] == {}
    assert mapping['header_mapping_by_sheet'] == {'Payroll': {}}
