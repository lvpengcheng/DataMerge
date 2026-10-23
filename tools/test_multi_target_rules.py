"""Multi-target organization, portable scope boundaries and 200k rule limits."""
import ast
import copy
import io
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from backend.ai_engine.rule_organizer import RuleOrganizer
from backend.utils.multi_target_rules import (
    RULES_CONTENT_LIMIT, limit_rules_content, render_bundle, parse_bundle, rules_for_workbook,
    patch_bundle, rules_for_structure,
)
from backend.utils.multi_workbook import generate_workbooks, correct_workbooks


def bundle():
    targets = []
    for index, name in enumerate(['工资.xlsx', '银行.xlsx'], 1):
        tid = f'T{index:04}'
        fields = [{'id': f'{tid}.S0001.C{i:04}', 'name': col, 'column': chr(64+i)}
                  for i, col in enumerate(['工号', '金额'], 1)]
        targets.append({'id': tid, 'file_name': name, 'sheets': [{'id': tid+'.S0001', 'name': '明细', 'fields': fields}],
            'rules': f'## 基本信息\n目标{name}\n## 列级规则\n明细：工号、金额\nPRIVATE_{index}',
            'covered_field_ids': [f['id'] for f in fields], 'external_output_dependencies': []})
    return {'version': 1, 'common_rules': '公共规则：保留文本主键', 'targets': targets,
            'cross_target_checks': '仅在人员范围一致时核对合计', 'source_format': '源文件.xlsx / 明细：工号、金额'}


def test_portable_roundtrip_and_same_named_sheets_remain_separate():
    original = bundle()
    text = render_bundle(original)
    parsed = parse_bundle(text)
    assert parsed['targets'] == original['targets']
    for index, target in enumerate(original['targets'], 1):
        scoped = rules_for_workbook(text, target['file_name'])
        assert f'PRIVATE_{index}' in scoped and f'PRIVATE_{3-index}' not in scoped
        assert original['common_rules'] in scoped and original['source_format'] in scoped
    assert rules_for_workbook('旧版规则', 'anything.xlsx') == '旧版规则'
    assert 'PRIVATE_2' in rules_for_structure(text, {'file_name': '银行.xls'})
    with pytest.raises(ValueError, match='对应的目标'):
        rules_for_workbook(text, '不明目标.xlsx')


def test_additional_rule_attachments_are_preserved_without_leaking_other_targets():
    text = '用户新增规则：金额保留三位\n' + render_bundle(bundle()) + '\n其他补充要求'
    scoped = rules_for_workbook(text, '工资.xlsx')
    assert '金额保留三位' in scoped and '其他补充要求' in scoped
    assert 'PRIVATE_2' not in scoped


def test_200k_boundary_retains_tail_and_does_not_silently_truncate_a_bundle():
    assert RULES_CONTENT_LIMIT == 200000
    text = 'x' * 199999 + '尾'
    assert limit_rules_content(text) == text
    assert len(limit_rules_content(text+'超出')) == 200000
    data = bundle()
    normal = render_bundle(data)
    data['common_rules'] += 'x' * (200000-len(normal))
    assert len(render_bundle(data)) == 200000
    data['common_rules'] += 'x'
    with pytest.raises(ValueError, match='20万'):
        render_bundle(data)


def test_partial_unknown_and_cyclic_rules_are_rejected():
    data = bundle()
    with pytest.raises(ValueError, match='边界'):
        parse_bundle(render_bundle(data).split('<!-- DM_RULES BUNDLE END')[0])
    with pytest.raises(ValueError, match='覆盖'):
        patch_bundle(data, {'targets': [{'id': 'T0001', 'covered_field_ids': []}]})
    with pytest.raises(ValueError, match='未知'):
        patch_bundle(data, {'targets': [{'id': 'T9999', 'rules': 'x'}]})
    data['targets'][1]['external_output_dependencies'] = ['T0001']
    text = render_bundle(data)
    with pytest.raises(ValueError, match='其他结果文件'):
        rules_for_workbook(text, '银行.xlsx')
    data['targets'][0]['external_output_dependencies'] = ['T0002']
    with pytest.raises(ValueError, match='循环依赖'):
        render_bundle(data)


class Provider:
    def __init__(self, missing=False, patch=None):
        self.calls, self.missing, self.patch = [], missing, patch

    def chat(self, messages, **kwargs):
        self.calls.append((copy.deepcopy(messages), kwargs))
        system = messages[0]['content']
        if '规则修改助手' in system:
            return json.dumps(self.patch or {'reply': '请说明要修改哪个目标。', 'targets': []}, ensure_ascii=False)
        if '本次仅整理指定的一个目标' in system:
            target = json.loads(messages[1]['content'].split('## 本次目标字段清单\n')[1].split('\n## 全部目标身份清单')[0])
            fields = [f for s in target['sheets'] for f in s['fields']]
            return json.dumps({'rules': f'## 基本信息\n{target["file_name"]}\n## 列级规则\n明细：工号、金额\n'
                + ('PRIVATE_1' if target['id'] == 'T0001' else 'PRIVATE_2'),
                'covered_field_ids': [] if self.missing else [f['id'] for f in fields],
                'external_output_dependencies': [], 'summary': '各目标按各自人群处理'}, ensure_ascii=False)
        if '公共分析' in system:
            return json.dumps({'common_rules': '保留文本主键', 'cross_target_checks': '待确认两表人员范围'}, ensure_ascii=False)
        if '你是一位专业的数据处理规则分析师' in system:
            return '# 数据处理规则\n## 基本信息\n## 列级规则\n明细：工号、金额'
        return json.dumps({'cross_target_checks': '按已确认人群核对，不能直接假定合计相等'}, ensure_ascii=False)

    def chat_stream(self, messages, chunk_callback=None, thinking_callback=None, **kwargs):
        if thinking_callback:
            thinking_callback('模型返回的思考片段\n')
        result = self.chat(messages, **kwargs)
        if chunk_callback:
            chunk_callback(result)
        return result


def organizer(provider):
    item = RuleOrganizer.__new__(RuleOrganizer)
    item.ai_provider = provider
    item._extract_source_structures = lambda *args: 'source.xlsx：工号、金额（脱敏样本）'
    item._extract_design_docs = lambda *args: '明确工资表与银行表的数据范围'
    def target(path, *args):
        item._last_target_inventory = [{'name': '明细', 'fields': [
            {'name': '工号', 'column': 'A'}, {'name': '金额', 'column': 'B'}]}]
        item._last_target_parse_error = False
        return f'{Path(path).name}，明细：工号、金额，金额保留2位'
    item._extract_target_structure = target
    return item


def test_organization_checks_coverage_and_streams_only_complete_document():
    provider = Provider()
    item = organizer(provider)
    chunks, progress = [], []
    result = item.organize_rules_stream(['source.xlsx'], target_files=['工资.xlsx', '银行.xlsx'],
        chunk_callback=chunks.append, thinking_callback=progress.append)
    assert chunks == [result] and parse_bundle(result)['targets'] == item.last_rule_bundle['targets']
    assert len(provider.calls) == 4 and len(progress) >= 4
    assert progress.count('模型返回的思考片段\n') == 4
    assert all(kwargs['require_complete'] for _, kwargs in provider.calls)
    assert 'PRIVATE_2' not in rules_for_workbook(result, '工资.xlsx')
    failing = organizer(Provider(missing=True))
    chunks = []
    with pytest.raises(ValueError, match='覆盖'):
        failing.organize_rules_stream(['source.xlsx'], target_files=['工资.xlsx', '银行.xlsx'], chunk_callback=chunks.append)
    assert not chunks and failing.last_rule_bundle is None


def test_followup_and_long_sop_forward_provider_thinking_without_publishing_json(monkeypatch):
    item = organizer(Provider())
    thoughts, chunks = [], []
    result = item.chat_followup([
        {'role': 'assistant', 'content': render_bundle(bundle())},
        {'role': 'user', 'content': '需要明确什么？'}],
        chunk_callback=chunks.append, thinking_callback=thoughts.append)
    assert thoughts.count('模型返回的思考片段\n') == 1
    assert chunks == [result] and result == '请说明要修改哪个目标。'
    monkeypatch.setattr('backend.ai_engine.rule_organizer.MAX_DOC_CHARS', 1000)
    thoughts.clear()
    result = item._prepare_design_context('规则原文' * 600, progress=thoughts.append)
    assert thoughts.count('模型返回的思考片段\n') == 3
    assert '分段 3' in result


@pytest.mark.parametrize('kind', ['openai', 'claude'])
@pytest.mark.parametrize('truncated', [False, True])
def test_provider_stream_preserves_thinking_and_rejects_truncated_output(kind, truncated):
    from backend.ai_engine.ai_provider import OpenAIProvider, ClaudeProvider
    cls = OpenAIProvider if kind == 'openai' else ClaudeProvider
    provider = cls.__new__(cls)
    provider.max_tokens = 16000
    seen = []
    def stream(*args, thinking_callback=None, **kwargs):
        assert 'require_complete' not in kwargs
        thinking_callback('provider reasoning')
        yield '{"complete":true}', None
        yield '', ('length' if kind == 'openai' else 'max_tokens') if truncated else 'stop'
    setattr(provider, '_openai_chat_stream' if kind == 'openai' else '_claude_chat_stream', stream)
    def run():
        return provider.chat_stream([{'role': 'user', 'content': 'test'}],
                                    thinking_callback=seen.append, require_complete=True)
    if truncated:
        with pytest.raises(ValueError, match='AI_RESPONSE_TRUNCATED'):
            run()
    else:
        assert run() == '{"complete":true}'
    assert seen == ['provider reasoning']


def test_json_stream_retries_truncation_without_publishing_partial_document():
    class TruncatedOnce(Provider):
        attempts = 0
        def chat_stream(self, messages, thinking_callback=None, **kwargs):
            self.attempts += 1
            thinking_callback(f'attempt {self.attempts}')
            if self.attempts == 1:
                raise ValueError('AI_RESPONSE_TRUNCATED: incomplete')
            return '{"ok":true}'
    provider = TruncatedOnce()
    thoughts = []
    assert organizer(provider)._json_analysis([], thinking_callback=thoughts.append) == {'ok': True}
    assert thoughts == ['attempt 1', 'attempt 2']


def test_targeted_followup_preserves_other_targets_and_answers_do_not_replace_rules():
    original = bundle()
    update = {**original['targets'][0], 'rules': original['targets'][0]['rules']+'\n修改为三位小数'}
    provider = Provider(patch={'reply': '已修改工资表', 'targets': [update]})
    item = organizer(provider)
    result = item.chat_followup([{'role': 'assistant', 'content': render_bundle(original)},
        {'role': 'user', 'content': '只把工资.xlsx改成三位小数'}])
    changed = parse_bundle(result)
    assert changed['targets'][1] == original['targets'][1]
    assert '三位小数' in changed['targets'][0]['rules']
    provider.patch = {'reply': '请明确文件', 'targets': []}
    assert item.chat_followup([{'role': 'assistant', 'content': result}, {'role': 'user', 'content': '改一下'}]) == '请明确文件'
    assert item.last_rule_bundle is None


def test_generation_and_correction_receive_only_the_current_targets_rules():
    data = bundle()
    class Generator:
        def __init__(self): self.calls = []; self.last_prompt = ''
        def generate_code(self, **kwargs):
            self.calls.append(kwargs)
            return 'def main():\n    return {}\n', 'ok'
        def generate_correction_code(self, **kwargs):
            self.calls.append(kwargs)
            return 'def main():\n    return {}\n'
    generator = Generator()
    structure = {'workbooks': {t['file_name']: {'sheets': {'明细': {}}} for t in data['targets']}, 'sheets': {'明细': {}}}
    code, _ = generate_workbooks(generator, {'expected_structure': structure, 'rules_content': render_bundle(data)})
    assert 'PRIVATE_1' in generator.calls[0]['rules_content'] and 'PRIVATE_2' not in generator.calls[0]['rules_content']
    assert 'PRIVATE_2' in generator.calls[1]['rules_content'] and 'PRIVATE_1' not in generator.calls[1]['rules_content']
    generator.calls.clear()
    correct_workbooks(generator, {'original_code': code, 'rules_content': render_bundle(data)})
    assert 'PRIVATE_2' not in generator.calls[0]['rules_content']
    assert 'PRIVATE_1' not in generator.calls[1]['rules_content']


def test_single_target_keeps_original_markdown_and_prompt_path():
    provider = Provider()
    item = organizer(provider)
    result = item.organize_rules(['source.xlsx'], target_file='工资.xlsx')
    assert parse_bundle(result) is None and '## 列级规则' in result and '## 源数据格式' in result
    assert len(provider.calls) == 1
    chunks = []
    result = item.organize_rules_stream(['source.xlsx'], target_files=['工资.xlsx'], chunk_callback=chunks.append)
    assert ''.join(chunks) == result and parse_bundle(result) is None


def test_relevant_rules_and_source_refresh_keep_large_and_structured_rules_intact():
    from backend.ai_engine.formula_code_generator import FormulaCodeGenerator, replace_source_data_format_in_rules
    text = '金额：' + 'x' * 100000 + '尾部规则'
    generator = FormulaCodeGenerator.__new__(FormulaCodeGenerator)
    assert generator._extract_relevant_rules(text, ['金额']).endswith('尾部规则')
    refreshed = replace_source_data_format_in_rules(render_bundle(bundle()), '新的源数据结构')
    parsed = parse_bundle(refreshed)
    assert parsed['source_format'] == '新的源数据结构'
    assert parsed['targets'] == bundle()['targets']


def test_explicit_target_scope_rejects_ai_changes_to_another_target():
    original = bundle()
    provider = Provider(patch={'targets': [{**original['targets'][1],
        'rules': original['targets'][1]['rules'] + '\n不应发生的修改'}]})
    item = organizer(provider)
    with pytest.raises(ValueError, match='点名'):
        item.chat_followup([{'role': 'assistant', 'content': render_bundle(original)},
            {'role': 'user', 'content': '只修改工资.xlsx的金额精度'}])
    assert item.last_rule_bundle is None


def test_real_workbook_headers_and_schema_are_available_for_every_target(tmp_path):
    from openpyxl import Workbook
    paths = []
    for name in ['来源.xlsx', '工资.xlsx', '银行.xlsx']:
        book = Workbook()
        sheet = book.active
        sheet.title = '明细'
        sheet.append(['工号', '金额'])
        sheet.append(['0001', 100])
        sheet.append(['0002', 200])
        sheet['A2'].number_format = '@'
        sheet['B2'].number_format = '0.00'
        path = tmp_path / name
        book.save(path)
        paths.append(str(path))
    item = RuleOrganizer(Provider())
    result = item.organize_rules(paths[:1], target_files=paths[1:])
    targets = parse_bundle(result)['targets']
    assert [target['file_name'] for target in targets] == ['工资.xlsx', '银行.xlsx']
    assert all(target['sheets'][0]['name'] == '明细' for target in targets)
    assert all(target['sheets'][0]['fields'][0]['schema'] for target in targets)


def test_rule_session_migration_is_additive_and_idempotent(monkeypatch):
    from sqlalchemy import create_engine, inspect, text
    from sqlalchemy.orm import sessionmaker
    from backend.database.models import Base
    from backend.database import init_db
    engine = create_engine('sqlite://')
    Base.metadata.create_all(engine)
    with engine.begin() as connection:
        connection.exec_driver_sql('DROP TABLE rule_sessions')
        connection.exec_driver_sql('CREATE TABLE rule_sessions (id INTEGER PRIMARY KEY, target_file_name VARCHAR(200))')
        connection.exec_driver_sql("INSERT INTO rule_sessions(id,target_file_name) VALUES(1,'old.xlsx')")
    monkeypatch.setattr(init_db, 'engine', engine)
    monkeypatch.setattr(init_db, 'SessionLocal', sessionmaker(bind=engine))
    init_db._migrate_add_columns()
    init_db._migrate_add_columns()
    columns = {column['name'] for column in inspect(engine).get_columns('rule_sessions')}
    assert {'target_file_name', 'target_file_names', 'rules_bundle'} <= columns
    with engine.connect() as connection:
        assert connection.execute(text('SELECT target_file_name FROM rule_sessions WHERE id=1')).scalar() == 'old.xlsx'
    engine.dispose()
