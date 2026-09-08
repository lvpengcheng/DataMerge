import ast
import json
import logging
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from backend.utils.training_assets import asset_manifest, resolve_asset_config, stage_revision, validate_settings
from backend.utils import training_validation  # Load real runner binding before per-test worker mocks.


def file(path, data=b'workbook'):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def test_history_recovers_persisted_files_after_temp_cleanup(tmp_path):
    source = file(tmp_path / 'source' / '员工.XLSX')
    target = file(tmp_path / '目标.xlsx')
    cfg = {'source_dir': r'C:\Temp\deleted\source', 'expected_file': r'C:\Temp\deleted\目标.xlsx'}
    manifest = asset_manifest(cfg, tmp_path)
    assert manifest['source_file_names'] == [source.name]
    assert manifest['expected_file_name'] == target.name
    assert resolve_asset_config(cfg, tmp_path)['expected_file'] == str(target)


def test_moved_revision_never_falls_back_to_old_source(tmp_path):
    file(tmp_path / 'source' / 'old.xlsx')
    current = file(tmp_path / 'revisions' / 'r1' / 'source' / 'new.xlsx')
    cfg = {'source_dir': r'E:\old\training_chat\1\revisions\r1\source'}
    assert asset_manifest(cfg, tmp_path)['source_file_names'] == [current.name]


@pytest.mark.parametrize('mode,expected', [('merge', ['a.xlsx', 'b.xlsx']), ('replace', ['a.xlsx'])])
def test_source_changes_are_copy_on_write_and_preserve_unselected_files(tmp_path, mode, expected):
    old = tmp_path / 'session' / 'source'
    file(old / 'a.xlsx', b'old a')
    file(old / 'b.xlsx', b'old b')
    uploaded = file(tmp_path / 'upload' / 'a.xlsx', b'new a')
    cfg, revision, changed = stage_revision({'source_dir': str(old)}, tmp_path / 'session',
        {'source': [uploaded]}, {'ai_provider': 'claude'}, mode)
    assert changed and revision
    assert sorted(p.name for p in Path(cfg['source_dir']).iterdir()) == expected
    assert (old / 'a.xlsx').read_bytes() == b'old a'
    assert (Path(cfg['source_dir']) / 'a.xlsx').read_bytes() == b'new a'


def test_provider_switch_does_not_copy_or_invalidate_inputs(tmp_path):
    cfg, revision, changed = stage_revision({'ai_provider': 'deepseek'}, tmp_path, {}, {'ai_provider': 'claude'})
    assert cfg['ai_provider'] == 'claude' and revision is None and not changed
    assert not list(tmp_path.iterdir())


@pytest.mark.parametrize('settings', [{'ai_provider': 'bad'}, {'salary_month': 13},
    {'manual_headers': []}, {'monthly_standard_hours': float('nan')}, {'source_dir': 'outside'}])
def test_settings_validation_rejects_invalid_or_private_fields(settings):
    with pytest.raises(ValueError):
        validate_settings(settings)


def api_function(name, namespace):
    path = Path(__file__).parent / 'api' / 'training_chat.py'
    node = next(n for n in ast.parse(path.read_text(encoding='utf-8')).body
                if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)) and n.name == name)
    node.decorator_list = []
    for argument in node.args.args:
        argument.annotation = None
    node.args.defaults = [ast.Constant(None) for _ in node.args.defaults]
    node.returns = None
    namespace = {'__package__': 'backend.api', 'Path': Path, 'logger': logging.getLogger('assets-test'), **namespace}
    exec(compile(ast.fix_missing_locations(ast.Module(body=[node], type_ignores=[])), str(path), 'exec'), namespace)
    return namespace[name]


@pytest.mark.parametrize('fail', [False, True])
def test_session_update_commits_valid_revision_only(tmp_path, monkeypatch, fail):
    import backend.utils.subprocess_runner as runner
    old = file(tmp_path / 'source' / 'old.xlsx')
    uploaded = file(tmp_path / 'upload' / 'new.xlsx')
    original = {'source_dir': str(old.parent), 'ai_provider': 'deepseek', 'latest_detailed_diff': 'old difference'}
    session = SimpleNamespace(id=1, config=dict(original), mode='formula')
    db = MagicMock()
    db.query.return_value.filter_by.return_value.first.return_value = session
    def prepare(target, args, **kw):
        return SimpleNamespace(success=not fail, error='invalid workbook',
            result={'config': args[0], 'source_structure': {'files': {}}})
    monkeypatch.setattr(runner, 'run_in_subprocess', prepare)
    update = api_function('_apply_session_asset_update', {'SessionLocal': lambda: db,
        'TrainingSession': object(), 'flag_modified': lambda *a: None,
        '_session_asset_root': lambda s: tmp_path, '_add_message': MagicMock()})
    if fail:
        with pytest.raises(ValueError, match='invalid workbook'):
            update(1, {'source': [uploaded]}, {'ai_provider': 'claude'}, 'merge', {})
        assert session.config == original
        db.commit.assert_not_called()
    else:
        result = update(1, {'source': [uploaded]}, {'ai_provider': 'claude'}, 'merge', {})
        assert result['source_file_names'] == ['new.xlsx', 'old.xlsx']
        assert session.config['validation_stale'] and session.ai_provider == 'claude'
        assert 'latest_detailed_diff' not in session.config
        db.commit.assert_called_once()
    assert old.read_bytes() == b'workbook'


def test_download_uses_recovered_path_and_checks_tenant(tmp_path):
    import os
    from fastapi import HTTPException
    expected = file(tmp_path / '目标.xlsx')
    session = SimpleNamespace(id=1, tenant_id='allowed', config={'expected_file': r'C:\gone\目标.xlsx'})
    db = MagicMock()
    db.query.return_value.filter_by.return_value.first.return_value = session
    download = api_function('download_original_file', {'os': os, 'HTTPException': HTTPException,
        'TrainingSession': object(), '_session_asset_root': lambda s: tmp_path, '_NO_STORE_HEADERS': {}})
    result = download(1, 'expected', None, db, None, ['allowed'])
    assert Path(result.path) == expected
    with pytest.raises(HTTPException) as error:
        download(1, 'expected', None, db, None, ['other'])
    assert error.value.status_code == 403


def test_updated_template_and_rule_attachments_are_reparsed_with_real_workbooks(tmp_path, monkeypatch):
    import os
    import shutil
    import sys
    import openpyxl
    from backend.utils.training_assets import prepare_revision
    namespace = {'os': os, 'shutil': shutil}
    helpers = {name: api_function(name, namespace) for name in (
        '_prepare_training_uploads_subprocess', '_build_source_structure_from_dir_impl', '_analyze_expected_structure_impl')}
    monkeypatch.setitem(sys.modules, 'backend.api.training_chat', SimpleNamespace(**helpers))
    monkeypatch.setenv('_IN_SUBPROCESS_WORKER', '1')
    source = tmp_path / 'source'
    source.mkdir()
    book = openpyxl.Workbook()
    book.active.title = '员工'
    book.active.append(['工号', '金额'])
    book.active.append(['001', 10])
    book.save(source / '员工.xlsx')
    target = tmp_path / 'new_target.xlsx'
    book.active.title = '结果'
    book.active['B2'] = '=10*2'
    book.save(target)
    book.close()
    rules = file(tmp_path / 'rules.md', '新的规则条款'.encode())
    cfg, revision, changed = stage_revision({'source_dir': str(source), 'mode': 'template',
        'rules_content': '原始补充说明', 'target_sheets': ['旧表']}, tmp_path / 'session',
        {'expected': [target], 'rules': [rules]}, {})
    result = prepare_revision(cfg, str(revision), ['expected', 'rules'])
    updated = result['config']
    assert updated['template_path'] == updated['expected_file'] != str(target)
    assert 'target_sheets' not in updated
    assert updated['expected_structure']['sheets']['结果']['formulas']['B2'] == '=10*2'
    assert '新的规则条款' in updated['rules_content'] and '原始补充说明' in updated['rules_content']
    assert result['source_structure']['total_sheets'] == 1
    rules.write_text('替换后的规则条款', encoding='utf-8')
    cfg2, revision2, _ = stage_revision(updated, tmp_path / 'session', {'rules': [rules]}, {})
    result2 = prepare_revision(cfg2, str(revision2), ['rules'])['config']
    assert '替换后的规则条款' in result2['rules_content'] and '新的规则条款' not in result2['rules_content']
    assert Path(updated['rule_files_dir'], 'rules.md').read_text(encoding='utf-8') == '新的规则条款'
