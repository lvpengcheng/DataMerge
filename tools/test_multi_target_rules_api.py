"""Exercise the real rule endpoints without starting unrelated app services."""
import ast
import asyncio
import json
import io
import logging
import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Dict, List, Optional

import pytest
from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.testclient import TestClient

from backend.auth.dependencies import get_current_user
from backend.database import models as db_models
from backend.api.rules import router as session_router
from backend.tests.test_table_split_api import api
from tools.test_multi_target_rules import Provider, organizer, bundle
from backend.utils.multi_target_rules import parse_bundle, render_bundle


@pytest.fixture
def rules_api(api, monkeypatch):
    from backend.ai_engine.rule_organizer import RuleOrganizer
    from backend.ai_engine.ai_provider import AIProviderFactory
    from backend.utils import aspose_helper, banner_splitter
    _, state, sessions = api
    provider = Provider()
    reference = organizer(provider)
    monkeypatch.setattr(RuleOrganizer, '__init__', lambda self, ai_provider: setattr(self, 'ai_provider', ai_provider))
    monkeypatch.setattr(RuleOrganizer, '_extract_source_structures', lambda self, *args: 'source.xlsx：工号、金额')
    monkeypatch.setattr(RuleOrganizer, '_extract_design_docs', lambda self, *args: '设计说明')
    def target(self, path, *args):
        self._last_target_inventory = [{'name': '明细', 'fields': [{'name': '工号', 'column': 'A'}, {'name': '金额', 'column': 'B'}]}]
        self._last_target_parse_error = False
        return f'{Path(path).name}，明细：工号、金额'
    monkeypatch.setattr(RuleOrganizer, '_extract_target_structure', target)
    monkeypatch.setattr(AIProviderFactory, 'create_provider', lambda *args: provider)
    monkeypatch.setattr(aspose_helper, 'is_encrypted', lambda *args: False)
    monkeypatch.setattr(banner_splitter, 'preprocess_uploaded_files', lambda *args: None)
    app = FastAPI()
    app.include_router(session_router)
    app.dependency_overrides[get_current_user] = lambda: state.user
    namespace = dict(globals(), app=app, SessionLocal=sessions, logger=logging.getLogger('rules-test'),
                     AIProviderFactory=AIProviderFactory, __package__='backend.app')
    source = ast.parse(Path('backend/app/main.py').read_text(encoding='utf-8'))
    functions = [node for node in source.body if isinstance(node, ast.AsyncFunctionDef) and node.name in {
        'organize_rules_endpoint', 'organize_rules_stream_endpoint', 'rules_chat_endpoint'}]
    # Keep the actual chat endpoint name discovered from its route decorator.
    for node in source.body:
        if isinstance(node, ast.AsyncFunctionDef) and any(
                isinstance(d, ast.Call) and d.args and isinstance(d.args[0], ast.Constant) and d.args[0].value == '/api/rules/chat'
                for d in node.decorator_list) and node not in functions:
            functions.append(node)
    exec(compile(ast.Module(body=functions, type_ignores=[]), '<real-rule-endpoints>', 'exec'), namespace)
    with TestClient(app) as client:
        yield client, provider, sessions, state


def attachments(legacy=False):
    return [('source_files', ('source.xlsx', b'source', 'application/octet-stream')),
        ('target_file' if legacy else 'target_files', ('工资.xlsx', b'payroll', 'application/octet-stream'))] + ([] if legacy else [
        ('target_files', ('银行.xlsx', b'bank', 'application/octet-stream'))])


def events(response):
    return [json.loads(line[6:]) for line in response.text.splitlines() if line.startswith('data: ')]


def test_sync_and_stream_accept_multiple_targets_and_persist_full_rules(rules_api):
    client, _, sessions, _ = rules_api
    result = client.post('/api/rules/organize', files=attachments())
    assert result.status_code == 200, result.text
    assert result.json()['target_file_names'] == ['工资.xlsx', '银行.xlsx']
    assert len(parse_bundle(result.json()['content'])['targets']) == 2
    streamed = client.post('/api/rules/organize/stream', files=attachments())
    output = events(streamed)
    assert not [event for event in output if event['type'] == 'error'], output
    assert sum(event['type'] == 'thinking' and '模型返回的思考片段' in event.get('content', '')
               for event in output) == 4
    complete = next(event for event in output if event['type'] == 'complete')
    with sessions() as db:
        stored = db.get(db_models.RuleSession, complete['session_id'])
        assert stored.target_file_names == ['工资.xlsx', '银行.xlsx']
        assert stored.rules_bundle['targets'][1]['file_name'] == '银行.xlsx'
        assert stored.final_result == complete['content']
    detail = client.get(f'/api/rules/sessions/{complete["session_id"]}').json()
    assert detail['target_file_names'] == ['工资.xlsx', '银行.xlsx']
    assert detail['rules_bundle']['version'] == 1


def test_multitarget_failure_does_not_publish_complete_or_replace_existing_session(rules_api):
    client, provider, sessions, _ = rules_api
    with sessions() as db:
        saved = db_models.RuleSession(user_id=1, title='已有规则', messages=[], final_result='原规则')
        db.add(saved); db.commit(); sid = saved.id
    provider.missing = True
    response = client.post('/api/rules/organize/stream', files=attachments(), data={'session_id': str(sid)})
    output = events(response)
    assert any(event['type'] == 'error' for event in output)
    assert not any(event['type'] == 'complete' for event in output)
    with sessions() as db:
        assert db.get(db_models.RuleSession, sid).final_result == '原规则'


def test_targeted_followup_uses_saved_revision_and_preserves_other_rules(rules_api):
    client, provider, sessions, _ = rules_api
    data = bundle()
    with sessions() as db:
        saved = db_models.RuleSession(user_id=1, title='多个目标', messages=[], rules_bundle=data, final_result=render_bundle(data))
        db.add(saved); db.commit(); sid = saved.id
    updated = {**data['targets'][0], 'rules': data['targets'][0]['rules']+'\n调整精度'}
    provider.patch = {'reply': '已修改', 'targets': [updated]}
    result = client.post('/api/rules/chat', json={'session_id': sid, 'messages': [
        {'role': 'assistant', 'content': '过时的页面内容'}, {'role': 'user', 'content': '只修改工资.xlsx精度'}]})
    output = events(result)
    assert not any(e['type'] == 'error' for e in output), output
    assert any(e['type'] == 'thinking' and '模型返回的思考片段' in e.get('content', '') for e in output)
    complete = next(e for e in output if e['type'] == 'complete')
    assert complete['is_rules_document'] is True
    with sessions() as db:
        saved = db.get(db_models.RuleSession, sid)
        assert saved.rules_bundle['targets'][1] == data['targets'][1]
        previous = saved.final_result
    provider.patch = {'reply': '请说明哪个目标', 'targets': []}
    result = client.post('/api/rules/chat', json={'session_id': sid, 'messages': [{'role': 'user', 'content': '改一下'}]})
    complete = next(e for e in events(result) if e['type'] == 'complete')
    assert complete['is_rules_document'] is False
    with sessions() as db:
        assert db.get(db_models.RuleSession, sid).final_result == previous


def test_duplicate_targets_and_other_users_sessions_are_rejected(rules_api):
    client, _, sessions, _ = rules_api
    duplicate = attachments()+[('target_files', ('工资.xls', b'other', 'application/octet-stream'))]
    assert client.post('/api/rules/organize', files=duplicate).status_code == 400
    with sessions() as db:
        saved = db_models.RuleSession(user_id=2, title='其他人的规则', messages=[], final_result='PRIVATE')
        db.add(saved); db.commit(); sid = saved.id
    assert client.post('/api/rules/chat', json={'session_id': sid, 'messages': [{'role': 'user', 'content': '修改'}]}).status_code == 404
    assert client.post('/api/rules/organize/stream', files=attachments(), data={'session_id': str(sid)}).status_code == 404


def test_legacy_single_target_field_and_history_remain_compatible(rules_api):
    client, _, sessions, _ = rules_api
    response = client.post('/api/rules/organize', files=attachments(legacy=True))
    assert response.status_code == 200
    assert response.json()['target_file_names'] == ['工资.xlsx']
    assert parse_bundle(response.json()['content']) is None
    with sessions() as db:
        old = db_models.RuleSession(user_id=1, title='历史单目标', target_file_name='旧目标.xlsx', messages=[], final_result='旧规则')
        db.add(old); db.commit(); sid = old.id
    assert client.get(f'/api/rules/sessions/{sid}').json()['target_file_names'] == ['旧目标.xlsx']


def test_every_target_password_is_checked_and_scoped_by_category(rules_api, monkeypatch):
    from backend.utils import aspose_helper
    client, _, _, _ = rules_api
    calls = []
    monkeypatch.setattr(aspose_helper, 'is_encrypted', lambda path: Path(path).parent.name == 'target')
    def decrypt(path, password):
        calls.append((Path(path).name, password))
        decrypted = path + '.decrypted'
        Path(decrypted).write_bytes(b'decrypted')
        return decrypted
    monkeypatch.setattr(aspose_helper, 'decrypt_excel', decrypt)
    assert client.post('/api/rules/organize/stream', files=attachments()).status_code == 422
    response = client.post('/api/rules/organize/stream', files=attachments(), data={
        'file_passwords': json.dumps({'target:工资.xlsx': 'p1', 'target:银行.xlsx': 'p2'})})
    assert any(event['type'] == 'complete' for event in events(response)), response.text
    assert calls == [('工资.xlsx', 'p1'), ('银行.xlsx', 'p2')]


def test_same_named_source_and_target_are_stored_separately(tmp_path):
    from backend.utils.rule_uploads import save_rule_attachments
    result = asyncio.run(save_rule_attachments(tmp_path,
        [UploadFile(filename='同名.xlsx', file=io.BytesIO(b'source'))],
        [UploadFile(filename='同名.xlsx', file=io.BytesIO(b'target'))], []))
    assert Path(result['source_paths'][0]).read_bytes() == b'source'
    assert Path(result['target_paths'][0]).read_bytes() == b'target'
