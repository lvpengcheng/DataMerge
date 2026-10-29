"""Conversion/import activates the exact package and reuses same-name scripts."""
import asyncio
import io
import os
from pathlib import Path
from types import SimpleNamespace

import pytest

os.environ['DATABASE_URL'] = 'sqlite://'
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from fastapi import UploadFile
from backend.database.connection import Base
from backend.database.models import CalculationSolution, Script, TrainingSession, User
from backend.solutions import api, store, restore, execution, package


@pytest.fixture
def context(tmp_path, monkeypatch):
    for module in (api, store, restore, execution):
        monkeypatch.setattr(module, 'ROOT', tmp_path)
    for module in (api, store):
        monkeypatch.setattr(module, 'STORE', tmp_path / 'solution_store')
    monkeypatch.setattr(store, 'fingerprint', lambda root: {})
    engine = create_engine('sqlite://')
    Base.metadata.create_all(engine)
    with Session(engine) as db:
        db.add(User(id=1, username='test', password_hash='unused'))
        session = TrainingSession(tenant_id='训练', user_id=1, session_key='sample', status='completed')
        db.add(session); db.flush()
        folder = tmp_path / 'tenants/训练/training_chat' / str(session.id)
        folder.mkdir(parents=True)
        (folder / 'expected.xlsx').write_text('fixture')
        session.config = {'expected_file': str(folder / 'expected.xlsx')}
        script = Script(tenant_id='训练', name='工资', code='def main():\r\n    return 42\r\n', source_session_id=session.id)
        db.add(script); db.flush()
        session.final_script_id = script.id
        db.commit()
        for filename in ['excel_parser.py', 'aspose_init.py', 'split_by_banner.py', 'backend/solutions/worker.py']:
            path = tmp_path / filename
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text('# fixture')
        admin = SimpleNamespace(id=1, org_id=None, role=SimpleNamespace(name='admin'))
        yield db, script, admin, tmp_path
    engine.dispose()


def test_conversion_is_active_and_freezes_code(context):
    db, script, admin, root = context
    original = script.code
    result = api.convert({'script_id': script.id}, db, admin, ['训练'])
    assert result['is_active'] and result['status'] == 'published'
    artifact = store.artifact(result['digest'])
    script.code = 'def main(): return "later edit"'
    db.commit()
    assert (artifact / 'business/script.py').read_bytes().decode() == original
    package.verify(artifact)


def test_visible_multitarget_script_freezes_each_template_for_execution(context):
    import hashlib
    from backend.utils.multi_workbook import build_script, script_targets, template_reference_code
    db, script, admin, root = context
    targets = {}
    for index, name in enumerate(['明细.xlsx', '汇总.xlsx']):
        template = root / 'tenants/训练' / name
        template.write_text('template ' + str(index))
        code = (f'TEMPLATE_NAME = {name!r}\nTEMPLATE_PATH = {str(template)!r}\n'
                'def main(): pass\n')
        targets[name] = {'code': code, 'structure': {}}
    script.code = build_script(targets)
    db.commit()
    row = store.build(db, script, admin.id)
    artifact = store.artifact(row.digest)
    manifest = package.verify(artifact)
    for target in script_targets(script.code).values():
        code = target['code']
        full_key = hashlib.sha256(code.encode()).hexdigest()
        ref_key = hashlib.sha256(template_reference_code(code).encode()).hexdigest()
        assert manifest['templates'][full_key] == manifest['templates'][ref_key]
        assert (artifact / manifest['templates'][ref_key]).is_file()
        from backend.utils.template_resolver import extract_template_ref
        assert Path(manifest['templates'][ref_key]).name == extract_template_ref(code)[0]


def test_frozen_template_basename_preserves_business_owner_selection(context, monkeypatch):
    import hashlib
    from backend.utils.multi_workbook import build_script
    from backend.utils import template_resolver
    db, script, admin, root = context
    template = root / 'tenants/训练/本月工资-王五.xlsx'
    template.write_bytes(b'template')
    child = (f'from pathlib import Path\nTEMPLATE_NAME = {template.name!r}\n'
             f'TEMPLATE_PATH = {str(template)!r}\n'
             'def main():\n'
             '    owner = "王五" if "王五" in Path(_template_override_path).stem else "全部人员"\n'
             '    Path(output_folder, "result.xlsx").write_text(owner, encoding="utf-8")\n')
    script.code = build_script({'王五.xlsx': {'code': child, 'structure': {}}})
    db.commit()
    row = store.build(db, script, admin.id)
    artifact = store.artifact(row.digest)
    manifest = package.verify(artifact)
    def frozen_resolve(**kwargs):
        key = hashlib.sha256(kwargs['script_code'].encode()).hexdigest()
        relative = manifest['templates'].get(key)
        return str(artifact / relative) if relative else None
    monkeypatch.setattr(template_resolver, 'resolve_template_path', frozen_resolve)
    output = root / 'outputs'
    output.mkdir()
    env = {'output_folder': str(output), '__name__': 'compute_script', 'tenant_id': '训练'}
    exec(compile(script.code, 'script.py', 'exec'), env)
    env['main']()
    assert (output / '王五.xlsx').read_text(encoding='utf-8') == '王五'


def test_import_replaces_same_name_and_deactivates_old_binding(context):
    db, source, admin, root = context
    row = store.build(db, source, admin.id)
    archive = root / '方案.zip'
    package.export_zip(store.artifact(row.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)
    target = Script(tenant_id='正式', name='工资', code='old')
    other = Script(tenant_id='其他', name='工资', code='other tenant')
    db.add_all([target, other]); db.commit()

    def imported():
        return asyncio.run(api.import_package(
            UploadFile(filename='方案.zip', file=io.BytesIO(archive.read_bytes())),
            '正式', False, True, db, admin, ['正式'], script_name='工资'))

    first = imported()
    assert first['script_id'] == target.id
    assert first['is_active'] and target.is_active
    assert target.code == source.code
    count = db.query(TrainingSession).count()
    assert imported()['id'] == first['id']
    assert db.query(TrainingSession).count() == count  # CRLF code is also idempotent.
    source.code = 'def main(): return 99\n'; db.commit()
    newer = store.build(db, source, admin.id)
    package.export_zip(store.artifact(newer.digest), archive)
    second = imported()
    assert second['script_id'] == target.id
    assert second['id'] != first['id'] and second['is_active']
    assert not db.get(CalculationSolution, first['id']).is_active
    assert db.query(Script).filter_by(tenant_id='正式', name='工资').count() == 1
    assert other.code == 'other tenant'
    assert target.code == source.code
    package.verify(store.artifact(first['digest']))


def test_upgrade_activates_latest_pending_once_and_preserves_disabled_versions(context):
    from datetime import datetime, timedelta
    from backend.solutions.migration import activate_pending_solutions
    db, source, admin, root = context
    old = store.build(db, source, admin.id)
    old.created_at = datetime(2026, 1, 1)
    old.status, old.is_active = 'published', True
    source.code = 'def main(): return 99'; db.commit()
    pending = store.build(db, source, admin.id)
    pending.created_at = datetime(2026, 1, 2)
    disabled_script = Script(tenant_id='训练', name='停用方案', code='code', is_active=False)
    db.add(disabled_script); db.flush()
    disabled = CalculationSolution(id='disabled', tenant_id='训练', script_id=disabled_script.id,
                                   name='停用方案', digest=old.digest, status='published', is_active=False)
    missing = CalculationSolution(id='missing', tenant_id='训练', script_id=source.id,
                                  name='缺失方案', digest='a' * 64, status='candidate', is_active=False)
    db.add_all([disabled, missing]); db.commit()
    activate_pending_solutions(db.get_bind())
    db.expire_all()
    assert pending.is_active and pending.status == 'published'
    assert not old.is_active
    assert not disabled.is_active and not missing.is_active
    activate_pending_solutions(db.get_bind())
    db.expire_all()
    assert pending.is_active and not old.is_active
