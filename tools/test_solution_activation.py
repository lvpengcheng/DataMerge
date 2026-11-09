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


def test_import_replaces_current_solution_and_archives_old_binding(context):
    db, source, admin, root = context
    row = store.build(db, source, admin.id)
    archive = root / '方案.zip'
    package.export_zip(store.artifact(row.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)
    target = Script(tenant_id='正式', name='工资', code='old', created_by=1)
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
    existing = db.get(CalculationSolution, first['id'])
    existing.validation = {'digest': first['digest'], 'passed': True}
    db.commit()
    newer = store.build(db, source, admin.id)
    package.export_zip(store.artifact(newer.digest), archive)
    second = imported()
    assert second['script_id'] == target.id
    assert second['id'] == first['id'] and second['is_active']
    assert second['digest'] == newer.digest and second['digest'] != first['digest']
    assert second['validation'] is None
    assert db.query(CalculationSolution).filter_by(tenant_id='正式', name='工资', is_active=True).count() == 1
    previous = db.get(CalculationSolution, first['id'])
    assert previous.is_active and previous.digest == second['digest']
    assert previous.validation is None
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


def test_same_name_import_overwrites_even_when_script_binding_is_missing(context):
    db, source, admin, root = context
    built = store.build(db, source, admin.id)
    archive = root / 'migration.zip'
    package.export_zip(store.artifact(built.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)

    def imported():
        return asyncio.run(api.import_package(
            UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
            '正式', False, True, db, admin, ['正式'], script_name='工资'))

    first = imported()
    old_session = first['session_id']
    old_script_id = first['script_id']
    db.delete(db.get(Script, old_script_id)); db.commit()
    other = Script(id=old_script_id, tenant_id='其他', name='工资', code='other tenant')
    db.add(other); db.commit()
    second = imported()
    assert second['id'] == first['id']
    assert second['script_id'] != old_script_id
    assert db.query(CalculationSolution).filter_by(tenant_id='正式', name='工资').count() == 1
    assert db.get(TrainingSession, old_session)
    assert other.code == 'other tenant'
    assert db.get(CalculationSolution, built.id).digest == built.digest


def test_failed_overwrite_keeps_existing_solution_and_script(context, monkeypatch):
    from fastapi import HTTPException
    db, source, admin, root = context
    original = store.build(db, source, admin.id)
    archive = root / 'migration.zip'
    package.export_zip(store.artifact(original.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)

    def imported():
        return asyncio.run(api.import_package(
            UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
            '正式', False, True, db, admin, ['正式'], script_name='工资'))

    first = imported()
    previous_code = db.get(Script, first['script_id']).code
    source.code = 'def main(): return 123\n'; db.commit()
    newer = store.build(db, source, admin.id)
    package.export_zip(store.artifact(newer.digest), archive)
    real_restore = restore.restore

    def fail_restore(*args, **kwargs):
        real_restore(*args, **kwargs)
        raise ValueError('simulated restore failure')

    monkeypatch.setattr(restore, 'restore', fail_restore)
    with pytest.raises(HTTPException, match='simulated restore failure'):
        imported()
    current = db.get(CalculationSolution, first['id'])
    assert current.digest == first['digest'] and current.is_active
    assert db.get(Script, current.script_id).code == previous_code
    assert db.query(CalculationSolution).filter_by(tenant_id='正式', name='工资').count() == 1


def test_import_can_target_missing_script_via_loaded_solution(context):
    from fastapi import HTTPException
    db, source, admin, root = context
    built = store.build(db, source, admin.id)
    archive = root / 'migration.zip'
    package.export_zip(store.artifact(built.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)
    first = asyncio.run(api.import_package(
        UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
        '正式', False, True, db, admin, ['正式'], script_name='目标工资方案'))
    db.delete(db.get(Script, first['script_id'])); db.commit()
    fallback = api.listing('正式', db, admin, ['正式'])['scripts'][0]
    assert fallback['needs_restore']
    result = asyncio.run(api.import_package(
        UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
        '正式', False, True, db, admin, ['正式'], script_id=fallback['id']))
    assert result['id'] == first['id']
    assert result['name'] == '目标工资方案'  # Preserve selected name, not the package's source name.
    assert db.get(Script, result['script_id']).name == '目标工资方案'
    assert db.query(CalculationSolution).filter_by(tenant_id='正式').count() == 1
    with pytest.raises(HTTPException) as error:
        asyncio.run(api.import_package(
            UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
            '训练', False, True, db, admin, ['训练'], script_id=fallback['id']))
    assert error.value.status_code == 404


def test_missing_migrated_script_is_selectable_and_recoverable_without_replaying_source_ids(context):
    db,source,admin,root=context
    row=store.build(db,source,admin.id)
    archive=root/'migration.zip';package.export_zip(store.artifact(row.digest),archive)
    (root/'tenants/正式').mkdir(parents=True)
    imported=asyncio.run(api.import_package(UploadFile(filename='migration.zip',file=io.BytesIO(archive.read_bytes())),
        '正式',False,True,db,admin,['正式'],script_name='迁移脚本'))
    target=db.get(Script,imported['script_id']);old_id=target.id
    db.delete(target);db.commit()
    # A reused id in another tenant must never be rebound or overwritten.
    other=Script(id=old_id,tenant_id='其他',name='其他脚本',code='unrelated')
    db.add(other);db.commit()
    listing=api.listing('正式',db,admin,['正式'])
    fallback=listing['scripts'][0]
    assert fallback['needs_restore'] and fallback['id']=='solution_'+imported['id']
    restored=api.convert({'script_id':fallback['id'],'tenant_id':'正式'},db,admin,['正式'])
    current=db.get(Script,restored['script_id'])
    assert current.tenant_id=='正式' and current.code==source.code and current.is_active
    assert other.code=='unrelated'
    assert db.get(TrainingSession,current.source_session_id)
    assert len(api.listing('正式',db,admin,['正式'])['scripts'])==1
    package.verify(store.artifact(imported['digest']))


def test_training_entry_restores_missing_script_even_when_old_session_still_exists(context):
    db,source,admin,root=context
    row=store.build(db,source,admin.id)
    old_session=source.source_session_id
    db.delete(source);db.commit()
    assert db.get(TrainingSession,old_session)
    result=asyncio.run(api.training_workspace(row.id,db,admin,['训练']))
    script=db.get(Script,result['script_id'])
    assert script and script.tenant_id=='训练'
    assert db.get(TrainingSession,result['session_id']).final_script_id==script.id


def test_deleting_old_training_session_does_not_delete_reimported_current_script(context,monkeypatch):
    from backend.api import training_chat
    from backend.storage import storage_manager
    db,script,admin,root=context
    previous=script.source_session_id
    current=TrainingSession(tenant_id='训练',user_id=admin.id,session_key='reimported',status='completed')
    db.add(current);db.flush();script.source_session_id=current.id;current.final_script_id=script.id;db.commit()
    deleted=[]
    manager=SimpleNamespace(get_tenant_dir=lambda tenant:root/'tenants'/tenant,
                            delete_script_files_by_code=lambda *args:deleted.append(args))
    monkeypatch.setattr(storage_manager,'StorageManager',lambda:manager)
    training_chat.delete_session(previous,db,admin)
    assert db.get(Script,script.id) and db.get(TrainingSession,current.id)
    assert not deleted


def test_admin_toggle_controls_only_latest_published_solution(context, monkeypatch):
    from datetime import datetime
    from backend.admin.router import disable_script, enable_script
    from backend.storage import storage_manager
    db, script, admin, root = context
    old = store.build(db, script, admin.id)
    old.created_at = datetime(2026, 1, 1)
    api.activate(db, old)
    script.code = 'def main(): return 321\n'; db.commit()
    latest = store.build(db, script, admin.id)
    latest.created_at = datetime(2026, 1, 2)
    api.activate(db, latest)
    changes = []
    monkeypatch.setattr(storage_manager, 'StorageManager', lambda: SimpleNamespace(
        set_script_disabled=lambda *args: changes.append(args)))
    asyncio.run(disable_script(script.id, db, admin))
    assert not script.is_active and not latest.is_active and not old.is_active
    asyncio.run(enable_script(script.id, db, admin))
    assert script.is_active and latest.is_active and not old.is_active
    assert changes == [('训练', '工资', True), ('训练', '工资', False)]
    # Repair a pre-existing mismatch even if the script was already disabled.
    script.is_active = False; latest.is_active = True; db.commit()
    asyncio.run(disable_script(script.id, db, admin))
    assert not latest.is_active


def test_disabled_script_cannot_execute_active_package(context):
    db, script, admin, root = context
    row = store.build(db, script, admin.id)
    api.activate(db, row)
    script.is_active = False; db.commit()
    with pytest.raises(ValueError, match='脚本已停用'):
        execution.binding('训练', 'solution_' + row.id, db)


def test_explicit_import_target_reuses_its_own_same_digest_version(context):
    db, source, admin, root = context
    built = store.build(db, source, admin.id)
    archive = root / 'migration.zip'
    package.export_zip(store.artifact(built.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)
    targets = [Script(tenant_id='正式', name='工资', code='old') for _ in range(2)]
    db.add_all(targets); db.commit()

    def imported(target):
        return asyncio.run(api.import_package(
            UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
            '正式', False, True, db, admin, ['正式'], script_id=str(target.id)))

    first, second = [imported(target) for target in targets]
    again = imported(targets[0])
    assert again['id'] == first['id'] and again['script_id'] == targets[0].id
    assert not db.get(CalculationSolution, second['id']).is_active
    assert db.query(CalculationSolution).filter_by(tenant_id='正式', is_active=True).count() == 1


def test_solution_selectors_show_one_current_script_per_name(context):
    db, script, admin, root = context
    row = store.build(db, script, admin.id)
    api.activate(db, row)
    archived = Script(tenant_id='训练', name=script.name, code='archived',
                      version=999, is_active=False)
    db.add(archived); db.commit()
    result = api.listing('训练', db, admin, ['训练'])
    assert [item['id'] for item in result['scripts']] == [script.id]
    assert db.get(Script, archived.id)  # History is retained, not deleted.
    script.is_active = False; db.commit()
    result = api.listing('训练', db, admin, ['训练'])
    assert len(result['scripts']) == 1


def test_training_list_hides_replaced_import_session_but_history_keeps_it(context):
    from backend.api.training_chat import list_chat_sessions
    db, source, admin, root = context
    row = store.build(db, source, admin.id)
    archive = root / 'migration.zip'
    package.export_zip(store.artifact(row.digest), archive)
    (root / 'tenants/正式').mkdir(parents=True)

    def imported():
        return asyncio.run(api.import_package(
            UploadFile(filename='migration.zip', file=io.BytesIO(archive.read_bytes())),
            '正式', False, True, db, admin, ['正式'], script_name='工资'))

    first = imported()
    source.code = 'def main(): return 456\n'; db.commit()
    newer = store.build(db, source, admin.id)
    package.export_zip(store.artifact(newer.digest), archive)
    second = imported()
    assert first['script_id'] == second['script_id']
    result = list_chat_sessions('正式', db, admin)
    assert [item['id'] for item in result['sessions']] == [second['session_id']]
    history = list_chat_sessions('正式', db, admin, include_history=True)
    assert {item['id'] for item in history['sessions']} == {first['session_id'], second['session_id']}
    assert next(item for item in history['sessions'] if item['id'] == first['session_id'])['is_history']
    from backend.solutions.lifecycle import set_script_enabled
    set_script_enabled(db, db.get(Script, second['script_id']), False); db.commit()
    assert not list_chat_sessions('正式', db, admin)['sessions']
