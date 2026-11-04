"""Portable package and database round-trip regression tests (no Excel engine)."""
import json
import os
from pathlib import Path

import pytest

# Never connect to a developer's configured database in these tests.
os.environ['DATABASE_URL'] = 'sqlite://'
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from backend.database.connection import Base
from backend.database.models import Script, TrainingSession, TrainingMessage, TrainingIteration, DataAsset, User, CalculationSolution
from backend.solutions import package, store, restore

REPO = Path(__file__).resolve().parents[1]


def write(root, name, content=''):
    path = root / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding='utf-8', newline='')
    return path


def tiny_package(root):
    for name in ('business/script.py', 'archive/records.json', 'runtime/backend/solutions/worker.py'):
        write(root, name, '{}')
    return package.seal(root, {'name': 'test'})


@pytest.fixture
def db():
    engine = create_engine('sqlite://')
    Base.metadata.create_all(engine)
    with Session(engine) as session:
        session.add(User(id=1, username='test', password_hash='unused'))
        session.commit()
        yield session
    engine.dispose()


def make_training(db, root):
    session = TrainingSession(tenant_id='训练租户', user_id=1, session_key='training', status='completed')
    db.add(session); db.flush()
    folder = root / 'tenants/训练租户/training_chat' / str(session.id)
    source = write(folder, 'source/input.csv', 'id,value\n1,42')
    expected = write(folder, 'expected.xlsx', 'fixture')
    write(folder, 'unusual.bin', 'also archived')
    asset = DataAsset(tenant_id='训练租户', asset_type='source', name='input', file_path=str(source), file_name=source.name)
    db.add(asset); db.flush()
    session.source_asset_ids = [asset.id]
    session.config = {'source_dir': str(source.parent), 'expected_file': str(expected), 'source_asset_ids': [asset.id], 'api_key': 'secret'}
    script = Script(tenant_id='训练租户', name='工资', code='def main():\n    return 42\n', source_session_id=session.id)
    db.add(script); db.flush()
    session.final_script_id = script.id
    db.add(TrainingMessage(session_id=session.id, role='assistant', content='训练完成', metadata_={'asset_id': asset.id}))
    db.add(TrainingIteration(session_id=session.id, iteration_num=1, status='completed', generated_code=script.code, accuracy=1))
    db.commit()
    for name in ('excel_parser.py', 'aspose_init.py', 'split_by_banner.py'):
        write(root, name, '# frozen version\n')
    write(root, 'backend/solutions/worker.py', '# entry\n')
    return script


@pytest.mark.parametrize('count', [0, 1, 2])
def test_first_training_persists_all_expected_file_paths(tmp_path, count):
    """Run the actual config persistence block without AI or an Excel engine."""
    import ast
    from types import SimpleNamespace
    import shutil
    tree = ast.parse((REPO / 'backend/api/training_chat.py').read_text(encoding='utf-8'))
    body = next(node.body for node in ast.walk(tree) if isinstance(node, ast.Try)
                and any(isinstance(child, ast.Assign) and any(
                    isinstance(target, ast.Name) and target.id == '_cfg' for target in child.targets)
                        for child in node.body))
    start = next(i for i, node in enumerate(body) if isinstance(node, ast.Assign)
                 and any(isinstance(t, ast.Name) and t.id == '_cfg' for t in node.targets))
    end = next(i for i, node in enumerate(body[start:], start) if isinstance(node, ast.Assign)
               and any(isinstance(t, ast.Attribute) and t.attr == 'config' for t in node.targets))
    uploads = [write(tmp_path, f'uploads/target{i}.xlsx', 'expected') for i in range(count)]
    folder = tmp_path / 'persistent'
    folder.mkdir()
    expected = write(folder, uploads[0].name, 'expected') if uploads else None
    ts = SimpleNamespace(config={'expected_files': [str(p) for p in uploads]})
    scope = dict(ts=ts, p_source=folder, saved_rule_paths=[], p_expected=str(expected) if expected else None,
                 p_template=None, expected_files=[str(p) for p in uploads], Path=Path, shutil=shutil,
                 session_persist_dir=folder, mode='formula',
                 expected_struct_dict={'workbooks': {p.name: {} for p in uploads}})
    exec(compile(ast.Module(body=body[start:end + 1], type_ignores=[]), '<persist-config>', 'exec'), scope)
    assert ts.config['expected_files'] == [str(folder / p.name) for p in uploads]
    assert all(Path(path).is_file() for path in ts.config['expected_files'])
    if uploads:
        assert ts.config['expected_file'] == ts.config['expected_files'][0]


def test_old_single_target_session_exports_and_imports_without_temp_paths(tmp_path, db, monkeypatch):
    monkeypatch.setattr(store, 'ROOT', tmp_path)
    monkeypatch.setattr(store, 'STORE', tmp_path / 'solution_store')
    monkeypatch.setattr(restore, 'ROOT', tmp_path)
    monkeypatch.setattr(store, 'fingerprint', lambda root: {})
    script = make_training(db, tmp_path)
    session = db.get(TrainingSession, script.source_session_id)
    stale = str(tmp_path / 'deleted_uploads' / Path(session.config['expected_file']).name)
    session.config = dict(session.config, expected_files=[stale])
    db.commit()
    row = store.build(db, script, 1)
    frozen = store.artifact(row.digest)
    records = json.loads((frozen / 'archive/records.json').read_text(encoding='utf-8'))
    cfg = records['session']['config']
    assert cfg['expected_files'] == [cfg['expected_file']]
    assert cfg['expected_file'].startswith('asset://')
    assert session.config['expected_files'] == [stale]  # Snapshot only; live history is untouched.
    assert not package.verify(frozen)['missing_archive']
    package.export_zip(frozen, tmp_path / 'export.zip')
    package.extract_zip(tmp_path / 'export.zip', tmp_path / 'unpacked')
    imported, _ = restore.restore(db, tmp_path / 'unpacked', '目标租户', 1)
    restored = db.get(TrainingSession, imported.source_session_id)
    assert restored.config['expected_files'] == [restored.config['expected_file']]
    assert Path(restored.config['expected_file']).is_file()


def test_empty_rule_directory_survives_zip_roundtrip(tmp_path, db, monkeypatch):
    monkeypatch.setattr(store, 'ROOT', tmp_path)
    monkeypatch.setattr(store, 'STORE', tmp_path / 'solution_store')
    monkeypatch.setattr(restore, 'ROOT', tmp_path)
    monkeypatch.setattr(store, 'fingerprint', lambda root: {})
    script = make_training(db, tmp_path)
    session = db.get(TrainingSession, script.source_session_id)
    rules = tmp_path / 'tenants/训练租户/rules'
    rules.mkdir()
    session.config = dict(session.config, rule_files_dir=str(rules))
    db.commit()
    row = store.build(db, script, 1)
    package.export_zip(store.artifact(row.digest), tmp_path / 'export.zip')
    package.extract_zip(tmp_path / 'export.zip', tmp_path / 'unpacked')
    imported, _ = restore.restore(db, tmp_path / 'unpacked', '目标租户', 1)
    restored = db.get(TrainingSession, imported.source_session_id)
    assert Path(restored.config['rule_files_dir']).is_dir()


def test_build_rejects_missing_required_file_before_registering_solution(tmp_path, db, monkeypatch):
    monkeypatch.setattr(store, 'ROOT', tmp_path)
    monkeypatch.setattr(store, 'STORE', tmp_path / 'solution_store')
    script = make_training(db, tmp_path)
    session = db.get(TrainingSession, script.source_session_id)
    session.config = dict(session.config, expected_files=[str(tmp_path / 'missing/other.xlsx')])
    db.commit()
    with pytest.raises(ValueError, match=r'expected_files\[0\].*other.xlsx'):
        store.build(db, script, 1)
    assert db.query(CalculationSolution).count() == 0
    assert not list(store.STORE.iterdir())


@pytest.mark.parametrize('path', ['C:/old/upload.xlsx', 'asset://archive/missing.xlsx'])
def test_export_and_import_reject_same_reference_with_field_details(tmp_path, path):
    root = tmp_path / 'package'
    tiny_package(root)
    write(root, 'archive/records.json', json.dumps({'session': {'config': {'expected_files': [path]}}}))
    package.seal(root, {'name': 'incomplete'})
    with pytest.raises(ValueError, match=r'expected_files\[0\]') as exported:
        package.export_zip(root, tmp_path / 'export.zip')
    with pytest.raises(ValueError) as imported:
        restore.restore(None, root, 'tenant', 1)
    assert str(exported.value) == str(imported.value)
    assert not (tmp_path / 'export.zip').exists()


