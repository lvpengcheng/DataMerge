"""History source downloads resolve retained task inputs without mixing task or tenant files."""
import ast
import hashlib
import logging
from pathlib import Path
from types import SimpleNamespace

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from backend.api import assets, compute
from backend.auth.dependencies import get_current_user, get_accessible_tenants, get_operable_tenants
from backend.database import models
from backend.database.connection import get_db
from backend.utils.compute_source_files import preserve_compute_source


@pytest.fixture
def history(tmp_path, monkeypatch):
    engine = create_engine('sqlite://', connect_args={'check_same_thread':False},poolclass=StaticPool)
    models.Base.metadata.create_all(engine)
    sessions = sessionmaker(bind=engine)
    app = FastAPI()
    app.include_router(assets.router)
    app.include_router(compute.router)
    user = SimpleNamespace(id=1,role=SimpleNamespace(name='user'))
    allowed = ['tenant-a']
    def database():
        with sessions() as db:
            yield db
    app.dependency_overrides[get_db] = database
    app.dependency_overrides[get_current_user] = lambda:user
    app.dependency_overrides[get_accessible_tenants] = lambda:allowed
    app.dependency_overrides[get_operable_tenants] = lambda:allowed
    monkeypatch.setattr(assets,'PROJECT_ROOT',tmp_path)
    with sessions() as db:
        db.add_all([models.ComputeTask(id=42,tenant_id='tenant-a',status='completed'),
                    models.ComputeTask(id=43,tenant_id='tenant-b',status='completed')])
        db.add(models.DataAsset(id=10,tenant_id='tenant-a',asset_type='source',name='工资.xlsx',
            file_name='工资.xlsx',file_path='session://42/工资.xlsx'))
        db.flush()
        db.add(models.ComputeTaskInput(task_id=42,asset_id=10,role='source'))
        db.commit()
    path = tmp_path/'tenants'/'tenant-a'/'compute_runs'/'task_42'/'source'/'工资.xlsx'
    path.parent.mkdir(parents=True)
    path.write_bytes(b'original task 42 source')
    with TestClient(app) as client:
        yield SimpleNamespace(client=client,sessions=sessions,root=tmp_path,path=path,allowed=allowed)
    engine.dispose()


def test_history_input_download_resolves_existing_session_marker_without_database_rewrite(history):
    detail = history.client.get('/api/compute2/tasks/42')
    assert detail.status_code == 200
    aid = detail.json()['inputs'][0]['asset_id']
    response = history.client.get(f'/api/assets/{aid}/download')
    assert response.status_code == 200, response.text
    assert response.content == history.path.read_bytes()
    with history.sessions() as db:
        assert db.get(models.DataAsset,aid).file_path == 'session://42/工资.xlsx'


def test_normal_assets_still_download_and_tenant_permissions_remain_enforced(history):
    with history.sessions() as db:
        db.get(models.DataAsset,10).file_path = str(history.path)
        db.commit()
    assert history.client.get('/api/assets/10/download').content == history.path.read_bytes()
    history.allowed.clear()
    assert history.client.get('/api/assets/10/download').status_code == 403


@pytest.mark.parametrize('path', ['session://43/工资.xlsx','session://42/../工资.xlsx','session://42/别的文件.xlsx'])
def test_session_markers_cannot_access_another_task_or_another_file(history,path):
    with history.sessions() as db:
        db.get(models.DataAsset,10).file_path = path
        db.commit()
    response = history.client.get('/api/assets/10/download')
    assert response.status_code == 404


def test_missing_source_does_not_fall_back_to_same_name_in_another_task(history):
    history.path.unlink()
    other = history.root/'tenants'/'tenant-a'/'compute_runs'/'task_99'/'source'/'工资.xlsx'
    other.parent.mkdir(parents=True)
    other.write_bytes(b'unrelated')
    response = history.client.get('/api/assets/10/download')
    assert response.status_code == 404
    assert '未保留或已清理' in response.json()['detail']


@pytest.mark.parametrize('format', ['pdf','encrypted'])
def test_conversion_downloads_use_resolved_source(history,monkeypatch,format):
    import sys
    converted = history.root/('converted.'+('pdf' if format == 'pdf' else 'xlsx'))
    converted.write_bytes(b'converted')
    calls = []
    def convert(path,**kwargs):
        calls.append(path)
        return str(converted)
    monkeypatch.setitem(sys.modules,'backend.utils.aspose_helper',SimpleNamespace(convert_to_pdf=convert,encrypt_excel=convert))
    response = history.client.get('/api/assets/10/download',params={'format':format})
    assert response.status_code == 200 and response.content == b'converted'
    assert Path(calls[0]) == history.path


def test_registering_new_compute_sources_persists_real_paths_and_survives_legacy_temp_cleanup(history):
    # Load the production registration function without starting the whole app.
    main = Path(__file__).resolve().parents[1]/'backend'/'app'/'main.py'
    tree = ast.parse(main.read_text(encoding='utf-8-sig'))
    function = next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name == '_persist_source_file')
    namespace = {'Path':Path,'__file__':str(history.root/'backend'/'app'/'main.py'),
        'db_models':models,'logger':logging.getLogger(__name__),
        '_file_sha256':lambda path:hashlib.sha256(Path(path).read_bytes()).hexdigest()}
    exec(compile(ast.Module(body=[function],type_ignores=[]),str(main),'exec'),namespace)
    with history.sessions() as db:
        aid = namespace['_persist_source_file'](db,42,'tenant-a',history.path,'工资.xlsx',history.path.stat().st_size,sheet_summary=[])
        assert aid
        assert Path(db.get(models.DataAsset,aid).file_path) == history.path
    legacy = history.root/'temporary-upload.xlsx'
    legacy.write_bytes(b'legacy original')
    with history.sessions() as db:
        legacy_id = namespace['_persist_source_file'](db,42,'tenant-a',legacy,legacy.name,legacy.stat().st_size,sheet_summary=[])
        assert legacy_id
    legacy.unlink()
    response = history.client.get(f'/api/assets/{legacy_id}/download')
    assert response.status_code == 200 and response.content == b'legacy original'


def test_permanent_sources_are_not_duplicated_and_forced_copies_are_unique(history):
    assert preserve_compute_source(history.root,'tenant-a',42,history.path,history.path.name) == history.path
    one = preserve_compute_source(history.root,'tenant-a',42,history.path,history.path.name,copy_file=True)
    two = preserve_compute_source(history.root,'tenant-a',42,history.path,history.path.name,copy_file=True)
    assert one != two and one.read_bytes() == two.read_bytes() == history.path.read_bytes()
