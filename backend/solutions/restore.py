"""Restore a package into new tenant-local rows; never replay source primary keys."""
import json
import shutil
import uuid
import hashlib
from pathlib import Path
from datetime import date, datetime
from sqlalchemy import inspect, Date, DateTime
from .store import ROOT, materialize, replace_paths


def values(model, data):
    result = {}
    for attr in inspect(model).column_attrs:
        key = attr.key
        if key == 'id' or key not in data:
            continue
        value = data[key]
        typ = attr.columns[0].type
        if value and isinstance(typ, DateTime):
            value = datetime.fromisoformat(value)
            if not typ.timezone:
                from ..utils.timezone import naive_utc
                value = naive_utc(value)
        elif value and isinstance(typ, Date):
            value = date.fromisoformat(value)
        result[key] = value
    return result


def restore_solution_script(db, row, user_id):
    """Recover a missing editable binding from a verified immutable snapshot.

    Restoring creates tenant-local ids and files. It never reuses the exported
    primary key, overwrites another tenant's script, or executes packaged code.
    """
    from ..database.models import Script, CalculationSolution
    from .store import artifact
    from .package import verify
    script = db.query(Script).filter_by(id=row.script_id, tenant_id=row.tenant_id).first()
    if script:
        return script, None
    package = artifact(row.digest)
    verify(package)
    script, directory = restore(db, package, row.tenant_id, user_id, script_name=row.name)
    previous = row.script_id
    related = db.query(CalculationSolution).filter_by(
        tenant_id=row.tenant_id, name=row.name, script_id=previous).all()
    for version in related:
        version.script_id = script.id
    row.script_id = script.id
    row.source_session_id = script.source_session_id
    script.is_active = any(version.is_active for version in related) or row.is_active
    return script, directory


def prepare(package, tenant):
    """Copy potentially large archives off the API event loop, without a DB session."""
    destination = ROOT / 'tenants' / tenant / 'solution_imports' / uuid.uuid4().hex
    destination.mkdir(parents=True)
    try:
        for directory in ('archive', 'assets'):
            if (package / directory).exists():
                shutil.copytree(package / directory, destination / directory)
        return destination
    except Exception:
        shutil.rmtree(destination, ignore_errors=True)
        raise


def restore(db, package, tenant, user_id, prepared=None, target_script=None, script_name=None):
    from ..database.models import Script, TrainingSession, TrainingMessage, TrainingIteration, DataAsset
    raw = json.loads((package / 'archive/records.json').read_text(encoding='utf-8'))
    def require_archived(path):
        if path and (not isinstance(path, str) or not path.startswith('asset://') or
                     not Path(materialize(path, package)).exists()):
            raise ValueError('方案引用了未归档的训练文件，请在源环境补齐后重新导出')
    for item in raw['assets']:
        require_archived(item['file_path'])
    cfg = (raw.get('session') or {}).get('config') or {}
    for key in ('source_dir', 'expected_file', 'template_path', 'rule_files_dir'):
        require_archived(cfg.get(key))
    for path in cfg.get('expected_files') or []:
        require_archived(path)
    # The archive stays immutable. Editable training uses its own complete copy.
    destination = prepared or prepare(package, tenant)
    try:
        records = materialize(raw, destination)
        assets = {}
        for source in records['assets']:
            data = values(DataAsset, source)
            data.update(tenant_id=tenant, uploaded_by=user_id, category_id=None, source_task_id=None)
            asset = DataAsset(**data)
            db.add(asset)
            db.flush()
            assets[source['id']] = asset.id
        source = records.get('session')
        if not source:
            # Saved/legacy scripts may have no training graph. They still need
            # a destination session so editing never targets an unrelated one.
            saved = records['script']
            source = records['session'] = {
                'id': None, 'mode': saved.get('mode') or 'formula',
                'config': dict(saved.get('config') or {}),
                'source_structure': saved.get('source_structure'),
                'expected_structure': saved.get('expected_structure'),
                'rules_content': saved.get('rules_content'),
                'manual_headers': saved.get('manual_headers'),
                'total_iterations': 0,
            }
        session = None
        restored_rows = []
        id_maps = {'asset_id': assets}
        if source:
            data = values(TrainingSession, source)
            data.update(tenant_id=tenant, user_id=user_id, session_key='solution_' + uuid.uuid4().hex,
                        final_script_id=None, source_asset_ids=[assets[i] for i in source.get('source_asset_ids') or []],
                        expected_asset_id=assets.get(source.get('expected_asset_id')))
            session = TrainingSession(**data)
            db.add(session)
            db.flush()
            id_maps['session_id'] = {source['id']: session.id}
            # Keep imported attachments under the normal training session root
            # so subsequent editing, downloads and conversions work unchanged.
            session_root = ROOT / 'tenants' / tenant / 'training_chat' / str(session.id)
            session_root.parent.mkdir(parents=True, exist_ok=True)
            if session_root.exists():
                raise ValueError('目标训练目录已存在，导入已停止')
            shutil.move(str(destination), str(session_root))
            old_root, destination = destination, session_root
            records = replace_paths(records, {str(old_root): str(destination)})
            source = records['session']
            session.config = source.get('config')
            session.source_structure = source.get('source_structure')
            session.expected_structure = source.get('expected_structure')
            for old, new in assets.items():
                row = db.get(DataAsset, new)
                row.file_path = replace_paths(row.file_path, {str(old_root): str(destination)})
                restored_rows.append(row)
            for model, key in ((TrainingMessage, 'messages'), (TrainingIteration, 'iterations')):
                mapping = id_maps['message_id' if key == 'messages' else 'iteration_id'] = {}
                for item in records[key]:
                    data = values(model, item)
                    data['session_id'] = session.id
                    row = model(**data)
                    db.add(row)
                    db.flush()
                    mapping[item['id']] = row.id
                    restored_rows.append(row)
        data = values(Script, records['script'])
        data.update(tenant_id=tenant, created_by=user_id, source_session_id=session.id if session else None,
                    is_active=False, updated_at=datetime.utcnow(), code=(package / 'business/script.py').read_bytes().decode('utf-8'))
        if target_script is not None:
            if target_script.tenant_id != tenant:
                raise ValueError('目标脚本不属于当前租户')
            data.pop('name', None)
            data.pop('created_at', None)
            data.pop('created_by', None)
            script = target_script
            for key, value in data.items():
                setattr(script, key, value)
        else:
            if script_name:
                data['name'] = script_name
            script = Script(**data)
        db.add(script)
        db.flush()
        if session:
            session.final_script_id = script.id
        id_maps['script_id'] = {records['script']['id']: script.id}
        id_maps['final_script_id'] = id_maps['script_id']
        id_maps['source_session_id'] = id_maps.get('session_id', {})
        id_maps['expected_asset_id'] = assets
        id_maps['source_asset_ids'] = assets
        def remap(value, key=None):
            if isinstance(value, dict):
                return {k: remap(v, k) for k, v in value.items()}
            if isinstance(value, list):
                return [remap(v, key) for v in value]
            if key == 'tenant_id':
                return tenant
            if key in ('user_id', 'created_by', 'uploaded_by'):
                return user_id
            if key in id_maps and value is not None:
                return id_maps[key].get(value)
            return value
        from sqlalchemy import JSON
        for row in [script, session, *restored_rows]:
            if row is None:
                continue
            for attr in inspect(type(row)).column_attrs:
                if isinstance(attr.columns[0].type, JSON):
                    # source_asset_ids has already been remapped above.
                    if row is session and attr.key == 'source_asset_ids':
                        continue
                    setattr(row, attr.key, remap(getattr(row, attr.key)))
        if session:
            # The package script can differ from the last archived iteration.
            # Record the exact imported code as an editable revision, without
            # inventing a local execution or assigning it an accuracy score.
            number = max([session.total_iterations or 0] +
                         [r['iteration_num'] for r in records['iterations']]) + 1
            script_path = destination / f'iter_{number}' / 'script.py'
            script_path.parent.mkdir(parents=True, exist_ok=True)
            script_path.write_bytes(script.code.encode('utf-8'))
            config = dict(session.config or {})
            manifest_path = package / 'manifest.json'
            manifest = json.loads(manifest_path.read_text(encoding='utf-8')) if manifest_path.is_file() else {}
            template = (manifest.get('templates') or {}).get(hashlib.sha256(script.code.encode('utf-8')).hexdigest())
            if template:
                from .package import safe_name
                config['template_path'] = str(destination / safe_name(template))
            config.update(validation_stale=True, latest_detailed_diff='',
                          latest_files={'script_file': str(script_path)},
                          solution_import={'script_id': script.id, 'baseline_iteration': number,
                                           'digest': package.name})
            config['script_name'] = script.name
            session.config = config
            session.total_iterations = number
            session.status = 'completed'
            db.add(TrainingIteration(session_id=session.id, iteration_num=number, status='imported',
                                     generated_code=script.code, accuracy=None,
                                     prompt_text='[导入方案代码，尚未在本环境执行]',
                                     execution_result={'source': 'solution_import', 'validation_pending': True,
                                                       'digest': package.name, 'script_id': script.id},
                                     finished_at=datetime.utcnow()))
        return script, destination
    except Exception:
        shutil.rmtree(destination, ignore_errors=True)
        raise
