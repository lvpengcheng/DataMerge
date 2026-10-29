"""Package creation and training graph archival. SQL IDs never cross environments."""
import hashlib
import json
import os
import shutil
import tempfile
import uuid
import ast
from datetime import datetime, date
from pathlib import Path

from .package import canonical, digest_file, seal, verify, scrub
from .environment import fingerprint, calculation_settings

ROOT = Path(__file__).resolve().parents[2]
STORE = ROOT / 'solution_store'
FILES = {'.xlsx', '.xls', '.xlsm', '.csv', '.txt', '.md', '.pdf', '.docx', '.json', '.py', '.zip'}


def tenant_name(value):
    value = str(value or '').strip()
    if (not value or len(value) > 100 or value.startswith('.') or
            any(c in value for c in '/\\:*?"<>|') or value.endswith('.') or '..' in value):
        raise ValueError('请填写有效租户名称，不能包含路径字符')
    from .package import safe_name
    safe_name(value)
    return value


def artifact(digest):
    if len(digest) != 64 or any(c not in '0123456789abcdef' for c in digest):
        raise ValueError('无效方案摘要')
    return STORE / digest


def row_data(row):
    from sqlalchemy import inspect
    return {p.key: scrub(getattr(row, p.key).isoformat() if isinstance(getattr(row, p.key), (datetime, date))
                         else getattr(row, p.key)) for p in inspect(type(row)).column_attrs}


def replace_paths(value, replacements):
    if isinstance(value, dict):
        return {k: replace_paths(v, replacements) for k, v in value.items()}
    if isinstance(value, list):
        return [replace_paths(v, replacements) for v in value]
    if isinstance(value, str):
        if value.startswith(('{', '[')):
            try:
                return json.dumps(replace_paths(json.loads(value), replacements), ensure_ascii=False)
            except (ValueError, TypeError):
                pass
        for old in sorted(replacements, key=len, reverse=True):
            if value == old or value.startswith(old + '/') or value.startswith(old + '\\'):
                return replacements[old] + value[len(old):].replace('\\', '/')
    return value


def collect_records(db, script):
    from ..database.models import TrainingSession, TrainingMessage, TrainingIteration, DataAsset
    session = db.query(TrainingSession).filter_by(id=script.source_session_id).first() if script.source_session_id else None
    messages = db.query(TrainingMessage).filter_by(session_id=session.id).order_by(TrainingMessage.id).all() if session else []
    iterations = db.query(TrainingIteration).filter_by(session_id=session.id).order_by(TrainingIteration.iteration_num).all() if session else []
    ids = set((session.source_asset_ids or []) if session else [])
    if session and session.expected_asset_id:
        ids.add(session.expected_asset_id)
    assets = db.query(DataAsset).filter(DataAsset.id.in_(ids)).all() if ids else []
    if {r.id for r in assets} != ids:
        raise ValueError('训练关联资产已缺失，无法生成完整方案包')
    if session and session.tenant_id != script.tenant_id:
        raise ValueError('脚本关联了其他租户的训练会话')
    if any(a.tenant_id not in (None, script.tenant_id) for a in assets):
        raise ValueError('训练引用了其他租户的资产')
    return {'script': row_data(script), 'session': row_data(session) if session else None,
            'messages': [row_data(r) for r in messages], 'iterations': [row_data(r) for r in iterations],
            'assets': [row_data(r) for r in assets]}


def build(db, script, user_id):
    from ..database.models import CalculationSolution
    from ..utils.template_resolver import extract_template_ref, resolve_template_path, find_nonportable_absolute_paths
    from ..utils.multi_workbook import script_targets
    tenant_name(script.tenant_id)
    db.flush()
    records = collect_records(db, script)
    # A saved script can be frozen while its conversation continues training.
    # Archive the records visible now; subsequent iterations are a later revision.
    script_code = script.code
    script_record = records['script']
    STORE.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.build-', dir=STORE))
    original_hashes, paths, missing = {}, {}, []
    try:
        def copy_file(source, relative):
            original = str(source)
            source = Path(source).resolve()
            allowed = (ROOT / 'tenants' / script.tenant_id).resolve()
            if not (source.is_relative_to(allowed) or source.is_relative_to((ROOT / 'global_assets').resolve())):
                raise ValueError('训练资源不在当前租户或公共资料目录中')
            if not source.is_file():
                missing.append(str(source))
                return None
            target = stage / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            before = digest_file(source)
            shutil.copyfile(source, target)
            if digest_file(target) != before:
                raise ValueError('资源在打包期间变化，请重试')
            original_hashes[str(source)] = before
            paths[str(source)] = 'asset://' + relative
            paths[original] = 'asset://' + relative
            if source.is_relative_to(ROOT):
                paths[str(source.relative_to(ROOT))] = 'asset://' + relative
                paths[source.relative_to(ROOT).as_posix()] = 'asset://' + relative
            return relative

        session = records['session']
        if session:
            session_root = ROOT / 'tenants' / script.tenant_id / 'training_chat' / str(session['id'])
            if session_root.is_dir():
                for f in sorted(session_root.rglob('*')):
                    if f.is_symlink():
                        raise ValueError('训练目录包含链接，请先处理后转方案')
                    if f.is_file() and f.name != '.env' and '__pycache__' not in f.parts:
                        copy_file(f, 'archive/session/' + f.relative_to(session_root).as_posix())
                paths[str(session_root.resolve())] = 'asset://archive/session'
            elif (session.get('config') or {}).get('archive_origin') != 'legacy':
                missing.append('训练文件目录不存在')
        else:
            missing.append('脚本没有关联训练会话')
        for asset in records['assets']:
            copied = copy_file(asset['file_path'], f"archive/assets/{asset['id']}/{Path(asset['file_path']).name}")
            if not copied:
                missing.append(f"训练资产 {asset['id']} 文件不存在")
        # Include referenced revisions/attachments even when older sessions stored
        # them outside training_chat/<id>. Resolve only tenant/public paths.
        def referenced(value):
            if isinstance(value, dict):
                for item in value.values():
                    referenced(item)
            elif isinstance(value, list):
                for item in value:
                    referenced(item)
            elif isinstance(value, str) and len(value) < 1500 and '\n' not in value:
                try:
                    candidate = Path(value).resolve()
                    permitted = any(candidate.is_relative_to((ROOT / folder).resolve()) for folder in
                                    ('tenants/' + script.tenant_id, 'global_assets'))
                    if not permitted or str(candidate) in paths:
                        return
                    if not candidate.exists():
                        missing.append(value)
                    elif candidate.is_file():
                        key = hashlib.sha256(str(candidate).encode()).hexdigest()[:20]
                        copy_file(candidate, 'archive/attachments/' + key + '/' + candidate.name)
                    elif candidate.is_dir():
                        key = hashlib.sha256(str(candidate).encode()).hexdigest()[:20]
                        relative = 'archive/attachments/' + key
                        for f in candidate.rglob('*'):
                            if f.is_symlink():
                                raise ValueError('训练资源包含链接')
                            if f.is_file() and f.name != '.env' and '__pycache__' not in f.parts:
                                copy_file(f, relative + '/' + f.relative_to(candidate).as_posix())
                        paths[value] = paths[str(candidate)] = 'asset://' + relative
                except OSError:
                    return
        referenced(records)
        templates = {}
        codes = [script_code] + [t['code'] for t in script_targets(script_code).values()]
        for code in codes:
            checked = code
            if script_targets(code):
                from ..utils.multi_workbook import script_wrapper
                checked = script_wrapper(code)
            if find_nonportable_absolute_paths(checked):
                raise ValueError('脚本含模板以外的固定绝对路径，请改为运行参数后转方案')
            name, _, baked = extract_template_ref(code)
            if name or baked:
                bound = (script_record.get('config') or {}).get('template_path') if code == script_code else None
                if bound and not Path(bound).is_absolute():
                    bound = str(ROOT / bound)
                resolved = resolve_template_path(tenant_id=script.tenant_id, script_code=code,
                                                 uploaded_override=bound, project_root=ROOT)
                if not resolved:
                    raise ValueError(f'必要模板缺失: {name or baked}')
                key = hashlib.sha256(code.encode()).hexdigest()
                # Business scripts may use the template basename to select the
                # workbook's data scope. Isolate by hash directory, keep its name.
                templates[key] = copy_file(resolved, 'assets/templates/' + key + '/' + Path(resolved).name)
                from ..utils.multi_workbook import template_reference_code
                reference = template_reference_code(code)
                if reference:
                    templates[hashlib.sha256(reference.encode()).hexdigest()] = templates[key]
        # Include the entire Python application closure; no current-project fallback.
        suffixes = {'.py', '.json', '.txt', '.md', '.yaml', '.yml', '.j2', '.xlsx', '.xltx'}
        sources = [p for p in (ROOT / 'backend').rglob('*') if p.is_file() and p.suffix in suffixes]
        sources += [ROOT / n for n in ('excel_parser.py', 'aspose_init.py', 'split_by_banner.py')]
        for src in sources:
            if any(part in src.parts for part in ('__pycache__', 'tests', 'test_data', 'docs', 'logs')) or src.name.startswith('test_'):
                continue
            dst = stage / 'runtime' / src.relative_to(ROOT)
            dst.parent.mkdir(parents=True, exist_ok=True)
            original_hashes[str(src)] = digest_file(src)
            shutil.copyfile(src, dst)
        # Native binaries are included; deployment license is explicitly excluded.
        for src in (ROOT / 'libs').rglob('*'):
            if src.is_file() and src.suffix.lower() in ('.dll', '.so', '.json', '.dylib'):
                dst = stage / 'runtime/libs' / src.relative_to(ROOT / 'libs')
                dst.parent.mkdir(parents=True, exist_ok=True)
                original_hashes[str(src)] = digest_file(src)
                shutil.copyfile(src, dst)
        (stage / 'business').mkdir()
        (stage / 'business/script.py').write_bytes(script_code.encode('utf-8'))
        (stage / 'archive').mkdir(exist_ok=True)
        records = replace_paths(records, paths)
        (stage / 'archive/records.json').write_bytes(canonical(records))
        from .inputs import snapshot
        snapshot(ROOT, script.tenant_id, stage / 'archive/context', db)
        metadata = {k: records['script'].get(k) for k in ('manual_headers', 'source_structure', 'expected_structure', 'rules_content')}
        metadata['use_history'] = (script_record.get('config') or {}).get('use_history')
        if records['session']:
            cfg = records['session'].get('config') or {}
            for key in ('manual_headers', 'source_structure', 'expected_structure', 'rules_content'):
                metadata[key] = metadata.get(key) or records['session'].get(key) or cfg.get(key)
        # Files must be internally consistent; live training rows may advance
        # after the already captured archive without invalidating that archive.
        if any(digest_file(p) != h for p, h in original_hashes.items()):
            raise ValueError('代码或资源在打包期间变化，请重试')
        manifest = seal(stage, {'name': script_record['name'], 'created_at': datetime.utcnow().isoformat(),
            'environment': fingerprint(ROOT), 'settings': calculation_settings(), 'metadata': metadata,
            'templates': templates, 'missing_archive': missing, 'source_script_id': script_record['id'],
            'source_tenant': script_record['tenant_id'], 'protocol': 1, 'mode': script_record.get('mode')})
        destination = artifact(manifest['digest'])
        if destination.exists():
            verify(destination)
            shutil.rmtree(stage)
        else:
            os.replace(stage, destination)
        row = CalculationSolution(id=uuid.uuid4().hex, tenant_id=script_record['tenant_id'], script_id=script_record['id'],
            source_session_id=script_record.get('source_session_id'), name=script_record['name'], digest=manifest['digest'], created_by=user_id)
        db.add(row)
        db.commit()
        return row
    finally:
        if stage.exists() and stage.name.startswith('.build-'):
            shutil.rmtree(stage)


def materialize(value, root):
    if isinstance(value, dict):
        return {k: materialize(v, root) for k, v in value.items()}
    if isinstance(value, list):
        return [materialize(v, root) for v in value]
    if isinstance(value, str) and value.startswith('asset://'):
        from .package import safe_name
        return str(Path(root) / safe_name(value[8:]))
    if isinstance(value, str) and value.startswith(('{', '[')):
        try:
            return json.dumps(materialize(json.loads(value), root), ensure_ascii=False)
        except (ValueError, TypeError):
            pass
    return value
