"""Register a filesystem-only script without replacing its files or active pointer."""
import json
import uuid
from .store import tenant_name


def register_legacy_script(db, root, tenant, identifier, user_id):
    from ..database.models import Script, TrainingSession, TrainingIteration
    folder = root / 'tenants' / tenant_name(tenant)
    source = folder / 'scripts' / (identifier + '.py')
    if not source.is_file() or source.is_symlink():
        raise ValueError('旧脚本不存在或不是普通文件')
    code = source.read_bytes().decode('utf-8')
    existing = db.query(Script).filter_by(tenant_id=tenant, code=code).first()
    if existing and existing.source_session_id:
        return existing
    def read(path):
        return json.loads(path.read_text(encoding='utf-8')) if path.is_file() else {}
    info = read(folder / 'scripts' / (identifier + '_info.json'))
    training = folder / 'training'
    files = read(training / 'training_info.json')
    result = read(training / identifier / 'training_result.json')
    config = dict(info)
    config.update(source_dir=str(training / 'source'), expected_file=(files.get('file_paths') or {}).get('expected'),
                  rule_files_dir=str(training / 'rules'), legacy_training_dir=str(training / identifier),
                  legacy_training_info=files, legacy_training_result=result, archive_origin='legacy')
    session = TrainingSession(tenant_id=tenant, user_id=user_id, session_key='legacy_' + uuid.uuid4().hex,
                              status='completed', config=config, total_iterations=result.get('total_iterations', 0))
    db.add(session); db.flush()
    for number, item in enumerate(result.get('iteration_results') or [], 1):
        db.add(TrainingIteration(session_id=session.id, iteration_num=number, status='completed',
                                 ai_response=item.get('raw_response'), execution_result=item.get('execution_result'),
                                 error_details={'legacy_record': item}))
    script = existing or Script(tenant_id=tenant, name=identifier, code=code, created_by=user_id, config=info)
    script.source_session_id = session.id
    for key in ('manual_headers', 'source_structure', 'expected_structure', 'rules_content'):
        if not getattr(script, key):
            setattr(script, key, info.get(key) or result.get(key))
    db.add(script); db.flush()
    session.final_script_id = script.id
    return script
