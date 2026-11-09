"""Resolve a tenant's active script consistently by ID, name or content hash."""
import hashlib
import re


def canonical_script_id(value):
    value = str(value)
    match = re.fullmatch(r'script[_.]([0-9a-f]{12})', value)
    return 'script_' + match.group(1) if match else value


def find_script(db, model, tenant_id, script_id):
    identity = canonical_script_id(script_id)
    query = db.query(model).filter_by(tenant_id=tenant_id, is_active=True)
    ordered = query.order_by(model.version.desc(), model.id.desc())
    if re.fullmatch(r'script_[0-9a-f]{12}', identity):
        digest = identity[7:]
        for row in ordered.all():
            if row.code and hashlib.md5(row.code.encode('utf-8')).hexdigest()[:12] == digest:
                return row
        return None
    if identity.isdecimal():
        return query.filter_by(id=int(identity)).first()
    return ordered.filter_by(name=str(script_id)).first()


def current_training_scripts(db, tenant_id):
    """Choose one runnable script per training session without scanning disk history."""
    from backend.database.models import Script, TrainingSession

    scripts = db.query(Script).filter_by(tenant_id=tenant_id, is_active=True).order_by(Script.id.desc()).all()
    by_id = {script.id: script for script in scripts if script.code}
    latest_by_session = {}
    for script in by_id.values():
        if script.source_session_id is not None:
            latest_by_session.setdefault(script.source_session_id, script)

    selected = {}
    sessions = db.query(TrainingSession).filter_by(tenant_id=tenant_id).all()
    for session in sessions:
        script = by_id.get(session.final_script_id)
        # A migrated script may have moved to a new session since this pointer was saved.
        if script is None or script.source_session_id not in (None, session.id):
            script = latest_by_session.get(session.id)
        if script is not None:
            selected[script.id] = script
    return sorted(selected.values(), key=lambda script: script.id, reverse=True)
