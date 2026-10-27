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
