"""Switch the current scheme without modifying any immutable package."""
from sqlalchemy import or_
from ..database.models import CalculationSolution, Script


def activate(db, row, *, commit=True):
    for previous in db.query(CalculationSolution).filter_by(tenant_id=row.tenant_id, created_by=row.created_by).filter(
            or_(CalculationSolution.name == row.name, CalculationSolution.script_id == row.script_id)).all():
        if previous.id != row.id:
            previous.is_active = False
    for script in db.query(Script).filter_by(tenant_id=row.tenant_id, name=row.name, created_by=row.created_by).all():
        script.is_active = script.id == row.script_id
    row.status, row.is_active = 'published', True
    if commit:
        db.commit()
    return row


def set_script_enabled(db, script, enabled):
    """Admin script state controls its published packages in the same transaction."""
    rows = db.query(CalculationSolution).filter_by(
        tenant_id=script.tenant_id, script_id=script.id).order_by(
            CalculationSolution.created_at.desc(), CalculationSolution.id.desc()).all()
    script.is_active = enabled
    if not enabled:
        for row in rows:
            row.is_active = False
        return
    latest = next((row for row in rows if row.status == 'published'), None)
    if latest:
        activate(db, latest, commit=False)
