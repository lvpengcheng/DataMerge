"""Switch the current scheme without modifying any immutable package."""
from ..database.models import CalculationSolution, Script


def activate(db, row):
    for previous in db.query(CalculationSolution).filter_by(tenant_id=row.tenant_id, name=row.name).all():
        if previous.id != row.id:
            previous.is_active = False
    for script in db.query(Script).filter_by(tenant_id=row.tenant_id, name=row.name).all():
        script.is_active = script.id == row.script_id
    row.status, row.is_active = 'published', True
    db.commit()
    return row
