"""Preserve existing versions while scoping package bindings to scripts."""
from sqlalchemy import inspect
from sqlalchemy.schema import CreateTable, CreateIndex


def activate_pending_solutions(engine):
    """Upgrade the latest old candidate per tenant/name to the automatic policy."""
    if 'calculation_solutions' not in inspect(engine).get_table_names():
        return
    from sqlalchemy.orm import Session
    from ..database.models import CalculationSolution, Script
    from .store import artifact
    from .package import verify
    from .lifecycle import activate
    with Session(engine) as db:
        rows = db.query(CalculationSolution).order_by(
            CalculationSolution.created_at.desc(), CalculationSolution.id.desc()).all()
        seen = set()
        for row in rows:
            key = (row.tenant_id, row.name)
            if key in seen:
                continue
            seen.add(key)
            # A newer disabled published version remains disabled.
            if row.status != 'candidate':
                continue
            if not db.query(Script).filter_by(id=row.script_id, tenant_id=row.tenant_id).first():
                continue
            try:
                verify(artifact(row.digest))
            except (OSError, ValueError):
                continue
            activate(db, row)


def migrate_script_scope(engine):
    inspector = inspect(engine)
    if 'calculation_solutions' not in inspector.get_table_names():
        return
    if not any(c.get('name') == 'uq_solution_tenant_digest'
               for c in inspector.get_unique_constraints('calculation_solutions')):
        return
    if engine.dialect.name == 'sqlite':
        from ..database.models import CalculationSolution
        table = CalculationSolution.__table__
        temporary = 'calculation_solutions_script_migration'
        ddl = str(CreateTable(table).compile(engine)).replace(
            'CREATE TABLE calculation_solutions', 'CREATE TABLE ' + temporary, 1)
        columns = ', '.join('"' + c.name + '"' for c in table.columns)
        with engine.begin() as conn:
            conn.exec_driver_sql('BEGIN')
            conn.exec_driver_sql(ddl)
            conn.exec_driver_sql(f'INSERT INTO {temporary} ({columns}) SELECT {columns} FROM calculation_solutions')
            conn.exec_driver_sql('DROP TABLE calculation_solutions')
            conn.exec_driver_sql(f'ALTER TABLE {temporary} RENAME TO calculation_solutions')
            for index in table.indexes:
                conn.execute(CreateIndex(index))
    else:
        drop = 'DROP CONSTRAINT' if engine.dialect.name == 'postgresql' else 'DROP INDEX'
        with engine.begin() as conn:
            conn.exec_driver_sql(f'ALTER TABLE calculation_solutions {drop} uq_solution_tenant_digest, '
                                 'ADD CONSTRAINT uq_solution_script_digest UNIQUE (tenant_id, script_id, digest)')
