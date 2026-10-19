"""Preserve existing versions while scoping package bindings to scripts."""
from sqlalchemy import inspect
from sqlalchemy.schema import CreateTable, CreateIndex


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
