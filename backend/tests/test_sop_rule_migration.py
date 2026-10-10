"""在内存数据库验证规则版本迁移，不连接业务数据库。"""
import ast
from datetime import datetime
from pathlib import Path
import unittest

from sqlalchemy import (Boolean, Column, DateTime, ForeignKey, Integer, MetaData,
                        String, Table, Text, create_engine, inspect)
from sqlalchemy.orm import declarative_base, relationship
from sqlalchemy.schema import CreateTable


class SopRuleMigrationTests(unittest.TestCase):
    def test_project_number_migration_preserves_legacy_rows_and_enforces_unique_index(self):
        root = Path(__file__).parent.parent / "database"
        models = ast.parse((root / "models.py").read_text(encoding="utf-8"))
        model = next(node for node in models.body if isinstance(node, ast.ClassDef) and node.name == "SopEntry")
        namespace = dict(globals(), Base=declarative_base())
        exec(compile(ast.Module(body=[model], type_ignores=[]), "sop_entry_model", "exec"), namespace)
        engine = create_engine("sqlite://")
        self.addCleanup(engine.dispose)
        with engine.begin() as conn:
            conn.exec_driver_sql("CREATE TABLE sop_entries (id INTEGER PRIMARY KEY, customer_name VARCHAR(200))")
            conn.exec_driver_sql("INSERT INTO sop_entries VALUES (1, '旧客户甲'), (2, '旧客户乙')")
        tree = ast.parse((root / "init_db.py").read_text(encoding="utf-8"))
        migration = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_migrate_add_columns")
        cutoff = next(i for i, node in enumerate(migration.body)
                      if isinstance(node, ast.ImportFrom) and node.module == "sqlalchemy.orm.attributes")
        migration.body = migration.body[:cutoff]
        namespace.update(engine=engine, _migrate_sop_rule_versions=lambda: None)
        exec(compile(ast.Module(body=[migration], type_ignores=[]), "sop_column_migration", "exec"), namespace)
        namespace["_migrate_add_columns"]()
        namespace["_migrate_add_columns"]()
        indexes = inspect(engine).get_indexes("sop_entries")
        self.assertTrue(any(index["name"] == "ix_sop_entries_project_number" and index["unique"] for index in indexes))
        from sqlalchemy.exc import IntegrityError
        with engine.begin() as conn:
            self.assertEqual(conn.exec_driver_sql("SELECT count(*) FROM sop_entries WHERE project_number IS NULL").scalar(), 2)
            conn.exec_driver_sql("UPDATE sop_entries SET project_number='001' WHERE id=1")
        with self.assertRaises(IntegrityError), engine.begin() as conn:
            conn.exec_driver_sql("UPDATE sop_entries SET project_number='001' WHERE id=2")

    def test_preserves_rule_ids_and_review_references_and_allows_new_versions(self):
        root = Path(__file__).parent.parent / "database"
        models = ast.parse((root / "models.py").read_text(encoding="utf-8"))
        model = next(node for node in models.body if isinstance(node, ast.ClassDef) and node.name == "SopRuleFile")
        namespace = dict(globals(), Base=declarative_base())
        exec(compile(ast.Module(body=[model], type_ignores=[]), "sop_model", "exec"), namespace)
        rule_table = namespace["SopRuleFile"].__table__
        Table("users", rule_table.metadata, Column("id", Integer, primary_key=True))
        engine = create_engine("sqlite://")
        ddl = str(CreateTable(rule_table).compile(engine)).rstrip()
        ddl = ddl[:ddl.rfind(")")] + ', CONSTRAINT uq_sop_rule_scope_customer UNIQUE (scope, customer_name)\n)'
        with engine.begin() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=ON")
            conn.exec_driver_sql("CREATE TABLE users (id INTEGER PRIMARY KEY)")
            conn.exec_driver_sql(ddl)
            conn.execute(rule_table.insert().values(id=8, scope="customer", customer_name="客户", file_path="old.md", file_name="old.md"))
            conn.exec_driver_sql("CREATE TABLE review_reference (rule_id INTEGER REFERENCES sop_rule_files(id))")
            conn.exec_driver_sql("INSERT INTO review_reference VALUES (8)")
        migration_tree = ast.parse((root / "init_db.py").read_text(encoding="utf-8"))
        migration = next(node for node in migration_tree.body if isinstance(node, ast.FunctionDef)
                         and node.name == "_migrate_sop_rule_versions")
        namespace["engine"] = engine
        exec(compile(ast.Module(body=[migration], type_ignores=[]), "sop_migration", "exec"), namespace)
        namespace["_migrate_sop_rule_versions"]()
        namespace["_migrate_sop_rule_versions"]()
        self.assertEqual(inspect(engine).get_unique_constraints("sop_rule_files"), [])
        with engine.begin() as conn:
            self.assertEqual(conn.exec_driver_sql("SELECT id FROM sop_rule_files").scalar(), 8)
            self.assertEqual(conn.exec_driver_sql("SELECT rule_id FROM review_reference").scalar(), 8)
            self.assertEqual(conn.exec_driver_sql("PRAGMA foreign_keys").scalar(), 1)
            self.assertEqual(conn.exec_driver_sql("PRAGMA foreign_key_check").all(), [])
            conn.execute(rule_table.insert().values(scope="customer", customer_name="客户", file_path="new.md", file_name="new.md"))
            self.assertEqual(conn.exec_driver_sql("SELECT count(*) FROM sop_rule_files").scalar(), 2)
        engine.dispose()


if __name__ == "__main__":
    unittest.main()
