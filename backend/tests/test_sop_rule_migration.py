"""在内存数据库验证规则版本迁移，不连接业务数据库。"""
import ast
from datetime import datetime
from pathlib import Path
import unittest

from sqlalchemy import (Boolean, Column, DateTime, ForeignKey, Integer, MetaData,
                        String, Table, Text, Index, JSON, create_engine, inspect)
from sqlalchemy.orm import declarative_base, relationship
from sqlalchemy.schema import CreateTable


class SopRuleMigrationTests(unittest.TestCase):
    def test_customer_migration_removes_number_preserves_rows_and_enforces_unique_index(self):
        root = Path(__file__).parent.parent / "database"
        models = ast.parse((root / "models.py").read_text(encoding="utf-8"))
        model = next(node for node in models.body if isinstance(node, ast.ClassDef) and node.name == "SopEntry")
        namespace = dict(globals(), Base=declarative_base())
        exec(compile(ast.Module(body=[model], type_ignores=[]), "sop_entry_model", "exec"), namespace)
        engine = create_engine("sqlite://")
        self.addCleanup(engine.dispose)
        with engine.begin() as conn:
            conn.exec_driver_sql("CREATE TABLE sop_entries (id INTEGER PRIMARY KEY, customer_name VARCHAR(200), project_number VARCHAR(100))")
            conn.exec_driver_sql("CREATE UNIQUE INDEX uq_sop_project_customer ON sop_entries (project_number, customer_name)")
            conn.exec_driver_sql("CREATE INDEX ix_sop_entries_project_number ON sop_entries (project_number)")
            conn.exec_driver_sql("INSERT INTO sop_entries VALUES (1, '旧客户甲', '001'), (2, '旧客户乙', '001')")
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
        self.assertTrue(any(index["name"] == "uq_sop_customer_name" and index["unique"] for index in indexes))
        self.assertFalse(any(index["name"] in ("uq_sop_project_customer", "ix_sop_entries_project_number") for index in indexes))
        self.assertNotIn("project_number", {col["name"] for col in inspect(engine).get_columns("sop_entries")})
        self.assertIn("supplementary_files", {col["name"] for col in inspect(engine).get_columns("sop_entries")})
        from sqlalchemy.exc import IntegrityError
        with engine.begin() as conn:
            self.assertEqual(conn.exec_driver_sql("SELECT count(*) FROM sop_entries").scalar(), 2)
            self.assertEqual(conn.exec_driver_sql("SELECT count(*) FROM sop_entries WHERE supplementary_files IS NULL").scalar(), 2)
        with self.assertRaises(IntegrityError), engine.begin() as conn:
            conn.exec_driver_sql("UPDATE sop_entries SET customer_name='旧客户甲' WHERE id=2")

        # 历史同名租户保留原数据，暂缓唯一约束升级，但必须允许服务继续启动。
        duplicate_engine = create_engine("sqlite://")
        self.addCleanup(duplicate_engine.dispose)
        with duplicate_engine.begin() as conn:
            conn.exec_driver_sql("CREATE TABLE sop_entries (id INTEGER PRIMARY KEY, customer_name VARCHAR(200), project_number VARCHAR(100))")
            conn.exec_driver_sql("INSERT INTO sop_entries VALUES (1, '同名客户', '001'), (2, '同名客户', '002')")
        namespace["engine"] = duplicate_engine
        namespace["_migrate_add_columns"]()
        namespace["_migrate_add_columns"]()
        self.assertFalse(any(index["name"] == "uq_sop_customer_name" for index in inspect(duplicate_engine).get_indexes("sop_entries")))
        with duplicate_engine.connect() as conn:
            self.assertEqual(conn.exec_driver_sql("SELECT count(*) FROM sop_entries").scalar(), 2)
            self.assertEqual(conn.exec_driver_sql("SELECT project_number FROM sop_entries ORDER BY id").scalars().all(), ["001", "002"])

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
