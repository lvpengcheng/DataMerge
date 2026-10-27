"""Script content and metadata use identical tenant-scoped identity resolution."""
import hashlib
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from backend.storage.storage_manager import StorageManager
from backend.utils.script_lookup import find_script


class Field:
    def desc(self):
        return self


class ScriptModel:
    version = Field()
    id = Field()


class Query:
    def __init__(self, rows):
        self.rows = rows
    def filter_by(self, **filters):
        return Query([r for r in self.rows if all(getattr(r, k) == v for k, v in filters.items())])
    def order_by(self, *args):
        return Query(sorted(self.rows, key=lambda r: (r.version, r.id), reverse=True))
    def all(self):
        return self.rows
    def first(self):
        return next(iter(self.rows), None)


class ScriptLookupTests(unittest.TestCase):
    def setUp(self):
        self.code = 'def main():\n    return 1\n'
        self.digest = hashlib.md5(self.code.encode()).hexdigest()[:12]
        self.identity = 'script_' + self.digest
        self.rows = [SimpleNamespace(id=1, version=1, name='友好名称', tenant_id='本租户',
                                    code=self.code, is_active=True)]
        self.db = SimpleNamespace(query=lambda _: Query(self.rows), close=lambda: None)

    def test_id_name_and_hash_resolve_same_active_tenant_script(self):
        for identity in ('1', '友好名称', self.identity, 'script.' + self.digest):
            self.assertIs(find_script(self.db, ScriptModel, '本租户', identity), self.rows[0])
        self.assertIsNone(find_script(self.db, ScriptModel, '另一租户', self.identity))
        self.rows[0].is_active = False
        self.assertIsNone(find_script(self.db, ScriptModel, '本租户', self.identity))

    def test_training_archive_is_loaded_when_scripts_directory_is_missing(self):
        with tempfile.TemporaryDirectory() as directory:
            manager = StorageManager(directory)
            path = Path(directory) / '本租户/training' / self.identity / (self.identity + '.py')
            path.parent.mkdir(parents=True)
            path.write_text(self.code, encoding='utf-8')
            self.assertEqual(manager.get_script_content('本租户', self.identity), self.code)
            self.assertEqual(manager.get_script_content('本租户', 'script.' + self.digest), self.code)
            manager.set_script_disabled('本租户', self.identity, True)
            self.assertTrue(manager.is_script_disabled('本租户', 'script.' + self.digest))

    def test_db_hash_fallback_returns_code_when_all_files_are_missing(self):
        import sys
        with tempfile.TemporaryDirectory() as directory, patch.dict(sys.modules, {
            'backend.database.connection': SimpleNamespace(SessionLocal=lambda: self.db),
            'backend.database.models': SimpleNamespace(Script=ScriptModel),
        }):
            self.assertEqual(StorageManager(directory).get_script_content('本租户', self.identity), self.code)


if __name__ == '__main__':
    unittest.main()
