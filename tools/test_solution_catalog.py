"""Missing frozen packages must not break unrelated tenant script selection."""
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from backend.solutions.catalog import tenant_versions
from backend.solutions.package import verify


class SolutionCatalogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.patch = patch('backend.solutions.store.STORE', self.root)
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def row(self, number, status='published', active=True):
        return SimpleNamespace(id=str(number), name=f'方案{number}', digest=f'{number:064x}',
                               status=status, is_active=active, validation={'accuracy': 1.0})

    def write(self, row, content):
        directory = self.root / row.digest
        directory.mkdir()
        (directory / 'manifest.json').write_text(content, encoding='utf-8')

    def test_missing_package_does_not_hide_other_available_solutions(self):
        missing, healthy = self.row(1), self.row(2)
        self.write(healthy, json.dumps({'digest': healthy.digest, 'mode': 'template'}))
        versions, pending, unavailable = tenant_versions([missing, healthy])
        self.assertEqual([v['solution_id'] for v in versions], ['2'])
        self.assertTrue(versions[0]['is_active'])
        self.assertEqual(pending, [])
        self.assertEqual([v['id'] for v in unavailable], ['1'])
        self.assertNotIn(str(self.root), unavailable[0]['reason'])
        # A missing frozen version remains an execution error, never a new package.
        with self.assertRaisesRegex(ValueError, '清单缺失'):
            verify(self.root / missing.digest)

    def test_malformed_and_mismatched_manifests_are_unavailable(self):
        rows = [self.row(i) for i in range(1, 5)]
        for row, content in zip(rows, ('not json', '[]', '{}', json.dumps({'digest': 'a' * 64}))):
            self.write(row, content)
        versions, _, unavailable = tenant_versions(rows)
        self.assertEqual(versions, [])
        self.assertEqual(len(unavailable), 4)

    def test_candidates_and_disabled_versions_do_not_require_files(self):
        versions, pending, unavailable = tenant_versions([self.row(1, 'candidate'), self.row(2, active=False)])
        self.assertEqual(versions, [])
        self.assertEqual(len(pending), 1)
        self.assertEqual(unavailable, [])

    def test_package_store_is_persistent_in_container_and_writable_in_native_deployment(self):
        compose = Path('docker-compose.yml').read_text(encoding='utf-8')
        self.assertIn('./solution_store:/app/solution_store', compose)
        native = Path('deploy_native_ubuntu.sh').read_text(encoding='utf-8')
        self.assertIn('global_assets,solution_store}', native)


if __name__ == '__main__':
    unittest.main()
