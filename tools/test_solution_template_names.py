"""Legacy package repair preserves all original code and restores template basenames."""
import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from backend.solutions.package import seal, verify
from backend.solutions.repairs import repair_template_names
from backend.utils.multi_workbook import build_script, script_targets, template_reference_code


class SolutionTemplateNameTests(unittest.TestCase):
    def test_legacy_repair_only_adds_named_template_copies_and_changes_bindings(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            package = root / 'original'
            targets = {}
            for owner in ['王五', '张一']:
                name = '本月工资费用账单-' + owner + '.xlsx'
                code = (f'TEMPLATE_NAME = {name!r}\n'
                        f'TEMPLATE_PATH = {("/app/tenants/满记/templates/" + name)!r}\n'
                        'def main(): return TEMPLATE_PATH\n')
                targets[name] = {'code': code, 'structure': {}}
            code = build_script(targets)
            files = {'business/script.py': code, 'archive/records.json': '{}',
                     'runtime/backend/solutions/worker.py': '# original runtime\n'}
            for name, text in files.items():
                file = package / name
                file.parent.mkdir(parents=True, exist_ok=True)
                file.write_bytes(text.encode())
            bindings = {}
            for target in script_targets(code).values():
                full_key = hashlib.sha256(target['code'].encode()).hexdigest()
                ref_key = hashlib.sha256(template_reference_code(target['code']).encode()).hexdigest()
                relative = 'assets/templates/' + full_key + '.xlsx'
                file = package / relative
                file.parent.mkdir(parents=True, exist_ok=True)
                file.write_bytes(full_key.encode())
                bindings[full_key] = bindings[ref_key] = relative
            before = seal(package, {'templates': bindings, 'settings': {'OUTPUT_TEST': '1'}, 'environment': {'test': 1}})
            repaired, after = repair_template_names(package, root / 'repairs')
            self.assertNotEqual(after['digest'], before['digest'])
            self.assertEqual(verify(package), before)
            self.assertEqual(after['settings'], before['settings'])
            self.assertEqual(after['environment'], before['environment'])
            for relative in before['files']:
                self.assertEqual((package / relative).read_bytes(), (repaired / relative).read_bytes())
            for name, target in script_targets(code).items():
                key = hashlib.sha256(target['code'].encode()).hexdigest()
                self.assertEqual(Path(after['templates'][key]).name, name)
                self.assertEqual((repaired / after['templates'][key]).read_bytes(), (package / bindings[key]).read_bytes())
            self.assertEqual(repair_template_names(repaired, root / 'repairs'), (repaired, after))


if __name__ == '__main__':
    unittest.main()
