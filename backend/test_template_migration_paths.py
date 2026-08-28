import hashlib
import tempfile
import unittest
from pathlib import Path

from backend.utils.template_resolver import (
    _find_by_name,
    extract_template_ref,
    find_nonportable_absolute_paths,
    portable_basename,
)


class TemplateMigrationPathTests(unittest.TestCase):
    def test_windows_baked_path_is_parsed_on_any_platform(self):
        code = 'TEMPLATE_PATH = r"E:\\deploy\\tenants\\old\\模板.xlsx"'
        name, template_hash, baked = extract_template_ref(code)
        self.assertEqual(name, "模板.xlsx")
        self.assertIsNone(template_hash)
        self.assertEqual(baked, r"E:\deploy\tenants\old\模板.xlsx")

    def test_portable_basename_handles_both_separators(self):
        self.assertEqual(portable_basename(r"C:\data\a.xlsx"), "a.xlsx")
        self.assertEqual(portable_basename("/app/data/a.xlsx"), "a.xlsx")

    def test_hash_mismatch_never_falls_back_to_same_name(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "模板.xlsx"
            path.write_bytes(b"actual-template")
            wrong_hash = hashlib.md5(b"different-template").hexdigest()
            self.assertIsNone(_find_by_name(td, path.name, wrong_hash))
            actual_hash = hashlib.md5(path.read_bytes()).hexdigest()
            self.assertEqual(_find_by_name(td, path.name, actual_hash), str(path))

    def test_source_absolute_path_is_rejected_but_template_is_rebindable(self):
        code = (
            'TEMPLATE_PATH = r"E:\\old\\template.xlsx"\n'
            'SOURCE_PATH = r"E:\\old\\tenant_a\\source.xlsx"\n'
        )
        self.assertEqual(
            find_nonportable_absolute_paths(code),
            [r"E:\old\tenant_a\source.xlsx"],
        )


if __name__ == "__main__":
    unittest.main()
