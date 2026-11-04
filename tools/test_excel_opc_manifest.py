"""Excel requires content type registration for the script-format manifest."""
import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from backend.utils.script_formats import read_explicit_formats, write_explicit_formats


class ManifestOpcTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'test.xlsx'
        book = Workbook(); book.active['A1'] = 123; book.save(self.path); book.close()

    def assert_manifest_valid(self, expected):
        with ZipFile(self.path) as z:
            self.assertIsNone(z.testzip())
            self.assertEqual(z.namelist().count('datamerge/explicit_number_formats.json'), 1)
            self.assertEqual(z.namelist().count('[Content_Types].xml'), 1)
            ns = {'c': 'http://schemas.openxmlformats.org/package/2006/content-types'}
            types = ET.fromstring(z.read('[Content_Types].xml'))
            matches = [n for n in types.findall('c:Override', ns)
                       if n.get('PartName') == '/datamerge/explicit_number_formats.json']
            self.assertEqual(len(matches), 1)
            self.assertEqual(matches[0].get('ContentType'), 'application/json')
        self.assertEqual(read_explicit_formats(self.path), expected)
        book = load_workbook(self.path, read_only=True)
        self.assertEqual(book.active['A1'].value, 123); book.close()

    def test_new_updated_and_idempotent_manifest(self):
        first = {'Sheet': {'A1': '0.00'}}
        second = {'Sheet': {'A1': '0.000'}}
        write_explicit_formats(self.path, first); self.assert_manifest_valid(first)
        original = self.path.read_bytes()
        write_explicit_formats(self.path, first); self.assertEqual(self.path.read_bytes(), original)
        write_explicit_formats(self.path, second); self.assert_manifest_valid(second)

    def test_repairs_existing_unregistered_manifest_even_if_content_is_unchanged(self):
        formats = {'Sheet': {'A1': '0.00'}}
        with ZipFile(self.path, 'a') as z:
            z.writestr('datamerge/explicit_number_formats.json', json.dumps(formats))
        write_explicit_formats(self.path, formats)
        self.assert_manifest_valid(formats)


if __name__ == '__main__': unittest.main()
