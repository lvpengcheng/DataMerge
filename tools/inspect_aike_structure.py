"""Read-only XML structure inspection; do not disclose employee records."""
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import sys
sys.stdout.reconfigure(encoding='utf-8')
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for path in Path(sys.argv[1]).glob('*.xlsx'):
    with ZipFile(path) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(n.itertext()) for n in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        print('\nFILE', path.name)
        for sheet in ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', ns):
            if sheet.attrib['name'] != 'Kmart_Payroll_detail':
                continue
            target = rels[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
            name = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            root = ET.fromstring(z.read(name))
            rows = root.findall('m:sheetData/m:row', ns)
            def value(c):
                if c is None: return ''
                v = c.find('m:v', ns)
                if c.attrib.get('t') == 's' and v is not None: return strings[int(v.text)]
                if c.attrib.get('t') == 'inlineStr': return ''.join(c.find('m:is', ns).itertext())
                return v.text if v is not None else ''
            populated = []
            for row in rows:
                r = row.attrib['r']
                v = value(row.find('m:c[@r="A'+r+'"]', ns))
                if v: populated.append(int(r))
            merges = [m.attrib['ref'] for m in root.findall('m:mergeCells/m:mergeCell', ns)]
            print(sheet.attrib['name'], 'rows=',len(rows), 'A_nonempty=',len(populated), 'A_last=',populated[-8:], 'merges=',merges[:25])
            print('headers=', [(c.attrib['r'], value(c)) for c in rows[0]][:42] if rows else [])
            if path.name.startswith('薪资报告'):
                for row in rows[1:4]:
                    print('HEADER ROW', row.attrib['r'], [(c.attrib['r'], value(c)) for c in row if value(c)])
