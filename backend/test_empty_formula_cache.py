import zipfile
from backend.utils.excel_comparator import inspect_formula_cache


def test_empty_string_is_calculated_but_absent_numeric_cache_is_missing(tmp_path):
    path = tmp_path / 'result.xlsx'
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('xl/workbook.xml', '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Summary" sheetId="1" r:id="rId1"/></sheets></workbook>''')
        archive.writestr('xl/_rels/workbook.xml.rels', '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>''')
        archive.writestr('xl/worksheets/sheet1.xml', '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">
        <c r="A1" t="str"><f>IF(1,"",1)</f><v/></c>
        <c r="B1"><f>1+1</f><v/></c>
        <c r="C1" t="str"><f>IF(1,"",1)</f></c>
        <c r="D1"><f>0</f><v>0</v></c>
        </row></sheetData></worksheet>''')
    report = inspect_formula_cache(path)
    assert report['formula_count'] == 4
    assert report['empty_cache_count'] == 2
    assert [s['cell'] for s in report['empty_cache_samples']] == ['B1', 'C1']
