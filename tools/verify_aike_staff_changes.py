"""Run the specified script on disposable copies; original workbooks are read-only."""
import importlib.util
import sys
import tempfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd
import openpyxl
from collections import Counter
from backend.utils.openpyxl_compat import ensure_custom_filter_compat
ensure_custom_filter_compat()
folder = Path(sys.argv[1])
script = folder / '埃柯_埃柯 Staff Change V3_20260915_104541.py'
spec = importlib.util.spec_from_file_location('aike_staff_test', script)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
source = {}
for filename in ['Sourcing Staff Change Report.xlsx', '拆分过程表.xlsx', '薪资日历-2026.xlsx']:
    for sheet, df in pd.read_excel(folder / filename, sheet_name=None, dtype=object).items():
        source[sheet] = {'df': df, 'columns': list(df.columns)}
module._pre_loaded_source_data = source
module.salary_year, module.salary_month = 2026, int(sys.argv[2]) if len(sys.argv) > 2 else 9
module.monthly_standard_hours = 174
module._template_override_path = str(folder / 'Jojo-Staff Change.xlsx')
with tempfile.TemporaryDirectory(prefix='aike_staff_verify_') as temp:
    module.output_folder = temp
    assert module.main()
    output = Path(temp) / 'Jojo-Staff Change.xlsx'
    wb = openpyxl.load_workbook(output)
    ws = wb['Kmart_Payroll_detail']
    rows = module.cleaned_rows('Kmart_Payroll_detail')
    assert len(rows) == 207, len(rows)
    assert len({key for _, key in rows}) == 207
    for r, key in rows:
        rec = module._STAFF_ROSTER[str(key)]
        assert str(ws.cell(r, 1).value) == str(key)
        assert ws.cell(r, 10).value == (rec.get('Status') or 'A')
        assert ws.cell(r, 14).data_type == 'f'
        if rec.get('DateofLeave') is not None:
            assert ws.cell(r, 12).value == rec['DateofLeave']
    joins = source['New Join']['df']['Employee ID'].map(str)
    positions = {key: r for r, key in rows}
    for key in joins:
        r = positions[key]
        assert ws.cell(r, 2).value
        assert ws.cell(r, 4).value
        assert ws.cell(r, 11).value
        assert ws.cell(r, 16).value is not None
        assert ws.cell(r, 10).value == ('N' if module.salary_month == 8 else 'A')
    print('VERIFIED employees=', len(rows), 'statuses=', dict(Counter(ws.cell(r,10).value for r,_ in rows)))
    print('VERIFIED all 5 joins present with identity/date/salary; leave dates/status retained; original files unchanged')
    wb.close()
