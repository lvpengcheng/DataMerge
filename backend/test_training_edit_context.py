import json

from backend.utils.training_edit_context import build_training_edit_context
from backend.ai_engine import precise_edit


def test_updated_inputs_override_old_description():
    text = build_training_edit_context({
        'source_structure_desc': 'old.xlsx', 'input_revision': 'revision-new',
        'expected_structure': {'sheets': {'部门汇总': {'headers': ['部门', '金额']}}},
        'validation_stale': True,
    }, {'source_structure': {'files': {'补贴.xlsx': {'sheets': {'明细': {'headers': ['工号', '补贴']}}}}}})
    assert '补贴.xlsx' in text and '部门汇总' in text and 'revision-new' in text
    assert 'old.xlsx' not in text
    assert '历史评分和差异已失效' in text


def test_precise_edit_adds_cleaning_and_assembly_with_call_sites(monkeypatch):
    original = 'def main(rows):\n    return rows\n\ndef unrelated():\n    return 42\n'
    replacement = (
        'def clean(rows):\n'
        '    return [dict(row, 工号=str(row["工号"]).strip()) for row in rows]\n\n'
        'def assemble(rows):\n'
        '    totals = {}\n'
        '    for row in rows:\n'
        '        key = row["工号"]\n'
        '        totals[key] = totals.get(key, 0) + row["补贴"]\n'
        '    return totals\n\n'
        'def main(rows):\n'
        '    return assemble(clean(rows))\n'
    )
    monkeypatch.setattr(precise_edit, '_call_provider', lambda *args: json.dumps({
        'edits': [{'find': 'def main(rows):\n    return rows\n', 'replace': replacement}]}))
    patched = precise_edit.run_precise_edit(object(), original,
        '工号去首尾空格并保留前导零，按工号汇总补贴，其余逻辑不变')
    assert patched and 'def unrelated():\n    return 42\n' in patched
    namespace = {}
    exec(patched, namespace)
    assert namespace['main']([{'工号': ' 001 ', '补贴': 2}, {'工号': '001', '补贴': 3}]) == {'001': 5}
