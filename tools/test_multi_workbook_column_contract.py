import json

import pytest

from backend.ai_engine import precise_edit
from backend.utils.column_contract import read_column_specs
from backend.utils.column_spec_edit import normalize_spec_updates, canonicalize_column_specs, merge_column_spec_edit, detach_generated_specs
from backend.utils.generated_code_check import check_generated_code
from backend.utils.multi_workbook import build_script, script_targets, precise_edit_workbooks


def child(fmt='yyyy-m-d', bad=False):
    return (f"_COLUMN_SPECS = {{'明细': {{'日期': {{'type': 'date', 'format': {fmt!r}}}}}}}\n"
            "def main():\n"
            "    rows = {}\n"
            "    if callable(globals().get('clean_source_data')):\n"
            "        rows = clean_source_data(rows)\n" +
            ("    missing_helper()\n" if bad else '') +
            "    return rows\n")


def test_multi_workbook_declarations_and_optional_cleaners_are_independent():
    code = build_script({'工资.xlsx': {'code': child()}, '汇总.xlsx': {'code': child('yyyy/mm/dd')}})
    assert check_generated_code(code) == []
    targets = script_targets(code)
    assert read_column_specs(targets['工资.xlsx']['code'])['明细']['日期']['format'] == 'yyyy-m-d'
    assert read_column_specs(targets['汇总.xlsx']['code'])['明细']['日期']['format'] == 'yyyy/mm/dd'


def test_sibling_definition_does_not_hide_child_missing_helper():
    good = child() + '\ndef missing_helper():\n    return 1\n'
    code = build_script({'工资.xlsx': {'code': good}, '汇总.xlsx': {'code': child(bad=True)}})
    issues = check_generated_code(code)
    assert len(issues) == 1
    assert '[汇总.xlsx]' in issues[0]['message'] and 'missing_helper' in issues[0]['message']


def test_runtime_contract_mutation_still_rejected_inside_one_target():
    code = build_script({'工资.xlsx': {'code': child() + '\n_COLUMN_SPECS.clear()\n'},
                         '汇总.xlsx': {'code': child()}})
    issues = check_generated_code(code)
    assert any(issue['kind'] == 'ColumnSpecContract' and '[工资.xlsx]' in issue['message'] for issue in issues)


@pytest.mark.parametrize('guard', ["callable(globals().get('clean_source_data'))", "'clean_source_data' in globals()"])
def test_old_optional_cleaning_guard_is_valid(guard):
    code = f'def main(rows):\n    if {guard}:\n        rows = clean_source_data(rows)\n    return rows\n'
    assert check_generated_code(code) == []


def test_unguarded_missing_cleaner_still_fails():
    issues = check_generated_code('def main(rows):\n    return clean_source_data(rows)\n')
    assert any('clean_source_data' in issue['message'] for issue in issues)


def test_delayed_nested_call_is_not_covered_by_outer_guard():
    code = "if callable(globals().get('clean_source_data')):\n    def delayed():\n        clean_source_data({})\n"
    assert check_generated_code(code)


@pytest.mark.parametrize('payload', [
    {'日期': {'type': 'date', 'format': 'yyyy-m-d'}},
    {'*': {'日期': {'type': 'date', 'number_format': 'yyyy-m-d'}}},
    {'_COLUMN_SPECS': {'*': {'日期': {'type': 'date', 'format': 'yyyy-m-d'}}}},
])
def test_common_ai_shapes_are_normalized(payload):
    assert normalize_spec_updates(payload) == {'*': {'日期': {'type': 'date', 'format': 'yyyy-m-d'}}}


def test_workbook_nested_response_is_routed_to_current_target_only(monkeypatch):
    original = build_script({'工资.xlsx': {'code': child()}, '汇总.xlsx': {'code': child()}})
    payload = {'edits': [], 'column_specs': {
        '工资.xlsx': {'明细': {'日期': {'format': 'yyyy/mm/dd'}}},
        '汇总.xlsx': {'明细': {'日期': {'format': 'yy-m-d'}}}}}
    monkeypatch.setattr(precise_edit, '_call_provider', lambda *args: json.dumps(payload))
    edited = precise_edit_workbooks(object(), original, '修改日期格式')
    assert edited and check_generated_code(edited) == []
    targets = script_targets(edited)
    assert read_column_specs(targets['工资.xlsx']['code'])['明细']['日期']['format'] == 'yyyy/mm/dd'
    assert read_column_specs(targets['汇总.xlsx']['code'])['明细']['日期']['format'] == 'yy-m-d'


def test_named_target_edit_leaves_other_workbook_untouched(monkeypatch):
    original = build_script({'工资.xlsx': {'code': child()}, '汇总.xlsx': {'code': child()}})
    monkeypatch.setattr(precise_edit, '_call_provider', lambda *args: json.dumps({
        'edits': [], 'column_specs': {'日期': {'format': 'yyyy/mm/dd'}}}))
    edited = precise_edit_workbooks(object(), original, '工资.xlsx 的日期格式改为 yyyy/mm/dd')
    assert edited
    assert script_targets(edited)['汇总.xlsx']['code'] == script_targets(original)['汇总.xlsx']['code']


def test_unknown_file_nested_response_is_rejected():
    with pytest.raises(ValueError, match='无法对应当前交付文件'):
        normalize_spec_updates({'其他.xlsx': {'明细': {'日期': 'date'}}}, '工资.xlsx')


def test_empty_sheet_is_still_rejected_with_actionable_example():
    with pytest.raises(ValueError, match='入职日期'):
        normalize_spec_updates({'': {'日期': {'format': 'yyyy-m-d'}}})


def test_sheet_name_ending_in_xlsx_is_not_mistaken_for_file_wrapper():
    values = {'明细.xlsx': {'日期': {'format': 'yyyy-m-d'}}}
    assert normalize_spec_updates(values) == values


def test_legacy_dictionary_bundle_is_checked_per_workbook():
    targets = {'工资.xlsx': {'code': child()}, '汇总.xlsx': {'code': child()}}
    code = ('from backend.utils.multi_workbook import execute_workbooks\n'
            + 'MULTI_WORKBOOK_SCRIPTS = ' + repr(targets) + '\n'
            + 'def main():\n    return execute_workbooks(MULTI_WORKBOOK_SCRIPTS, globals())\n')
    assert check_generated_code(code) == []


def test_multi_workbook_exec_and_export_keep_separate_column_formats(tmp_path, monkeypatch):
    import openpyxl
    from backend.utils.output_postprocess import finalize_output_workbook
    monkeypatch.setenv('COLUMN_CONTRACT_MODE', 'unified')
    targets = {}
    for name, fmt in [('工资.xlsx', 'yyyy-m-d'), ('汇总.xlsx', 'yyyy/mm/dd')]:
        code = child(fmt).split('def main():')[0] + '''
def main(output_folder):
    from pathlib import Path
    from datetime import datetime
    import openpyxl
    rows = {}
    if callable(globals().get('clean_source_data')):
        rows = clean_source_data(rows)
    wb = openpyxl.Workbook()
    wb.active.title = '明细'
    wb.active.append(['日期'])
    wb.active.append([datetime(2026, 1, 1)])
    wb.save(Path(output_folder) / 'result.xlsx')
    wb.close()
'''
        targets[name] = {'code': code}
    bundle = build_script(targets)
    assert check_generated_code(bundle) == []
    namespace = {'__name__': 'training_bundle', 'input_folder': str(tmp_path), 'output_folder': str(tmp_path)}
    exec(compile(bundle, '<bundle>', 'exec'), namespace)
    namespace['main']()
    for name, fmt in [('工资.xlsx', 'yyyy-m-d'), ('汇总.xlsx', 'yyyy/mm/dd')]:
        path = tmp_path / name
        finalize_output_workbook(path, script_code=targets[name]['code'])
        wb = openpyxl.load_workbook(path)
        try:
            assert wb['明细']['A2'].number_format == fmt
            assert wb['明细']['A2'].value.year == 2026
        finally:
            wb.close()


def test_real_ai_record_list_is_canonicalized_without_losing_rules():
    code = "_COLUMN_SPECS = [{'sheet': '源_员工档案', 'column': '工号', 'type': 'key'}, {'sheet': '结果', 'column': '部门', 'type': 'key'}]\ndef main():\n    return 1\n"
    fixed = canonicalize_column_specs(code)
    assert check_generated_code(fixed) == []
    assert read_column_specs(fixed)['源_员工档案']['工号']['type'] == 'key'
    assert read_column_specs(fixed)['结果']['部门']['type'] == 'key'
    assert fixed.endswith('def main():\n    return 1\n')
    assert merge_column_spec_edit(code, code, {}) == fixed


def test_ai_can_repair_invalid_original_list_and_add_compatible_format():
    old = "_COLUMN_SPECS = [{'sheet': '结果', 'column': '工号', 'type': 'key'}]\nx = 1\n"
    edited = "_COLUMN_SPECS = {'结果': {'工号': {'type': 'key', 'format': '@'}}}\nx = 1\n"
    assert merge_column_spec_edit(old, edited, {}) == edited


def test_repairing_list_cannot_drop_or_change_existing_keys():
    old = "_COLUMN_SPECS = [{'sheet': '结果', 'column': '工号', 'type': 'key'}]\nx = 1\n"
    for changed in ["_COLUMN_SPECS = {}\nx = 1\n", "_COLUMN_SPECS = {'结果': {'工号': 'number'}}\nx = 1\n"]:
        with pytest.raises(ValueError, match='删除或改写'):
            merge_column_spec_edit(old, changed, {})


def test_generation_detaches_record_lists_as_valid_module_dicts():
    code = "_COLUMN_SPECS = [{'sheet': '结果', 'column': '工号', 'type': 'key'}]\ndef fill():\n    pass\n"
    prefix, body = detach_generated_specs(code)
    assert read_column_specs(prefix)['结果']['工号']['type'] == 'key'
    assert body == 'def fill():\n    pass\n'


def test_conflicting_records_are_not_silently_merged():
    values = [{'sheet': '结果', 'column': '工号', 'type': 'key'},
              {'sheet': '结果', 'column': '工号', 'type': 'number'}]
    with pytest.raises(ValueError, match='冲突'):
        normalize_spec_updates(values)


def test_target_execution_error_identifies_workbook_and_preserves_cause(tmp_path):
    code = build_script({'部门薪资汇总.xlsx': {'code':
        "def main():\n    raise KeyError('源_员工档案')\n"}})
    namespace = {'__name__': 'bundle', 'output_folder': str(tmp_path), 'input_folder': str(tmp_path)}
    exec(compile(code, '<bundle>', 'exec'), namespace)
    with pytest.raises(ValueError, match='部门薪资汇总.xlsx') as error:
        namespace['main']()
    assert isinstance(error.value.__cause__, KeyError)
    assert not list(tmp_path.glob('*.xlsx'))


def test_legacy_record_conversion_preserves_future_import_and_business_ast():
    import ast
    code = '"""说明"""\nfrom __future__ import annotations\n_COLUMN_SPECS = [{"sheet":"结果","column":"工号","type":"key"}]\ndef main():\n    return 1\n'
    fixed = canonicalize_column_specs(code)
    compile(fixed, '<fixed>', 'exec')
    before, after = ast.parse(code), ast.parse(fixed)
    before.body = [n for n in before.body if not isinstance(n, ast.Assign)]
    after.body = [n for n in after.body if not isinstance(n, ast.Assign)]
    assert ast.dump(before) == ast.dump(after)
