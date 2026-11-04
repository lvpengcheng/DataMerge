"""Global column declaration protocol, including the reported `.source` failure."""
import ast
from pathlib import Path

import pytest

from backend.utils.column_spec_schema import decode_column_specs
from backend.utils.column_spec_edit import canonicalize_column_specs, normalize_spec_updates, merge_column_spec_edit, prepare_column_contract_script
from backend.utils.column_contract import normalize_column_specs, read_column_specs, normalize_source_keys
from backend.utils.generated_code_check import check_generated_code
from backend.utils.multi_workbook import build_script, script_targets
from backend.utils.script_source_scope import source_scoped_generation


KEY = {'源_员工档案': {'工号': {'type': 'key'}}}
SHAPES = [
    KEY,
    {'源_员工档案': {'source': '员工档案.xlsx', '工号': {'type': 'key'}}},
    {'源_员工档案': {'source': '员工档案.xlsx', 'type': 'source', '工号': {'type': 'key'}}},
    {'源_员工档案': {'source': '员工档案.xlsx', 'type': 'source', 'columns': {'工号': {'type': 'key'}}}},
    {'源_员工档案': {'source': '员工档案.xlsx', 'columns': [{'name': '工号', 'type': 'key'}]}},
    [{'sheet': '源_员工档案', 'column': '工号', 'type': 'key'}],
    [{'sheet_name': '源_员工档案', 'columns': {'工号': {'type': 'key'}}, 'source': '员工档案.xlsx'}],
    {'sheet': '源_员工档案', 'columns': {'工号': {'type': 'key'}}, 'source': '员工档案.xlsx'},
    {'_COLUMN_SPECS': {'源_员工档案': {'工号': {'type': 'key'}}}},
    {'源_员工档案': {'工号': {'type': 'key', 'source': '员工档案.工号'}}},
]


@pytest.mark.parametrize('shape', SHAPES)
def test_every_boundary_uses_same_schema(shape):
    assert decode_column_specs(shape) == KEY
    assert normalize_spec_updates(shape) == KEY
    script = '_COLUMN_SPECS = ' + repr(shape) + '\ndef main():\n    return 1\n'
    fixed = canonicalize_column_specs(script)
    assert ast.literal_eval(ast.parse(fixed).body[0].value) == KEY
    assert check_generated_code(fixed) == []
    assert read_column_specs(script) == normalize_column_specs(KEY)
    assert normalize_column_specs(shape) == normalize_column_specs(KEY)
    assert normalize_column_specs(normalize_column_specs(shape)) == normalize_column_specs(KEY)
    assert merge_column_spec_edit(script, script, {}) == fixed


@pytest.mark.parametrize('alias,expected', [('str', 'text'), ('string', 'text'), ('integer', 'number'),
                                         ('float', 'number'), ('decimal', 'number'), ('日期', 'date')])
def test_type_aliases_are_consistent(alias, expected):
    spec = {'结果': {'字段': {'field_type': alias, 'number_format': 'General'}}}
    assert decode_column_specs(spec) == {'结果': {'字段': {'type': expected, 'format': 'General'}}}


def test_source_and_type_can_be_real_column_names():
    spec = {'结果': {'source': {'type': 'text'}, 'type': {'type': 'text'}, 'columns': {'type': 'text'}}}
    assert decode_column_specs(spec) == spec


def test_unknown_real_type_is_not_silently_dropped():
    with pytest.raises(ValueError, match='结果.金额.*source.*不支持'):
        decode_column_specs({'结果': {'金额': {'type': 'source', 'format': '0.00'}}})


def test_ambiguous_metadata_does_not_become_a_column():
    with pytest.raises(ValueError, match='只有来源元信息'):
        decode_column_specs({'源_员工档案': {'source': '员工档案.xlsx'}})


def test_alias_and_rule_conflicts_are_rejected():
    with pytest.raises(ValueError, match='冲突'):
        decode_column_specs({'结果': {'金额': {'type': 'text', 'dtype': 'number'}}})
    with pytest.raises(ValueError, match='冲突'):
        decode_column_specs([{'sheet': '结果', 'column': '工号', 'type': 'key'},
                             {'sheet': '结果', 'column': '工号', 'type': 'number'}])


def test_runtime_cache_is_not_mutated_with_metadata_shape():
    import pandas as pd
    data = {'员工档案': {'df': pd.DataFrame({'工号': [123]})}}
    fixed = normalize_source_keys(data, SHAPES[1])
    assert fixed['员工档案']['df']['工号'].iloc[0] == '123'
    assert data['员工档案']['df']['工号'].iloc[0] == 123


def test_generation_decorator_normalizes_all_returned_code():
    class Generator:
        @source_scoped_generation
        def generate_code(self, input_folder=None, rules_content='', expected_structure=None):
            return '_COLUMN_SPECS = ' + repr(SHAPES[1]) + '\ndef main(): return 1\n', 'response'
    code, response = Generator().generate_code()
    assert check_generated_code(code) == []
    assert ast.literal_eval(ast.parse(code).body[0].value) == KEY
    assert response == 'response'


def test_multi_workbook_preparation_preserves_per_file_rules_and_wrapper():
    a = '_COLUMN_SPECS = ' + repr(SHAPES[1]) + '\ndef main(): return 1\n'
    b = "_COLUMN_SPECS={'结果': {'columns': {'金额': {'dtype': 'decimal', 'number_format': '0.00'}}}}\ndef main(): return 2\n"
    bundle = build_script({'员工.xlsx': {'code': a}, '汇总.xlsx': {'code': b}})
    fixed = prepare_column_contract_script(bundle)
    assert check_generated_code(fixed) == []
    targets = script_targets(fixed)
    assert read_column_specs(targets['员工.xlsx']['code'])['源_员工档案']['工号']['type'] == 'key'
    assert read_column_specs(targets['汇总.xlsx']['code'])['结果']['金额']['type'] == 'number'
    assert prepare_column_contract_script(fixed) == fixed


def test_runtime_mutation_is_still_rejected_before_normalization():
    script = '_COLUMN_SPECS = ' + repr(SHAPES[1]) + '\n_COLUMN_SPECS.clear()\n'
    with pytest.raises(ValueError, match='运行时修改'):
        prepare_column_contract_script(script)


def test_duplicate_literal_keys_cannot_disappear_at_any_ast_boundary():
    code = "_COLUMN_SPECS = {'结果': {'金额': {'type': 'number', 'type': 'text'}}}\n"
    for reader in (canonicalize_column_specs, read_column_specs):
        with pytest.raises(ValueError):
            reader(code)
    assert any(issue['kind'] == 'ColumnSpecContract' for issue in check_generated_code(code))


def test_reading_a_dynamic_declaration_never_executes_it():
    with pytest.raises(ValueError, match='未执行代码'):
        read_column_specs("_COLUMN_SPECS = __import__('os').getenv('SECRET')")
