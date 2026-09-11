from pathlib import Path
import pytest


def test_structure_defaults_for_renamed_sheet_with_partial_columns():
    from backend.utils.compute_precheck import _suggest_structural_columns
    expected = [{'file': 'old.xlsx', 'sheet': '训练表',
                 'expected_columns': ['编号', '姓名', '部门', '金额', '备注']}]
    actual = [f'new.xlsx > 本月数据 > {col}' for col in ['编号', '姓名', '部门', '金额']]
    suggestions = _suggest_structural_columns(expected, actual)
    assert len(suggestions) == 4
    assert {s['suggested_path'] for s in suggestions} == set(actual)
    assert all(s['expected_path'].startswith('old.xlsx > 训练表 > ') for s in suggestions)


def test_structure_defaults_leave_ambiguous_sheets_manual():
    from backend.utils.compute_precheck import _suggest_structural_columns
    expected = [{'file': 'old.xlsx', 'sheet': '训练表', 'expected_columns': ['编号', '金额']}]
    actual = [f'new.xlsx > {sheet} > {col}' for sheet in ['本月', '上月'] for col in ['编号', '金额']]
    assert _suggest_structural_columns(expected, actual) == []


def test_structure_defaults_preserve_exact_sheet_over_duplicate_structure():
    from backend.utils.compute_precheck import _suggest_structural_columns
    expected = [{'file': 'same.xlsx', 'sheet': '本月', 'expected_columns': ['编号', '金额']}]
    actual = [f'same.xlsx > {sheet} > {col}' for sheet in ['本月', '上月'] for col in ['编号', '金额']]
    assert {s['suggested_path'] for s in _suggest_structural_columns(expected, actual)} == set(actual[:2])


def test_structure_defaults_do_not_reuse_source_sheet():
    from backend.utils.compute_precheck import _suggest_structural_columns
    expected = [{'file': 'old.xlsx', 'sheet': sheet, 'expected_columns': ['编号', '金额']}
                for sheet in ['一部', '二部']]
    actual = ['new.xlsx > 数据 > 编号', 'new.xlsx > 数据 > 金额']
    assert _suggest_structural_columns(expected, actual) == []


def test_structure_defaults_reject_unrelated_columns():
    from backend.utils.compute_precheck import _suggest_structural_columns
    expected = [{'file': 'old.xlsx', 'sheet': '数据', 'expected_columns': ['编号', '金额', '姓名']}]
    assert _suggest_structural_columns(expected, ['new.xlsx > 数据 > 编号', 'new.xlsx > 数据 > 其他']) == []


@pytest.mark.parametrize('mapping_failed', [False, True])
def test_empty_confirmation_still_validates_once_in_existing_worker(monkeypatch, tmp_path, mapping_failed):
    from backend.utils import compute_precheck as pre, source_auto_filler as filler
    (tmp_path / 'input.xlsx').touch()
    monkeypatch.setattr(filler, 'auto_rename_uploaded_by_combined_score', lambda **kw: ([], [], {}))
    monkeypatch.setattr(filler, 'auto_fill_missing_sources', lambda **kw: ([], []))
    monkeypatch.setattr(pre, '_check_history', lambda *a: None)
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *a: None)
    monkeypatch.setattr(pre, '_apply_confirmed_mapping', lambda *a: (_ for _ in ()).throw(AssertionError('empty mapping applied')))
    calls = []
    def match(*a):
        calls.append(a)
        return False, 'not matched', None, {'mapping_failed': mapping_failed, 'actual_paths': ['input.xlsx > Sheet > ID']}
    monkeypatch.setattr(pre, '_header_match_subprocess', match)
    monkeypatch.setattr(pre, '_collect_uploaded_columns', lambda *a: (_ for _ in ()).throw(AssertionError('parsed twice')))
    result = pre.precheck_compute(str(tmp_path), {'files': {}}, None, '', 'test', None, None,
        None, confirmed_mapping={'file_mapping': {}}, ai_provider_name='claude', in_worker=True)
    assert result.ok == mapping_failed and len(calls) == 1
    if mapping_failed:
        assert result._source_mapping_warning == 'not matched'
    assert result.actual_paths == ['input.xlsx > Sheet > ID']


def test_confirmed_rewrite_is_not_overwritten_by_original(monkeypatch, tmp_path):
    from backend.utils.compute_precheck import _apply_confirmed_mapping
    from backend.utils.fast_header_matcher import FastHeaderMatcher
    original = tmp_path / 'new.xlsx'
    original.write_bytes(b'old headers')
    def rewrite(info, directory):
        result = Path(directory) / info['expected_file']
        result.write_bytes(b'mapped headers')
        return str(result)
    monkeypatch.setattr(FastHeaderMatcher, 'rewrite_excel', rewrite)
    _apply_confirmed_mapping(str(tmp_path), {}, {'new.xlsx': {'expected_file': 'expected.xlsx'}})
    assert (tmp_path / 'expected.xlsx').read_bytes() == b'mapped headers'
    assert not original.exists()
