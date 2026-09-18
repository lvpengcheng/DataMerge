import pytest


def test_column_suggestions_cannot_reuse_an_already_mapped_sheet():
    from backend.utils.compute_precheck import _suggest_structural_columns
    entries = [{'file':'上月.xlsx', 'sheet':'第二批', 'expected_columns':['编号','金额']}]
    actual = ['上传.xlsx > 第一批 > 编号', '上传.xlsx > 第一批 > 金额']
    mapping = {'上传.xlsx': {'expected_file':'上月.xlsx', 'sheet_mapping':{'第一批':'第一批'}}}
    assert _suggest_structural_columns(entries, actual, mapping) == []


def test_column_suggestions_stay_inside_the_selected_sheet_and_file():
    from backend.utils.compute_precheck import _suggest_structural_columns
    entries = [{'file':'上月.xlsx', 'sheet':'第二批', 'expected_columns':['编号','金额']}]
    actual = [f'上传.xlsx > {sheet} > {col}' for sheet in ['第一批','第二批'] for col in ['编号','金额']]
    mapping = {'上传.xlsx': {'expected_file':'上月.xlsx', 'sheet_mapping':{'第二批':'第二批'}}}
    result = _suggest_structural_columns(entries, actual, mapping)
    assert {row['suggested_path'] for row in result} == set(actual[2:])
    assert _suggest_structural_columns(entries, actual, {}, {'上传.xlsx':'当月.xlsx'}) == []


def test_refresh_confirmation_never_dispatches_calculation(tmp_path, monkeypatch):
    """Exercise the actual endpoint body without initializing the entire web app."""
    import ast
    import asyncio
    import json
    from pathlib import Path
    from backend.utils import compute_ingest
    from backend.utils.compute_precheck import PrecheckResult
    source = Path(__file__).parent / 'app' / 'main.py'
    tree = ast.parse(source.read_text(encoding='utf-8'))
    endpoint = next(n for n in tree.body if isinstance(n, ast.AsyncFunctionDef)
                    and n.name == 'compute_session_confirm')
    endpoint.decorator_list = []
    endpoint.returns = None
    for arg in endpoint.args.args:
        arg.annotation = None
    endpoint.args.defaults = [ast.Constant(None) for _ in endpoint.args.defaults]
    calls = []
    async def dispatch(*args, **kwargs):
        calls.append(True)
        return {'started': True}
    namespace = {'asyncio': asyncio,
                 '_get_compute_session': lambda *args: {'temp_dir': str(tmp_path), 'params': {}},
                 '_compute_pending_payload': lambda pc, session: {'session_id': session},
                 '_dispatch_compute_task': dispatch,
                 'logger': __import__('logging').getLogger(__name__)}
    module = ast.fix_missing_locations(ast.Module(body=[endpoint], type_ignores=[]))
    exec(compile(module, str(source), 'exec'), namespace)
    monkeypatch.setattr(compute_ingest, 'ingest_ready', lambda *args: True)
    monkeypatch.setattr(compute_ingest, 'read_meta', lambda *args: object())
    monkeypatch.setattr(compute_ingest, 'resolve_with_confirmations', lambda *args, **kwargs: PrecheckResult(ok=True))
    fn = namespace['compute_session_confirm']
    refreshed = asyncio.run(fn('s1', {'refresh_only': True, 'confirmed_target_map': {'当月2': '202608(2)'}}))
    assert refreshed['mapping_refreshed'] is True
    assert calls == []
    saved = json.loads((tmp_path / '_confirmations.json').read_text(encoding='utf-8'))
    assert saved['confirmed_target_map'] == {'当月2': '202608(2)'}
    assert 'refresh_only' not in saved
    assert asyncio.run(fn('s1', {})) == {'started': True}
    assert calls == [True]


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
    """空确认封装不算通过；源文件只解析一次；诊断可用时不阻断、只记 warning。"""
    from backend.utils import compute_precheck as pre, source_auto_filler as filler
    from backend.utils.fast_header_matcher import FastHeaderMatcher
    (tmp_path / 'input.xlsx').touch()
    monkeypatch.setattr(filler, 'auto_rename_uploaded_by_combined_score', lambda **kw: ([], [], {}))
    monkeypatch.setattr(filler, 'auto_fill_missing_sources', lambda **kw: ([], []))
    monkeypatch.setattr(pre, '_check_history', lambda *a: None)
    monkeypatch.setattr(pre, '_check_target_sheets', lambda *a: None)

    parses = []
    def parse_inputs(self, input_files, manual_headers=None, multi_sheet_source=False):
        parses.append(list(input_files))
        return [{'file_name': 'input.xlsx', 'file_path': str(tmp_path / 'input.xlsx'),
                 'sheet_name': 'Sheet', 'headers': {'ID': 'A'}}], {}
    monkeypatch.setattr(FastHeaderMatcher, 'parse_inputs', parse_inputs)
    monkeypatch.setattr(FastHeaderMatcher, 'match_headers_only', lambda *a: {
        'success': False, 'error': 'not matched',
        'diagnostics': {'mapping_failed': mapping_failed,
                        'actual_paths': ['input.xlsx > Sheet > ID']}})
    monkeypatch.setattr(pre, '_collect_uploaded_columns', lambda *a: (_ for _ in ()).throw(AssertionError('parsed twice')))

    result = pre.precheck_compute(str(tmp_path), {'files': {'train.xlsx': {'sheets': {'S': {'headers': {'编号': 'A'}}}}}},
        None, '', 'test', None, None, None, confirmed_mapping={'file_mapping': {}},
        ai_provider_name='claude', in_worker=True)
    assert result.ok == mapping_failed and len(parses) == 1
    if mapping_failed:
        assert result._source_mapping_warning == 'not matched'
    assert result.actual_paths == ['input.xlsx > Sheet > ID']


def test_confirmed_mapping_is_composed_in_memory_without_rewriting_files(tmp_path):
    """确认的列映射在内存里复合回真实坐标（不再重写文件、不再二次解析）。"""
    from backend.utils.compute_ingest import _build_virtual_sheets, _compose_file_mapping
    original = tmp_path / 'new.xlsx'
    original.write_bytes(b'old headers')
    input_sheets = [{'file_name': 'new.xlsx', 'file_path': str(original),
                     'sheet_name': '本月数据', 'headers': {'编号': 'A', '金额': 'B'}}]
    confirmed = {'new.xlsx': {'expected_file': 'expected.xlsx',
                              'sheet_mapping': {'本月数据': '训练表'},
                              'header_mapping': {'编号': '工号'}}}

    virtual, xlate = _build_virtual_sheets(input_sheets, {}, confirmed)
    assert virtual[0]['file_name'] == 'expected.xlsx' and virtual[0]['sheet_name'] == '训练表'
    assert set(virtual[0]['headers']) == {'工号', '金额'}

    # 匹配器在虚拟坐标上给出恒等映射 → 复合回真实文件名/sheet/列
    match_fm = {'expected.xlsx': {'expected_file': 'expected.xlsx', 'file_path': str(original),
                                  'sheet_mapping': {'训练表': '训练表'},
                                  'header_mapping': {'工号': '工号', '金额': '金额'}}}
    composed = _compose_file_mapping(match_fm, xlate)
    assert composed['new.xlsx']['expected_file'] == 'expected.xlsx'
    assert composed['new.xlsx']['sheet_mapping'] == {'本月数据': '训练表'}
    assert composed['new.xlsx']['header_mapping'] == {'编号': '工号', '金额': '金额'}
    assert composed['new.xlsx']['needs_rewrite'] is True  # 列改名在内存里生效
    assert original.read_bytes() == b'old headers'  # 源文件一个字节都没动
    assert not (tmp_path / 'expected.xlsx').exists()
