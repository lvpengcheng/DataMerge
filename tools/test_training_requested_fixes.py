from pathlib import Path
from types import SimpleNamespace

import pytest

from backend.utils.template_resolver import extract_template_ref, resolve_template_path


def test_mapped_compute_source_runs_with_same_input_as_training(tmp_path, monkeypatch):
    from backend.api.training_chat import _load_full_source_data
    from backend.utils.fast_header_matcher import FastHeaderMatcher
    from backend.utils.script_entry import invoke_script_main
    from excel_parser import ExcelRegion, IntelligentExcelParser, SheetData

    source_dir = tmp_path / 'training'
    source_dir.mkdir()
    (source_dir / '员工.xlsx').touch()
    formats = {'B': '@'}
    rows = [{'A': None, 'B': 123}, {'A': None, 'B': '00456'}]
    training_sheet = SheetData('员工明细', [ExcelRegion(
        head_data={'序号': 'A', 'PSID': 'B'}, data=rows, column_formats=formats)])
    monkeypatch.setattr(IntelligentExcelParser, 'parse_excel_file', lambda *args, **kwargs: [training_sheet])
    training = _load_full_source_data(str(source_dir), reserved_sheet_names={'员工明细'})

    upload = str(tmp_path / '本月.xlsx')
    upload_sheet = SheetData('当月', [ExcelRegion(
        head_data={'行号': 'A', '员工代码': 'B'}, data=rows, column_formats=formats)])
    mapping = {'本月.xlsx': {
        'expected_file': '员工.xlsx', 'file_path': upload,
        'sheet_mapping': {'当月': '员工明细'},
        'header_mapping_by_sheet': {'当月': {'行号': '序号', '员工代码': 'PSID'}},
    }}
    compute = FastHeaderMatcher()._build_pre_loaded_from_memory(
        mapping, {(upload, '当月'): upload_sheet}, reserved_sheet_names={'员工明细'})

    namespace = {}
    exec('''
def load_source_data():
    raise AssertionError('preloaded sheets should not be read again')

def main():
    sources = load_source_data()
    assert sources is _pre_loaded_source_data
    entry = sources['员工_员工明细']
    return entry['columns'], entry['column_formats'], entry['df'].to_dict('list')
''', namespace)
    expected = (['序号', 'PSID'], {'PSID': '@'}, {'序号': [1, 2], 'PSID': ['123', '00456']})
    assert invoke_script_main(namespace['main'], {'_pre_loaded_source_data': training}) == expected
    assert invoke_script_main(namespace['main'], {'_pre_loaded_source_data': compute}) == expected
    assert '_pre_loaded_source_data' not in namespace
    assert compute['员工_员工明细']['source_origins'][0]['sheet_name'] == '当月'
    assert compute['员工_员工明细']['df']['PSID'].iloc[0] == 123


def test_default_template_ignores_stale_comments(tmp_path):
    directory = tmp_path / 'tenants' / 'demo' / 'templates'
    directory.mkdir(parents=True)
    template = directory / '435.xlsx'
    template.write_bytes(b'template')
    code = "# TEMPLATE_NAME = '55.xlsx'\n# TEMPLATE_PATH = '/old/55.xlsx'\nTEMPLATE_NAME = '435.xlsx'\nTEMPLATE_PATH = '/old/435.xlsx'"
    assert extract_template_ref(code) == ('435.xlsx', None, '/old/435.xlsx')
    assert resolve_template_path(tenant_id='demo', script_code=code, project_root=tmp_path) == str(template)


@pytest.mark.parametrize('failure', ['timeout', 'sheet_key', 'exception'])
def test_training_comparison_failure_keeps_output_and_selected_sheets(tmp_path, monkeypatch, failure):
    from backend.api.training_chat import _persist_iteration_files, _run_single_iteration
    from backend.sandbox.code_sandbox import CodeSandbox
    from backend.storage.storage_manager import StorageManager
    from backend.utils import training_validation, subprocess_runner
    from openpyxl import Workbook
    sources = tmp_path / 'sources'
    sources.mkdir()
    def execute(self, code, env):
        wb = Workbook()
        wb.active.title = '目标'
        wb.active.append(['工号', '金额'])
        wb.active.append(['001', 100])
        wb.save(Path(env['output_folder']) / 'result.xlsx')
        return {'success': True}
    monkeypatch.setattr(CodeSandbox, 'execute_script', execute)
    monkeypatch.setattr(training_validation, 'finalize_training_output', lambda *args: {})
    captured = {}
    def compare_worker(*args, **kwargs):
        captured.update(kwargs['kwargs'])
        if failure == 'timeout':
            return SimpleNamespace(success=False, timed_out=True, killed=True, killed_by_memory=False, error='timeout')
        if failure == 'exception':
            raise RuntimeError('对比进程启动失败')
        Path(captured['output_file']).write_bytes(b'partial diff')
        return SimpleNamespace(success=True, result={
            'comparison_complete': False, 'total_cells': 20, 'matched_cells': 20,
            'error': "部分 Sheet 对比失败: {'Payroll history': '缺少 PSID'}",
        })
    monkeypatch.setattr(subprocess_runner, 'run_in_subprocess', compare_worker)
    result = _run_single_iteration(1, 'pass', 'demo', str(sources), 'expected.xlsx', 1,
                                   expected_structure={'sheets': {'目标': {}, 'Payroll history': {}}},
                                   target_sheets=['目标'])
    import shutil
    try:
        assert result['success']
        assert result['comparison_complete'] is False
        reason = {'timeout': '超时', 'sheet_key': '缺少 PSID', 'exception': '对比进程启动失败'}
        assert reason[failure] in result['comparison_warning']
        assert result['accuracy'] == 0
        assert result['diff_file'] is None
        assert captured['selected_sheets'] == ['目标']
        monkeypatch.setattr(StorageManager, 'get_tenant_dir', lambda *args: tmp_path)
        files = _persist_iteration_files('demo', 1, 1, 'pass', result)
        assert Path(files['output_file']).is_file()
        assert 'diff_file' not in files
    finally:
        shutil.rmtree(Path(result['output_dir']).parent)
