from pathlib import Path
from types import SimpleNamespace

import pytest

from backend.utils.template_resolver import extract_template_ref, resolve_template_path


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
