import copy
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import pickle
import shutil

import pytest

from backend.solutions import environment
from backend.solutions.package import digest_file
from backend.utils.workbook_staging import workbook_staging, install_legacy_cleanup


def runtime():
    return dict(python='3.11.17', implementation='CPython', platform='linux', machine='x86_64',
                python_binary='linux-binary', dotnet=['Microsoft.NETCore.App 9.0.21'],
                packages={'pandas': '2.1.4', 'python_dateutil': '2.9.0.post0', 'uvloop': '0.23.0'}, native={})


def test_windows_platform_and_host_tools_are_not_calculation_identity(tmp_path, monkeypatch):
    expected = runtime()
    actual = copy.deepcopy(expected)
    actual.update(platform='win32', machine='AMD64', python_binary='windows-binary')
    actual['packages'].pop('uvloop')
    actual['packages']['colorama'] = '0.4.6'
    actual['dotnet'].append('Microsoft.AspNetCore.App 9.0.21')
    assert environment.compatibility_errors(expected, actual) == []
    monkeypatch.setattr(environment, 'fingerprint', lambda root: actual)
    assert environment.runtime_python(expected, tmp_path) == sys.executable


@pytest.mark.parametrize('key,value,description', [
    ('python', '3.12.0', 'python'), ('machine', 'ARM64', 'CPU'),
    ('dotnet', ['Microsoft.NETCore.App 8.0.0'], '.NET'),
    ('packages', {'pandas': '2.2.0'}, 'pandas'),
])
def test_incompatible_calculation_runtime_is_rejected(key, value, description):
    actual = runtime()
    actual[key] = value
    assert description in ';'.join(environment.compatibility_errors(runtime(), actual))


def test_native_adaptation_preserves_managed_code_and_rejects_wrong_version(tmp_path, monkeypatch):
    monkeypatch.setattr(sys, 'platform', 'win32')
    frozen, host = tmp_path / 'frozen', tmp_path / 'host'
    for root in (frozen, host):
        (root / 'libs').mkdir(parents=True)
        (root / 'libs/SkiaSharp.dll').write_bytes(b'managed')
    (host / 'libs/libSkiaSharp.dll').write_bytes(b'windows-native')
    expected = {'native': {'SkiaSharp.dll': digest_file(frozen / 'libs/SkiaSharp.dll')}}
    environment.prepare_platform_runtime(expected, frozen, host)
    assert (frozen / 'libs/libSkiaSharp.dll').read_bytes() == b'windows-native'
    assert digest_file(frozen / 'libs/SkiaSharp.dll') == expected['native']['SkiaSharp.dll']
    (frozen / 'libs/libSkiaSharp.dll').unlink()
    (host / 'libs/SkiaSharp.dll').write_bytes(b'wrong-version')
    with pytest.raises(ValueError, match='版本不一致'):
        environment.prepare_platform_runtime(expected, frozen, host)


@pytest.mark.skipif(os.name != 'nt', reason='Windows sharing semantics')
def test_windows_locked_output_does_not_fail_success_or_mask_business_error(tmp_path, caplog):
    from ctypes import wintypes
    create = ctypes.windll.kernel32.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    close = ctypes.windll.kernel32.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    for fail in (False, True):
        handle = None
        def run():
            nonlocal handle
            with workbook_staging(dir=tmp_path) as staging:
                file = Path(staging) / '工资.xlsx'
                file.write_bytes(b'completed')
                handle = create(str(file), 0x80000000, 1, None, 3, 0, None)
                assert handle != ctypes.c_void_p(-1).value
                (tmp_path / 'delivered.xlsx').write_bytes(file.read_bytes())
                if fail:
                    raise ValueError('business error')
        try:
            if fail:
                with pytest.raises(ValueError, match='business error'):
                    run()
            else:
                run()
            assert (tmp_path / 'delivered.xlsx').read_bytes() == b'completed'
            assert '仍被占用' in caplog.text
        finally:
            if handle:
                close(handle)


def test_legacy_cleanup_does_not_modify_global_tempfile(tmp_path):
    import tempfile
    from types import SimpleNamespace
    original = tempfile.TemporaryDirectory
    module = SimpleNamespace(tempfile=tempfile)
    install_legacy_cleanup(module)
    assert tempfile.TemporaryDirectory is original
    with module.tempfile.TemporaryDirectory(prefix='multi_', dir=tmp_path) as folder:
        assert Path(folder).is_dir()
    assert not Path(folder).exists()


def test_iis_calculation_constraints_match_docker_lock():
    root = Path(__file__).resolve().parents[1]
    def requirements(path):
        return {line.split('==')[0].replace('_', '-'): line.split('==')[1]
                for line in path.read_text().splitlines() if '==' in line and not line.startswith('#')}
    docker = requirements(root / 'requirements.docker.lock')
    iis = requirements(root / 'tools/iis/calculation-constraints.txt')
    assert set(iis) == environment.CALCULATION_PACKAGES
    assert all(docker[name] == version for name, version in iis.items())


@pytest.mark.skipif(os.name != 'nt' or not os.environ.get('DOTNET_ROOT'), reason='IIS runtime integration')
def test_linux_package_runs_in_windows_with_host_bootstrap_and_frozen_code(tmp_path, monkeypatch):
    from backend.solutions import execution, package
    repo = Path(__file__).resolve().parents[1]
    frozen = tmp_path / 'package'
    runtime_dir = frozen / 'runtime'
    def write(relative, text):
        path = frozen / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding='utf-8')
    for name in ('backend/__init__.py', 'backend/solutions/__init__.py', 'backend/utils/__init__.py',
                 'backend/app/__init__.py', 'backend/storage/__init__.py'):
        write('runtime/' + name, '')
    for name in ('backend/solutions/package.py', 'backend/solutions/environment.py',
                 'backend/solutions/store.py', 'backend/utils/subprocess_worker.py', 'aspose_init.py'):
        write('runtime/' + name, (repo / name).read_text(encoding='utf-8'))
    write('runtime/backend/solutions/worker.py', 'raise AssertionError("old bootstrap must not run")')
    write('runtime/backend/utils/template_resolver.py', 'def extract_template_ref(code): return None,None,None\n')
    write('runtime/backend/storage/storage_manager.py', 'class StorageManager:\n def __init__(self,*args): pass\n')
    write('runtime/backend/app/main.py', '')
    write('runtime/probe.py', '''def run():
    from Aspose.Cells import Workbook
    from System import Environment
    from pathlib import Path
    import tempfile
    import openpyxl
    with tempfile.TemporaryDirectory() as folder:
        file = str(Path(folder) / 'result.xlsx')
        book = Workbook()
        try:
            book.Worksheets[0].Cells['A1'].PutValue(21)
            book.Worksheets[0].Cells['B1'].Formula = '=A1*2'
            book.CalculateFormula()
            book.Save(file)
        finally:
            book.Dispose()
        with open(file, 'rb') as stream:
            values = openpyxl.load_workbook(stream, data_only=True)
            try:
                return {'value': values.active['B1'].value, 'dotnet': str(Environment.Version)}
            finally:
                values.close()
''')
    write('business/script.py', '# frozen original business')
    write('archive/records.json', '{}')
    (runtime_dir / 'libs').mkdir()
    for name in ('Aspose.Cells.dll', 'SkiaSharp.dll', 'libSkiaSharp.so'):
        shutil.copy2(repo / 'libs' / name, runtime_dir / 'libs' / name)
    shutil.copy2(repo / 'tools/iis/runtimeconfig.json', runtime_dir / 'libs/runtimeconfig.json')
    expected = environment.fingerprint(runtime_dir)
    expected.update(platform='linux', machine='x86_64', python_binary='linux-executable')
    manifest = package.seal(frozen, dict(environment=expected, settings={}, templates={}, metadata={}))
    monkeypatch.setattr(execution, 'artifact', lambda digest: frozen)
    params = tmp_path / 'params.pkl'
    params.write_bytes(pickle.dumps({'entry': 'probe:run', 'args': (), 'kwargs': {}}))
    command = execution.launch(params, {'digest': manifest['digest']}, 'rpc')
    assert Path(command[5]) == repo / 'backend/solutions/worker.py'
    result = subprocess.run(command, capture_output=True, encoding='utf-8', timeout=60)
    assert result.returncode == 0, result.stdout + result.stderr
    assert '@@RESULT@@' in result.stdout, result.stdout + result.stderr
    output = next(line[10:] for line in result.stdout.splitlines() if line.startswith('@@RESULT@@'))
    try:
        data = pickle.loads(Path(output).read_bytes())
        assert data == {'value': 42, 'dotnet': '9.0.21'}
    finally:
        Path(output).unlink()
    assert package.verify(frozen) == manifest
    assert not (runtime_dir / 'libs/libSkiaSharp.dll').exists()
