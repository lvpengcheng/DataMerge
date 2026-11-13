"""Environment compatibility is fail-closed; never pip-install an imported package."""
import importlib.metadata
import os
import platform
import sys
import subprocess
import json
import hashlib
import ast
import re
from pathlib import Path
from .package import digest_file

# This is the spreadsheet execution contract, not a hash of the host's tools,
# web server, GPU drivers or Python executable. Platform wheels may differ.
CALCULATION_PACKAGES = frozenset({
    'numpy', 'pandas', 'openpyxl', 'xlrd', 'pythonnet', 'clr-loader',
    'et-xmlfile', 'python-dateutil', 'pytz', 'tzdata', 'six', 'cffi',
})


def compatibility_errors(expected, actual):
    errors = []
    for key in ('python', 'implementation'):
        if not expected.get(key) or expected.get(key) != actual.get(key):
            errors.append(f'{key}: 方案={expected.get(key)}，本机={actual.get(key)}')
    aliases = {'amd64': 'x86_64', 'aarch64': 'arm64'}
    machines = [aliases.get(str(e.get('machine', '')).lower(), str(e.get('machine', '')).lower())
                for e in (expected, actual)]
    if not machines[0] or machines[0] != machines[1]:
        errors.append('CPU 架构不兼容')
    if expected.get('platform') not in ('win32', 'linux') or actual.get('platform') not in ('win32', 'linux'):
        if expected.get('platform') != actual.get('platform'):
            errors.append('不支持此操作系统组合')
    required = {v for v in expected.get('dotnet', []) if v.startswith('Microsoft.NETCore.App ')}
    if not required or not required.issubset(set(actual.get('dotnet', []))):
        errors.append('.NET Core: 需要 ' + ', '.join(sorted(required)))
    def packages(env):
        return {re.sub(r'[-_.]+', '-', k.lower()): v for k, v in env.get('packages', {}).items()}
    wanted, installed = packages(expected), packages(actual)
    for name in sorted(CALCULATION_PACKAGES & wanted.keys()):
        if installed.get(name) != wanted[name]:
            errors.append(f'{name}: 方案={wanted[name]}，本机={installed.get(name, "未安装")}')
    return errors


def prepare_platform_runtime(expected, runtime, host):
    """Add only the destination Skia native binary to the private execution copy."""
    import shutil
    runtime, host = Path(runtime), Path(host)
    filename = 'libSkiaSharp.dll' if sys.platform == 'win32' else 'libSkiaSharp.so'
    target = runtime / 'libs' / filename
    if target.is_file() or not (runtime / 'libs/SkiaSharp.dll').is_file():
        return
    native = host / 'libs' / filename
    managed = host / 'libs/SkiaSharp.dll'
    if not native.is_file() or not managed.is_file():
        raise ValueError('本机缺少平台运行库：' + filename)
    if digest_file(managed) != expected.get('native', {}).get('SkiaSharp.dll'):
        raise ValueError('本机 SkiaSharp 程序集与方案版本不一致，无法适配平台运行库')
    shutil.copy2(native, target)


def fingerprint(root):
    native = {}
    libs = Path(root) / 'libs'
    for p in sorted(libs.rglob('*')) if libs.is_dir() else []:
        if p.is_file() and p.suffix.lower() in ('.dll', '.so', '.json', '.dylib'):
            native[p.relative_to(libs).as_posix()] = digest_file(p)
    try:
        local = Path(os.environ.get('DOTNET_ROOT', '')) / ('dotnet.exe' if os.name == 'nt' else 'dotnet')
        dotnet = subprocess.run([str(local.resolve()) if local.is_file() else 'dotnet', '--list-runtimes'], capture_output=True, text=True, timeout=10)
        dotnet_versions = sorted(' '.join(line.split()[:2]) for line in dotnet.stdout.splitlines())
    except (OSError, subprocess.TimeoutExpired):
        dotnet_versions = []
    return {
        'python': platform.python_version(), 'implementation': platform.python_implementation(),
        'python_binary': digest_file(sys.executable), 'dotnet': dotnet_versions,
        'platform': sys.platform, 'machine': platform.machine(),
        'packages': dict(sorted((d.metadata['Name'].lower(), d.version)
                               for d in importlib.metadata.distributions() if d.metadata.get('Name'))),
        'native': native,
    }


def require_environment(expected, root):
    runtime_python(expected, root)


def environment_id(expected):
    from .package import canonical
    return hashlib.sha256(canonical(expected)).hexdigest()


def runtime_python(expected, root):
    """A local deployment registry retains old environments after application upgrades.

    Registry paths are never accepted from an imported ZIP. Native libraries are
    always loaded from the verified package and may differ from the current app.
    """
    current = fingerprint(root)
    current['native'] = expected['native']
    if current == expected:
        return sys.executable
    errors = compatibility_errors(expected, current)
    if not errors:
        return sys.executable
    registry = Path(root) / 'solution_store' / 'runtimes.json'
    candidates = json.loads(registry.read_text(encoding='utf-8')) if registry.exists() else {}
    executable = candidates.get(environment_id(expected))
    if executable and Path(executable).is_file():
        probe = ('import sys,json;sys.path.insert(0,sys.argv[1]);'
                 'from backend.solutions.environment import fingerprint;'
                 'print(json.dumps(fingerprint(sys.argv[1])))')
        result = subprocess.run([executable, '-I', '-c', probe, str(root)],
                                capture_output=True, text=True, timeout=60)
        if result.returncode == 0:
            actual = json.loads(result.stdout.strip().splitlines()[-1])
            actual['native'] = expected['native']
            if actual == expected or not compatibility_errors(expected, actual):
                return executable
    raise ValueError('方案运行环境不兼容，禁止自动使用不兼容依赖。请安装对应版本或登记兼容环境：'
                     + '；'.join(errors) + '。环境标识：' + environment_id(expected))


def calculation_settings():
    # A snapshot of ALL non-secret calculation settings, including absent keys, is
    # applied in the worker after disabling dotenv. Operational secrets stay local.
    prefixes = ('COMPUTE_', 'EXCEL_', 'ASPOSE_', 'SOURCE_', 'COLUMN_', 'TEMPLATE_',
                'OUTPUT_', 'FORMULA_', 'MATCH_', 'HISTORY_', 'TRAINING_', 'HEADER_', 'RESTORE_')
    from .package import SECRET
    # Also discover literal environment reads, so a calculation flag without one
    # of the historical prefixes cannot silently drift after deployment.
    root = Path(__file__).resolve().parents[2]
    keys = {'TZ', 'LANG', 'LC_ALL'}
    for file in [*root.glob('*.py'), *(root / 'backend').rglob('*.py')]:
        if any(part in file.parts for part in ('tests', 'test_data', '__pycache__')):
            continue
        try:
            for node in ast.walk(ast.parse(file.read_text(encoding='utf-8-sig'))):
                if not isinstance(node, ast.Call) or not node.args:
                    continue
                name = ast.unparse(node.func)
                if name not in ('os.getenv', 'os.environ.get', 'environ.get', 'getenv', 'env_int', 'env_float', 'env_bool'):
                    continue
                if isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str):
                    keys.add(node.args[0].value)
        except (OSError, SyntaxError):
            continue
    operational = ('DATABASE_', 'DB_', 'JWT_', 'SMTP_', 'MAIL_', 'EMAIL_', 'REDIS_',
                   'CELERY_', 'AWS_', 'PYTHON', 'DOTNET_', 'SOLUTION_', 'SOP_')
    return {k: v for k, v in os.environ.items()
            if (k.startswith(prefixes) or k in keys) and not SECRET.search(k)
            and not k.startswith(operational) and k not in ('PATH', 'HOME', 'TEMP', 'TMP', 'USERPROFILE', 'SYSTEMROOT')}
