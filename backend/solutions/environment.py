"""Environment compatibility is fail-closed; never pip-install an imported package."""
import importlib.metadata
import os
import platform
import sys
import subprocess
import json
import hashlib
import ast
from pathlib import Path
from .package import digest_file


def fingerprint(root):
    native = {}
    libs = Path(root) / 'libs'
    for p in sorted(libs.rglob('*')) if libs.is_dir() else []:
        if p.is_file() and p.suffix.lower() in ('.dll', '.so', '.json', '.dylib'):
            native[p.relative_to(libs).as_posix()] = digest_file(p)
    try:
        dotnet = subprocess.run(['dotnet', '--list-runtimes'], capture_output=True, text=True, timeout=10)
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
            if actual == expected:
                return executable
    raise ValueError('方案运行环境不匹配，请保留并登记原 Python/依赖/.NET 环境；禁止自动使用当前依赖。环境标识：' + environment_id(expected))


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
