"""Narrow runtime repairs create a new verified artifact; frozen versions stay intact."""
import ast
import os
import shutil
import tempfile
import textwrap
from datetime import datetime, timezone
from pathlib import Path

from .package import seal, verify


def initialize_compute_structure(source):
    """Move the existing metadata initialization before its first template guard."""
    tree = ast.parse(source)
    task = next((n for n in tree.body if isinstance(n, ast.AsyncFunctionDef)
                 and n.name == 'run_compute_task'), None)
    if task is None:
        raise ValueError('方案缺少可修复的智算入口')
    guards = [n for n in ast.walk(task) if isinstance(n, ast.If)
              and 'workbook_definitions(expected_structure)' in ast.unparse(n.test)]
    if not guards:
        raise ValueError('方案运行代码与此修复不兼容')
    guard = min(guards, key=lambda n: n.lineno)
    assignments = {}
    for node in ast.walk(task):
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
            if name == '_script_info' or (name == 'expected_structure'
                    and isinstance(node.value, ast.Call) and isinstance(node.value.func, ast.Attribute)
                    and isinstance(node.value.func.value, ast.Name) and node.value.func.value.id == '_script_info'):
                assignments.setdefault(name, []).append(node)
    if any(len(assignments.get(name, [])) != 1 for name in ('_script_info', 'expected_structure')):
        raise ValueError('方案结构初始化与此修复不兼容')
    info, structure = (assignments[name][0] for name in ('_script_info', 'expected_structure'))
    if structure.lineno < guard.lineno:
        return source
    decoders = [n for n in ast.walk(task) if isinstance(n, ast.If)
                and ast.unparse(n.test) == 'isinstance(expected_structure, str)']
    if (len(decoders) != 1 or ast.unparse(info.value) != '_load_script_info_for_precheck(tenant_id, script_id)'
            or ast.unparse(structure.value) != "_script_info.get('expected_structure')"):
        raise ValueError('方案初始化逻辑与此修复不兼容')
    moved = [info, structure, decoders[0]]
    lines = source.splitlines(keepends=True)
    blocks = [textwrap.indent(textwrap.dedent(''.join(lines[n.lineno - 1:n.end_lineno])),
                              ' ' * guard.col_offset) for n in moved]
    removed = {i for n in moved for i in range(n.lineno - 1, n.end_lineno)}
    output = []
    for i, line in enumerate(lines):
        if i == guard.lineno - 1:
            output.extend(blocks)
            output.append('\n')
        if i not in removed:
            output.append(line)
    result = ''.join(output)
    compile(result, '<repaired-compute-runtime>', 'exec')
    return result


def repair_compute_startup(package, destination):
    """Preserve business code, assets, environment and settings in a new version."""
    package, destination = Path(package), Path(destination)
    original = verify(package)
    relative = Path('runtime/backend/app/main.py')
    source = (package / relative).read_text(encoding='utf-8')
    repaired = initialize_compute_structure(source)
    if repaired == source:
        return package, original
    destination.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.repair-', dir=destination))
    try:
        shutil.copytree(package, stage, dirs_exist_ok=True)
        (stage / relative).write_text(repaired, encoding='utf-8', newline='')
        metadata = {k: v for k, v in original.items() if k not in ('files', 'digest', 'format_version')}
        metadata['runtime_repair'] = {
            'kind': 'compute-expected-structure-initialization', 'parent_digest': original['digest'],
            'created_at': datetime.now(timezone.utc).isoformat(),
        }
        manifest = seal(stage, metadata)
        target = destination / manifest['digest']
        if target.exists():
            verify(target)
        else:
            os.replace(stage, target)
        verify(package)
        return target, verify(target)
    finally:
        if stage.exists():
            shutil.rmtree(stage)


def repair_template_names(package, destination):
    """Create a new version with named templates; never alter business/runtime code."""
    import hashlib
    import re
    from .package import safe_name
    from ..utils.multi_workbook import script_targets, template_reference_code
    from ..utils.template_resolver import extract_template_ref, portable_basename
    package, destination = Path(package), Path(destination)
    original = verify(package)
    code = (package / 'business/script.py').read_bytes().decode('utf-8')
    replacements, copies = {}, {}
    for source in [code, *(t['code'] for t in script_targets(code).values())]:
        key = hashlib.sha256(source.encode()).hexdigest()
        relative = (original.get('templates') or {}).get(key)
        if not relative or not re.fullmatch(r'[0-9a-f]{64}', Path(relative).stem):
            continue
        name, _, baked = extract_template_ref(source)
        filename = portable_basename(name or baked or '')
        if not filename or filename == Path(relative).name:
            continue
        safe_name(filename)
        target = 'assets/templates/' + key + '/' + filename
        # Include aliases used by the visible single-file multi-workbook format.
        reference_key = hashlib.sha256(template_reference_code(source).encode()).hexdigest()
        for alias in (key, reference_key):
            if alias in original['templates']:
                replacements[alias] = target
        copies[target] = relative
    if not replacements:
        return package, original
    destination.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='.repair-', dir=destination))
    try:
        shutil.copytree(package, stage, dirs_exist_ok=True)
        for target, source in copies.items():
            file = stage / target
            file.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(package / source, file)
        metadata = {k: v for k, v in original.items() if k not in ('files', 'digest', 'format_version')}
        metadata['templates'] = {**original['templates'], **replacements}
        metadata['template_name_repair'] = {
            'parent_digest': original['digest'], 'created_at': datetime.now(timezone.utc).isoformat(),
        }
        manifest = seal(stage, metadata)
        target = destination / manifest['digest']
        if target.exists():
            verify(target)
        else:
            os.replace(stage, target)
        verify(package)
        return target, verify(target)
    finally:
        if stage.exists():
            shutil.rmtree(stage)
