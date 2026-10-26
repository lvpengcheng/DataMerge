"""Explicit per-script input files; omission preserves the existing input set."""
import ast
import inspect
import json
import re
import shutil
import tempfile
from contextlib import contextmanager
from functools import wraps
from pathlib import Path


def _file_stem(name):
    path = Path(name)
    return path.stem if path.suffix.lower() in ('.xlsx', '.xlsm', '.xls') else path.name


def validate_source_files(value):
    if value is None:
        return None
    if not isinstance(value, list) or any(not isinstance(n, str) or not n.strip() for n in value):
        raise ValueError('SOURCE_FILES 必须是文件名列表；不限制来源时请省略或设为 None')
    result = []
    for name in value:
        name = name.strip()
        if '/' in name or '\\' in name or Path(name).suffix.lower() not in ('.xlsx', '.xlsm', '.xls'):
            raise ValueError(f'源文件须使用不含路径的完整文件名: {name}')
        if name.casefold() in {n.casefold() for n in result}:
            raise ValueError(f'重复指定源文件: {name}')
        result.append(name)
    return result


def source_files_from_rules(content):
    """An explicit list, never infer a whitelist from incidental file mentions."""
    text = str(content or '')
    matches = re.findall(r'^\s*(?:SOURCE_FILES|使用源文件|使用源)\s*[:：=]\s*(\[[^\n]*\])\s*$', text, re.M)
    blocks = re.findall(r'^\s*(?:使用源文件|使用源)\s*[:：]\s*\n((?:[ \t]*[-*][ \t]+[^\n]+\n?)+)', text, re.M)
    lists = []
    for raw in matches:
        try:
            lists.append(validate_source_files(json.loads(raw)))
        except json.JSONDecodeError as exc:
            raise ValueError('使用源文件列表须使用 JSON 双引号') from exc
    for block in blocks:
        lists.append(validate_source_files([re.sub(r'^\s*[-*]\s+', '', line).strip().strip('`')
                                           for line in block.splitlines() if line.strip()]))
    if not lists:
        inline = re.findall(r'^[ \t]*(?:SOURCE_FILES|使用源文件|使用源)[ \t]*[:：=][ \t]*([^\n]*)$', text, re.M)
        for raw in inline:
            if raw.strip() in ('None', 'null', '全部'):
                lists.append(None)
            elif raw.strip():
                lists.append(validate_source_files([n.strip().strip('`') for n in re.split('[,，]', raw)]))
            else:
                raise ValueError('使用源文件配置为空，请列出完整文件名；取消限制请填写 null')
    if len(lists) > 1 and any(value != lists[0] for value in lists[1:]):
        raise ValueError('当前目标存在多份不同的使用源文件配置，请保留一份明确的列表')
    return lists[0] if lists else None


def source_files_from_code(code):
    value = None
    for node in ast.parse(code or '').body:
        if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == 'SOURCE_FILES' for t in node.targets):
            value = validate_source_files(ast.literal_eval(node.value))
    return value


@contextmanager
def selected_input_folder(input_folder, selection, source_data=None):
    selection = validate_source_files(selection)
    if selection is None:
        yield input_folder
        return
    root = Path(input_folder)
    available = {}
    for path in root.iterdir():
        if path.is_file():
            available.setdefault(path.name.casefold(), []).append(path)
    chosen = []
    for name in selection:
        hits = available.get(name.casefold(), [])
        if not hits and source_data:
            origins = {origin['file_path'] for entry in source_data.values()
                       if _file_stem(entry.get('source_file') or '').casefold() == Path(name).stem.casefold()
                       for origin in entry.get('source_origins') or []}
            hits = [Path(p) for p in origins if Path(p).is_file()]
        if len(hits) != 1:
            raise ValueError(f'指定源文件缺失或不唯一: {name}')
        chosen.append((hits[0], name))
    with tempfile.TemporaryDirectory(prefix='selected_sources_') as temp:
        for path, name in chosen:
            shutil.copy2(path, Path(temp) / name)
        yield temp


def filter_source_data(data, selection, reserved_names=(), input_folder=None):
    if selection is None:
        return data
    from backend.utils.data_helpers import assign_sheet_keys
    stems = {Path(n).stem.casefold() for n in selection}
    if len(stems) != len(selection):
        raise ValueError('指定源文件存在同名不同扩展名，无法唯一匹配预加载数据')
    entries, seen = [], set()
    legacy = None
    for key, entry in data.items():
        if id(entry) in seen:
            continue
        seen.add(id(entry))
        source = entry.get('source_file')
        sheet = entry.get('source_sheet')
        if not source or not sheet:
            if legacy is None:
                # Old saved scripts predate provenance fields. Resolve their keys
                # from workbook names, never guess from substrings.
                import pandas as pd
                pairs = []
                for path in sorted(Path(input_folder).iterdir()):
                    if path.suffix.lower() in ('.xlsx', '.xlsm', '.xls') and not path.name.startswith('~'):
                        with pd.ExcelFile(path) as book:
                            pairs.extend((path.stem, sn) for sn in book.sheet_names)
                mapping = assign_sheet_keys(pairs, reserved_names=set(reserved_names))
                legacy = {}
                for pair, canonical in mapping.items():
                    for alias in {canonical, pair[1], pair[0] + '_' + pair[1]}:
                        legacy.setdefault(alias, set()).add(pair)
            hits = legacy.get(key, set())
            if len(hits) != 1:
                raise ValueError(f'旧脚本源表 {key} 无法唯一确定文件来源，请重新智训或补充 source_file/source_sheet')
            source, sheet = next(iter(hits))
            entry = {**entry, 'source_file': source, 'source_sheet': sheet}
        if _file_stem(source).casefold() in stems:
            entries.append((_file_stem(source), sheet, entry))
    found = {fb.casefold() for fb, _, _ in entries}
    missing = stems - found
    if missing:
        raise ValueError('指定源文件没有可读取的工作表: ' + '、'.join(sorted(missing)))
    keys = assign_sheet_keys([(fb, sn) for fb, sn, _ in entries], reserved_names=set(reserved_names))
    return {keys[fb, sn]: entry for fb, sn, entry in entries}


def source_scoped_generation(function):
    """Use the same input subset when generating code and when running it."""
    @wraps(function)
    def wrapped(*args, **kwargs):
        bound = inspect.signature(function).bind(*args, **kwargs)
        bound.apply_defaults()
        options = bound.arguments
        structure = options.get('expected_structure') or {}
        original = options.get('original_code', options.get('full_code', ''))
        if ((isinstance(structure, dict) and len(structure.get('workbooks') or {}) > 1)
                or 'MULTI_WORKBOOK_SCRIPTS =' in original):
            return function(*args, **kwargs)
        from backend.utils.multi_target_rules import rules_for_structure
        rules = rules_for_structure(options.get('rules_content'), structure,
                                    Path(options.get('template_path') or '').name)
        selection = source_files_from_rules(rules)
        configured = bool(re.search(r'^[ \t]*(?:SOURCE_FILES|使用源文件|使用源)[ \t]*[:：=]', str(rules or ''), re.M))
        if selection is None and original and not configured:
            selection = source_files_from_code(original)
        with selected_input_folder(options.get('input_folder'), selection if options.get('input_folder') else None) as folder:
            if 'input_folder' in options:
                options['input_folder'] = folder
            result = function(*bound.args, **bound.kwargs)
        if (selection is not None or configured) and result:
            code = result[0] if isinstance(result, tuple) else result
            if code:
                code += '\n\nSOURCE_FILES = ' + repr(selection) + '\n'
                result = (code, *result[1:]) if isinstance(result, tuple) else code
        return result
    return wrapped
