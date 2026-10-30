"""Validation and role adaptation for independently configured output sheets."""
import copy
import re


def validate_tasks(tasks, tables):
    if not isinstance(tasks, list) or not 1 <= len(tasks) <= 50:
        raise ValueError('请配置1至50个目标Sheet')
    destinations = set()
    template_groups = {}
    targets = set()
    for task in tasks:
        if not isinstance(task, dict):
            raise ValueError('目标Sheet配置无效')
        main = task.get('main_file')
        active = task.get('active_files')
        if not isinstance(main, str) or main not in tables or not isinstance(active, list) or main not in active or any(not isinstance(f, str) for f in active):
            raise ValueError('目标Sheet或参与来源选择无效')
        if len(set(active)) != len(active) or not set(active) <= set(tables) or len(active) < 2:
            raise ValueError('每个目标Sheet至少需要一张来源表，且不能重复选择')
        if not isinstance(task.get('key_map'), dict) or set(task['key_map']) != set(active) or not all(isinstance(c, str) and c for c in task['key_map'].values()):
            raise ValueError('请为目标及每张来源表设置关联键')
        filename, sheet = task.get('output_file', ''), task.get('output_sheet', '')
        if (not isinstance(filename, str) or len(filename) > 150 or
                re.search(r'[\\/:*?"<>|\x00-\x1f]', filename) or
                not filename.lower().endswith('.xlsx') or filename.startswith('.') or
                filename.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *('COM'+str(i) for i in range(1,10)), *('LPT'+str(i) for i in range(1,10))}):
            raise ValueError('输出文件名须为有效的.xlsx文件名，不能包含路径')
        if not isinstance(sheet, str) or not sheet.strip() or len(sheet) > 31 or re.search(r"[\\/*?:\[\]\x00-\x1f]", sheet) or sheet.startswith("'") or sheet.endswith("'"):
            raise ValueError('输出Sheet名称无效（最多31个字符）')
        dest = (filename.casefold(), sheet.casefold())
        if dest in destinations:
            raise ValueError('同一输出文件中的目标Sheet不能重名')
        destinations.add(dest)
        target = (filename.casefold(), main)
        if target in targets:
            raise ValueError('同一输出文件中不能重复配置同一个模板Sheet，请使用不同输出文件名')
        targets.add(target)
        # A workbook preserves one template and its inter-sheet formulas.
        physical = tables[main].get('physical_file') or main
        previous = template_groups.setdefault(filename.casefold(), physical)
        if previous != physical:
            raise ValueError('同一输出文件请选择同一个模板工作簿中的Sheet；不同模板请设置不同输出文件名')
        if task.get('split_config') or task.get('output_tasks'):
            raise ValueError('目标Sheet不支持嵌套任务或拆表规则')
        if not task.get('overwrite_pairs'):
            raise ValueError('每个目标Sheet至少需要一项覆盖字段配置')
        if task.get('output_mode', 1) not in (1, 2) or (task.get('output_mode') == 2 and not task.get('compare_pairs')):
            raise ValueError('输出差异Sheet时请配置对比字段')
        for kind in ('overwrite_pairs', 'compare_pairs'):
            pairs = task.get(kind, [])
            if not isinstance(pairs, list) or len(pairs) > 2000:
                raise ValueError('字段配置无效')
            for pair in pairs:
                if not isinstance(pair, dict) or pair.get('source_file') not in active:
                    raise ValueError('公式默认来源必须在本目标选中的表中')
                if len(str(pair.get('source_expr') or pair.get('source_col') or '')) > 10000:
                    raise ValueError('公式长度超过限制')


def task_scheme_config(config, task):
    """Reuse the single-target matcher with a task-specific main role."""
    local = copy.deepcopy(config)
    local.pop('output_tasks', None)
    by_file = {r['file']: r for r in config['roles']}
    files = [task['main_file']] + [f for f in task['active_files'] if f != task['main_file']]
    local['roles'] = [by_file[f] for f in files]
    local['main_fp'] = by_file[files[0]]['fp']
    local['source_fps'] = [by_file[f]['fp'] for f in files[1:]]
    local.pop('required_cols_by_role', None)
    local['key_map_by_role'] = {str(i): task['key_map'][f] for i, f in enumerate(files)}
    for field in ('data_scope', 'name_col', 'id_col', 'output_mode', 'diff_order', 'normalize_keys', 'date_key_mode', 'union_source_keys'):
        local[field] = task.get(field)
    for kind in ('overwrite_pairs', 'compare_pairs'):
        local[kind] = [{**p, 'source_role': files.index(p['source_file']),
                        'source_fp': by_file[p['source_file']]['fp']} for p in task.get(kind, [])]
    return local


def resolve_task(task, config, match, resolve_config):
    local = task_scheme_config(config, task)
    indices = [next(i for i, r in enumerate(match['roles']) if r['file'] == lr['file']) for lr in local['roles']]
    local_match = {**match, 'roles': local['roles'],
                   'role_files': [match['role_files'][i] for i in indices],
                   'column_maps_by_role': {j: match['column_maps_by_role'].get(i, {}) for j, i in enumerate(indices)}}
    resolved = resolve_config(local, local_match)
    files = local_match['role_files']
    result = {**copy.deepcopy(task), 'main_file': files[0], 'active_files': files,
              'key_map': {files[int(i)]: col for i, col in resolved['key_map_by_role'].items()}}
    for field in ('data_scope', 'name_col', 'id_col'):
        result[field] = resolved.get(field)
    for kind in ('overwrite_pairs', 'compare_pairs'):
        result[kind] = [{**p, 'source_file': files[p['source_role']]} for p in resolved[kind]]
    return result


def execute_tasks(session_dir, request, execute_one):
    """Update sheets in one workbook per output group, preserving cross-sheet references."""
    import json
    import uuid
    import zipfile
    from pathlib import Path
    import aspose_init  # noqa: F401
    from Aspose.Cells import Workbook, SaveFormat
    aspose_init.ensure_license()
    root = Path(session_dir)
    meta = json.loads((root / '_meta.json').read_text(encoding='utf-8'))
    tables = {f['name']: f for f in meta.get('files', [])}
    tasks = request['output_tasks']
    validate_tasks(tasks, tables)
    outdir = root / ('batch_' + uuid.uuid4().hex)
    outdir.mkdir()
    totals = {k: 0 for k in ('matched_rows', 'overwritten_cells', 'added_rows', 'diff_rows')}
    groups = {}
    for task in tasks:
        groups.setdefault(task['output_file'].casefold(), []).append(task)
    results = []
    try:
        for group in groups.values():
            template = root / (tables[group[0]['main_file']].get('physical_file') or group[0]['main_file'])
            path = outdir / group[0]['output_file']
            book = Workbook(str(template))
            try:
                original_names = {book.Worksheets[i].Name.casefold() for i in range(book.Worksheets.Count)}
                target_names = {tables[t['main_file']]['sheet'].casefold() for t in group}
                for task in group:
                    if task['output_sheet'].casefold() in original_names - target_names:
                        raise ValueError('输出Sheet名称与模板中保留的其他Sheet重名：' + task['output_sheet'])
                book.Save(str(path), SaveFormat.Xlsx)
            finally:
                book.Dispose()
            for task in group:
                data = {**request, **task, 'output_tasks': None, '_template_path': str(path), '_output_path': str(path)}
                stat = execute_one(session_dir, data)
                for key in totals:
                    totals[key] += stat.get(key, 0)
            book = Workbook(str(path))
            try:
                # Temporary names allow a pair of target sheets to exchange names.
                sheets = [book.Worksheets[tables[t['main_file']]['sheet']] for t in group]
                for sheet in sheets:
                    sheet.Name = '_target_' + uuid.uuid4().hex[:20]
                for sheet, task in zip(sheets, group):
                    sheet.Name = task['output_sheet']
                book.CalculateFormula()
                book.Save(str(path), SaveFormat.Xlsx)
            finally:
                book.Dispose()
            results.append(path)
        if len(results) == 1:
            result = results[0]
        else:
            result = outdir / '整合结果.zip'
            with zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as archive:
                for path in results:
                    archive.write(path, path.name)
        return {**totals, 'batch_path': str(result), 'output_count': len(results), 'task_count': len(tasks)}
    except Exception:
        import shutil
        shutil.rmtree(outdir, ignore_errors=True)
        raise
