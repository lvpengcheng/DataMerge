"""Per-workbook compute template uploads, keeping the legacy single path contract."""
import json
from pathlib import Path


def target_key(workbook, sheet):
    return json.dumps([workbook, sheet], ensure_ascii=False, separators=(',', ':'))


def workbook_target_map(confirmed, workbook):
    result = {}
    for key, value in (confirmed or {}).items():
        try:
            parts = json.loads(key)
        except (ValueError, TypeError):
            continue
        if isinstance(parts, list) and len(parts) == 2 and parts[0] == workbook:
            result[parts[1]] = value
    return result


async def save_template_uploads(uploads, run_dir, script_content):
    from .multi_workbook import script_targets
    from .upload_stream import save_upload_file, safe_upload_name

    uploads = [file for file in uploads if file is not None and file.filename]
    if not uploads:
        return None
    targets = script_targets(script_content)
    multiple = len(targets) > 1
    if not multiple and len(uploads) > 1:
        raise ValueError('当前脚本仅有一个目标文件，只能上传一份模板')
    names = {Path(name).with_suffix('.xlsx').name.casefold(): name for name in targets}
    selected, seen = [], set()
    for file in uploads:
        name = safe_upload_name(file.filename, 'template.xlsx')
        if Path(name).suffix.lower() not in ('.xlsx', '.xls', '.xlsm'):
            raise ValueError(f'模板文件必须是 Excel：{name}')
        normalized = Path(name).with_suffix('.xlsx').name.casefold()
        if normalized in seen:
            raise ValueError(f'模板文件重名：{name}')
        seen.add(normalized)
        if multiple and normalized not in names:
            raise ValueError(f'模板「{name}」没有对应的目标文件，请使用目标文件名：' + '、'.join(targets))
        selected.append((file, name, names.get(normalized)))
    directory = Path(run_dir) / 'template'
    directory.mkdir(parents=True, exist_ok=True)
    result = {}
    for file, name, target in selected:
        path = directory / name
        await save_upload_file(file, path)
        result[target or name] = str(path.resolve())
    return result if multiple else next(iter(result.values()))
