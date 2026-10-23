"""Keep source, target and design attachments separate while retaining names."""
from pathlib import Path
from fastapi import HTTPException
from .upload_stream import save_upload_file, safe_upload_name


async def save_rule_attachments(directory, source_files, target_files, design_docs):
    result = {}
    for category, files in [('source', source_files), ('target', target_files), ('design', design_docs)]:
        folder = Path(directory) / category
        folder.mkdir(parents=True, exist_ok=True)
        paths, names, seen = [], [], set()
        for file in files:
            if not file or not file.filename:
                continue
            name = safe_upload_name(file.filename, category + '.xlsx')
            key = (Path(name).with_suffix('.xlsx').name if category == 'target' else name).casefold()
            if key in seen:
                raise HTTPException(400, detail=f'{category}附件重名：{name}，请先重命名')
            seen.add(key)
            path = folder / name
            await save_upload_file(file, path)
            paths.append(str(path))
            names.append(name)
        result[category + '_paths'] = paths
        result[category + '_names'] = names
    return result
