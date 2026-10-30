"""Durable source asset paths and compatibility with session-only history records."""
from pathlib import Path
import re
import shutil
import uuid


def preserve_compute_source(project_root, tenant_id, task_id, source_path, file_name, *, copy_file=False):
    source = Path(source_path).resolve(strict=True)
    root = Path(project_root).resolve()
    tenant = (root / 'tenants' / tenant_id).resolve()
    if not tenant.is_relative_to(root / 'tenants'):
        raise ValueError('源文件租户目录无效')
    run_source = tenant / 'compute_runs' / f'task_{int(task_id)}' / 'source'
    # New computations already write into their permanent tenant task directory.
    if not copy_file and source.is_relative_to(run_source) and source.is_file():
        return source
    # Legacy execution directories are removed after computing. Archive those
    # inputs before execution, rather than registering a path that will vanish.
    name = Path(str(file_name).replace('\\', '/')).name
    if not name or name in ('.', '..'):
        raise ValueError('源文件名称无效')
    directory = tenant / 'assets' / 'source'
    directory.mkdir(parents=True, exist_ok=True)
    destination = directory / f'{uuid.uuid4().hex}_{name}'
    shutil.copy2(source, destination)
    return destination


def resolve_asset_download_path(asset, db, project_root):
    """Resolve only this asset's associated compute task, never another run by name."""
    value = str(asset.file_path or '')
    if not value.startswith('session://'):
        path = Path(value)
        return path if value and path.is_file() else None
    match = re.fullmatch(r'session://([1-9][0-9]*)/([^/\\]+)', value)
    if not match or asset.asset_type != 'source' or not asset.tenant_id:
        return None
    task_id, name = int(match[1]), match[2]
    if name != asset.file_name or name in ('.', '..'):
        return None
    from ..database.models import ComputeTask, ComputeTaskInput
    linked = (db.query(ComputeTaskInput).join(ComputeTask, ComputeTask.id == ComputeTaskInput.task_id)
              .filter(ComputeTaskInput.asset_id == asset.id, ComputeTaskInput.task_id == task_id,
                      ComputeTask.tenant_id == asset.tenant_id).first())
    if linked is None:
        return None
    root = Path(project_root).resolve()
    tenant = (root / 'tenants' / asset.tenant_id).resolve()
    if not tenant.is_relative_to(root / 'tenants'):
        return None
    source_dir = tenant / 'compute_runs' / f'task_{task_id}' / 'source'
    path = (source_dir / name).resolve()
    if path.is_relative_to(source_dir) and path.is_file():
        return path
    return None
