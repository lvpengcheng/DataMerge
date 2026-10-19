"""Pin reference assets and historical results for one task or package validation."""
import json
import shutil
from pathlib import Path
from types import SimpleNamespace
from .package import canonical, digest_file, safe_name


def snapshot(root, tenant, destination, db):
    from ..database.models import DataAsset
    from .store import row_data
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=False)
    files, originals = {}, {}
    def copy(source, relative):
        source = Path(source).resolve()
        if not source.is_file():
            return None
        allowed = ((Path(root) / 'tenants' / tenant).resolve(), (Path(root) / 'global_assets').resolve())
        if not any(source.is_relative_to(folder) for folder in allowed):
            raise ValueError('历史或基础资料不在当前租户/公共资料目录内')
        target = destination / safe_name(relative)
        target.parent.mkdir(parents=True, exist_ok=True)
        before = digest_file(source)
        shutil.copyfile(source, target)
        if digest_file(source) != before or digest_file(target) != before:
            raise ValueError('基础资料在快照期间变化，请重试')
        files[relative] = before
        originals[source] = before
        return relative
    assets = db.query(DataAsset).filter(DataAsset.asset_type == 'reference', DataAsset.is_active == True,
        (DataAsset.tenant_id == tenant) | DataAsset.tenant_id.is_(None)).order_by(DataAsset.created_at.desc()).all()
    records = []
    asset_records = [row_data(asset) for asset in assets]
    for asset in assets:
        record = row_data(asset)
        record['file_path'] = copy(asset.file_path, f'references/{asset.id}/{Path(asset.file_path).name}')
        records.append(record)
    folder = Path(root) / 'tenants' / tenant
    history = folder / 'history'
    history_files = sorted(str(p.resolve()) for p in history.rglob('*') if p.is_file()) if history.is_dir() else []
    if history.is_dir():
        for file in history.rglob('*'):
            if file.is_symlink():
                raise ValueError('历史资料包含链接')
            if file.is_file():
                copy(file, 'tenants/' + tenant + '/history/' + file.relative_to(history).as_posix())
    copy(folder / 'calculation_history.json', 'tenants/' + tenant + '/calculation_history.json')
    # A collection of individually valid copies must also describe one stable
    # set of rows/files, rather than mix revisions changed during the copy.
    if any(not path.is_file() or digest_file(path) != digest for path, digest in originals.items()):
        raise ValueError('历史或基础资料在快照期间变化，请重试')
    current_history = sorted(str(p.resolve()) for p in history.rglob('*') if p.is_file()) if history.is_dir() else []
    if current_history != history_files:
        raise ValueError('历史资料目录在快照期间变化，请重试')
    db.expire_all()
    current_assets = db.query(DataAsset).filter(DataAsset.asset_type == 'reference', DataAsset.is_active == True,
        (DataAsset.tenant_id == tenant) | DataAsset.tenant_id.is_(None)).order_by(DataAsset.created_at.desc()).all()
    if [row_data(asset) for asset in current_assets] != asset_records:
        raise ValueError('基础资料记录在快照期间变化，请重试')
    data = {'tenant_id': tenant, 'assets': records, 'files': files}
    (destination / 'inputs.json').write_bytes(canonical(data))
    return digest_file(destination / 'inputs.json')


def install(destination, tenant, expected_digest=None):
    """Replace only the frozen runtime's data readers; result persistence stays live."""
    destination = Path(destination)
    if expected_digest and digest_file(destination / 'inputs.json') != expected_digest:
        raise ValueError('任务基础资料快照清单已变化')
    data = json.loads((destination / 'inputs.json').read_text(encoding='utf-8'))
    for relative, digest in data['files'].items():
        if digest_file(destination / safe_name(relative)) != digest:
            raise ValueError('任务基础资料快照已变化')
    from backend.utils import source_auto_filler, historical_data
    records = []
    for item in data['assets']:
        record = dict(item)
        record['tenant_id'] = tenant if record.get('tenant_id') is not None else None
        record['file_path'] = str(destination / safe_name(record['file_path'])) if record.get('file_path') else ''
        records.append(SimpleNamespace(**record))
    def reference_assets(requested_tenant, db):
        if requested_tenant != tenant:
            raise ValueError('方案不能读取其他租户的基础资料')
        return [a for a in records if a.tenant_id == tenant], [a for a in records if a.tenant_id is None]
    source_auto_filler._load_reference_assets = reference_assets
    original = historical_data.HistoricalDataProvider
    class PinnedHistory(original):
        def __init__(self, requested_tenant, base_dir='tenants'):
            if requested_tenant != tenant:
                raise ValueError('方案不能读取其他租户的历史数据')
            super().__init__(data['tenant_id'], str(destination / 'tenants'))
            self.tenant_id = tenant
    historical_data.HistoricalDataProvider = PinnedHistory
