"""Portable, non-executable package validation; no application/ORM imports."""
import hashlib
import json
import re
import stat
import zipfile
from pathlib import Path, PurePosixPath

FORMAT = 1
MAX_FILES = 30000
MAX_BYTES = 4 * 1024 ** 3


def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'),
                      allow_nan=False).encode('utf-8')


def digest_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def safe_name(name):
    parts = PurePosixPath(name).parts
    if (not name or '\\' in name or ':' in name or name.startswith('/')
            or not parts or any(p in ('', '.', '..') or p.endswith((' ', '.')) for p in name.split('/'))):
        raise ValueError('方案包含非法路径')
    for part in parts:
        if part.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *(f'COM{i}' for i in range(10)), *(f'LPT{i}' for i in range(10))}:
            raise ValueError('方案包含保留文件名')
    return name


def inventory(root):
    result = {}
    for path in sorted(Path(root).rglob('*')):
        if path.is_symlink():
            raise ValueError('方案不允许符号链接')
        if path.is_file() and path != root / 'manifest.json':
            name = safe_name(path.relative_to(root).as_posix())
            result[name] = {'sha256': digest_file(path), 'size': path.stat().st_size}
    return result


def seal(root, metadata):
    manifest = dict(metadata, format_version=FORMAT, files=inventory(root))
    if (len(manifest['files']) + 1 > MAX_FILES or sum(item['size'] for item in manifest['files'].values()) > MAX_BYTES
            or any(item['size'] > 512 * 1024 ** 2 for item in manifest['files'].values())):
        raise ValueError('方案超出导入支持的文件数量或大小限制')
    manifest['digest'] = hashlib.sha256(canonical(manifest)).hexdigest()
    (Path(root) / 'manifest.json').write_bytes(canonical(manifest))
    return manifest


def verify(root):
    root = Path(root)
    try:
        manifest = json.loads((root / 'manifest.json').read_text(encoding='utf-8'))
    except (OSError, ValueError) as exc:
        raise ValueError('方案清单缺失或无效') from exc
    if not isinstance(manifest, dict):
        raise ValueError('方案清单必须是对象')
    if manifest.get('format_version') != FORMAT:
        raise ValueError('不支持的方案格式版本')
    body = {k: v for k, v in manifest.items() if k != 'digest'}
    if hashlib.sha256(canonical(body)).hexdigest() != manifest.get('digest'):
        raise ValueError('方案清单校验失败')
    if re.fullmatch(r'[0-9a-f]{64}', root.name) and root.name != manifest['digest']:
        raise ValueError('方案内容与绑定版本不一致')
    if inventory(root) != manifest.get('files'):
        raise ValueError('方案文件缺失、被修改或包含未登记文件')
    for required in ('business/script.py', 'archive/records.json', 'runtime/backend/solutions/worker.py'):
        if required not in manifest['files']:
            raise ValueError(f'方案缺少必要文件: {required}')
    return manifest


def export_zip(root, destination):
    manifest = verify(root)
    with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name in ['manifest.json', *sorted(manifest['files'])]:
            archive.write(Path(root) / name, name)


def extract_zip(source, destination):
    """Extract only after validating the complete directory, never run imports."""
    root = Path(destination).resolve()
    root.mkdir(parents=True, exist_ok=False)
    with zipfile.ZipFile(source) as archive:
        entries = archive.infolist()
        if len(entries) > MAX_FILES or sum(i.file_size for i in entries) > MAX_BYTES:
            raise ValueError('方案解压大小或文件数量超过限制')
        seen = set()
        for info in entries:
            if info.is_dir():
                continue
            name = safe_name(info.filename)
            if name.casefold() in seen or stat.S_ISLNK(info.external_attr >> 16) or info.flag_bits & 1:
                raise ValueError('方案包含重名、链接或加密条目')
            seen.add(name.casefold())
            if info.file_size > 512 * 1024 ** 2:
                raise ValueError('方案单文件过大')
            if root not in (root / name).resolve().parents:
                raise ValueError('方案路径越界')
        used = 0
        for info in entries:
            if info.is_dir():
                continue
            path = root / info.filename
            path.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as src, path.open('xb') as dst:
                for block in iter(lambda: src.read(1024 * 1024), b''):
                    used += len(block)
                    if used > MAX_BYTES:
                        raise ValueError('方案解压大小超过限制')
                    dst.write(block)
    return verify(root)


SECRET = re.compile(r'(password|passwd|secret|token|api[_-]?key|authorization|database_url|credential)', re.I)


def scrub(value):
    if isinstance(value, dict):
        return {k: scrub(v) for k, v in value.items() if not SECRET.search(str(k))}
    if isinstance(value, list):
        return [scrub(v) for v in value]
    if isinstance(value, str):
        return re.sub(r'(?i)(Bearer\s+|sk-)[A-Za-z0-9_.-]{10,}', '[REDACTED]', value)
    return value
