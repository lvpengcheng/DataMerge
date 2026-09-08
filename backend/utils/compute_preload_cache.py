"""Task-local parsed data cache; avoids reopening Excel between precheck and compute."""
import hashlib
import json
import os
import pickle
from pathlib import Path


def _signature(source_dir, context):
    files = []
    for path in sorted(Path(source_dir).iterdir()):
        if path.is_file() and path.suffix.lower() in (".xls", ".xlsx", ".xlsm"):
            stat = path.stat()
            files.append((path.name, stat.st_size, stat.st_mtime_ns))
    # Normalize JSON columns consistently across SQLite and other databases.
    def normalize(value):
        for _ in range(2):
            if isinstance(value, str):
                try:
                    value = json.loads(value)
                except ValueError:
                    break
        return value
    canonical = json.dumps([files, [normalize(x) for x in context]], sort_keys=True,
                           ensure_ascii=False, default=str).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest()


def save_preload(source_dir, data, mapping, context):
    """Only called on server-created task directories; never accepts uploaded pickle."""
    path = Path(source_dir).parent / "_validated_source.pkl"
    staged = path.with_suffix(".tmp")
    try:
        with staged.open("wb") as stream:
            pickle.dump({"signature": _signature(source_dir, context),
                         "data": data, "mapping": mapping}, stream, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(staged, path)
    finally:
        if staged.exists():
            staged.unlink()


def load_preload(source_dir, context):
    path = Path(source_dir).parent / "_validated_source.pkl"
    if not path.exists():
        return None
    with path.open("rb") as stream:
        cached = pickle.load(stream)
    if cached.get("signature") != _signature(source_dir, context):
        return None
    return cached["mapping"], cached["data"]
