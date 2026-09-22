"""Tenant-scoped compute working directories and their retention policy."""

from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Optional


def delete_after_finish() -> bool:
    """Whether a completed or failed compute workspace should be deleted."""
    return os.getenv("COMPUTE_DELETE_RUN_FILES_AFTER_FINISH", "false").strip().lower() in {
        "1", "true", "yes", "on",
    }


def _safe_component(value: Any, fallback: str) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z_.-]+", "_", str(value or "")).strip("._")
    return cleaned or fallback


def create_tenant_compute_run_dir(storage_manager, tenant_id: str, task_id: Any) -> Path:
    """Create ``tenants/<tenant>/compute_runs/task_<id>`` before uploads are saved."""
    tenant_dir = Path(storage_manager.get_tenant_dir(tenant_id)).resolve()
    run_root = (tenant_dir / "compute_runs").resolve()
    run_root.mkdir(parents=True, exist_ok=True)
    safe_task_id = _safe_component(task_id, "unknown")
    run_dir = run_root / f"task_{safe_task_id}"
    # A database task id is unique; silently reusing an old directory could mix
    # two calculations and is therefore more dangerous than failing explicitly.
    run_dir.mkdir(parents=False, exist_ok=False)
    (run_dir / "run_manifest.json").write_text(
        json.dumps({
            "task_id": str(task_id),
            "tenant_id": str(tenant_id),
            "status": "created",
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "run_dir": str(run_dir),
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return run_dir


def finalize_compute_run(
    run_dir: Path | str,
    storage_manager,
    tenant_id: str,
    task_id: Any,
    *,
    status: str,
    script_id: Optional[Any] = None,
    error: Optional[str] = None,
    extra: Optional[Mapping[str, Any]] = None,
) -> Optional[Path]:
    """Write the terminal manifest in place, then apply the env retention policy."""
    path = Path(run_dir).resolve()
    if not path.is_dir():
        return None

    tenant_dir = Path(storage_manager.get_tenant_dir(tenant_id)).resolve()
    run_root = (tenant_dir / "compute_runs").resolve()
    safe_task_id = _safe_component(task_id, datetime.now().strftime("%Y%m%d_%H%M%S"))
    expected = (run_root / f"task_{safe_task_id}").resolve()
    if path != expected or run_root not in path.parents:
        raise ValueError(f"计算目录不属于当前租户任务: {path}")

    manifest = {
        "task_id": str(task_id),
        "tenant_id": str(tenant_id),
        "script_id": None if script_id is None else str(script_id),
        "status": str(status),
        "finished_at": datetime.now().isoformat(timespec="seconds"),
        "run_dir": str(path),
    }
    if error:
        manifest["error"] = str(error)
    if extra:
        manifest.update(dict(extra))
    (path / "run_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    if delete_after_finish():
        shutil.rmtree(path)
        return None
    return path
