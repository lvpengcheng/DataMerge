import json
from pathlib import Path

from backend.utils.compute_run_archive import (
    create_tenant_compute_run_dir,
    finalize_compute_run,
)


class _Storage:
    def __init__(self, root: Path):
        self.root = root

    def get_tenant_dir(self, tenant_id: str) -> Path:
        path = self.root / tenant_id
        path.mkdir(parents=True, exist_ok=True)
        return path


def test_compute_workspace_is_final_tenant_run_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPUTE_DELETE_RUN_FILES_AFTER_FINISH", "false")
    storage = _Storage(tmp_path / "tenants")
    work = create_tenant_compute_run_dir(storage, "tenant-a", 42)
    assert work == (tmp_path / "tenants" / "tenant-a" / "compute_runs" / "task_42").resolve()
    initial_manifest = json.loads((work / "run_manifest.json").read_text(encoding="utf-8"))
    assert initial_manifest["status"] == "created"

    (work / "source").mkdir()
    (work / "source" / "input.xlsx").write_bytes(b"excel")
    (work / "_confirmations.json").write_text('{"ok": true}', encoding="utf-8")

    archived = finalize_compute_run(
        work, storage, "tenant-a", 42, status="completed", script_id="salary",
    )
    assert archived == (tmp_path / "tenants" / "tenant-a" / "compute_runs" / "task_42").resolve()
    assert (archived / "source" / "input.xlsx").read_bytes() == b"excel"
    assert (archived / "_confirmations.json").exists()
    manifest = json.loads((archived / "run_manifest.json").read_text(encoding="utf-8"))
    assert manifest["task_id"] == "42"
    assert manifest["tenant_id"] == "tenant-a"
    assert manifest["status"] == "completed"


def test_compute_run_can_be_deleted_after_finish(tmp_path, monkeypatch):
    monkeypatch.setenv("COMPUTE_DELETE_RUN_FILES_AFTER_FINISH", "true")
    storage = _Storage(tmp_path / "tenants")
    work = create_tenant_compute_run_dir(storage, "tenant-a", 7)
    assert finalize_compute_run(work, storage, "tenant-a", 7, status="failed") is None
    assert not work.exists()
