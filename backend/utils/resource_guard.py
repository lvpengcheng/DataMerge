"""Memory admission and bounded process termination, without reading task disks."""
import ctypes
import os
import signal
import subprocess
import sys
import threading
import time
from pathlib import Path

_fault = None
_fault_lock = threading.Lock()


def memory_snapshot_mb():
    """Return (effective total, available); honor Linux container memory limits."""
    try:
        if sys.platform == "win32":
            class Status(ctypes.Structure):
                _fields_ = [("length", ctypes.c_ulong), ("load", ctypes.c_ulong)] + [
                    (name, ctypes.c_ulonglong) for name in
                    ("total", "available", "page_total", "page_available", "virtual", "virtual_available", "extended")]
            status = Status()
            status.length = ctypes.sizeof(status)
            if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
                return None
            return status.total / 1048576, status.available / 1048576
        info = {}
        for line in Path("/proc/meminfo").read_text().splitlines():
            key, value = line.split(":", 1)
            info[key] = int(value.strip().split()[0]) / 1024
        total, available = info["MemTotal"], info.get("MemAvailable", info["MemFree"])
        for limit_path, used_path in (
            ("/sys/fs/cgroup/memory.max", "/sys/fs/cgroup/memory.current"),
            ("/sys/fs/cgroup/memory/memory.limit_in_bytes", "/sys/fs/cgroup/memory/memory.usage_in_bytes"),
        ):
            try:
                limit = int(Path(limit_path).read_text()) / 1048576
                used = int(Path(used_path).read_text()) / 1048576
                total, available = min(total, limit), min(available, max(0, limit - used))
            except (OSError, ValueError):
                continue
        return total, available
    except (OSError, ValueError, KeyError):
        return None


def reserve_mb(snapshot):
    return max(256, snapshot[0] * 0.15, float(os.getenv("EXCEL_MEMORY_RESERVE_MB", "256")))


def memory_limit_mb(configured=0, existing_rss=0):
    """0 means available-memory budget; explicit caps may only lower that budget."""
    snapshot = memory_snapshot_mb()
    budget = max(128, int(snapshot[1] - reserve_mb(snapshot) + existing_rss)) if snapshot else 1024
    budget = min(budget, max(128, int(os.getenv("EXCEL_TASK_MEMORY_CEILING_MB", "8192"))))
    return min(int(configured), budget) if configured and configured > 0 else budget


def admission_ready():
    ensure_healthy()
    snapshot = memory_snapshot_mb()
    return snapshot is None or snapshot[1] >= reserve_mb(snapshot) + 256


def wait_for_memory(timeout, cancel_event=None):
    deadline = time.monotonic() + max(0, timeout)
    while not admission_ready():
        if cancel_event and cancel_event.is_set():
            raise RuntimeError("任务已取消")
        if time.monotonic() >= deadline:
            raise RuntimeError("可用内存不足，资源等待超时；请释放内存或增加机器内存后重试")
        time.sleep(min(0.5, max(0, deadline - time.monotonic())))


def ensure_healthy():
    if _fault:
        raise RuntimeError(_fault)


def trip_circuit(pid):
    global _fault
    with _fault_lock:
        _fault = (f"进程 {pid} 在终止期限内仍未退出，可能存在磁盘/内核 I/O 阻塞；"
                  "本服务进程已暂停派发新的 Excel 重任务，请先检查存储和残留进程，再重启服务")


def process_group_options():
    return {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP} if sys.platform == "win32" else {"start_new_session": True}


def kill_tree(pid):
    if sys.platform == "win32":
        try:
            result = subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)],
                                    capture_output=True, timeout=5,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                return
        except (OSError, subprocess.TimeoutExpired):
            pass
        from ctypes import wintypes
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = kernel.OpenProcess(1, False, pid)
        if handle:
            try:
                kernel.TerminateProcess(handle, 1)
            finally:
                kernel.CloseHandle(handle)
    else:
        try:
            # Only kill our own isolated group, never the API server's group.
            if os.getpgid(pid) == pid:
                os.killpg(pid, signal.SIGKILL)
            else:
                os.kill(pid, signal.SIGKILL)
        except ProcessLookupError:
            pass


def terminate_process(proc, grace_seconds=None):
    """Never wait indefinitely after kill. A surviving process stops new dispatch."""
    if proc is None or proc.poll() is not None:
        return True
    try:
        kill_tree(proc.pid)
    except OSError:
        pass
    grace = max(0.1, min(10, float(os.getenv("SUBPROCESS_KILL_GRACE_SECONDS", "5")))) if grace_seconds is None else grace_seconds
    try:
        proc.wait(timeout=grace)
        return True
    except subprocess.TimeoutExpired:
        trip_circuit(proc.pid)
        return False
