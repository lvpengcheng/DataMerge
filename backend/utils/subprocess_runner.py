"""通用子进程执行器：把主进程内的"重活"（Aspose 解析/计算、AI 脚本执行）隔离到独立子进程。

背景: 多表整合对比/智训时，特定文件（公式密集/超大/含外部链接）会让 Aspose 在主进程内
长时间计算、内存暴涨 → VM 内存耗尽 → pagefile 疯狂读写（宿主 C 盘 IO 100%）→ 虚拟机假死。
Python 线程无法强杀，threading.join(timeout) 是假超时（超时后线程继续跑、内存照涨）。
唯一可靠的隔离是: 子进程 + 真超时强杀 + 内存护栏——子进程内怎么爆都只炸自己，主进程安全。

用法:
    res = run_in_subprocess(
        "backend.utils.excel_comparator:_aspose_calc_impl",
        (str(file_path),),
        timeout=120, max_memory_mb=4096,
        progress_cb=my_cb,          # 可选；目标函数签名须含 progress_cb 参数
    )
    if not res.success:
        # res.timed_out / res.killed_by_memory 区分原因
"""

import base64
import logging
import os
import pickle
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional
from backend.utils.resource_guard import (
    kill_tree as _kill_tree, terminate_process, process_group_options,
    memory_limit_mb, wait_for_memory, ensure_healthy,
)

logger = logging.getLogger(__name__)

_PROGRESS_MARK = "__SUBPROC_PROGRESS__"   # 与 subprocess_worker 中一致
_PROJ_ROOT = str(Path(__file__).resolve().parent.parent.parent)

# 并发上限：最多同时跑 N 个子进程，防止多任务并发时子进程内存叠加（N × 1GB 以内可控）。
# 其余调用排队等待（排队有上限 SUBPROCESS_QUEUE_TIMEOUT，满负载返回"系统繁忙"而非无限等）。
# N 用 lazy 信号量：首次使用时才从 .env SUBPROCESS_CONCURRENCY 读取（默认 1），
# 避免模块 import 时 .env 尚未加载（load_dotenv 在应用启动早期执行）导致配置读不到。
_semaphore = None
_semaphore_lock = threading.Lock()


def _get_semaphore() -> threading.Semaphore:
    """lazy 创建并发信号量：读 .env SUBPROCESS_CONCURRENCY（默认 1）。"""
    global _semaphore
    if _semaphore is None:
        with _semaphore_lock:
            if _semaphore is None:
                _semaphore = threading.Semaphore(max(1, env_int("SUBPROCESS_CONCURRENCY", 1)))
    return _semaphore


def queue_timeout() -> int:
    """排队等待上限（秒）：.env SUBPROCESS_QUEUE_TIMEOUT，默认 600。
    大文件复杂计算本身可长达数分钟，排队必须等得起——默认 600s（10 分钟）。"""
    return env_int("SUBPROCESS_QUEUE_TIMEOUT", 600)


def env_int(name: str, default: int) -> int:
    """读 .env 整数配置，无效/缺失回退默认值。"""
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


def default_max_memory_mb() -> int:
    """默认子进程内存上限（MB），由 .env SUBPROCESS_MAX_MEMORY_MB 控制，默认 0 表示自动预算。"""
    return env_int("SUBPROCESS_MAX_MEMORY_MB", 0)


def default_timeout(kind: str) -> int:
    """各类重活的默认超时（秒），由 .env SUBPROCESS_*_TIMEOUT 控制。"""
    defaults = {
        "parse": 300,      # SUBPROCESS_PARSE_TIMEOUT 解析
        "calc": 120,       # SUBPROCESS_CALC_TIMEOUT 公式计算
        "write": 600,      # SUBPROCESS_WRITE_TIMEOUT 整合回填/写操作
    }
    return env_int(f"SUBPROCESS_{kind.upper()}_TIMEOUT", defaults.get(kind, 300))


# ---------------- Windows 进程内存读取（ctypes，不依赖 psutil） ----------------

def _process_rss_mb(pid: int) -> float:
    """读取进程常驻内存（RSS）MB；失败返回 0（调用方视为不可测）。
    Windows: PSAPI WorkingSetSize；Linux: /proc/<pid>/status 的 VmRSS（Docker 容器内同样有效）。
    """
    if sys.platform != "win32":
        try:
            with open(f"/proc/{pid}/status", "r") as _f:
                for _line in _f:
                    if _line.startswith("VmRSS:"):
                        return float(_line.split()[1]) / 1024.0   # kB → MB
        except Exception:
            pass
        return 0.0
    import ctypes
    from ctypes import wintypes

    class _PROCESS_MEMORY_COUNTERS(ctypes.Structure):
        _fields_ = [
            ("cb", wintypes.DWORD),
            ("PageFaultCount", wintypes.DWORD),
            ("PeakWorkingSetSize", ctypes.c_size_t),
            ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
            ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t),
            ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    # 注意: 必须显式声明 argtypes/restype。ctypes.windll 默认 restype 是 c_int(32位)，
    # Windows x64 上 HANDLE 是 64 位指针会被截断 → OpenProcess 拿到假 handle。
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    k32.CloseHandle.restype = wintypes.BOOL
    psapi = ctypes.WinDLL("psapi", use_last_error=True)   # GetProcessMemoryInfo 属 PSAPI
    psapi.GetProcessMemoryInfo.argtypes = [
        wintypes.HANDLE, ctypes.POINTER(_PROCESS_MEMORY_COUNTERS), wintypes.DWORD,
    ]
    psapi.GetProcessMemoryInfo.restype = wintypes.BOOL

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    PROCESS_QUERY_INFORMATION = 0x0400
    handle = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_QUERY_INFORMATION,
                             False, pid)
    if not handle:
        return 0.0
    try:
        counters = _PROCESS_MEMORY_COUNTERS()
        counters.cb = ctypes.sizeof(_PROCESS_MEMORY_COUNTERS)
        if psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb):
            return counters.WorkingSetSize / (1024 * 1024)
        return 0.0
    finally:
        k32.CloseHandle(handle)


@dataclass
class SubprocessResult:
    """子进程执行结果。success=False 时 result=None，error 含明确原因。"""
    success: bool = False
    result: object = None
    error: str = ""
    timed_out: bool = False          # 超时被杀
    killed_by_memory: bool = False   # 内存超限被杀
    termination_failed: bool = False
    peak_memory_mb: float = 0.0
    duration: float = 0.0
    log_lines: list = field(default_factory=list)

    @property
    def killed(self) -> bool:
        return self.timed_out or self.killed_by_memory


def _run_single(entry, args, kwargs, timeout, max_memory_mb, progress_cb,
                cancel_event=None) -> SubprocessResult:
    """Isolated process with finite queue, execution, pipe drain and kill deadlines."""
    res = SubprocessResult()
    started = time.monotonic()
    sem = _get_semaphore()
    if not sem.acquire(timeout=queue_timeout()):
        res.error = "Excel 执行槽排队超时，请稍后重试"
        return res
    proc = None
    params_file = None
    done = threading.Event()
    state = {"result_path": None, "error_b64": None}
    try:
        wait_for_memory(queue_timeout(), cancel_event)
        if cancel_event and cancel_event.is_set():
            raise RuntimeError("任务已取消")
        cap = memory_limit_mb(default_max_memory_mb() if max_memory_mb is None else max_memory_mb)
        kwargs = dict(kwargs or {})
        if progress_cb:
            kwargs[_PROGRESS_MARK] = True
        else:
            kwargs.pop("progress_cb", None)
        fd, params_file = tempfile.mkstemp(suffix=".pkl", prefix="subproc_params_")
        with os.fdopen(fd, "wb") as f:
            pickle.dump({"entry": entry, "args": args, "kwargs": kwargs}, f, protocol=pickle.HIGHEST_PROTOCOL)
        env = os.environ.copy()
        env.update(PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        proc = subprocess.Popen(
            [sys.executable, "-u", "-m", "backend.utils.subprocess_worker", params_file],
            cwd=_PROJ_ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            env=env, **process_group_options())

        def read_output():
            try:
                for raw in proc.stdout:
                    line = raw.decode("utf-8", "replace").rstrip("\r\n")
                    if line.startswith("@@RESULT@@"):
                        state["result_path"] = line[len("@@RESULT@@"):]
                    elif line.startswith("@@ERROR@@"):
                        state["error_b64"] = line[len("@@ERROR@@"):]
                    elif line.startswith("@@PROG@@"):
                        if progress_cb:
                            try:
                                progress_cb(line[len("@@PROG@@"):])
                            except Exception:
                                pass
                    elif line.strip():
                        res.log_lines.append(line)
                        if len(res.log_lines) > 2000:
                            del res.log_lines[:1000]
            finally:
                done.set()

        threading.Thread(target=read_output, daemon=True, name="excel-output").start()
        deadline = time.monotonic() + (timeout if timeout and timeout > 0 else 3600)
        while proc.poll() is None:
            rss = _process_rss_mb(proc.pid)
            current_cap = memory_limit_mb(cap, rss)
            res.peak_memory_mb = max(res.peak_memory_mb, rss)
            if cancel_event and cancel_event.is_set():
                res.error = "任务已取消"
                break
            if rss > current_cap:
                res.killed_by_memory = True
                res.error = f"任务内存超过安全预算 {current_cap}MB（峰值 {rss:.0f}MB）"
                break
            if time.monotonic() >= deadline:
                res.timed_out = True
                res.error = f"执行超时（{timeout}s）"
                break
            time.sleep(0.2)
        if res.error:
            res.termination_failed = not terminate_process(proc)
            res.error += "；进程未退出，已暂停派发新任务，请检查存储 I/O" if res.termination_failed else "；进程已终止"
        elif not done.wait(timeout=2):
            res.error = "进程退出后输出管道未关闭，结果不完整"
            _kill_tree(proc.pid)
        elif proc.returncode != 0:
            res.error = f"子进程异常退出(code={proc.returncode})"
        elif state["error_b64"]:
            res.error = base64.b64decode(state["error_b64"]).decode("utf-8", "replace").strip()
        elif state["result_path"]:
            with open(state["result_path"], "rb") as f:
                res.result = pickle.load(f)
            res.success = True
        else:
            res.error = "子进程未返回结果"
    except Exception as exc:
        res.error = f"子进程执行失败: {exc}"
    finally:
        if proc is not None and proc.poll() is None and not res.termination_failed:
            res.termination_failed = not terminate_process(proc)
        if not res.termination_failed:
            for path in (params_file, state["result_path"]):
                if path:
                    try:
                        os.remove(path)
                    except OSError:
                        pass
        res.duration = time.monotonic() - started
        sem.release()
    return res


# ==================== 常驻 Worker 池（省掉每次启动的 Aspose 初始化开销） ====================

class _WorkerSlot:
    """一个常驻 worker：Popen + 常驻 reader 线程 + 当前任务上下文。

    任务协议（daemon 模式，stdout 行带 task_id）：
        @@RESULT@@{task_id}@@{result_path} / @@ERROR@@{task_id}@@{b64}
        @@PROG@@{task_id}@@{msg}
    """

    def __init__(self, idx: int):
        self.idx = idx
        self.proc = None
        self.lock = threading.Lock()
        self.idle = True
        self.pending: Dict[str, dict] = {}
        self.dead = True
        self.start()

    def start(self):
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        self.proc = subprocess.Popen(
            [sys.executable, "-u", "-m", "backend.utils.subprocess_worker", "--daemon"],
            cwd=_PROJ_ROOT, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, env=env, bufsize=0, **process_group_options())
        self.dead = False
        self.reader = threading.Thread(target=self._read_loop, args=(self.proc,), daemon=True,
                                       name=f"pool-reader-{self.idx}")
        self.reader.start()

    def restart(self):
        """崩溃/被杀后重启补位（pending 任务已由 kill/EOF 置完成）。"""
        ensure_healthy()
        if self.proc is not None and self.proc.poll() is None:
            if not terminate_process(self.proc):
                ensure_healthy()
        try:
            self.start()
        except Exception as e:
            logger.error(f"[subproc-pool/{self.idx}] worker 重启失败: {e}")
            self.dead = True

    def _read_loop(self, proc):
        try:
            for raw in proc.stdout:
                line = raw.decode("utf-8", "replace").rstrip("\r\n")
                if line.startswith("@@READY@@"):
                    continue
                if line.startswith(("@@RESULT@@", "@@ERROR@@", "@@PROG@@")):
                    # 行格式: @@RESULT@@{task_id}@@{payload} —— 先去掉前缀 @@ 再 partition
                    kind, _, rest = line[2:].partition("@@")
                    task_id, _, payload = rest.partition("@@")
                    with self.lock:
                        ctx = self.pending.get(task_id)
                    if ctx is None:
                        continue
                    if kind == "RESULT":
                        ctx["result_path"] = payload
                        ctx["event"].set()
                    elif kind == "ERROR":
                        ctx["error_b64"] = payload
                        ctx["event"].set()
                    elif kind == "PROG" and ctx.get("progress_cb"):
                        try:
                            ctx["progress_cb"](payload)
                        except Exception:
                            pass
                elif line.strip():
                    logger.debug(f"[subproc-pool/{self.idx}] {line}")
        except Exception:
            pass
        finally:
            # EOF = worker 退出：pending 任务全部置失败（父进程 wait 返回）
            with self.lock:
                if self.proc is not proc:
                    return  # An old reader must never invalidate its replacement.
                self.dead = True
                for tid, ctx in self.pending.items():
                    if not ctx["event"].is_set():
                        ctx["error_b64"] = base64.b64encode("worker 进程意外退出".encode("utf-8"))
                        ctx["event"].set()
                self.pending.clear()

    def submit(self, task_id: str, params_file: str, ctx: dict) -> bool:
        with self.lock:
            if self.dead or self.proc is None or self.proc.poll() is not None:
                return False
            self.pending[task_id] = ctx
            self.idle = False
        try:
            self.proc.stdin.write((params_file + "\n").encode("utf-8"))
            self.proc.stdin.flush()
            return True
        except Exception:
            return False

    def kill_and_restart(self):
        """强杀当前 worker 并重启补位（任务超时/内存超限时调用）。"""
        with self.lock:
            self.dead = True
            for tid, ctx in self.pending.items():
                if not ctx["event"].is_set():
                    ctx["error_b64"] = base64.b64encode("worker 被强杀（超时/内存超限）".encode("utf-8"))
                    ctx["event"].set()
            self.pending.clear()
        # Reap first; defer replacement until the next admitted task needs it.
        return terminate_process(self.proc)



class _WorkerPool:
    """常驻 worker 池：按 SUBPROCESS_POOL_SIZE 保留少量进程，超限时强杀重启。"""

    def __init__(self, size: int):
        self.slots = [_WorkerSlot(i) for i in range(size)]
        self._sem = threading.Semaphore(size)
        self._acquire_lock = threading.Lock()

    def _acquire_slot(self) -> "_WorkerSlot":
        """等一个空闲 slot（死亡的自动重启）。"""
        deadline = time.monotonic() + queue_timeout()
        while time.monotonic() < deadline:
            ensure_healthy()
            for slot in self.slots:
                with slot.lock:
                    if slot.idle and (slot.dead or slot.proc is None or slot.proc.poll() is not None):
                        slot.restart()
                    if slot.idle and not slot.dead:
                        slot.idle = False
                        return slot
            time.sleep(0.05)
        raise RuntimeError("工作进程等待超时")

    def run(self, entry: str, args: tuple, kwargs: dict,
            timeout: float, max_memory_mb: int, progress_cb, cancel_event=None) -> SubprocessResult:
        res = SubprocessResult()
        t0 = time.time()
        kwargs = dict(kwargs or {})
        if progress_cb is not None:
            kwargs[_PROGRESS_MARK] = True
        else:
            kwargs.pop("progress_cb", None)

        if not self._sem.acquire(timeout=queue_timeout()):
            res.error = (f"系统繁忙：并发计算任务已满（上限 {len(self.slots)}），"
                         f"排队超过 {queue_timeout()}s，请稍后重试")
            return res

        shared = _get_semaphore()
        if not shared.acquire(timeout=queue_timeout()):
            self._sem.release()
            res.error = "Excel 执行槽排队超时"
            return res
        slot = None
        params_file = None
        import uuid
        task_id = uuid.uuid4().hex
        try:
            wait_for_memory(queue_timeout(), cancel_event)
            fd, params_file = tempfile.mkstemp(suffix=".pkl", prefix="subproc_params_")
            with os.fdopen(fd, "wb") as f:
                pickle.dump({"entry": entry, "args": args, "kwargs": kwargs,
                             "task_id": task_id}, f, protocol=pickle.HIGHEST_PROTOCOL)

            slot = self._acquire_slot()
            max_memory_mb = memory_limit_mb(max_memory_mb, _process_rss_mb(slot.proc.pid))
            ctx = {"event": threading.Event(), "result_path": None,
                   "error_b64": None, "progress_cb": progress_cb}
            if not slot.submit(task_id, params_file, ctx):
                raise RuntimeError(f"worker {slot.idx} 提交失败")

            # 等待 + 超时/内存监控（0.5s 粒度轮询）
            timed_out = False
            killed_mem = False
            deadline = time.monotonic() + (timeout if (timeout and timeout > 0) else 3600)
            while not ctx["event"].is_set():
                if cancel_event and cancel_event.is_set():
                    res.termination_failed = not slot.kill_and_restart()
                    raise RuntimeError("任务已取消")
                if time.monotonic() > deadline:
                    timed_out = True
                    res.termination_failed = not slot.kill_and_restart()
                    break
                if max_memory_mb and max_memory_mb > 0:
                    try:
                        rss = _process_rss_mb(slot.proc.pid)
                        res.peak_memory_mb = max(res.peak_memory_mb, rss)
                        if rss > memory_limit_mb(max_memory_mb, rss):
                            killed_mem = True
                            res.termination_failed = not slot.kill_and_restart()
                            break
                    except Exception:
                        pass
                ctx["event"].wait(0.5)

            res.duration = time.time() - t0

            if timed_out:
                res.timed_out = True
                res.error = f"执行超时（{timeout}s），已强杀"
                logger.error(f"[subproc-pool/{entry}] {res.error}")
            elif killed_mem:
                res.killed_by_memory = True
                res.error = f"内存超限被杀（上限 {max_memory_mb}MB）"
            elif ctx["error_b64"]:
                try:
                    tb = base64.b64decode(ctx["error_b64"]).decode("utf-8", "replace")
                except Exception:
                    tb = ctx["error_b64"]
                res.error = tb.strip()
            elif ctx["result_path"] and os.path.exists(ctx["result_path"]):
                try:
                    with open(ctx["result_path"], "rb") as f:
                        res.result = pickle.load(f)
                    res.success = True
                except Exception as e:
                    res.error = f"结果反序列化失败: {e}"
                finally:
                    try:
                        os.remove(ctx["result_path"])
                    except Exception:
                        pass
            else:
                res.error = "worker 无结果返回"
        except Exception as e:
            res.error = f"子进程池执行异常: {e}"
            logger.error(f"[subproc-pool/{entry}] {res.error}", exc_info=True)
        finally:
            shared.release()
            self._sem.release()
            if slot is not None:
                with slot.lock:
                    slot.pending.pop(task_id, None)
                    slot.idle = True
            try:
                if not res.termination_failed and params_file and os.path.exists(params_file):
                    os.remove(params_file)
            except Exception:
                pass
        if res.termination_failed:
            res.success = False
            res.error = "进程在强杀后仍未退出，已暂停派发新任务，请检查存储 I/O"
        return res


_pool: Optional[_WorkerPool] = None
_pool_lock = threading.Lock()
_pool_failed = False


def _get_pool() -> Optional[_WorkerPool]:
    """惰性创建常驻 worker 池；创建失败返回 None（调用方回退单任务模式）。"""
    global _pool, _pool_failed
    if _pool is not None or _pool_failed:
        return _pool
    with _pool_lock:
        if _pool is not None or _pool_failed:
            return _pool
        try:
            # 低内存环境默认无常驻池；显式开启后仍与一次性进程共享执行槽。
            size = max(0, env_int("SUBPROCESS_POOL_SIZE", 0))
            if size == 0:
                return None
            wait_for_memory(queue_timeout())
            _pool = _WorkerPool(size)
            logger.info(f"[subproc-pool] 常驻 worker 池已启动（{size} 个，预加载 Aspose）")
            return _pool
        except Exception as e:
            _pool_failed = True
            logger.error(f"[subproc-pool] 启动失败，回退单任务模式: {e}")
            return None


def _run_entry_sync(entry: str, args: tuple, kwargs: dict, progress_cb) -> SubprocessResult:
    """子进程内同步执行（防嵌套死锁）：不 Popen，直接在本进程调用 entry。

    场景：对比/整合的 impl 内部还会调用 run_in_subprocess（如公式计算 _aspose_calc_impl），
    在常驻池里 worker A 执行时嵌套请求池 → 并发满时排队等 queue_timeout(600s) → 卡死。
    检测到 _IN_SUBPROCESS_WORKER 标记后直接同步执行，同进程内调用无并发槽占用。
    """
    import importlib
    import traceback as _tb

    res = SubprocessResult()
    t0 = time.time()
    try:
        kwargs = dict(kwargs or {})
        if progress_cb is not None:
            kwargs[_PROGRESS_MARK] = True
        module_path, _, func_name = entry.partition(":")
        module = importlib.import_module(module_path)
        func = getattr(module, func_name)
        if kwargs.pop(_PROGRESS_MARK, None):
            kwargs["progress_cb"] = progress_cb
        res.result = func(*args, **kwargs)
        res.success = True
    except Exception as e:
        res.error = "".join(_tb.format_exception_only(type(e), e)).strip()
    res.duration = time.time() - t0
    return res


def run_in_subprocess(
    entry: str,
    args: tuple = (),
    kwargs: dict = None,
    timeout: float = 300,
    max_memory_mb: int = None,
    progress_cb=None,
    cancel_event=None,
) -> SubprocessResult:
    """在独立子进程执行 entry 指向的模块级函数，超时/超内存强杀。

    优先走常驻 worker 池（省掉每次启动的 Aspose 初始化开销，20-30s → <1s）；
    池不可用时回退单任务模式（每次新进程）。接口与行为语义不变。

    子进程内（_IN_SUBPROCESS_WORKER 标记）直接同步执行：防池内嵌套请求池死锁。
    """
    max_memory_mb = default_max_memory_mb() if max_memory_mb is None else max_memory_mb
    if os.environ.get("_IN_SUBPROCESS_WORKER") == "1":
        return _run_entry_sync(entry, args, kwargs, progress_cb)
    try:
        ensure_healthy()
    except RuntimeError as exc:
        return SubprocessResult(error=str(exc))
    pool = _get_pool()
    if pool is not None:
        return pool.run(entry, args, kwargs, timeout, max_memory_mb, progress_cb, cancel_event)
    return _run_single(entry, args, kwargs, timeout, max_memory_mb, progress_cb, cancel_event)


async def _async_run(runner, entry, args, kwargs, timeout, max_memory_mb, progress_cb):
    import asyncio
    cancelled = threading.Event()
    future = asyncio.create_task(asyncio.to_thread(
        runner, entry, args, kwargs or {}, timeout, max_memory_mb, progress_cb,
        cancel_event=cancelled))
    try:
        return await asyncio.shield(future)
    except asyncio.CancelledError:
        cancelled.set()
        # The executor owns the reservation until the child has actually exited.
        def consume_result(task):
            if not task.cancelled():
                task.exception()
        future.add_done_callback(consume_result)
        raise


async def run_in_subprocess_async(entry, args=(), kwargs=None, timeout=300,
                                  max_memory_mb=None, progress_cb=None):
    return await _async_run(run_in_subprocess, entry, args, kwargs, timeout,
                            max_memory_mb, progress_cb)


async def run_in_fresh_subprocess_async(entry, args=(), kwargs=None, timeout=300,
                                        max_memory_mb=None, progress_cb=None):
    return await _async_run(run_in_fresh_subprocess, entry, args, kwargs, timeout,
                            max_memory_mb, progress_cb)


def run_in_fresh_subprocess(entry, args=(), kwargs=None, timeout=300,
                            max_memory_mb=None, progress_cb=None, cancel_event=None):
    if os.environ.get("_IN_SUBPROCESS_WORKER") == "1":
        return _run_entry_sync(entry, args, kwargs, progress_cb)
    return _run_single(entry, args, kwargs or {}, timeout, max_memory_mb,
                       progress_cb, cancel_event)
