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

logger = logging.getLogger(__name__)

_PROGRESS_MARK = "__SUBPROC_PROGRESS__"   # 与 subprocess_worker 中一致
_PROJ_ROOT = str(Path(__file__).resolve().parent.parent.parent)

# 并发上限：最多同时跑 N 个子进程，防止多任务并发时子进程内存叠加（N × 1GB 以内可控）。
# 其余调用排队等待（排队有上限 SUBPROCESS_QUEUE_TIMEOUT，满负载返回"系统繁忙"而非无限等）。
# N 用 lazy 信号量：首次使用时才从 .env SUBPROCESS_CONCURRENCY 读取（默认 2），
# 避免模块 import 时 .env 尚未加载（load_dotenv 在应用启动早期执行）导致配置读不到。
_semaphore = None
_semaphore_lock = threading.Lock()


def _get_semaphore() -> threading.Semaphore:
    """lazy 创建并发信号量：读 .env SUBPROCESS_CONCURRENCY（默认 2）。"""
    global _semaphore
    if _semaphore is None:
        with _semaphore_lock:
            if _semaphore is None:
                _semaphore = threading.Semaphore(env_int("SUBPROCESS_CONCURRENCY", 2))
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
    """默认子进程内存上限（MB），由 .env SUBPROCESS_MAX_MEMORY_MB 控制，默认 4096。"""
    return env_int("SUBPROCESS_MAX_MEMORY_MB", 4096)


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


def _kill_tree(pid: int) -> None:
    """强杀进程树。Windows 用 taskkill /F /T（连子进程一起杀）；其他平台 SIGKILL。"""
    if sys.platform == "win32":
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True, timeout=10,
            )
            return
        except Exception as e:
            logger.warning(f"[subproc] taskkill 失败（回退 TerminateProcess）: {e}")
            import ctypes
            try:
                handle = ctypes.windll.kernel32.OpenProcess(0x0001, False, pid)  # PROCESS_TERMINATE
                if handle:
                    ctypes.windll.kernel32.TerminateProcess(handle, 1)
                    ctypes.windll.kernel32.CloseHandle(handle)
            except Exception:
                pass
            return
    try:
        os.kill(pid, 9)
    except Exception:
        pass


@dataclass
class SubprocessResult:
    """子进程执行结果。success=False 时 result=None，error 含明确原因。"""
    success: bool = False
    result: object = None
    error: str = ""
    timed_out: bool = False          # 超时被杀
    killed_by_memory: bool = False   # 内存超限被杀
    duration: float = 0.0
    log_lines: list = field(default_factory=list)

    @property
    def killed(self) -> bool:
        return self.timed_out or self.killed_by_memory


def run_in_subprocess(
    entry: str,
    args: tuple = (),
    kwargs: dict = None,
    timeout: float = 300,
    max_memory_mb: int = None,
    progress_cb=None,
) -> SubprocessResult:
    """在独立子进程执行 entry 指向的模块级函数，超时/超内存强杀。

    Args:
        entry: "module.path:func_name"
        args/kwargs: 函数参数（必须可 pickle；kwargs 中的 progress_cb 会被移除，
            以标记替换后由子进程注入真回调）
        timeout: 超时秒数（0/None = 不限）
        max_memory_mb: 内存上限 MB，超限强杀；None 用 .env SUBPROCESS_MAX_MEMORY_MB（默认 4096）
        progress_cb: 进度回调（函数签名须支持 progress_cb= 参数），子进程打 @@PROG@@ 后转发

    Returns:
        SubprocessResult
    """
    max_memory_mb = default_max_memory_mb() if max_memory_mb is None else max_memory_mb
    kwargs = dict(kwargs or {})
    if progress_cb is not None:
        kwargs[_PROGRESS_MARK] = True   # 移除真函数（不可 pickle），子进程侧替换为 stdout 包装
    else:
        kwargs.pop("progress_cb", None)

    res = SubprocessResult()
    t0 = time.time()
    state = {"result_path": None, "error_b64": None, "start_error": None,
             "killed_by_memory": False}
    done = threading.Event()

    # 并发上限排队：等待空位（最多 queue_timeout 秒，满了返回"系统繁忙"而非无限等）。
    # 注意：本函数是同步阻塞的——async 端点请用 run_in_subprocess_async，否则会冻结事件循环。
    _sem = _get_semaphore()
    if not _sem.acquire(timeout=queue_timeout()):
        res.error = (f"系统繁忙：并发计算任务已满（上限 "
                     f"{env_int('SUBPROCESS_CONCURRENCY', 2)}），排队超过 "
                     f"{queue_timeout()}s，请稍后重试")
        logger.warning(f"[subproc/{entry}] {res.error}")
        return res
    proc = None
    params_file = None
    mem_thread = None
    try:
        # 1) 参数 pickle 到临时文件
        fd, params_file = tempfile.mkstemp(suffix=".json", prefix="subproc_params_")
        with os.fdopen(fd, "wb") as f:
            pickle.dump({"entry": entry, "args": args, "kwargs": kwargs}, f,
                        protocol=pickle.HIGHEST_PROTOCOL)

        # 2) 主线程启动子进程（reader 线程只读管道，避免 proc 就绪竞态）
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"   # 与智算子进程一致，避免 Windows GBK 乱码
            env["PYTHONUTF8"] = "1"
            proc = subprocess.Popen(
                [sys.executable, "-u", "-m", "backend.utils.subprocess_worker", params_file],
                cwd=_PROJ_ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=env,
            )
        except Exception as e:
            state["start_error"] = e

        def _reader():
            try:
                for raw in proc.stdout:
                    line = raw.decode("utf-8", "replace").rstrip("\r\n")
                    if line.startswith("@@PROG@@"):
                        msg = line[len("@@PROG@@"):]
                        if progress_cb:
                            try:
                                progress_cb(msg)
                            except Exception:
                                pass
                    elif line.startswith("@@RESULT@@"):
                        state["result_path"] = line[len("@@RESULT@@"):]
                    elif line.startswith("@@ERROR@@"):
                        state["error_b64"] = line[len("@@ERROR@@"):]
                    elif line.strip():
                        res.log_lines.append(line)
                        logger.debug(f"[subproc/{entry}] {line}")
                proc.wait()
            except Exception as e:
                state["start_error"] = e
            finally:
                done.set()

        _t = threading.Thread(target=_reader, name=f"subproc-{entry.rsplit(':', 1)[-1]}", daemon=True)
        _t.start()

        if state["start_error"] is not None:
            done.set()   # 启动失败：让监控线程退出，走错误分支

        # 3) 内存监控线程：每秒轮询 RSS，超限强杀
        def _mem_watchdog():
            while not done.is_set():
                if proc is None or proc.poll() is not None:
                    return
                try:
                    rss = _process_rss_mb(proc.pid)
                    if rss > max_memory_mb:
                        logger.error(
                            f"[subproc/{entry}] 内存超限 {rss:.0f}MB > {max_memory_mb}MB，强杀")
                        state["killed_by_memory"] = True
                        _kill_tree(proc.pid)
                        return
                except Exception:
                    pass
                time.sleep(0.5)

        if max_memory_mb and max_memory_mb > 0:
            mem_thread = threading.Thread(target=_mem_watchdog,
                                          name=f"subproc-mem-{entry.rsplit(':', 1)[-1]}",
                                          daemon=True)
            mem_thread.start()

        # 4) 等待完成（真超时）
        timed_out = False
        if state["start_error"] is None:
            if timeout and timeout > 0:
                try:
                    proc.wait(timeout=timeout)
                except subprocess.TimeoutExpired:
                    timed_out = True
                    logger.error(f"[subproc/{entry}] 执行超时（{timeout}s），强杀")
                    _kill_tree(proc.pid)
                    proc.wait()   # 收尸；reader 线程随后因 stdout EOF 自然结束
            else:
                proc.wait()

        done.wait()

        res.duration = time.time() - t0

        # 4) 结果/错误
        if state["start_error"] is not None:
            res.error = f"子进程启动/读取失败: {state['start_error']}"
            logger.error(f"[subproc/{entry}] {res.error}")
        elif state["killed_by_memory"]:
            res.killed_by_memory = True
            res.error = f"内存超限被杀（上限 {max_memory_mb}MB）"
        elif timed_out:
            res.timed_out = True
            res.error = f"执行超时（{timeout}s），已强杀"
        elif state["error_b64"]:
            try:
                tb = base64.b64decode(state["error_b64"]).decode("utf-8", "replace")
            except Exception:
                tb = state["error_b64"]
            res.error = tb.strip()
            logger.error(f"[subproc/{entry}] 子进程异常: {tb.strip().splitlines()[-1] if tb.strip() else '未知'}")
        elif state["result_path"] and os.path.exists(state["result_path"]):
            try:
                with open(state["result_path"], "rb") as f:
                    res.result = pickle.load(f)
                res.success = True
            except Exception as e:
                res.error = f"结果反序列化失败: {e}"
                logger.error(f"[subproc/{entry}] {res.error}")
            finally:
                try:
                    os.remove(state["result_path"])
                except Exception:
                    pass
        else:
            rc = proc.returncode if proc is not None else None
            res.error = f"子进程退出(code={rc})，无结果"
            logger.error(f"[subproc/{entry}] {res.error}")
    finally:
        _get_semaphore().release()
        try:
            if params_file and os.path.exists(params_file):
                os.remove(params_file)
        except Exception:
            pass
        if proc is not None and proc.poll() is None:
            _kill_tree(proc.pid)

    return res


async def run_in_subprocess_async(entry: str, args: tuple = (), kwargs: dict = None,
                                  timeout: float = 300, max_memory_mb: int = None,
                                  progress_cb=None) -> SubprocessResult:
    """async 版 run_in_subprocess：在后台线程执行，不冻结事件循环。

    用于 async 端点（整合对比/合并/智算等）。事件循环内直接调同步版会在
    Popen 等待/排队期间冻结所有用户的请求——多客户并发时表现为整体卡死。
    """
    import asyncio
    return await asyncio.to_thread(
        run_in_subprocess, entry, args, kwargs or {},
        timeout=timeout, max_memory_mb=max_memory_mb, progress_cb=progress_cb,
    )
