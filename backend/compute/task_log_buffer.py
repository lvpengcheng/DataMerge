"""内存级日志缓冲区，支持 SSE 断线重连。

每个计算任务的日志事件存储在内存中，SSE 读者可以从任意 offset 续读。
任务完成后保留 30 分钟，之后自动清理。
"""
import asyncio
import logging
import threading
import time
from typing import AsyncGenerator, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


class TaskLogBuffer:
    """进程全局单例，线程安全的计算任务日志缓冲区。"""

    _instance: Optional["TaskLogBuffer"] = None

    @classmethod
    def get_instance(cls) -> "TaskLogBuffer":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self._buffers: Dict[str, List[Tuple[int, str]]] = {}   # task_id -> [(event_id, json_str)]
        self._finished: Dict[str, bool] = {}                    # task_id -> done?
        self._waiters: Dict[str, List[asyncio.Event]] = {}      # task_id -> [asyncio.Event]
        self._finish_times: Dict[str, float] = {}               # task_id -> finish timestamp
        self._lock = threading.Lock()
        self._TTL = 1800  # 30 分钟后清理

    def create_task(self, task_id: str):
        """为新任务初始化缓冲区。"""
        with self._lock:
            self._buffers[task_id] = []
            self._finished[task_id] = False
            self._waiters[task_id] = []
        logger.info(f"[TaskLogBuffer] 创建任务缓冲区: {task_id}")

    def push(self, task_id: str, event_json: str) -> int:
        """线程安全推送事件。返回 event_id (从 1 开始)，-1 表示任务不存在。"""
        with self._lock:
            buf = self._buffers.get(task_id)
            if buf is None:
                return -1
            event_id = len(buf) + 1
            buf.append((event_id, event_json))
            waiters = list(self._waiters.get(task_id, []))

        # 在锁外唤醒等待者
        self._wake_waiters(waiters)
        return event_id

    def finish(self, task_id: str):
        """标记任务完成，唤醒所有读者。"""
        with self._lock:
            self._finished[task_id] = True
            self._finish_times[task_id] = time.time()
            waiters = list(self._waiters.get(task_id, []))

        self._wake_waiters(waiters)
        logger.info(f"[TaskLogBuffer] 任务完成: {task_id}")

    async def read_from(self, task_id: str, from_id: int = 0) -> AsyncGenerator[Tuple[int, Optional[str]], None]:
        """异步生成器：从 from_id 开始读事件。

        产出 (event_id, json_str) 为正常事件，(-1, None) 为心跳信号。
        任务完成且所有事件读完后返回。
        """
        waiter = asyncio.Event()
        with self._lock:
            if task_id not in self._buffers:
                return
            self._waiters[task_id].append(waiter)

        try:
            cursor = from_id
            while True:
                # 读取当前所有新事件
                with self._lock:
                    buf = self._buffers.get(task_id, [])
                    finished = self._finished.get(task_id, False)

                while cursor < len(buf):
                    event_id, data = buf[cursor]
                    yield event_id, data
                    cursor += 1

                if finished:
                    break

                # 等待新事件或超时（心跳）
                waiter.clear()
                try:
                    await asyncio.wait_for(waiter.wait(), timeout=5)
                except asyncio.TimeoutError:
                    yield -1, None  # 心跳信号
        finally:
            with self._lock:
                wl = self._waiters.get(task_id, [])
                if waiter in wl:
                    wl.remove(waiter)

    def get_status(self, task_id: str) -> Optional[dict]:
        """获取任务缓冲区状态。返回 None 表示缓冲区不存在（已过期或未创建）。"""
        with self._lock:
            if task_id not in self._buffers:
                return None
            return {
                "event_count": len(self._buffers[task_id]),
                "finished": self._finished.get(task_id, False),
            }

    def cleanup_expired(self):
        """清理已过 TTL 的已完成任务缓冲区。"""
        now = time.time()
        with self._lock:
            expired = [
                tid for tid, ts in self._finish_times.items()
                if now - ts > self._TTL
            ]
            for tid in expired:
                self._buffers.pop(tid, None)
                self._finished.pop(tid, None)
                self._waiters.pop(tid, None)
                self._finish_times.pop(tid, None)

        if expired:
            logger.info(f"[TaskLogBuffer] 清理过期缓冲区: {expired}")

    @staticmethod
    def _wake_waiters(waiters: list):
        """线程安全地唤醒 asyncio.Event 等待者。"""
        for w in waiters:
            try:
                loop = asyncio.get_running_loop()
                loop.call_soon_threadsafe(w.set)
            except RuntimeError:
                # 没有运行中的 event loop，直接 set
                w.set()
