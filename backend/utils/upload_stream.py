"""上传文件分块落盘。"""

import asyncio
import os
from pathlib import Path


UPLOAD_CHUNK_SIZE = 1024 * 1024
_excel_work_semaphore = None


class ExcelWorkGate(asyncio.Semaphore):
    """FIFO semaphore with finite queue length/wait and memory admission."""

    def __init__(self, value=1):
        super().__init__(value)
        self.pending = 0

    async def acquire(self):
        from backend.utils.resource_guard import admission_ready, ensure_healthy
        ensure_healthy()
        if self.pending >= max(1, int(os.getenv("EXCEL_QUEUE_MAX_PENDING", "32"))):
            raise RuntimeError("Excel 等待队列已满，请稍后提交")
        self.pending += 1
        acquired = False
        try:
            async def wait():
                nonlocal acquired
                await super(ExcelWorkGate, self).acquire()
                acquired = True
                while not admission_ready():
                    await asyncio.sleep(0.5)
                return True
            return await asyncio.wait_for(wait(), max(1, int(os.getenv("EXCEL_QUEUE_MAX_WAIT", "1800"))))
        except BaseException:
            if acquired:
                self.release()
            raise
        finally:
            self.pending -= 1


def get_excel_work_semaphore():
    """进程级 Excel 重任务闸门：基础资料、智算、智训共享可配置并发。"""
    global _excel_work_semaphore
    if _excel_work_semaphore is None:
        try:
            concurrency = int(os.getenv("EXCEL_WORK_CONCURRENCY", "1"))
        except (TypeError, ValueError):
            concurrency = 1
        # 防止误配置无限并发；需要更高吞吐应先扩大容器/VM内存。
        _excel_work_semaphore = ExcelWorkGate(max(1, min(concurrency, 5)))
    return _excel_work_semaphore


async def save_upload_file(upload, destination, chunk_size: int = UPLOAD_CHUNK_SIZE) -> int:
    """把 FastAPI UploadFile 分块写入目标文件，并显式关闭上传句柄。"""
    path = Path(destination)
    await asyncio.to_thread(path.parent.mkdir, parents=True, exist_ok=True)
    written = 0
    try:
        output = await asyncio.to_thread(path.open, "wb")
        try:
            while True:
                chunk = await upload.read(chunk_size)
                if not chunk:
                    break
                await asyncio.to_thread(output.write, chunk)
                written += len(chunk)
        finally:
            await asyncio.to_thread(output.close)
    finally:
        await upload.close()
    return written


def safe_upload_name(filename: str, fallback: str = "upload.xlsx") -> str:
    """只保留文件名，阻止 multipart 文件名携带目录穿越。"""
    return Path(filename or fallback).name
