"""上传文件分块落盘。"""

import asyncio
import os
from pathlib import Path


UPLOAD_CHUNK_SIZE = 1024 * 1024
_excel_work_semaphore = None


def get_excel_work_semaphore():
    """进程级 Excel 重任务闸门：基础资料、智算、智训共享可配置并发。"""
    global _excel_work_semaphore
    if _excel_work_semaphore is None:
        try:
            concurrency = int(os.getenv("EXCEL_WORK_CONCURRENCY", "3"))
        except (TypeError, ValueError):
            concurrency = 3
        # 防止误配置无限并发；需要更高吞吐应先扩大容器/VM内存。
        _excel_work_semaphore = asyncio.Semaphore(max(1, min(concurrency, 5)))
    return _excel_work_semaphore


async def save_upload_file(upload, destination, chunk_size: int = UPLOAD_CHUNK_SIZE) -> int:
    """把 FastAPI UploadFile 分块写入目标文件，并显式关闭上传句柄。"""
    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    written = 0
    try:
        with path.open("wb") as output:
            while True:
                chunk = await upload.read(chunk_size)
                if not chunk:
                    break
                output.write(chunk)
                written += len(chunk)
    finally:
        await upload.close()
    return written


def safe_upload_name(filename: str, fallback: str = "upload.xlsx") -> str:
    """只保留文件名，阻止 multipart 文件名携带目录穿越。"""
    return Path(filename or fallback).name
