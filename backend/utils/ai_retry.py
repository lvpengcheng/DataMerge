"""AI 服务瞬时故障的轻量重试工具。"""

import logging
import time
from typing import Callable, Optional


logger = logging.getLogger(__name__)


class AIServiceBusyError(RuntimeError):
    """服务繁忙/暂时不可用，且已用完自动重试。"""


class AITransientResponseError(RuntimeError):
    """HTTP 请求成功，但 AI 网关返回了空或不完整的临时响应。"""


def _status_code(exc: Exception):
    status = getattr(exc, "status_code", None)
    if status is None:
        status = getattr(getattr(exc, "response", None), "status_code", None)
    try:
        return int(status) if status is not None else None
    except (TypeError, ValueError):
        return None


def is_transient_ai_error(exc: Exception) -> bool:
    """判断是否适合原请求重试，不对鉴权/参数错误做无效重试。"""
    if isinstance(exc, AITransientResponseError):
        return True
    status = _status_code(exc)
    if status in (408, 409, 425, 429, 500, 502, 503, 504):
        return True

    text = str(exc).casefold().replace("\\_", "_")
    markers = (
        "server overloaded",
        "service_unavailable",
        "service unavailable",
        "temporarily unavailable",
        "rate limit",
        "rate_limit",
        "too many requests",
        "timeout",
        "timed out",
        "connection error",
        "connection reset",
        "peer closed connection",
        "incomplete chunked read",
        "remoteprotocolerror",
        "server disconnected",
        "connection terminated",
        "connection closed",
        "bad gateway",
        "gateway timeout",
        "empty choices",
        "空 choices",
        "未返回 choices",
    )
    return any(marker in text for marker in markers)


def chat_with_transient_retry(
    provider,
    messages,
    *,
    stage: str = "AI 调用",
    max_attempts: int = 4,
    base_delay: float = 2.0,
    max_delay: float = 15.0,
    emit: Optional[Callable[[dict], None]] = None,
    sleep_fn: Callable[[float], None] = time.sleep,
):
    """调用 provider.chat；遇到 429/5xx/超时等瞬时错误时指数退避重试。"""
    attempts = max(1, min(int(max_attempts or 1), 8))
    retry_total = attempts - 1
    for attempt in range(1, attempts + 1):
        try:
            return provider.chat(messages)
        except Exception as exc:
            transient = is_transient_ai_error(exc)
            if not transient:
                raise
            if attempt >= attempts:
                status = _status_code(exc)
                status_text = f"（HTTP {status}）" if status else ""
                raise AIServiceBusyError(
                    f"AI 服务当前繁忙{status_text}，已自动重试 {retry_total} 次，请稍后再试"
                ) from exc

            delay = min(max(0.0, float(base_delay)) * (2 ** (attempt - 1)),
                        max(0.0, float(max_delay)))
            logger.warning(
                "%s遇到瞬时 AI 服务错误，%.1f 秒后重试（%d/%d）: %s",
                stage, delay, attempt, retry_total, exc,
            )
            if emit:
                emit({
                    "type": "status",
                    "message": (
                        f"{stage}遇到 AI 服务繁忙，{delay:g} 秒后自动重试"
                        f"（{attempt}/{retry_total}）…"
                    ),
                })
            sleep_fn(delay)


def chat_stream_with_transient_retry(
    provider,
    messages,
    *,
    stage: str = "AI 流式调用",
    max_attempts: int = 4,
    base_delay: float = 2.0,
    max_delay: float = 15.0,
    emit: Optional[Callable[[dict], None]] = None,
    chunk_callback: Optional[Callable[[str], None]] = None,
    thinking_callback: Optional[Callable[[str], None]] = None,
    sleep_fn: Callable[[float], None] = time.sleep,
):
    """流式调用 provider.chat_stream，并对连接中断/5xx 做有界重试。

    每次重试都重新收集完整响应；调用方的 chunk 回调仅用于显示进度，
    最终结果以最后一次成功调用为准，避免把失败尝试的半截内容拼进去。
    """
    attempts = max(1, min(int(max_attempts or 1), 8))
    retry_total = attempts - 1
    for attempt in range(1, attempts + 1):
        try:
            received = []

            def _on_chunk(chunk):
                if chunk:
                    received.append(chunk)
                    if chunk_callback:
                        chunk_callback(chunk)

            result = provider.chat_stream(
                messages,
                chunk_callback=_on_chunk,
                thinking_callback=thinking_callback,
            )
            return result or "".join(received)
        except Exception as exc:
            transient = is_transient_ai_error(exc)
            if not transient:
                raise
            if attempt >= attempts:
                status = _status_code(exc)
                status_text = f"（HTTP {status}）" if status else ""
                raise AIServiceBusyError(
                    f"AI 服务当前繁忙或流式连接中断{status_text}，"
                    f"已自动重试 {retry_total} 次，请稍后再试"
                ) from exc

            delay = min(max(0.0, float(base_delay)) * (2 ** (attempt - 1)),
                        max(0.0, float(max_delay)))
            logger.warning(
                "%s遇到瞬时 AI 流式错误，%.1f 秒后重试（%d/%d）: %s",
                stage, delay, attempt, retry_total, exc,
            )
            if emit:
                emit({
                    "type": "status",
                    "message": (
                        f"{stage}连接中断，{delay:g} 秒后自动重试"
                        f"（{attempt}/{retry_total}）…"
                    ),
                })
            sleep_fn(delay)
