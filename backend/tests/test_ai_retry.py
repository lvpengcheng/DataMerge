"""AI 瞬时服务故障重试回归测试。"""

import pytest

from backend.utils.ai_retry import (
    AIServiceBusyError,
    AITransientResponseError,
    chat_stream_with_transient_retry,
    chat_with_transient_retry,
    is_transient_ai_error,
)


class _StatusError(Exception):
    def __init__(self, status_code, message):
        super().__init__(message)
        self.status_code = status_code


class _Provider:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def chat(self, _messages):
        outcome = self.outcomes[self.calls]
        self.calls += 1
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


class _StreamProvider:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def chat_stream(self, _messages, chunk_callback=None, thinking_callback=None):
        outcome = self.outcomes[self.calls]
        self.calls += 1
        if isinstance(outcome, Exception):
            if chunk_callback:
                chunk_callback("失败尝试的半截")
            raise outcome
        if thinking_callback:
            thinking_callback("分析摘要")
        if chunk_callback:
            for chunk in outcome:
                chunk_callback(chunk)
        return "".join(outcome)


def test_retries_503_then_returns_result():
    provider = _Provider([
        _StatusError(503, "Server Overloaded"),
        _StatusError(503, "service_unavailable_error"),
        "最终规则",
    ])
    sleeps, events = [], []
    result = chat_with_transient_retry(
        provider, [], stage="最终规则", max_attempts=4,
        base_delay=2, emit=events.append, sleep_fn=sleeps.append,
    )
    assert result == "最终规则"
    assert provider.calls == 3
    assert sleeps == [2, 4]
    assert len(events) == 2 and all(e["type"] == "status" for e in events)


def test_retries_deepseek_empty_choices_response():
    provider = _Provider([
        AITransientResponseError("AI 服务返回空 choices[]"),
        "正常结果",
    ])
    assert chat_with_transient_retry(
        provider, [], max_attempts=3, base_delay=0, sleep_fn=lambda _s: None,
    ) == "正常结果"
    assert provider.calls == 2


def test_non_transient_error_is_not_retried():
    provider = _Provider([ValueError("参数错误")])
    with pytest.raises(ValueError, match="参数错误"):
        chat_with_transient_retry(provider, [], max_attempts=4, sleep_fn=lambda _s: None)
    assert provider.calls == 1


def test_exhausted_503_returns_clear_busy_message():
    provider = _Provider([_StatusError(503, "Server Overloaded")] * 3)
    with pytest.raises(AIServiceBusyError, match="HTTP 503.*已自动重试 2 次"):
        chat_with_transient_retry(
            provider, [], max_attempts=3, base_delay=0, sleep_fn=lambda _s: None,
        )
    assert provider.calls == 3


def test_escaped_service_unavailable_code_is_transient():
    assert is_transient_ai_error(
        RuntimeError("Error code: 503 - service\\_unavailable\\_error")
    )


def test_stream_retry_returns_only_successful_attempt_result():
    provider = _StreamProvider([
        RuntimeError("peer closed connection without sending complete message body"),
        ["最终", "规则"],
    ])
    progress, thoughts = [], []
    result = chat_stream_with_transient_retry(
        provider, [], max_attempts=3, base_delay=0,
        chunk_callback=progress.append, thinking_callback=thoughts.append,
        sleep_fn=lambda _s: None,
    )
    assert result == "最终规则"
    assert provider.calls == 2
    assert progress[-2:] == ["最终", "规则"]
    assert thoughts == ["分析摘要"]
