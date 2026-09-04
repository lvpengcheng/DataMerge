"""AI 模型请求参数兼容性回归测试。"""

import os
from contextlib import contextmanager
from types import SimpleNamespace

from backend.ai_engine.ai_provider import (
    ClaudeProvider,
    _is_claude_5_plus_model,
    _model_omits_temperature,
)


@contextmanager
def _without_global_temperature_override():
    old = os.environ.pop("AI_DISABLE_TEMPERATURE", None)
    try:
        yield
    finally:
        if old is not None:
            os.environ["AI_DISABLE_TEMPERATURE"] = old


def test_claude_opus_5_omits_temperature():
    with _without_global_temperature_override():
        assert _model_omits_temperature("claude-opus-5")
        assert _model_omits_temperature("claude-sonnet-5-20260901")
        assert _model_omits_temperature("anthropic.claude-haiku-6-v1:0")
        assert _model_omits_temperature("claude-5-1")


def test_older_claude_model_keeps_existing_temperature_behavior():
    with _without_global_temperature_override():
        assert not _model_omits_temperature("claude-3-sonnet-20240229")
        assert not _model_omits_temperature("claude-3-opus")


def test_claude_5_model_detection_ignores_version_dates():
    assert _is_claude_5_plus_model("claude-opus-5")
    assert _is_claude_5_plus_model("anthropic.claude-sonnet-6-v1:0")
    assert not _is_claude_5_plus_model("claude-3-sonnet-20240229")


class _FakeStream:
    def __init__(self, texts=None):
        self.text_stream = iter(texts or ["规则内容"])

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def get_final_message(self):
        return SimpleNamespace(stop_reason="end_turn")


class _RawClaudeEventStream:
    """模拟新版 Anthropic SDK：raw delta 后再发一个便捷合成事件。"""

    def __init__(self):
        self._events = [
            SimpleNamespace(
                type="content_block_delta",
                delta=SimpleNamespace(type="thinking_delta", thinking="先分析列关系"),
            ),
            SimpleNamespace(type="thinking", thinking="先分析列关系"),
            SimpleNamespace(
                type="content_block_delta",
                delta=SimpleNamespace(type="signature_delta", signature="hidden"),
            ),
            SimpleNamespace(
                type="content_block_delta",
                delta=SimpleNamespace(type="text_delta", text="def fill_template():\n"),
            ),
            SimpleNamespace(type="text", text="def fill_template():\n"),
            SimpleNamespace(
                type="content_block_delta",
                delta=SimpleNamespace(type="text_delta", text="    return True\n"),
            ),
            SimpleNamespace(type="text", text="    return True\n"),
        ]

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def __iter__(self):
        return iter(self._events)

    def get_final_message(self):
        return SimpleNamespace(stop_reason="end_turn")


def test_claude_opus_5_stream_does_not_send_temperature():
    captured = {}

    def _stream(**kwargs):
        captured.update(kwargs)
        return _FakeStream()

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://api.anthropic.com"
    provider.anthropic_version = ""
    provider._client = SimpleNamespace(
        messages=SimpleNamespace(stream=_stream),
    )

    with _without_global_temperature_override():
        chunks = list(provider._claude_chat_stream(
            "系统提示", [{"role": "user", "content": "整理规则"}], max_tokens=100,
        ))
    assert chunks[-1] == ("", "end_turn")
    assert "temperature" not in captured


def test_claude_opus_5_stream_sends_adaptive_summarized_thinking():
    captured = {}

    def _stream(**kwargs):
        captured.update(kwargs)
        return _FakeStream()

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 64000
    provider.base_url = "https://api.anthropic.com"
    provider.anthropic_version = ""
    provider.thinking = {"type": "adaptive", "display": "summarized"}
    provider._client = SimpleNamespace(messages=SimpleNamespace(stream=_stream))

    list(provider._claude_chat_stream(
        "系统提示", [{"role": "user", "content": "生成代码"}],
    ))
    assert captured["thinking"] == {
        "type": "adaptive",
        "display": "summarized",
    }
    assert "temperature" not in captured


def test_claude_non_stream_uses_text_blocks_after_thinking_block():
    captured = {}

    def _create(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(
            content=[
                SimpleNamespace(type="thinking", thinking="内部摘要"),
                SimpleNamespace(type="text", text="正式代码"),
            ],
            stop_reason="end_turn",
        )

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 64000
    provider.base_url = "https://api.anthropic.com"
    provider.anthropic_version = ""
    provider.thinking = {"type": "adaptive", "display": "summarized"}
    provider._client = SimpleNamespace(messages=SimpleNamespace(create=_create))

    content, stop_reason = provider._claude_chat(
        "系统提示", [{"role": "user", "content": "生成代码"}],
    )
    assert (content, stop_reason) == ("正式代码", "end_turn")
    assert captured["thinking"]["type"] == "adaptive"


def test_claude_stream_separates_thinking_and_text_without_duplicates():
    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://claude.forward-thinking.example"
    provider.anthropic_version = "bedrock-2023-05-31"
    provider._client = SimpleNamespace(
        messages=SimpleNamespace(stream=lambda **_kwargs: _RawClaudeEventStream()),
    )

    thoughts = []
    chunks = list(provider._claude_chat_stream(
        "系统提示",
        [{"role": "user", "content": "生成代码"}],
        max_tokens=100,
        thinking_callback=thoughts.append,
    ))

    assert thoughts == ["先分析列关系"]
    assert "".join(text for text, _ in chunks) == "def fill_template():\n    return True\n"
    assert chunks[-1] == ("", "end_turn")


class _EnterErrorStream:
    def __init__(self, exc):
        self.exc = exc

    def __enter__(self):
        raise self.exc

    def __exit__(self, *_args):
        return False


def test_aws_forward_stream_retries_with_anthropic_version():
    calls = []

    def _stream(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            return _EnterErrorStream(RuntimeError(
                "aws_invoke_error: Bedrock Runtime ValidationException: "
                "anthropic_version: Field required"
            ))
        return _FakeStream()

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://claude.forward.example"
    provider.anthropic_version = ""
    provider._client = SimpleNamespace(messages=SimpleNamespace(stream=_stream))

    chunks = list(provider._claude_chat_stream(
        "系统提示", [{"role": "user", "content": "整理规则"}], max_tokens=100,
    ))
    assert chunks[-1] == ("", "end_turn")
    assert len(calls) == 2
    assert "extra_body" not in calls[0]
    assert calls[1]["extra_body"] == {
        "anthropic_version": "bedrock-2023-05-31",
    }


def test_aws_forward_non_stream_retries_with_anthropic_version():
    calls = []

    def _create(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            raise RuntimeError(
                "aws_invoke_error: InvokeModel Bedrock ValidationException: "
                "anthropic\\_version: Field required"
            )
        return SimpleNamespace(
            content=[SimpleNamespace(text="规则内容")], stop_reason="end_turn",
        )

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://claude.forward-non-stream.example"
    provider.anthropic_version = ""
    provider._client = SimpleNamespace(messages=SimpleNamespace(create=_create))

    content, stop_reason = provider._claude_chat(
        "系统提示", [{"role": "user", "content": "整理规则"}], max_tokens=100,
    )
    assert (content, stop_reason) == ("规则内容", "end_turn")
    assert len(calls) == 2
    assert calls[1]["extra_body"]["anthropic_version"] == "bedrock-2023-05-31"


def test_aws_forward_downgrades_unsupported_thinking_display():
    calls = []

    def _stream(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            return _EnterErrorStream(RuntimeError(
                "Error code: 400 ValidationException: thinking.display is not supported"
            ))
        return _FakeStream()

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 64000
    provider.base_url = "https://claude.aws-forward.example"
    provider.anthropic_version = "bedrock-2023-05-31"
    provider.thinking = {"type": "adaptive", "display": "summarized"}
    provider._client = SimpleNamespace(messages=SimpleNamespace(stream=_stream))

    chunks = list(provider._claude_chat_stream(
        "系统提示", [{"role": "user", "content": "生成代码"}],
    ))
    assert chunks[-1] == ("", "end_turn")
    assert calls[0]["thinking"]["display"] == "summarized"
    assert calls[1]["thinking"] == {"type": "adaptive"}


class _DisconnectingTextStream:
    def __init__(self):
        self.step = 0

    def __iter__(self):
        return self

    def __next__(self):
        if self.step == 0:
            self.step += 1
            return "代码前半段"
        raise RuntimeError(
            "peer closed connection without sending complete message body "
            "(incomplete chunked read)"
        )


class _PartialDisconnectStream(_FakeStream):
    def __init__(self):
        self.text_stream = _DisconnectingTextStream()


@contextmanager
def _without_transport_wait():
    old_delay = os.environ.get("CLAUDE_TRANSPORT_RETRY_DELAY")
    old_attempts = os.environ.get("CLAUDE_TRANSPORT_MAX_ATTEMPTS")
    os.environ["CLAUDE_TRANSPORT_RETRY_DELAY"] = "0"
    os.environ["CLAUDE_TRANSPORT_MAX_ATTEMPTS"] = "3"
    try:
        yield
    finally:
        if old_delay is None:
            os.environ.pop("CLAUDE_TRANSPORT_RETRY_DELAY", None)
        else:
            os.environ["CLAUDE_TRANSPORT_RETRY_DELAY"] = old_delay
        if old_attempts is None:
            os.environ.pop("CLAUDE_TRANSPORT_MAX_ATTEMPTS", None)
        else:
            os.environ["CLAUDE_TRANSPORT_MAX_ATTEMPTS"] = old_attempts


def test_claude_stream_resumes_after_incomplete_chunked_read():
    calls = []

    def _stream(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            return _PartialDisconnectStream()
        return _FakeStream(["代码后半段"])

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://claude.forward-stream.example"
    provider.anthropic_version = "bedrock-2023-05-31"
    provider._client = SimpleNamespace(messages=SimpleNamespace(stream=_stream))

    with _without_transport_wait():
        chunks = list(provider._claude_chat_stream(
            "系统提示", [{"role": "user", "content": "生成代码"}], max_tokens=100,
        ))
    assert "".join(text for text, _ in chunks) == "代码前半段代码后半段"
    assert len(calls) == 2
    retry_messages = calls[1]["messages"]
    assert retry_messages[-2] == {"role": "assistant", "content": "代码前半段"}
    assert "不要重复" in retry_messages[-1]["content"]


def test_claude_non_stream_retries_incomplete_chunked_read():
    calls = []

    def _create(**kwargs):
        calls.append(kwargs)
        if len(calls) == 1:
            raise RuntimeError(
                "peer closed connection without sending complete message body "
                "(incomplete chunked read)"
            )
        return SimpleNamespace(
            content=[SimpleNamespace(text="完整代码")], stop_reason="end_turn",
        )

    provider = object.__new__(ClaudeProvider)
    provider.model = "claude-opus-5"
    provider.max_tokens = 1000
    provider.base_url = "https://claude.forward-create.example"
    provider.anthropic_version = "bedrock-2023-05-31"
    provider._client = SimpleNamespace(messages=SimpleNamespace(create=_create))

    with _without_transport_wait():
        content, stop_reason = provider._claude_chat(
            "系统提示", [{"role": "user", "content": "生成代码"}], max_tokens=100,
        )
    assert (content, stop_reason) == ("完整代码", "end_turn")
    assert len(calls) == 2
