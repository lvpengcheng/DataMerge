"""OpenAI SDK 3.x / Responses API 兼容性回归测试。"""

from types import SimpleNamespace

from backend.ai_engine.ai_provider import OpenAIProvider, _is_openai_reasoning_model


def _provider(model="gpt-5.2", base_url="https://api.openai.com/v1", api_mode="auto"):
    provider = object.__new__(OpenAIProvider)
    provider.model = model
    provider.base_url = base_url
    provider.api_mode = api_mode
    provider.max_tokens = 64000
    provider.reasoning_effort = "medium"
    provider.reasoning_summary = "auto"
    provider.verbosity = "medium"
    return provider


def test_reasoning_model_detection_does_not_match_gpt4():
    assert _is_openai_reasoning_model("gpt-5.2")
    assert _is_openai_reasoning_model("openai/gpt-5-mini")
    assert _is_openai_reasoning_model("o3")
    assert not _is_openai_reasoning_model("gpt-4.1")


def test_auto_uses_responses_only_for_official_reasoning_models():
    assert _provider()._use_responses_api()
    assert not _provider(model="gpt-4.1")._use_responses_api()
    assert not _provider(base_url="https://gateway.example/v1")._use_responses_api()


def test_responses_non_stream_uses_new_parameters_and_output_text():
    captured = {}
    provider = _provider()

    def _create(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(
            output_text="最终内容",
            status="completed",
            incomplete_details=None,
        )

    provider._client = SimpleNamespace(
        responses=SimpleNamespace(create=_create),
    )
    content, finish_reason = provider._openai_chat(
        [{"role": "user", "content": "生成代码"}], temperature=0.1,
    )

    assert (content, finish_reason) == ("最终内容", "stop")
    assert captured["max_output_tokens"] == 64000
    assert captured["reasoning"] == {"effort": "medium", "summary": "auto"}
    assert captured["text"] == {"verbosity": "medium"}
    assert "temperature" not in captured
    assert "max_tokens" not in captured


def test_responses_stream_separates_reasoning_summary_and_text():
    events = [
        SimpleNamespace(type="response.reasoning_summary_text.delta", delta="先分析"),
        SimpleNamespace(type="response.output_text.delta", delta="print("),
        SimpleNamespace(type="response.output_text.delta", delta="1)"),
        SimpleNamespace(
            type="response.incomplete",
            response=SimpleNamespace(
                status="incomplete",
                incomplete_details=SimpleNamespace(reason="max_output_tokens"),
            ),
        ),
    ]
    provider = _provider()
    provider._client = SimpleNamespace(
        responses=SimpleNamespace(create=lambda **_kwargs: iter(events)),
    )
    thoughts = []
    chunks = list(provider._openai_chat_stream(
        [{"role": "user", "content": "生成代码"}],
        thinking_callback=thoughts.append,
    ))

    assert thoughts == ["先分析"]
    assert "".join(text for text, _ in chunks) == "print(1)"
    assert chunks[-1] == ("", "length")


def test_chat_fallback_uses_max_completion_tokens_for_gpt5():
    captured = {}
    provider = _provider(base_url="https://gateway.example/v1")

    def _create(**kwargs):
        captured.update(kwargs)
        return SimpleNamespace(
            choices=[SimpleNamespace(
                message=SimpleNamespace(content="完成"), finish_reason="stop",
            )],
        )

    provider._client = SimpleNamespace(
        chat=SimpleNamespace(completions=SimpleNamespace(create=_create)),
    )
    assert provider._openai_chat([{"role": "user", "content": "测试"}]) == ("完成", "stop")
    assert captured["max_completion_tokens"] == 64000
    assert "max_tokens" not in captured
    assert "temperature" not in captured
