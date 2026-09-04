"""模板/自动代码生成的流式失败降级回归测试。"""

import sys
import types
import unittest

# Codex 的精简测试 Python 未捆绑 requests；本测试不发网络请求，只需让
# ai_provider 模块完成导入。生产容器仍使用真实 requests。
if "requests" not in sys.modules:
    requests_stub = types.ModuleType("requests")
    requests_stub.Session = object
    sys.modules["requests"] = requests_stub

from backend.ai_engine.auto_code_generator import AutoCodeGenerator
from backend.ai_engine.template_code_generator import TemplateCodeGenerator


class _FailingStreamProvider:
    max_tokens = 8000

    def __init__(self):
        self.stream_calls = 0
        self.chat_calls = 0

    def chat_stream(self, _messages, chunk_callback=None, **_kwargs):
        self.stream_calls += 1
        raise RuntimeError(
            "peer closed connection without sending complete message body "
            "(incomplete chunked read)"
        )

    def chat(self, _messages, **_kwargs):
        self.chat_calls += 1
        return "```python\ndef generated():\n    return True\n```"


class _SuccessfulStreamProvider:
    max_tokens = 8000

    def __init__(self):
        self.chat_calls = 0

    def chat_stream(self, _messages, chunk_callback=None, **_kwargs):
        if chunk_callback:
            chunk_callback("def generated():\n")
            chunk_callback("    return True\n")
        return "def generated():\n    return True\n"

    def chat(self, _messages, **_kwargs):
        self.chat_calls += 1
        return "unexpected"


class CodeGenerationStreamFallbackTests(unittest.TestCase):
    def test_stream_disconnect_falls_back_and_replaces_partial_code(self):
        for generator_type in (TemplateCodeGenerator, AutoCodeGenerator):
            with self.subTest(generator=generator_type.__name__):
                provider = _FailingStreamProvider()
                generator = generator_type(ai_provider=provider)
                events = []

                response = generator._call_ai("生成代码", events.append)

                self.assertIn("def generated", response)
                self.assertEqual(provider.stream_calls, 1)
                self.assertEqual(provider.chat_calls, 1)
                self.assertTrue(any("流式调用失败" in item for item in events))
                self.assertTrue(any("[CODE_REPLACE]" in item for item in events))

    def test_successful_stream_does_not_make_duplicate_non_stream_request(self):
        for generator_type in (TemplateCodeGenerator, AutoCodeGenerator):
            with self.subTest(generator=generator_type.__name__):
                provider = _SuccessfulStreamProvider()
                generator = generator_type(ai_provider=provider)
                events = []

                response = generator._call_ai("生成代码", events.append)

                self.assertIn("def generated", response)
                self.assertEqual(provider.chat_calls, 0)
                self.assertTrue(any("[CODE]" in item for item in events))
                self.assertFalse(any("[CODE_REPLACE]" in item for item in events))


if __name__ == "__main__":
    unittest.main()
