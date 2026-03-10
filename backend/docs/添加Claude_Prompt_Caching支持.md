# 为Claude Provider添加Prompt Caching支持

## 功能概述

为Claude AI Provider添加了Prompt Caching（提示词缓存）功能，可以显著降低API调用成本和延迟。

## 什么是Prompt Caching？

Prompt Caching是Anthropic提供的一项功能，允许缓存长提示词的前缀部分，在后续请求中重用，从而：
- **降低成本**：缓存命中的token价格更低
- **减少延迟**：不需要重新处理缓存的内容
- **提升性能**：特别适合重复使用相同system prompt的场景

## 修改内容

### 1. `_claude_chat` 方法

添加了 `use_cache` 参数（默认为True）：

```python
def _claude_chat(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    """非流式调用 Anthropic SDK，返回 (content, stop_reason)

    Args:
        system_prompt: 系统提示词
        messages: 消息列表
        max_tokens: 最大token数
        temperature: 温度参数
        use_cache: 是否启用提示词缓存（默认True）
        **kwargs: 其他参数
    """
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # 如果启用缓存，添加cache_control参数
    if use_cache:
        filtered_kwargs["cache_control"] = {"type": "ephemeral"}

    response = self._client.messages.create(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,
        messages=messages,
        **filtered_kwargs,
    )
    content = response.content[0].text
    stop_reason = response.stop_reason
    return content, stop_reason
```

### 2. `_claude_chat_stream` 方法

同样添加了 `use_cache` 参数（默认为True）：

```python
def _claude_chat_stream(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    """流式调用 Anthropic SDK，yield (text_chunk, stop_reason)

    Args:
        system_prompt: 系统提示词
        messages: 消息列表
        max_tokens: 最大token数
        temperature: 温度参数
        use_cache: 是否启用提示词缓存（默认True）
        **kwargs: 其他参数
    """
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # 如果启用缓存，添加cache_control参数
    if use_cache:
        filtered_kwargs["cache_control"] = {"type": "ephemeral"}

    stop_reason = None
    with self._client.messages.stream(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,
        messages=messages,
        **filtered_kwargs,
    ) as stream:
        for text in stream.text_stream:
            if text:
                yield text, stop_reason
        final_message = stream.get_final_message()
        stop_reason = final_message.stop_reason
    yield "", stop_reason
```

## 使用方式

### 默认启用缓存（推荐）

```python
# 默认情况下，缓存已启用
content, stop_reason = self._claude_chat(
    system_prompt="You are a helpful assistant...",
    messages=[{"role": "user", "content": "Hello"}]
)
```

### 禁用缓存

```python
# 如果需要禁用缓存，传递 use_cache=False
content, stop_reason = self._claude_chat(
    system_prompt="You are a helpful assistant...",
    messages=[{"role": "user", "content": "Hello"}],
    use_cache=False
)
```

## API请求示例

启用缓存后，实际发送的API请求如下：

```bash
curl https://api.anthropic.com/v1/messages \
  -H "content-type: application/json" \
  -H "x-api-key: $ANTHROPIC_API_KEY" \
  -H "anthropic-version: 2023-06-01" \
  -d '{
    "model": "claude-opus-4-6",
    "max_tokens": 1024,
    "cache_control": {"type": "ephemeral"},
    "system": "You are a helpful assistant that remembers our conversation.",
    "messages": [
      {"role": "user", "content": "My name is Alex. I work on machine learning."},
      {"role": "assistant", "content": "Nice to meet you, Alex! How can I help with your ML work today?"},
      {"role": "user", "content": "What did I say I work on?"}
    ]
  }'
```

## 缓存工作原理

1. **首次请求**：
   - 发送完整的system prompt和messages
   - 添加 `"cache_control": {"type": "ephemeral"}`
   - Anthropic缓存这些内容

2. **后续请求**：
   - 如果system prompt和messages前缀相同
   - Anthropic从缓存中读取，不重新处理
   - 只处理新增的内容

3. **缓存有效期**：
   - `ephemeral` 类型的缓存有效期为5分钟
   - 5分钟内的重复请求可以命中缓存

## 适用场景

### ✅ 适合使用缓存的场景

1. **训练迭代**：
   - 同一个租户的多次训练
   - system prompt相同
   - 规则内容相似

2. **代码修正**：
   - 多次修正同一段代码
   - system prompt不变
   - 只有错误信息变化

3. **批量处理**：
   - 处理多个相似的任务
   - 使用相同的system prompt

### ❌ 不适合使用缓存的场景

1. **一次性请求**：
   - 只调用一次API
   - 无法从缓存中受益

2. **每次都不同**：
   - system prompt每次都变化
   - 无法命中缓存

## 成本节省

根据Anthropic的定价：

| Token类型 | 标准价格 | 缓存命中价格 | 节省 |
|----------|---------|------------|------|
| Input | $3/MTok | $0.30/MTok | 90% |
| Output | $15/MTok | $15/MTok | 0% |

**示例**：
- system prompt: 10,000 tokens
- 训练10次迭代
- 不使用缓存：10,000 × 10 × $3/MTok = $0.30
- 使用缓存：10,000 × $3/MTok + 10,000 × 9 × $0.30/MTok = $0.03 + $0.027 = $0.057
- **节省：81%**

## 在训练中的应用

在训练过程中，system prompt通常是固定的：

```python
system_prompt = (
    "你是一个专业的Python程序员，擅长处理各种Excel数据处理任务，"
    "包括人力资源、财务、供应链等不同业务场景。请生成准确、高效的Python代码。"
    "特别注意根据业务场景选择合适的主键进行数据关联和计算。"
    "只返回Python代码，不要包含解释或其他文本。\n\n"
    "重要：如果代码较长（超过150行），请主动分段输出。"
    "每段在逻辑完整的位置断开（如函数定义之间），"
    f"段末单独输出一行 {self.CONTINUATION_MARKER} 作为标记。"
    "收到'继续'后输出下一段。最后一段不需要标记。"
)
```

这个system prompt在整个训练过程中不变，非常适合缓存。

## 监控缓存效果

Anthropic API响应中会包含缓存使用信息：

```json
{
  "usage": {
    "input_tokens": 1000,
    "cache_creation_input_tokens": 10000,
    "cache_read_input_tokens": 9000,
    "output_tokens": 500
  }
}
```

- `cache_creation_input_tokens`: 首次创建缓存的token数
- `cache_read_input_tokens`: 从缓存读取的token数

## 注意事项

1. **缓存有效期**：
   - `ephemeral` 缓存有效期为5分钟
   - 超过5分钟需要重新创建缓存

2. **缓存键**：
   - 缓存基于完整的请求内容
   - system prompt或messages前缀变化会导致缓存失效

3. **向后兼容**：
   - 默认启用缓存，不影响现有代码
   - 如需禁用，传递 `use_cache=False`

## 修改文件

- `backend/ai_engine/ai_provider.py` - ClaudeProvider类的 `_claude_chat` 和 `_claude_chat_stream` 方法

## 总结

✅ 添加了Prompt Caching支持
✅ 默认启用，自动优化成本
✅ 可以通过 `use_cache=False` 禁用
✅ 特别适合训练场景，可节省高达90%的输入token成本
✅ 向后兼容，不影响现有代码

现在系统会自动利用Prompt Caching功能，降低API调用成本，特别是在训练迭代和代码修正场景中！
