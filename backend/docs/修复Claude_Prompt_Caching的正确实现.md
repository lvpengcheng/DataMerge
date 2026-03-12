# 修复Claude Prompt Caching的正确实现

## 问题

用户检查缓存设置，发现当前实现方式不正确。

### 错误的实现（修复前）

```python
# ❌ 错误：cache_control作为顶层参数传递
filtered_kwargs["cache_control"] = {"type": "ephemeral"}

response = self._client.messages.create(
    model=self.model,
    max_tokens=max_tokens,
    temperature=temperature,
    system=system_prompt,  # 字符串格式
    messages=messages,
    **filtered_kwargs,  # cache_control在这里
)
```

**问题**：
- `cache_control` 不是顶层参数
- `system` 应该是数组格式，而不是字符串
- `cache_control` 应该添加到数组中的文本块上

### 正确的实现（官方示例）

```python
# ✅ 正确：cache_control添加到system数组的文本块上
response = client.messages.create(
    model="claude-opus-4-6",
    max_tokens=1024,
    system=[
        {
            "type": "text",
            "text": "You are an AI assistant..."
        },
        {
            "type": "text",
            "text": "<large content to cache>",
            "cache_control": {"type": "ephemeral"}  # 在这里添加
        }
    ],
    messages=[...]
)
```

## 解决方案

### 修复 `_claude_chat` 方法

**修改前**：
```python
def _claude_chat(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # ❌ 错误的实现
    if use_cache:
        filtered_kwargs["cache_control"] = {"type": "ephemeral"}

    response = self._client.messages.create(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,  # 字符串格式
        messages=messages,
        **filtered_kwargs,
    )
    content = response.content[0].text
    stop_reason = response.stop_reason
    return content, stop_reason
```

**修改后**：
```python
def _claude_chat(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # ✅ 正确的实现：将system_prompt转换为数组格式并添加cache_control
    if use_cache and isinstance(system_prompt, str):
        system_prompt = [
            {
                "type": "text",
                "text": system_prompt,
                "cache_control": {"type": "ephemeral"}
            }
        ]

    response = self._client.messages.create(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,  # 数组格式
        messages=messages,
        **filtered_kwargs,
    )
    content = response.content[0].text
    stop_reason = response.stop_reason
    return content, stop_reason
```

### 修复 `_claude_chat_stream` 方法

**修改前**：
```python
def _claude_chat_stream(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # ❌ 错误的实现
    if use_cache:
        filtered_kwargs["cache_control"] = {"type": "ephemeral"}

    stop_reason = None
    with self._client.messages.stream(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,  # 字符串格式
        messages=messages,
        **filtered_kwargs,
    ) as stream:
        # ...
```

**修改后**：
```python
def _claude_chat_stream(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    filtered_kwargs = {k: v for k, v in kwargs.items() if k not in ("extra_headers", "stream", "use_cache")}

    # ✅ 正确的实现：将system_prompt转换为数组格式并添加cache_control
    if use_cache and isinstance(system_prompt, str):
        system_prompt = [
            {
                "type": "text",
                "text": system_prompt,
                "cache_control": {"type": "ephemeral"}
            }
        ]

    stop_reason = None
    with self._client.messages.stream(
        model=self.model,
        max_tokens=max_tokens or max(self.max_tokens, 64000),
        temperature=temperature,
        system=system_prompt,  # 数组格式
        messages=messages,
        **filtered_kwargs,
    ) as stream:
        # ...
```

## 关键改进

### 1. System参数格式

**修改前**：
```python
system="You are an AI assistant..."  # 字符串
```

**修改后**：
```python
system=[
    {
        "type": "text",
        "text": "You are an AI assistant...",
        "cache_control": {"type": "ephemeral"}
    }
]  # 数组格式
```

### 2. Cache Control位置

**修改前**：
```python
# ❌ 作为顶层参数
filtered_kwargs["cache_control"] = {"type": "ephemeral"}
response = client.messages.create(..., **filtered_kwargs)
```

**修改后**：
```python
# ✅ 添加到system数组的文本块上
system=[
    {
        "type": "text",
        "text": "...",
        "cache_control": {"type": "ephemeral"}  # 在这里
    }
]
```

### 3. 兼容性处理

```python
# 检查system_prompt是否是字符串
if use_cache and isinstance(system_prompt, str):
    # 转换为数组格式
    system_prompt = [
        {
            "type": "text",
            "text": system_prompt,
            "cache_control": {"type": "ephemeral"}
        }
    ]
```

**优点**：
- ✅ 如果system_prompt已经是数组格式，不会重复转换
- ✅ 如果use_cache=False，保持原样
- ✅ 向后兼容

## 缓存工作原理

### 缓存的内容

当启用缓存时，Claude会缓存system prompt的内容：

```python
system=[
    {
        "type": "text",
        "text": "You are an AI assistant...",  # 这部分会被缓存
        "cache_control": {"type": "ephemeral"}
    }
]
```

### 缓存的好处

1. **首次请求**：
   - 正常计费
   - 创建缓存

2. **后续请求**（5分钟内）：
   - 缓存命中：90%折扣
   - 只对新内容正常计费

### 示例

假设system prompt有10,000 tokens：

**不使用缓存**：
```
请求1: 10,000 tokens × $15/M = $0.15
请求2: 10,000 tokens × $15/M = $0.15
请求3: 10,000 tokens × $15/M = $0.15
总计: $0.45
```

**使用缓存**：
```
请求1: 10,000 tokens × $15/M = $0.15 (创建缓存)
请求2: 10,000 tokens × $1.5/M = $0.015 (缓存命中，90%折扣)
请求3: 10,000 tokens × $1.5/M = $0.015 (缓存命中，90%折扣)
总计: $0.18 (节省60%)
```

## 验证方法

### 1. 检查API响应

启用缓存后，API响应会包含缓存信息：

```python
response = client.messages.create(...)

# 检查usage信息
print(response.usage)
# {
#   "input_tokens": 1000,
#   "cache_creation_input_tokens": 10000,  # 首次请求
#   "cache_read_input_tokens": 0,
#   "output_tokens": 500
# }

# 后续请求
# {
#   "input_tokens": 1000,
#   "cache_creation_input_tokens": 0,
#   "cache_read_input_tokens": 10000,  # 缓存命中！
#   "output_tokens": 500
# }
```

### 2. 添加日志

```python
def _claude_chat(self, system_prompt, messages, max_tokens=None, temperature=0.1, use_cache=True, **kwargs):
    # ...
    response = self._client.messages.create(...)

    # 记录缓存使用情况
    usage = response.usage
    logger.info(f"Token使用: input={usage.input_tokens}, "
                f"cache_creation={getattr(usage, 'cache_creation_input_tokens', 0)}, "
                f"cache_read={getattr(usage, 'cache_read_input_tokens', 0)}, "
                f"output={usage.output_tokens}")

    return content, stop_reason
```

### 3. 测试脚本

```python
import anthropic

client = anthropic.Anthropic(api_key="your-api-key")

# 第一次请求
response1 = client.messages.create(
    model="claude-opus-4-6",
    max_tokens=1024,
    system=[
        {
            "type": "text",
            "text": "You are an AI assistant..." * 1000,  # 大量文本
            "cache_control": {"type": "ephemeral"}
        }
    ],
    messages=[{"role": "user", "content": "Hello"}]
)
print("请求1:", response1.usage)

# 第二次请求（5分钟内）
response2 = client.messages.create(
    model="claude-opus-4-6",
    max_tokens=1024,
    system=[
        {
            "type": "text",
            "text": "You are an AI assistant..." * 1000,  # 相同的文本
            "cache_control": {"type": "ephemeral"}
        }
    ],
    messages=[{"role": "user", "content": "Hi"}]
)
print("请求2:", response2.usage)  # 应该显示cache_read_input_tokens
```

## 注意事项

### 1. 缓存有效期

- 缓存有效期：5分钟
- 5分钟后需要重新创建缓存

### 2. 缓存粒度

- 缓存是基于system prompt的完整内容
- 内容改变会导致缓存失效

### 3. 最小缓存大小

- 建议缓存的内容至少1024 tokens
- 太小的内容缓存收益不明显

### 4. 多个缓存块

可以缓存多个块：

```python
system=[
    {
        "type": "text",
        "text": "Base instructions..."
    },
    {
        "type": "text",
        "text": "Large context 1...",
        "cache_control": {"type": "ephemeral"}
    },
    {
        "type": "text",
        "text": "Large context 2...",
        "cache_control": {"type": "ephemeral"}
    }
]
```

## 修改文件

- `backend/ai_engine/ai_provider.py`
  - `_claude_chat` 方法（第1347-1374行）
  - `_claude_chat_stream` 方法（第1376-1408行）

## 总结

✅ 修复了Prompt Caching的实现方式
✅ 将system参数从字符串改为数组格式
✅ 将cache_control添加到文本块上，而不是顶层参数
✅ 添加了兼容性检查
✅ 保持向后兼容

现在缓存功能可以正常工作，能够节省高达90%的输入token成本！🎉
