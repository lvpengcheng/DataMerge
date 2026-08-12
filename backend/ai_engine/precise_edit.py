# -*- coding: utf-8 -*-
"""模式无关的"外科手术式"精确编辑工具。

三种代码生成模式（模板 / 公式 / 自动）共用：AI 只产出 {find, replace} 替换对，
宿主对每个 find 逐字校验在代码里**唯一出现**才 replace —— 未点名的代码物理上
不可能被改动。片段不唯一 / 无可套用 / 语法不过 → 返回 None，交调用方兜底。

不依赖模板结构、不依赖列标记，因此对任何 Python 脚本（含现有旧脚本）都生效。
"""
import re
import json
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)

_SYSTEM_PROMPT = (
    "你是 Python + Excel（openpyxl / 公式）专家，擅长 HR 薪酬场景。"
    "现在做外科手术式精确修改：只输出结构化替换对，绝不重写整段代码。严格按给定输出格式回答。"
)


def parse_precise_edits(ai_response: str) -> list:
    """从 AI 响应里稳健抽取 {"edits":[{find,replace,note},...]}。

    兼容：纯 JSON、```json 代码块包裹、前后带解释文字。取第一个能解析出
    含 edits 数组的 JSON 对象。失败返回 []。
    """
    if not ai_response:
        return []
    text = ai_response.strip()
    candidates = []
    for m in re.finditer(r"```(?:json)?\s*([\s\S]*?)```", text):
        candidates.append(m.group(1).strip())
    candidates.append(text)
    b = text.find("{")
    e = text.rfind("}")
    if b >= 0 and e > b:
        candidates.append(text[b:e + 1])

    for cand in candidates:
        if not cand or ('"edits"' not in cand and "'edits'" not in cand):
            continue
        try:
            obj = json.loads(cand)
        except Exception:
            continue
        edits = obj.get("edits") if isinstance(obj, dict) else None
        if isinstance(edits, list) and edits:
            out = [it for it in edits
                   if isinstance(it, dict) and "find" in it and "replace" in it]
            if out:
                return out
    return []


def apply_precise_edits(code_segment: str, edits: list):
    """把 edits 逐条套用到 code_segment（find 必须唯一出现才替换）。

    Returns: (patched_segment, applied_notes, failed) —— failed 为 [(note, 出现次数)]。
    """
    patched = code_segment
    applied, failed = [], []
    for e in edits:
        find = e.get("find")
        repl = e.get("replace")
        note = (e.get("note") or "").strip() or (find[:40] if isinstance(find, str) else "")
        if not isinstance(find, str) or not isinstance(repl, str) or not find:
            continue
        cnt = patched.count(find)
        if cnt == 1:
            patched = patched.replace(find, repl, 1)
            applied.append(note)
        else:
            failed.append((note, cnt))
    return patched, applied, failed


def build_prompt(code_segment: str, user_feedback: str, code_label: str = "当前代码",
                 extra_context: str = "") -> str:
    return f"""你正在做**外科手术式**精确修改。下面是{code_label}。请**只**根据【修改指示】改动直接相关的代码，
以 JSON 输出一组精确替换，**严禁重写整段代码**。

## 输出格式（只输出这一个 JSON，前后不要任何解释文字）
```json
{{"edits": [
  {{"find": "<在下面代码里逐字原样、唯一出现的片段（含缩进）>", "replace": "<替换后的片段>", "note": "<这一处改了什么>"}}
]}}
```

## 铁律（违反必败）
1. `find` 必须是下面代码中**逐字原样、且唯一**出现的片段（连同前导空格/缩进一起复制）；片段太短会不唯一，请多带 1~2 行上下文保证唯一
2. **只**为【修改指示】直接涉及的内容生成 edit；其它列、其它逻辑**一个字都不要动**，也**不要**为它们生成 edit
3. 不要改 import、变量名、辅助函数签名，除非指示明确要求
4. f-string 引号规范：外层 `f"..."` 双引号，公式内 sheet 名用 `'` 单引号，空串拼接用 `EMPTY`（=`'""'`）
5. 要新增/删除整行时，`find` 带上相邻锚点行、`replace` 里体现增删；不要凭空插入无锚点代码

## {code_label}
```python
{code_segment}
```

## 修改指示（最高优先级，只改这里提到的内容）
{user_feedback}

{extra_context}"""


def _call_provider(ai_provider, prompt: str, stream_callback: Callable = None,
                   thinking_callback: Callable = None) -> str:
    """统一走 provider 的 chat / chat_stream（所有 provider 都实现 BaseAIProvider 接口）。

    thinking_callback: 推理模型思考过程逐块回调（如 DeepSeek reasoning_content），
    与正式内容(stream_callback)分开，不污染代码流。
    """
    messages = [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    # 输出预算用 provider 配置：deepseek 推理模型的 max_tokens 是【含思考的总预算】，
    # 写死 8000 会被思考阶段吃掉大半、代码截断（Claude 思考不占输出预算所以没暴露）
    _effective_max_tokens = getattr(ai_provider, "max_tokens", None) or 8000
    if stream_callback and hasattr(ai_provider, "chat_stream"):
        from datetime import datetime as _dt

        def _on_chunk(chunk: str):
            if chunk:
                stream_callback(f"[{_dt.now().strftime('%H:%M:%S')}] [CODE] {chunk}")

        return ai_provider.chat_stream(messages, chunk_callback=_on_chunk,
                                       thinking_callback=thinking_callback,
                                       temperature=0.1, max_tokens=_effective_max_tokens) or ""
    return ai_provider.chat(messages, temperature=0.1, max_tokens=_effective_max_tokens) or ""


def _validate_syntax(code: str) -> Optional[str]:
    """语法校验；通过返回 None，否则返回错误描述。"""
    try:
        compile(code, "<precise_edit>", "exec")
        return None
    except SyntaxError as e:
        return f"{e.msg} (line {e.lineno})"


def run_precise_edit(
    ai_provider,
    full_code: str,
    user_feedback: str,
    *,
    code_segment: Optional[str] = None,
    splice_fn: Optional[Callable[[str, str], str]] = None,
    code_label: str = "当前代码",
    extra_context: str = "",
    stream_callback: Callable = None,
    thinking_callback: Callable = None,
    indent_fixer=None,
    training_logger=None,
    reason_sink: Optional[list] = None,
) -> Optional[str]:
    """模式无关精确编辑。

    Args:
        ai_provider: 任一 BaseAIProvider 实例
        full_code: 完整脚本（最终返回/拼接的对象）
        user_feedback: 用户这轮的修改指示（只喂这个，不灌历史）
        code_segment: 展示给 AI 且套用替换的片段；None = 用 full_code 整份
        splice_fn: (full_code, patched_segment) -> 新 full_code；None = 片段即整份
        code_label: prompt 里对代码块的称呼（如 "当前 fill_template 函数"）
        extra_context: 追加到 prompt 末尾的上下文（_COL_MAP / 规则等，可选）
        stream_callback: 代码内容流回调（[CODE] 前缀由 _call_provider 封装）
        thinking_callback: 推理模型思考过程逐块回调（如 DeepSeek reasoning_content）
        reason_sink: 可选 list；失败返回 None 时向其追加一句人类可读的失败原因，
                     供调用方拼进给用户的提示（不影响返回值）。

    Returns:
        修正后的完整脚本；无可套用替换 / 语法不过 → None（调用方兜底）。
    """
    def log(msg):
        logger.info(msg)
        if stream_callback:
            stream_callback(msg)

    def _fail(reason: str):
        if reason_sink is not None:
            reason_sink.append(reason)
        return None

    if not full_code:
        return _fail("没有可修改的原始代码")
    segment = code_segment if code_segment is not None else full_code

    prompt = build_prompt(segment, user_feedback, code_label=code_label, extra_context=extra_context)
    if training_logger and hasattr(training_logger, "log_full_prompt"):
        try:
            training_logger.log_full_prompt(prompt, "precise_edit")
        except Exception:
            pass

    ai_response = _call_provider(ai_provider, prompt, stream_callback, thinking_callback)
    if training_logger and hasattr(training_logger, "log_full_ai_response"):
        try:
            training_logger.log_full_ai_response(ai_response, "precise_edit")
        except Exception:
            pass

    edits = parse_precise_edits(ai_response)
    if not edits:
        log("[精确编辑] 未解析到有效替换对，交由兜底")
        # AI 常在无法完成时用自然语言说明原因，截一段带回给用户
        _hint = (ai_response or "").strip().replace("\n", " ")
        _hint = f"（AI 说明：{_hint[:200]}）" if _hint else ""
        return _fail(f"AI 未给出可套用的精确替换，可能是指示不够具体或该改动无法用局部替换表达{_hint}")

    patched_segment, applied, failed = apply_precise_edits(segment, edits)
    for note, cnt in failed:
        log(f"[精确编辑] ⚠ 片段无法唯一定位（出现 {cnt} 次），已跳过: {note}")
    if not applied:
        log("[精确编辑] 无任何可套用的替换，交由兜底")
        _locs = "；".join(
            f"“{note}”在代码中出现 {cnt} 次（需唯一才能定位）" for note, cnt in failed
        ) if failed else "AI 给出的替换片段与当前代码对不上"
        return _fail(f"要改的位置无法在当前代码里精确定位：{_locs}")
    log(f"[精确编辑] 已精确套用 {len(applied)} 处修改: {applied}；跳过 {len(failed)} 处")

    # 修复管线（可选，provider/indent_fixer 存在才走）
    if indent_fixer is not None:
        try:
            patched_segment = indent_fixer.fix(patched_segment)
        except Exception as e:
            logger.warning(f"缩进修复异常（不阻断）: {e}")
    try:
        if hasattr(ai_provider, "validate_and_fix_code_format"):
            fixed = ai_provider.validate_and_fix_code_format(patched_segment)
            if fixed and isinstance(fixed, str):
                patched_segment = fixed
    except Exception as e:
        logger.warning(f"provider 代码格式修复异常（不阻断）: {e}")

    complete_code = splice_fn(full_code, patched_segment) if splice_fn else patched_segment

    try:
        if hasattr(ai_provider, "_fix_fstring_quotes"):
            _full_fixed = ai_provider._fix_fstring_quotes(complete_code)
            if _full_fixed and isinstance(_full_fixed, str):
                complete_code = _full_fixed
    except Exception as e:
        logger.warning(f"f-string 二次修复异常（不阻断）: {e}")

    syn_err = _validate_syntax(complete_code)
    if syn_err:
        log(f"[精确编辑] 语法校验未通过：{syn_err}，交由兜底")
        return _fail(f"改动后代码语法不通过（{syn_err}），已放弃以免写入坏代码")

    log(f"[精确编辑] 完成（complete_code={len(complete_code)} 字符）")
    return complete_code
