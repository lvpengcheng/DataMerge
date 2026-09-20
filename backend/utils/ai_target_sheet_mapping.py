"""模板目标表 AI 语义匹配。

只在规则/列签名匹配出现歧义时调用；AI 只做候选推荐，最终仍由人工确认。
"""
import json
import logging

logger = logging.getLogger(__name__)


def suggest_target_sheet_mapping(ambiguous, key_signatures, sheet_vocab, ai_provider_name):
    """为歧义的目标表 key 推荐模板 sheet。

    Args:
        ambiguous: {key: [{"sheet": ..., "score": ...}, ...]}
        key_signatures: {key: [训练列名, ...]}
        sheet_vocab: {sheet_name: [模板顶部表头文本, ...]}
        ai_provider_name: provider 名称；为空则跳过。

    Returns:
        {key: {"ai_recommended": sheet, "ai_confidence": float, "ai_reason": str}}
    """
    if not ambiguous or not ai_provider_name:
        return {}
    payload = []
    for key, candidates in ambiguous.items():
        payload.append({
            "key": str(key),
            "training_columns": [str(c) for c in (key_signatures.get(key) or key_signatures.get(str(key)) or [])][:80],
            "candidates": [
                {"sheet": str(c.get("sheet")),
                 "rule_score": c.get("score"),
                 "sheet_headers": [str(h) for h in (sheet_vocab.get(str(c.get("sheet"))) or [])][:80]}
                for c in (candidates or [])[:8]
            ],
        })
    if not payload:
        return {}
    prompt = (
        "你是 Excel 模板目标表匹配专家。训练时固化了多个目标表 key，"
        "每个 key 带有训练列签名；当月模板 sheet 名称可能随月份/机构/地区变化。\n"
        "请根据训练列签名、当月模板 sheet 顶部表头和候选规则分数，"
        "为每个 key 推荐一个最可能对应的模板 sheet。\n"
        "注意：名称相似但业务含义不同的表不要推荐；证据不足时不要输出该项。\n"
        "严格只输出 JSON 对象，格式："
        '{"mappings":[{"key":"训练目标表 key","sheet":"推荐模板 sheet",'
        '"confidence":0.0-1.0,"reason":"简短中文依据"}]}。\n'
        + json.dumps({"ambiguous": payload}, ensure_ascii=False)
    )
    try:
        from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout
        provider = AIProviderFactory.create_provider(ai_provider_name)
        raw = chat_with_timeout(provider, [{"role": "user", "content": prompt}], max_tokens=4000)
        if not raw:
            return {}
        text = raw.strip()
        if text.startswith("```"):
            text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
        data = json.loads(text)
        mappings = data.get("mappings") if isinstance(data, dict) else None
        if not isinstance(mappings, list):
            return {}
        allowed = {str(k): {str(c.get("sheet")) for c in cands} for k, cands in ambiguous.items()}
        out = {}
        for item in mappings:
            if not isinstance(item, dict):
                continue
            key = str(item.get("key") or "")
            sheet = str(item.get("sheet") or "")
            if key not in allowed or sheet not in allowed[key]:
                continue
            try:
                confidence = float(item.get("confidence"))
            except (TypeError, ValueError):
                confidence = None
            if confidence is not None and not 0.0 <= confidence <= 1.0:
                confidence = None
            out[key] = {
                "ai_recommended": sheet,
                "ai_confidence": confidence,
                "ai_reason": str(item.get("reason") or "")[:500],
            }
        return out
    except Exception as exc:
        logger.warning("[TargetSheet/AI] 推荐失败: %s", exc)
        return {}
