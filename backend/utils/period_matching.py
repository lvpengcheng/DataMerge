"""月份角色约束：训练文件或 Sheet 中带 上月/本月 时，必须与工资年月的 YYYYMM 对应。

这类约束是确定性的，不能交给 AI 猜。AI 只在没有明确月份角色时做语义匹配。
"""
import re
from typing import Any, Dict, Optional

_CURRENT_WORDS = ("本月", "当月", "当前月", "本期", "当期")
_PREVIOUS_WORDS = ("上月", "前月", "上期", "前期")
_BASENAME_STRIP = re.compile(r"\.(xlsx|xls|xlsm|xltx|xltm)$", re.IGNORECASE)


def period_role(name: Any) -> Optional[str]:
    """识别名称中的本月/上月角色。"""
    base = _BASENAME_STRIP.sub("", str(name or "")).lower()
    if any(word.lower() in base for word in _CURRENT_WORDS):
        return "current"
    if any(word.lower() in base for word in _PREVIOUS_WORDS):
        return "previous"
    return None


def period_index(name: Any, reference_idx: Optional[int] = None) -> Optional[int]:
    """从名称中提取月份索引。

    支持：
    - 202607 / 2026-07 / 2026年7月
    - 7月 / 7月份（需要 reference_idx 才能确定年份）
    """
    base = _BASENAME_STRIP.sub("", str(name or ""))
    match = re.search(r"(?<!\d)(20\d{2})\D{0,3}(0?[1-9]|1[0-2])(?!\d)", base)
    if match:
        return int(match.group(1)) * 12 + int(match.group(2)) - 1
    if reference_idx is None:
        return None
    month_match = re.search(r"(?<!\d)(0?[1-9]|1[0-2])\s*月", base)
    if not month_match:
        return None
    month = int(month_match.group(1))
    current_year = reference_idx // 12
    current_month = reference_idx - current_year * 12 + 1
    year = current_year if month <= current_month else current_year - 1
    return year * 12 + month - 1


def period_context(matching_context: Optional[Dict[str, Any]]) -> Optional[int]:
    """从 matching_context 取工资年月，返回月份索引；无效或缺失返回 None。"""
    context = matching_context or {}
    try:
        year = int(context.get("salary_year"))
        month = int(context.get("salary_month"))
    except (TypeError, ValueError):
        return None
    if not (1 <= month <= 12):
        return None
    return year * 12 + month - 1


def _naming_candidates(sheet: Dict[str, Any]) -> list:
    return [
        sheet.get("file_name"), sheet.get("original_file_name"),
        sheet.get("sheet_name"), sheet.get("original_sheet_name"),
    ]


def period_candidate_allowed(
    train_sheet: Dict[str, Any],
    actual_sheet: Dict[str, Any],
    matching_context: Optional[Dict[str, Any]],
) -> bool:
    """检查训练 上月/本月 角色与上传名称中的 YYYYMM 是否冲突。"""
    current_idx = period_context(matching_context)
    if current_idx is None:
        return True
    train_names = [str(value) for value in _naming_candidates(train_sheet) if value]
    actual_names = [str(value) for value in _naming_candidates(actual_sheet) if value]
    actual_indices = {period_index(name, current_idx) for name in actual_names}
    actual_indices.discard(None)
    if not actual_indices:
        return True
    for name in train_names:
        role = period_role(name)
        if role == "current":
            if current_idx not in actual_indices:
                return False
        elif role == "previous":
            if (current_idx - 1) not in actual_indices:
                return False
    return True
