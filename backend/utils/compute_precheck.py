"""智算事前校验模块

在 compute_submit 同步阶段拦截，避免事后脚本运行时炸出难定位的二手错误。

校验顺序（全部失败项收集后一次性返回，便于前端一次展示）：
1. 基础资料兜底：用 source_auto_filler 自动从 reference assets 补全缺失文件
2. 表头匹配：复用 FastHeaderMatcher
3. AI 辅助列匹配建议（仅当 step 2 失败时调用，与训练同 provider）
4. 历史数据校验（仅当脚本含 history_provider/load_history 等关键字时检查）

confirmed_mapping 透传：用户在前端弹窗确认 AI 建议后，重提时携带，本模块直接使用、跳过 AI 步骤。
"""

import os
import re
import json
import shutil
import logging
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

# 历史数据关键字（脚本含任一即视为依赖历史数据）
_HISTORY_KEYWORDS = [
    "history_provider",
    "load_history",
    "get_available_months",
    "get_employee_history",
    "HistoricalDataProvider",
    "历史数据",  # sheet 名引用
]


@dataclass
class PrecheckResult:
    ok: bool = True
    missing_files: List[str] = field(default_factory=list)
    auto_filled: List[Dict[str, Any]] = field(default_factory=list)
    auto_renamed: List[Dict[str, Any]] = field(default_factory=list)
    rename_candidates: List[Dict[str, Any]] = field(default_factory=list)
    missing_columns: List[Dict[str, Any]] = field(default_factory=list)
    ai_suggestions: List[Dict[str, Any]] = field(default_factory=list)
    history_warnings: List[str] = field(default_factory=list)
    file_mapping: Optional[Dict[str, Any]] = None

    def to_dict(self) -> dict:
        return asdict(self)


def precheck_compute(
    source_dir: str,
    source_structure: dict,
    manual_headers: Optional[dict],
    script_content: str,
    tenant_id: str,
    salary_year: Optional[int],
    salary_month: Optional[int],
    db_session,
    ai_provider_name: Optional[str] = None,
    confirmed_mapping: Optional[Dict[str, Any]] = None,
    confirmed_renames: Optional[Dict[str, str]] = None,
) -> PrecheckResult:
    """智算事前校验主入口"""
    result = PrecheckResult()

    if not source_structure:
        # 老脚本可能没存 source_structure，无法校验，直接放行
        logger.info("[Precheck] 缺少 source_structure，跳过校验")
        return result

    # 步骤 0：用户已确认的改名映射先落地
    if confirmed_renames:
        try:
            from .source_auto_filler import apply_confirmed_renames
            applied = apply_confirmed_renames(source_dir, confirmed_renames)
            if applied:
                result.auto_renamed.extend(applied)
                logger.info(f"[Precheck] 应用用户确认改名: {applied}")
        except Exception as e:
            logger.warning(f"[Precheck] 应用 confirmed_renames 失败: {e}", exc_info=True)

    # 步骤 0.5：列头+文件名组合评分自动改名（处理用户改名上传场景）
    try:
        from .source_auto_filler import auto_rename_uploaded_by_combined_score, ai_disambiguate_rename_candidates
        renamed, ambiguous, uploaded_headers_map = auto_rename_uploaded_by_combined_score(
            source_dir=source_dir,
            source_structure=source_structure,
        )
        if renamed:
            result.auto_renamed.extend(renamed)
        if ambiguous:
            # 程序无法决断 → 调 AI 给语义裁决，但仍交由前端弹窗确认
            if ai_provider_name:
                try:
                    ambiguous = ai_disambiguate_rename_candidates(
                        ambiguous=ambiguous,
                        source_structure=source_structure,
                        uploaded_headers_map=uploaded_headers_map,
                        ai_provider_name=ai_provider_name,
                    )
                except Exception as ai_err:
                    logger.warning(f"[Precheck] AI 改名裁决失败（不阻断）: {ai_err}", exc_info=True)
            result.ok = False
            result.rename_candidates = ambiguous
            logger.warning(f"[Precheck] 改名候选需用户确认: {[c['uploaded'] for c in ambiguous]}")
            return result
    except Exception as e:
        logger.warning(f"[Precheck] 自动改名评分异常: {e}", exc_info=True)

    # 步骤 1：基础资料兜底
    try:
        from .source_auto_filler import auto_fill_missing_sources
        filled, still_missing = auto_fill_missing_sources(
            source_dir=source_dir,
            source_structure=source_structure,
            tenant_id=tenant_id,
            db_session=db_session,
        )
        result.auto_filled = filled or []
        if still_missing:
            result.ok = False
            result.missing_files = list(still_missing)
            logger.warning(f"[Precheck] 缺失文件无法兜底: {still_missing}")
    except Exception as e:
        logger.warning(f"[Precheck] 基础资料兜底异常: {e}", exc_info=True)

    # 步骤 2：confirmed_mapping 短路（用户已确认 AI 建议）
    if confirmed_mapping:
        try:
            _apply_confirmed_mapping(source_dir, source_structure, confirmed_mapping)
            result.file_mapping = confirmed_mapping.get("file_mapping") or confirmed_mapping
            # 文件已按用户确认改写，跳过表头匹配/AI 步骤
            _check_history(script_content, tenant_id, salary_year, salary_month, result)
            return result
        except Exception as e:
            logger.error(f"[Precheck] 应用 confirmed_mapping 失败: {e}", exc_info=True)
            result.ok = False
            result.missing_columns.append({
                "file": "(confirmed_mapping)",
                "sheet": "",
                "expected_columns": [],
                "error": f"应用确认映射失败: {e}",
            })
            return result

    # 步骤 3：表头匹配
    if not result.missing_files:
        try:
            from .fast_header_matcher import FastHeaderMatcher
            matcher = FastHeaderMatcher()
            input_files = _collect_input_files(source_dir)
            if input_files:
                ok, err, file_mapping, _pre = matcher.match_parse_and_prepare(
                    source_structure=source_structure,
                    input_files=input_files,
                    manual_headers=manual_headers,
                    output_dir=None,  # 校验阶段不写 fallback
                )
                if ok and file_mapping:
                    result.file_mapping = file_mapping
                else:
                    result.ok = False
                    missing = _extract_missing_columns(source_structure, input_files, err)
                    result.missing_columns = missing
                    logger.warning(f"[Precheck] 表头匹配失败: {err}")
                    # 步骤 4：调 AI 给建议
                    if ai_provider_name:
                        try:
                            result.ai_suggestions = _ai_suggest_column_mapping(
                                missing, source_structure, input_files, ai_provider_name
                            )
                        except Exception as ai_err:
                            logger.warning(f"[Precheck] AI 建议失败（不阻断）: {ai_err}", exc_info=True)
        except Exception as e:
            logger.warning(f"[Precheck] 表头匹配异常: {e}", exc_info=True)

    # 步骤 5：历史数据
    _check_history(script_content, tenant_id, salary_year, salary_month, result)

    return result


# ==================== 内部工具 ====================

def _collect_input_files(source_dir: str) -> List[str]:
    p = Path(source_dir)
    files = []
    for ext in ("*.xlsx", "*.xls", "*.xlsm"):
        files.extend(str(f) for f in p.glob(ext) if not f.name.startswith("~"))
    return files


def _extract_missing_columns(
    source_structure: dict,
    input_files: List[str],
    error_msg: Optional[str],
) -> List[Dict[str, Any]]:
    """从训练 source_structure 与上传文件名集合的差集，提取缺失列明细

    简化策略：枚举训练每个文件每个 sheet 的列，标注「未在上传文件里出现的」
    """
    uploaded_names = {os.path.basename(f).lower() for f in input_files}
    missing = []
    files = source_structure.get("files", {}) if isinstance(source_structure, dict) else {}

    for file_name, file_data in files.items():
        if not isinstance(file_data, dict) or "error" in file_data:
            continue
        for sheet_name, sheet_info in (file_data.get("sheets") or {}).items():
            headers = sheet_info.get("headers") if isinstance(sheet_info, dict) else None
            cols = []
            if isinstance(headers, dict):
                cols = list(headers.keys())
            elif isinstance(headers, list):
                cols = list(headers)

            entry = {
                "file": file_name,
                "sheet": sheet_name,
                "expected_columns": cols,
                "uploaded_present": file_name.lower() in uploaded_names,
            }
            missing.append(entry)

    if error_msg:
        missing.append({"file": "(matcher)", "sheet": "", "expected_columns": [], "error": error_msg})
    return missing


def _apply_confirmed_mapping(source_dir: str, source_structure: dict, confirmed_mapping: dict) -> None:
    """按用户确认的列映射改写文件

    confirmed_mapping 格式（与 FastHeaderMatcher.file_mapping 兼容）：
    {
      "<上传文件名>": {
        "expected_file": "<训练文件名>",
        "sheet_mapping": {"<上传 sheet>": "<训练 sheet>", ...},
        "header_mapping": {"<上传列>": "<训练列>", ...}
      }
    }
    """
    from .fast_header_matcher import FastHeaderMatcher

    # 兼容嵌套形式：{"file_mapping": {...}}
    fm = confirmed_mapping.get("file_mapping") if isinstance(confirmed_mapping.get("file_mapping"), dict) else confirmed_mapping

    for input_name, info in fm.items():
        if not isinstance(info, dict):
            continue
        info.setdefault("needs_rewrite", True)
        info.setdefault("input_file", input_name)
        if not info.get("file_path"):
            info["file_path"] = os.path.join(source_dir, input_name)
        try:
            FastHeaderMatcher.rewrite_excel(info, source_dir)
            expected = info.get("expected_file")
            if expected and expected != input_name:
                old = os.path.join(source_dir, input_name)
                new = os.path.join(source_dir, expected)
                if os.path.exists(old) and old != new:
                    if os.path.exists(new):
                        os.remove(new)
                    shutil.move(old, new)
        except Exception as e:
            raise RuntimeError(f"改写文件 {input_name} 失败: {e}") from e


def _ai_suggest_column_mapping(
    missing: List[Dict[str, Any]],
    source_structure: dict,
    input_files: List[str],
    ai_provider_name: str,
) -> List[Dict[str, Any]]:
    """调用 AI 在「训练期望列」与「上传文件实际列」之间给出映射建议"""
    expected_paths = []
    for file_name, file_data in (source_structure.get("files") or {}).items():
        if not isinstance(file_data, dict) or "error" in file_data:
            continue
        for sheet_name, sheet_info in (file_data.get("sheets") or {}).items():
            headers = sheet_info.get("headers") if isinstance(sheet_info, dict) else None
            cols = list(headers.keys()) if isinstance(headers, dict) else (list(headers) if isinstance(headers, list) else [])
            for col in cols:
                expected_paths.append(f"{file_name} > {sheet_name} > {col}")

    actual_paths = _collect_uploaded_columns(input_files)
    if not expected_paths or not actual_paths:
        return []

    prompt = _build_column_match_prompt(expected_paths, actual_paths)

    try:
        from backend.ai_engine.ai_provider import AIProviderFactory
        provider = AIProviderFactory.create_provider(ai_provider_name)
        messages = [
            {"role": "system", "content": "你是一个 Excel 表头匹配专家，擅长在不同命名习惯之间找到语义等价的列对应关系。"},
            {"role": "user", "content": prompt},
        ]
        raw = provider.chat(messages, temperature=0.1, max_tokens=2000)
    except Exception as e:
        logger.warning(f"[Precheck/AI] 调用失败: {e}")
        return []

    return _parse_ai_response(raw)


def _collect_uploaded_columns(input_files: List[str]) -> List[str]:
    """解析上传文件，列出 {file > sheet > col} 路径列表"""
    paths = []
    try:
        from excel_parser import IntelligentExcelParser
        parser = IntelligentExcelParser()
        for fp in input_files:
            try:
                results = parser.parse_excel_file(fp, max_data_rows=1, read_formulas=False)
                fname = os.path.basename(fp)
                for sheet_data in results:
                    sheet_name = sheet_data.sheet_name
                    for region in (sheet_data.regions or []):
                        head = region.head_data or {}
                        for col in head.keys():
                            paths.append(f"{fname} > {sheet_name} > {col}")
            except Exception as e:
                logger.warning(f"[Precheck/AI] 解析上传文件 {fp} 失败: {e}")
    except Exception as e:
        logger.warning(f"[Precheck/AI] 加载 parser 失败: {e}")
    return paths


def _build_column_match_prompt(expected: List[str], actual: List[str]) -> str:
    return (
        "下面是两个列清单，每条格式为 `文件名 > Sheet名 > 列名`（Sheet名可能含 banner 后缀，如 `数据-合同工`）。\n\n"
        f"## 训练期望的列（共 {len(expected)} 项）\n"
        + "\n".join(f"- {p}" for p in expected[:200])
        + "\n\n"
        f"## 用户上传文件实际有的列（共 {len(actual)} 项）\n"
        + "\n".join(f"- {p}" for p in actual[:200])
        + "\n\n"
        "请在期望列与实际列之间做语义匹配，对每个期望列给出最可能的实际列对应（如果完全找不到合理对应可省略该项）。\n"
        "**严格只输出 JSON 数组**，每项格式：\n"
        '  {"expected_path": "...", "suggested_path": "...", "confidence": 0.0-1.0, "reason": "简短中文原因"}\n'
        "不要输出任何 JSON 之外的文字、解释、代码块标记。"
    )


def _parse_ai_response(raw: str) -> List[Dict[str, Any]]:
    if not raw:
        return []
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)

    m = re.search(r"\[[\s\S]*\]", text)
    if m:
        text = m.group(0)
    try:
        data = json.loads(text)
        if isinstance(data, list):
            cleaned = []
            for item in data:
                if not isinstance(item, dict):
                    continue
                cleaned.append({
                    "expected_path": str(item.get("expected_path", "")),
                    "suggested_path": str(item.get("suggested_path", "")),
                    "confidence": float(item.get("confidence", 0.0) or 0.0),
                    "reason": str(item.get("reason", "")),
                })
            return cleaned
    except Exception as e:
        logger.warning(f"[Precheck/AI] 响应 JSON 解析失败: {e}; raw={raw[:300]}")
    return []


def _check_history(
    script_content: Optional[str],
    tenant_id: str,
    salary_year: Optional[int],
    salary_month: Optional[int],
    result: PrecheckResult,
) -> None:
    """脚本含历史数据关键字 + month>1 → 校验前序月份历史数据是否齐全"""
    if not script_content:
        return
    if not any(kw in script_content for kw in _HISTORY_KEYWORDS):
        return
    if not salary_year or not salary_month or salary_month <= 1:
        return

    try:
        from .historical_data import HistoricalDataProvider
        provider = HistoricalDataProvider(tenant_id)
        available = set(provider.get_available_months(salary_year))
        expected = set(range(1, salary_month))
        missing = sorted(expected - available)
        if missing:
            warning = (
                f"脚本依赖历史数据，但 {salary_year} 年缺少 "
                f"{', '.join(str(m) + '月' for m in missing)} 的历史结果"
            )
            result.history_warnings.append(warning)
            logger.warning(f"[Precheck/History] {warning}")
    except Exception as e:
        logger.warning(f"[Precheck/History] 校验异常: {e}", exc_info=True)
