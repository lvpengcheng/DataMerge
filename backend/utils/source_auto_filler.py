"""源文件自动补全模块

计算时，如果用户未上传某些训练时使用的源文件，
自动从基础资料（DataAsset, asset_type="reference"）中查找匹配文件并补全。

优先级：用户上传 > 租户基础资料 > 全局基础资料
匹配策略：文件名精确匹配 > 表头结构模糊匹配

另：用户改名上传场景由 auto_rename_uploaded_by_combined_score 处理：
列头 Jaccard + 文件名相似度组合打分；高置信度自动改名，模糊场景返回候选交前端确认。
"""

import os
import re
import shutil
import logging
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from sqlalchemy.orm import Session
else:
    Session = Any

logger = logging.getLogger(__name__)

# 表头匹配阈值：交集占比 >= 此值视为匹配
HEADER_MATCH_THRESHOLD = 0.7

# 组合评分自动改名阈值
RENAME_AUTO_SCORE = 0.85       # 综合分门槛（列头+文件名加权）
RENAME_AUTO_LEAD = 0.15        # 综合分领先第二名的差距
RENAME_JACCARD_AUTO = 0.85     # 列头主导判定：jaccard 单独门槛（不依赖文件名）
RENAME_JACCARD_LEAD = 0.20     # 列头主导判定：jaccard 领先第二名的差距
RENAME_CANDIDATE_FLOOR = 0.10  # 候选展示最低分（低于此分不按常规评分显示；
                               # 若某上传文件对所有期望都低于此分，仍兜底列出全部期望供手动选择）
RENAME_HEADER_WEIGHT = 0.7
RENAME_NAME_WEIGHT = 0.3

_CURRENT_PERIOD_WORDS = ("本月", "当月", "当前月", "本期", "当期")
_PREVIOUS_PERIOD_WORDS = ("上月", "前月", "上期", "前期")


def auto_fill_missing_sources(
    source_dir: str,
    source_structure: dict,
    tenant_id: str,
    db_session: Session,
) -> Tuple[List[dict], List[str]]:
    """自动补全缺失的源文件

    Args:
        source_dir: 用户上传文件所在目录
        source_structure: 训练时记录的源文件结构（Script.source_structure）
        tenant_id: 当前租户 ID
        db_session: 数据库会话

    Returns:
        (filled_list, missing_list)
        - filled_list: [{"file_name": "xx.xlsx", "source": "tenant/global", "asset_name": "xx"}]
        - missing_list: ["缺失文件名1.xlsx", ...]
    """
    if not source_structure or "files" not in source_structure:
        return [], []

    expected_files = set(source_structure["files"].keys())
    uploaded_files = {
        f.name for f in Path(source_dir).iterdir()
        if f.is_file() and not f.name.startswith("~")
    }

    missing = expected_files - uploaded_files
    if not missing:
        return [], []

    logger.info(f"[AutoFill] 训练期望文件: {expected_files}, 已上传: {uploaded_files}, 缺失: {missing}")

    # 加载候选基础资料
    tenant_assets, global_assets = _load_reference_assets(tenant_id, db_session)

    filled_list = []
    missing_list = []

    for file_name in missing:
        asset, source_scope = _find_matching_asset(
            file_name,
            source_structure["files"].get(file_name, {}),
            tenant_assets,
            global_assets,
        )

        if asset:
            # 复制文件到 source_dir
            src_path = asset.file_path
            dst_path = os.path.join(source_dir, file_name)

            if os.path.exists(src_path):
                shutil.copy2(src_path, dst_path)
                filled_list.append({
                    "file_name": file_name,
                    "source": source_scope,
                    "asset_name": asset.name,
                    "asset_id": asset.id,
                })
                logger.info(f"[AutoFill] 补全 '{file_name}' ← {source_scope}基础资料 '{asset.name}'")
            else:
                logger.warning(f"[AutoFill] 基础资料文件不存在: {src_path}")
                missing_list.append(file_name)
        else:
            missing_list.append(file_name)
            logger.info(f"[AutoFill] '{file_name}' 未找到匹配的基础资料")

    return filled_list, missing_list


def _load_reference_assets(
    tenant_id: str,
    db_session: Session,
) -> Tuple[list, list]:
    """加载租户级和全局级的基础资料

    Returns:
        (tenant_assets, global_assets)
    """
    from backend.database.models import DataAsset

    all_assets = (
        db_session.query(DataAsset)
        .filter(
            DataAsset.asset_type == "reference",
            DataAsset.is_active == True,
            (DataAsset.tenant_id == tenant_id) | (DataAsset.tenant_id.is_(None)),
        )
        .order_by(DataAsset.created_at.desc())
        .all()
    )

    tenant_assets = [a for a in all_assets if a.tenant_id == tenant_id]
    global_assets = [a for a in all_assets if a.tenant_id is None]

    return tenant_assets, global_assets


def _find_matching_asset(
    expected_name: str,
    file_structure: dict,
    tenant_assets: list,
    global_assets: list,
) -> Tuple[Optional[object], Optional[str]]:
    """按优先级查找匹配的基础资料

    匹配顺序：
    1. 租户级 - 文件名匹配
    2. 全局级 - 文件名匹配
    3. 租户级 - 表头结构匹配
    4. 全局级 - 表头结构匹配

    Returns:
        (asset, scope) 或 (None, None)
    """
    # 1. 租户级文件名匹配
    match = _match_by_filename(expected_name, tenant_assets)
    if match:
        return match, "租户"

    # 2. 全局级文件名匹配
    match = _match_by_filename(expected_name, global_assets)
    if match:
        return match, "全局"

    # 3. 表头结构匹配（需要 source_structure 中的 headers）
    expected_headers = _extract_headers(file_structure)
    if not expected_headers:
        return None, None

    # 4. 租户级表头匹配
    match = _match_by_headers(expected_headers, tenant_assets)
    if match:
        return match, "租户"

    # 5. 全局级表头匹配
    match = _match_by_headers(expected_headers, global_assets)
    if match:
        return match, "全局"

    return None, None


def _match_by_filename(expected_name: str, assets: list) -> Optional[object]:
    """按文件名匹配基础资料（精确匹配原始文件名）"""
    expected_lower = expected_name.lower()
    for asset in assets:
        if asset.file_name and asset.file_name.lower() == expected_lower:
            return asset
    return None


def _extract_headers(file_structure: dict) -> set:
    """从 source_structure 中提取某个文件的所有表头名称"""
    headers = set()
    sheets = file_structure.get("sheets", {})
    for sheet_info in sheets.values():
        # headers 格式：{"列名": "列字母"} — key 是列名
        sheet_headers = sheet_info.get("headers", {})
        if isinstance(sheet_headers, dict):
            headers.update(sheet_headers.keys())
        elif isinstance(sheet_headers, list):
            headers.update(sheet_headers)
    return headers


def _match_by_headers(expected_headers: set, assets: list) -> Optional[object]:
    """按表头结构匹配基础资料

    计算 交集大小 / max(expected, asset_headers) >= 阈值
    返回得分最高的资产
    """
    best_asset = None
    best_score = 0.0

    for asset in assets:
        asset_headers = _get_asset_headers(asset)
        if not asset_headers:
            continue

        intersection = expected_headers & asset_headers
        denominator = max(len(expected_headers), len(asset_headers))
        if denominator == 0:
            continue

        score = len(intersection) / denominator
        if score >= HEADER_MATCH_THRESHOLD and score > best_score:
            best_score = score
            best_asset = asset

    return best_asset


def _get_asset_headers(asset) -> set:
    """从 DataAsset 中提取表头集合"""
    headers = set()

    # 优先用 parsed_headers
    if asset.parsed_headers:
        ph = asset.parsed_headers
        if isinstance(ph, dict):
            # 可能是 {"Sheet1": ["col1", "col2"]} 或 {"col1": "A", ...}
            for v in ph.values():
                if isinstance(v, list):
                    headers.update(v)
                elif isinstance(v, dict):
                    headers.update(v.keys())
                elif isinstance(v, str):
                    headers.add(v)
            # 如果 value 都是字符串（列字母），key 就是表头名
            if not headers and all(isinstance(v, str) for v in ph.values()):
                headers.update(ph.keys())
        elif isinstance(ph, list):
            headers.update(ph)

    # 兜底用 sheet_summary
    if not headers and asset.sheet_summary:
        for sheet in asset.sheet_summary:
            if isinstance(sheet, dict):
                h = sheet.get("headers", [])
                if isinstance(h, list):
                    headers.update(h)
                elif isinstance(h, dict):
                    headers.update(h.keys())

    return headers


# ==================== 改名上传场景：组合评分自动改名 ====================

def auto_rename_uploaded_by_combined_score(
    source_dir: str,
    source_structure: dict,
    salary_year: Optional[int] = None,
    salary_month: Optional[int] = None,
) -> Tuple[List[Dict[str, str]], List[Dict[str, object]], Dict[str, set]]:
    """对"上传文件名与训练期望不一致"的场景做组合评分匹配

    用户场景：上传文件 = `2026年薪资.xlsx`，训练期望 = `薪资明细表.xlsx`。
    用列头 Jaccard + 文件名相似度组合打分。
    - 高置信度：直接物理改名（os.rename），让后续 auto_fill 能正确识别
    - 模糊场景：返回候选清单，由前端弹窗交用户选择（可继续交 AI 分析语义）

    Returns:
        (renamed_list, ambiguous_candidates, uploaded_headers_map)
        - renamed_list: [{"from": "上传名", "to": "期望名", "score": 0.92, ...}]
        - ambiguous_candidates: [
              {"uploaded": "2026年薪资.xlsx",
               "candidates": [
                  {"expected": "薪资明细表.xlsx", "score": 0.62,
                   "header_jaccard": 0.5, "name_similarity": 0.8},
                  ...
               ]}
          ]
        - uploaded_headers_map: {"2026年薪资.xlsx": {"工号", "姓名", ...}}
          仅包含尚未被自动改名的 uploaded extras 的列头集合，便于后续 AI 复用
    """
    if not source_structure or "files" not in source_structure:
        return [], [], {}

    expected_files: Dict[str, dict] = source_structure.get("files") or {}
    expected_names = set(expected_files.keys())

    src_path = Path(source_dir)
    if not src_path.exists():
        return [], [], {}

    uploaded_files = [
        f for f in src_path.iterdir()
        if f.is_file() and not f.name.startswith("~") and f.suffix.lower() in (".xlsx", ".xls", ".xlsm")
    ]
    uploaded_names = {f.name for f in uploaded_files}

    # 已经有正确文件名的 → 不参与
    extras = [f for f in uploaded_files if f.name not in expected_names]
    missing = expected_names - uploaded_names
    if not extras or not missing:
        return [], [], {}

    logger.info(f"[Rename] 上传未命名匹配: {[f.name for f in extras]}, 缺失期望: {missing}")

    # 工资月份角色优先于模糊相似度。典型场景：训练文件叫“上月/本月”，计算上传
    # “202607/202608”。这类文件的列头往往完全相同，仅靠 Jaccard 无法区分。
    period_renamed = _rename_by_period_role(
        src_path, extras, missing, salary_year=salary_year, salary_month=salary_month,
    )
    if period_renamed:
        uploaded_files = [
            f for f in src_path.iterdir()
            if f.is_file() and not f.name.startswith("~")
            and f.suffix.lower() in (".xlsx", ".xls", ".xlsm")
        ]
        uploaded_names = {f.name for f in uploaded_files}
        extras = [f for f in uploaded_files if f.name not in expected_names]
        missing = expected_names - uploaded_names
        if not extras or not missing:
            return period_renamed, [], {}

    # 1:1 无歧义直通：只缺 1 个期望文件、且只多出 1 个上传文件 → 直接映射，忽略文件名差异。
    # 场景：训练时模板/源叫 aa.xlsx，计算时改名成 bb.xlsx —— 只有一个候选，没有歧义，
    # 不应因文件名不同而拒绝计算（列头是否相似无关紧要，用户显然就是拿它替换）。
    if len(extras) == 1 and len(missing) == 1:
        u_name = extras[0].name
        e_name = next(iter(missing))
        try:
            old = src_path / u_name
            new = src_path / e_name
            if new.exists():
                new.unlink()
            old.rename(new)
            logger.info(f"[Rename] 1:1 唯一匹配，忽略文件名差异自动映射: {u_name} → {e_name}")
            unique_record = {"from": u_name, "to": e_name, "score": 1.0,
                             "jaccard": None, "name_sim": None, "reason": "1:1 唯一匹配"}
            return (period_renamed + [unique_record], [], {})
        except Exception as _ex:
            logger.warning(f"[Rename] 1:1 自动映射失败，回退打分匹配: {_ex}")

    # 解析每个上传文件的 Sheet + 列头；一次解析同时供规则评分和 AI 使用。
    uploaded_signatures = {f.name: _read_uploaded_signature(str(f)) for f in extras}
    uploaded_headers = {
        name: signature["headers"] for name, signature in uploaded_signatures.items()
    }

    # 提取每个 missing 期望文件的列头
    expected_headers_map: Dict[str, set] = {}
    expected_sheets_map: Dict[str, set] = {}
    for name in missing:
        file_info = expected_files.get(name) or {}
        expected_headers_map[name] = _extract_headers(file_info)
        expected_sheets_map[name] = set((file_info.get("sheets") or {}).keys())

    # 全配对打分
    # (score, uploaded, expected, header_jaccard, name_sim, sheet_sim)
    scored_pairs: List[Tuple[float, str, str, float, float, float]] = []
    for u_name, u_headers in uploaded_headers.items():
        for e_name, e_headers in expected_headers_map.items():
            jaccard = _jaccard(u_headers, e_headers)
            name_sim = _filename_similarity(u_name, e_name)
            sheet_sim = _set_name_similarity(
                uploaded_signatures[u_name]["sheet_names"], expected_sheets_map[e_name]
            )
            if uploaded_signatures[u_name]["sheet_names"] and expected_sheets_map[e_name]:
                score = 0.6 * jaccard + 0.2 * sheet_sim + 0.2 * name_sim
            else:
                score = RENAME_HEADER_WEIGHT * jaccard + RENAME_NAME_WEIGHT * name_sim
            scored_pairs.append((score, u_name, e_name, jaccard, name_sim, sheet_sim))

    # 贪心匹配：按分数降序，每个 uploaded / expected 各只能用一次
    scored_pairs.sort(key=lambda x: -x[0])
    renamed_list: List[Dict[str, str]] = list(period_renamed)
    consumed_uploaded: set = set()
    consumed_expected: set = set()

    # 第一遍：自动改名（高置信度 + 领先第二名）
    # 为了判断"领先",每个 uploaded 收集 top-2 分数
    top_by_uploaded: Dict[str, List[Tuple[float, str, float, float, float]]] = {}
    for score, u, e, j, n, sn in scored_pairs:
        top_by_uploaded.setdefault(u, []).append((score, e, j, n, sn))

    for u_name, ranked in top_by_uploaded.items():
        if not ranked:
            continue
        ranked = ranked[:5]  # 已按分降序

        # 综合分（列头+文件名加权）的第一/二名
        first_score, first_expected, first_j, first_n, first_sn = ranked[0]
        second_score = ranked[1][0] if len(ranked) >= 2 else 0.0
        score_lead = first_score - second_score

        # 列头主导判定：单独按 jaccard 排，找 top-1 / top-2
        by_j = sorted(ranked, key=lambda r: -r[2])
        j_first_score, j_first_expected, j_first_j, j_first_n, j_first_sn = by_j[0]
        j_second_j = by_j[1][2] if len(by_j) >= 2 else 0.0
        j_lead = j_first_j - j_second_j

        # 决策：列头主导优先（列头是真凭实据），综合分作兜底
        auto_target = None
        auto_meta = None
        decision = None
        if j_first_j >= RENAME_JACCARD_AUTO and j_lead >= RENAME_JACCARD_LEAD:
            auto_target = j_first_expected
            auto_meta = (j_first_score, j_first_j, j_first_n, j_first_sn)
            decision = "header"
        elif first_score >= RENAME_AUTO_SCORE and score_lead >= RENAME_AUTO_LEAD:
            auto_target = first_expected
            auto_meta = (first_score, first_j, first_n, first_sn)
            decision = "combined"

        if (
            auto_target
            and auto_target not in consumed_expected
            and u_name not in consumed_uploaded
        ):
            picked_score, picked_j, picked_n, picked_sn = auto_meta
            try:
                old = src_path / u_name
                new = src_path / auto_target
                if new.exists():
                    new.unlink()
                old.rename(new)
                consumed_uploaded.add(u_name)
                consumed_expected.add(auto_target)
                renamed_list.append({
                    "from": u_name,
                    "to": auto_target,
                    "score": round(picked_score, 3),
                    "header_jaccard": round(picked_j, 3),
                    "name_similarity": round(picked_n, 3),
                    "sheet_similarity": round(picked_sn, 3),
                    "decision": decision,
                })
                logger.info(
                    f"[Rename] 自动改名 '{u_name}' → '{auto_target}' "
                    f"(decision={decision}, score={picked_score:.2f}, jaccard={picked_j:.2f}, "
                    f"name_sim={picked_n:.2f}, score_lead={score_lead:.2f}, j_lead={j_lead:.2f})"
                )
            except Exception as e:
                logger.warning(f"[Rename] 改名失败 '{u_name}' → '{auto_target}': {e}")

    # 第二遍：剩下的算候选。未匹配的上传文件必须始终能手动选择——
    # 1) 候选全量（不再截断 top-5）：上传/期望文件多时，被截断的期望候选用户就选不到；
    # 2) 即使与所有期望文件的分数都低于 RENAME_CANDIDATE_FLOOR（列头/文件名都不像，
    #    正是"名称、列头都不一致"的场景），也要兜底列出全部剩余期望作候选，
    #    否则该文件不会出现在前端选择列表里，用户想手动指定也选不到。
    ambiguous: List[Dict[str, object]] = []
    remaining_headers: Dict[str, set] = {}
    all_remaining_expected = [e for e in expected_names if e not in consumed_expected]

    # 全量配对（按 uploaded 分组，分数降序）——不依赖 top_by_uploaded 的 top-5 截断
    pairs_by_uploaded: Dict[str, List[Tuple[float, str, float, float, float]]] = {}
    for score, u, e, j, n, sn in scored_pairs:
        if u not in consumed_uploaded and e not in consumed_expected:
            pairs_by_uploaded.setdefault(u, []).append((score, e, j, n, sn))

    for u_name, ranked in pairs_by_uploaded.items():
        ranked.sort(key=lambda x: -x[0])
        candidates = []
        for score, e, j, n, sn in ranked:
            if score < RENAME_CANDIDATE_FLOOR:
                continue
            candidates.append({
                "expected": e,
                "score": round(score, 3),
                "header_jaccard": round(j, 3),
                "name_similarity": round(n, 3),
                "sheet_similarity": round(sn, 3),
            })
        if not candidates:
            # 兜底：全部剩余期望列为候选（低分也列出，按分数排序），
            # 保证该上传文件始终出现在手动选择列表里
            for e_name in all_remaining_expected:
                m = next((s for s in scored_pairs if s[1] == u_name and s[2] == e_name), None)
                if m:
                    candidates.append({
                        "expected": e_name,
                        "score": round(m[0], 3),
                        "header_jaccard": round(m[3], 3),
                        "name_similarity": round(m[4], 3),
                        "sheet_similarity": round(m[5], 3),
                    })
            candidates.sort(key=lambda c: -c["score"])
        if candidates:
            ambiguous.append({
                "uploaded": u_name,
                "candidates": candidates,   # 全量候选，前端下拉可滚动选择
                "uploaded_sheets": uploaded_signatures.get(u_name, {}).get("sheets", []),
            })
            remaining_headers[u_name] = uploaded_headers.get(u_name, set())

    # 即使系统未启用 AI provider，也给前端预选一个结构候选；它不会自动执行，
    # 仍需用户在事前校验弹窗中确认。
    return renamed_list, _add_structural_fallback_recommendations(ambiguous), remaining_headers


def ai_disambiguate_rename_candidates(
    ambiguous: List[Dict[str, object]],
    source_structure: dict,
    uploaded_headers_map: Dict[str, set],
    ai_provider_name: str,
) -> List[Dict[str, object]]:
    """对模糊改名候选调 AI 做语义裁决

    给 AI 同时看：
    - 上传文件名 + 列头集合
    - 训练期望文件名 + 列头集合（仅 candidates 里出现过的 expected）
    AI 给每个 uploaded 文件返回最匹配的 expected + 置信度 + 简短理由。

    将结果写回 ambiguous 各项的 `ai_recommended` / `ai_confidence` / `ai_reason` 字段。
    AI 失败时 ambiguous 原样返回（不破坏程序评分结果）。
    """
    if not ambiguous or not ai_provider_name:
        return ambiguous

    expected_files = (source_structure or {}).get("files") or {}

    # 收集出现在候选里的 expected 名（避免把不相关的 expected 全塞给 AI）。
    # 候选现为全量（上传/期望文件多时可能几十个），AI 裁决只看 top-8 高分候选，
    # 不影响前端下拉的全量展示。
    relevant_expected: set = set()
    for item in ambiguous:
        for c in (item.get("candidates") or [])[:8]:
            if c.get("expected"):
                relevant_expected.add(c["expected"])

    expected_payload = []
    for name in relevant_expected:
        file_info = expected_files.get(name) or {}
        headers = sorted(_extract_headers(file_info))
        sheets = []
        for sheet_name, sheet_info in (file_info.get("sheets") or {}).items():
            sheet_headers = sheet_info.get("headers") or {}
            if isinstance(sheet_headers, dict):
                sheet_headers = list(sheet_headers.keys())
            sheets.append({"name": sheet_name, "headers": list(sheet_headers)[:50]})
        expected_payload.append({"file": name, "headers": headers[:50], "sheets": sheets})

    uploaded_payload = []
    for item in ambiguous:
        u_name = item.get("uploaded")
        headers = sorted(uploaded_headers_map.get(u_name, set()))
        candidate_files = [c["expected"] for c in (item.get("candidates") or [])[:8] if c.get("expected")]
        uploaded_payload.append({
            "file": u_name,
            "headers": headers[:50],
            "sheets": item.get("uploaded_sheets") or [],
            "candidate_expected_files": candidate_files,
        })

    if not expected_payload or not uploaded_payload:
        return _add_structural_fallback_recommendations(ambiguous)

    prompt = _build_rename_disambiguation_prompt(uploaded_payload, expected_payload)

    try:
        from backend.ai_engine.ai_provider import AIProviderFactory, chat_with_timeout
        provider = AIProviderFactory.create_provider(ai_provider_name)
        messages = [
            {"role": "system", "content": "你是一个 Excel 数据语义识别专家，擅长在不同命名习惯下识别文件之间的对应关系。"},
            {"role": "user", "content": prompt},
        ]
        # 带总超时：AI 卡住时原样返回程序评分结果（提交请求同步等待，超网关读超时会被判 504）
        raw = chat_with_timeout(provider, messages, max_tokens=1500)
        if raw is None:
            return _add_structural_fallback_recommendations(ambiguous)
    except Exception as e:
        logger.warning(f"[Rename/AI] 调用失败: {e}")
        return _add_structural_fallback_recommendations(ambiguous)

    ai_map = _parse_rename_disambiguation_response(raw)
    if not ai_map:
        return _add_structural_fallback_recommendations(ambiguous)

    # 写回 ambiguous
    for item in ambiguous:
        u_name = item.get("uploaded")
        rec = ai_map.get(u_name)
        if not rec:
            continue
        rec_expected = rec.get("expected")
        # 仅当 AI 推荐项确实在候选列表里才采纳
        cand_names = {c.get("expected") for c in (item.get("candidates") or [])}
        if rec_expected and rec_expected in cand_names:
            item["ai_recommended"] = rec_expected
            item["ai_confidence"] = float(rec.get("confidence", 0.0) or 0.0)
            item["ai_reason"] = str(rec.get("reason", "") or "")
            item["recommendation_source"] = "ai"

    return _add_structural_fallback_recommendations(ambiguous)


def _build_rename_disambiguation_prompt(uploaded_payload, expected_payload) -> str:
    import json as _json
    return (
        "下面是「用户上传的文件」和「训练期望文件」两份清单，每个文件附带其列头集合。\n"
        "由于用户改名了上传文件，需要你判断每个上传文件最可能对应训练期望中的哪一个。\n"
        "判断依据：列头语义重合度优先；文件名语义为辅。\n\n"
        f"## 用户上传文件（共 {len(uploaded_payload)} 个）\n"
        + _json.dumps(uploaded_payload, ensure_ascii=False, indent=2)
        + "\n\n"
        f"## 训练期望文件候选（共 {len(expected_payload)} 个）\n"
        + _json.dumps(expected_payload, ensure_ascii=False, indent=2)
        + "\n\n"
        "请综合文件名、月份/期间语义、Sheet 名和列名判断。像‘202607/202608’与‘上月/本月’这类名称，"
        "应按月份先后关系匹配。\n"
        "对每个上传文件，都必须从其 candidate_expected_files 里返回一个最可能的候选；把握低也不要省略，"
        "请降低 confidence 并说明原因，最终仍由用户确认。\n"
        "**严格只输出 JSON 数组**，每项格式：\n"
        '  {"uploaded": "上传文件名", "expected": "训练期望文件名", "confidence": 0.0-1.0, "reason": "简短中文原因"}\n'
        "不要输出任何 JSON 之外的文字、解释、代码块标记。"
    )


def _parse_rename_disambiguation_response(raw: str) -> Dict[str, Dict[str, object]]:
    if not raw:
        return {}
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    m = re.search(r"\[[\s\S]*\]", text)
    if m:
        text = m.group(0)
    import json as _json
    try:
        data = _json.loads(text)
    except Exception as e:
        logger.warning(f"[Rename/AI] 响应 JSON 解析失败: {e}; raw={raw[:300]}")
        return {}

    out: Dict[str, Dict[str, object]] = {}
    if isinstance(data, list):
        for item in data:
            if not isinstance(item, dict):
                continue
            u = str(item.get("uploaded") or "").strip()
            e = str(item.get("expected") or "").strip()
            if not u or not e:
                continue
            out[u] = {
                "expected": e,
                "confidence": float(item.get("confidence", 0.0) or 0.0),
                "reason": str(item.get("reason", "") or ""),
            }
    return out


def apply_confirmed_renames(source_dir: str, confirmed_renames: Dict[str, str]) -> List[Dict[str, str]]:
    """应用前端用户确认的改名映射

    Args:
        source_dir: 源文件目录
        confirmed_renames: {"上传文件名": "目标期望文件名", ...}

    Returns:
        实际改名记录 [{"from": ..., "to": ...}, ...]
    """
    src_path = Path(source_dir)
    applied: List[Dict[str, str]] = []
    for u_name, target in (confirmed_renames or {}).items():
        if not u_name or not target or u_name == target:
            continue
        old = src_path / u_name
        new = src_path / target
        if not old.exists():
            logger.warning(f"[Rename] 用户确认改名但源不存在: {u_name}")
            continue
        try:
            if new.exists() and new != old:
                new.unlink()
            old.rename(new)
            applied.append({"from": u_name, "to": target})
            logger.info(f"[Rename] 用户确认改名 '{u_name}' → '{target}'")
        except Exception as e:
            logger.error(f"[Rename] 用户确认改名失败 '{u_name}' → '{target}': {e}")
    return applied


# ==================== 评分内部工具 ====================

def _read_uploaded_signature(file_path: str) -> dict:
    """快速解析上传文件的 Sheet 名和列头（每个 sheet 第一个 region 即可）。"""
    headers: set = set()
    sheets = []
    try:
        from excel_parser import IntelligentExcelParser
        parser = IntelligentExcelParser()
        results = parser.parse_excel_file(file_path, max_data_rows=1, read_formulas=False)
        for sheet_data in results:
            sheet_headers = set()
            for region in (sheet_data.regions or []):
                head = region.head_data or {}
                headers.update(head.keys())
                sheet_headers.update(head.keys())
            sheets.append({
                "name": sheet_data.sheet_name,
                "headers": sorted(str(h) for h in sheet_headers if h and str(h).strip())[:50],
            })
    except Exception as e:
        logger.warning(f"[Rename] 解析 {file_path} 列头失败: {e}")
    clean_headers = {h for h in headers if h and str(h).strip()}
    return {
        "headers": clean_headers,
        "sheet_names": {str(s["name"]) for s in sheets if s.get("name")},
        "sheets": sheets,
    }


def _read_uploaded_headers(file_path: str) -> set:
    """兼容旧调用：只返回上传文件列头集合。"""
    return _read_uploaded_signature(file_path)["headers"]


def _jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    inter = a & b
    union = a | b
    return len(inter) / len(union) if union else 0.0


def _set_name_similarity(a: set, b: set) -> float:
    """两个 Sheet 名集合的最佳成对相似度。"""
    if not a or not b:
        return 0.0
    return max(
        SequenceMatcher(None, str(left).lower().strip(), str(right).lower().strip()).ratio()
        for left in a for right in b
    )


_BASENAME_STRIP = re.compile(r"\.(xlsx|xls|xlsm)$", re.IGNORECASE)


def _filename_similarity(a: str, b: str) -> float:
    """文件名（去后缀、转小写）SequenceMatcher 比例"""
    a_base = _BASENAME_STRIP.sub("", a or "").lower().strip()
    b_base = _BASENAME_STRIP.sub("", b or "").lower().strip()
    if not a_base or not b_base:
        return 0.0
    return SequenceMatcher(None, a_base, b_base).ratio()


def _period_index(file_name: str) -> Optional[int]:
    """从 202607 / 2026-07 / 2026年7月 等文件名中提取连续月份索引。"""
    base = _BASENAME_STRIP.sub("", file_name or "")
    match = re.search(r"(?<!\d)(20\d{2})\D{0,3}(0?[1-9]|1[0-2])(?!\d)", base)
    if not match:
        return None
    return int(match.group(1)) * 12 + int(match.group(2)) - 1


def _period_role(file_name: str) -> Optional[str]:
    base = _BASENAME_STRIP.sub("", file_name or "").lower()
    if any(word in base for word in _CURRENT_PERIOD_WORDS):
        return "current"
    if any(word in base for word in _PREVIOUS_PERIOD_WORDS):
        return "previous"
    return None


def _rename_by_period_role(
    src_path: Path,
    extras: List[Path],
    missing: set,
    salary_year: Optional[int] = None,
    salary_month: Optional[int] = None,
) -> List[Dict[str, object]]:
    """把日期文件名确定性映射到上月/本月角色；不能唯一确定时不做自动改名。"""
    expected_by_role: Dict[str, List[str]] = {"current": [], "previous": []}
    for name in missing:
        role = _period_role(name)
        if role:
            expected_by_role[role].append(name)

    dated = [(f, _period_index(f.name)) for f in extras]
    dated = [(f, idx) for f, idx in dated if idx is not None]
    if not dated:
        return []

    current_idx = None
    if salary_year and salary_month and 1 <= int(salary_month) <= 12:
        current_idx = int(salary_year) * 12 + int(salary_month) - 1
    elif len({idx for _, idx in dated}) >= 2:
        current_idx = max(idx for _, idx in dated)
    if current_idx is None:
        return []

    target_idx = {"current": current_idx, "previous": current_idx - 1}
    planned = []
    used_uploads = set()
    for role in ("previous", "current"):
        expected = expected_by_role[role]
        matches = [f for f, idx in dated if idx == target_idx[role] and f.name not in used_uploads]
        if len(expected) == 1 and len(matches) == 1:
            planned.append((matches[0], expected[0], role))
            used_uploads.add(matches[0].name)

    applied: List[Dict[str, object]] = []
    for old, target, role in planned:
        new = src_path / target
        try:
            old.rename(new)
            applied.append({
                "from": old.name, "to": target, "score": 1.0,
                "decision": "period_role", "reason": "按工资年月识别上月/本月",
            })
            logger.info(f"[Rename] 月份角色自动匹配 '{old.name}' → '{target}' ({role})")
        except Exception as exc:
            logger.warning(f"[Rename] 月份角色改名失败 '{old.name}' → '{target}': {exc}")
    return applied


def _add_structural_fallback_recommendations(
    ambiguous: List[Dict[str, object]],
) -> List[Dict[str, object]]:
    """AI 不返回时仍给出低置信结构候选；仅预选，前端仍要求人工确认。"""
    claimed = {
        str(item.get("ai_recommended"))
        for item in ambiguous if item.get("ai_recommended")
    }
    for item in ambiguous:
        if item.get("ai_recommended"):
            continue
        candidates = item.get("candidates") or []
        candidate = next((c for c in candidates if c.get("expected") not in claimed), None)
        if candidate is None and candidates:
            candidate = candidates[0]
        if not candidate or not candidate.get("expected"):
            continue
        expected = str(candidate["expected"])
        claimed.add(expected)
        item["ai_recommended"] = expected
        item["ai_confidence"] = float(candidate.get("score", 0.0) or 0.0)
        item["ai_reason"] = "AI 未返回有效建议；按文件名和列结构给出低置信候选，请人工确认"
        item["recommendation_source"] = "structure_fallback"
    return ambiguous
