"""源文件自动补全模块

计算时，如果用户未上传某些训练时使用的源文件，
自动从基础资料（DataAsset, asset_type="reference"）中查找匹配文件并补全。

优先级：用户上传 > 租户基础资料 > 全局基础资料
匹配策略：文件名精确匹配 > 表头结构模糊匹配
"""

import os
import shutil
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from sqlalchemy.orm import Session

logger = logging.getLogger(__name__)

# 表头匹配阈值：交集占比 >= 此值视为匹配
HEADER_MATCH_THRESHOLD = 0.7


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
