"""智算源文件摄取（ingest）——整条链路里唯一"打开源文件"的地方。

设计要点（对应"点击计算后只读一次盘"的目标）：

1. **一次解析**：`ingest_source_dir` 内只调用 `FastHeaderMatcher.parse_inputs` 一次，
   每个上传文件仅被 Aspose 打开 1 次；改名评分复用同一份表头，不再单独开文件。
2. **一次确认**：只确认上传文件、源 Sheet 与目标表歧义；源列不进入 AI 或人工匹配。
3. **确认轮不碰 Excel**：解析结果落 `_ingest_sources.pkl`（大），表头层元数据落
   `_ingest_meta.pkl`（小）。人工确认后只读 meta，在内存里重算映射（毫秒级，
   不占 Excel 闸门），确认通过后由计算进程读 sources.pkl 构建预加载数据。
4. **不再重写文件**：以前"确认列映射 → rewrite_excel 整文件重写 → 再解析一次"
   的三步，现在等价地在内存里完成（虚拟改名 + 重新匹配 + 映射复合）。

本模块只做编排，匹配/评分/校验算法全部复用既有实现。
"""

import os
import json
import pickle
import logging
import tempfile
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


# 解析逻辑或产物结构变化时 +1，旧会话产物自动失效
INGEST_VERSION = 10

_META_NAME = "_ingest_meta.pkl"
_SOURCES_NAME = "_ingest_sources.pkl"
_EXCEL_EXTS = (".xlsx", ".xls", ".xlsm")


@dataclass
class IngestMeta:
    """表头层元数据（KB 级，不含 DataFrame），确认轮只读这一份。"""

    source_dir: str = ""
    version: int = INGEST_VERSION
    # 解析参数
    multi_sheet_source: bool = False
    ai_provider_name: Optional[str] = None
    # 训练侧基准
    source_structure: dict = field(default_factory=dict)
    manual_headers: Optional[dict] = None
    expected_structure: Optional[dict] = None
    train_sheets: List[Dict[str, Any]] = field(default_factory=list)
    # 解析结果（表头层）
    input_sheets: List[Dict[str, Any]] = field(default_factory=list)
    signatures: Dict[str, dict] = field(default_factory=dict)
    # 资产登记用的表头摘要 {文件名: [{sheet_name, rows, headers, regions}]}，省掉登记时再开一次文件
    sheet_summary: Dict[str, list] = field(default_factory=dict)
    # 文件级决策
    auto_renamed: List[Dict[str, Any]] = field(default_factory=list)
    rename_candidates: List[Dict[str, Any]] = field(default_factory=list)
    auto_filled: List[Dict[str, Any]] = field(default_factory=list)
    # 基础资料补全后由确定性结构规则直接锁定的映射；不再进入 AI/人工候选。
    auto_filled_mapping: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    missing_files: List[str] = field(default_factory=list)
    parse_error: Optional[str] = None

    def rename_probe(self) -> Dict[str, str]:
        """改名歧义未确认时的试探映射（取最高分候选），用于把后续问题一次性暴露出来。"""
        probe = {}
        for cand in self.rename_candidates or []:
            uploaded = cand.get("uploaded")
            options = cand.get("candidates") or []
            if uploaded and options:
                best = max(options, key=lambda c: float(c.get("score") or 0))
                if best.get("expected"):
                    probe[uploaded] = best["expected"]
        return probe

# ==================== 加密 / 格式预处理 ====================

def detect_encrypted_files(files: List[Tuple[str, str]]) -> List[str]:
    """判断哪些文件是加密的。

    先看 8 字节魔数：现代扩展名 + `PK` 一定不加密，现代扩展名 + OLE 头一定加密；
    只有真正含糊的才落到 Aspose（省掉每个文件一次 DetectFileFormat）。
    """
    encrypted: List[str] = []
    ambiguous: List[Tuple[str, str]] = []
    modern_exts = {".xlsx", ".xlsm", ".xltx", ".xltm"}
    for path, name in files:
        ext = Path(name).suffix.lower()
        try:
            with open(path, "rb") as stream:
                head = stream.read(8)
        except Exception:
            head = b""
        if ext in modern_exts and head.startswith(b"PK"):
            continue
        if ext in modern_exts and head.startswith(b"\xd0\xcf\x11\xe0"):
            encrypted.append(name)
            continue
        ambiguous.append((path, name))

    if ambiguous:
        from backend.utils.aspose_helper import is_encrypted
        encrypted.extend(name for path, name in ambiguous if is_encrypted(path))
    return encrypted


def prepare_source_dir(
    source_dir: str,
    passwords_dict: Optional[Dict[str, str]] = None,
    template_override_path: Optional[str] = None,
) -> Tuple[List[str], Optional[str]]:
    """解密 + xls 转换。返回 (缺密码的加密文件名, 处理后的模板路径)。"""
    import shutil
    from backend.utils.aspose_helper import decrypt_excel
    from backend.utils.source_normalizer import convert_xls_to_xlsx

    passwords_dict = dict(passwords_dict or {})
    src = Path(source_dir)
    names_path = src / '_upload_names.json'
    original_names = json.loads(names_path.read_text(encoding='utf-8')) if names_path.exists() else {}
    candidates = [(str(fp.resolve()), fp.name) for fp in src.iterdir()
                  if fp.is_file() and fp.suffix.lower() in _EXCEL_EXTS]
    if template_override_path:
        template_name = Path(template_override_path).name
        template_key = 'template::' + template_name
        # Older clients used bare names; do not reuse a source password for a same-named template.
        if template_key not in passwords_dict and not any(name == template_name for _, name in candidates):
            passwords_dict[template_key] = passwords_dict.get(template_name)
        candidates.append((template_override_path, template_key))

    encrypted_names = set(detect_encrypted_files(candidates))
    blocked = [name for name in encrypted_names if not passwords_dict.get(name)]
    if blocked:
        return blocked, template_override_path

    for path, name in candidates:
        if name in encrypted_names:
            decrypted = None
            try:
                decrypted = decrypt_excel(path, password=passwords_dict.get(name))
                if detect_encrypted_files([(decrypted, name)]):
                    blocked.append(name)
                    continue
                shutil.move(decrypted, path)
            except Exception:
                # Return to the password dialog without exposing passwords or losing the upload.
                blocked.append(name)
            finally:
                if decrypted and Path(decrypted).resolve() != Path(path).resolve():
                    Path(decrypted).unlink(missing_ok=True)
    if blocked:
        return blocked, template_override_path

    for fp in list(src.iterdir()):
        if fp.is_file() and fp.suffix.lower() == ".xls":
            # A separately uploaded .xlsx with the same stem must not be overwritten.
            if fp.with_suffix('.xlsx').exists():
                continue
            converted = Path(convert_xls_to_xlsx(str(fp.resolve())))
            if converted.name != fp.name:
                original_names[converted.name] = original_names.get(fp.name, fp.name)
    names_path.write_text(json.dumps(original_names, ensure_ascii=False), encoding='utf-8')
    if template_override_path:
        template_override_path = convert_xls_to_xlsx(template_override_path)
    return [], template_override_path


def collect_input_files(source_dir: str) -> List[str]:
    p = Path(source_dir)
    files: List[str] = []
    for ext in _EXCEL_EXTS:
        files.extend(str(f) for f in p.glob(f"*{ext}") if not f.name.startswith("~"))
    return sorted(files)


# ==================== 摄取主流程 ====================

def ingest_source_dir(
    source_dir: str,
    source_structure: Any,
    manual_headers: Optional[dict] = None,
    expected_structure: Optional[dict] = None,
    tenant_id: Optional[str] = None,
    db_session=None,
    ai_provider_name: Optional[str] = None,
    salary_year: Optional[int] = None,
    salary_month: Optional[int] = None,
    confirmed_renames: Optional[Dict[str, str]] = None,
    confirmed_mapping: Optional[Dict[str, Any]] = None,
    defer_base_fill: bool = False,
) -> Tuple[IngestMeta, Dict[tuple, Any]]:
    """解析一次源目录，产出表头层元数据 + 内存解析结果。

    Returns: (meta, parsed_sheets_map)；parsed_sheets_map 形如 {(file_path, sheet): SheetData}
    """
    from .fast_header_matcher import FastHeaderMatcher

    matcher = FastHeaderMatcher()
    structure, struct_err = FastHeaderMatcher.normalize_structure(source_structure or {})
    meta = IngestMeta(source_dir=str(source_dir), ai_provider_name=ai_provider_name)
    if struct_err:
        meta.parse_error = struct_err
        return meta, {}
    meta.source_structure = structure
    meta.manual_headers = manual_headers
    meta.expected_structure = expected_structure

    # 人工文件关系也只作为约束，禁止覆盖上传文件名。
    meta.multi_sheet_source = bool(FastHeaderMatcher._infer_multi_sheet_source(structure))
    meta.train_sheets = matcher._build_training_sheets(structure)

    # 步骤 1：先且只解析本次上传文件。基础资料不能在上传关系尚未判断前抢占角色。
    input_files = collect_input_files(source_dir)
    parsed_sheets_map: Dict[tuple, Any] = {}
    input_sheets: List[Dict[str, Any]] = []
    if input_files:
        try:
            input_sheets, parsed_sheets_map = matcher.parse_inputs(
                input_files, manual_headers, multi_sheet_source=meta.multi_sheet_source)
        except Exception as e:
            meta.parse_error = str(e)
            logger.error(f"[Ingest] 解析源文件失败: {e}", exc_info=True)
            return meta, {}
    meta.signatures = _signatures_from_sheets(input_sheets)
    names_path = Path(source_dir) / '_upload_names.json'
    original_names = json.loads(names_path.read_text(encoding='utf-8')) if names_path.exists() else {}
    for sheet in input_sheets:
        sheet['original_file_name'] = original_names.get(sheet['file_name'], sheet['file_name'])
        sheet['original_sheet_name'] = sheet['sheet_name']
        sheet['source_origin'] = 'upload'

    # 先且只用确定性的结构规则判断本次上传已经覆盖的训练角色。这里禁止调用 AI：
    # 首轮只整理上传关系；人工确认后才允许基础资料补全未覆盖角色。
    matching_context = {'salary_year': salary_year, 'salary_month': salary_month}
    structural = matcher.match_headers_only(
        meta.train_sheets, input_sheets, None, matching_context)
    coverage_mapping = dict((structural.get('mapping') or {}).get('file_mapping') or {})

    structural_files = {info['expected_file'] for info in coverage_mapping.values()
                        if info.get('expected_file')}
    provisional_auto_renamed = []
    provisional_candidates = []
    if not structural['success']:
        from .structural_source_mapping import suggest_file_relations
        provisional_auto_renamed, provisional_candidates = suggest_file_relations(
            meta.train_sheets, input_sheets, coverage_mapping, salary_year, salary_month)
    meta.auto_renamed = list(provisional_auto_renamed)
    meta.rename_candidates = list(provisional_candidates)

    # 步骤 3：基础资料兜底。
    # 严格顺序：本次上传 > 租户基础资料 > 全局基础资料。只有确定性的结构匹配、
    # 月份角色自动关系或人工确认关系才能声明训练角色已覆盖；模糊候选不能阻止补全。
    if (not defer_base_fill) and db_session is not None and tenant_id:
        try:
            from .source_auto_filler import auto_fill_missing_sources
            _confirmed_raw = ((confirmed_mapping or {}).get(
                'file_mapping', confirmed_mapping or {}) or {})
            _confirmed_sheets: Dict[str, set] = {}
            for filename, info in _confirmed_raw.items():
                if not isinstance(info, dict):
                    continue
                target_file = str(info.get('expected_file') or filename)
                _confirmed_sheets.setdefault(target_file, set()).update(
                    str(sheet) for sheet in (info.get('sheet_mapping') or {}).values())
            # 只有训练文件的全部 Sheet 都已由上传关系覆盖，才阻止基础资料补入；
            # 只匹配了其中一个 Sheet 时，仍允许基础资料补齐该文件的其他 Sheet。
            _confirmed_files = set()
            for target_file, covered_sheets in _confirmed_sheets.items():
                expected_sheets = set(((structure.get('files') or {}).get(target_file) or {})
                                      .get('sheets', {}).keys())
                if expected_sheets and expected_sheets <= covered_sheets:
                    _confirmed_files.add(target_file)
            filled, still_missing = auto_fill_missing_sources(
                source_dir=source_dir,
                source_structure=structure,
                tenant_id=tenant_id,
                db_session=db_session,
                assume_present=(structural_files |
                                {r['to'] for r in provisional_auto_renamed} |
                                {v for v in (confirmed_renames or {}).values() if v} |
                                _confirmed_files),
            )
            meta.auto_filled = list(filled or [])
            meta.missing_files = list(still_missing or [])
            # 步骤 3.1：只解析新补进来的文件（增量，不重解析已有文件）
            new_paths = [str(Path(source_dir) /
                             (f.get("stored_file_name") or f["file_name"]))
                         for f in (filled or [])]
            new_paths = [p for p in new_paths if os.path.exists(p)]
            if new_paths:
                extra_sheets, extra_map = matcher.parse_inputs(
                    new_paths, manual_headers, multi_sheet_source=meta.multi_sheet_source)
                filled_by_name = {
                    (item.get('stored_file_name') or item.get('file_name')): item
                    for item in (filled or [])
                }
                for sheet in extra_sheets:
                    fill_info = filled_by_name.get(sheet.get('file_name')) or {}
                    source_scope = fill_info.get('source')
                    sheet['original_file_name'] = sheet.get('file_name')
                    sheet['original_sheet_name'] = sheet.get('sheet_name')
                    sheet['source_origin'] = (
                        'tenant_base' if source_scope == '租户'
                        else 'global_base' if source_scope == '全局'
                        else 'base'
                    )
                    sheet['source_asset_name'] = fill_info.get('asset_name')
                input_sheets.extend(extra_sheets)
                parsed_sheets_map.update(extra_map)
                meta.signatures = _signatures_from_sheets(input_sheets)
        except Exception as e:
            logger.warning(f"[Ingest] 基础资料兜底异常: {e}", exc_info=True)
    elif defer_base_fill:
        meta.missing_files = []
        logger.info('[Ingest] 首轮仅匹配本次上传文件，基础资料补全延后到人工确认之后')

    # 基础资料现在已经在 input_sheets 中。基于完整源集合重新整理文件关系，
    # 后续 AI 和人工弹窗由此看到“本次上传 + 自动补全”的全部文件与 Sheet。
    final_structural = matcher.match_headers_only(
        meta.train_sheets, input_sheets, None, matching_context)
    if final_structural.get('success'):
        meta.auto_renamed = []
        meta.rename_candidates = []
    else:
        from .structural_source_mapping import suggest_file_relations
        meta.auto_renamed, meta.rename_candidates = suggest_file_relations(
            meta.train_sheets, input_sheets,
            (final_structural.get('mapping') or {}).get('file_mapping') or {},
            salary_year, salary_month)

    # 每个基础资料文件的训练角色在补全时已经确定。用代码单独完成其结构映射并
    # 锁定，避免它与本次上传一起进入 AI 候选或人工 Sheet/字段选择框。逐文件
    # 匹配也能避免多个基础文件结构相似时互相形成歧义。
    for filled_info in meta.auto_filled:
        actual_name = filled_info.get('stored_file_name') or filled_info.get('file_name')
        expected_name = filled_info.get('file_name')
        base_actual = [s for s in input_sheets if s.get('file_name') == actual_name]
        base_training = [s for s in meta.train_sheets if s.get('file_name') == expected_name]
        if not base_actual or not base_training:
            continue
        base_result = matcher.match_headers_only(
            base_training, base_actual, None, matching_context)
        base_mapping = (base_result.get('mapping') or {}).get('file_mapping') or {}
        if base_result.get('success') and actual_name in base_mapping:
            locked_info = deepcopy(base_mapping[actual_name])
            locked_info['auto_filled'] = True
            locked_info['source_origin'] = base_actual[0].get('source_origin', 'base')
            meta.auto_filled_mapping[actual_name] = locked_info
        else:
            # 理论上基础资料与训练结构一致；若资产本身已失效，保留日志但绝不把
            # 它推入 AI/人工候选，防止用户被要求再次确认一个“已补全”的文件。
            logger.warning(
                "[Ingest] 基础资料已补入但无法按训练结构锁定，已从 AI/人工候选排除: %s → %s",
                actual_name, expected_name)

    meta.input_sheets = input_sheets
    meta.sheet_summary = _summaries_from_map(parsed_sheets_map)
    logger.info(
        f"[Ingest] 完成：{len(meta.signatures)} 个文件 / {len(input_sheets)} 个 sheet，"
        f"自动改名 {len(meta.auto_renamed)}、待确认改名 {len(meta.rename_candidates)}、"
        f"兜底 {len(meta.auto_filled)}、缺失 {len(meta.missing_files)}"
    )
    return meta, parsed_sheets_map


# ==================== 确认与映射解算（纯内存，无 Aspose） ====================

def resolve_with_confirmations(
    meta: IngestMeta,
    confirmed_renames: Optional[Dict[str, str]] = None,
    confirmed_mapping: Optional[Dict[str, Any]] = None,
    confirmed_target_map: Optional[Dict[str, str]] = None,
    script_content: Optional[str] = None,
    tenant_id: Optional[str] = None,
    salary_year: Optional[int] = None,
    salary_month: Optional[int] = None,
    use_history: Optional[bool] = None,
    template_override_path: Optional[str] = None,
    skip_history_check: bool = False,
    skipped_missing_files: Optional[List[str]] = None,
    allow_ai_matching: bool = True,
):
    """在表头层解算最终映射，收集全部待确认项。首轮与确认轮共用这段代码。

    「跳过」是显式决定，不是空白：`confirmed_renames`/`confirmed_target_map` 里
    **键存在即已表态**，值为空串表示"不映射/跳过"；`skipped_missing_files` 是用户
    确认可以缺失的文件。任何已表态的项都不再回传前端，否则同一个框会反复弹。

    Returns: PrecheckResult（额外挂 `_effective_renames`，供上层记录实际采用的改名）
    """
    from .compute_precheck import (
        PrecheckResult,
        _check_target_sheets, _check_history,
    )
    from .fast_header_matcher import FastHeaderMatcher
    active_ai_provider = meta.ai_provider_name if allow_ai_matching else None

    result = PrecheckResult()
    interactive_input_sheets = [
        s for s in meta.input_sheets if s.get('source_origin', 'upload') == 'upload'
    ]
    result.actual_sources = [{'file': s['file_name'], 'sheet': s['sheet_name'],
                              'original_file': s.get('original_file_name', s['file_name']),
                              'original_sheet': s.get('original_sheet_name', s['sheet_name']),
                              'origin': s.get('source_origin', 'upload'),
                              'asset_name': s.get('source_asset_name')}
                             for s in interactive_input_sheets]
    from .confirmed_source_mapping import fully_unmatched_sheets
    result.unmatched_columns = list((confirmed_mapping or {}).get('unmatched_columns', []))
    unmatched_sheets = fully_unmatched_sheets(meta.source_structure, result.unmatched_columns)
    result.auto_renamed = list(meta.auto_renamed or [])
    result.auto_filled = list(meta.auto_filled or [])
    # 用户确认"这个文件本月确实没有"→ 不再阻断（缺列走已有的非阻断告警路径）
    _skipped_files = {str(f) for f in (skipped_missing_files or []) if f}
    _skipped_files.update(f for f, info in meta.source_structure.get('files', {}).items()
                          if info.get('sheets') and all((f, s) in unmatched_sheets for s in info['sheets']))
    result.missing_files = [f for f in (meta.missing_files or []) if f not in _skipped_files]
    _accepted_missing = [f for f in (meta.missing_files or []) if f in _skipped_files]
    if _accepted_missing:
        logger.warning("[Ingest] 用户确认跳过缺失文件，继续计算: %s", _accepted_missing)
    if result.missing_files:
        result.ok = False

    if meta.parse_error:
        result.ok = False
        result.missing_columns = [{"file": "", "sheet": "", "expected_columns": [],
                                  "error": meta.parse_error}]
        return result

    # 改名：用户确认优先，未确认则用最高分候选试探（同时把候选继续报给前端）
    # 未确认的名称候选不能伪装成同名文件，干扰结构优先级或 AI 判断。
    effective_renames = {r['from']: r['to'] for r in meta.auto_renamed if r.get('decision') == 'period_role'}
    _raw_renames = confirmed_renames or {}
    _decided_files = {str(k) for k in _raw_renames.keys() if k}
    confirmed_renames = {str(k): str(v) for k, v in _raw_renames.items() if k and str(v).strip()}
    # 文件层必须一一对应：多个上传文件选择同一个训练文件时直接阻断，不能进入匹配。
    _target_to_uploads = {}
    for _upload, _target in confirmed_renames.items():
        _target_to_uploads.setdefault(_target, []).append(_upload)
    _dup_file_targets = {target: uploads for target, uploads in _target_to_uploads.items()
                         if len(uploads) > 1}
    if _dup_file_targets:
        result.ok = False
        _dup_uploads = {upload for uploads in _dup_file_targets.values() for upload in uploads}
        _existing_rename = {str(c.get('uploaded')): c for c in (meta.rename_candidates or [])
                            if c.get('uploaded')}
        result.rename_candidates = []
        for _upload in sorted(_dup_uploads):
            _candidate = dict(_existing_rename.get(_upload) or {
                'uploaded': _upload, 'candidates': [],
            })
            _candidate['candidates'] = list(_candidate.get('candidates') or [])
            result.rename_candidates.append(_candidate)
        _dup_desc = '；'.join(f"{target} ← {','.join(uploads)}"
                             for target, uploads in _dup_file_targets.items())
        result.mapping_notice = f'多个上传文件不能映射到同一个训练文件：{_dup_desc}'
        result.missing_columns = [{
            'file': '', 'sheet': '', 'expected_columns': [], 'error': result.mapping_notice,
        }]
        result.actual_paths = _actual_paths(interactive_input_sheets)
        if not skip_history_check:
            _check_history(script_content, tenant_id, salary_year, salary_month, result, use_history)
        _check_target_sheets(script_content, tenant_id, template_override_path,
                             confirmed_target_map, result, active_ai_provider)
        return result
    # 明确选了"不映射"的文件：撤掉试探改名，别背着用户按最高分改
    for _skip_name in _decided_files - set(confirmed_renames):
        effective_renames.pop(_skip_name, None)
    effective_renames.update(confirmed_renames)
    unconfirmed = [c for c in (meta.rename_candidates or [])
                   if c.get("uploaded") not in _decided_files]
    if unconfirmed:
        result.ok = False
        result.rename_candidates = unconfirmed
    result._effective_renames = effective_renames

    # 先固定人工选择，仅剩余表参与自动匹配，避免再次推断覆盖确认结果。
    auto_filled_locked = deepcopy(meta.auto_filled_mapping or {})
    locked = deepcopy(auto_filled_locked)
    if confirmed_mapping:
        from .confirmed_source_mapping import apply_confirmed_mapping
        try:
            locked = apply_confirmed_mapping(meta, locked, confirmed_mapping)
        except ValueError as exc:
            result.ok = False
            # 即使列关系尚不完整，也要保留已经人工指定的文件/Sheet，并补齐真实
            # file_path。最终审核放行后，计算进程才能按该关系生成对应的 ``源_*``。
            result.file_mapping = deepcopy(
                (confirmed_mapping or {}).get('file_mapping', confirmed_mapping) or {})
            _actual_by_pair = {(s['file_name'], s['sheet_name']): s for s in meta.input_sheets}
            for _filename, _info in result.file_mapping.items():
                _mapped_sheets = list((_info.get('sheet_mapping') or {}).keys())
                _source = next((_actual_by_pair.get((_filename, sn)) for sn in _mapped_sheets
                                if _actual_by_pair.get((_filename, sn)) is not None), None)
                if _source is not None:
                    _info['file_path'] = _source['file_path']
                    _info['needs_rewrite'] = True
            result.actual_paths = _actual_paths(interactive_input_sheets)
            result.missing_columns = [{
                'file': '', 'sheet': '', 'expected_columns': [], 'error': str(exc),
            }]
            result.ai_suggestions = []
            # 一次展示剩余的所有类别，避免修完源列后下一轮才出现目标表/历史确认。
            if not skip_history_check:
                _check_history(script_content, tenant_id, salary_year, salary_month, result, use_history)
            _check_target_sheets(script_content, tenant_id, template_override_path,
                                 confirmed_target_map, result, active_ai_provider)
            return result
    locked_targets = {(info['expected_file'], sheet) for info in locked.values()
                      for sheet in info['sheet_mapping'].values()}
    locked_targets.update(unmatched_sheets)
    locked_inputs = {(filename, sheet) for filename, info in locked.items() for sheet in info['sheet_mapping']}
    train_remaining = [s for s in meta.train_sheets
                       if (s['file_name'], s['sheet_name']) not in locked_targets
                       and s['file_name'] not in _skipped_files]
    input_remaining = [s for s in interactive_input_sheets
                       if (s['file_name'], s['sheet_name']) not in locked_inputs]
    effective_renames.update({filename: info['expected_file'] for filename, info in locked.items()})
    virtual_sheets, xlate = _build_virtual_sheets(input_remaining, effective_renames, None)
    # File-level confirmation is a constraint, not merely a filename score bonus.
    # Leave unconfirmed probes free to participate in normal automatic matching.
    for sheet in virtual_sheets:
        actual_file = os.path.basename(sheet['file_path'])
        # Matching keys and AI suggestions must retain the physical upload identity.
        sheet['file_name'] = actual_file
        if actual_file in effective_renames:
            sheet['_confirmed_file'] = effective_renames[actual_file]
    if (locked or result.unmatched_columns) and not train_remaining:
        match_result = {"success": True, "mapping": {"file_mapping": {}}}
    elif not virtual_sheets:
        match_result = {"success": False, "error": "上传的文件无法读取或为空"}
    elif train_remaining:
        match_result = FastHeaderMatcher().match_headers_only(
            train_remaining, virtual_sheets, active_ai_provider,
            {'script_content': script_content or '',
             'template_name': os.path.basename(template_override_path or ''),
             'target_sheets': list((meta.expected_structure or {}).get('sheets', {})),
             'target_columns': {name: list(info.get('headers') or {}) for name, info in
                                (meta.expected_structure or {}).get('sheets', {}).items()
                                if isinstance(info, dict)},
             'salary_year': salary_year,
             'salary_month': salary_month,
             'original_file_names': meta.auto_renamed})
    else:
        match_result = {"success": True, "mapping": {"file_mapping": {}}}
    composed = deepcopy(auto_filled_locked)
    composed.update(_compose_file_mapping(
        (match_result.get("mapping") or {}).get("file_mapping") or {}, xlate))
    if 'AI 匹配未通过' in (match_result.get('error') or ''):
        result.mapping_notice = ((match_result.get('ai_failure_reason') or 'AI 推荐未通过校验')
                                 + '；已保留程序匹配结果，其余来源请确认。')
    elif not match_result.get('success') and match_result.get('match_method') == 'ai':
        result.mapping_notice = 'AI 已保留能够确定的文件和 Sheet；其余来源请确认。'
    if confirmed_mapping:
        composed = apply_confirmed_mapping(meta, composed, confirmed_mapping)
    if locked:
        # 文件/Sheet 已被明确指定后，同文件的改名候选不再重复询问。
        result.rename_candidates = [c for c in result.rename_candidates if c.get('uploaded') not in locked]
        mapped_files = {info['expected_file'] for info in composed.values()}
        result.missing_files = [f for f in result.missing_files if f not in mapped_files]
        result.ok = not (result.rename_candidates or result.missing_files)
    if composed:
        # 已由结构或语义匹配解决的来源不再因旧的文件名候选重复弹窗。
        result.rename_candidates = [c for c in result.rename_candidates if c.get('uploaded') not in composed]
        mapped_files = {info['expected_file'] for info in composed.values()}
        result.missing_files = [f for f in result.missing_files if f not in mapped_files]
        result.ok = not (result.rename_candidates or result.missing_files)
        result.file_mapping = composed
        covered = {(info['expected_file'], sheet) for info in composed.values()
                   for sheet in info.get('sheet_mapping', {}).values()}
        remaining = [s for s in meta.train_sheets
                     if (s['file_name'], s['sheet_name']) not in covered | unmatched_sheets and s['file_name'] not in _skipped_files]
        # 智算只审核文件/Sheet。列差异由训练脚本按改名后的源文件处理，不再
        # 转成 missing_columns，也不再触发列级 AI/人工确认。
        if remaining:
            result.ok = False
            result.actual_paths = _actual_paths(interactive_input_sheets)
            result.source_sheet_reviews = [{
                'expected_file': s['file_name'], 'expected_sheet': s['sheet_name'],
                'suggested_file': '', 'suggested_sheet': '', 'confidence': None,
                'reason': '程序和 AI 均未确定来源，请选择对应文件/Sheet；若基础资料也没有则报告缺失文件。',
                'recommendation_source': 'manual',
            } for s in remaining]
    elif result.unmatched_columns and not train_remaining:
        result.file_mapping = {}
        result.rename_candidates = [c for c in result.rename_candidates
                                   if any(o.get('expected') not in _skipped_files for o in c.get('candidates', []))]
        result.ok = not (result.rename_candidates or result.missing_files)
    else:
        # 未确定的来源必须先确认，不能静默把未匹配文件交给脚本猜测。
        diagnostics = match_result.get("diagnostics") or {}
        result.ok = False
        result.actual_paths = diagnostics.get("actual_paths") or _actual_paths(interactive_input_sheets)
        result.source_sheet_reviews = [{
            'expected_file': s['file_name'], 'expected_sheet': s['sheet_name'],
            'suggested_file': '', 'suggested_sheet': '', 'confidence': None,
            'reason': match_result.get('error') or '无法自动确定来源，请选择对应文件/Sheet。',
            'recommendation_source': 'manual',
        } for s in train_remaining]

    if match_result.get('needs_confirmation'):
        result.ok = False
        result.mapping_requires_confirmation = True
        result.actual_paths = _actual_paths(interactive_input_sheets)
        # AI 仅输出文件/Sheet 推荐；禁止在此重新生成列建议。
        _existing = {(str(item.get('expected_file')), str(item.get('expected_sheet'))): item
                     for item in result.source_sheet_reviews or []}
        for item in (match_result.get('source_sheet_reviews') or []):
            key = (str(item.get('expected_file')), str(item.get('expected_sheet')))
            _existing[key] = item
        result.source_sheet_reviews = list(_existing.values())
        result.ai_suggestions = []
        result.missing_columns = []

    if not skip_history_check:
        _check_history(script_content, tenant_id, salary_year, salary_month, result, use_history)
    _check_target_sheets(script_content, tenant_id, template_override_path,
                         confirmed_target_map, result, active_ai_provider)
    return result


def _summaries_from_map(parsed_sheets_map: Dict[tuple, Any]) -> Dict[str, list]:
    """把解析结果压成资产登记用的表头摘要（与 _persist_source_file 原格式一致）。"""
    summaries: Dict[str, list] = {}
    for (file_path, sheet_name), sheet_data in (parsed_sheets_map or {}).items():
        headers: List[str] = []
        rows = 0
        regions = getattr(sheet_data, "regions", None) or []
        for region in regions:
            headers.extend(list((region.head_data or {}).keys()))
            rows += len(region.data or [])
        summaries.setdefault(os.path.basename(file_path), []).append(
            {"sheet_name": sheet_name, "rows": rows, "headers": headers[:50], "regions": len(regions)})
    return summaries


def _signatures_from_sheets(input_sheets: List[Dict[str, Any]]) -> Dict[str, dict]:
    """把解析结果压成改名评分需要的签名（复用同一份表头，避免二次开文件）。"""
    sigs: Dict[str, dict] = {}
    for s in input_sheets:
        entry = sigs.setdefault(s["file_name"], {"headers": set(), "sheet_names": set(), "sheets": []})
        cols = [str(h) for h in (s.get("headers") or {}).keys() if h and str(h).strip()]
        entry["headers"].update(cols)
        entry["sheet_names"].add(str(s["sheet_name"]))
        entry["sheets"].append({"name": s["sheet_name"], "headers": sorted(cols)[:50]})
    return sigs


def _remap_after_rename(
    source_dir: str,
    renamed: List[Dict[str, Any]],
    input_sheets: List[Dict[str, Any]],
    parsed_sheets_map: Dict[tuple, Any],
) -> Tuple[List[Dict[str, Any]], Dict[tuple, Any]]:
    """物理改名后，把内存里的 file_name/file_path 与 parsed_sheets_map 键一起搬过去。"""
    name_map = {r.get("from"): r.get("to") for r in renamed if r.get("from") and r.get("to")}
    if not name_map:
        return input_sheets, parsed_sheets_map
    path_map: Dict[str, str] = {}
    for s in input_sheets:
        new_name = name_map.get(s["file_name"])
        if not new_name:
            continue
        new_path = str(Path(source_dir) / new_name)
        path_map[s["file_path"]] = new_path
        s["file_name"] = new_name
        s["file_path"] = new_path
    if not path_map:
        return input_sheets, parsed_sheets_map
    remapped = {(path_map.get(fp, fp), sn): sd for (fp, sn), sd in parsed_sheets_map.items()}
    return input_sheets, remapped


def _build_virtual_sheets(
    input_sheets: List[Dict[str, Any]],
    renames: Dict[str, str],
    confirmed_mapping: Optional[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], Dict[tuple, dict]]:
    """按"用户确认后的样子"造一份虚拟表头，并记录回真实坐标的翻译表。

    xlate[(file_path, virtual_sheet)] = {
        "actual_sheet": 真实 sheet 名,
        "header_map": {真实列名: 虚拟列名},
    }
    """
    fm = confirmed_mapping or {}
    if isinstance(fm.get("file_mapping"), dict):
        fm = fm["file_mapping"]

    virtual: List[Dict[str, Any]] = []
    xlate: Dict[tuple, dict] = {}
    for s in input_sheets:
        actual_name = s["file_name"]
        info = fm.get(actual_name) if isinstance(fm.get(actual_name), dict) else {}
        sheet_mapping = info.get("sheet_mapping") or {}
        header_mapping = (info.get("header_mapping_by_sheet") or {}).get(
            s['sheet_name'], info.get("header_mapping") or {})
        v_file = info.get("expected_file") or renames.get(actual_name) or actual_name
        v_sheet = sheet_mapping.get(s["sheet_name"], s["sheet_name"])

        header_map: Dict[Any, Any] = {}
        v_headers: Dict[Any, Any] = {}
        for col, sample in (s.get("headers") or {}).items():
            v_col = header_mapping.get(col, col)
            header_map[col] = v_col
            v_headers[v_col] = sample

        key = (s["file_path"], v_sheet)
        if key in xlate:
            # 同一文件两张 sheet 被确认成同名：无法区分，保留先到的那张
            logger.warning(f"[Ingest] 确认映射把 {actual_name} 的多张 sheet 归并为 '{v_sheet}'，忽略后到的")
            continue
        xlate[key] = {"actual_sheet": s["sheet_name"], "header_map": header_map}
        virtual.append({
            "file_name": v_file,
            "file_path": s["file_path"],
            "sheet_name": v_sheet,
            "headers": v_headers,
            "original_file_name": s.get('original_file_name', actual_name),
            "original_sheet_name": s.get('original_sheet_name', s['sheet_name']),
        })
    return virtual, xlate


def _compose_file_mapping(match_fm: Dict[str, Any], xlate: Dict[tuple, dict]) -> Dict[str, Any]:
    """把"虚拟坐标下的匹配结果"复合回真实文件/sheet/列，供 build_preload 直接使用。

    同时保留匹配阶段产生的逐列/Sheet/文件置信度，供预检层判断是否需要列人工确认。
    """
    out: Dict[str, Any] = {}
    for _v_file, info in (match_fm or {}).items():
        if not isinstance(info, dict):
            continue
        file_path = info.get("file_path", "")
        actual_name = os.path.basename(file_path) or _v_file
        entry = out.setdefault(actual_name, {
            "expected_file": info.get("expected_file", _v_file),
            "sheet_mapping": {},
            "header_mapping": {},
            "header_mapping_by_sheet": {},
            "header_confidence_by_sheet": {},
            "sheet_confidence": {},
            "file_confidence": None,
            "needs_rewrite": False,
            "file_path": file_path,
        })
        _file_conf = info.get('file_confidence')
        if _file_conf is not None:
            try:
                entry["file_confidence"] = max(float(entry.get("file_confidence") or 0.0), float(_file_conf))
            except (TypeError, ValueError):
                pass
        for v_sheet, train_sheet in (info.get("sheet_mapping") or {}).items():
            m_headers = (info.get('header_mapping_by_sheet') or {}).get(
                v_sheet, info.get('header_mapping') or {})
            m_col_conf = (info.get('header_confidence_by_sheet') or {}).get(v_sheet) or {}
            m_sheet_conf = (info.get('sheet_confidence') or {}).get(v_sheet)
            tr = xlate.get((file_path, v_sheet))
            if tr is None:
                # 没有翻译记录说明该 sheet 未经虚拟改名，直接透传
                entry["sheet_mapping"][v_sheet] = train_sheet
                entry["header_mapping"].update(m_headers)
                entry['header_mapping_by_sheet'][v_sheet] = dict(m_headers)
                entry['header_confidence_by_sheet'][v_sheet] = {
                    str(col): float(conf) for col, conf in m_col_conf.items()
                    if conf is not None
                }
                if m_sheet_conf is not None:
                    entry['sheet_confidence'][v_sheet] = float(m_sheet_conf)
                continue
            actual_sheet = tr["actual_sheet"]
            entry["sheet_mapping"][actual_sheet] = train_sheet
            scoped = entry['header_mapping_by_sheet'].setdefault(actual_sheet, {})
            scoped_conf = entry['header_confidence_by_sheet'].setdefault(actual_sheet, {})
            for actual_col, v_col in tr["header_map"].items():
                # 只登记真正匹配到的训练列；多余源列仍可保留在 DataFrame，
                # 但不能出现在人工确认表中伪装成训练期望列。
                if v_col not in m_headers:
                    continue
                target_col = m_headers.get(v_col, v_col)
                entry["header_mapping"][actual_col] = target_col
                scoped[actual_col] = target_col
                raw_conf = m_col_conf.get(v_col)
                if raw_conf is not None:
                    try:
                        scoped_conf[str(actual_col)] = float(raw_conf)
                    except (TypeError, ValueError):
                        pass
            if m_sheet_conf is not None:
                entry['sheet_confidence'][actual_sheet] = float(m_sheet_conf)
    for entry in out.values():
        entry["needs_rewrite"] = (any(k != v for k, v in entry['sheet_mapping'].items()) or
            any(k != v for columns in entry['header_mapping_by_sheet'].values() for k, v in columns.items()))
    return out

def _actual_paths(input_sheets: List[Dict[str, Any]]) -> List[str]:
    return [f"{s['file_name']} > {s['sheet_name']} > {col}"
            for s in input_sheets for col in (s.get("headers") or {})]


def build_preload(meta: IngestMeta, parsed_sheets_map: Dict[tuple, Any],
                  file_mapping: Dict[str, Any]) -> Dict[str, Any]:
    """由最终映射 + 内存解析结果构建预加载源数据（含列改名，不落盘）。"""
    from .fast_header_matcher import FastHeaderMatcher
    return FastHeaderMatcher().build_preload(
        file_mapping, parsed_sheets_map,
        source_structure=meta.source_structure,
        expected_structure=meta.expected_structure,
    )


# ==================== 会话产物读写 ====================

def _dump(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as fh:
            pickle.dump(obj, fh, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def write_ingest(session_dir: str, meta: IngestMeta, parsed_sheets_map: Dict[tuple, Any]) -> None:
    """meta 与 sources 分文件落盘：确认轮只读 meta，大表不进 API 进程。"""
    base = Path(session_dir)
    _dump(base / _META_NAME, meta)
    _dump(base / _SOURCES_NAME, parsed_sheets_map)
    logger.info(f"[Ingest] 会话产物已写入 {base}（sheet 数 {len(parsed_sheets_map)}）")


def _load(path: Path) -> Optional[Any]:
    # 只读服务端自己写的产物；上传路径永远不会经过这里
    if not path.exists():
        return None
    try:
        with open(path, "rb") as fh:
            return pickle.load(fh)
    except Exception as e:
        logger.warning(f"[Ingest] 读取产物失败 {path}: {e}")
        return None


def read_meta(session_dir: str) -> Optional[IngestMeta]:
    meta = _load(Path(session_dir) / _META_NAME)
    if isinstance(meta, IngestMeta) and getattr(meta, "version", None) == INGEST_VERSION:
        return meta
    if meta is not None:
        logger.info("[Ingest] 会话产物版本不符，忽略")
    return None


def ingest_ready(session_dir: str) -> bool:
    """会话产物是否齐全（meta + sources 都在）。缺一半就必须退回重传路径，
    否则确认轮的虚拟改名无处落地（计算进程会回退到未改名的真实文件）。"""
    base = Path(session_dir)
    return (base / _META_NAME).exists() and (base / _SOURCES_NAME).exists()


def read_sources(session_dir: str) -> Optional[Dict[tuple, Any]]:
    data = _load(Path(session_dir) / _SOURCES_NAME)
    return data if isinstance(data, dict) else None
