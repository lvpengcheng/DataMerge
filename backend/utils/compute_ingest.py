"""智算源文件摄取（ingest）——整条链路里唯一"打开源文件"的地方。

设计要点（对应"点击计算后只读一次盘"的目标）：

1. **一次解析**：`ingest_source_dir` 内只调用 `FastHeaderMatcher.parse_inputs` 一次，
   每个上传文件仅被 Aspose 打开 1 次；改名评分复用同一份表头，不再单独开文件。
2. **一次确认**：改名候选、低置信列名、目标表歧义在同一轮里全部收集返回，
   不再"遇到改名歧义就 return"从而逼用户走三轮。
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
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# 解析逻辑或产物结构变化时 +1，旧会话产物自动失效
INGEST_VERSION = 4

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

    def candidate_targets(self) -> set:
        """所有改名候选可能占用的期望文件名（不能被基础资料抢先覆盖）。"""
        names = set()
        for cand in self.rename_candidates or []:
            for c in cand.get("candidates") or []:
                if c.get("expected"):
                    names.add(c["expected"])
        return names


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

    # 步骤 1：唯一一次全量解析（每文件 1 次 Aspose）
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

    # 文件名、Sheet 名改变不应先触发 AI。完整结构能唯一确定时直接复用。
    structural = matcher.match_headers_only(meta.train_sheets, input_sheets)
    structural_files = {info['expected_file'] for info in
                        (structural.get('mapping') or {}).get('file_mapping', {}).values()}
    if not structural['success']:
        from .structural_source_mapping import suggest_file_relations
        meta.auto_renamed, meta.rename_candidates = suggest_file_relations(
            meta.train_sheets, input_sheets,
            (structural.get('mapping') or {}).get('file_mapping') or {}, salary_year, salary_month)

    # 步骤 3：基础资料兜底；改名候选可能占用的期望名不参与兜底（避免覆盖用户文件）
    if db_session is not None and tenant_id:
        try:
            from .source_auto_filler import auto_fill_missing_sources
            filled, still_missing = auto_fill_missing_sources(
                source_dir=source_dir,
                source_structure=structure,
                tenant_id=tenant_id,
                db_session=db_session,
                assume_present=(meta.candidate_targets() | structural_files |
                                {r['to'] for r in meta.auto_renamed} |
                                {v for v in (confirmed_renames or {}).values() if v}),
            )
            meta.auto_filled = filled or []
            meta.missing_files = list(still_missing or [])
            # 步骤 3.1：只解析新补进来的文件（增量，不重解析已有文件）
            new_paths = [str(Path(source_dir) / f["file_name"]) for f in meta.auto_filled]
            new_paths = [p for p in new_paths if os.path.exists(p)]
            if new_paths:
                extra_sheets, extra_map = matcher.parse_inputs(
                    new_paths, manual_headers, multi_sheet_source=meta.multi_sheet_source)
                input_sheets.extend(extra_sheets)
                parsed_sheets_map.update(extra_map)
                meta.signatures = _signatures_from_sheets(input_sheets)
        except Exception as e:
            logger.warning(f"[Ingest] 基础资料兜底异常: {e}", exc_info=True)

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
):
    """在表头层解算最终映射，收集全部待确认项。首轮与确认轮共用这段代码。

    「跳过」是显式决定，不是空白：`confirmed_renames`/`confirmed_target_map` 里
    **键存在即已表态**，值为空串表示"不映射/跳过"；`skipped_missing_files` 是用户
    确认可以缺失的文件。任何已表态的项都不再回传前端，否则同一个框会反复弹。

    Returns: PrecheckResult（额外挂 `_effective_renames`，供上层记录实际采用的改名）
    """
    from .compute_precheck import (
        PrecheckResult, _extract_missing_columns, _suggest_structural_columns,
        _check_target_sheets, _check_history,
    )
    from .fast_header_matcher import FastHeaderMatcher

    result = PrecheckResult()
    result.actual_sources = [{'file': s['file_name'], 'sheet': s['sheet_name'],
                              'original_file': s.get('original_file_name', s['file_name']),
                              'original_sheet': s.get('original_sheet_name', s['sheet_name'])}
                             for s in meta.input_sheets]
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
    locked = {}
    if confirmed_mapping:
        from .confirmed_source_mapping import apply_confirmed_mapping
        try:
            locked = apply_confirmed_mapping(meta, {}, confirmed_mapping)
        except ValueError as exc:
            result.ok = False
            result.file_mapping = (confirmed_mapping or {}).get('file_mapping', confirmed_mapping)
            result.actual_paths = _actual_paths(meta.input_sheets)
            result.missing_columns = _extract_missing_columns(
                meta.source_structure, [s['file_path'] for s in meta.input_sheets], str(exc))
            result.ai_suggestions = _suggest_structural_columns(
                result.missing_columns, result.actual_paths, result.file_mapping, confirmed_renames)
            # 一次展示剩余的所有类别，避免修完源列后下一轮才出现目标表/历史确认。
            if not skip_history_check:
                _check_history(script_content, tenant_id, salary_year, salary_month, result, use_history)
            _check_target_sheets(script_content, tenant_id, template_override_path,
                                 confirmed_target_map, result)
            return result
    locked_targets = {(info['expected_file'], sheet) for info in locked.values()
                      for sheet in info['sheet_mapping'].values()}
    locked_targets.update(unmatched_sheets)
    locked_inputs = {(filename, sheet) for filename, info in locked.items() for sheet in info['sheet_mapping']}
    train_remaining = [s for s in meta.train_sheets
                       if (s['file_name'], s['sheet_name']) not in locked_targets
                       and s['file_name'] not in _skipped_files]
    input_remaining = [s for s in meta.input_sheets
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
            train_remaining, virtual_sheets, meta.ai_provider_name,
            {'script_content': script_content or '',
             'template_name': os.path.basename(template_override_path or ''),
             'target_sheets': list((meta.expected_structure or {}).get('sheets', {})),
             'target_columns': {name: list(info.get('headers') or {}) for name, info in
                                (meta.expected_structure or {}).get('sheets', {}).items()
                                if isinstance(info, dict)},
             'original_file_names': meta.auto_renamed})
    else:
        match_result = {"success": True, "mapping": {"file_mapping": {}}}
    composed = _compose_file_mapping(
        (match_result.get("mapping") or {}).get("file_mapping") or {}, xlate)
    if 'AI 匹配未通过' in (match_result.get('error') or ''):
        result.mapping_notice = ((match_result.get('ai_failure_reason') or 'AI 推荐未通过校验')
                                 + '；已保留程序匹配结果，其余来源请确认。')
    elif not match_result.get('success') and match_result.get('match_method') == 'ai':
        result.mapping_notice = 'AI 已保留能够确定的匹配；剩余缺失或不确定的来源、字段请确认。'
    if locked:
        composed = apply_confirmed_mapping(meta, composed, confirmed_mapping)
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
        missing_columns = [{'file': s['file_name'], 'sheet': s['sheet_name'],
                            'expected_columns': list(s['headers'])} for s in remaining]
        # A high similarity score is not full column coverage. Surface the few
        # unresolved columns now, rather than discarding all mappings next round.
        skipped_columns = {tuple(item) for item in result.unmatched_columns}
        mapped_columns = {}
        for info in composed.values():
            for source_sheet, target_sheet in info.get('sheet_mapping', {}).items():
                columns = (info.get('header_mapping_by_sheet') or {}).get(
                    source_sheet, info.get('header_mapping') or {})
                mapped_columns.setdefault((info['expected_file'], target_sheet), set()).update(columns.values())
        for sheet in meta.train_sheets:
            key = (sheet['file_name'], sheet['sheet_name'])
            if key not in covered:
                continue
            missing = [col for col in sheet['headers'] if col not in mapped_columns.get(key, set())
                       and (*key, col) not in skipped_columns]
            if missing:
                missing_columns.append({'file': key[0], 'sheet': key[1], 'expected_columns': missing})
        if missing_columns:
            result.ok = False
            result.actual_paths = _actual_paths(meta.input_sheets)
            result.missing_columns = missing_columns
            result.ai_suggestions = _suggest_structural_columns(
                result.missing_columns, result.actual_paths, composed, effective_renames)
    elif result.unmatched_columns and not train_remaining:
        result.file_mapping = {}
        result.rename_candidates = [c for c in result.rename_candidates
                                   if any(o.get('expected') not in _skipped_files for o in c.get('candidates', []))]
        result.ok = not (result.rename_candidates or result.missing_files)
    else:
        # 未确定的来源必须先确认，不能静默把未匹配文件交给脚本猜测。
        diagnostics = match_result.get("diagnostics") or {}
        result.ok = False
        result.actual_paths = diagnostics.get("actual_paths") or _actual_paths(meta.input_sheets)
        missing = _extract_missing_columns(
            meta.source_structure, [s["file_path"] for s in meta.input_sheets],
            match_result.get("error"))
        result.missing_columns = missing
        result.ai_suggestions = _suggest_structural_columns(missing, result.actual_paths)

    if match_result.get('needs_confirmation'):
        result.ok = False
        result.mapping_requires_confirmation = True
        result.actual_paths = _actual_paths(meta.input_sheets)
        result.ai_suggestions.extend(match_result.get('ai_suggestions') or [])

    if not skip_history_check:
        _check_history(script_content, tenant_id, salary_year, salary_month, result, use_history)
    _check_target_sheets(script_content, tenant_id, template_override_path,
                         confirmed_target_map, result)
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
    """把"虚拟坐标下的匹配结果"复合回真实文件/sheet/列，供 build_preload 直接使用。"""
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
            "needs_rewrite": False,
            "file_path": file_path,
        })
        for v_sheet, train_sheet in (info.get("sheet_mapping") or {}).items():
            m_headers = (info.get('header_mapping_by_sheet') or {}).get(
                v_sheet, info.get('header_mapping') or {})
            tr = xlate.get((file_path, v_sheet))
            if tr is None:
                # 没有翻译记录说明该 sheet 未经虚拟改名，直接透传
                entry["sheet_mapping"][v_sheet] = train_sheet
                entry["header_mapping"].update(m_headers)
                entry['header_mapping_by_sheet'][v_sheet] = dict(m_headers)
                continue
            entry["sheet_mapping"][tr["actual_sheet"]] = train_sheet
            scoped = entry['header_mapping_by_sheet'].setdefault(tr['actual_sheet'], {})
            for actual_col, v_col in tr["header_map"].items():
                # 只登记真正匹配到的训练列；多余源列仍可保留在 DataFrame，
                # 但不能出现在人工确认表中伪装成训练期望列。
                if v_col not in m_headers:
                    continue
                entry["header_mapping"][actual_col] = m_headers.get(v_col, v_col)
                scoped[actual_col] = m_headers.get(v_col, v_col)
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
