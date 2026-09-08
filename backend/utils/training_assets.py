"""Persistent, revisioned training inputs shared by history, downloads and edits."""
import shutil
import uuid
from pathlib import Path, PureWindowsPath

EXCEL_SUFFIXES = {'.xlsx', '.xls', '.xlsm'}
RULE_SUFFIXES = EXCEL_SUFFIXES | {'.txt', '.md', '.pdf', '.docx'}
PUBLIC_SETTINGS = ('ai_provider', 'salary_year', 'salary_month', 'monthly_standard_hours',
                   'manual_headers', 'multi_sheet_source', 'use_history', 'mode')


def basename(value):
    return PureWindowsPath(str(value or '')).name


def excel_files(directory):
    directory = Path(directory) if directory else None
    if directory is None or not directory.is_dir():
        return []
    return sorted((p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES
                   and not p.name.startswith('~')), key=lambda p: p.name.casefold())


def resolve_asset_config(config, session_root):
    """Recover this session's persisted files after a host/path migration; never scan other sessions."""
    cfg = dict(config or {})
    root = Path(session_root)
    for key in ('expected_file', 'template_path'):
        original = cfg.get(key)
        if original and Path(original).is_file():
            continue
        name = basename(original)
        candidates = [root / name, root / 'expected' / name] if name else []
        # Rebase a revision path after moving from Windows to Docker or another disk.
        parts = PureWindowsPath(str(original or '')).parts
        if 'revisions' in parts:
            candidates = [root.joinpath(*parts[parts.index('revisions'):])]
        cfg[key] = next((str(p) for p in candidates if p.is_file()), original)
    original_source = cfg.get('source_dir')
    parts = PureWindowsPath(str(original_source or '')).parts
    if not excel_files(original_source) and 'revisions' in parts:
        candidate = root.joinpath(*parts[parts.index('revisions'):])
        if excel_files(candidate):
            cfg['source_dir'] = str(candidate)
    if not excel_files(cfg.get('source_dir')) and 'revisions' not in parts and excel_files(root / 'source'):
        cfg['source_dir'] = str(root / 'source')
    original_rules = cfg.get('rule_files_dir')
    if original_rules and not Path(original_rules).is_dir():
        parts = PureWindowsPath(str(original_rules)).parts
        candidate = root.joinpath(*parts[parts.index('revisions'):]) if 'revisions' in parts else root / 'rule_files'
        if candidate.is_dir():
            cfg['rule_files_dir'] = str(candidate)
    return cfg


def asset_manifest(config, session_root):
    cfg = resolve_asset_config(config, session_root)
    expected = cfg.get('expected_file')
    rules_dir = cfg.get('rule_files_dir')
    rules = sorted(p.name for p in Path(rules_dir).iterdir() if p.is_file()) if rules_dir and Path(rules_dir).is_dir() else []
    return {
        'source_file_names': [p.name for p in excel_files(cfg.get('source_dir'))],
        'expected_file_name': basename(expected) if expected and Path(expected).is_file() else None,
        'rule_file_names': rules,
        'settings': {key: cfg.get(key) for key in PUBLIC_SETTINGS},
        'input_revision': cfg.get('input_revision'),
        'validation_stale': bool(cfg.get('validation_stale')),
    }


def validate_settings(settings):
    if not isinstance(settings, dict):
        raise ValueError('配置必须是 JSON 对象')
    if set(settings) - set(PUBLIC_SETTINGS):
        raise ValueError('包含不支持的配置项')
    cfg = dict(settings)
    if 'ai_provider' in cfg and cfg['ai_provider'] not in ('openai', 'claude', 'deepseek', 'ollama', 'local'):
        raise ValueError('不支持的 AI 提供者')
    if 'mode' in cfg and cfg['mode'] not in ('formula', 'auto', 'template', 'direct'):
        raise ValueError('不支持的生成模式')
    for key, low, high in [('salary_year', 1900, 9999), ('salary_month', 1, 12)]:
        if cfg.get(key) is not None and (type(cfg[key]) is not int or not low <= cfg[key] <= high):
            raise ValueError(f'{key} 超出有效范围')
    if cfg.get('monthly_standard_hours') is not None:
        import math
        value = float(cfg['monthly_standard_hours'])
        if not math.isfinite(value) or value <= 0:
            raise ValueError('标准工时必须大于零')
        cfg['monthly_standard_hours'] = value
    for key in ('multi_sheet_source', 'use_history'):
        if key in cfg and type(cfg[key]) is not bool:
            raise ValueError(f'{key} 必须是布尔值')
    if cfg.get('manual_headers') is not None and not isinstance(cfg['manual_headers'], dict):
        raise ValueError('手动表头必须是 JSON 对象')
    return cfg


def stage_revision(config, session_root, uploads, settings, source_mode='merge'):
    """Copy-on-write: validation/DB failures leave every previous revision untouched."""
    if source_mode not in ('merge', 'replace'):
        raise ValueError('源文件更新方式无效')
    cfg = resolve_asset_config(config, session_root)
    settings = validate_settings(settings)
    changed = any(cfg.get(k) != v for k, v in settings.items() if k != 'ai_provider')
    cfg.update(settings)
    if not any(uploads.values()):
        return cfg, None, changed
    revision = Path(session_root) / 'revisions' / uuid.uuid4().hex
    revision.mkdir(parents=True)
    if uploads.get('source'):
        target = revision / 'source'
        target.mkdir()
        if source_mode == 'merge':
            for path in excel_files(cfg.get('source_dir')):
                shutil.copy2(path, target / path.name)
        seen = set()
        for path in uploads['source']:
            path = Path(path)
            if path.name.casefold() in seen:
                raise ValueError(f'本次上传源文件重名: {path.name}')
            seen.add(path.name.casefold())
            for old in target.iterdir():
                if old.name.casefold() == path.name.casefold():
                    old.unlink()
            shutil.copy2(path, target / path.name)
        cfg['source_dir'] = str(target)
    if uploads.get('expected'):
        target = revision / 'expected'
        target.mkdir()
        path = Path(uploads['expected'][0])
        shutil.copy2(path, target / path.name)
        cfg['expected_file'] = str(target / path.name)
        cfg.pop('target_sheets', None)  # Sheet selections belonged to the previous workbook.
    if uploads.get('rules'):
        cfg.setdefault('rule_base_content', cfg.get('rules_content', '') if not cfg.get('rule_files_dir') else '')
        target = revision / 'rules'
        target.mkdir()
        old_dir = cfg.get('rule_files_dir')
        if old_dir and Path(old_dir).is_dir():
            for path in Path(old_dir).iterdir():
                if path.is_file():
                    shutil.copy2(path, target / path.name)
        for path in uploads['rules']:
            path = Path(path)
            for old in target.iterdir():
                if old.name.casefold() == path.name.casefold():
                    old.unlink()
            shutil.copy2(path, target / path.name)
        cfg['rule_files_dir'] = str(target)
    return cfg, revision, True


def prepare_revision(config, revision_path, changed_categories):
    """Native workbook work runs only inside the existing bounded subprocess runner."""
    from backend.api.training_chat import (_prepare_training_uploads_subprocess,
        _build_source_structure_from_dir_impl, _analyze_expected_structure_impl)
    cfg = dict(config)
    if revision_path:
        revision = Path(revision_path)
        source = revision / 'source'
        source.mkdir(exist_ok=True)
        result = _prepare_training_uploads_subprocess({
            'source_dir': str(source),
            'expected_file': cfg.get('expected_file') if 'expected' in changed_categories else None,
            'passwords': cfg.get('file_passwords') or {}, 'mode': cfg.get('mode')})
        if result.get('error_type'):
            raise ValueError(f"文件需要有效密码: {result.get('files')}")
        if 'expected' in changed_categories:
            cfg['expected_file'] = result['expected_file']
    has_sources = bool(excel_files(cfg.get('source_dir')))
    sources = (_build_source_structure_from_dir_impl(cfg.get('source_dir', ''),
        cfg.get('manual_headers'), cfg.get('multi_sheet_source', False)) if has_sources
        else {'files': {}, 'total_sheets': 0, 'total_regions': 0})
    if has_sources and (not sources.get('total_sheets') or any(v.get('error') or not v.get('sheets')
                                             for v in sources.get('files', {}).values())):
        raise ValueError('源文件结构解析失败，请检查文件、密码及手动表头')
    expected = cfg.get('expected_file')
    if expected and Path(expected).is_file():
        cfg['expected_structure'] = _analyze_expected_structure_impl(expected)
        if not any(sheet.get('headers') for sheet in cfg['expected_structure'].get('sheets', {}).values()):
            raise ValueError('目标文件没有可识别的数据区域')
    if cfg.get('mode') == 'template':
        cfg['template_path'] = expected
    if 'rules' in changed_categories:
        from backend.ai_engine.document_parser import get_document_parser
        parts = [cfg.get('rule_base_content', '')]
        for path in sorted(Path(cfg['rule_files_dir']).iterdir()):
            parsed = get_document_parser().parse_document(str(path))
            if not parsed.strip():
                raise ValueError(f'规则文件未能解析: {path.name}')
            parts.append(f'=== 规则文件: {path.name} ===\n{parsed}')
        cfg['rules_content'] = '\n\n'.join(parts)
    return {'config': cfg, 'source_structure': sources}
