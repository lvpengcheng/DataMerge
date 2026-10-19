"""Lightweight final approval and deferred validation with visible execution logs."""
import asyncio
import json
import os
import time
from pathlib import Path


def verify_confirmed_source_mapping(confirmed, actual):
    """Check approved relations without requiring every historical training sheet."""
    approved = (confirmed or {}).get('file_mapping', confirmed or {})
    has_selected_sheet = False
    for filename, info in approved.items():
        if not isinstance(info, dict):
            continue
        selected = info.get('sheet_mapping') or {}
        if not selected:
            continue
        has_selected_sheet = True
        effective = (actual or {}).get(filename) or {}
        expected_file = info.get('expected_file') or filename
        if (effective.get('expected_file') or filename) != expected_file:
            raise ValueError(f'人工确认的源文件关系未被采用，已停止计算：{filename} → {expected_file}')
        for sheet, target in selected.items():
            if effective.get('sheet_mapping', {}).get(sheet) != target:
                raise ValueError(f'人工确认的源 Sheet 关系未被采用，已停止计算：{filename} / {sheet} → {expected_file} / {target}')
    return has_selected_sheet


def validate_target_confirmation(script_content, rows, confirmed):
    from .compute_precheck import _extract_colmap
    keys = list((_extract_colmap(script_content or '') or {}).keys())
    reviews = {row['key']: row for row in rows or [] if row.get('key')}
    keys = list(dict.fromkeys(keys + list(reviews)))
    claimed = set()
    for key in keys:
        if key not in (confirmed or {}):
            raise ValueError(f'请确认关键目标 Sheet「{key}」的对应关系或明确跳过')
        value = confirmed[key]
        if not isinstance(value, str):
            raise ValueError(f'目标 Sheet「{key}」的选择无效')
        if not value:
            continue
        if value not in reviews.get(key, {}).get('all_sheets', []):
            raise ValueError(f'目标 Sheet「{key}」选择的模板表不存在，请重新选择')
        identity = (reviews.get(key, {}).get('workbook', ''), value)
        if identity in claimed:
            raise ValueError(f'模板 Sheet「{value}」被重复选择，请修正对应关系')
        claimed.add(identity)


async def prepare_confirmed_compute(params_file, emit, runner=None):
    """Run expensive post-confirmation work after the UI has received a task id."""
    path = Path(params_file)
    params = json.loads(path.read_text(encoding='utf-8'))
    payload = params.get('_post_confirmation_checks')
    if not payload:
        return
    if runner is None:
        from .subprocess_runner import run_in_fresh_subprocess_async
        runner = run_in_fresh_subprocess_async
    emit('info', '已接受最终匹配关系；开始后台校验和基础资料补全，保持人工选择，不再调用 AI。')
    for key, sheet in (params.get('target_sheet_manual_map') or {}).items():
        emit('info', f'目标 Sheet：{key} → {sheet}' if sheet else f'目标 Sheet：{key} → 本次明确跳过（不填充结果）')
    started = time.monotonic()
    operation = asyncio.create_task(runner(
        'backend.app.main:_compute_upload_precheck_subprocess', args=(payload,),
        timeout=int(os.getenv('COMPUTE_PRECHECK_TIMEOUT', '600')),
        max_memory_mb=int(os.getenv('COMPUTE_PRECHECK_MAX_MEMORY_MB', '1536'))))
    try:
        while not operation.done():
            done, _ = await asyncio.wait({operation}, timeout=10)
            if not done:
                emit('info', f'后台校验与基础资料补全进行中（已用时 {int(time.monotonic() - started)} 秒）')
        completed = await operation
    finally:
        if not operation.done():
            operation.cancel()
            await asyncio.gather(operation, return_exceptions=True)
    if not completed.success:
        raise ValueError(f'后台校验失败：{completed.error}')
    result = completed.result or {}
    if result.get('encrypted_files'):
        raise ValueError('后台校验失败：补全文件无法解密')
    check = result.get('pc_result')
    if check is None:
        raise ValueError('后台校验未生成结果，已停止计算')
    if getattr(check, '_reused_source_parse', False):
        emit('info', '已复用上传文件的解析结果，未重复打开源 Excel；仅解析新增基础资料。')
    for item in check.auto_filled or []:
        emit('info', f'基础资料补全：{item.get("file_name") or item.get("file")}（{item.get("source") or "基础资料"}）')
    if check.missing_files:
        raise ValueError('基础资料补全后仍缺少源文件，已停止计算：' + '、'.join(check.missing_files))
    if getattr(check, 'target_confirmation_required', False):
        raise ValueError('关键目标 Sheet 的确认关系无法应用，已停止计算')
    errors = [row['error'] for row in check.missing_columns or [] if row.get('error')]
    if errors:
        raise ValueError('后台校验未通过，已停止计算：' + '；'.join(errors))
    approved_relations_verified = False
    if payload.get('mapping_finalized'):
        approved_relations_verified = verify_confirmed_source_mapping(
            payload.get('confirmed_mapping'), check.file_mapping)
    if check.source_sheet_reviews or check.rename_candidates:
        pending = '、'.join(f'{row.get("expected_file", "")} / {row.get("expected_sheet", "")}'
                           for row in check.source_sheet_reviews or [])
        if not approved_relations_verified:
            raise ValueError('后台校验发现源文件或 Sheet 关系仍不完整，已停止计算：' + pending)
        if pending:
            emit('warning', '以下训练源 Sheet 未纳入本次匹配，仅记录提示，不覆盖已确认关系：' + pending)
        if check.rename_candidates:
            names = '、'.join(str(row.get('uploaded') or '') for row in check.rename_candidates)
            emit('warning', '仍有未选用的文件候选，按最终确认关系计算：' + names)
    if check.mapping_notice:
        emit('warning', check.mapping_notice)
    for warning in check.history_warnings or []:
        emit('warning', str(warning))
    for filename, info in (check.file_mapping or {}).items():
        for sheet, target in (info.get('sheet_mapping') or {}).items():
            emit('info', f'源 Sheet：{filename} / {sheet} → {info.get("expected_file") or filename} / {target}')
    params.update(pre_validated_mapping=check.file_mapping,
                  unmatched_columns=check.unmatched_columns,
                  precheck_auto_filled=check.auto_filled,
                  source_mapping_warning=None,
                  template_override_path=result.get('template_override_path') or params.get('template_override_path'))
    params.pop('_post_confirmation_checks', None)
    # Replacing the parameter file is the handoff: compute cannot observe partial JSON.
    import tempfile
    fd, staged = tempfile.mkstemp(dir=str(path.parent), suffix='.json.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(params, stream, ensure_ascii=False, default=str)
        os.replace(staged, path)
    finally:
        Path(staged).unlink(missing_ok=True)
    emit('info', f'后台校验与基础资料补全完成（用时 {time.monotonic() - started:.1f} 秒），开始执行计算脚本。')
