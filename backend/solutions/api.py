"""Authenticated solution lifecycle. Importing Python code requires administrator trust."""
import asyncio
import os
import shutil
import tempfile
import uuid
import hashlib
import re
from datetime import datetime, timezone
from pathlib import Path
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Body
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
from sqlalchemy.orm import Session
from ..database.connection import get_db
from ..database.models import CalculationSolution, Script, TrainingSession, TenantAuthorization
from ..auth.dependencies import get_current_user, get_operable_tenants, has_permission
from .store import ROOT, STORE, artifact, build, tenant_name
from .package import verify, export_zip, extract_zip
from .environment import require_environment

router = APIRouter(prefix='/api/solutions', tags=['计算方案'])


def allowed(tenant, tenants, user):
    if tenant not in tenants and not has_permission(user, 'admin'):
        raise HTTPException(403, '无权操作此租户')


def get_row(db, identity, tenants, user):
    row = db.query(CalculationSolution).filter_by(id=identity).first()
    if not row:
        raise HTTPException(404, '方案不存在')
    allowed(row.tenant_id, tenants, user)
    return row


def describe(row):
    return {'id': row.id, 'name': row.name, 'tenant_id': row.tenant_id, 'digest': row.digest,
            'version': row.digest[:12], 'status': row.status, 'is_active': row.is_active,
            'script_id': row.script_id, 'session_id': row.source_session_id, 'validation': row.validation}


def validation_result(data, digest):
    missing = []
    complete = bool(data.get('total_cells', 0)) and not data.get('error') and data.get('comparison_complete') is not False
    warning = bool(data.get('formula_warning'))
    def inspect_result(result, prefix=''):
        nonlocal complete, warning
        if not isinstance(result, dict):
            return
        missing.extend(prefix + str(sheet) for sheet in result.get('missing_sheets') or [])
        warning = warning or bool(result.get('formula_warning'))
        if result.get('error') or result.get('comparison_complete') is False:
            complete = False
        children = result.get('diff_details')
        if isinstance(children, dict):
            for name, child in children.items():
                inspect_result(child, prefix + str(name) + '/')
    inspect_result(data)
    return {'digest': digest, 'passed': complete and data.get('total_differences', 1) == 0 and not warning and not missing,
            'completed': complete, 'formula_warning': warning, 'missing_sheets': missing,
            'total_cells': data.get('total_cells'), 'total_differences': data.get('total_differences'),
            'accuracy': data.get('accuracy') if data.get('accuracy') is not None else
                        data.get('matched_cells', 0) / max(1, data.get('total_cells', 0)),
            'details': {k: data[k] for k in ('per_sheet', 'diff_details', 'field_diff_samples', 'detailed_text', 'detailed_diff', 'formula_report') if k in data}}


@router.get('')
def listing(tenant_id: str, db: Session = Depends(get_db), user=Depends(get_current_user), tenants=Depends(get_operable_tenants)):
    allowed(tenant_id, tenants, user)
    tenant_name(tenant_id)
    rows = db.query(Script).filter_by(tenant_id=tenant_id).all()
    scripts = [{'id': r.id, 'name': r.name, 'session_id': r.source_session_id} for r in rows]
    hashes = {'script_' + hashlib.md5(r.code.encode()).hexdigest()[:12] for r in rows}
    for file in (ROOT / 'tenants' / tenant_id / 'scripts').glob('script_*.py'):
        if re.fullmatch(r'script_[0-9a-f]{12}', file.stem) and file.stem not in hashes:
            scripts.append({'id': file.stem, 'name': file.stem + '（旧脚本）', 'session_id': None})
    return {'solutions': [describe(r) for r in db.query(CalculationSolution).filter_by(tenant_id=tenant_id)
                          .order_by(CalculationSolution.created_at.desc()).all()],
            'scripts': scripts}


@router.post('')
def convert(body: dict = Body(...), db: Session = Depends(get_db), user=Depends(get_current_user), tenants=Depends(get_operable_tenants)):
    identifier = str(body.get('script_id') or '')
    script = db.query(Script).filter_by(id=int(identifier)).first() if identifier.isdigit() else None
    if re.fullmatch(r'script_[0-9a-f]{12}', identifier):
        tenant = tenant_name(body.get('tenant_id'))
        allowed(tenant, tenants, user)
        from .legacy import register_legacy_script
        try:
            script = register_legacy_script(db, ROOT, tenant, identifier, user.id)
        except ValueError as exc:
            db.rollback()
            raise HTTPException(422, str(exc)) from exc
    if not script:
        raise HTTPException(404, '请先将训练代码保存为最佳脚本，或选择已有脚本')
    allowed(script.tenant_id, tenants, user)
    try:
        if not script.source_session_id:
            legacy_id = 'script_' + hashlib.md5(script.code.encode()).hexdigest()[:12]
            legacy_file = ROOT / 'tenants' / script.tenant_id / 'scripts' / (legacy_id + '.py')
            if legacy_file.is_file():
                if legacy_file.read_bytes().decode('utf-8') != script.code:
                    raise ValueError('磁盘脚本与数据库版本不同，请先在智训中保存为最佳脚本后再转换')
                from .legacy import register_legacy_script
                script = register_legacy_script(db, ROOT, script.tenant_id, legacy_id, user.id)
        return describe(build(db, script, user.id))
    except ValueError as exc:
        db.rollback()
        raise HTTPException(422, str(exc)) from exc


@router.post('/{identity}/validate')
def validate(identity: str, db: Session = Depends(get_db), user=Depends(get_current_user), tenants=Depends(get_operable_tenants)):
    row = get_row(db, identity, tenants, user)
    from ..utils.subprocess_runner import run_in_fresh_subprocess
    try:
        manifest = verify(artifact(row.digest))
        require_environment(manifest['environment'], ROOT)
        if manifest.get('missing_archive'):
            raise ValueError('归档资料不完整：' + '；'.join(manifest['missing_archive']))
        result = run_in_fresh_subprocess('backend.solutions.worker:validate',
            args=({'_solution': {'digest': row.digest}, 'package': str(artifact(row.digest)), 'tenant_id': row.tenant_id},),
            timeout=1800, max_memory_mb=2048)
        if not result.success:
            raise ValueError(result.error)
        data = result.result
        row.validation = validation_result(data, row.digest)
        db.commit()
        return describe(row)
    except ValueError as exc:
        row.validation = {'digest': row.digest, 'passed': False, 'error': str(exc)}
        db.commit()
        raise HTTPException(422, str(exc)) from exc


@router.post('/{identity}/publish')
def publish(identity: str, db: Session = Depends(get_db), user=Depends(get_current_user), tenants=Depends(get_operable_tenants), body: dict = Body(None)):
    row = get_row(db, identity, tenants, user)
    result = row.validation or {}
    body = body if isinstance(body, dict) else {}
    accepted = (body.get('accept_differences') is True and bool(str(body.get('reason') or '').strip())
                and result.get('completed') and not result.get('formula_warning') and not result.get('missing_sheets'))
    if (not result.get('passed') and not accepted) or result.get('digest') != row.digest:
        raise HTTPException(409, '请先完成本环境验收，结果与训练期望完全一致后再启用')
    try:
        require_environment(verify(artifact(row.digest))['environment'], ROOT)
    except ValueError as exc:
        raise HTTPException(409, str(exc)) from exc
    # Existing versions remain selectable. Aliases resolve the newest published version.
    row.status, row.is_active = 'published', True
    row.validation = {**result, 'accepted_by': user.id, 'accepted_at': datetime.now(timezone.utc).isoformat(),
                      'accepted_differences': bool(accepted and not result.get('passed')), 'acceptance_reason': body.get('reason')}
    db.commit()
    return describe(row)


@router.get('/{identity}/export')
def export(identity: str, db: Session = Depends(get_db), user=Depends(get_current_user), tenants=Depends(get_operable_tenants)):
    row = get_row(db, identity, tenants, user)
    fd, path = tempfile.mkstemp(suffix='.zip')
    os.close(fd)
    try:
        export_zip(artifact(row.digest), path)
        return FileResponse(path, filename=f'solution_{row.digest[:12]}.zip', media_type='application/zip',
                            background=BackgroundTask(os.unlink, path))
    except Exception:
        os.unlink(path)
        raise


@router.post('/{identity}/training')
async def training_workspace(identity: str, db: Session = Depends(get_db),
                             user=Depends(get_current_user), tenants=Depends(get_operable_tenants)):
    row = get_row(db, identity, tenants, user)
    if row.source_session_id:
        session = db.get(TrainingSession, row.source_session_id)
        if session and session.tenant_id == row.tenant_id:
            return describe(row)
    destination = artifact(row.digest)
    restored = None
    try:
        await asyncio.to_thread(verify, destination)
        from .restore import restore, prepare
        restored = await asyncio.to_thread(prepare, destination, row.tenant_id)
        target = db.query(Script).filter_by(id=row.script_id, tenant_id=row.tenant_id).first()
        script, restored = restore(db, destination, row.tenant_id, user.id, prepared=restored,
                                   target_script=target, script_name=row.name)
        row.source_session_id = script.source_session_id
        row.script_id = script.id
        db.commit()
        return describe(row)
    except Exception as exc:
        db.rollback()
        if restored:
            shutil.rmtree(restored, ignore_errors=True)
        raise HTTPException(422, f'恢复修正会话失败：{exc}') from exc


@router.post('/import')
async def import_package(file: UploadFile = File(...), tenant_id: str = Form(...), create_tenant: bool = Form(False),
                         trust_code: bool = Form(False), db: Session = Depends(get_db),
                         user=Depends(get_current_user), tenants=Depends(get_operable_tenants),
                         script_id: str = Form(None), script_name: str = Form(None)):
    # A package contains executable application code, not merely spreadsheet data.
    if not has_permission(user, 'admin') or not trust_code:
        raise HTTPException(403, '方案包含可执行代码，须由管理员确认来源可信后导入')
    try:
        tenant = tenant_name(tenant_id)
    except ValueError as exc:
        raise HTTPException(422, str(exc)) from exc
    identifier = str(script_id).strip() if isinstance(script_id, (str, int)) else ''
    target_id = int(identifier) if identifier.isdigit() else None
    legacy_id = identifier if re.fullmatch(r'script_[0-9a-f]{12}', identifier) else None
    if identifier and target_id is None and legacy_id is None:
        raise HTTPException(422, '目标脚本标识无效')
    new_name = script_name.strip() if isinstance(script_name, str) else ''
    if len(new_name) > 200:
        raise HTTPException(422, '脚本名称不能超过 200 个字符')
    target = None
    if target_id is not None:
        target = db.query(Script).filter_by(id=target_id, tenant_id=tenant).first()
        if not target:
            raise HTTPException(404, '目标脚本不存在或不属于所选租户')
    if identifier and new_name:
        raise HTTPException(422, '请选择现有脚本或填写新脚本名称，不能同时指定')
    exists = (ROOT / 'tenants' / tenant).exists() or db.query(TrainingSession).filter_by(tenant_id=tenant).first() or db.query(Script).filter_by(tenant_id=tenant).first() or db.query(TenantAuthorization).filter_by(tenant_id=tenant).first()
    if create_tenant and exists:
        raise HTTPException(409, '租户已存在，请选择已有租户导入')
    if not create_tenant and not exists:
        raise HTTPException(404, '租户不存在，请选择创建租户')
    if create_tenant and not user.org_id:
        raise HTTPException(422, '请先给当前管理员配置所属组织，以便建立新租户授权')
    STORE.mkdir(parents=True, exist_ok=True)
    restored = None
    with tempfile.TemporaryDirectory(prefix='.import-', dir=STORE) as temporary:
        temp = Path(temporary)
        try:
            size = 0
            try:
                with (temp / 'upload.zip').open('wb') as output:
                    while chunk := await file.read(1024 * 1024):
                        size += len(chunk)
                        if size > 1024 ** 3:
                            raise ValueError('方案 ZIP 不能超过 1GB')
                        await asyncio.to_thread(output.write, chunk)
            finally:
                await file.close()
            manifest = await asyncio.to_thread(extract_zip, temp / 'upload.zip', temp / 'package')
            if legacy_id:
                from .legacy import register_legacy_script
                target = next((s for s in db.query(Script).filter_by(tenant_id=tenant).all()
                               if (s.config or {}).get('solution_import_legacy_id') == legacy_id), None)
                target = target or register_legacy_script(db, ROOT, tenant, legacy_id, user.id)
            # Import stores the package without executing any of its Python.
            name = target.name if target else new_name or manifest['name']
            matching = db.query(Script).filter_by(tenant_id=tenant, name=name).all() if target is None else [target]
            existing = (db.query(CalculationSolution).filter_by(tenant_id=tenant, digest=manifest['digest'])
                        .filter(CalculationSolution.script_id.in_([s.id for s in matching])).first())
            if existing and existing.source_session_id and db.get(TrainingSession, existing.source_session_id):
                return describe(existing)
            if target is None and matching and not existing:
                raise ValueError('脚本名称已存在，请选择该脚本导入，或使用新的脚本名称')
            if existing and target is None:
                target = db.get(Script, existing.script_id)
            destination = artifact(manifest['digest'])
            if destination.exists():
                verify(destination)
            else:
                os.replace(temp / 'package', destination)
            from .restore import restore, prepare
            restored = await asyncio.to_thread(prepare, destination, tenant)
            script, restored = restore(db, destination, tenant, user.id, prepared=restored,
                                       target_script=target, script_name=name)
            if legacy_id:
                script.config = {**(script.config or {}), 'solution_import_legacy_id': legacy_id}
            if create_tenant:
                db.add(TenantAuthorization(tenant_id=tenant, org_id=user.org_id, granted_by=user.id))
            row = existing or CalculationSolution(id=uuid.uuid4().hex, tenant_id=tenant,
                                                  digest=manifest['digest'], created_by=user.id)
            row.script_id = script.id
            row.source_session_id = script.source_session_id
            row.name = script.name
            db.add(row)
            db.commit()
            return describe(row)
        except Exception as exc:
            db.rollback()
            if restored:
                shutil.rmtree(restored, ignore_errors=True)
            raise HTTPException(422, f'导入失败：{exc}') from exc
