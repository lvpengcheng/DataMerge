"""Host side routing: pin a solution before upload and never re-resolve mid-task."""
import hashlib
import json
from pathlib import Path
from .store import ROOT, artifact, materialize
from .package import verify
from .environment import require_environment


def binding(tenant, script_id, db=None):
    from ..database.connection import SessionLocal
    from ..database.models import CalculationSolution, Script
    owned = db is None
    db = db or SessionLocal()
    try:
        q = db.query(CalculationSolution).filter_by(tenant_id=tenant, status='published', is_active=True)
        if str(script_id).startswith('solution_'):
            row = q.filter_by(id=str(script_id)[9:]).first()
            if not row:
                raise ValueError('方案不存在、已停用或无权访问')
        else:
            row = None
            for candidate in q.order_by(CalculationSolution.created_at.desc()).all():
                script = db.query(Script).filter_by(id=candidate.script_id, tenant_id=tenant).first()
                aliases = (str(script.id), script.name, 'script_' + hashlib.md5(script.code.encode()).hexdigest()[:12]) if script else ()
                if str(script_id) in aliases:
                    row = candidate
                    break
        if not row:
            return None
        root = artifact(row.digest)
        manifest = verify(root)
        require_environment(manifest['environment'], ROOT)
        return {'id': row.id, 'digest': row.digest, 'tenant_id': tenant, 'script_id': row.script_id,
                'metadata': materialize(manifest['metadata'], root),
                'code': (root / 'business/script.py').read_bytes().decode('utf-8')}
    finally:
        if owned:
            db.close()


def pin_run(run_dir, solution):
    if solution:
        from .inputs import snapshot
        from ..database.connection import SessionLocal
        inputs_dir = Path(run_dir) / '_solution_inputs'
        with SessionLocal() as db:
            solution['inputs_digest'] = snapshot(ROOT, solution['tenant_id'], inputs_dir, db)
        solution['inputs_dir'] = str(inputs_dir.resolve())
        path = Path(run_dir) / 'solution.json'
        identity = {k: solution[k] for k in ('id', 'digest', 'tenant_id', 'script_id')}
        identity['inputs_digest'] = solution['inputs_digest']
        path.write_text(json.dumps(identity, ensure_ascii=False), encoding='utf-8')
        from ..database.models import ComputeTask, Script
        with SessionLocal() as db:
            task = db.query(ComputeTask).filter_by(id=int(Path(run_dir).name.removeprefix('task_'))).first()
            if task:
                if db.query(Script).filter_by(id=solution['script_id'], tenant_id=solution['tenant_id']).first():
                    task.script_id = solution['script_id']
                task.analysis_report = {**(task.analysis_report or {}), 'solution': identity}
                db.commit()


def launch(params_file, solution, mode):
    """Only local verified artifacts can be launched. Imported code needs admin approval."""
    if not solution:
        return None
    root = artifact(solution['digest'])
    manifest = verify(root)
    from .environment import runtime_python
    executable = runtime_python(manifest['environment'], ROOT)
    return [executable, '-I', '-X', 'utf8', '-u', str(root / 'runtime/backend/solutions/worker.py'),
            str(root), mode, str(params_file), str(ROOT)]


def solution_for_params(params):
    return params.get('_solution') if isinstance(params, dict) else None
