"""One code selection rule for training discussion, editing and downloads."""
from types import SimpleNamespace


def latest_code_revision(db, session):
    from backend.database.models import TrainingIteration, Script, CalculationSolution
    iteration = (db.query(TrainingIteration).filter_by(session_id=session.id)
                 .filter(TrainingIteration.generated_code.isnot(None))
                 .filter(TrainingIteration.generated_code != '')
                 .order_by(TrainingIteration.iteration_num.desc()).first())
    if iteration and (iteration.generated_code or '').strip():
        return iteration
    script = (db.query(Script).filter_by(id=session.final_script_id, tenant_id=session.tenant_id).first()
              if session.final_script_id else None)
    if not script:
        solution = (db.query(CalculationSolution).filter_by(source_session_id=session.id, tenant_id=session.tenant_id)
                    .order_by(CalculationSolution.created_at.desc()).first())
        if solution:
            script = db.query(Script).filter_by(id=solution.script_id, tenant_id=session.tenant_id).first()
    if not script:
        # Legacy migration imports did not always restore final_script_id, but
        # their saved script still points back to the destination session.
        script = (db.query(Script).filter_by(source_session_id=session.id, tenant_id=session.tenant_id)
                  .filter(Script.code.isnot(None), Script.code != '')
                  .order_by(Script.created_at.desc(), Script.id.desc()).first())
    # Older imported sessions may have no iteration code at all. Their saved
    # script is still a valid editing starting point, even before publication.
    if not script or not (script.code or '').strip():
        return None
    return SimpleNamespace(id=None, generated_code=script.code, accuracy=None,
                           iteration_num=session.total_iterations or 0, error_details=None,
                           execution_result={'source': 'saved_script', 'validation_pending': True})
