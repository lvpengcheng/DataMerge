"""SOP 专用初审队列：持久化间隔调度、原子认领、租约恢复。"""
import logging
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

from sqlalchemy.exc import IntegrityError
from ..database.connection import SessionLocal
from ..database.models import SopAiJob, SopAiSchedule, SopEntry, SopRound

logger = logging.getLogger(__name__)
POLL_SECONDS = 5
LEASE_SECONDS = 180


def ensure_schedule(db):
    row = db.query(SopAiSchedule).filter_by(id=1).first()
    if row is None:
        try:
            db.add(SopAiSchedule(id=1, mode="immediate", interval_hours=6,
                                 next_run_at=datetime.utcnow()))
            db.commit()
        except IntegrityError:
            db.rollback()  # 另一进程已建立同一配置。
        row = db.query(SopAiSchedule).filter_by(id=1).first()
    return row


def lock_job(db, job_id, token):
    """在结果写入事务中锁定有效租约；失去认领权的旧进程不能覆盖结果。"""
    now = datetime.utcnow()
    return db.query(SopAiJob).filter(
        SopAiJob.id == job_id, SopAiJob.claim_token == token,
        SopAiJob.status == "running", SopAiJob.lease_until > now,
    ).update({"lease_until": now + timedelta(seconds=LEASE_SECONDS)}, synchronize_session=False) == 1


def run_uploaded_round(round_id, entry_id, session_factory=SessionLocal, analyze=None):
    """立即模式由上传请求直接执行并返回结果；定时模式仅保存任务。"""
    with session_factory() as db:
        config = ensure_schedule(db)
        if config.mode == "scheduled":
            return {"status": "ai_queued", "ai_comment": "", "error_message": ""}
        job = db.query(SopAiJob).filter_by(round_id=round_id, entry_id=entry_id).one()
        job_id = job.id
        token = uuid.uuid4().hex
        claimed = db.query(SopAiJob).filter(
            SopAiJob.id == job_id, SopAiJob.status.in_(("pending", "ready")),
        ).update({"status": "running", "claim_token": token,
                  "lease_until": datetime.utcnow() + timedelta(seconds=LEASE_SECONDS)}, synchronize_session=False)
        if claimed:
            db.query(SopRound).filter_by(id=round_id).update({"status": "ai_analyzing"})
            db.query(SopEntry).filter_by(id=entry_id, latest_round_id=round_id).update({"status": "ai_analyzing"})
        db.commit()
    if claimed:
        done = threading.Event()
        def renew_lease():
            while not done.wait(LEASE_SECONDS / 3):
                try:
                    with session_factory() as db:
                        valid = lock_job(db, job_id, token)
                        db.commit()
                        if not valid:
                            return
                except Exception:
                    logger.exception("[sop] 上传初审租约续期失败 job=%s", job_id)
        heartbeat = threading.Thread(target=renew_lease, name="sop-upload-lease", daemon=True)
        heartbeat.start()
        try:
            if analyze is None:
                from ..api.tools import _run_sop_ai_analysis
                analyze = _run_sop_ai_analysis
            analyze(round_id, entry_id, job_id, token)
        finally:
            done.set()
            heartbeat.join(timeout=2)
    with session_factory() as db:
        row = db.query(SopRound).filter_by(id=round_id).one()
        return {"status": row.status, "ai_comment": row.ai_comment or "",
                "error_message": row.error_message or ""}


class SopAiDispatcher:
    def __init__(self, session_factory=SessionLocal, analyze=None):
        self.session_factory = session_factory
        self.analyze = analyze
        self.stop_event = threading.Event()
        self.wake_event = threading.Event()
        self.thread = None
        self.executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="sop-ai")
        self.active = {}

    def initialize(self):
        with self.session_factory() as db:
            ensure_schedule(db)
            # 兼容升级前尚未完成的初审，只补建未存在的任务。
            rounds = db.query(SopRound).join(SopEntry, SopEntry.latest_round_id == SopRound.id).filter(
                SopRound.status.in_(("ai_queued", "ai_analyzing")),
                ~SopRound.id.in_(db.query(SopAiJob.round_id)),
            ).all()
            for row in rounds:
                try:
                    with db.begin_nested():
                        db.add(SopAiJob(round_id=row.id, entry_id=row.sop_id, status="pending"))
                        db.flush()
                        row.status = "ai_queued"
                        db.query(SopEntry).filter_by(id=row.sop_id, latest_round_id=row.id).update(
                            {"status": "ai_queued"}, synchronize_session=False)
                except IntegrityError:
                    pass
            if ensure_schedule(db).mode == "immediate":
                # 仅启动时恢复上传中断后尚未认领的记录；正常上传不依赖扫描。
                db.query(SopAiJob).filter_by(status="pending").update({"status": "ready"})
            db.commit()

    def tick(self):
        now = datetime.utcnow()
        with self.session_factory() as db:
            for future, (job_id, token) in list(self.active.items()):
                if future.done():
                    try:
                        future.result()
                    except Exception:
                        logger.exception("[sop] 初审工作线程异常 job=%s", job_id)
                    del self.active[future]
                else:
                    # 过期租约不续约，交由恢复流程处理。
                    lock_job(db, job_id, token)
            db.commit()
            expired = db.query(SopAiJob).filter(
                SopAiJob.status == "running", SopAiJob.lease_until <= now).all()
            for job in expired:
                changed = db.query(SopAiJob).filter(
                    SopAiJob.id == job.id, SopAiJob.status == "running", SopAiJob.lease_until <= now,
                ).update({"status": "ready", "claim_token": None, "lease_until": None}, synchronize_session=False)
                if changed:
                    db.query(SopRound).filter_by(id=job.round_id, status="ai_analyzing").update({"status": "ai_queued"})
                    db.query(SopEntry).filter_by(id=job.entry_id, latest_round_id=job.round_id,
                                                status="ai_analyzing").update({"status": "ai_queued"})
            db.commit()
            config = ensure_schedule(db)
            if config.mode == "scheduled" and config.next_run_at <= now:
                interval = timedelta(hours=config.interval_hours)
                changed = db.query(SopAiSchedule).filter(
                    SopAiSchedule.id == 1, SopAiSchedule.next_run_at == config.next_run_at,
                    SopAiSchedule.updated_at == config.updated_at,
                ).update({"last_run_at": now, "next_run_at": now + interval}, synchronize_session=False)
                if changed:
                    db.query(SopAiJob).filter_by(status="pending").update({"status": "ready"})
                    logger.info("[sop] 扫描初审队列 mode=%s", config.mode)
                db.commit()
            capacity = max(0, 2 - len(self.active))
            candidates = db.query(SopAiJob).join(SopRound, SopRound.id == SopAiJob.round_id).join(
                SopEntry, SopEntry.id == SopAiJob.entry_id).filter(
                    SopAiJob.status == "ready", SopRound.status == "ai_queued",
                    SopEntry.latest_round_id == SopRound.id,
                ).order_by(SopAiJob.id).limit(capacity).all() if capacity else []
            for candidate in candidates:
                job_id, round_id, entry_id = candidate.id, candidate.round_id, candidate.entry_id
                token = uuid.uuid4().hex
                claimed = db.query(SopAiJob).filter_by(id=job_id, status="ready").update(
                    {"status": "running", "claim_token": token,
                     "lease_until": now + timedelta(seconds=LEASE_SECONDS)}, synchronize_session=False)
                if not claimed:
                    db.rollback()
                    continue
                db.query(SopRound).filter_by(id=round_id).update({"status": "ai_analyzing"})
                db.query(SopEntry).filter_by(id=entry_id, latest_round_id=round_id).update({"status": "ai_analyzing"})
                db.commit()
                try:
                    future = self.executor.submit(self._analyze, round_id, entry_id, job_id, token)
                    self.active[future] = (job_id, token)
                except Exception:
                    # 保留持久化记录，租约过期后自动恢复。
                    logger.exception("[sop] 无法启动初审任务 job=%s", job_id)

    def _analyze(self, round_id, entry_id, job_id, token):
        if self.analyze:
            return self.analyze(round_id, entry_id, job_id, token)
        from ..api.tools import _run_sop_ai_analysis
        return _run_sop_ai_analysis(round_id, entry_id, job_id, token)

    def _loop(self):
        while not self.stop_event.is_set():
            try:
                self.tick()
            except Exception:
                logger.exception("[sop] 初审队列扫描失败，将于下次扫描重试")
            self.wake_event.wait(POLL_SECONDS)
            self.wake_event.clear()

    def start(self):
        self.initialize()
        self.thread = threading.Thread(target=self._loop, name="sop-ai-dispatcher", daemon=True)
        self.thread.start()

    def stop(self):
        self.stop_event.set()
        self.wake_event.set()
        if self.thread:
            self.thread.join(timeout=2)
        self.executor.shutdown(wait=False)


_dispatcher = None


def start_sop_dispatcher():
    global _dispatcher
    if _dispatcher is None:
        _dispatcher = SopAiDispatcher()
        _dispatcher.start()


def stop_sop_dispatcher():
    global _dispatcher
    if _dispatcher:
        _dispatcher.stop()
        _dispatcher = None


def wake_sop_dispatcher():
    if _dispatcher:
        _dispatcher.wake_event.set()
