"""SOP 调度回归：真实 SQLite 队列，隔离 AI 和业务数据库。"""
import ast
from concurrent.futures import Future
from datetime import datetime, timedelta
from pathlib import Path
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import Mock
import json
import re

from sqlalchemy import Column, DateTime, Integer, String, Text, JSON, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker


class HoldingExecutor:
    def __init__(self):
        self.calls = []

    def submit(self, callback, *args):
        self.calls.append(args)
        return Future()

    def shutdown(self, **kwargs):
        pass


class SopSchedulerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.engine = create_engine("sqlite:///" + str(Path(self.temp.name) / "queue.db"))
        self.addCleanup(self.engine.dispose)
        base = declarative_base()
        class Entry(base):
            __tablename__ = "sop_entries"
            id = Column(Integer, primary_key=True)
            latest_round_id = Column(Integer)
            status = Column(String)
            customer_name = Column(String, default="测试客户")
            description = Column(Text, default="")
        class Round(base):
            __tablename__ = "sop_rounds"
            id = Column(Integer, primary_key=True)
            sop_id = Column(Integer)
            status = Column(String)
            source_file_paths = Column(JSON, default=list)
            source_file_path = Column(String, default="")
            result_file_path = Column(String, default="")
            rule_file_path = Column(String, default="")
            rule_used_id = Column(Integer)
            prompt_text = Column(Text)
            ai_response = Column(Text)
            ai_analysis = Column(JSON)
            ai_comment = Column(Text)
            ai_provider = Column(String)
            finished_at = Column(DateTime)
            error_message = Column(Text)
        root = Path(__file__).parent.parent
        model_tree = ast.parse((root / "database/models.py").read_text(encoding="utf-8"))
        classes = [node for node in model_tree.body if isinstance(node, ast.ClassDef)
                   and node.name in {"SopAiJob", "SopAiSchedule"}]
        self.ns = dict(Base=base, Column=Column, DateTime=DateTime, Integer=Integer,
                       String=String, datetime=datetime)
        exec(compile(ast.Module(body=classes, type_ignores=[]), "queue_models", "exec"), self.ns)
        self.Job, self.Schedule = self.ns["SopAiJob"], self.ns["SopAiSchedule"]
        self.Entry, self.Round = Entry, Round
        base.metadata.create_all(self.engine)
        self.factory = sessionmaker(bind=self.engine, autoflush=False)
        module = ast.parse((root / "utils/sop_scheduler.py").read_text(encoding="utf-8"))
        module.body = [node for node in module.body if not (isinstance(node, ast.ImportFrom) and node.level)]
        self.ns.update(SessionLocal=self.factory, SopEntry=Entry, SopRound=Round, __name__="test_scheduler")
        exec(compile(module, "sop_scheduler", "exec"), self.ns)
        self.first = self.dispatcher()
        self.first.initialize()

    def dispatcher(self):
        instance = self.ns["SopAiDispatcher"](self.factory)
        instance.executor.shutdown()
        instance.executor = HoldingExecutor()
        return instance

    def add(self, identifier, status="ai_queued", with_job=True, job_status="ready"):
        with self.factory() as db:
            db.add(self.Entry(id=identifier, latest_round_id=identifier, status=status))
            db.add(self.Round(id=identifier, sop_id=identifier, status=status))
            if with_job:
                db.add(self.Job(round_id=identifier, entry_id=identifier, status=job_status))
            db.commit()

    def scheduled(self, due=False):
        with self.factory() as db:
            db.query(self.Schedule).update({"mode":"scheduled", "interval_hours":6,
                "next_run_at": datetime.utcnow() + timedelta(hours=-1 if due else 6)})
            db.commit()

    def test_immediate_upload_returns_verdict_without_dispatcher_tick(self):
        for identifier, fails in [(1, False), (2, True)]:
            self.add(identifier, job_status="pending")
            provider = SimpleNamespace(model="fake", chat=Mock(return_value='{"passed":true}'))
            if fails:
                provider.chat.side_effect = RuntimeError("模拟 AI 服务失败")
            result = self.ns["run_uploaded_round"](identifier, identifier, self.factory, self.analyzer(provider))
            self.assertEqual(result["status"], "failed" if fails else "ai_passed")
            self.assertTrue(result["error_message"] if fails else result["ai_comment"])
            self.assertEqual(self.first.executor.calls, [])
            again = self.ns["run_uploaded_round"](identifier, identifier, self.factory, self.analyzer(provider))
            self.assertEqual(again, result)
            self.assertEqual(provider.chat.call_count, 1)

    def test_immediate_tick_does_not_scan_pending_uploads(self):
        self.add(1, job_status="pending")
        with self.factory() as db:
            before = db.query(self.Schedule).one().next_run_at
        self.first.tick()
        self.assertEqual(self.first.executor.calls, [])
        with self.factory() as db:
            self.assertEqual(db.query(self.Job).one().status, "pending")
            self.assertEqual(db.query(self.Schedule).one().next_run_at, before)
            self.assertIsNone(db.query(self.Schedule).one().last_run_at)

    def test_scheduled_upload_returns_queue_without_ai_call(self):
        self.scheduled()
        self.add(1, job_status="pending")
        analyze = Mock()
        result = self.ns["run_uploaded_round"](1, 1, self.factory, analyze)
        self.assertEqual(result["status"], "ai_queued")
        analyze.assert_not_called()

    def test_startup_recovers_interrupted_immediate_upload(self):
        self.add(1, job_status="pending")
        self.first.initialize()
        self.first.tick()
        self.assertEqual(len(self.first.executor.calls), 1)

    def test_immediate_scan_claims_once_across_workers(self):
        self.add(1)
        second = self.dispatcher()
        self.first.tick()
        second.tick()
        self.first.tick()
        self.assertEqual(len(self.first.executor.calls), 1)
        self.assertEqual(len(second.executor.calls), 0)
        with self.factory() as db:
            self.assertEqual(db.query(self.Round).one().status, "ai_analyzing")

    def test_scheduled_waits_then_processes_only_unreviewed(self):
        self.scheduled()
        self.add(1, job_status="pending")
        self.add(2, "ai_passed", with_job=False)
        self.first.tick()
        self.assertEqual(self.first.executor.calls, [])
        self.scheduled(due=True)
        before = datetime.utcnow()
        self.first.tick()
        self.assertEqual(len(self.first.executor.calls), 1)
        with self.factory() as db:
            config = db.query(self.Schedule).one()
            self.assertGreater(config.next_run_at, before + timedelta(hours=5, minutes=59))
            self.assertEqual(db.query(self.Round).filter_by(id=2).one().status, "ai_passed")
        # 本次扫描后上传的记录等待下一次定时扫描。
        self.add(3, job_status="pending")
        self.first.tick()
        self.assertEqual(len(self.first.executor.calls), 1)

    def test_restart_preserves_waiting_queue_and_recovers_expired_claim(self):
        self.scheduled()
        self.add(1, job_status="pending")
        restarted = self.dispatcher()
        restarted.initialize()
        restarted.tick()
        self.assertEqual(restarted.executor.calls, [])
        self.scheduled(due=True)
        self.first.tick()
        round_id, entry_id, job_id, old_token = self.first.executor.calls[0]
        with self.factory() as db:
            db.query(self.Job).filter_by(id=job_id).update({"lease_until": datetime.utcnow()-timedelta(seconds=1)})
            db.commit()
        restarted.tick()
        self.assertEqual(len(restarted.executor.calls), 1)
        new_token = restarted.executor.calls[0][3]
        self.assertNotEqual(old_token, new_token)
        with self.factory() as db:
            self.assertFalse(self.ns["lock_job"](db, job_id, old_token))
            self.assertTrue(self.ns["lock_job"](db, job_id, new_token))
            db.commit()

    def test_legacy_interrupted_round_is_enqueued_only_once(self):
        self.add(1, "ai_analyzing", with_job=False)
        self.first.initialize()
        self.first.initialize()
        with self.factory() as db:
            self.assertEqual(db.query(self.Job).count(), 1)
            self.assertEqual(db.query(self.Round).one().status, "ai_queued")

    def test_terminal_jobs_and_obsolete_rounds_are_not_reprocessed(self):
        self.add(1, "completed")
        self.add(2)
        with self.factory() as db:
            db.query(self.Entry).filter_by(id=2).update({"latest_round_id":999})
            db.query(self.Job).filter_by(round_id=1).update({"status":"completed"})
            db.commit()
        self.first.tick()
        self.assertEqual(self.first.executor.calls, [])

    def analyzer(self, provider):
        tree = ast.parse(Path(__file__).parent.parent.joinpath("api/tools.py").read_text(encoding="utf-8"))
        node = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_run_sop_ai_analysis")
        node.body = [statement for statement in node.body if not isinstance(statement, ast.ImportFrom)]
        verdict = SimpleNamespace(passed=True, summary="通过", issues=[], model_dump=lambda: {"passed":True})
        context = dict(self.ns, AIProviderFactory=SimpleNamespace(create_with_fallback=lambda: provider),
                       SopAiVerdict=SimpleNamespace(model_validate=lambda value: verdict),
                       SopRuleFile=self.Job, _sop_effective_rule=lambda rows, customer: None,
                       _sop_read_file=lambda *args, **kwargs: "测试文件", _build_sop_prompt=lambda *args: "测试提示",
                       SOP_SYSTEM_PROMPT="测试", _sop_notify_ai_pass=Mock(), _sop_notify_rejection=Mock(),
                       re=re, json=json)
        exec(compile(ast.Module(body=[node], type_ignores=[]), "sop_analyzer", "exec"), context)
        return context["_run_sop_ai_analysis"]

    def test_ai_result_and_failure_finish_the_claimed_job(self):
        for identifier, fails in [(1, False), (2, True)]:
            self.add(identifier)
            with self.factory() as db:
                db.query(self.Schedule).update({"next_run_at":datetime.utcnow()-timedelta(seconds=1)})
                db.commit()
            self.first.tick()
            provider = SimpleNamespace(model="fake", chat=Mock(return_value='{"passed":true}'))
            if fails:
                provider.chat.side_effect = RuntimeError("模拟 AI 服务失败")
            self.analyzer(provider)(*self.first.executor.calls[-1])
            with self.factory() as db:
                self.assertEqual(db.query(self.Job).filter_by(round_id=identifier).one().status, "failed" if fails else "completed")
                self.assertEqual(db.query(self.Round).filter_by(id=identifier).one().status, "failed" if fails else "ai_passed")
                self.assertEqual(db.query(self.Entry).filter_by(id=identifier).one().status, "failed" if fails else "ai_passed")

    def test_ai_response_cannot_overwrite_new_owner(self):
        self.add(1)
        self.first.tick()
        round_id, entry_id, job_id, token = self.first.executor.calls[0]
        def response(_messages):
            with self.factory() as db:
                db.query(self.Job).filter_by(id=job_id).update({"claim_token":"new-owner"})
                db.commit()
            return '{"passed":true}'
        self.analyzer(SimpleNamespace(model="fake", chat=response))(round_id, entry_id, job_id, token)
        with self.factory() as db:
            self.assertEqual(db.query(self.Job).one().status, "running")
            self.assertEqual(db.query(self.Round).one().status, "ai_analyzing")
            self.assertIsNone(db.query(self.Round).one().ai_analysis)


if __name__ == "__main__":
    unittest.main()
