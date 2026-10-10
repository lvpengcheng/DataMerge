"""SOP 数据范围回归测试；只提取纯权限函数，避免依赖完整服务运行环境。"""

import ast
import asyncio
import types
import unittest
from unittest.mock import Mock
from pathlib import Path


class HTTPException(Exception):
    def __init__(self, status_code, detail):
        self.status_code = status_code
        self.detail = detail


source = (Path(__file__).parent.parent / "api").joinpath("tools.py").read_text(encoding="utf-8")
tree = ast.parse(source)
names = {"_sop_is_superadmin", "_sop_is_owner", "_sop_can_read", "_sop_check_read", "_sop_check_upload", "_require_sop_action", "_sop_create_round", "_sop_reviewer_issue", "_sop_reviewer_eligible", "_sop_get_reviewer", "_sop_notify_ai_pass", "sop_reviewer_config"}
names.add("_sop_get_reviewers")
names.update({"_sop_current_rules", "_sop_effective_rule"})
functions = [node for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in names]
for function in functions:
    function.decorator_list = []
namespace = {
    "HTTPException": HTTPException,
    "Path": Path,
    "SopEntry": object,
    "SopRound": object,
    "Depends": lambda dependency: None,
    "get_current_user": lambda: None,
    "require_permission": lambda key: None,
    "has_permission": lambda user, key: key in user.role.permissions or user.role.name == "admin",
}
exec(compile(ast.Module(body=functions, type_ignores=[]), "sop_access", "exec"), namespace)


def user(identifier, *permissions, role="custom"):
    return types.SimpleNamespace(id=identifier, role=types.SimpleNamespace(
        name=role, permissions={key: True for key in permissions}))


def entry(owner, status="draft", creator=None):
    return types.SimpleNamespace(assignee_user_id=owner, created_by=creator, status=status)


class SopAccessTests(unittest.TestCase):
    def test_only_latest_rules_are_visible_and_effective(self):
        def rule(identifier, scope="global", customer=None, active=True, path="rule.md"):
            return types.SimpleNamespace(id=identifier, scope=scope, customer_name=customer,
                                         is_active=active, file_path=path)
        old, latest = rule(1), rule(4)
        customer_old, customer_latest = rule(2, "customer", "客户"), rule(5, "customer", "客户")
        rows = [latest, customer_old, old, customer_latest]
        self.assertEqual({row.id for row in namespace["_sop_current_rules"](rows)}, {4, 5})
        effective = namespace["_sop_effective_rule"]
        self.assertEqual(effective(rows, "客户").id, 5)
        self.assertEqual(effective(rows, "其他客户").id, 4)
        customer_latest.is_active = False
        self.assertEqual(effective(rows, "客户").id, 4)
        latest.is_active = False
        self.assertIsNone(effective(rows, "客户"))
        latest.file_path = customer_latest.file_path = ""
        self.assertEqual(namespace["_sop_current_rules"](rows), [])
        self.assertIsNone(effective(rows, "客户"))

    def test_reviewer_options_query_user_table_and_explain_exclusions(self):
        allowed = user(42, "menu.sop", "sop.review")
        allowed.is_active, allowed.email = True, "reviewer@example.com"
        allowed.username, allowed.display_name = "reviewer", "复审人"
        inactive = user(43)
        inactive.is_active, inactive.email = False, None
        inactive.username, inactive.display_name = "inactive", "停用用户"
        user_model = types.SimpleNamespace(username=object())
        config_model = object()
        config_query, user_query, db = Mock(), Mock(), Mock()
        config_query.filter_by.return_value.first.return_value = types.SimpleNamespace(value={"user_id": 42})
        user_query.filter_by.return_value.first.return_value = allowed
        user_query.order_by.return_value.all.return_value = [allowed, inactive]
        db.query.side_effect = lambda model: config_query if model is config_model else user_query
        namespace.update(SystemConfig=config_model, User=user_model, SessionLocal=lambda: db)
        data = asyncio.run(namespace["sop_reviewer_config"]())
        self.assertEqual(data["user_id"], 42)
        self.assertTrue(data["reviewer_valid"])
        self.assertEqual([row["id"] for row in data["users"]], [42, 43])
        self.assertTrue(data["users"][0]["selectable"])
        self.assertFalse(data["users"][1]["selectable"])
        self.assertIn("账号已停用", data["users"][1]["disabled_reason"])
        self.assertIn("未填写邮箱", data["users"][1]["disabled_reason"])
        user_query.order_by.assert_called_once_with(user_model.username)
        db.close.assert_called_once()

    def test_reviewer_requires_active_email_and_both_permissions(self):
        reviewer = user(6, "menu.sop", "sop.review")
        reviewer.is_active = True
        reviewer.email = "reviewer@example.com"
        eligible = namespace["_sop_reviewer_eligible"]
        self.assertTrue(eligible(reviewer))
        reviewer.is_active = False
        self.assertFalse(eligible(reviewer))
        reviewer.is_active = True
        reviewer.email = "   "
        self.assertFalse(eligible(reviewer))
        reviewer.email = "reviewer@example.com"
        del reviewer.role.permissions["sop.review"]
        self.assertFalse(eligible(reviewer))
        reviewer.role.permissions = {"sop.review": True}
        self.assertFalse(eligible(reviewer))
        self.assertFalse(eligible(None))

    def test_notice_uses_stored_user_id_and_current_email(self):
        reviewer = user(42, "menu.sop", "sop.review")
        reviewer.is_active = True
        reviewer.email = " selected@example.com "
        reviewer.username = "selected.reviewer"
        reviewer.display_name = "所选复审人"
        config_query, user_query = Mock(), Mock()
        config_query.filter_by.return_value.first.return_value = types.SimpleNamespace(value={"user_id": 42})
        user_query.filter_by.return_value.first.return_value = reviewer
        config_model, user_model = object(), object()
        namespace.update(SystemConfig=config_model, User=user_model,
                         _sop_executor=Mock(), send_sop_mail=Mock(),
                         _sop_public_url=lambda path: "https://example.com" + path)
        db = Mock()
        db.query.side_effect = lambda model: config_query if model is config_model else user_query
        namespace["_sop_notify_ai_pass"](db, types.SimpleNamespace(salesperson_name="业务员", customer_name="客户"))
        user_query.filter_by.assert_called_once_with(id=42)
        args = namespace["_sop_executor"].submit.call_args.args
        self.assertEqual(args[1], "selected@example.com")
        self.assertIn("您好 所选复审人", args[3])
        reviewer.is_active = False
        namespace["_sop_executor"].reset_mock()
        with self.assertRaises(HTTPException):
            namespace["_sop_notify_ai_pass"](db, types.SimpleNamespace())
        namespace["_sop_executor"].submit.assert_not_called()

    def test_missing_config_does_not_fall_back_to_a_username(self):
        config_model = object()
        namespace["SystemConfig"] = config_model
        db = Mock()
        db.query.return_value.filter_by.return_value.first.return_value = None
        self.assertIsNone(namespace["_sop_get_reviewer"](db))
        db.query.assert_called_once_with(config_model)

    def test_multiple_reviewers_receive_notices_with_duplicate_emails_deduplicated(self):
        people = []
        for identifier, email, active in [(1, "one@example.com", True), (2, "two@example.com", True),
                                           (3, "ONE@example.com", True), (4, "inactive@example.com", False)]:
            person = user(identifier, "menu.sop", "sop.review")
            person.email, person.is_active = email, active
            person.display_name = person.username = str(identifier)
            people.append(person)
        config_model, user_model = object(), object()
        db, config_query, user_query = Mock(), Mock(), Mock()
        config_query.filter_by.return_value.first.return_value = types.SimpleNamespace(value={"user_ids": [1, 2, 3, 4]})
        user_query.filter_by.side_effect = lambda id: types.SimpleNamespace(first=lambda: people[id - 1])
        db.query.side_effect = lambda model: config_query if model is config_model else user_query
        namespace.update(SystemConfig=config_model, User=user_model, _sop_executor=Mock(),
                         send_sop_mail=Mock(), _sop_public_url=lambda path: None)
        namespace["_sop_notify_ai_pass"](db, types.SimpleNamespace(salesperson_name="业务员", customer_name="客户"))
        recipients = [call.args[1] for call in namespace["_sop_executor"].submit.call_args_list]
        self.assertEqual(recipients, ["one@example.com", "two@example.com"])

    def test_business_user_sees_only_assigned_entries(self):
        seller = user(1, "menu.sop", "sop.upload")
        self.assertTrue(namespace["_sop_can_read"](entry(1), seller))
        self.assertFalse(namespace["_sop_can_read"](entry(2), seller))
        with self.assertRaises(HTTPException) as raised:
            namespace["_sop_check_upload"](entry(2), seller)
        self.assertEqual(raised.exception.status_code, 403)

    def test_reviewer_sees_pending_and_submitted_but_cannot_upload(self):
        reviewer = user(3, "menu.sop", "sop.review")
        self.assertTrue(namespace["_sop_can_read"](entry(1, "draft"), reviewer))
        self.assertTrue(namespace["_sop_can_read"](entry(1, "awaiting_files"), reviewer))
        self.assertTrue(namespace["_sop_can_read"](entry(1, "ai_passed"), reviewer))
        with self.assertRaises(HTTPException):
            namespace["_sop_check_upload"](entry(1, "rejected"), reviewer)

    def test_report_role_reads_all_but_cannot_upload(self):
        manager = user(4, "menu.sop", "sop.report")
        self.assertTrue(namespace["_sop_can_read"](entry(1), manager))
        with self.assertRaises(HTTPException):
            namespace["_sop_check_upload"](entry(1), manager)

    def test_legacy_creator_can_access_unassigned_entry(self):
        seller = user(1, "menu.sop", "sop.upload")
        self.assertTrue(namespace["_sop_can_read"](entry(None, creator=1), seller))
        self.assertFalse(namespace["_sop_can_read"](entry(None, creator=2), seller))

    def test_superadmin_can_access_all(self):
        admin = user(5, role="admin")
        self.assertTrue(namespace["_sop_can_read"](entry(1), admin))
        namespace["_sop_check_upload"](entry(1), admin)

    def test_action_also_requires_sop_menu(self):
        check = namespace["_require_sop_action"]("sop.upload")
        with self.assertRaises(HTTPException):
            asyncio.run(check(user(1, "sop.upload")))
        self.assertEqual(asyncio.run(check(user(1, "menu.sop", "sop.upload"))).id, 1)

    def test_upload_requires_rule_file_and_open_status(self):
        upload = namespace["_sop_create_round"]
        with self.assertRaises(HTTPException) as closed:
            upload(None, entry(1, "completed"), [], None, None)
        self.assertEqual(closed.exception.status_code, 400)
        source = types.SimpleNamespace(filename="source.xlsx")
        result = types.SimpleNamespace(filename="result.xlsx")
        with self.assertRaises(HTTPException) as missing_rule:
            upload(None, entry(1, "draft"), [source], result, None)
        self.assertIn("规则文件", missing_rule.exception.detail)

    def test_upload_rejects_disallowed_formats_before_creating_files_or_rounds(self):
        for slot in range(3):
            for extension in ("xlsm", "pdf", "docx", "txt", "zip"):
                files = [types.SimpleNamespace(filename="file.xlsx") for _ in range(3)]
                files[slot].filename = "file." + extension
                with self.subTest(slot=slot, extension=extension), self.assertRaises(HTTPException) as rejected:
                    namespace["_sop_create_round"](None, entry(1), [files[0]], files[1], files[2])
                self.assertEqual(rejected.exception.status_code, 400)


if __name__ == "__main__":
    unittest.main()
