"""使用真实临时文件验证 SOP ZIP；不连接业务数据库。"""
import ast
import io
from pathlib import Path
import tempfile
import types
import unittest
import zipfile


class HTTPException(Exception):
    def __init__(self, status_code, detail):
        self.status_code, self.detail = status_code, detail


source = Path(__file__).parent.parent.joinpath("api/tools.py").read_text(encoding="utf-8")
nodes = [node for node in ast.parse(source).body if isinstance(node, ast.FunctionDef)
         and node.name in {"_sop_package_files", "_sop_build_package", "_sop_round_dir"}]
namespace = dict(Path=Path, io=io, zipfile=zipfile, HTTPException=HTTPException)
exec(compile(ast.Module(body=nodes, type_ignores=[]), "sop_package", "exec"), namespace)


class SopPackageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        namespace["SOP_ROOT"] = Path(self.temp.name)
        self.folder = namespace["_sop_round_dir"](7, 2)
        self.folder.mkdir(parents=True)
        self.paths = []
        for name, content in [("source_0.xlsx", b"source one\x00\xff"),
                              ("source_1.xlsx", b"source two"), ("result.xlsx", b"result"),
                              ("rule.docx", b"rule")]:
            path = self.folder / name
            path.write_bytes(content)
            self.paths.append(str(path))
        self.round = types.SimpleNamespace(
            sop_id=7, round_no=2, source_file_path=self.paths[0], source_file_name="工资.xlsx",
            source_file_paths=self.paths[:2], source_file_names=["工资.xlsx", "工资.xlsx"],
            result_file_path=self.paths[2], result_file_name="结果.xlsx",
            rule_file_path=self.paths[3], rule_file_name="规则.docx")

    def test_all_files_and_duplicate_original_names_are_preserved(self):
        with zipfile.ZipFile(namespace["_sop_build_package"](self.round)) as archive:
            self.assertEqual(archive.namelist(), ["源文件/1_工资.xlsx", "源文件/2_工资.xlsx", "结果表/结果.xlsx", "规则/规则.docx"])
            self.assertEqual(archive.read("源文件/1_工资.xlsx"), b"source one\x00\xff")
            self.assertEqual(archive.read("源文件/2_工资.xlsx"), b"source two")
            self.assertIsNone(archive.testzip())

    def test_migrated_linux_and_windows_paths_resolve_in_same_round(self):
        self.round.source_file_path = "/app/old/source_0.xlsx"
        self.round.source_file_paths = ["/app/old/source_0.xlsx", r"Z:\old\source_1.xlsx"]
        self.round.result_file_path = "/app/old/result.xlsx"
        self.round.rule_file_path = r"Z:\old\rule.docx"
        with zipfile.ZipFile(namespace["_sop_build_package"](self.round)) as archive:
            self.assertEqual(len(archive.namelist()), 4)

    def test_legacy_source_not_in_array_is_included(self):
        self.round.source_file_paths = [self.paths[1]]
        self.round.source_file_names = ["工资.xlsx"]
        with zipfile.ZipFile(namespace["_sop_build_package"](self.round)) as archive:
            self.assertEqual(len(archive.namelist()), 4)

    def test_missing_files_fail_instead_of_empty_or_partial_zip(self):
        Path(self.paths[1]).unlink()
        with self.assertRaises(HTTPException) as raised:
            namespace["_sop_build_package"](self.round)
        self.assertEqual(raised.exception.status_code, 409)
        self.assertIn("源文件/工资.xlsx", raised.exception.detail)
        for path in self.folder.iterdir():
            path.unlink()
        with self.assertRaises(HTTPException) as empty:
            namespace["_sop_build_package"](self.round)
        self.assertIn("结果表/结果.xlsx", empty.exception.detail)
        self.assertIn("规则/规则.docx", empty.exception.detail)

    def test_missing_path_with_recorded_name_is_not_ignored(self):
        self.round.source_file_names.append("遗漏.xlsx")
        with self.assertRaises(HTTPException) as raised:
            namespace["_sop_build_package"](self.round)
        self.assertIn("遗漏.xlsx", raised.exception.detail)


if __name__ == "__main__":
    unittest.main()
