"""AI 样本隐私与源表日期回归，不调用外部 AI。"""
import ast
import copy
import json
import logging
import os
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List, Optional

import openpyxl
import pandas as pd
import pytest
from openpyxl.utils.datetime import from_excel

from backend.utils.desensitize import mask_ai_samples
from backend.utils.source_sheet_writer import write_source_dataframe
from backend.ai_engine.rule_organizer import RuleOrganizer
from backend.ai_engine.prompt_generator import PromptGenerator


@pytest.mark.parametrize("header", ["姓名", "英文名", "联系人", "身份证号码", "公司名称", "单位名称",
                                    "通讯地址", "手机号码", "邮箱", "银行账号", "employee_name",
                                    "ID Number", "Company Name", "Home Address"])
def test_sensitive_samples_are_fully_masked_without_mutation(header):
    rows = [{"A": "Sensitive-31012345678901234X", "B": 42}, {"A": None}, {"A": ""}]
    before = copy.deepcopy(rows)
    masked = mask_ai_samples({header: "A", "数量": "B"}, rows)
    assert masked == [{"A": "[已脱敏]", "B": 42}, {"A": None}, {"A": ""}]
    assert rows == before


@pytest.mark.parametrize("schema", [{}, {"field_type": "date", "number_format": "General"},
                                    {"field_type": "decimal", "number_format": "0.00"},
                                    {"field_type": "text", "number_format": "@"}])
@pytest.mark.parametrize("source_format", ["General", "yyyy/mm/dd"])
def test_date_serial_roundtrip_overrides_stale_metadata(tmp_path, schema, source_format):
    wb = openpyxl.Workbook()
    write_source_dataframe(wb.active, pd.DataFrame({"日期": [45992, None], "金额": [45992, 1]}),
                           {"日期": schema}, {"日期": source_format})
    path = tmp_path / "source.xlsx"
    wb.save(path)
    out = openpyxl.load_workbook(path)
    assert out.active["A2"].value == from_excel(45992)
    assert out.active["A2"].number_format == ("yyyy/mm/dd" if source_format != "General" else "yyyy-mm-dd")
    assert out.active["A3"].value is None
    assert out.active["B2"].value == 45992
    assert out.active["B2"].number_format == "General"
    out.close()


def fake_regions():
    return [SimpleNamespace(sheet_name="数据", regions=[SimpleNamespace(
        head_data={"姓名": "A", "身份证": "B", "公司名": "C", "地址": "D", "数量": "E"},
        data=[{"A": "隐私姓名", "B": "31012345678901234X", "C": "隐私公司", "D": "隐私地址", "E": 42}],
        formula={}, column_schemas={}), SimpleNamespace(
        head_data={"数量": "A"}, data=[{"A": 99}], formula={}, column_schemas={})])]


def assert_private(result):
    text = json.dumps(result, ensure_ascii=False, default=str)
    for secret in ("隐私姓名", "31012345678901234X", "隐私公司", "隐私地址"):
        assert secret not in text
    assert "[已脱敏]" in text
    assert "42" in text


def test_rule_source_and_target_samples_are_masked():
    parsed = fake_regions()
    organizer = RuleOrganizer.__new__(RuleOrganizer)
    organizer.excel_parser = SimpleNamespace(parse_excel_file=lambda *a, **kw: parsed)
    organizer._formula_context = lambda path: ""
    assert_private(organizer._extract_source_structures(["source.xlsx"]))
    assert_private(organizer._extract_target_structure("target.xlsx"))
    assert parsed[0].regions[0].data[0]["A"] == "隐私姓名"


@pytest.mark.parametrize("filename,owner,function", [
    ("api/training_chat.py", None, "_build_source_structure_from_dir_impl"),
    ("api/training_chat.py", None, "_analyze_expected_structure_impl"),
    ("ai_engine/training_engine.py", "TrainingEngine", "_analyze_source_structure"),
    ("ai_engine/training_engine.py", "TrainingEngine", "_analyze_expected_structure"),
])
def test_training_structure_entry_points_mask_each_region(tmp_path, monkeypatch, filename, owner, function):
    import excel_parser
    import backend.utils.formula_evidence as evidence
    parsed = fake_regions()
    parser = SimpleNamespace(parse_excel_file=lambda *a, **kw: parsed)
    monkeypatch.setattr(excel_parser, "IntelligentExcelParser", lambda: parser)
    monkeypatch.setattr(evidence, "collect_formula_evidence", lambda *a: {})
    path = tmp_path / "source.xlsx"
    path.touch()
    tree = ast.parse((Path(__file__).parent / filename).read_text(encoding="utf-8"))
    nodes = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == owner).body if owner else tree.body
    node = next(n for n in nodes if isinstance(n, ast.FunctionDef) and n.name == function)
    ns = dict(Dict=Dict, Any=Any, List=List, Optional=Optional, Path=Path, os=os,
              logger=logging.getLogger(__name__), mask_ai_samples=mask_ai_samples)
    exec(compile(ast.Module(body=[node], type_ignores=[]), filename, "exec"), ns)
    arg = str(tmp_path) if "from_dir" in function else [str(path)] if "source_structure" in function else str(path)
    args = [SimpleNamespace(excel_parser=parser, logger=ns["logger"], training_logger=SimpleNamespace(log_error=lambda msg: pytest.fail(msg))), arg] if owner else [arg]
    result = ns[function](*args)
    assert_private(result)
    assert parsed[0].regions[0].data[0]["A"] == "隐私姓名"


def test_prompt_masks_older_raw_samples_again():
    region = fake_regions()[0].regions[0]
    structure = {"sheets": {"数据": {"headers": region.head_data, "data_sample": region.data}}}
    assert_private(PromptGenerator()._compress_structure(structure))
    assert region.data[0]["A"] == "隐私姓名"
