"""Exercise the real reparse handler without loading the application/database."""
import ast
import asyncio
import json
import logging
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest
from fastapi import HTTPException


def handler(tmp_path):
    source = Path(__file__).resolve().parents[1] / 'backend/api/tools.py'
    tree = ast.parse(source.read_text(encoding='utf-8'))
    node = next(n for n in tree.body if isinstance(n, ast.AsyncFunctionDef)
                and n.name == '_integrate_reparse_headers_impl')
    files = [dict(name=name, sheet='Sheet1', columns=['id'], header_start=1,
                  header_end=1, header_manual=False) for name in ('visible.xlsx', 'hidden.xlsx')]
    for file in files:
        (tmp_path / file['name']).touch()
    (tmp_path / '_meta.json').write_text(json.dumps({'files': files}), encoding='utf-8')
    parse = AsyncMock(side_effect=AssertionError('unchanged ranges should reuse parsed data'))
    scope = dict(json=json, time=time, HTTPException=HTTPException, EXCEL_EXTS={'.xlsx'},
                 logger=logging.getLogger(__name__), _workflow_session_dir=lambda *args: tmp_path,
                 _match_integrate_schemes=lambda *args: [], _parse_file_to_df_fresh=parse)
    exec(compile(ast.Module(body=[node], type_ignores=[]), str(source), 'exec'), scope)
    return scope['_integrate_reparse_headers_impl'], files, parse


@pytest.mark.parametrize('workflow', ['integrate', 'split'])
def test_one_row_header_with_hidden_upload(tmp_path, workflow):
    run, files, parse = handler(tmp_path)
    request = SimpleNamespace(session_id='test', header_ranges={
        'visible.xlsx': {'start_row': 1, 'end_row': 1}}, sheet_selections={}, scheme_id=None)
    result = asyncio.run(run(request, None, workflow))
    assert result['files'][0]['header_manual'] is True
    assert result['files'][1] == files[1]
    assert result['files'][0]['header_start'] == result['files'][0]['header_end'] == 1
    parse.assert_not_called()


@pytest.mark.parametrize('value', [None, '', 1.5, '1.5', True, 0, -1])
def test_invalid_rows_remain_errors(tmp_path, value):
    run, _, _ = handler(tmp_path)
    before = (tmp_path / '_meta.json').read_bytes()
    request = SimpleNamespace(session_id='test', header_ranges={
        'visible.xlsx': {'start_row': value, 'end_row': 1}}, sheet_selections={}, scheme_id=None)
    with pytest.raises(HTTPException) as error:
        asyncio.run(run(request, None, 'integrate'))
    assert error.value.status_code == 400
    assert (tmp_path / '_meta.json').read_bytes() == before
