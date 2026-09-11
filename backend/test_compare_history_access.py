import ast
import asyncio
from pathlib import Path
from types import SimpleNamespace
import pytest
from fastapi import HTTPException


def test_compare_history_detail_and_download_are_owner_only():
    tree = ast.parse((Path(__file__).parent / 'app/main.py').read_text(encoding='utf-8'))
    names = {'list_compare_history', 'get_compare_detail', 'download_compare_result'}
    functions = [n for n in tree.body if isinstance(n, ast.AsyncFunctionDef) and n.name in names]
    for fn in functions:
        fn.decorator_list = []
        fn.args.defaults = []
    scope = {'HTTPException': HTTPException, '_compare_history': [
        {'session_id': 'mine', 'user_id': 1}, {'session_id': 'other', 'user_id': 2},
        {'session_id': 'legacy'}]}
    exec(compile(ast.fix_missing_locations(ast.Module(body=functions, type_ignores=[])), '<access>', 'exec'), scope)
    user = SimpleNamespace(id=1)
    assert asyncio.run(scope['list_compare_history'](user)) == [{'session_id': 'mine', 'user_id': 1}]
    for session in ['other', 'legacy']:
        with pytest.raises(HTTPException) as error:
            asyncio.run(scope['get_compare_detail'](session, user))
        assert error.value.status_code == 404
        with pytest.raises(HTTPException) as error:
            asyncio.run(scope['download_compare_result'](session, 'result.xlsx', user))
        assert error.value.status_code == 404
