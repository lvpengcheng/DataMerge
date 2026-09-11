"""Exercise the runtime publication gate without importing the web application."""
import ast
import os
from pathlib import Path

import pytest


@pytest.mark.parametrize('setting,blocked', [(None, False), ('false', False), ('true', True)])
def test_formula_errors_default_to_downloadable_warning(monkeypatch, setting, blocked):
    if setting is None:
        monkeypatch.delenv('COMPUTE_STRICT_FORMULA_ERRORS', raising=False)
    else:
        monkeypatch.setenv('COMPUTE_STRICT_FORMULA_ERRORS', setting)
    tree = ast.parse((Path(__file__).parent / 'app' / 'main.py').read_text(encoding='utf-8'))
    gate = next(node for node in ast.walk(tree) if isinstance(node, ast.If)
                and 'COMPUTE_STRICT_FORMULA_ERRORS' in ast.unparse(node.test))
    program = compile(ast.fix_missing_locations(ast.Module(body=[gate], type_ignores=[])), '<gate>', 'exec')
    scope = {'os': os, '_formula_errors': 4, '_formula_bad_refs': 2}
    if blocked:
        with pytest.raises(RuntimeError, match='4 个公式错误'):
            exec(program, scope)
    else:
        exec(program, scope)
