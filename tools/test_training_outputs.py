import ast
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock
from datetime import datetime

import pytest
from fastapi import HTTPException

from backend.utils.training_outputs import iteration_files


@pytest.mark.parametrize('depth',[0,1,2])
def test_revision_results_survive_repeated_migration(tmp_path,depth):
    root=tmp_path
    for _ in range(depth):root=root/'archive'/'session'
    folder=root/'iter_18';folder.mkdir(parents=True)
    for name in ('script.py','汇总.xlsx','交付结果.zip','差异对比.zip'):
        (folder/name).write_bytes(name.encode())
    files=iteration_files(tmp_path,18)
    assert Path(files['output_file']).name=='交付结果.zip'
    assert Path(files['diff_file']).name=='差异对比.zip'
    assert files['archived']==(depth>0)
    assert iteration_files(tmp_path,19)=={}


def test_imported_code_revision_is_not_presented_as_a_new_execution(tmp_path):
    (tmp_path/'iter_19').mkdir();(tmp_path/'iter_19/script.py').write_text('code')
    old=tmp_path/'archive/session/iter_18';old.mkdir(parents=True);(old/'交付结果.zip').write_bytes(b'old results')
    assert 'output_file' not in iteration_files(tmp_path,19)
    assert iteration_files(tmp_path,18)['archived']


def test_requested_iteration_never_downloads_another_iterations_latest_result(tmp_path):
    tree=ast.parse(Path('backend/api/training_chat.py').read_text(encoding='utf-8'))
    fn=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='download_iteration_file')
    fn.decorator_list=[];fn.returns=None
    for arg in fn.args.args:arg.annotation=None
    fn.args.defaults=[ast.Constant(None)]*len(fn.args.defaults)
    ast.fix_missing_locations(fn)
    namespace={'__package__':'backend.api','TrainingSession':object(),'HTTPException':HTTPException,
               '_session_asset_root':lambda session:tmp_path,'datetime':datetime,'re':__import__('re'),'os':__import__('os'),
               'Path':Path,'_NO_STORE_HEADERS':{}}
    exec(compile(ast.Module(body=[fn],type_ignores=[]),'<download>','exec'),namespace)
    latest=tmp_path/'latest.xlsx';latest.write_bytes(b'not requested')
    session=SimpleNamespace(config={'latest_files':{'output_file':str(latest)}},tenant_id='tenant')
    db=MagicMock();db.query.return_value.filter_by.return_value.first.return_value=session
    with pytest.raises(HTTPException) as exc:
        namespace['download_iteration_file'](1,'output',iteration=18,db=db)
    assert exc.value.status_code==404
