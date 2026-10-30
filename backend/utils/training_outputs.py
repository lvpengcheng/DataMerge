"""Locate the exact revision, including archives nested by repeated migrations."""
from pathlib import Path


def iteration_files(session_root, number):
    root=Path(session_root).resolve()
    current=root
    for depth in range(12):
        directory=(current/f'iter_{int(number)}').resolve()
        if not directory.is_relative_to(root):
            return {}
        files={}
        if directory.is_dir():
            for path in sorted(directory.iterdir()):
                if not path.is_file() or not path.resolve().is_relative_to(root) or path.name.startswith('~'):
                    continue
                if path.name=='script.py':files['script_file']=str(path)
                elif path.name=='交付结果.zip':files['output_file']=str(path)
                elif path.name=='差异对比.zip':files['diff_file']=str(path)
                elif path.suffix.lower() in ('.xlsx','.xls','.xlsm'):
                    key='diff_file' if 'diff' in path.name.lower() or '差异对比' in path.name else 'output_file'
                    files.setdefault(key,str(path))
            for key,name in (('output_file','交付结果.zip'),('diff_file','差异对比.zip')):
                if (directory/name).is_file() and (directory/name).resolve().is_relative_to(root):
                    files[key]=str(directory/name)
            if files:
                return {**files,'archived':depth>0}
        current=current/'archive'/'session'
        if not current.is_dir():
            break
    return {}
