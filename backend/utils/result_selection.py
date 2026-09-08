"""Choose an unambiguous business workbook before validation or publication."""
import hashlib
from pathlib import Path


def pick_result_output(output_files, template_path=None):
    files = sorted({Path(path) for path in output_files}, key=lambda p: p.name.casefold())
    files = [p for p in files if p.is_file() and p.suffix.lower() == '.xlsx'
             and not p.name.startswith(('~', '_'))
             and not p.stem.startswith('差异对比')
             and p.stem.casefold() not in ('diff', 'comparison')]
    if len(files) <= 1:
        return files[0] if files else None
    if template_path and Path(template_path).is_file():
        template = Path(template_path)
        size = template.stat().st_size
        def digest(path):
            with path.open('rb') as stream:
                return hashlib.file_digest(stream, 'sha256').digest()
        template_digest = None
        candidates = []
        for path in files:
            if path.stat().st_size != size:
                candidates.append(path)
                continue
            if template_digest is None:
                template_digest = digest(template)
            if digest(path) != template_digest:
                candidates.append(path)
        files = candidates
    if len(files) == 1:
        return files[0]
    if not files:
        raise ValueError('输出文件均为未修改的模板副本，未找到计算结果')
    raise ValueError('存在多个结果工作簿，无法确定最终结果；请让脚本仅保留最终工作簿: '
                     + ', '.join(path.name for path in files))
