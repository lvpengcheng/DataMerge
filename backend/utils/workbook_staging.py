"""Best-effort staging cleanup after publishing explicit workbook outputs."""
import gc
import logging
import shutil
import sys
import tempfile
import time
from contextlib import contextmanager


@contextmanager
def workbook_staging(*, prefix='multi_', dir=None, **kwargs):
    path = tempfile.mkdtemp(prefix=prefix, dir=dir)
    try:
        yield path
    finally:
        for attempt in range(3):
            try:
                shutil.rmtree(path)
                break
            except FileNotFoundError:
                break
            except OSError as exc:
                if not isinstance(exc, PermissionError) and getattr(exc, 'winerror', None) not in (32, 33):
                    # Cleanup must never replace the original calculation error.
                    logging.getLogger(__name__).warning('临时工作簿目录清理失败: %s: %s', path, exc)
                    break
                gc.collect()
                if 'System' in sys.modules:
                    try:
                        from System import GC
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    except Exception:
                        logging.getLogger(__name__).debug('CLR 清理未完成', exc_info=True)
                if attempt == 2:
                    logging.getLogger(__name__).warning('临时工作簿仍被占用，保留至任务目录清理: %s', path)
                else:
                    time.sleep(0.1)


def install_legacy_cleanup(module):
    """Scope the old frozen module's tempfile facade to multi-workbook staging."""
    from types import SimpleNamespace
    original = module.tempfile
    def directory(*args, **kwargs):
        if not args and kwargs.get('prefix') == 'multi_':
            return workbook_staging(**kwargs)
        return original.TemporaryDirectory(*args, **kwargs)
    module.tempfile = SimpleNamespace(**{
        **{name: getattr(original, name) for name in dir(original) if not name.startswith('__')},
        'TemporaryDirectory': directory,
    })
