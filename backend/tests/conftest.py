"""统一测试导入路径，兼容早期 ai_engine 顶层引用。"""
import sys
from pathlib import Path

backend_dir = Path(__file__).resolve().parents[1]
project_dir = backend_dir.parent
for directory in (project_dir, backend_dir):
    if str(directory) not in sys.path:
        sys.path.insert(0, str(directory))
