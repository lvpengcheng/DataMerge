"""Deployment packages include app dependencies and exclude local runtime data."""
import tarfile
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import deploy


class DeployPackagingTests(unittest.TestCase):
    def test_runtime_data_and_local_dependencies_are_pruned(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            included = {
                'backend/app/main.py', 'backend/api/tools.py',
                'frontend/templates/training.html', 'frontend/static/js/training.js',
                'backend/admin/templates/user_import_template.xlsx',
                'libs/Aspose.Cells.dll', 'libs/libSkiaSharp.so', 'fonts/app.ttf',
                '.env.example', 'requirements.txt', 'Dockerfile', '.dockerignore',
                'docker-compose.yml', 'run.py', 'excel_parser.py',
                'aspose_init.py', 'split_by_banner.py', 'deploy_native_ubuntu.sh',
            }
            excluded = {
                'solution_store/version/archive/source.xlsx',
                'tenants/example/training_chat/input.xlsx', 'global_assets/reference.xlsx',
                'tmp/pip/pandas.py', '.codex-build/testdeps/pandas.py',
                'outputs/result.xlsx', 'output/result.xlsx', 'temp/old.tar.gz',
                '_deploy_backups/old.tar.gz', 'tests/test_case.py', 'tools/debug.py',
                'backend/tests/test_case.py', 'backend/test_data/source.xlsx',
                'docs/manual.md', 'design_docs/design.md', 'examples/input.xlsx',
                '参考/notes.md', 'dist/application.exe', 'build_iis/package.zip',
                '.codex/settings.json', '.git/objects/data', 'venv/library.py',
                '.env', '.env.dev', 'data.db', 'deploy.config.dev.json',
                'deploy.native.config.json', 'backend/__pycache__/main.pyc',
                'datamerge_deploy_old.tar.gz', 'datamerge_backup_old.tar.gz',
            }
            for name in included | excluded:
                file = root / name
                file.parent.mkdir(parents=True, exist_ok=True)
                file.write_text('fixture', encoding='utf-8')
            with patch.object(deploy, 'PROJECT_ROOT', root):
                archive = root / 'temp/package.tar.gz'
                count = deploy.build_tarball(archive, deploy.EXCLUDE_DIRS)
                with tarfile.open(archive) as stream:
                    names = set(stream.getnames())
                self.assertEqual(names, included)
                self.assertEqual(count, len(included))
                code_archive = root / 'temp/code.tar.gz'
                deploy.build_tarball(code_archive, deploy.EXCLUDE_DIRS | deploy.CODE_ONLY_EXTRA_DIRS)
                with tarfile.open(code_archive) as stream:
                    names = set(stream.getnames())
                self.assertEqual(names, {name for name in included if not name.startswith(('libs/', 'fonts/'))})

    def test_output_archive_cannot_include_itself(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / 'run.py').write_text('# fixture', encoding='utf-8')
            with patch.object(deploy, 'PROJECT_ROOT', root):
                archive = root / 'package.tar.gz'
                self.assertEqual(deploy.build_tarball(archive, deploy.EXCLUDE_DIRS), 1)
                with tarfile.open(archive) as stream:
                    self.assertEqual(stream.getnames(), ['run.py'])


if __name__ == '__main__':
    unittest.main()
