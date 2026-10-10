from contextlib import contextmanager

import pytest

from backend.database import connection


class BrokenEngine:
    def __init__(self, error):
        self.error = error

    @contextmanager
    def connect(self):
        raise self.error
        yield


@pytest.mark.parametrize('error,expected', [
    (OSError('timeout expired'), '目标主机或端口不可达'),
    (OSError('connection refused'), '数据库端口未监听'),
    (OSError('password authentication failed'), '账号或密码'),
])
def test_database_startup_error_is_actionable_and_never_exposes_password(monkeypatch, error, expected):
    monkeypatch.setattr(connection, 'engine', BrokenEngine(error))
    monkeypatch.setattr(connection, 'DATABASE_URL',
                        'postgresql+psycopg2://dbuser:top-secret@db.internal:5432/payroll')
    with pytest.raises(RuntimeError) as caught:
        connection.verify_database_connection()
    message = str(caught.value)
    assert expected in message
    assert 'db.internal:5432/payroll' in message
    assert 'dbuser' not in message and 'top-secret' not in message


def test_database_timeout_configuration_is_applied_to_engine():
    assert connection.engine.url.get_backend_name() == 'postgresql'
    assert connection.DATABASE_CONNECT_ARGS['connect_timeout'] == connection.DATABASE_CONNECT_TIMEOUT


def test_startup_does_not_overwrite_existing_examples(tmp_path, monkeypatch):
    import run
    monkeypatch.setattr(run, 'project_root', tmp_path)
    examples = tmp_path / 'examples'
    examples.mkdir()
    existing = examples / 'source1.xlsx'
    existing.write_bytes(b'user workbook')
    run.create_example_files()
    assert existing.read_bytes() == b'user workbook'
    assert (examples / 'source2.xlsx').is_file()
    assert (examples / 'expected_result.xlsx').is_file()
