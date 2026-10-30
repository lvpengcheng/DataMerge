from datetime import datetime,timedelta,timezone
from backend.utils.timezone import utc_iso,naive_utc
from backend.solutions.restore import values
from backend.database.models import Script


def test_recorded_timestamps_declare_utc_without_rewriting_database_values():
    value=datetime(2026,10,29,0,15,30)
    assert utc_iso(value)=='2026-10-29T00:15:30Z'
    assert value.hour==0 and value.tzinfo is None
    assert utc_iso(None) is None


def test_migration_preserves_instants_when_timestamp_has_an_explicit_offset():
    aware=datetime(2026,10,29,8,15,30,tzinfo=timezone(timedelta(hours=8)))
    assert naive_utc(aware)==datetime(2026,10,29,0,15,30)
    assert values(Script,{'created_at':aware.isoformat()})['created_at']==datetime(2026,10,29,0,15,30)
