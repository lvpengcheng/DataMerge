"""Recorded database timestamps are naive UTC; public timestamps declare UTC."""
from datetime import datetime, timezone


def naive_utc(value: datetime) -> datetime:
    if value.tzinfo is not None:
        return value.astimezone(timezone.utc).replace(tzinfo=None)
    return value


def utc_iso(value):
    if value is None:
        return None
    return naive_utc(value).replace(tzinfo=timezone.utc).isoformat().replace('+00:00', 'Z')
