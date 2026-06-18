"""Tests de la resolución de `period` (app/services/period.py).

Los atajos con nombre dependen de la hora actual, así que se validan por
invariantes (límites de medianoche, semiabierto, contiene "ahora"). Los rangos
explícitos y la conversión a UTC se comprueban de forma determinista.
"""

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from app.core.exceptions import GraphAPIError
from app.services.period import resolve_period

_FMT = "%Y-%m-%dT%H:%M:%SZ"


def test_none_period_returns_none():
    assert resolve_period(None, "UTC") is None


def test_explicit_range_utc_is_half_open_and_to_inclusive():
    start, end = resolve_period(("2026-01-01", "2026-06-30"), "UTC")
    assert start == "2026-01-01T00:00:00Z"
    # `to` inclusive → fin = medianoche del día siguiente.
    assert end == "2026-07-01T00:00:00Z"


def test_explicit_range_converts_tenant_tz_to_utc():
    # Madrid en enero es UTC+1 → medianoche local = 23:00 UTC del día anterior.
    start, end = resolve_period(("2026-01-01", "2026-01-31"), "Europe/Madrid")
    assert start == "2025-12-31T23:00:00Z"
    assert end == "2026-01-31T23:00:00Z"


def test_current_month_bounds_are_first_of_month_utc():
    now = datetime.now(ZoneInfo("UTC"))
    start, end = resolve_period("current_month", "UTC")
    assert start == now.replace(
        day=1, hour=0, minute=0, second=0, microsecond=0
    ).strftime(_FMT)
    # "ahora" cae dentro del intervalo semiabierto.
    assert start <= now.strftime(_FMT) < end


def test_current_day_is_one_day_window():
    start, end = resolve_period("current_day", "UTC")
    s = datetime.strptime(start, _FMT)
    e = datetime.strptime(end, _FMT)
    assert e - s == timedelta(days=1)
    assert s.hour == 0 and s.minute == 0


def test_current_week_is_seven_day_window_starting_monday():
    start, end = resolve_period("current_week", "UTC")
    s = datetime.strptime(start, _FMT)
    e = datetime.strptime(end, _FMT)
    assert e - s == timedelta(days=7)
    assert s.weekday() == 0  # lunes


def test_unknown_named_period_raises_400():
    with pytest.raises(GraphAPIError) as exc:
        resolve_period("current_decade", "UTC")
    assert exc.value.status_code == 400


def test_invalid_explicit_date_raises_400():
    with pytest.raises(GraphAPIError) as exc:
        resolve_period(("2026-13-01", "2026-12-31"), "UTC")
    assert exc.value.status_code == 400


def test_to_before_from_raises_400():
    with pytest.raises(GraphAPIError) as exc:
        resolve_period(("2026-06-30", "2026-01-01"), "UTC")
    assert exc.value.status_code == 400


def test_invalid_timezone_raises_500():
    with pytest.raises(GraphAPIError) as exc:
        resolve_period("current_month", "Not/AZone")
    assert exc.value.status_code == 500
