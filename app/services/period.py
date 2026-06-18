"""Traducción del `period` del upsert (atajo con nombre o rango explícito) a un
intervalo UTC ``[inicio, fin)`` listo para el ``$filter`` de Graph.

Los límites se calculan en la **zona horaria del tenant** (config
``TENANT_TIMEZONE``, IANA; por defecto UTC) y luego se convierten a UTC, porque
Graph almacena y compara las fechas de las columnas en UTC. Así el corte de
"mes en curso" coincide con cómo el usuario ve las fechas en SharePoint.

El intervalo es **semiabierto**: incluye el instante inicial y excluye el final,
de modo que periodos contiguos no se solapen. Los atajos con nombre soportados
son ``current_day``, ``current_week`` (semana ISO, empieza lunes),
``current_month`` y ``current_year``.
"""

from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from app.core.exceptions import GraphAPIError

# Formato de literal de fecha-hora UTC que se interpola en el $filter de Graph.
_GRAPH_DT_FMT = "%Y-%m-%dT%H:%M:%SZ"

_NAMED_PERIODS = frozenset(
    {"current_day", "current_week", "current_month", "current_year"}
)


def _load_tz(name: str) -> ZoneInfo:
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as exc:
        # Config del servidor errónea (o falta tzdata en el host) → 500, no 400:
        # no es culpa del caller.
        raise GraphAPIError(
            500, f"Zona horaria del tenant inválida ({name!r}): {exc}"
        ) from exc


def _midnight(d: date, tz: ZoneInfo) -> datetime:
    """Medianoche (00:00) de la fecha ``d`` en la zona ``tz`` (tz-aware)."""
    return datetime(d.year, d.month, d.day, tzinfo=tz)


def _first_of_next_month(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)


def _named_bounds(name: str, now: datetime, tz: ZoneInfo) -> tuple[datetime, datetime]:
    today = now.date()
    if name == "current_day":
        return _midnight(today, tz), _midnight(today + timedelta(days=1), tz)
    if name == "current_week":
        monday = today - timedelta(days=today.weekday())
        return _midnight(monday, tz), _midnight(monday + timedelta(days=7), tz)
    if name == "current_month":
        first = today.replace(day=1)
        return _midnight(first, tz), _midnight(_first_of_next_month(first), tz)
    if name == "current_year":
        return _midnight(date(today.year, 1, 1), tz), _midnight(
            date(today.year + 1, 1, 1), tz
        )
    raise GraphAPIError(
        400,
        f"period desconocido: {name!r}. Use un atajo válido "
        f"({', '.join(sorted(_NAMED_PERIODS))}) o un rango explícito {{from, to}}.",
    )


def _parse_day(value: str, field: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise GraphAPIError(
            400, f"Fecha inválida en period.{field}: {value!r} (formato esperado YYYY-MM-DD)"
        ) from exc


def _explicit_bounds(
    period_from: str, period_to: str, tz: ZoneInfo
) -> tuple[datetime, datetime]:
    start = _parse_day(period_from, "from")
    end = _parse_day(period_to, "to")
    if end < start:
        raise GraphAPIError(
            400, f"Rango de period inválido: 'to' ({period_to}) es anterior a 'from' ({period_from})"
        )
    # `to` es inclusivo: el intervalo semiabierto llega hasta el inicio del día
    # siguiente para cubrir todas las horas del día `to`.
    return _midnight(start, tz), _midnight(end + timedelta(days=1), tz)


def resolve_period(
    period: "str | tuple[str, str] | None", tz_name: str
) -> tuple[str, str] | None:
    """Resuelve ``period`` a un par ``(inicio_utc, fin_utc)`` en formato Graph.

    - ``None`` → devuelve ``None`` (la coincidencia será solo por clave).
    - ``str`` → atajo con nombre (``current_month``, …).
    - ``(from, to)`` → rango explícito de fechas ``YYYY-MM-DD`` (``to`` inclusive).
    """
    if period is None:
        return None

    tz = _load_tz(tz_name)
    if isinstance(period, str):
        start, end = _named_bounds(period, datetime.now(tz), tz)
    else:
        start, end = _explicit_bounds(period[0], period[1], tz)

    utc = ZoneInfo("UTC")
    return start.astimezone(utc).strftime(_GRAPH_DT_FMT), end.astimezone(utc).strftime(
        _GRAPH_DT_FMT
    )
