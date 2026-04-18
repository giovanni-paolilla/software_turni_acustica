"""Generazione automatica delle settimane (giovedi-domenica) da calendario."""
from __future__ import annotations

import calendar
from datetime import date, timedelta

_MESI_IT: dict[str, int] = {
    "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4,
    "maggio": 5, "giugno": 6, "luglio": 7, "agosto": 8,
    "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12,
}

_MESI_ABBR: dict[int, str] = {
    1: "Gen", 2: "Feb", 3: "Mar", 4: "Apr",
    5: "Mag", 6: "Giu", 7: "Lug", 8: "Ago",
    9: "Set", 10: "Ott", 11: "Nov", 12: "Dic",
}


def month_name_to_number(name: str) -> int | None:
    """Converte un nome mese italiano (case-insensitive) in numero 1-12."""
    return _MESI_IT.get(name.strip().lower())


def generate_weeks_for_month(year: int, month_name: str) -> list[dict]:
    """Genera le coppie giovedi-domenica per un dato mese.

    Ogni settimana e' definita dal giovedi e dalla domenica successiva (+3gg).
    Se la domenica cade nel mese successivo, viene inclusa normalmente.

    Returns:
        Lista di dict con *month*, *week* (label), *thursday* e *sunday*
        (date ISO 8601) e *week_of_month* (indice 1-based del giovedi nel mese).
    """
    month_num = month_name_to_number(month_name)
    if month_num is None:
        return []

    thursdays: list[date] = [
        d
        for d in calendar.Calendar().itermonthdates(year, month_num)
        if d.month == month_num and d.weekday() == 3
    ]

    weeks: list[dict] = []
    for thu in thursdays:
        sun = thu + timedelta(days=3)
        ta, sa = _MESI_ABBR[thu.month], _MESI_ABBR[sun.month]
        if thu.month == sun.month:
            label = f"Sett. {thu.day:02d}-{sun.day:02d} {ta}"
        else:
            label = f"Sett. {thu.day:02d} {ta} - {sun.day:02d} {sa}"

        weeks.append({
            "month": month_name.strip(),
            "week": label,
            "thursday": thu.isoformat(),
            "sunday": sun.isoformat(),
            "week_of_month": _thursday_index(thu),
        })
    return weeks


def _thursday_index(thu: date) -> int:
    """Indice 1-based del giovedi nel suo mese (1o, 2o, 3o, ...)."""
    return sum(
        1 for day in range(1, thu.day + 1)
        if date(thu.year, thu.month, day).weekday() == 3
    )
