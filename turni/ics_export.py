"""Esportazione turni in formato iCalendar (.ics)."""
from __future__ import annotations

import re
import uuid
from datetime import date, datetime, timedelta, timezone

_MESI_ABBR_NUM: dict[str, int] = {
    "gen": 1, "feb": 2, "mar": 3, "apr": 4,
    "mag": 5, "giu": 6, "lug": 7, "ago": 8,
    "set": 9, "ott": 10, "nov": 11, "dic": 12,
}


def _parse_week_dates(week_label: str, anno: str) -> tuple[date | None, date | None]:
    """Estrae le date giovedi/domenica dal label della settimana."""
    year = int(anno)
    compact = re.sub(r"\s*-\s*", "-", (week_label or "").strip())
    m = re.search(
        r"(?i)(\d{1,2})\s*([A-Za-z\u00C0-\u00FF]{3,})?\s*-\s*(\d{1,2})\s*([A-Za-z\u00C0-\u00FF]{3,})?",
        compact,
    )
    if not m:
        return None, None
    d1, m1, d2, m2 = m.groups()
    if not m1 and m2:
        m1 = m2
    if not m2 and m1:
        m2 = m1
    if not m1:
        return None, None
    mo1 = _MESI_ABBR_NUM.get(m1[:3].lower())
    mo2 = _MESI_ABBR_NUM.get(m2[:3].lower()) if m2 else mo1
    if not mo1 or not mo2:
        return None, None
    try:
        return date(year, mo1, int(d1)), date(year, mo2, int(d2))
    except ValueError:
        return None, None


def _esc(text: str) -> str:
    return (text
            .replace("\\", "\\\\")
            .replace(";", "\\;")
            .replace(",", "\\,")
            .replace("\n", "\\n"))


def build_ics(anno: str, result_rows: list[dict]) -> str:
    """Genera un calendario ICS con due eventi per settimana (giovedi + domenica)."""
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//SoftwareTurniAcustica//IT",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Turni Audio/Video",
    ]
    now = datetime.now(tz=timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    for row in result_rows:
        thu, sun = _parse_week_dates(row.get("week", ""), anno)
        if thu is None or sun is None:
            continue
        audio = row.get("audio", "")
        video = row.get("video", "")
        sabato = row.get("sabato", "")
        site = row.get("site", "")
        tag = f" [{site}]" if site else ""

        thu_s = thu.strftime("%Y%m%d")
        thu_e = (thu + timedelta(days=1)).strftime("%Y%m%d")
        sun_s = sun.strftime("%Y%m%d")
        sun_e = (sun + timedelta(days=1)).strftime("%Y%m%d")

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uuid.uuid4()}",
            f"DTSTAMP:{now}",
            f"DTSTART;VALUE=DATE:{thu_s}",
            f"DTEND;VALUE=DATE:{thu_e}",
            f"SUMMARY:{_esc(f'Giovedi{tag} - Audio: {audio}, Video: {video}')}",
            f"DESCRIPTION:{_esc(f'Audio: {audio}\\nVideo: {video}')}",
            "END:VEVENT",
            "BEGIN:VEVENT",
            f"UID:{uuid.uuid4()}",
            f"DTSTAMP:{now}",
            f"DTSTART;VALUE=DATE:{sun_s}",
            f"DTEND;VALUE=DATE:{sun_e}",
            f"SUMMARY:{_esc(f'Domenica{tag} - A/V: {sabato}')}",
            f"DESCRIPTION:{_esc(f'Operatore A/V: {sabato}')}",
            "END:VEVENT",
        ]

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


def format_whatsapp(anno: str, result_rows: list[dict],
                    operatori: list[str], counts: list[int]) -> str:
    """Formatta i risultati per WhatsApp/Telegram (copia-incolla)."""
    parts: list[str] = []
    parts.append(f"*TURNI AUDIO/VIDEO {anno}*")
    parts.append("")

    cur_month = ""
    cur_site = ""
    for r in result_rows:
        site = r.get("site", "")
        if site and site != cur_site:
            cur_site = site
            parts.append(f"--- {site.upper()} ---")
        month = r.get("month", "")
        if month != cur_month:
            cur_month = month
            parts.append(f"\n*{month.upper()}*")

        parts.append(f"{r['week']}")
        parts.append(f"  Gio: Audio={r['audio']}, Video={r['video']}")
        parts.append(f"  Dom: {r['sabato']}")
        if r.get("busy"):
            parts.append(f"  Busy: {r['busy']}")

    parts.append("\n*RIEPILOGO*")
    paired = sorted(zip(operatori, counts), key=lambda x: -x[1])
    for op, cnt in paired:
        parts.append(f"  {op}: {cnt} turni")

    return "\n".join(parts)
