"""Funzioni pure di utilità condivise tra UI, solver e I/O."""
from __future__ import annotations


def normalize_name(name: str) -> str:
    return " ".join(name.strip().split()).lower()


def _safe_csv_cell(value: str) -> str:
    """Neutralizza formule potenzialmente eseguibili nei fogli di calcolo."""
    if not isinstance(value, str):
        return value
    stripped = value.lstrip()
    if stripped and stripped[0] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def _dedupe_normalized_texts(values: list[str]) -> tuple[list[str], list[str], list[str]]:
    """Deduplica stringhe preservando l'ordine e confrontando via normalize_name()."""
    unique_values: list[str] = []
    unique_norms: list[str] = []
    duplicates: list[str] = []
    for value in values:
        norm = normalize_name(value)
        if norm not in unique_norms:
            unique_values.append(value)
            unique_norms.append(norm)
        else:
            duplicates.append(value)
    return unique_values, unique_norms, duplicates


def _group_weeks_by_normalized_month(weeks: list[dict]) -> dict[str, list[dict]]:
    """Raggruppa le settimane usando una chiave mese normalizzata."""
    grouped: dict[str, list[dict]] = {}
    for week in weeks:
        grouped.setdefault(normalize_name(week["month"]), []).append(week)
    return grouped


def _order_weeks_by_declared_months(weeks: list[dict], declared_months: list[str]) -> list[dict]:
    """Ordina le settimane per mese dichiarato, preservando l'ordine interno al mese.

    Evita output interleavati del tipo Gennaio1, Febbraio1, Marzo1, Gennaio2...
    quando la UI o il caricamento sessione hanno costruito la lista in quell'ordine.
    """
    grouped = _group_weeks_by_normalized_month(weeks)
    ordered: list[dict] = []
    seen_months: set[str] = set()

    for month in declared_months:
        month_norm = normalize_name(month)
        if month_norm in grouped:
            ordered.extend(grouped[month_norm])
            seen_months.add(month_norm)

    for week in weeks:
        month_norm = normalize_name(week["month"])
        if month_norm not in seen_months:
            ordered.append(week)
            seen_months.add(month_norm)

    return ordered


def _find_blank_text_entries(values: list[str]) -> list[str]:
    """Restituisce gli elementi vuoti o composti solo da spazi."""
    return [value for value in values if not normalize_name(value)]


def _normalized_week_key(month: str, week: str) -> tuple[str, str]:
    """Chiave canonica di una settimana: mese + nome settimana normalizzati."""
    return normalize_name(month), normalize_name(week)


def _find_duplicate_week_keys(weeks: list[dict]) -> list[tuple[str, str]]:
    """Individua settimane duplicate nello stesso mese dopo normalizzazione."""
    duplicates: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    dup_seen: set[tuple[str, str]] = set()
    for week in weeks:
        key = _normalized_week_key(week["month"], week["week"])
        if key in seen and key not in dup_seen:
            duplicates.append((week["month"], week["week"]))
            dup_seen.add(key)
        else:
            seen.add(key)
    return duplicates


def _parse_operatori_text(raw: str) -> list[str]:
    return [x.strip() for part in raw.splitlines() for x in part.split(",") if x.strip()]


def _parse_mesi_text(raw: str) -> list[str]:
    return [m.strip() for m in raw.split(",") if m.strip()]
