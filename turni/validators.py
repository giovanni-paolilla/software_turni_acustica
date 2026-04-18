"""Validazione e canonizzazione dei dati di sessione e di Step 0."""
from __future__ import annotations
import re

from turni.helpers import (
    normalize_name,
    _dedupe_normalized_texts,
    _find_blank_text_entries,
    _find_duplicate_week_keys,
    _parse_operatori_text,
    _parse_mesi_text,
)


class SessionValidationError(ValueError):
    """Errore di validazione dei dati di sessione."""


def _canonicalize_timeout_text(timeout_value, *, error_message: str) -> str:
    timeout_text = str(timeout_value).strip()
    try:
        timeout_val = float(timeout_text)
        if timeout_val <= 0:
            raise ValueError
    except (TypeError, ValueError) as exc:
        raise SessionValidationError(error_message) from exc
    # Rappresentazione intera per valori interi (es. "60" non "60.0")
    if timeout_val == int(timeout_val):
        return str(int(timeout_val))
    return str(timeout_val)


def _validate_operator_pool(operatori: list[str], *, error_message: str) -> None:
    if len(operatori) < 2:
        raise SessionValidationError(error_message)


def _week_available_count(week: dict, operator_count: int) -> int:
    if "available" in week:
        available = week.get("available")
        if isinstance(available, list):
            return len(set(available))
    busy_indices = week.get("busy_indices")
    if isinstance(busy_indices, list):
        return operator_count - len(set(busy_indices))
    busy_names = week.get("busy")
    if isinstance(busy_names, list):
        return operator_count - len(set(busy_names))
    return operator_count


def _validate_solver_ready_weeks(weeks: list[dict], *, operator_count: int,
                                  error_prefix: str = "File sessione non valido") -> None:
    if operator_count < 2:
        raise SessionValidationError(
            f"{error_prefix}: servono almeno 2 operatori totali per pianificare i turni."
        )
    for week in weeks:
        if _week_available_count(week, operator_count) < 2:
            raise SessionValidationError(
                f"{error_prefix}: '{week['week']}' ({week['month']}): meno di 2 operatori disponibili."
            )


def _validate_week_entries(weeks: list[dict], *, declared_months: list[str] | None = None,
                            error_prefix: str = "File sessione non valido") -> list[dict]:
    """Valida e canonizza le settimane condividendo le regole tra UI e sessione."""
    declared_months_norm = None
    if declared_months is not None:
        declared_months_norm = {normalize_name(m) for m in declared_months}

    canonical_weeks: list[dict] = []
    for week in weeks:
        month = week.get("month")
        week_name = week.get("week")

        if not isinstance(month, str):
            raise SessionValidationError(
                f"{error_prefix}: il campo 'month' di ogni settimana deve essere testo."
            )
        if not normalize_name(month):
            raise SessionValidationError(
                f"{error_prefix}: il campo 'month' di una settimana e' vuoto o composto solo da spazi."
            )
        if declared_months_norm is not None and normalize_name(month) not in declared_months_norm:
            raise SessionValidationError(
                f"{error_prefix}: presenti settimane associate a mesi non dichiarati nella sessione."
            )
        if not isinstance(week_name, str):
            raise SessionValidationError(
                f"{error_prefix}: il nome di ogni settimana deve essere testo."
            )
        if not normalize_name(week_name):
            raise SessionValidationError(
                f"{error_prefix}: presenti settimane con nome vuoto o solo spazi."
            )

        canonical_weeks.append({
            **week,
            "month": month.strip(),
            "week": week_name.strip(),
        })

    duplicate_weeks = _find_duplicate_week_keys(canonical_weeks)
    if duplicate_weeks:
        dup_month, dup_week = duplicate_weeks[0]
        raise SessionValidationError(
            f"{error_prefix}: settimana duplicata '{dup_week}' nel mese '{dup_month}'."
        )

    return canonical_weeks


def _parse_step0_inputs(anno: str, mesi_text: str, operatori_text: str, timeout_text: str
                        ) -> tuple[str, str, list[str], list[str], list[str], list[str], list[str]]:
    """Parsa e valida i campi dello step 0 restituendo una forma canonica."""
    anno = anno.strip()
    if not re.fullmatch(r"\d{4}", anno):
        raise SessionValidationError("Anno non valido. Inserisci 4 cifre (es. 2025).")

    timeout_text = _canonicalize_timeout_text(
        timeout_text,
        error_message="Timeout non valido. Inserisci un numero positivo (es. 60).",
    )

    raw_ops = _parse_operatori_text(operatori_text)
    if not raw_ops:
        raise SessionValidationError("Inserisci almeno un operatore.")
    if _find_blank_text_entries(raw_ops):
        raise SessionValidationError(
            "Gli operatori non possono essere vuoti o composti solo da spazi."
        )

    raw_mesi = _parse_mesi_text(mesi_text)
    if not raw_mesi:
        raise SessionValidationError("Inserisci almeno un mese.")
    if _find_blank_text_entries(raw_mesi):
        raise SessionValidationError(
            "I mesi non possono essere vuoti o composti solo da spazi."
        )

    new_ops, new_norm, dup_ops = _dedupe_normalized_texts(raw_ops)
    new_mesi, _, dup_mesi = _dedupe_normalized_texts(raw_mesi)
    _validate_operator_pool(
        new_ops,
        error_message=(
            "Servono almeno 2 operatori totali per pianificare i turni "
            "(giovedi Audio e Video devono essere assegnati a persone diverse)."
        ),
    )
    return anno, timeout_text, new_ops, new_norm, new_mesi, dup_ops, dup_mesi


def _validate_session_payload(session: dict) -> dict:
    """Valida e canonizza il payload di sessione usato sia in save che in load."""
    if not isinstance(session, dict):
        raise SessionValidationError(
            "File sessione non valido: la radice JSON deve essere un oggetto."
        )

    required_session_keys = {"anno", "operatori", "mesi", "weeks"}
    if not required_session_keys.issubset(session):
        missing = required_session_keys - session.keys()
        raise SessionValidationError(
            f"File sessione non valido.\nCampi mancanti: {', '.join(sorted(missing))}"
        )

    ops = session["operatori"]
    mesi = session["mesi"]
    weeks = session["weeks"]

    if (not isinstance(ops, list) or not all(isinstance(o, str) for o in ops)
            or not isinstance(mesi, list)
            or not all(isinstance(m, str) for m in mesi)):
        raise SessionValidationError(
            "File sessione non valido: 'operatori' e 'mesi' devono essere liste di testo."
        )

    ops = [o.strip() for o in ops]
    mesi = [m.strip() for m in mesi]

    if not ops:
        raise SessionValidationError("File sessione non valido: inserisci almeno un operatore.")
    if not mesi:
        raise SessionValidationError("File sessione non valido: inserisci almeno un mese.")

    if _find_blank_text_entries(ops):
        raise SessionValidationError(
            "File sessione non valido: operatori vuoti o composti solo da spazi."
        )
    if _find_blank_text_entries(mesi):
        raise SessionValidationError(
            "File sessione non valido: mesi vuoti o composti solo da spazi."
        )

    _, _, dup_ops = _dedupe_normalized_texts(ops)
    if dup_ops:
        raise SessionValidationError(
            "File sessione non valido: operatori duplicati dopo normalizzazione.\n"
            + "\n".join(dup_ops)
        )

    _validate_operator_pool(
        ops,
        error_message=(
            "File sessione non valido: servono almeno 2 operatori totali per "
            "pianificare i turni."
        ),
    )

    _, _, dup_mesi = _dedupe_normalized_texts(mesi)
    if dup_mesi:
        raise SessionValidationError(
            "File sessione non valido: mesi duplicati dopo normalizzazione.\n"
            + "\n".join(dup_mesi)
        )

    if not isinstance(weeks, list):
        raise SessionValidationError("File sessione non valido: 'weeks' deve essere una lista.")
    if not all(isinstance(w, dict) for w in weeks):
        raise SessionValidationError(
            "File sessione non valido: ogni settimana deve essere un oggetto."
        )

    required_week_keys = {"month", "week", "busy_indices"}
    bad_weeks = [w for w in weeks if not required_week_keys.issubset(w)]
    if bad_weeks:
        raise SessionValidationError(
            f"File sessione non valido: {len(bad_weeks)} settimana/e malformata/e."
        )

    canonical_weeks = _validate_week_entries(
        weeks,
        declared_months=mesi,
        error_prefix="File sessione non valido",
    )
    for week in canonical_weeks:
        busy_indices = week.get("busy_indices", [])
        if (not isinstance(busy_indices, list)
                or not all(isinstance(i, int) and i >= 0 for i in busy_indices)):
            raise SessionValidationError(
                f"busy_indices non valido nella settimana '{week['week']}'.\nDevono essere interi non negativi."
            )
        week["busy_indices"] = list(busy_indices)

    too_large_idx = [
        (w["week"], idx)
        for w in canonical_weeks
        for idx in w["busy_indices"]
        if idx >= len(ops)
    ]
    if too_large_idx:
        week_name, idx = too_large_idx[0]
        raise SessionValidationError(
            f"busy_indices fuori range nella settimana '{week_name}' (indice {idx})."
        )

    _validate_solver_ready_weeks(
        canonical_weeks,
        operator_count=len(ops),
        error_prefix="File sessione non valido",
    )

    anno_val = str(session.get("anno", "")).strip()
    if not re.fullmatch(r"\d{4}", anno_val):
        raise SessionValidationError(
            f"Anno nella sessione non valido: '{anno_val}' (atteso formato YYYY)."
        )

    validated = {
        "anno": anno_val,
        "operatori": ops,
        "mesi": mesi,
        "weeks": canonical_weeks,
    }
    if "solver_timeout" in session:
        validated["solver_timeout"] = _canonicalize_timeout_text(
            session["solver_timeout"],
            error_message="File sessione non valido: solver_timeout deve essere un numero positivo.",
        )
    return validated
