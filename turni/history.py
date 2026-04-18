"""Memoria storica dei turni per equita' inter-sessione."""
from __future__ import annotations

import json
import logging
import os
from datetime import datetime

from turni.helpers import normalize_name
from turni.io_utils import _write_text_file_atomic

logger = logging.getLogger(__name__)

_DEFAULT_DIR = os.path.join(os.path.expanduser("~"), ".turni_acustica")
_HISTORY_FILE = "history.json"


class HistoryStore:
    """Conteggio cumulativo dei turni assegnati, persistito su disco."""

    def __init__(self, directory: str | None = None) -> None:
        self.directory = directory or _DEFAULT_DIR
        self.filepath = os.path.join(self.directory, _HISTORY_FILE)
        self._data: dict = self._load()

    def _load(self) -> dict:
        if not os.path.exists(self.filepath):
            return {"sessions": [], "cumulative": {}}
        try:
            with open(self.filepath, encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                return data
        except (OSError, json.JSONDecodeError, ValueError):
            logger.warning("Storico non leggibile: %s", self.filepath)
        return {"sessions": [], "cumulative": {}}

    def _save(self) -> None:
        os.makedirs(self.directory, exist_ok=True)
        _write_text_file_atomic(
            self.filepath,
            json.dumps(self._data, ensure_ascii=False, indent=2),
        )

    def record_session(
        self,
        anno: str,
        mesi: list[str],
        operatori: list[str],
        counts: list[int],
    ) -> None:
        """Registra i conteggi di una pianificazione completata."""
        self._data.setdefault("sessions", []).append({
            "date": datetime.now().isoformat(timespec="seconds"),
            "anno": anno,
            "mesi": mesi,
            "counts": dict(zip(operatori, counts)),
        })
        cum = self._data.setdefault("cumulative", {})
        for op, cnt in zip(operatori, counts):
            key = normalize_name(op)
            cum[key] = cum.get(key, 0) + cnt
        self._save()

    def get_cumulative_counts(self, operatori: list[str]) -> list[int]:
        """Conteggi cumulativi allineati alla lista *operatori* fornita."""
        cum = self._data.get("cumulative", {})
        return [cum.get(normalize_name(op), 0) for op in operatori]

    def get_sessions(self) -> list[dict]:
        return list(self._data.get("sessions", []))

    def clear(self) -> None:
        self._data = {"sessions": [], "cumulative": {}}
        if os.path.exists(self.filepath):
            self._save()
