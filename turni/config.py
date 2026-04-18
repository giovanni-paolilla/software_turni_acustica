"""Configurazione utente persistente (sessioni recenti, template DOCX, sedi)."""
from __future__ import annotations

import json
import logging
import os

from turni.io_utils import _write_text_file_atomic

logger = logging.getLogger(__name__)

_DEFAULT_DIR = os.path.join(os.path.expanduser("~"), ".turni_acustica")
_CONFIG_FILE = "config.json"
_AUTOSAVE_FILE = "autosave.json"
_MAX_RECENT = 8


class UserConfig:
    """Preferenze utente e recovery automatico."""

    def __init__(self, directory: str | None = None) -> None:
        self.directory = directory or _DEFAULT_DIR
        self.config_path = os.path.join(self.directory, _CONFIG_FILE)
        self.autosave_path = os.path.join(self.directory, _AUTOSAVE_FILE)
        self._data: dict = self._load()

    def _load(self) -> dict:
        if not os.path.exists(self.config_path):
            return self._defaults()
        try:
            with open(self.config_path, encoding="utf-8") as fh:
                d = json.load(fh)
            return d if isinstance(d, dict) else self._defaults()
        except (OSError, json.JSONDecodeError):
            return self._defaults()

    @staticmethod
    def _defaults() -> dict:
        return {
            "recent_sessions": [],
            "docx_template": {
                "title": "AUDIO/VIDEO MESSINA-GANZIRRI",
                "title_color": "4C79C5",
                "subtitle_color": "A8D033",
                "font_title": "Times New Roman",
                "font_body": "Arial",
            },
            "sites": ["Messina", "Ganzirri"],
            "history_weight": 0.5,
            "autosave_interval_seconds": 120,
        }

    def _save(self) -> None:
        os.makedirs(self.directory, exist_ok=True)
        try:
            _write_text_file_atomic(
                self.config_path,
                json.dumps(self._data, ensure_ascii=False, indent=2),
            )
        except OSError:
            logger.warning("Impossibile salvare config: %s", self.config_path)

    # -- sessioni recenti --------------------------------------------------
    @property
    def recent_sessions(self) -> list[str]:
        return list(self._data.get("recent_sessions", []))

    def add_recent(self, path: str) -> None:
        path = os.path.abspath(path)
        recents = [p for p in self.recent_sessions if p != path]
        recents.insert(0, path)
        self._data["recent_sessions"] = recents[:_MAX_RECENT]
        self._save()

    def remove_recent(self, path: str) -> None:
        path = os.path.abspath(path)
        self._data["recent_sessions"] = [
            p for p in self.recent_sessions if p != path
        ]
        self._save()

    # -- template DOCX -----------------------------------------------------
    @property
    def docx_template(self) -> dict:
        defaults = self._defaults()["docx_template"]
        return {**defaults, **self._data.get("docx_template", {})}

    def set_docx_template(self, **kw) -> None:
        self._data.setdefault("docx_template", {}).update(kw)
        self._save()

    # -- sedi --------------------------------------------------------------
    @property
    def sites(self) -> list[str]:
        return list(self._data.get("sites", ["Messina", "Ganzirri"]))

    def set_sites(self, sites: list[str]) -> None:
        self._data["sites"] = list(sites)
        self._save()

    # -- peso storico ------------------------------------------------------
    @property
    def history_weight(self) -> float:
        return float(self._data.get("history_weight", 0.5))

    def set_history_weight(self, w: float) -> None:
        self._data["history_weight"] = max(0.0, min(1.0, w))
        self._save()

    # -- auto-save ---------------------------------------------------------
    @property
    def autosave_interval(self) -> int:
        return int(self._data.get("autosave_interval_seconds", 120))

    def save_autosave(self, session_data: dict) -> None:
        os.makedirs(self.directory, exist_ok=True)
        try:
            _write_text_file_atomic(
                self.autosave_path,
                json.dumps(session_data, ensure_ascii=False, indent=2),
            )
        except OSError:
            logger.warning("Auto-save fallito: %s", self.autosave_path)

    def load_autosave(self) -> dict | None:
        if not os.path.exists(self.autosave_path):
            return None
        try:
            with open(self.autosave_path, encoding="utf-8") as fh:
                d = json.load(fh)
            return d if isinstance(d, dict) else None
        except (OSError, json.JSONDecodeError):
            return None

    def clear_autosave(self) -> None:
        try:
            os.unlink(self.autosave_path)
        except FileNotFoundError:
            pass
        except OSError:
            logger.warning("Impossibile rimuovere auto-save")
