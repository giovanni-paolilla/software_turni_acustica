"""Costanti globali: solver, tema grafico e font."""
from __future__ import annotations
import platform


# ══════════════════════════════════════════════════════════════
#  COSTANTI SOLVER
#  Peso equita': penalizza fortemente la differenza max-min turni.
#  100x > peso_penalty garantisce che l'equita' abbia sempre precedenza
#  sulla preferenza di non riusare l'operatore sab+mart nella stessa sett.
# ══════════════════════════════════════════════════════════════
PESO_DIFF          = 100
PESO_PENALTY       = 1
PESO_HISTORY       = 50
SOLVER_TIMEOUT     = 60.0
LOCK_STALE_SECONDS = 24 * 60 * 60

ALL_ROLES: frozenset[str] = frozenset({"audio", "video", "sabato"})


# ══════════════════════════════════════════════════════════════
#  FONT  cross-platform: Segoe UI non esiste su macOS/Linux
# ══════════════════════════════════════════════════════════════
_SYS = platform.system()
_FONT_FAMILY = ("Segoe UI"    if _SYS == "Windows" else
                "SF Pro Text" if _SYS == "Darwin"  else
                "DejaVu Sans")


# ══════════════════════════════════════════════════════════════
#  TEMA
# ══════════════════════════════════════════════════════════════
BG           = "#F0F4F8"
PANEL_BG     = "#FFFFFF"
ACCENT       = "#2563EB"
SUCCESS      = "#16A34A"
DANGER       = "#DC2626"
WARN         = "#D97706"
MUTED        = "#64748B"
TEXT         = "#1E293B"
BORDER       = "#CBD5E1"
HEADER_BG    = "#1E3A5F"
MONTH_HDR_BG = "#EFF6FF"   # sfondo header sezione mese nello step 1
ACCENT_LIGHT = "#DBEAFE"   # sfondo pulsante "+ Aggiungi" e checkbox selezionato
BTN_SECONDARY= "#E2E8F0"   # pulsanti navigazione secondari (Indietro, Modifica)
DANGER_LIGHT = "#FEE2E2"   # sfondo pulsante "X" elimina settimana
HEADER_BTN   = "#1E4A7A"   # pulsanti Salva/Carica sessione nell'header
INFO_WARN    = "#FCA5A5"   # testo avviso ortools mancante nell'header
CYAN         = "#0891B2"   # pulsante "Salva CSV"

FONT_TITLE = (_FONT_FAMILY, 18, "bold")
FONT_LABEL = (_FONT_FAMILY, 10)
FONT_BOLD  = (_FONT_FAMILY, 10, "bold")
FONT_SMALL = (_FONT_FAMILY, 9)
FONT_MONO  = ("Courier New", 10)
