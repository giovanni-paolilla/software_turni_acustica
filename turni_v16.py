"""Gestione Turni v16 — entry point.

Struttura del package turni/:
  constants.py   costanti globali (solver, tema, font)
  helpers.py     funzioni pure di utilita'
  validators.py  SessionValidationError + logica di validazione
  io_utils.py    I/O atomico e locking distribuito
  docx_export.py generazione documento Word
  solver.py      SolvePhase + TurniSolver (CP-SAT)
  ui/
    widgets.py   ScrollableFrame, styled_btn, card
    app.py       TurniApp (wizard a 3 step)
"""
from __future__ import annotations
import logging
import tkinter as tk
from tkinter import messagebox

from turni.solver import ORTOOLS_OK
from turni.ui.app import TurniApp

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    if not ORTOOLS_OK:
        _r = tk.Tk()
        _r.withdraw()
        messagebox.showwarning(
            "Dipendenza mancante",
            "ortools non e' installato.\n\n"
            "Installa con:\n    pip install ortools")
        _r.destroy()

    TurniApp().mainloop()
