"""Applicazione principale Tkinter: wizard a 3 step per la gestione turni."""
from __future__ import annotations
import datetime
import json
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from turni.constants import (
    BG, PANEL_BG, ACCENT, SUCCESS, DANGER, WARN, MUTED, TEXT, BORDER,
    HEADER_BG, MONTH_HDR_BG, ACCENT_LIGHT, BTN_SECONDARY, DANGER_LIGHT,
    HEADER_BTN, INFO_WARN, CYAN, SOLVER_TIMEOUT,
    FONT_TITLE, FONT_LABEL, FONT_BOLD, FONT_SMALL, FONT_MONO,
)
from turni.helpers import (
    normalize_name,
    _order_weeks_by_declared_months,
    _group_weeks_by_normalized_month,
)
from turni.validators import (
    SessionValidationError,
    _parse_step0_inputs,
    _validate_week_entries,
    _validate_solver_ready_weeks,
    _validate_session_payload,
)
from turni.io_utils import _write_text_file_atomic
from turni.solver import TurniSolver, ORTOOLS_OK
from turni.docx_export import build_turni_docx
from turni.ui.widgets import configure_ttk_style, styled_btn, card, ScrollableFrame

logger = logging.getLogger(__name__)


class TurniApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Turni")
        self.geometry("880x700")
        self.minsize(760, 580)
        self.configure(bg=BG)
        self.resizable(True, True)
        configure_ttk_style()

        # Dati sessione
        self.operatori       = []
        self.operatori_norm  = []
        self.anno            = tk.StringVar(value=str(datetime.date.today().year))
        self.mesi            = []
        self.weeks_data      = []
        self.solver_timeout  = tk.StringVar(value=str(int(SOLVER_TIMEOUT)))

        # Inizializzazione garantita
        self._last_result_text = ""
        self._last_solver      = None

        # Controllo stato calcolo
        self._solving       = False
        self._cancel_event  = threading.Event()
        self._pending_close   = False
        self._pending_restart = False

        # Stato wizard
        self.current_step = 0
        self.steps        = ["Operatori & Anno", "Settimane", "Risultati"]

        # Per evitare ricostruzione inutile di step1
        self._step1_built  = False
        self._prev_ops:  list[str] = []
        self._prev_mesi: list[str] = []
        self._week_rows: list[dict]= []
        self._month_frames: dict[str, tk.Frame] = {}
        self._btn_save_session = None
        self._btn_load_session = None
        self._btn_modify = None
        self._btn_save_txt = None
        self._btn_save_csv = None
        self._btn_restart = None
        self._cancel_btn = None

        self._build_ui()
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ─── Layout ───────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_stepbar()
        self._build_content()

    def _build_header(self):
        hdr = tk.Frame(self, bg=HEADER_BG, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  Gestione Turni",
                 bg=HEADER_BG, fg="white",
                 font=FONT_TITLE).pack(side="left", padx=12, pady=10)
        mb = tk.Frame(hdr, bg=HEADER_BG)
        mb.pack(side="right", padx=12)
        self._btn_save_session = styled_btn(mb, "Salva sessione",  self._save_session,
                                            bg=HEADER_BTN)
        self._btn_save_session.pack(side="left", padx=4)
        self._btn_load_session = styled_btn(mb, "Carica sessione", self._load_session,
                                            bg=HEADER_BTN)
        self._btn_load_session.pack(side="left", padx=4)
        if not ORTOOLS_OK:
            tk.Label(hdr, text="ATTENZIONE: ortools mancante!",
                     bg=HEADER_BG, fg=INFO_WARN,
                     font=FONT_SMALL).pack(side="right", padx=8)

    def _build_stepbar(self):
        bar = tk.Frame(self, bg=PANEL_BG, height=46,
                       highlightbackground=BORDER, highlightthickness=1)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        self._step_labels = []
        for i, name in enumerate(self.steps):
            lbl = tk.Label(bar, text=f"  {i+1}. {name}  ",
                           bg=PANEL_BG, fg=MUTED, font=FONT_SMALL,
                           padx=8, pady=12)
            lbl.pack(side="left")
            self._step_labels.append(lbl)
            if i < len(self.steps)-1:
                from turni.constants import _FONT_FAMILY
                tk.Label(bar, text=">", bg=PANEL_BG,
                         fg=BORDER, font=(_FONT_FAMILY, 14)).pack(side="left")

    def _build_content(self):
        self.content = tk.Frame(self, bg=BG)
        self.content.pack(fill="both", expand=True)
        self.frames = {}
        for i in range(len(self.steps)):
            f = tk.Frame(self.content, bg=BG)
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.frames[i] = f
        self._build_step0(self.frames[0])
        self._build_step1(self.frames[1])
        self._build_step2(self.frames[2])

    def _show_step(self, n):
        self.current_step = n
        for i, lbl in enumerate(self._step_labels):
            if i == n:
                lbl.config(fg=ACCENT, font=FONT_BOLD, bg=MONTH_HDR_BG)
            elif i < n:
                lbl.config(fg=SUCCESS, font=FONT_SMALL, bg=PANEL_BG)
            else:
                lbl.config(fg=MUTED, font=FONT_SMALL, bg=PANEL_BG)
        self.frames[n].lift()

    def _set_solver_ui_state(self, solving: bool):
        state = "disabled" if solving else "normal"
        for btn in (
            self._btn_save_session,
            self._btn_load_session,
            self._btn_modify,
            self._btn_save_txt,
            self._btn_save_csv,
            self._btn_restart,
        ):
            if btn is not None:
                btn.config(state=state)
        if self._cancel_btn is not None:
            self._cancel_btn.config(state="normal" if solving else "disabled")

    def _request_cancel(self):
        if not self._solving:
            return
        self._cancel_event.set()
        cp_solver = getattr(self._last_solver, "cp_solver", None)
        for method_name in ("stop_search", "StopSearch"):
            stop_method = getattr(cp_solver, method_name, None)
            if callable(stop_method):
                try:
                    stop_method()
                except Exception:
                    logger.exception("Impossibile interrompere il solver con %s", method_name)
                break
        if self._cancel_btn is not None:
            self._cancel_btn.config(state="disabled")

    def _perform_restart(self):
        self.weeks_data        = []
        self._last_solver      = None
        self._last_result_text = ""
        self._pending_restart  = False
        self._result_subtitle.config(text="", fg=MUTED)
        self._set_result("")
        if self.mesi:
            self._build_step1_content(add_default_row=False)
            self._step1_built = True
        else:
            self._step1_built = False
            self._week_rows = []
            self._month_frames = {}
        self._show_step(0)

    def _on_close(self):
        """Chiusura sicura anche durante il calcolo."""
        if self._solving:
            if not messagebox.askyesno(
                    "Calcolo in corso",
                    "Il calcolo e' ancora in esecuzione.\n"
                    "Vuoi annullarlo e uscire?", parent=self):
                return
            self._pending_close = True
            self._request_cancel()
            self._result_subtitle.config(text="Annullamento in corso...", fg=WARN)
            self._set_result("Annullamento in corso. La finestra verra' chiusa al termine del solver.\n")
            self._set_solver_ui_state(True)
            if self._cancel_btn is not None:
                self._cancel_btn.config(state="disabled")
            return
        self.destroy()

    # ═══ STEP 0 – Operatori & Anno ════════════════════════════
    def _build_step0(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(expand=True)

        tk.Label(wrap, text="Operatori e Anno",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(28, 4))
        tk.Label(wrap, text="Inserisci i dati base per pianificare i turni.",
                 bg=BG, fg=MUTED, font=FONT_LABEL).pack(pady=(0, 16))

        c = card(wrap)
        c.pack(padx=40, pady=8, fill="x")
        inner = tk.Frame(c, bg=PANEL_BG)
        inner.pack(padx=24, pady=20, fill="x")
        inner.columnconfigure(1, weight=1)

        def lbl(row, text):
            tk.Label(inner, text=text, bg=PANEL_BG,
                     fg=TEXT, font=FONT_BOLD).grid(
                row=row, column=0, sticky="nw", pady=8)

        def entry(row, var=None, width=44):
            e = tk.Entry(inner, font=FONT_LABEL, width=width,
                         relief="solid",
                         highlightbackground=BORDER, highlightthickness=1,
                         textvariable=var)
            e.grid(row=row, column=1, sticky="ew", padx=12, pady=8)
            return e

        lbl(0, "Anno")
        self._anno_entry = entry(0, var=self.anno, width=12)

        lbl(1, "Mesi (separati da virgola)")
        self._mesi_entry = entry(1)
        self._mesi_entry.insert(0, "Gennaio, Febbraio")

        lbl(2, "Operatori\n(uno per riga o virgola)")
        tf = tk.Frame(inner, bg=PANEL_BG)
        tf.grid(row=2, column=1, sticky="ew", padx=12, pady=8)
        self._ops_text = tk.Text(tf, font=FONT_LABEL, width=44, height=5,
                                  relief="solid", wrap="word",
                                  highlightbackground=BORDER,
                                  highlightthickness=1)
        self._ops_text.pack(side="left", fill="both", expand=True)
        _ops_sb = ttk.Scrollbar(tf, command=self._ops_text.yview)
        self._ops_text.configure(yscrollcommand=_ops_sb.set)
        _ops_sb.pack(side="right", fill="y")
        self._ops_text.insert("1.0",
            "Mario Rossi\nLuca Bianchi\nAnna Verdi\nGiulia Neri")

        lbl(3, "Timeout solver (sec)")
        timeout_frame = tk.Frame(inner, bg=PANEL_BG)
        timeout_frame.grid(row=3, column=1, sticky="w", padx=12, pady=8)
        tk.Entry(timeout_frame, textvariable=self.solver_timeout,
                 font=FONT_LABEL, width=8, relief="solid",
                 highlightbackground=BORDER, highlightthickness=1).pack(side="left")
        tk.Label(timeout_frame,
                 text="  (aumenta se il solver non trova soluzioni in tempo)",
                 bg=PANEL_BG, fg=MUTED, font=FONT_SMALL).pack(side="left")

        styled_btn(wrap, "Avanti  >", self._step0_next).pack(pady=20)

    def _step0_next(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso",
                "Attendi il termine del calcolo prima di modificare i dati.",
                parent=self)
            return

        try:
            anno, timeout_str, new_ops, new_norm, new_mesi, dups, dup_mesi = _parse_step0_inputs(
                self.anno.get(),
                self._mesi_entry.get(),
                self._ops_text.get("1.0", "end"),
                self.solver_timeout.get(),
            )
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return

        self.anno.set(anno)
        self.solver_timeout.set(timeout_str)

        if dups:
            messagebox.showwarning("Attenzione",
                "Operatori duplicati rimossi:\n" + "\n".join(dups),
                parent=self)

        if dup_mesi:
            messagebox.showwarning("Attenzione",
                "Mesi duplicati rimossi:\n" + "\n".join(dup_mesi),
                parent=self)

        ops_ok  = new_ops  == self._prev_ops
        mesi_ok = new_mesi == self._prev_mesi

        if self._step1_built and ops_ok and mesi_ok:
            self.operatori = new_ops
            self.operatori_norm = new_norm
            self.mesi = new_mesi
            self._show_step(1)
            return

        if self._step1_built and self._week_rows and not (ops_ok and mesi_ok):
            if not messagebox.askyesno(
                    "Attenzione",
                    "Hai modificato operatori o mesi.\n"
                    "Le settimane inserite verranno azzerate. Continuare?",
                    parent=self):
                return

        self.operatori      = new_ops
        self.operatori_norm = new_norm
        self.mesi           = new_mesi
        self._prev_ops      = list(new_ops)
        self._prev_mesi     = list(new_mesi)
        self.weeks_data     = []

        self._build_step1_content()
        self._step1_built = True
        self._show_step(1)

    # ═══ STEP 1 – Settimane ═══════════════════════════════════
    def _build_step1(self, parent):
        self._step1_parent = parent

    def _build_step1_content(self, add_default_row: bool = True):
        for w in self._step1_parent.winfo_children():
            w.destroy()
        self._week_rows    = []
        self._month_frames = {}

        tk.Label(self._step1_parent, text="Pianificazione Settimane",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(20, 2))
        tk.Label(self._step1_parent,
                 text="Aggiungi settimane e spunta gli operatori non disponibili.",
                 bg=BG, fg=MUTED, font=FONT_LABEL).pack(pady=(0, 8))

        sf = ScrollableFrame(self._step1_parent)
        sf.pack(fill="both", expand=True, padx=20)

        for mese in self.mesi:
            mc = card(sf.inner)
            mc.pack(fill="x", pady=6, padx=4)
            hdr = tk.Frame(mc, bg=MONTH_HDR_BG)
            hdr.pack(fill="x")
            tk.Label(hdr, text=f"  {mese}",
                     bg=MONTH_HDR_BG, fg=ACCENT,
                     font=FONT_BOLD, anchor="w").pack(
                side="left", padx=10, pady=8)
            styled_btn(hdr, "+ Aggiungi settimana",
                       lambda m=mese: self._add_week_row(m),
                       bg=ACCENT_LIGHT, fg=ACCENT).pack(
                side="right", padx=10, pady=6)
            rows_frame = tk.Frame(mc, bg=PANEL_BG)
            rows_frame.pack(fill="x", padx=8, pady=4)
            self._month_frames[mese] = rows_frame
            if add_default_row:
                self._add_week_row(mese)

        nav = tk.Frame(self._step1_parent, bg=BG)
        nav.pack(pady=14)
        styled_btn(nav, "<  Indietro",
                   lambda: self._show_step(0),
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=8)
        styled_btn(nav, "Genera Turni  >",
                   self._step1_next, bg=SUCCESS).pack(side="left", padx=8)

    def _add_week_row(self, mese):
        rows_frame = self._month_frames[mese]
        row_idx    = sum(1 for r in self._week_rows if r["month"] == mese)

        row = tk.Frame(rows_frame, bg=PANEL_BG,
                       highlightbackground=BTN_SECONDARY, highlightthickness=1)
        row.pack(fill="x", pady=3, padx=4)

        lf = tk.Frame(row, bg=PANEL_BG)
        lf.pack(side="left", padx=8, pady=6)
        num_lbl = tk.Label(lf, text=f"Settimana {row_idx+1}:",
                           bg=PANEL_BG, fg=TEXT, font=FONT_BOLD)
        num_lbl.pack(anchor="w")
        name_var = tk.StringVar(value=f"Sett.{row_idx+1} {mese[:3]}")
        tk.Entry(lf, textvariable=name_var, font=FONT_LABEL,
                 width=22, relief="solid",
                 highlightbackground=BORDER, highlightthickness=1).pack(pady=2)

        bf = tk.Frame(row, bg=PANEL_BG)
        bf.pack(side="left", padx=16, pady=6, fill="x", expand=True)
        tk.Label(bf, text="Operatori busy:",
                 bg=PANEL_BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w")
        chk = tk.Frame(bf, bg=PANEL_BG)
        chk.pack(anchor="w")
        busy_vars = {}
        cols = 4
        for idx, op in enumerate(self.operatori):
            var = tk.BooleanVar(value=False)
            busy_vars[idx] = var
            tk.Checkbutton(chk, text=op, variable=var,
                           bg=PANEL_BG, fg=TEXT, font=FONT_SMALL,
                           activebackground=PANEL_BG,
                           selectcolor=ACCENT_LIGHT).grid(
                row=idx//cols, column=idx%cols,
                sticky="w", padx=8, pady=1)

        rd = {
            "month":      mese,
            "name_var":   name_var,
            "busy_vars":  busy_vars,
            "row_widget": row,
            "num_label":  num_lbl,
        }
        self._week_rows.append(rd)
        styled_btn(row, "X",
                   lambda r=rd, rw=row: self._remove_week_row(r, rw),
                   bg=DANGER_LIGHT, fg=DANGER).pack(side="right", padx=8, pady=6)

    def _remove_week_row(self, rd, rw):
        mese = rd["month"]
        rw.destroy()
        if rd in self._week_rows:
            self._week_rows.remove(rd)
        self._renumber_weeks(mese)

    def _renumber_weeks(self, mese):
        for i, rd in enumerate(r for r in self._week_rows if r["month"] == mese):
            rd["num_label"].config(text=f"Settimana {i+1}:")

    def _step1_next(self):
        if self._solving:
            messagebox.showinfo("Calcolo in corso",
                "Il calcolo e' ancora in esecuzione.", parent=self)
            return
        if not self._week_rows:
            messagebox.showerror("Errore",
                "Aggiungi almeno una settimana.", parent=self)
            return

        raw_weeks = [
            {
                "month": rd["month"],
                "week": rd["name_var"].get(),
                "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()],
            }
            for rd in self._week_rows
        ]

        try:
            canonical_weeks = _validate_week_entries(
                raw_weeks,
                declared_months=self.mesi,
                error_prefix="Errore Step 1",
            )
            _validate_solver_ready_weeks(
                canonical_weeks,
                operator_count=len(self.operatori),
                error_prefix="Errore Step 1",
            )
        except SessionValidationError as e:
            msg = str(e)
            if msg.startswith("Errore Step 1: "):
                msg = msg[len("Errore Step 1: "):]
            messagebox.showerror("Errori", msg, parent=self)
            return

        new_weeks_data = []
        for week in canonical_weeks:
            busy_norm = [self.operatori_norm[i] for i in week["busy_indices"]]
            available = [
                i for i, op in enumerate(self.operatori_norm)
                if op not in busy_norm
            ]
            new_weeks_data.append({
                "month": week["month"],
                "week": week["week"],
                "busy": busy_norm,
                "available": available,
            })

        if not new_weeks_data:
            messagebox.showerror("Errore",
                "Nessuna settimana valida.", parent=self)
            return
        self.weeks_data = _order_weeks_by_declared_months(new_weeks_data, self.mesi)
        self._show_step(2)
        self._run_solver()

    # ═══ STEP 2 – Risultati ═══════════════════════════════════
    def _build_step2(self, parent):
        tk.Label(parent, text="Risultati",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(20, 2))
        self._result_subtitle = tk.Label(parent, text="",
                                          bg=BG, fg=MUTED, font=FONT_LABEL)
        self._result_subtitle.pack(pady=(0, 6))

        rc = card(parent)
        rc.pack(fill="both", expand=True, padx=20, pady=4)
        self._result_text = tk.Text(rc, font=FONT_MONO, wrap="none",
                                     bg=PANEL_BG, fg=TEXT, relief="flat",
                                     state="disabled", padx=12, pady=8)
        sb_y = ttk.Scrollbar(rc, command=self._result_text.yview)
        sb_x = ttk.Scrollbar(rc, orient="horizontal",
                              command=self._result_text.xview)
        self._result_text.configure(yscrollcommand=sb_y.set,
                                     xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self._result_text.pack(fill="both", expand=True)

        self._prog_frame = tk.Frame(parent, bg=BG)
        self._prog_frame.pack(pady=4)
        self._progress = ttk.Progressbar(
            self._prog_frame, mode="indeterminate", length=280)
        self._progress.pack(side="left", padx=6)
        self._cancel_btn = styled_btn(
            self._prog_frame, "Annulla", self._cancel_solving, bg=DANGER)
        self._cancel_btn.pack(side="left", padx=6)
        self._prog_frame.pack_forget()

        nav = tk.Frame(parent, bg=BG)
        nav.pack(pady=10)
        self._btn_modify = styled_btn(nav, "<  Modifica",
                                      lambda: self._show_step(1),
                                      bg=BTN_SECONDARY, fg=TEXT)
        self._btn_modify.pack(side="left", padx=6)
        self._btn_save_txt = styled_btn(nav, "Salva TXT",
                                        lambda: self._export("txt"),
                                        bg=SUCCESS)
        self._btn_save_txt.pack(side="left", padx=6)
        self._btn_save_csv = styled_btn(nav, "Salva CSV",
                                        lambda: self._export("csv"),
                                        bg=CYAN)
        self._btn_save_csv.pack(side="left", padx=6)
        self._btn_save_docx = styled_btn(nav, "Salva DOCX",
                                         lambda: self._export("docx"),
                                         bg="#2563EB")
        self._btn_save_docx.pack(side="left", padx=6)
        self._btn_restart = styled_btn(nav, "Nuovo calcolo",
                                       self._restart, bg=WARN)
        self._btn_restart.pack(side="left", padx=6)

    def _run_solver(self):
        """Snapshot immutabili passati al thread worker."""
        ops_snap = list(self.operatori)
        weeks_snap = [
            {**w, "available": list(w["available"]), "busy": list(w["busy"])}
            for w in self.weeks_data
        ]

        self._solving      = True
        self._cancel_event = threading.Event()
        self._pending_close = False
        self._pending_restart = False
        self._last_solver  = TurniSolver(ops_snap, weeks_snap)

        self._result_subtitle.config(text="Ottimizzazione in corso...", fg=MUTED)
        self._prog_frame.pack(pady=4)
        self._progress.start(10)
        self._set_solver_ui_state(True)
        self._set_result("Calcolo in corso, attendere...\n")

        ev = self._cancel_event
        try:
            timeout = float(self.solver_timeout.get())
            if timeout <= 0:
                raise ValueError("timeout non positivo")
        except ValueError:
            timeout = SOLVER_TIMEOUT
            self.solver_timeout.set(str(int(SOLVER_TIMEOUT)))
            logger.warning("Timeout non valido in _run_solver, usato default %.1f", SOLVER_TIMEOUT)
        solver_ref = self._last_solver

        def worker():
            try:
                txt = solver_ref.solve(ev, timeout=timeout)
            except Exception as exc:
                logger.exception("Errore imprevisto nel solver")
                txt = f"Errore imprevisto durante il calcolo:\n{exc}"
            try:
                self.after(0, lambda: self._display_result(txt))
            except tk.TclError:
                logger.debug("UI gia' chiusa: risultato solver ignorato")

        threading.Thread(target=worker, daemon=True).start()

    def _cancel_solving(self):
        self._request_cancel()
        self._result_subtitle.config(text="Annullamento in corso...", fg=WARN)
        self._set_result("Annullamento in corso, attendere...\n")

    def _display_result(self, text):
        if not self.winfo_exists():
            return
        self._solving = False
        self._progress.stop()
        self._prog_frame.pack_forget()
        self._set_solver_ui_state(False)
        self._last_result_text = text
        ok = self._last_solver is not None and self._last_solver.solved_ok
        cancelled = self._last_solver is not None and self._last_solver.cancelled
        partial = (self._last_solver is not None
                   and self._last_solver.partial_result_available)
        if ok:
            subtitle, color = "Pianificazione completata.", SUCCESS
        elif cancelled and partial:
            subtitle, color = "Calcolo annullato: mostrata l'ultima soluzione parziale trovata.", WARN
        elif cancelled:
            subtitle, color = "Calcolo annullato.", WARN
        else:
            subtitle, color = text.split("\n")[0], DANGER
        self._result_subtitle.config(text=subtitle, fg=color)
        self._set_result(text)
        if self._pending_close:
            self._pending_close = False
            self.destroy()
            return
        if self._pending_restart:
            self._perform_restart()

    def _set_result(self, text):
        self._result_text.config(state="normal")
        self._result_text.delete("1.0", "end")
        self._result_text.insert("1.0", text)
        self._result_text.config(state="disabled")

    def _restart(self):
        if self._solving:
            self._pending_restart = True
            self._request_cancel()
            self._result_subtitle.config(
                text="Annullamento in corso... verra' avviato un nuovo calcolo.",
                fg=WARN,
            )
            self._set_result(
                "Annullamento in corso. Il reset verra' eseguito al termine del solver.\n")
            return
        self._perform_restart()

    # ─── Export ──────────────────────────────────────────────
    def _export(self, fmt):
        solver_ok = self._last_solver is not None and self._last_solver.solved_ok
        if not solver_ok:
            messagebox.showinfo("Info",
                "Nessun risultato da esportare.", parent=self)
            return
        anno = self.anno.get()
        path = None
        if fmt == "txt":
            path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Testo", "*.txt"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.txt", parent=self)
            if path:
                try:
                    _write_text_file_atomic(path, self._last_result_text)
                except OSError as e:
                    messagebox.showerror("Errore salvataggio",
                        f"Impossibile scrivere il file:\n{e}", parent=self)
                    return
        elif fmt == "csv":
            path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.csv", parent=self)
            if path:
                try:
                    _write_text_file_atomic(
                        path,
                        self._last_solver.format_csv(anno),
                        newline="",
                    )
                except OSError as e:
                    messagebox.showerror("Errore salvataggio",
                        f"Impossibile scrivere il file:\n{e}", parent=self)
                    return
        elif fmt == "docx":
            path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Documento Word", "*.docx"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.docx", parent=self)
            if path:
                try:
                    build_turni_docx(
                        path,
                        anno,
                        list(self._last_solver.result_rows),
                        list(self._last_solver.operatori),
                        list(self._last_solver.counts),
                        self._last_solver.status_str,
                        self._last_solver.diff_val,
                        self._last_solver.penalty_val,
                    )
                except OSError as e:
                    messagebox.showerror("Errore salvataggio",
                        f"Impossibile scrivere il file:\n{e}", parent=self)
                    return
                except Exception as e:
                    logger.exception("Errore durante la creazione del DOCX")
                    messagebox.showerror("Errore esportazione DOCX",
                        f"Impossibile creare il documento Word:\n{e}", parent=self)
                    return
        if path:
            messagebox.showinfo("Salvato",
                f"File salvato:\n{path}", parent=self)

    # ─── Sessione JSON ────────────────────────────────────────
    def _save_session(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso",
                "Attendi il termine del calcolo prima di salvare la sessione.",
                parent=self)
            return

        try:
            anno, timeout_str, parsed_ops, _parsed_norms, parsed_mesi, dup_ops, dup_mesi = _parse_step0_inputs(
                self.anno.get(),
                self._mesi_entry.get(),
                self._ops_text.get("1.0", "end"),
                self.solver_timeout.get(),
            )
        except SessionValidationError as e:
            messagebox.showerror("Errore",
                f"Impossibile salvare la sessione:\n{e}", parent=self)
            return

        if dup_ops:
            messagebox.showwarning("Attenzione",
                "Operatori duplicati rimossi nel salvataggio:\n" + "\n".join(dup_ops),
                parent=self)
        if dup_mesi:
            messagebox.showwarning("Attenzione",
                "Mesi duplicati rimossi nel salvataggio:\n" + "\n".join(dup_mesi),
                parent=self)

        if self._step1_built and (parsed_ops != self.operatori or parsed_mesi != self.mesi):
            messagebox.showerror("Errore",
                "Hai modificato operatori o mesi nello Step 0 ma non hai ancora applicato "
                "le modifiche. Clicca 'Avanti' per confermarle prima di salvare la sessione.",
                parent=self)
            return

        raw_session = {
            "anno": anno,
            "operatori": parsed_ops,
            "mesi": parsed_mesi,
            "solver_timeout": timeout_str,
            "weeks": [
                {
                    "month": rd["month"],
                    "week": rd["name_var"].get(),
                    "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()],
                }
                for rd in self._week_rows
            ],
        }

        try:
            session = _validate_session_payload(raw_session)
            session["solver_timeout"] = timeout_str
        except SessionValidationError as e:
            messagebox.showerror("Errore",
                f"Impossibile salvare la sessione:\n{e}", parent=self)
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("Sessione", "*.json"), ("Tutti", "*.*")],
            initialfile=f"sessione_{session['anno']}.json",
            parent=self)
        if path:
            try:
                _write_text_file_atomic(
                    path,
                    json.dumps(session, ensure_ascii=False, indent=2),
                )
            except OSError as e:
                messagebox.showerror("Errore salvataggio",
                    f"Impossibile scrivere il file:\n{e}", parent=self)
                return
            messagebox.showinfo("Salvato",
                f"Sessione salvata:\n{path}", parent=self)

    def _load_session(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso",
                "Attendi il termine del calcolo prima di caricare una sessione.",
                parent=self)
            return
        path = filedialog.askopenfilename(
            filetypes=[("Sessione", "*.json"), ("Tutti", "*.*")],
            parent=self)
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                raw_session = json.load(f)
        except Exception as e:
            messagebox.showerror("Errore",
                f"Impossibile leggere il file:\n{e}", parent=self)
            return

        try:
            session = _validate_session_payload(raw_session)
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return

        ops = session["operatori"]
        mesi = session["mesi"]
        weeks = session["weeks"]
        anno_val = session["anno"]

        self._last_solver      = None
        self._last_result_text = ""

        self.anno.set(anno_val)
        if "solver_timeout" in session:
            self.solver_timeout.set(session["solver_timeout"])
        self._ops_text.delete("1.0", "end")
        self._ops_text.insert("1.0", "\n".join(ops))
        self._mesi_entry.delete(0, "end")
        self._mesi_entry.insert(0, ", ".join(mesi))

        self.operatori      = ops
        self.operatori_norm = [normalize_name(o) for o in ops]
        self.mesi           = mesi
        self._prev_ops      = list(self.operatori)
        self._prev_mesi     = list(self.mesi)

        self._build_step1_content(add_default_row=False)
        self._step1_built = True

        by_month = _group_weeks_by_normalized_month(weeks)

        for mese in self.mesi:
            for w in by_month.get(normalize_name(mese), []):
                self._add_week_row(mese)
                rd = self._week_rows[-1]
                rd["name_var"].set(w.get("week", ""))
                for idx in w.get("busy_indices", []):
                    if idx in rd["busy_vars"]:
                        rd["busy_vars"][idx].set(True)
            self._renumber_weeks(mese)

        self._show_step(1)
        messagebox.showinfo("Caricato",
            "Sessione caricata. Verifica i dati e clicca 'Genera Turni'.",
            parent=self)
