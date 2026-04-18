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
    HEADER_BTN, INFO_WARN, CYAN, SOLVER_TIMEOUT, ALL_ROLES,
    FONT_TITLE, FONT_LABEL, FONT_BOLD, FONT_SMALL, FONT_MONO,
    _FONT_FAMILY,
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
from turni.calendar_utils import generate_weeks_for_month
from turni.history import HistoryStore
from turni.config import UserConfig
from turni.ics_export import build_ics, format_whatsapp
from turni.ui.widgets import configure_ttk_style, styled_btn, card, ScrollableFrame

logger = logging.getLogger(__name__)


class TurniApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Turni")
        self.geometry("920x750")
        self.minsize(800, 620)
        self.configure(bg=BG)
        self.resizable(True, True)
        configure_ttk_style()

        # Infrastruttura persistente
        self._config  = UserConfig()
        self._history = HistoryStore()

        # Dati sessione
        self.operatori       = []
        self.operatori_norm  = []
        self.anno            = tk.StringVar(value=str(datetime.date.today().year))
        self.mesi            = []
        self.sites           = list(self._config.sites)
        self.weeks_data      = []
        self.solver_timeout  = tk.StringVar(value=str(int(SOLVER_TIMEOUT)))

        # Feature 4: ruoli operatore   {indice_op: {"audio","video","sabato"}}
        self.operator_roles:  dict[int, set[str]] = {}
        # Feature 5: pattern ricorrenti  {indice_op: [1,3]}  settimane del mese
        self.recurring_busy:  dict[int, list[int]] = {}
        # Feature 9: sedi per operatore  {indice_op: ["Messina","Ganzirri"]}
        self.operator_sites:  dict[int, list[str]] = {}

        # Feature 2: equita' storica
        self.use_history     = tk.BooleanVar(value=False)
        self.history_weight  = tk.DoubleVar(value=self._config.history_weight)

        # Stato risultati
        self._last_result_text = ""
        self._last_solver      = None
        self._result_rows_edit: list[dict] = []

        # Controllo stato calcolo
        self._solving       = False
        self._cancel_event  = threading.Event()
        self._pending_close   = False
        self._pending_restart = False

        # Stato wizard
        self.current_step = 0
        self.steps        = ["Operatori & Anno", "Settimane", "Risultati"]

        # Step 1 cache
        self._step1_built  = False
        self._prev_ops:  list[str] = []
        self._prev_mesi: list[str] = []
        self._week_rows: list[dict]= []
        self._month_frames: dict[str, tk.Frame] = {}

        # Button references
        self._btn_save_session = None
        self._btn_load_session = None
        self._btn_modify = None
        self._btn_save_txt = None
        self._btn_save_csv = None
        self._btn_save_docx = None
        self._btn_restart = None
        self._cancel_btn = None

        # Step 2 view mode
        self._view_mode = "text"   # "text" | "table"
        self._result_tree = None
        self._tree_frame = None
        self._text_frame = None
        self._edit_combo = None

        self._build_ui()
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Feature 7: auto-save e recovery
        self._autosave_id = None
        self.after(500, self._check_autosave_recovery)
        self._schedule_autosave()

    # ================================================================
    #  LAYOUT
    # ================================================================
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

        # Feature 7: menu sessioni recenti
        self._load_mb = tk.Menubutton(mb, text="Carica sessione",
                                      bg=HEADER_BTN, fg="white",
                                      font=FONT_BOLD, relief="flat",
                                      cursor="hand2", padx=14, pady=6)
        self._load_menu = tk.Menu(self._load_mb, tearoff=0)
        self._load_mb["menu"] = self._load_menu
        self._load_mb.pack(side="left", padx=4)
        self._btn_load_session = self._load_mb
        self._build_recent_menu()

        if not ORTOOLS_OK:
            tk.Label(hdr, text="ATTENZIONE: ortools mancante!",
                     bg=HEADER_BG, fg=INFO_WARN,
                     font=FONT_SMALL).pack(side="right", padx=8)

    def _build_recent_menu(self):
        menu = self._load_menu
        menu.delete(0, "end")
        for path in self._config.recent_sessions:
            name = os.path.basename(path)
            menu.add_command(label=name,
                             command=lambda p=path: self._load_session_from_path(p))
        if self._config.recent_sessions:
            menu.add_separator()
        menu.add_command(label="Sfoglia...", command=self._load_session)

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
            self._btn_save_docx,
            self._btn_restart,
        ):
            if btn is not None:
                try:
                    btn.config(state=state)
                except tk.TclError:
                    pass
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
        self._result_rows_edit = []
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
        self._config.clear_autosave()
        self.destroy()

    # ================================================================
    #  FEATURE 7: AUTO-SAVE
    # ================================================================
    def _schedule_autosave(self):
        interval_ms = self._config.autosave_interval * 1000
        self._autosave_id = self.after(interval_ms, self._do_autosave)

    def _do_autosave(self):
        if self._solving:
            self._schedule_autosave()
            return
        try:
            data = self._collect_session_data()
            if data:
                self._config.save_autosave(data)
        except Exception:
            logger.debug("Auto-save fallito", exc_info=True)
        self._schedule_autosave()

    def _check_autosave_recovery(self):
        recovery = self._config.load_autosave()
        if recovery is None:
            return
        try:
            _validate_session_payload(recovery)
        except SessionValidationError:
            self._config.clear_autosave()
            return
        if messagebox.askyesno(
                "Recupero sessione",
                "E' stata trovata una sessione non salvata.\n"
                "Vuoi ripristinarla?", parent=self):
            self._apply_session(recovery)
        self._config.clear_autosave()

    # ================================================================
    #  STEP 0 - Operatori & Anno
    # ================================================================
    def _build_step0(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(expand=True)

        tk.Label(wrap, text="Operatori e Anno",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(28, 4))
        tk.Label(wrap, text="Inserisci i dati base per pianificare i turni.",
                 bg=BG, fg=MUTED, font=FONT_LABEL).pack(pady=(0, 12))

        c = card(wrap)
        c.pack(padx=40, pady=8, fill="x")
        inner = tk.Frame(c, bg=PANEL_BG)
        inner.pack(padx=24, pady=16, fill="x")
        inner.columnconfigure(1, weight=1)

        def lbl(row, text):
            tk.Label(inner, text=text, bg=PANEL_BG,
                     fg=TEXT, font=FONT_BOLD).grid(
                row=row, column=0, sticky="nw", pady=6)

        def entry(row, var=None, width=44):
            e = tk.Entry(inner, font=FONT_LABEL, width=width,
                         relief="solid",
                         highlightbackground=BORDER, highlightthickness=1,
                         textvariable=var)
            e.grid(row=row, column=1, sticky="ew", padx=12, pady=6)
            return e

        lbl(0, "Anno")
        self._anno_entry = entry(0, var=self.anno, width=12)

        lbl(1, "Mesi (virgola)")
        self._mesi_entry = entry(1)
        self._mesi_entry.insert(0, "Gennaio, Febbraio")

        # Feature 9: sedi
        lbl(2, "Sedi (virgola)")
        self._sites_entry = entry(2)
        self._sites_entry.insert(0, ", ".join(self.sites))

        lbl(3, "Operatori\n(uno per riga)")
        tf = tk.Frame(inner, bg=PANEL_BG)
        tf.grid(row=3, column=1, sticky="ew", padx=12, pady=6)
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

        lbl(4, "Timeout solver (sec)")
        timeout_frame = tk.Frame(inner, bg=PANEL_BG)
        timeout_frame.grid(row=4, column=1, sticky="w", padx=12, pady=6)
        tk.Entry(timeout_frame, textvariable=self.solver_timeout,
                 font=FONT_LABEL, width=8, relief="solid",
                 highlightbackground=BORDER, highlightthickness=1).pack(side="left")
        tk.Label(timeout_frame,
                 text="  (aumenta se il solver non trova soluzioni in tempo)",
                 bg=PANEL_BG, fg=MUTED, font=FONT_SMALL).pack(side="left")

        # Feature 2: equita' storica
        hist_frame = tk.Frame(inner, bg=PANEL_BG)
        hist_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=6)
        tk.Checkbutton(hist_frame, text="Usa equita' storica",
                       variable=self.use_history, bg=PANEL_BG, fg=TEXT,
                       font=FONT_BOLD, activebackground=PANEL_BG,
                       selectcolor=ACCENT_LIGHT).pack(side="left")
        tk.Label(hist_frame, text="Peso:", bg=PANEL_BG, fg=MUTED,
                 font=FONT_SMALL).pack(side="left", padx=(16, 4))
        tk.Scale(hist_frame, from_=0, to=100, orient="horizontal",
                 variable=self.history_weight, bg=PANEL_BG, fg=TEXT,
                 highlightthickness=0, length=120,
                 font=FONT_SMALL).pack(side="left")
        tk.Label(hist_frame, text="%", bg=PANEL_BG, fg=MUTED,
                 font=FONT_SMALL).pack(side="left")

        # Feature 4 + 5: pulsanti config ruoli e pattern
        cfg_frame = tk.Frame(inner, bg=PANEL_BG)
        cfg_frame.grid(row=6, column=0, columnspan=2, sticky="w", padx=12, pady=6)
        styled_btn(cfg_frame, "Configura ruoli",
                   self._open_roles_dialog,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)
        styled_btn(cfg_frame, "Pattern indisponibilita'",
                   self._open_patterns_dialog,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)

        styled_btn(wrap, "Avanti  >", self._step0_next).pack(pady=16)

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

        # Parsa sedi
        raw_sites = [s.strip() for s in self._sites_entry.get().split(",") if s.strip()]
        self.sites = raw_sites if raw_sites else [""]

        if dups:
            messagebox.showwarning("Attenzione",
                "Operatori duplicati rimossi:\n" + "\n".join(dups), parent=self)
        if dup_mesi:
            messagebox.showwarning("Attenzione",
                "Mesi duplicati rimossi:\n" + "\n".join(dup_mesi), parent=self)

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

        # Reset ruoli per nuovi operatori
        self.operator_roles = {i: set(ALL_ROLES) for i in range(len(new_ops))}
        self.operator_sites = {i: list(self.sites) for i in range(len(new_ops))}

        self._build_step1_content()
        self._step1_built = True
        self._show_step(1)

    # ================================================================
    #  FEATURE 4: DIALOG RUOLI
    # ================================================================
    def _open_roles_dialog(self):
        if not self.operatori:
            ops_raw = self._ops_text.get("1.0", "end").strip()
            if not ops_raw:
                messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
                return
            from turni.helpers import _parse_operatori_text
            temp_ops = _parse_operatori_text(ops_raw)
            if not temp_ops:
                messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
                return
        else:
            temp_ops = self.operatori

        dlg = tk.Toplevel(self)
        dlg.title("Configura Ruoli Operatori")
        dlg.geometry("500x400")
        dlg.configure(bg=BG)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Ruoli per operatore", bg=BG, fg=TEXT,
                 font=FONT_TITLE).pack(pady=(12, 8))

        sf = ScrollableFrame(dlg)
        sf.pack(fill="both", expand=True, padx=16)

        role_vars: dict[int, dict[str, tk.BooleanVar]] = {}
        for i, op in enumerate(temp_ops):
            row = tk.Frame(sf.inner, bg=PANEL_BG,
                           highlightbackground=BORDER, highlightthickness=1)
            row.pack(fill="x", pady=2, padx=4)
            tk.Label(row, text=op, bg=PANEL_BG, fg=TEXT,
                     font=FONT_BOLD, width=20, anchor="w").pack(side="left", padx=8)
            role_vars[i] = {}
            current = self.operator_roles.get(i, ALL_ROLES)
            for role in ["audio", "video", "sabato"]:
                var = tk.BooleanVar(value=(role in current))
                role_vars[i][role] = var
                tk.Checkbutton(row, text=role.capitalize(), variable=var,
                               bg=PANEL_BG, fg=TEXT, font=FONT_SMALL,
                               activebackground=PANEL_BG,
                               selectcolor=ACCENT_LIGHT).pack(side="left", padx=8)

        def apply():
            for i in role_vars:
                roles = {r for r, v in role_vars[i].items() if v.get()}
                self.operator_roles[i] = roles if roles else set(ALL_ROLES)
            dlg.destroy()

        styled_btn(dlg, "Applica", apply, bg=SUCCESS).pack(pady=12)

    # ================================================================
    #  FEATURE 5: DIALOG PATTERN INDISPONIBILITA'
    # ================================================================
    def _open_patterns_dialog(self):
        if not self.operatori:
            ops_raw = self._ops_text.get("1.0", "end").strip()
            if not ops_raw:
                messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
                return
            from turni.helpers import _parse_operatori_text
            temp_ops = _parse_operatori_text(ops_raw)
            if not temp_ops:
                messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
                return
        else:
            temp_ops = self.operatori

        dlg = tk.Toplevel(self)
        dlg.title("Pattern Indisponibilita' Ricorrente")
        dlg.geometry("600x400")
        dlg.configure(bg=BG)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Settimane del mese sempre non disponibili",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(12, 4))
        tk.Label(dlg, text="Spunta le settimane in cui l'operatore e' regolarmente impegnato.",
                 bg=BG, fg=MUTED, font=FONT_SMALL).pack(pady=(0, 8))

        sf = ScrollableFrame(dlg)
        sf.pack(fill="both", expand=True, padx=16)

        pattern_vars: dict[int, dict[int, tk.BooleanVar]] = {}
        for i, op in enumerate(temp_ops):
            row = tk.Frame(sf.inner, bg=PANEL_BG,
                           highlightbackground=BORDER, highlightthickness=1)
            row.pack(fill="x", pady=2, padx=4)
            tk.Label(row, text=op, bg=PANEL_BG, fg=TEXT,
                     font=FONT_BOLD, width=18, anchor="w").pack(side="left", padx=8)
            pattern_vars[i] = {}
            current = self.recurring_busy.get(i, [])
            for wk in range(1, 6):
                var = tk.BooleanVar(value=(wk in current))
                pattern_vars[i][wk] = var
                tk.Checkbutton(row, text=f"{wk}a sett.", variable=var,
                               bg=PANEL_BG, fg=TEXT, font=FONT_SMALL,
                               activebackground=PANEL_BG,
                               selectcolor=ACCENT_LIGHT).pack(side="left", padx=4)

        def apply():
            for i in pattern_vars:
                weeks = [wk for wk, v in pattern_vars[i].items() if v.get()]
                if weeks:
                    self.recurring_busy[i] = weeks
                else:
                    self.recurring_busy.pop(i, None)
            dlg.destroy()

        styled_btn(dlg, "Applica", apply, bg=SUCCESS).pack(pady=12)

    # ================================================================
    #  STEP 1 - Settimane
    # ================================================================
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

        multi_site = len(self.sites) > 1 and self.sites[0]

        for mese in self.mesi:
            mc = card(sf.inner)
            mc.pack(fill="x", pady=6, padx=4)
            hdr = tk.Frame(mc, bg=MONTH_HDR_BG)
            hdr.pack(fill="x")
            tk.Label(hdr, text=f"  {mese}",
                     bg=MONTH_HDR_BG, fg=ACCENT,
                     font=FONT_BOLD, anchor="w").pack(
                side="left", padx=10, pady=8)
            # Feature 1: auto-genera settimane
            styled_btn(hdr, "Auto-genera",
                       lambda m=mese: self._auto_generate_weeks(m),
                       bg="#7C3AED", fg="white").pack(
                side="right", padx=4, pady=6)
            styled_btn(hdr, "+ Aggiungi settimana",
                       lambda m=mese: self._add_week_row(m),
                       bg=ACCENT_LIGHT, fg=ACCENT).pack(
                side="right", padx=4, pady=6)
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

    # Feature 1: auto-genera settimane dal calendario
    def _auto_generate_weeks(self, mese):
        try:
            year = int(self.anno.get())
        except ValueError:
            messagebox.showerror("Errore", "Anno non valido.", parent=self)
            return

        weeks = generate_weeks_for_month(year, mese)
        if not weeks:
            messagebox.showinfo("Info",
                f"Impossibile generare settimane per '{mese}' {year}.\n"
                "Verifica che il nome del mese sia in italiano.", parent=self)
            return

        existing = [r for r in self._week_rows if r["month"] == mese]
        if existing:
            if not messagebox.askyesno("Attenzione",
                    f"Ci sono gia' {len(existing)} settimane per {mese}.\n"
                    "Vuoi sostituirle?", parent=self):
                return
            for rd in list(existing):
                rd["row_widget"].destroy()
                self._week_rows.remove(rd)

        multi_site = len(self.sites) > 1 and self.sites[0]

        for wk in weeks:
            if multi_site:
                for site in self.sites:
                    self._add_week_row(mese)
                    rd = self._week_rows[-1]
                    rd["name_var"].set(wk["week"])
                    rd["site_var"].set(site)
                    rd["date_key"] = wk["thursday"]
                    rd["week_of_month"] = wk.get("week_of_month", 0)
                    # Feature 5: applica pattern ricorrenti
                    self._apply_recurring_busy(rd, wk.get("week_of_month", 0))
            else:
                self._add_week_row(mese)
                rd = self._week_rows[-1]
                rd["name_var"].set(wk["week"])
                rd["date_key"] = wk["thursday"]
                rd["week_of_month"] = wk.get("week_of_month", 0)
                self._apply_recurring_busy(rd, wk.get("week_of_month", 0))

        self._renumber_weeks(mese)

    def _apply_recurring_busy(self, rd, week_of_month):
        """Applica i pattern di indisponibilita' ricorrente a una riga settimana."""
        if not week_of_month:
            return
        for op_idx, weeks_pattern in self.recurring_busy.items():
            if week_of_month in weeks_pattern and op_idx in rd["busy_vars"]:
                rd["busy_vars"][op_idx].set(True)

    def _add_week_row(self, mese):
        rows_frame = self._month_frames[mese]
        row_idx    = sum(1 for r in self._week_rows if r["month"] == mese)
        multi_site = len(self.sites) > 1 and self.sites[0]

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

        # Feature 9: selezione sede
        site_var = tk.StringVar(value=self.sites[0] if self.sites else "")
        if multi_site:
            sf = tk.Frame(lf, bg=PANEL_BG)
            sf.pack(anchor="w", pady=2)
            tk.Label(sf, text="Sede:", bg=PANEL_BG, fg=MUTED,
                     font=FONT_SMALL).pack(side="left")
            ttk.Combobox(sf, textvariable=site_var, values=self.sites,
                         state="readonly", width=14,
                         font=FONT_SMALL).pack(side="left", padx=4)

        # Feature 6: blocco settimana
        lock_var = tk.BooleanVar(value=False)
        tk.Checkbutton(lf, text="Blocca", variable=lock_var,
                       bg=PANEL_BG, fg=MUTED, font=FONT_SMALL,
                       activebackground=PANEL_BG,
                       selectcolor=ACCENT_LIGHT).pack(anchor="w")

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
            "site_var":   site_var,
            "lock_var":   lock_var,
            "row_widget": row,
            "num_label":  num_lbl,
            "date_key":   "",
            "week_of_month": 0,
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

        raw_weeks = []
        for rd in self._week_rows:
            entry = {
                "month": rd["month"],
                "week": rd["name_var"].get(),
                "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()],
            }
            site = rd["site_var"].get()
            if site:
                entry["site"] = site
            if rd["date_key"]:
                entry["date_key"] = rd["date_key"]
            if rd["lock_var"].get():
                entry["locked"] = True
            raw_weeks.append(entry)

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
            wd: dict = {
                "month": week["month"],
                "week": week["week"],
                "busy": busy_norm,
                "available": available,
            }
            if week.get("site"):
                wd["site"] = week["site"]
            if week.get("date_key"):
                wd["date_key"] = week["date_key"]
            # Feature 6: settimana bloccata con assegnazione precedente
            if week.get("locked") and week.get("locked_assignment"):
                wd["locked"] = True
                wd["locked_assignment"] = week["locked_assignment"]
            new_weeks_data.append(wd)

        if not new_weeks_data:
            messagebox.showerror("Errore",
                "Nessuna settimana valida.", parent=self)
            return
        self.weeks_data = _order_weeks_by_declared_months(new_weeks_data, self.mesi)
        self._show_step(2)
        self._run_solver()

    # ================================================================
    #  STEP 2 - Risultati
    # ================================================================
    def _build_step2(self, parent):
        tk.Label(parent, text="Risultati",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(16, 2))
        self._result_subtitle = tk.Label(parent, text="",
                                          bg=BG, fg=MUTED, font=FONT_LABEL)
        self._result_subtitle.pack(pady=(0, 4))

        # Feature 3/8: toggle testo/tabella
        toggle_frame = tk.Frame(parent, bg=BG)
        toggle_frame.pack(pady=2)
        self._btn_view_text = styled_btn(toggle_frame, "Testo",
                                         lambda: self._toggle_view("text"),
                                         bg=ACCENT)
        self._btn_view_text.pack(side="left", padx=2)
        self._btn_view_table = styled_btn(toggle_frame, "Tabella",
                                          lambda: self._toggle_view("table"),
                                          bg=BTN_SECONDARY, fg=TEXT)
        self._btn_view_table.pack(side="left", padx=2)

        # Contenitore per text view
        self._text_frame = card(parent)
        self._text_frame.pack(fill="both", expand=True, padx=20, pady=4)
        self._result_text = tk.Text(self._text_frame, font=FONT_MONO, wrap="none",
                                     bg=PANEL_BG, fg=TEXT, relief="flat",
                                     state="disabled", padx=12, pady=8)
        sb_y = ttk.Scrollbar(self._text_frame, command=self._result_text.yview)
        sb_x = ttk.Scrollbar(self._text_frame, orient="horizontal",
                              command=self._result_text.xview)
        self._result_text.configure(yscrollcommand=sb_y.set,
                                     xscrollcommand=sb_x.set)
        sb_y.pack(side="right", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self._result_text.pack(fill="both", expand=True)

        # Feature 3: contenitore per table view (nascosto di default)
        self._tree_frame = card(parent)
        cols = ("settimana", "audio", "video", "sabato", "sede", "stato")
        self._result_tree = ttk.Treeview(self._tree_frame, columns=cols,
                                          show="headings", height=14)
        for col, heading, w in [
            ("settimana", "Settimana", 180),
            ("audio", "Audio", 130),
            ("video", "Video", 130),
            ("sabato", "Domenica A/V", 130),
            ("sede", "Sede", 100),
            ("stato", "Stato", 60),
        ]:
            self._result_tree.heading(col, text=heading)
            self._result_tree.column(col, width=w, anchor="center")
        tree_sb = ttk.Scrollbar(self._tree_frame, command=self._result_tree.yview)
        self._result_tree.configure(yscrollcommand=tree_sb.set)
        tree_sb.pack(side="right", fill="y")
        self._result_tree.pack(fill="both", expand=True)
        self._result_tree.bind("<Double-1>", self._on_table_double_click)

        self._prog_frame = tk.Frame(parent, bg=BG)
        self._prog_frame.pack(pady=4)
        self._progress = ttk.Progressbar(
            self._prog_frame, mode="indeterminate", length=280)
        self._progress.pack(side="left", padx=6)
        self._cancel_btn = styled_btn(
            self._prog_frame, "Annulla", self._cancel_solving, bg=DANGER)
        self._cancel_btn.pack(side="left", padx=6)
        self._prog_frame.pack_forget()

        # Riga 1: pulsanti principali
        nav = tk.Frame(parent, bg=BG)
        nav.pack(pady=4)
        self._btn_modify = styled_btn(nav, "<  Modifica",
                                      lambda: self._show_step(1),
                                      bg=BTN_SECONDARY, fg=TEXT)
        self._btn_modify.pack(side="left", padx=4)
        self._btn_save_txt = styled_btn(nav, "Salva TXT",
                                        lambda: self._export("txt"),
                                        bg=SUCCESS)
        self._btn_save_txt.pack(side="left", padx=4)
        self._btn_save_csv = styled_btn(nav, "Salva CSV",
                                        lambda: self._export("csv"),
                                        bg=CYAN)
        self._btn_save_csv.pack(side="left", padx=4)
        self._btn_save_docx = styled_btn(nav, "Salva DOCX",
                                         lambda: self._export("docx"),
                                         bg="#2563EB")
        self._btn_save_docx.pack(side="left", padx=4)
        self._btn_restart = styled_btn(nav, "Nuovo calcolo",
                                       self._restart, bg=WARN)
        self._btn_restart.pack(side="left", padx=4)

        # Riga 2: pulsanti secondari (Feature 2, 10)
        nav2 = tk.Frame(parent, bg=BG)
        nav2.pack(pady=(0, 8))
        styled_btn(nav2, "Salva ICS",
                   lambda: self._export("ics"),
                   bg="#7C3AED").pack(side="left", padx=4)
        styled_btn(nav2, "Copia",
                   self._copy_to_clipboard,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)
        styled_btn(nav2, "Registra storico",
                   self._record_history,
                   bg="#0D9488").pack(side="left", padx=4)
        styled_btn(nav2, "Personalizza DOCX",
                   self._open_docx_config,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)

    # -- View toggle -------------------------------------------------------
    def _toggle_view(self, mode: str):
        self._view_mode = mode
        if mode == "text":
            self._tree_frame.pack_forget()
            self._text_frame.pack(fill="both", expand=True, padx=20, pady=4,
                                  after=self._result_subtitle)
            self._btn_view_text.config(bg=ACCENT, fg="white")
            self._btn_view_table.config(bg=BTN_SECONDARY, fg=TEXT)
        else:
            self._text_frame.pack_forget()
            self._tree_frame.pack(fill="both", expand=True, padx=20, pady=4,
                                  after=self._result_subtitle)
            self._btn_view_text.config(bg=BTN_SECONDARY, fg=TEXT)
            self._btn_view_table.config(bg=ACCENT, fg="white")
            self._populate_result_tree()

    def _populate_result_tree(self):
        tree = self._result_tree
        tree.delete(*tree.get_children())
        for i, r in enumerate(self._result_rows_edit):
            locked = r.get("locked", False)
            stato = "BLK" if locked else "OK"
            tree.insert("", "end", iid=str(i), values=(
                r.get("week", ""),
                r.get("audio", ""),
                r.get("video", ""),
                r.get("sabato", ""),
                r.get("site", ""),
                stato,
            ))

    # Feature 3: modifica manuale post-solver
    def _on_table_double_click(self, event):
        tree = self._result_tree
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item or not column:
            return

        col_idx = int(column[1:]) - 1
        col_names = ["settimana", "audio", "video", "sabato", "sede", "stato"]
        if col_idx not in (1, 2, 3):
            return

        row_idx = int(item)
        row_data = self._result_rows_edit[row_idx]
        if row_data.get("locked"):
            return

        # Determina operatori disponibili
        if row_idx < len(self.weeks_data):
            avail_indices = self.weeks_data[row_idx].get("available", [])
            avail_names = [self.operatori[i] for i in avail_indices
                           if 0 <= i < len(self.operatori)]
        else:
            avail_names = list(self.operatori)

        bbox = tree.bbox(item, column)
        if not bbox:
            return
        x, y, w, h = bbox

        if self._edit_combo is not None:
            self._edit_combo.destroy()

        combo = ttk.Combobox(tree, values=avail_names, state="readonly",
                             font=FONT_SMALL)
        current_val = tree.set(item, column)
        combo.set(current_val)
        combo.place(x=x, y=y, width=w, height=h)
        combo.focus_set()
        self._edit_combo = combo

        col_key = col_names[col_idx]

        def on_select(e=None):
            new_val = combo.get()
            if new_val:
                tree.set(item, column, new_val)
                self._result_rows_edit[row_idx][col_key] = new_val
                self._validate_table_row(row_idx)
                # Aggiorna testo risultati
                if self._last_solver:
                    self._last_solver.result_rows = list(self._result_rows_edit)
                    self._last_result_text = self._last_solver._format_text()
                    self._set_result(self._last_result_text)
            combo.destroy()
            self._edit_combo = None

        combo.bind("<<ComboboxSelected>>", on_select)
        combo.bind("<FocusOut>", lambda e: (combo.destroy(), setattr(self, '_edit_combo', None)))
        combo.bind("<Escape>", lambda e: (combo.destroy(), setattr(self, '_edit_combo', None)))

    def _validate_table_row(self, row_idx):
        row = self._result_rows_edit[row_idx]
        audio = row.get("audio", "")
        video = row.get("video", "")
        tree = self._result_tree
        item = str(row_idx)
        if audio == video and audio:
            tree.set(item, "stato", "!!")
        else:
            tree.set(item, "stato", "OK")

    # -- Solver execution --------------------------------------------------
    def _run_solver(self):
        ops_snap = list(self.operatori)
        weeks_snap = [
            {**w, "available": list(w["available"]), "busy": list(w["busy"])}
            for w in self.weeks_data
        ]

        # Prepara parametri feature
        roles = dict(self.operator_roles) if self.operator_roles else None
        hist_counts = None
        hist_weight = 0.0
        if self.use_history.get():
            hist_counts = self._history.get_cumulative_counts(ops_snap)
            hist_weight = self.history_weight.get() / 100.0

        self._solving      = True
        self._cancel_event = threading.Event()
        self._pending_close = False
        self._pending_restart = False
        self._last_solver  = TurniSolver(
            ops_snap, weeks_snap,
            operator_roles=roles,
            historical_counts=hist_counts,
            history_weight=hist_weight,
        )

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

        _s = self._last_solver
        ok = _s is not None and _s.solved_ok
        cancelled = _s is not None and _s.cancelled
        partial = _s is not None and _s.partial_result_available

        if ok:
            subtitle, color = "Pianificazione completata.", SUCCESS
        elif cancelled and partial:
            subtitle, color = "Calcolo annullato: mostrata soluzione parziale.", WARN
        elif cancelled:
            subtitle, color = "Calcolo annullato.", WARN
        else:
            subtitle, color = text.split("\n")[0], DANGER

        self._result_subtitle.config(text=subtitle, fg=color)
        self._set_result(text)

        # Copia editabile dei risultati
        if _s and _s.result_rows:
            self._result_rows_edit = [dict(r) for r in _s.result_rows]
            # Salva locked_assignment nei risultati
            for i, r in enumerate(self._result_rows_edit):
                if i < len(self.weeks_data):
                    wd = self.weeks_data[i]
                    if wd.get("locked") and wd.get("locked_assignment"):
                        r["locked"] = True
                        r["locked_assignment"] = wd["locked_assignment"]
        else:
            self._result_rows_edit = []

        if self._view_mode == "table":
            self._populate_result_tree()

        if self._pending_close:
            self._pending_close = False
            self._config.clear_autosave()
            self.destroy()
            return
        if self._pending_restart:
            self._perform_restart()

    def _set_result(self, text):
        self._result_text.config(state="normal")
        self._result_text.delete("1.0", "end")
        self._result_text.insert("1.0", text)
        self._result_text.config(state="disabled")

    # Feature 10: copia negli appunti
    def _copy_to_clipboard(self):
        _s = self._last_solver
        if not _s or not (_s.solved_ok or _s.partial_result_available):
            messagebox.showinfo("Info", "Nessun risultato da copiare.", parent=self)
            return
        text = format_whatsapp(
            self.anno.get(), self._result_rows_edit,
            _s.operatori, _s.counts)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("Copiato",
            "Risultati copiati negli appunti (formato WhatsApp).", parent=self)

    # Feature 2: registra nello storico
    def _record_history(self):
        _s = self._last_solver
        if not _s or not _s.solved_ok:
            messagebox.showinfo("Info",
                "Registra nello storico solo dopo una soluzione completa.", parent=self)
            return
        if not messagebox.askyesno("Conferma",
                "Vuoi registrare questi turni nello storico?\n"
                "I conteggi verranno sommati ai precedenti per l'equita' futura.",
                parent=self):
            return
        self._history.record_session(
            self.anno.get(), self.mesi, _s.operatori, _s.counts)
        messagebox.showinfo("Registrato",
            "Turni registrati nello storico.", parent=self)

    # Feature 10: personalizza DOCX
    def _open_docx_config(self):
        tpl = self._config.docx_template
        dlg = tk.Toplevel(self)
        dlg.title("Personalizza DOCX")
        dlg.geometry("450x300")
        dlg.configure(bg=BG)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Template Documento Word",
                 bg=BG, fg=TEXT, font=FONT_TITLE).pack(pady=(12, 8))

        inner = tk.Frame(dlg, bg=PANEL_BG)
        inner.pack(padx=20, pady=8, fill="x")

        fields: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate([
            ("title", "Titolo"),
            ("title_color", "Colore titolo (hex)"),
            ("subtitle_color", "Colore sottotitolo (hex)"),
            ("font_title", "Font titolo"),
            ("font_body", "Font corpo"),
        ]):
            tk.Label(inner, text=label, bg=PANEL_BG, fg=TEXT,
                     font=FONT_BOLD).grid(row=i, column=0, sticky="w", padx=8, pady=4)
            var = tk.StringVar(value=tpl.get(key, ""))
            fields[key] = var
            tk.Entry(inner, textvariable=var, font=FONT_LABEL, width=30,
                     relief="solid", highlightbackground=BORDER,
                     highlightthickness=1).grid(row=i, column=1, padx=8, pady=4)

        def apply():
            self._config.set_docx_template(**{k: v.get() for k, v in fields.items()})
            dlg.destroy()

        styled_btn(dlg, "Salva", apply, bg=SUCCESS).pack(pady=12)

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

    # -- Export ------------------------------------------------------------
    def _export(self, fmt):
        _s = self._last_solver
        has_result = _s is not None and (_s.solved_ok or _s.partial_result_available)
        if not has_result:
            messagebox.showinfo("Info",
                "Nessun risultato da esportare.", parent=self)
            return
        anno = self.anno.get()
        rows = self._result_rows_edit or _s.result_rows
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
                    messagebox.showerror("Errore", f"Impossibile scrivere:\n{e}", parent=self)
                    return

        elif fmt == "csv":
            path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.csv", parent=self)
            if path:
                try:
                    _write_text_file_atomic(path, _s.format_csv(anno), newline="")
                except OSError as e:
                    messagebox.showerror("Errore", f"Impossibile scrivere:\n{e}", parent=self)
                    return

        elif fmt == "docx":
            path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Documento Word", "*.docx"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.docx", parent=self)
            if path:
                try:
                    build_turni_docx(
                        path, anno, rows,
                        list(_s.operatori), list(_s.counts),
                        _s.status_str, _s.diff_val, _s.penalty_val,
                        template=self._config.docx_template,
                    )
                except OSError as e:
                    messagebox.showerror("Errore", f"Impossibile scrivere:\n{e}", parent=self)
                    return
                except Exception as e:
                    logger.exception("Errore DOCX")
                    messagebox.showerror("Errore DOCX", str(e), parent=self)
                    return

        elif fmt == "ics":
            path = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("iCalendar", "*.ics"), ("Tutti", "*.*")],
                initialfile=f"turni_{anno}.ics", parent=self)
            if path:
                try:
                    ics_data = build_ics(anno, rows)
                    _write_text_file_atomic(path, ics_data)
                except OSError as e:
                    messagebox.showerror("Errore", f"Impossibile scrivere:\n{e}", parent=self)
                    return

        if path:
            messagebox.showinfo("Salvato", f"File salvato:\n{path}", parent=self)

    # ================================================================
    #  SESSIONE JSON
    # ================================================================
    def _collect_session_data(self) -> dict | None:
        """Raccoglie lo stato corrente in formato sessione."""
        try:
            anno = self.anno.get().strip()
            ops = self.operatori or []
            mesi = self.mesi or []
            if not ops or not mesi:
                return None

            weeks = []
            for rd in self._week_rows:
                entry: dict = {
                    "month": rd["month"],
                    "week": rd["name_var"].get(),
                    "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()],
                }
                site = rd["site_var"].get()
                if site:
                    entry["site"] = site
                if rd["date_key"]:
                    entry["date_key"] = rd["date_key"]
                if rd["lock_var"].get():
                    entry["locked"] = True
                weeks.append(entry)

            data: dict = {
                "anno": anno,
                "operatori": ops,
                "mesi": mesi,
                "weeks": weeks,
                "solver_timeout": self.solver_timeout.get(),
            }
            # Feature 9: sedi
            if self.sites and self.sites != [""]:
                data["sites"] = self.sites
            # Feature 4: ruoli
            if self.operator_roles:
                data["operator_roles"] = {
                    str(k): sorted(v) for k, v in self.operator_roles.items()
                }
            # Feature 5: pattern
            if self.recurring_busy:
                data["recurring_busy"] = {
                    str(k): v for k, v in self.recurring_busy.items()
                }
            # Feature 6: assegnazioni bloccate
            if self._result_rows_edit:
                locked = []
                for r in self._result_rows_edit:
                    if r.get("locked") and r.get("locked_assignment"):
                        locked.append({
                            "week": r["week"],
                            "month": r["month"],
                            "locked_assignment": r["locked_assignment"],
                        })
                if locked:
                    data["locked_results"] = locked

            return data
        except Exception:
            return None

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

        data = self._collect_session_data()
        if data is None:
            messagebox.showerror("Errore", "Dati sessione non disponibili.", parent=self)
            return

        # Sovrascrivi con valori validati
        data["anno"] = anno
        data["operatori"] = parsed_ops
        data["mesi"] = parsed_mesi
        data["solver_timeout"] = timeout_str

        try:
            session = _validate_session_payload(data)
            session["solver_timeout"] = timeout_str
            # Preserva campi extra che la validazione non tocca
            for key in ("sites", "operator_roles", "recurring_busy", "locked_results"):
                if key in data:
                    session[key] = data[key]
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
            self._config.add_recent(path)
            self._build_recent_menu()
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
        self._load_session_from_path(path)

    def _load_session_from_path(self, path: str):
        if self._solving:
            messagebox.showwarning("Calcolo in corso",
                "Attendi il termine del calcolo.", parent=self)
            return
        if not os.path.exists(path):
            messagebox.showerror("Errore",
                f"File non trovato:\n{path}", parent=self)
            self._config.remove_recent(path)
            self._build_recent_menu()
            return
        try:
            with open(path, encoding="utf-8") as f:
                raw_session = json.load(f)
        except Exception as e:
            messagebox.showerror("Errore",
                f"Impossibile leggere il file:\n{e}", parent=self)
            return

        try:
            _validate_session_payload(raw_session)
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return

        self._config.add_recent(path)
        self._build_recent_menu()
        self._apply_session(raw_session)

    def _apply_session(self, session: dict):
        """Applica i dati di sessione alla UI."""
        ops = session.get("operatori", [])
        mesi = session.get("mesi", [])
        weeks = session.get("weeks", [])
        anno_val = session.get("anno", "")

        self._last_solver      = None
        self._last_result_text = ""
        self._result_rows_edit = []

        self.anno.set(anno_val)
        if "solver_timeout" in session:
            self.solver_timeout.set(session["solver_timeout"])

        self._ops_text.delete("1.0", "end")
        self._ops_text.insert("1.0", "\n".join(ops))
        self._mesi_entry.delete(0, "end")
        self._mesi_entry.insert(0, ", ".join(mesi))

        # Feature 9: sedi
        saved_sites = session.get("sites", [])
        if saved_sites:
            self.sites = saved_sites
            self._sites_entry.delete(0, "end")
            self._sites_entry.insert(0, ", ".join(saved_sites))

        self.operatori      = [o.strip() for o in ops]
        self.operatori_norm = [normalize_name(o) for o in self.operatori]
        self.mesi           = [m.strip() for m in mesi]
        self._prev_ops      = list(self.operatori)
        self._prev_mesi     = list(self.mesi)

        # Feature 4: ruoli
        saved_roles = session.get("operator_roles", {})
        self.operator_roles = {}
        for k, v in saved_roles.items():
            try:
                self.operator_roles[int(k)] = set(v)
            except (ValueError, TypeError):
                pass

        # Feature 5: pattern
        saved_patterns = session.get("recurring_busy", {})
        self.recurring_busy = {}
        for k, v in saved_patterns.items():
            try:
                self.recurring_busy[int(k)] = list(v)
            except (ValueError, TypeError):
                pass

        # Feature 6: assegnazioni bloccate
        locked_results = session.get("locked_results", [])
        locked_map: dict[str, dict] = {}
        for lr in locked_results:
            key = f"{lr.get('month', '')}|{lr.get('week', '')}"
            locked_map[key] = lr.get("locked_assignment", {})

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
                # Sede
                if w.get("site"):
                    rd["site_var"].set(w["site"])
                # Date key
                if w.get("date_key"):
                    rd["date_key"] = w["date_key"]
                # Lock
                if w.get("locked"):
                    rd["lock_var"].set(True)
                    key = f"{w.get('month', '')}|{w.get('week', '')}"
                    if key in locked_map:
                        # Salva l'assegnazione bloccata nei dati della riga
                        pass  # Will be handled in _step1_next via locked_assignment
            self._renumber_weeks(mese)

        self._show_step(1)
        messagebox.showinfo("Caricato",
            "Sessione caricata. Verifica i dati e clicca 'Genera Turni'.",
            parent=self)
