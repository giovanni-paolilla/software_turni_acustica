"""Applicazione principale CustomTkinter: wizard a 3 step per la gestione turni."""
from __future__ import annotations
import datetime
import json
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

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


class TurniApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestione Turni")
        self.geometry("920x750")
        self.minsize(800, 620)
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

        self.operator_roles:  dict[int, set[str]] = {}
        self.recurring_busy:  dict[int, list[int]] = {}
        self.operator_sites:  dict[int, list[str]] = {}

        self.use_history     = tk.BooleanVar(value=False)
        self.history_weight  = tk.DoubleVar(value=self._config.history_weight * 100)

        self._last_result_text = ""
        self._last_solver      = None
        self._result_rows_edit: list[dict] = []

        self._solving       = False
        self._cancel_event  = threading.Event()
        self._pending_close   = False
        self._pending_restart = False

        self.current_step = 0
        self.steps        = ["Operatori & Anno", "Settimane", "Risultati"]

        self._step1_built  = False
        self._prev_ops:  list[str] = []
        self._prev_mesi: list[str] = []
        self._week_rows: list[dict]= []
        self._month_frames: dict[str, ctk.CTkFrame] = {}

        self._btn_save_session = None
        self._btn_load_session = None
        self._btn_modify = None
        self._btn_save_txt = None
        self._btn_save_csv = None
        self._btn_save_docx = None
        self._btn_restart = None
        self._cancel_btn = None

        self._view_mode = "text"
        self._result_tree = None
        self._tree_frame = None
        self._text_frame = None
        self._edit_combo = None

        self._build_ui()
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

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
        hdr = ctk.CTkFrame(self, fg_color=HEADER_BG, height=56, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  Gestione Turni",
                     text_color="white",
                     font=FONT_TITLE).pack(side="left", padx=12, pady=10)
        mb = ctk.CTkFrame(hdr, fg_color=HEADER_BG, corner_radius=0)
        mb.pack(side="right", padx=12)
        self._btn_save_session = styled_btn(mb, "Salva sessione", self._save_session,
                                            bg=HEADER_BTN)
        self._btn_save_session.pack(side="left", padx=4)

        self._btn_load_session = styled_btn(mb, "Carica sessione",
                                            self._show_load_menu, bg=HEADER_BTN)
        self._btn_load_session.pack(side="left", padx=4)
        self._load_menu = tk.Menu(self, tearoff=0)

        if not ORTOOLS_OK:
            ctk.CTkLabel(hdr, text="ATTENZIONE: ortools mancante!",
                         text_color=INFO_WARN,
                         font=FONT_SMALL).pack(side="right", padx=8)

    def _show_load_menu(self):
        menu = self._load_menu
        menu.delete(0, "end")
        for path in self._config.recent_sessions:
            name = os.path.basename(path)
            menu.add_command(label=name,
                             command=lambda p=path: self._load_session_from_path(p))
        if self._config.recent_sessions:
            menu.add_separator()
        menu.add_command(label="Sfoglia...", command=self._load_session)
        btn = self._btn_load_session
        menu.post(btn.winfo_rootx(), btn.winfo_rooty() + btn.winfo_height())

    def _build_stepbar(self):
        bar = ctk.CTkFrame(self, fg_color=PANEL_BG, height=46, corner_radius=0,
                           border_width=1, border_color=BORDER)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        self._step_labels = []
        for i, name in enumerate(self.steps):
            lbl = ctk.CTkLabel(bar, text=f"  {i+1}. {name}  ",
                               text_color=MUTED, font=FONT_SMALL)
            lbl.pack(side="left", padx=4, pady=12)
            self._step_labels.append(lbl)
            if i < len(self.steps)-1:
                ctk.CTkLabel(bar, text=">", text_color=BORDER,
                             font=(_FONT_FAMILY, 14)).pack(side="left")

    def _build_content(self):
        self.content = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.content.pack(fill="both", expand=True)
        self.frames = {}
        for i in range(len(self.steps)):
            f = ctk.CTkFrame(self.content, fg_color=BG, corner_radius=0)
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.frames[i] = f
        self._build_step0(self.frames[0])
        self._build_step1(self.frames[1])
        self._build_step2(self.frames[2])

    def _show_step(self, n):
        self.current_step = n
        for i, lbl in enumerate(self._step_labels):
            if i == n:
                lbl.configure(text_color=ACCENT, font=FONT_BOLD, fg_color=MONTH_HDR_BG)
            elif i < n:
                lbl.configure(text_color=SUCCESS, font=FONT_SMALL, fg_color=PANEL_BG)
            else:
                lbl.configure(text_color=MUTED, font=FONT_SMALL, fg_color=PANEL_BG)
        self.frames[n].lift()

    def _set_solver_ui_state(self, solving: bool):
        state = "disabled" if solving else "normal"
        for btn in (
            self._btn_save_session, self._btn_load_session,
            self._btn_modify, self._btn_save_txt, self._btn_save_csv,
            self._btn_save_docx, self._btn_restart,
        ):
            if btn is not None:
                try:
                    btn.configure(state=state)
                except (tk.TclError, ValueError):
                    pass
        if self._cancel_btn is not None:
            self._cancel_btn.configure(state="normal" if solving else "disabled")

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
            self._cancel_btn.configure(state="disabled")

    def _perform_restart(self):
        self.weeks_data        = []
        self._last_solver      = None
        self._last_result_text = ""
        self._result_rows_edit = []
        self._pending_restart  = False
        self._result_subtitle.configure(text="", text_color=MUTED)
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
            self._result_subtitle.configure(text="Annullamento in corso...", text_color=WARN)
            self._set_result("Annullamento in corso.\n")
            self._set_solver_ui_state(True)
            if self._cancel_btn is not None:
                self._cancel_btn.configure(state="disabled")
            return
        self._config.clear_autosave()
        self.destroy()

    # ================================================================
    #  AUTO-SAVE
    # ================================================================
    def _schedule_autosave(self):
        self._autosave_id = self.after(
            self._config.autosave_interval * 1000, self._do_autosave)

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
        if messagebox.askyesno("Recupero sessione",
                "E' stata trovata una sessione non salvata.\nVuoi ripristinarla?",
                parent=self):
            self._apply_session(recovery)
        self._config.clear_autosave()

    # ================================================================
    #  STEP 0
    # ================================================================
    def _build_step0(self, parent):
        wrap = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        wrap.pack(expand=True)

        ctk.CTkLabel(wrap, text="Operatori e Anno",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(28, 4))
        ctk.CTkLabel(wrap, text="Inserisci i dati base per pianificare i turni.",
                     text_color=MUTED, font=FONT_LABEL).pack(pady=(0, 12))

        c = card(wrap)
        c.pack(padx=40, pady=8, fill="x")
        inner = ctk.CTkFrame(c, fg_color=PANEL_BG, corner_radius=0)
        inner.pack(padx=24, pady=16, fill="x")
        inner.columnconfigure(1, weight=1)

        def lbl(row, text):
            ctk.CTkLabel(inner, text=text, text_color=TEXT,
                         font=FONT_BOLD).grid(row=row, column=0, sticky="nw", pady=6)

        def entry(row, var=None, width=350):
            e = ctk.CTkEntry(inner, font=FONT_LABEL, width=width,
                             textvariable=var)
            e.grid(row=row, column=1, sticky="ew", padx=12, pady=6)
            return e

        lbl(0, "Anno")
        self._anno_entry = entry(0, var=self.anno, width=100)

        lbl(1, "Mesi (virgola)")
        self._mesi_entry = entry(1)
        self._mesi_entry.insert(0, "Gennaio, Febbraio")

        lbl(2, "Sedi (virgola)")
        self._sites_entry = entry(2)
        self._sites_entry.insert(0, ", ".join(self.sites))

        lbl(3, "Operatori\n(uno per riga)")
        self._ops_text = ctk.CTkTextbox(inner, font=FONT_LABEL, width=350, height=110,
                                         fg_color=PANEL_BG, border_width=1,
                                         border_color=BORDER)
        self._ops_text.grid(row=3, column=1, sticky="ew", padx=12, pady=6)
        self._ops_text.insert("1.0",
            "Mario Rossi\nLuca Bianchi\nAnna Verdi\nGiulia Neri")

        lbl(4, "Timeout solver (sec)")
        timeout_frame = ctk.CTkFrame(inner, fg_color=PANEL_BG, corner_radius=0)
        timeout_frame.grid(row=4, column=1, sticky="w", padx=12, pady=6)
        ctk.CTkEntry(timeout_frame, textvariable=self.solver_timeout,
                     font=FONT_LABEL, width=80).pack(side="left")
        ctk.CTkLabel(timeout_frame,
                     text="  (aumenta se necessario)",
                     text_color=MUTED, font=FONT_SMALL).pack(side="left")

        # Equita' storica
        hist_frame = ctk.CTkFrame(inner, fg_color=PANEL_BG, corner_radius=0)
        hist_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=12, pady=6)
        ctk.CTkCheckBox(hist_frame, text="Usa equita' storica",
                        variable=self.use_history,
                        text_color=TEXT, font=FONT_BOLD).pack(side="left")
        ctk.CTkLabel(hist_frame, text="Peso:", text_color=MUTED,
                     font=FONT_SMALL).pack(side="left", padx=(16, 4))
        self._hist_weight_label = ctk.CTkLabel(hist_frame,
                                               text=f"{int(self.history_weight.get())}%",
                                               text_color=TEXT, font=FONT_SMALL, width=40)
        ctk.CTkSlider(hist_frame, from_=0, to=100, variable=self.history_weight,
                      width=120, command=lambda v: self._hist_weight_label.configure(
                          text=f"{int(v)}%")).pack(side="left")
        self._hist_weight_label.pack(side="left")

        # Ruoli e pattern
        cfg_frame = ctk.CTkFrame(inner, fg_color=PANEL_BG, corner_radius=0)
        cfg_frame.grid(row=6, column=0, columnspan=2, sticky="w", padx=12, pady=6)
        styled_btn(cfg_frame, "Configura ruoli", self._open_roles_dialog,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)
        styled_btn(cfg_frame, "Pattern indisponibilita'", self._open_patterns_dialog,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)

        styled_btn(wrap, "Avanti  >", self._step0_next).pack(pady=16)

    def _step0_next(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso",
                "Attendi il termine del calcolo.", parent=self)
            return
        try:
            anno, timeout_str, new_ops, new_norm, new_mesi, dups, dup_mesi = _parse_step0_inputs(
                self.anno.get(), self._mesi_entry.get(),
                self._ops_text.get("1.0", "end"), self.solver_timeout.get())
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return

        self.anno.set(anno)
        self.solver_timeout.set(timeout_str)
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
            self.operatori, self.operatori_norm, self.mesi = new_ops, new_norm, new_mesi
            self._show_step(1)
            return
        if self._step1_built and self._week_rows and not (ops_ok and mesi_ok):
            if not messagebox.askyesno("Attenzione",
                    "Hai modificato operatori o mesi.\nLe settimane verranno azzerate.",
                    parent=self):
                return

        self.operatori, self.operatori_norm, self.mesi = new_ops, new_norm, new_mesi
        self._prev_ops, self._prev_mesi = list(new_ops), list(new_mesi)
        self.weeks_data = []
        self.operator_roles = {i: set(ALL_ROLES) for i in range(len(new_ops))}
        self.operator_sites = {i: list(self.sites) for i in range(len(new_ops))}
        self._build_step1_content()
        self._step1_built = True
        self._show_step(1)

    # ================================================================
    #  DIALOG RUOLI
    # ================================================================
    def _open_roles_dialog(self):
        temp_ops = self.operatori
        if not temp_ops:
            from turni.helpers import _parse_operatori_text
            temp_ops = _parse_operatori_text(self._ops_text.get("1.0", "end"))
        if not temp_ops:
            messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Configura Ruoli Operatori")
        dlg.geometry("520x420")
        dlg.transient(self)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Ruoli per operatore",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(12, 8))

        sf = ScrollableFrame(dlg)
        sf.pack(fill="both", expand=True, padx=16)

        role_vars: dict[int, dict[str, tk.BooleanVar]] = {}
        for i, op in enumerate(temp_ops):
            row = ctk.CTkFrame(sf.inner, fg_color=PANEL_BG, corner_radius=6,
                               border_width=1, border_color=BORDER)
            row.pack(fill="x", pady=2, padx=4)
            ctk.CTkLabel(row, text=op, text_color=TEXT, font=FONT_BOLD,
                         width=160, anchor="w").pack(side="left", padx=8)
            role_vars[i] = {}
            current = self.operator_roles.get(i, ALL_ROLES)
            for role in ["audio", "video", "sabato"]:
                var = tk.BooleanVar(value=(role in current))
                role_vars[i][role] = var
                ctk.CTkCheckBox(row, text=role.capitalize(), variable=var,
                                text_color=TEXT, font=FONT_SMALL).pack(side="left", padx=8)

        def apply():
            for i in role_vars:
                roles = {r for r, v in role_vars[i].items() if v.get()}
                self.operator_roles[i] = roles if roles else set(ALL_ROLES)
            dlg.destroy()
        styled_btn(dlg, "Applica", apply, bg=SUCCESS).pack(pady=12)

    # ================================================================
    #  DIALOG PATTERN
    # ================================================================
    def _open_patterns_dialog(self):
        temp_ops = self.operatori
        if not temp_ops:
            from turni.helpers import _parse_operatori_text
            temp_ops = _parse_operatori_text(self._ops_text.get("1.0", "end"))
        if not temp_ops:
            messagebox.showinfo("Info", "Inserisci prima gli operatori.", parent=self)
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Pattern Indisponibilita' Ricorrente")
        dlg.geometry("620x420")
        dlg.transient(self)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Settimane del mese sempre non disponibili",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(12, 4))
        ctk.CTkLabel(dlg, text="Spunta le settimane in cui l'operatore e' regolarmente impegnato.",
                     text_color=MUTED, font=FONT_SMALL).pack(pady=(0, 8))

        sf = ScrollableFrame(dlg)
        sf.pack(fill="both", expand=True, padx=16)

        pattern_vars: dict[int, dict[int, tk.BooleanVar]] = {}
        for i, op in enumerate(temp_ops):
            row = ctk.CTkFrame(sf.inner, fg_color=PANEL_BG, corner_radius=6,
                               border_width=1, border_color=BORDER)
            row.pack(fill="x", pady=2, padx=4)
            ctk.CTkLabel(row, text=op, text_color=TEXT, font=FONT_BOLD,
                         width=140, anchor="w").pack(side="left", padx=8)
            pattern_vars[i] = {}
            current = self.recurring_busy.get(i, [])
            for wk in range(1, 6):
                var = tk.BooleanVar(value=(wk in current))
                pattern_vars[i][wk] = var
                ctk.CTkCheckBox(row, text=f"{wk}a sett.", variable=var,
                                text_color=TEXT, font=FONT_SMALL).pack(side="left", padx=4)

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
    #  STEP 1
    # ================================================================
    def _build_step1(self, parent):
        self._step1_parent = parent

    def _build_step1_content(self, add_default_row: bool = True):
        for w in self._step1_parent.winfo_children():
            w.destroy()
        self._week_rows    = []
        self._month_frames = {}

        ctk.CTkLabel(self._step1_parent, text="Pianificazione Settimane",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(20, 2))
        ctk.CTkLabel(self._step1_parent,
                     text="Aggiungi settimane e spunta gli operatori non disponibili.",
                     text_color=MUTED, font=FONT_LABEL).pack(pady=(0, 8))

        sf = ScrollableFrame(self._step1_parent)
        sf.pack(fill="both", expand=True, padx=20)

        for mese in self.mesi:
            mc = card(sf.inner)
            mc.pack(fill="x", pady=6, padx=4)
            hdr = ctk.CTkFrame(mc, fg_color=MONTH_HDR_BG, corner_radius=6)
            hdr.pack(fill="x")
            ctk.CTkLabel(hdr, text=f"  {mese}", text_color=ACCENT,
                         font=FONT_BOLD, anchor="w").pack(side="left", padx=10, pady=8)
            styled_btn(hdr, "Auto-genera",
                       lambda m=mese: self._auto_generate_weeks(m),
                       bg="#7C3AED", fg="white").pack(side="right", padx=4, pady=6)
            styled_btn(hdr, "+ Aggiungi settimana",
                       lambda m=mese: self._add_week_row(m),
                       bg=ACCENT_LIGHT, fg=ACCENT).pack(side="right", padx=4, pady=6)
            rows_frame = ctk.CTkFrame(mc, fg_color=PANEL_BG, corner_radius=0)
            rows_frame.pack(fill="x", padx=8, pady=4)
            self._month_frames[mese] = rows_frame
            if add_default_row:
                self._add_week_row(mese)

        nav = ctk.CTkFrame(self._step1_parent, fg_color=BG, corner_radius=0)
        nav.pack(pady=14)
        styled_btn(nav, "<  Indietro", lambda: self._show_step(0),
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=8)
        styled_btn(nav, "Genera Turni  >", self._step1_next,
                   bg=SUCCESS).pack(side="left", padx=8)

    def _auto_generate_weeks(self, mese):
        try:
            year = int(self.anno.get())
        except ValueError:
            messagebox.showerror("Errore", "Anno non valido.", parent=self)
            return
        weeks = generate_weeks_for_month(year, mese)
        if not weeks:
            messagebox.showinfo("Info",
                f"Impossibile generare settimane per '{mese}' {year}.", parent=self)
            return
        existing = [r for r in self._week_rows if r["month"] == mese]
        if existing:
            if not messagebox.askyesno("Attenzione",
                    f"Sostituire le {len(existing)} settimane esistenti per {mese}?",
                    parent=self):
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
        if not week_of_month:
            return
        for op_idx, weeks_pattern in self.recurring_busy.items():
            if week_of_month in weeks_pattern and op_idx in rd["busy_vars"]:
                rd["busy_vars"][op_idx].set(True)

    def _add_week_row(self, mese):
        rows_frame = self._month_frames[mese]
        row_idx    = sum(1 for r in self._week_rows if r["month"] == mese)
        multi_site = len(self.sites) > 1 and self.sites[0]

        row = ctk.CTkFrame(rows_frame, fg_color=PANEL_BG, corner_radius=6,
                           border_width=1, border_color=BTN_SECONDARY)
        row.pack(fill="x", pady=3, padx=4)

        lf = ctk.CTkFrame(row, fg_color=PANEL_BG, corner_radius=0)
        lf.pack(side="left", padx=8, pady=6)
        num_lbl = ctk.CTkLabel(lf, text=f"Settimana {row_idx+1}:",
                               text_color=TEXT, font=FONT_BOLD)
        num_lbl.pack(anchor="w")
        name_var = tk.StringVar(value=f"Sett.{row_idx+1} {mese[:3]}")
        ctk.CTkEntry(lf, textvariable=name_var, font=FONT_LABEL,
                     width=180).pack(pady=2)

        site_var = tk.StringVar(value=self.sites[0] if self.sites else "")
        if multi_site:
            sf = ctk.CTkFrame(lf, fg_color=PANEL_BG, corner_radius=0)
            sf.pack(anchor="w", pady=2)
            ctk.CTkLabel(sf, text="Sede:", text_color=MUTED,
                         font=FONT_SMALL).pack(side="left")
            ctk.CTkComboBox(sf, variable=site_var, values=self.sites,
                            state="readonly", width=120,
                            font=FONT_SMALL).pack(side="left", padx=4)

        lock_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(lf, text="Blocca", variable=lock_var,
                        text_color=MUTED, font=FONT_SMALL).pack(anchor="w")

        bf = ctk.CTkFrame(row, fg_color=PANEL_BG, corner_radius=0)
        bf.pack(side="left", padx=16, pady=6, fill="x", expand=True)
        ctk.CTkLabel(bf, text="Operatori busy:", text_color=MUTED,
                     font=FONT_SMALL).pack(anchor="w")
        chk = ctk.CTkFrame(bf, fg_color=PANEL_BG, corner_radius=0)
        chk.pack(anchor="w")
        busy_vars = {}
        cols = 4
        for idx, op in enumerate(self.operatori):
            var = tk.BooleanVar(value=False)
            busy_vars[idx] = var
            ctk.CTkCheckBox(chk, text=op, variable=var,
                            text_color=TEXT, font=FONT_SMALL).grid(
                row=idx//cols, column=idx%cols, sticky="w", padx=8, pady=1)

        rd = {
            "month": mese, "name_var": name_var, "busy_vars": busy_vars,
            "site_var": site_var, "lock_var": lock_var, "row_widget": row,
            "num_label": num_lbl, "date_key": "", "week_of_month": 0,
        }
        self._week_rows.append(rd)
        styled_btn(row, "X", lambda r=rd, rw=row: self._remove_week_row(r, rw),
                   bg=DANGER_LIGHT, fg=DANGER).pack(side="right", padx=8, pady=6)

    def _remove_week_row(self, rd, rw):
        mese = rd["month"]
        rw.destroy()
        if rd in self._week_rows:
            self._week_rows.remove(rd)
        self._renumber_weeks(mese)

    def _renumber_weeks(self, mese):
        for i, rd in enumerate(r for r in self._week_rows if r["month"] == mese):
            rd["num_label"].configure(text=f"Settimana {i+1}:")

    def _step1_next(self):
        if self._solving:
            messagebox.showinfo("Calcolo in corso",
                "Il calcolo e' ancora in esecuzione.", parent=self)
            return
        if not self._week_rows:
            messagebox.showerror("Errore", "Aggiungi almeno una settimana.", parent=self)
            return

        raw_weeks = []
        for rd in self._week_rows:
            entry = {"month": rd["month"], "week": rd["name_var"].get(),
                     "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()]}
            site = rd["site_var"].get()
            if site:
                entry["site"] = site
            if rd["date_key"]:
                entry["date_key"] = rd["date_key"]
            if rd["lock_var"].get():
                entry["locked"] = True
            raw_weeks.append(entry)

        try:
            canonical_weeks = _validate_week_entries(raw_weeks, declared_months=self.mesi,
                                                     error_prefix="Errore Step 1")
            _validate_solver_ready_weeks(canonical_weeks, operator_count=len(self.operatori),
                                         error_prefix="Errore Step 1")
        except SessionValidationError as e:
            msg = str(e)
            if msg.startswith("Errore Step 1: "):
                msg = msg[len("Errore Step 1: "):]
            messagebox.showerror("Errori", msg, parent=self)
            return

        new_weeks_data = []
        for week in canonical_weeks:
            busy_norm = [self.operatori_norm[i] for i in week["busy_indices"]]
            available = [i for i, op in enumerate(self.operatori_norm) if op not in busy_norm]
            wd: dict = {"month": week["month"], "week": week["week"],
                        "busy": busy_norm, "available": available}
            if week.get("site"):
                wd["site"] = week["site"]
            if week.get("date_key"):
                wd["date_key"] = week["date_key"]
            if week.get("locked") and week.get("locked_assignment"):
                wd["locked"] = True
                wd["locked_assignment"] = week["locked_assignment"]
            new_weeks_data.append(wd)

        if not new_weeks_data:
            messagebox.showerror("Errore", "Nessuna settimana valida.", parent=self)
            return
        self.weeks_data = _order_weeks_by_declared_months(new_weeks_data, self.mesi)
        self._show_step(2)
        self._run_solver()

    # ================================================================
    #  STEP 2
    # ================================================================
    def _build_step2(self, parent):
        ctk.CTkLabel(parent, text="Risultati",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(16, 2))
        self._result_subtitle = ctk.CTkLabel(parent, text="",
                                              text_color=MUTED, font=FONT_LABEL)
        self._result_subtitle.pack(pady=(0, 4))

        toggle = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        toggle.pack(pady=2)
        self._btn_view_text = styled_btn(toggle, "Testo",
                                         lambda: self._toggle_view("text"), bg=ACCENT)
        self._btn_view_text.pack(side="left", padx=2)
        self._btn_view_table = styled_btn(toggle, "Tabella",
                                          lambda: self._toggle_view("table"),
                                          bg=BTN_SECONDARY, fg=TEXT)
        self._btn_view_table.pack(side="left", padx=2)

        # Text view
        self._text_frame = card(parent)
        self._text_frame.pack(fill="both", expand=True, padx=20, pady=4)
        self._result_text = ctk.CTkTextbox(self._text_frame, font=FONT_MONO,
                                            wrap="none", fg_color=PANEL_BG,
                                            text_color=TEXT, state="disabled")
        self._result_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Table view (ttk.Treeview — no CTk equivalent)
        self._tree_frame = card(parent)
        cols = ("settimana", "audio", "video", "sabato", "sede", "stato")
        self._result_tree = ttk.Treeview(self._tree_frame, columns=cols,
                                          show="headings", height=14)
        for col, heading, w in [
            ("settimana", "Settimana", 180), ("audio", "Audio", 130),
            ("video", "Video", 130), ("sabato", "Domenica A/V", 130),
            ("sede", "Sede", 100), ("stato", "Stato", 60),
        ]:
            self._result_tree.heading(col, text=heading)
            self._result_tree.column(col, width=w, anchor="center")
        tree_sb = ttk.Scrollbar(self._tree_frame, command=self._result_tree.yview)
        self._result_tree.configure(yscrollcommand=tree_sb.set)
        tree_sb.pack(side="right", fill="y")
        self._result_tree.pack(fill="both", expand=True)
        self._result_tree.bind("<Double-1>", self._on_table_double_click)

        self._prog_frame = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        self._prog_frame.pack(pady=4)
        self._progress = ctk.CTkProgressBar(self._prog_frame, mode="indeterminate", width=280)
        self._progress.pack(side="left", padx=6)
        self._cancel_btn = styled_btn(self._prog_frame, "Annulla",
                                      self._cancel_solving, bg=DANGER)
        self._cancel_btn.pack(side="left", padx=6)
        self._prog_frame.pack_forget()

        nav = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        nav.pack(pady=4)
        self._btn_modify = styled_btn(nav, "<  Modifica", lambda: self._show_step(1),
                                      bg=BTN_SECONDARY, fg=TEXT)
        self._btn_modify.pack(side="left", padx=4)
        self._btn_save_txt = styled_btn(nav, "Salva TXT", lambda: self._export("txt"),
                                        bg=SUCCESS)
        self._btn_save_txt.pack(side="left", padx=4)
        self._btn_save_csv = styled_btn(nav, "Salva CSV", lambda: self._export("csv"),
                                        bg=CYAN)
        self._btn_save_csv.pack(side="left", padx=4)
        self._btn_save_docx = styled_btn(nav, "Salva DOCX", lambda: self._export("docx"),
                                         bg="#2563EB")
        self._btn_save_docx.pack(side="left", padx=4)
        self._btn_restart = styled_btn(nav, "Nuovo calcolo", self._restart, bg=WARN)
        self._btn_restart.pack(side="left", padx=4)

        nav2 = ctk.CTkFrame(parent, fg_color=BG, corner_radius=0)
        nav2.pack(pady=(0, 8))
        styled_btn(nav2, "Salva ICS", lambda: self._export("ics"),
                   bg="#7C3AED").pack(side="left", padx=4)
        styled_btn(nav2, "Copia", self._copy_to_clipboard,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)
        styled_btn(nav2, "Registra storico", self._record_history,
                   bg="#0D9488").pack(side="left", padx=4)
        styled_btn(nav2, "Personalizza DOCX", self._open_docx_config,
                   bg=BTN_SECONDARY, fg=TEXT).pack(side="left", padx=4)

    def _toggle_view(self, mode: str):
        self._view_mode = mode
        if mode == "text":
            self._tree_frame.pack_forget()
            self._text_frame.pack(fill="both", expand=True, padx=20, pady=4,
                                  after=self._result_subtitle)
            self._btn_view_text.configure(fg_color=ACCENT, text_color="white")
            self._btn_view_table.configure(fg_color=BTN_SECONDARY, text_color=TEXT)
        else:
            self._text_frame.pack_forget()
            self._tree_frame.pack(fill="both", expand=True, padx=20, pady=4,
                                  after=self._result_subtitle)
            self._btn_view_text.configure(fg_color=BTN_SECONDARY, text_color=TEXT)
            self._btn_view_table.configure(fg_color=ACCENT, text_color="white")
            self._populate_result_tree()

    def _populate_result_tree(self):
        tree = self._result_tree
        tree.delete(*tree.get_children())
        for i, r in enumerate(self._result_rows_edit):
            stato = "BLK" if r.get("locked") else "OK"
            tree.insert("", "end", iid=str(i), values=(
                r.get("week", ""), r.get("audio", ""), r.get("video", ""),
                r.get("sabato", ""), r.get("site", ""), stato))

    def _on_table_double_click(self, event):
        tree = self._result_tree
        if tree.identify_region(event.x, event.y) != "cell":
            return
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item or not column:
            return
        col_idx = int(column[1:]) - 1
        if col_idx not in (1, 2, 3):
            return
        row_idx = int(item)
        if self._result_rows_edit[row_idx].get("locked"):
            return

        if row_idx < len(self.weeks_data):
            avail = [self.operatori[i] for i in self.weeks_data[row_idx].get("available", [])
                     if 0 <= i < len(self.operatori)]
        else:
            avail = list(self.operatori)

        bbox = tree.bbox(item, column)
        if not bbox:
            return
        x, y, w, h = bbox
        if self._edit_combo is not None:
            self._edit_combo.destroy()

        combo = ttk.Combobox(tree, values=avail, state="readonly", font=FONT_SMALL)
        combo.set(tree.set(item, column))
        combo.place(x=x, y=y, width=w, height=h)
        combo.focus_set()
        self._edit_combo = combo
        col_key = ["settimana", "audio", "video", "sabato", "sede", "stato"][col_idx]

        def on_select(e=None):
            val = combo.get()
            if val:
                tree.set(item, column, val)
                self._result_rows_edit[row_idx][col_key] = val
                self._validate_table_row(row_idx)
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
        stato = "!!" if row.get("audio") == row.get("video") and row.get("audio") else "OK"
        self._result_tree.set(str(row_idx), "stato", stato)

    def _run_solver(self):
        ops_snap = list(self.operatori)
        weeks_snap = [{**w, "available": list(w["available"]), "busy": list(w["busy"])}
                      for w in self.weeks_data]
        roles = dict(self.operator_roles) if self.operator_roles else None
        hist_counts, hist_weight = None, 0.0
        if self.use_history.get():
            hist_counts = self._history.get_cumulative_counts(ops_snap)
            hist_weight = self.history_weight.get() / 100.0

        self._solving, self._cancel_event = True, threading.Event()
        self._pending_close = self._pending_restart = False
        self._last_solver = TurniSolver(ops_snap, weeks_snap, operator_roles=roles,
                                        historical_counts=hist_counts,
                                        history_weight=hist_weight)
        self._result_subtitle.configure(text="Ottimizzazione in corso...", text_color=MUTED)
        self._prog_frame.pack(pady=4)
        self._progress.start()
        self._set_solver_ui_state(True)
        self._set_result("Calcolo in corso, attendere...\n")

        ev = self._cancel_event
        try:
            timeout = float(self.solver_timeout.get())
            if timeout <= 0:
                raise ValueError
        except ValueError:
            timeout = SOLVER_TIMEOUT
            self.solver_timeout.set(str(int(SOLVER_TIMEOUT)))
        solver_ref = self._last_solver

        def worker():
            try:
                txt = solver_ref.solve(ev, timeout=timeout)
            except Exception as exc:
                logger.exception("Errore imprevisto nel solver")
                txt = f"Errore imprevisto:\n{exc}"
            try:
                self.after(0, lambda: self._display_result(txt))
            except tk.TclError:
                pass
        threading.Thread(target=worker, daemon=True).start()

    def _cancel_solving(self):
        self._request_cancel()
        self._result_subtitle.configure(text="Annullamento in corso...", text_color=WARN)
        self._set_result("Annullamento in corso...\n")

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
            sub, col = "Pianificazione completata.", SUCCESS
        elif cancelled and partial:
            sub, col = "Annullato: soluzione parziale.", WARN
        elif cancelled:
            sub, col = "Calcolo annullato.", WARN
        else:
            sub, col = text.split("\n")[0], DANGER
        self._result_subtitle.configure(text=sub, text_color=col)
        self._set_result(text)

        if _s and _s.result_rows:
            self._result_rows_edit = [dict(r) for r in _s.result_rows]
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
        self._result_text.configure(state="normal")
        self._result_text.delete("1.0", "end")
        self._result_text.insert("1.0", text)
        self._result_text.configure(state="disabled")

    def _copy_to_clipboard(self):
        _s = self._last_solver
        if not _s or not (_s.solved_ok or _s.partial_result_available):
            messagebox.showinfo("Info", "Nessun risultato da copiare.", parent=self)
            return
        text = format_whatsapp(self.anno.get(), self._result_rows_edit,
                               _s.operatori, _s.counts)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("Copiato", "Risultati copiati (formato WhatsApp).", parent=self)

    def _record_history(self):
        _s = self._last_solver
        if not _s or not _s.solved_ok:
            messagebox.showinfo("Info", "Solo dopo una soluzione completa.", parent=self)
            return
        if not messagebox.askyesno("Conferma",
                "Registrare questi turni nello storico?", parent=self):
            return
        self._history.record_session(self.anno.get(), self.mesi, _s.operatori, _s.counts)
        messagebox.showinfo("Registrato", "Turni registrati nello storico.", parent=self)

    def _open_docx_config(self):
        tpl = self._config.docx_template
        dlg = ctk.CTkToplevel(self)
        dlg.title("Personalizza DOCX")
        dlg.geometry("480x320")
        dlg.transient(self)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Template Documento Word",
                     text_color=TEXT, font=FONT_TITLE).pack(pady=(12, 8))
        inner = ctk.CTkFrame(dlg, fg_color=PANEL_BG, corner_radius=8)
        inner.pack(padx=20, pady=8, fill="x")

        fields: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate([
            ("title", "Titolo"), ("title_color", "Colore titolo (hex)"),
            ("subtitle_color", "Colore sottotitolo (hex)"),
            ("font_title", "Font titolo"), ("font_body", "Font corpo"),
        ]):
            ctk.CTkLabel(inner, text=label, text_color=TEXT,
                         font=FONT_BOLD).grid(row=i, column=0, sticky="w", padx=8, pady=4)
            var = tk.StringVar(value=tpl.get(key, ""))
            fields[key] = var
            ctk.CTkEntry(inner, textvariable=var, font=FONT_LABEL,
                         width=240).grid(row=i, column=1, padx=8, pady=4)

        def apply():
            self._config.set_docx_template(**{k: v.get() for k, v in fields.items()})
            dlg.destroy()
        styled_btn(dlg, "Salva", apply, bg=SUCCESS).pack(pady=12)

    def _restart(self):
        if self._solving:
            self._pending_restart = True
            self._request_cancel()
            self._result_subtitle.configure(text="Annullamento...", text_color=WARN)
            self._set_result("Annullamento in corso.\n")
            return
        self._perform_restart()

    def _export(self, fmt):
        _s = self._last_solver
        if not _s or not (_s.solved_ok or _s.partial_result_available):
            messagebox.showinfo("Info", "Nessun risultato da esportare.", parent=self)
            return
        anno = self.anno.get()
        rows = self._result_rows_edit or _s.result_rows
        path = None

        if fmt == "txt":
            path = filedialog.asksaveasfilename(
                defaultextension=".txt", filetypes=[("Testo", "*.txt")],
                initialfile=f"turni_{anno}.txt", parent=self)
            if path:
                try:
                    _write_text_file_atomic(path, self._last_result_text)
                except OSError as e:
                    messagebox.showerror("Errore", str(e), parent=self)
                    return
        elif fmt == "csv":
            path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                initialfile=f"turni_{anno}.csv", parent=self)
            if path:
                try:
                    _write_text_file_atomic(path, _s.format_csv(anno), newline="")
                except OSError as e:
                    messagebox.showerror("Errore", str(e), parent=self)
                    return
        elif fmt == "docx":
            path = filedialog.asksaveasfilename(
                defaultextension=".docx", filetypes=[("Word", "*.docx")],
                initialfile=f"turni_{anno}.docx", parent=self)
            if path:
                try:
                    build_turni_docx(path, anno, rows, list(_s.operatori), list(_s.counts),
                                    _s.status_str, _s.diff_val, _s.penalty_val,
                                    template=self._config.docx_template)
                except Exception as e:
                    logger.exception("Errore DOCX")
                    messagebox.showerror("Errore", str(e), parent=self)
                    return
        elif fmt == "ics":
            path = filedialog.asksaveasfilename(
                defaultextension=".ics", filetypes=[("iCalendar", "*.ics")],
                initialfile=f"turni_{anno}.ics", parent=self)
            if path:
                try:
                    _write_text_file_atomic(path, build_ics(anno, rows))
                except OSError as e:
                    messagebox.showerror("Errore", str(e), parent=self)
                    return
        if path:
            messagebox.showinfo("Salvato", f"File salvato:\n{path}", parent=self)

    # ================================================================
    #  SESSIONE JSON
    # ================================================================
    def _collect_session_data(self) -> dict | None:
        try:
            anno = self.anno.get().strip()
            ops, mesi = self.operatori or [], self.mesi or []
            if not ops or not mesi:
                return None
            weeks = []
            for rd in self._week_rows:
                entry: dict = {"month": rd["month"], "week": rd["name_var"].get(),
                               "busy_indices": [i for i, v in rd["busy_vars"].items() if v.get()]}
                site = rd["site_var"].get()
                if site:
                    entry["site"] = site
                if rd["date_key"]:
                    entry["date_key"] = rd["date_key"]
                if rd["lock_var"].get():
                    entry["locked"] = True
                weeks.append(entry)
            data: dict = {"anno": anno, "operatori": ops, "mesi": mesi,
                          "weeks": weeks, "solver_timeout": self.solver_timeout.get()}
            if self.sites and self.sites != [""]:
                data["sites"] = self.sites
            if self.operator_roles:
                data["operator_roles"] = {str(k): sorted(v) for k, v in self.operator_roles.items()}
            if self.recurring_busy:
                data["recurring_busy"] = {str(k): v for k, v in self.recurring_busy.items()}
            if self._result_rows_edit:
                locked = [{"week": r["week"], "month": r["month"],
                           "locked_assignment": r["locked_assignment"]}
                          for r in self._result_rows_edit
                          if r.get("locked") and r.get("locked_assignment")]
                if locked:
                    data["locked_results"] = locked
            return data
        except Exception:
            return None

    def _save_session(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso", "Attendi il termine.", parent=self)
            return
        try:
            anno, timeout_str, parsed_ops, _, parsed_mesi, dup_ops, dup_mesi = _parse_step0_inputs(
                self.anno.get(), self._mesi_entry.get(),
                self._ops_text.get("1.0", "end"), self.solver_timeout.get())
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        if dup_ops:
            messagebox.showwarning("Attenzione", "Duplicati rimossi:\n" + "\n".join(dup_ops), parent=self)
        if dup_mesi:
            messagebox.showwarning("Attenzione", "Mesi duplicati:\n" + "\n".join(dup_mesi), parent=self)
        if self._step1_built and (parsed_ops != self.operatori or parsed_mesi != self.mesi):
            messagebox.showerror("Errore",
                "Modifiche non applicate. Clicca 'Avanti' prima di salvare.", parent=self)
            return

        data = self._collect_session_data()
        if data is None:
            messagebox.showerror("Errore", "Dati non disponibili.", parent=self)
            return
        data.update({"anno": anno, "operatori": parsed_ops, "mesi": parsed_mesi,
                     "solver_timeout": timeout_str})
        try:
            session = _validate_session_payload(data)
            session["solver_timeout"] = timeout_str
            for key in ("sites", "operator_roles", "recurring_busy", "locked_results"):
                if key in data:
                    session[key] = data[key]
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("Sessione", "*.json")],
            initialfile=f"sessione_{session['anno']}.json", parent=self)
        if path:
            try:
                _write_text_file_atomic(path, json.dumps(session, ensure_ascii=False, indent=2))
            except OSError as e:
                messagebox.showerror("Errore", str(e), parent=self)
                return
            self._config.add_recent(path)
            messagebox.showinfo("Salvato", f"Sessione salvata:\n{path}", parent=self)

    def _load_session(self):
        if self._solving:
            messagebox.showwarning("Calcolo in corso", "Attendi il termine.", parent=self)
            return
        path = filedialog.askopenfilename(
            filetypes=[("Sessione", "*.json"), ("Tutti", "*.*")], parent=self)
        if path:
            self._load_session_from_path(path)

    def _load_session_from_path(self, path: str):
        if self._solving:
            return
        if not os.path.exists(path):
            messagebox.showerror("Errore", f"File non trovato:\n{path}", parent=self)
            self._config.remove_recent(path)
            return
        try:
            with open(path, encoding="utf-8") as f:
                raw = json.load(f)
        except Exception as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        try:
            _validate_session_payload(raw)
        except SessionValidationError as e:
            messagebox.showerror("Errore", str(e), parent=self)
            return
        self._config.add_recent(path)
        self._apply_session(raw)

    def _apply_session(self, session: dict):
        ops = session.get("operatori", [])
        mesi = session.get("mesi", [])
        weeks = session.get("weeks", [])

        self._last_solver = None
        self._last_result_text = ""
        self._result_rows_edit = []

        self.anno.set(session.get("anno", ""))
        if "solver_timeout" in session:
            self.solver_timeout.set(session["solver_timeout"])
        self._ops_text.configure(state="normal")
        self._ops_text.delete("1.0", "end")
        self._ops_text.insert("1.0", "\n".join(ops))
        self._mesi_entry.delete(0, "end")
        self._mesi_entry.insert(0, ", ".join(mesi))

        saved_sites = session.get("sites", [])
        if saved_sites:
            self.sites = saved_sites
            self._sites_entry.delete(0, "end")
            self._sites_entry.insert(0, ", ".join(saved_sites))

        self.operatori = [o.strip() for o in ops]
        self.operatori_norm = [normalize_name(o) for o in self.operatori]
        self.mesi = [m.strip() for m in mesi]
        self._prev_ops, self._prev_mesi = list(self.operatori), list(self.mesi)

        self.operator_roles = {}
        for k, v in session.get("operator_roles", {}).items():
            try:
                self.operator_roles[int(k)] = set(v)
            except (ValueError, TypeError):
                pass
        self.recurring_busy = {}
        for k, v in session.get("recurring_busy", {}).items():
            try:
                self.recurring_busy[int(k)] = list(v)
            except (ValueError, TypeError):
                pass

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
                if w.get("site"):
                    rd["site_var"].set(w["site"])
                if w.get("date_key"):
                    rd["date_key"] = w["date_key"]
                if w.get("locked"):
                    rd["lock_var"].set(True)
            self._renumber_weeks(mese)

        self._show_step(1)
        messagebox.showinfo("Caricato",
            "Sessione caricata. Verifica e clicca 'Genera Turni'.", parent=self)
