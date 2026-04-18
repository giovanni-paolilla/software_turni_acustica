"""Solver CP-SAT per la pianificazione turni, completamente disaccoppiato dalla UI."""
from __future__ import annotations
import csv
import io
import logging
import threading
from enum import Enum, auto
from typing import Any

from turni.constants import (
    PESO_DIFF, PESO_PENALTY, PESO_HISTORY, SOLVER_TIMEOUT, ALL_ROLES,
)
from turni.helpers import normalize_name, _safe_csv_cell

logger = logging.getLogger(__name__)

try:
    from ortools.sat.python import cp_model
    ORTOOLS_OK = True
except ImportError:
    ORTOOLS_OK = False


class SolvePhase(Enum):
    """Stato del ciclo di vita di una singola esecuzione di solve()."""
    IDLE              = auto()
    SOLVED            = auto()
    CANCELLED         = auto()
    CANCELLED_PARTIAL = auto()
    ERROR             = auto()


class _CancelCallback(cp_model.CpSolverSolutionCallback if ORTOOLS_OK else object):
    """Ferma la ricerca quando cancel_event e' impostato."""
    def __init__(self, cancel_event):
        if ORTOOLS_OK:
            super().__init__()
        self._ev = cancel_event

    def on_solution_callback(self):
        if self._ev.is_set():
            self.StopSearch()


class TurniSolver:
    """Logica CP-SAT completamente disaccoppiata dalla UI.

    Attributi pubblici dopo solve():
        result_rows  -- lista di dict con month/week/audio/video/sabato/busy/site
        counts       -- turni totali per operatore (indice parallelo a operatori)
        diff_val     -- delta max-min (obiettivo di equita')
        penalty_val  -- penalita' sabato sovrapposto a turno settimanale
        status_str   -- "OTTIMALE" | "FATTIBILE (timeout)"
        phase        -- SolvePhase corrente

    Parametri opzionali:
        operator_roles     -- dict[int, set[str]]: ruoli abilitati per operatore
        historical_counts  -- list[int]: turni storici cumulativi per operatore
        history_weight     -- float 0..1: peso dell'equita' storica nell'obiettivo
    """

    def __init__(self, operatori: list[str], weeks_data: list[dict], *,
                 operator_roles: dict[int, set[str]] | None = None,
                 historical_counts: list[int] | None = None,
                 history_weight: float = 0.0) -> None:
        self.operatori:   list[str]  = operatori
        self.weeks_data:  list[dict] = weeks_data
        self.result_rows: list[dict] = []
        self.counts:      list[int]  = []
        self.diff_val:    int        = 0
        self.penalty_val: int        = 0
        self.status_str:  str        = ""
        self.phase:       SolvePhase = SolvePhase.IDLE
        self.cp_solver               = None

        self.operator_roles:    dict[int, set[str]] = operator_roles or {}
        self.historical_counts: list[int] | None    = historical_counts
        self.history_weight:    float                = history_weight

    # -- Proprieta' di retrocompatibilita' ---------------------------------
    @property
    def solved_ok(self) -> bool:
        return self.phase == SolvePhase.SOLVED

    @property
    def cancelled(self) -> bool:
        return self.phase in (SolvePhase.CANCELLED, SolvePhase.CANCELLED_PARTIAL)

    @property
    def partial_result_available(self) -> bool:
        return self.phase == SolvePhase.CANCELLED_PARTIAL

    def _reset_result_state(self) -> None:
        self.result_rows = []
        self.counts = []
        self.diff_val = 0
        self.penalty_val = 0
        self.status_str = ""
        self.phase = SolvePhase.IDLE
        self.cp_solver = None

    # ======================================================================
    #  SOLVE
    # ======================================================================
    def solve(self, cancel_event: threading.Event | None = None,
              timeout: float = SOLVER_TIMEOUT) -> str:
        self._reset_result_state()

        if not ORTOOLS_OK:
            self.phase = SolvePhase.ERROR
            return ("ERRORE: ortools non installato.\n"
                    "Esegui:  pip install ortools")

        if cancel_event is None:
            cancel_event = threading.Event()
        if cancel_event.is_set():
            self.phase = SolvePhase.CANCELLED
            return "Calcolo annullato prima di trovare una soluzione."

        operatori  = self.operatori
        weeks_data = self.weeks_data
        n          = len(operatori)
        roles      = self.operator_roles
        all_r      = ALL_ROLES
        model      = cp_model.CpModel()
        turn_vars: dict[int, dict[str, Any]] = {}
        locked_indices: set[int] = set()

        # -- variabili per settimana ---------------------------------------
        for week_idx, week in enumerate(weeks_data):
            avail = week.get("available", [])

            # Settimana bloccata: variabili a dominio fisso
            la = week.get("locked_assignment")
            if week.get("locked") and isinstance(la, dict):
                locked_indices.add(week_idx)
                ta = model.NewIntVar(la["audio"], la["audio"], f"ta_w{week_idx}")
                tv = model.NewIntVar(la["video"], la["video"], f"tv_w{week_idx}")
                sat = model.NewIntVar(la["sabato"], la["sabato"], f"sat_w{week_idx}")
                turn_vars[week_idx] = {"t_audio": ta, "t_video": tv, "sat": sat}
                continue

            if not avail:
                self.phase = SolvePhase.ERROR
                return (f"Errore: nessun operatore per "
                        f"'{week['week']}' ({week['month']}).")

            # Filtra per ruolo (nessun fallback: se il pool e' vuoto, e' un errore)
            avail_a = [i for i in avail if "audio"  in roles.get(i, all_r)]
            avail_v = [i for i in avail if "video"  in roles.get(i, all_r)]
            avail_s = [i for i in avail if "sabato" in roles.get(i, all_r)]

            if not avail_a or not avail_v or not avail_s:
                missing = []
                if not avail_a:
                    missing.append("audio")
                if not avail_v:
                    missing.append("video")
                if not avail_s:
                    missing.append("sabato")
                self.phase = SolvePhase.ERROR
                return (f"Errore: nessun operatore abilitato per {', '.join(missing)} in "
                        f"'{week['week']}' ({week['month']}). "
                        f"Verifica la configurazione dei ruoli.")

            ta  = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail_a), f"ta_w{week_idx}")
            tv  = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail_v), f"tv_w{week_idx}")
            sat = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail_s), f"sat_w{week_idx}")
            model.Add(ta != tv)
            turn_vars[week_idx] = {"t_audio": ta, "t_video": tv, "sat": sat}

        # -- penalita' sabato sovrapposto ----------------------------------
        penalties: list[Any] = []
        for week_idx in range(len(weeks_data)):
            ta  = turn_vars[week_idx]["t_audio"]
            tv  = turn_vars[week_idx]["t_video"]
            sat = turn_vars[week_idx]["sat"]
            for t_var, lbl in [(ta, "a"), (tv, "v")]:
                pen = model.NewBoolVar(f"pen_{lbl}_w{week_idx}")
                model.Add(sat == t_var).OnlyEnforceIf(pen)
                model.Add(sat != t_var).OnlyEnforceIf(pen.Not())
                penalties.append(pen)
        tot_pen = model.NewIntVar(0, 2 * len(weeks_data), "tp")
        model.Add(tot_pen == sum(penalties))

        # -- indicatori e conteggi -----------------------------------------
        ind: list[list[Any]] = [[] for _ in range(n)]
        for week_idx, week in enumerate(weeks_data):
            if week_idx in locked_indices:
                # Per settimane bloccate, aggiungi indicatori fissi
                la = week["locked_assignment"]
                for op_idx in {la["audio"], la["video"], la["sabato"]}:
                    if 0 <= op_idx < n:
                        # conta quante volte appare
                        count = sum(1 for v in [la["audio"], la["video"], la["sabato"]]
                                    if v == op_idx)
                        for c in range(count):
                            b = model.NewBoolVar(f"b_locked_w{week_idx}_{op_idx}_{c}")
                            model.Add(b == 1)
                            ind[op_idx].append(b)
                continue

            for ruolo, var in turn_vars[week_idx].items():
                for i in week["available"]:
                    b = model.NewBoolVar(f"b_w{week_idx}_{ruolo}_{i}")
                    model.Add(var == i).OnlyEnforceIf(b)
                    model.Add(var != i).OnlyEnforceIf(b.Not())
                    ind[i].append(b)

        ub = 3 * len(weeks_data)
        cnt_vars = []
        for i in range(n):
            c = model.NewIntVar(0, ub, f"cnt{i}")
            model.Add(c == sum(ind[i]))
            cnt_vars.append(c)

        mx = model.NewIntVar(0, ub, "mx")
        mn = model.NewIntVar(0, ub, "mn")
        for c in cnt_vars:
            model.Add(c <= mx)
            model.Add(c >= mn)
        diff = model.NewIntVar(0, ub, "diff")
        model.Add(diff == mx - mn)

        # -- vincoli cross-sede --------------------------------------------
        date_groups: dict[str, list[int]] = {}
        for idx, w in enumerate(weeks_data):
            dk = w.get("date_key")
            if dk:
                date_groups.setdefault(dk, []).append(idx)
        for indices in date_groups.values():
            for i in range(len(indices)):
                for j in range(i + 1, len(indices)):
                    wi, wj = indices[i], indices[j]
                    site_i = weeks_data[wi].get("site", "")
                    site_j = weeks_data[wj].get("site", "")
                    if site_i and site_j and site_i != site_j:
                        for vi in turn_vars[wi].values():
                            for vj in turn_vars[wj].values():
                                model.Add(vi != vj)

        # -- obiettivo -----------------------------------------------------
        hist = self.historical_counts
        if hist and self.history_weight > 0 and len(hist) == n:
            max_hist = max(hist) if hist else 0
            total_ub = ub + max_hist
            total_vars = []
            for i in range(n):
                t = model.NewIntVar(0, total_ub, f"tot{i}")
                model.Add(t == cnt_vars[i] + hist[i])
                total_vars.append(t)
            hmx = model.NewIntVar(0, total_ub, "hmx")
            hmn = model.NewIntVar(0, total_ub, "hmn")
            for t in total_vars:
                model.Add(t <= hmx)
                model.Add(t >= hmn)
            hdiff = model.NewIntVar(0, total_ub, "hdiff")
            model.Add(hdiff == hmx - hmn)

            hw = max(1, int(self.history_weight * PESO_HISTORY))
            model.Minimize(
                diff * PESO_DIFF + tot_pen * PESO_PENALTY + hdiff * hw
            )
        else:
            model.Minimize(diff * PESO_DIFF + tot_pen * PESO_PENALTY)

        # -- solve ---------------------------------------------------------
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = timeout
        self.cp_solver = solver
        if cancel_event.is_set():
            self.phase = SolvePhase.CANCELLED
            self.cp_solver = None
            return "Calcolo annullato prima di trovare una soluzione."
        try:
            status = solver.Solve(model, _CancelCallback(cancel_event))
            cancel_requested = cancel_event.is_set()
        finally:
            self.cp_solver = None

        if cancel_requested and status not in [
                cp_model.OPTIMAL, cp_model.FEASIBLE]:
            self.phase = SolvePhase.CANCELLED
            return "Calcolo annullato prima di trovare una soluzione."

        if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            self.phase = SolvePhase.ERROR
            return "Nessuna soluzione trovata. Controlla i dati inseriti."

        # -- estrai risultati ----------------------------------------------
        self.result_rows = []
        for week_idx, week in enumerate(weeks_data):
            ta  = solver.Value(turn_vars[week_idx]["t_audio"])
            tv  = solver.Value(turn_vars[week_idx]["t_video"])
            sat = solver.Value(turn_vars[week_idx]["sat"])
            busy_names = ", ".join(
                operatori[i] for i, op in enumerate(operatori)
                if normalize_name(op) in week["busy"]
            ) if week["busy"] else ""
            row: dict[str, Any] = {
                "month":  week["month"],
                "week":   week["week"],
                "audio":  operatori[ta],
                "video":  operatori[tv],
                "sabato": operatori[sat],
                "busy":   busy_names,
            }
            if week.get("site"):
                row["site"] = week["site"]
            if week.get("locked"):
                row["locked"] = True
                row["locked_assignment"] = {"audio": ta, "video": tv, "sabato": sat}
            self.result_rows.append(row)

        self.counts      = [solver.Value(cnt_vars[i]) for i in range(n)]
        self.diff_val    = solver.Value(diff)
        self.penalty_val = solver.Value(tot_pen)
        if cancel_requested:
            self.phase = SolvePhase.CANCELLED_PARTIAL
            self.status_str = "ANNULLATO DALL'UTENTE (soluzione parziale)"
            return self._format_text()
        self.status_str = ("OTTIMALE" if status == cp_model.OPTIMAL
                           else "FATTIBILE (timeout)")
        self.phase = SolvePhase.SOLVED
        return self._format_text()

    # ======================================================================
    #  FORMAT
    # ======================================================================
    def _format_text(self) -> str:
        lines = []
        sep   = "-" * 64
        lines += [sep, "  PROGRAMMA TURNI", sep]

        audio_col_w: int = max(
            (len(r["audio"]) for r in self.result_rows), default=20)

        cur_month = None
        cur_site  = None
        for r in self.result_rows:
            site = r.get("site", "")
            if site and site != cur_site:
                cur_site = site
                lines.append(f"\n  === {site.upper()} ===")
            if r["month"] != cur_month:
                cur_month = r["month"]
                lines.append(f"\n  {cur_month.upper()}")
                lines.append("-" * 40)
            lock_tag = " [BLOCCATO]" if r.get("locked") else ""
            lines.append(f"  {r['week']}{lock_tag}")
            lines.append(
                f"    Giovedi  | Audio: {r['audio']:<{audio_col_w}}  Video: {r['video']}")
            lines.append(f"    Domenica | Operatore A/V: {r['sabato']}")
            if r["busy"]:
                lines.append(f"    Busy     | {r['busy']}")
            lines.append("")

        lines += [sep, "  RIEPILOGO TURNI PER OPERATORE", sep]
        paired  = sorted(zip(self.operatori, self.counts), key=lambda x: -x[1])
        max_n   = max((len(p[0]) for p in paired), default=10)
        for op, cnt in paired:
            lines.append(f"  {op:<{max_n}}  {'|' * cnt}  ({cnt})")

        lines += [
            "",
            f"  Stato           : {self.status_str}",
            f"  Delta max-min   : {self.diff_val}",
            f"  Penalita' domenica: {self.penalty_val}",
            sep,
        ]
        return "\n".join(lines)

    def format_csv(self, anno: str) -> str:
        buf = io.StringIO()
        w   = csv.writer(buf)
        has_sites = any(r.get("site") for r in self.result_rows)
        if has_sites:
            w.writerow(["Anno", "Sede", "Mese", "Settimana",
                        "Giovedi Audio", "Giovedi Video", "Domenica A/V", "Busy"])
            for r in self.result_rows:
                w.writerow([
                    _safe_csv_cell(anno),
                    _safe_csv_cell(r.get("site", "")),
                    _safe_csv_cell(r["month"]),
                    _safe_csv_cell(r["week"]),
                    _safe_csv_cell(r["audio"]),
                    _safe_csv_cell(r["video"]),
                    _safe_csv_cell(r["sabato"]),
                    _safe_csv_cell(r["busy"]),
                ])
        else:
            w.writerow(["Anno", "Mese", "Settimana",
                        "Giovedi Audio", "Giovedi Video", "Domenica A/V", "Busy"])
            for r in self.result_rows:
                w.writerow([
                    _safe_csv_cell(anno),
                    _safe_csv_cell(r["month"]),
                    _safe_csv_cell(r["week"]),
                    _safe_csv_cell(r["audio"]),
                    _safe_csv_cell(r["video"]),
                    _safe_csv_cell(r["sabato"]),
                    _safe_csv_cell(r["busy"]),
                ])
        return buf.getvalue()
