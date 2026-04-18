"""Solver CP-SAT per la pianificazione turni, completamente disaccoppiato dalla UI."""
from __future__ import annotations
import csv
import io
import logging
import threading
from enum import Enum, auto

from turni.constants import PESO_DIFF, PESO_PENALTY, SOLVER_TIMEOUT
from turni.helpers import normalize_name, _safe_csv_cell

logger = logging.getLogger(__name__)

try:
    from ortools.sat.python import cp_model
    ORTOOLS_OK = True
except ImportError:
    ORTOOLS_OK = False


class SolvePhase(Enum):
    """Stato del ciclo di vita di una singola esecuzione di solve()."""
    IDLE              = auto()   # reset iniziale / dopo _reset_result_state
    SOLVED            = auto()   # soluzione trovata correttamente
    CANCELLED         = auto()   # annullato senza soluzione
    CANCELLED_PARTIAL = auto()   # annullato con soluzione parziale disponibile
    ERROR             = auto()   # errore imprevisto


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
        result_rows  -- lista di dict con month/week/audio/video/sabato/busy
        counts       -- turni totali per operatore (indice parallelo a operatori)
        diff_val     -- delta max-min (obiettivo di equita')
        penalty_val  -- penalita' sabato sovrapposto a turno settimanale
        status_str   -- "OTTIMALE" | "FATTIBILE (timeout)"
        phase        -- SolvePhase corrente (IDLE / SOLVED / CANCELLED / CANCELLED_PARTIAL / ERROR)

    Proprieta' di retrocompatibilita' (derivate da phase):
        solved_ok              -- True se phase == SOLVED
        cancelled              -- True se phase in (CANCELLED, CANCELLED_PARTIAL)
        partial_result_available -- True se phase == CANCELLED_PARTIAL
    """

    def __init__(self, operatori: list[str], weeks_data: list[dict]) -> None:
        self.operatori:   list[str]  = operatori
        self.weeks_data:  list[dict] = weeks_data
        self.result_rows: list[dict] = []
        self.counts:      list[int]  = []
        self.diff_val:    int        = 0
        self.penalty_val: int        = 0
        self.status_str:  str        = ""
        self.phase:       SolvePhase = SolvePhase.IDLE
        self.cp_solver               = None

    # ── Proprieta' di retrocompatibilita' ─────────────────────
    @property
    def solved_ok(self) -> bool:
        """True solo se la soluzione e' completa e valida."""
        return self.phase == SolvePhase.SOLVED

    @property
    def cancelled(self) -> bool:
        """True se il solve e' stato annullato (con o senza risultato parziale)."""
        return self.phase in (SolvePhase.CANCELLED, SolvePhase.CANCELLED_PARTIAL)

    @property
    def partial_result_available(self) -> bool:
        """True se esiste una soluzione parziale dopo annullamento."""
        return self.phase == SolvePhase.CANCELLED_PARTIAL

    def _reset_result_state(self) -> None:
        self.result_rows = []
        self.counts = []
        self.diff_val = 0
        self.penalty_val = 0
        self.status_str = ""
        self.phase = SolvePhase.IDLE
        self.cp_solver = None

    def solve(self, cancel_event: threading.Event | None = None,
              timeout: float = SOLVER_TIMEOUT) -> str:
        """Esegue il solver CP-SAT e restituisce il testo risultato.

        Args:
            cancel_event: Event impostato per interrompere la ricerca.
            timeout:      Limite di tempo in secondi (default: SOLVER_TIMEOUT).

        Returns:
            Stringa formattata del risultato, o messaggio di errore.
        """
        self._reset_result_state()

        if not ORTOOLS_OK:
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
        model      = cp_model.CpModel()
        turn_vars  = {}

        for week_idx, week in enumerate(weeks_data):
            key   = week_idx
            avail = week["available"]
            if not avail:
                return (f"Errore: nessun operatore per "
                        f"'{week['week']}' ({week['month']}).")
            ta  = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail), f"ta_w{key}")
            tv  = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail), f"tv_w{key}")
            sat = model.NewIntVarFromDomain(
                cp_model.Domain.FromValues(avail), f"sat_w{key}")
            model.Add(ta != tv)
            turn_vars[key] = {"t_audio": ta, "t_video": tv, "sat": sat}

        penalties = []
        for week_idx, week in enumerate(weeks_data):
            key = week_idx
            ta, tv, sat = (turn_vars[key]["t_audio"],
                           turn_vars[key]["t_video"],
                           turn_vars[key]["sat"])
            for t_var, lbl in [(ta, "a"), (tv, "v")]:
                pen = model.NewBoolVar(f"pen_{lbl}_w{key}")
                model.Add(sat == t_var).OnlyEnforceIf(pen)
                model.Add(sat != t_var).OnlyEnforceIf(pen.Not())
                penalties.append(pen)
        tot_pen = model.NewIntVar(0, 2*len(weeks_data), "tp")
        model.Add(tot_pen == sum(penalties))

        ind: list[list] = [[] for _ in range(n)]
        for week_idx, week in enumerate(weeks_data):
            key = week_idx
            for ruolo, var in turn_vars[key].items():
                for i in week["available"]:
                    b = model.NewBoolVar(f"b_w{key}_{ruolo}_{i}")
                    model.Add(var == i).OnlyEnforceIf(b)
                    model.Add(var != i).OnlyEnforceIf(b.Not())
                    ind[i].append(b)

        cnt_vars = []
        for i in range(n):
            c = model.NewIntVar(0, 3*len(weeks_data), f"cnt{i}")
            model.Add(c == sum(ind[i]))
            cnt_vars.append(c)

        mx = model.NewIntVar(0, 3*len(weeks_data), "mx")
        mn = model.NewIntVar(0, 3*len(weeks_data), "mn")
        for c in cnt_vars:
            model.Add(c <= mx)
            model.Add(c >= mn)
        diff = model.NewIntVar(0, 3*len(weeks_data), "diff")
        model.Add(diff == mx - mn)

        model.Minimize(diff * PESO_DIFF + tot_pen * PESO_PENALTY)

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
            return "Nessuna soluzione trovata. Controlla i dati inseriti."

        self.result_rows = []
        for week_idx, week in enumerate(weeks_data):
            key = week_idx
            ta  = solver.Value(turn_vars[key]["t_audio"])
            tv  = solver.Value(turn_vars[key]["t_video"])
            sat = solver.Value(turn_vars[key]["sat"])
            busy_names = ", ".join(
                operatori[i] for i, op in enumerate(operatori)
                if normalize_name(op) in week["busy"]
            ) if week["busy"] else ""
            self.result_rows.append({
                "month":  week["month"],
                "week":   week["week"],
                "audio":  operatori[ta],
                "video":  operatori[tv],
                "sabato": operatori[sat],
                "busy":   busy_names,
            })

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

    def _format_text(self) -> str:
        lines = []
        sep   = "-" * 64
        lines += [sep, "  PROGRAMMA TURNI", sep]

        audio_col_w: int = max((len(r["audio"]) for r in self.result_rows), default=20)

        cur_month = None
        for r in self.result_rows:
            if r["month"] != cur_month:
                cur_month = r["month"]
                lines.append(f"\n  {cur_month.upper()}")
                lines.append("-" * 40)
            lines.append(f"  {r['week']}")
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
        w.writerow(["Anno","Mese","Settimana",
                    "Giovedi Audio","Giovedi Video","Domenica A/V","Busy"])
        for r in self.result_rows:
            w.writerow([_safe_csv_cell(anno), _safe_csv_cell(r["month"]), _safe_csv_cell(r["week"]),
                        _safe_csv_cell(r["audio"]), _safe_csv_cell(r["video"]),
                        _safe_csv_cell(r["sabato"]), _safe_csv_cell(r["busy"])])
        return buf.getvalue()
