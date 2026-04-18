"""Test per le nuove funzionalita': calendario, storico, config, ICS, solver avanzato."""
import json
import os
import pathlib
import tempfile
import threading
import unittest

from turni.calendar_utils import generate_weeks_for_month, month_name_to_number
from turni.history import HistoryStore
from turni.config import UserConfig
from turni.ics_export import build_ics, format_whatsapp, _parse_week_dates
from turni.solver import TurniSolver, SolvePhase, ORTOOLS_OK
from turni.constants import ALL_ROLES


class CalendarUtilsTests(unittest.TestCase):
    def test_month_name_to_number(self):
        self.assertEqual(month_name_to_number("Gennaio"), 1)
        self.assertEqual(month_name_to_number(" febbraio "), 2)
        self.assertEqual(month_name_to_number("DICEMBRE"), 12)
        self.assertIsNone(month_name_to_number("January"))

    def test_generate_weeks_for_month_january_2026(self):
        weeks = generate_weeks_for_month(2026, "Gennaio")
        self.assertTrue(len(weeks) >= 4)
        for w in weeks:
            self.assertEqual(w["month"], "Gennaio")
            self.assertIn("Sett.", w["week"])
            self.assertTrue(w["thursday"].startswith("2026-01"))
            self.assertIn("week_of_month", w)
            self.assertGreaterEqual(w["week_of_month"], 1)

    def test_generate_weeks_cross_month_boundary(self):
        # Aprile 2026: ultimo giovedi 30 aprile, domenica 3 maggio
        weeks = generate_weeks_for_month(2026, "Aprile")
        last = weeks[-1]
        self.assertIn("Apr", last["week"])
        self.assertTrue(
            last["sunday"].startswith("2026-05") or last["sunday"].startswith("2026-04"))

    def test_generate_weeks_unknown_month_returns_empty(self):
        self.assertEqual(generate_weeks_for_month(2026, "Foobar"), [])

    def test_week_of_month_increments(self):
        weeks = generate_weeks_for_month(2026, "Gennaio")
        indices = [w["week_of_month"] for w in weeks]
        self.assertEqual(indices, list(range(1, len(indices) + 1)))


class HistoryStoreTests(unittest.TestCase):
    def test_record_and_retrieve(self):
        with tempfile.TemporaryDirectory() as tmp:
            h = HistoryStore(directory=tmp)
            h.record_session("2026", ["Gen"], ["Mario", "Luca"], [5, 3])
            counts = h.get_cumulative_counts(["Mario", "Luca"])
            self.assertEqual(counts, [5, 3])
            self.assertEqual(len(h.get_sessions()), 1)

    def test_cumulative_across_sessions(self):
        with tempfile.TemporaryDirectory() as tmp:
            h = HistoryStore(directory=tmp)
            h.record_session("2026", ["Gen"], ["Mario", "Luca"], [5, 3])
            h.record_session("2026", ["Feb"], ["Mario", "Luca"], [4, 6])
            counts = h.get_cumulative_counts(["Mario", "Luca"])
            self.assertEqual(counts, [9, 9])

    def test_persistence(self):
        with tempfile.TemporaryDirectory() as tmp:
            h1 = HistoryStore(directory=tmp)
            h1.record_session("2026", ["Gen"], ["Mario"], [5])
            h2 = HistoryStore(directory=tmp)
            self.assertEqual(h2.get_cumulative_counts(["Mario"]), [5])

    def test_clear(self):
        with tempfile.TemporaryDirectory() as tmp:
            h = HistoryStore(directory=tmp)
            h.record_session("2026", ["Gen"], ["Mario"], [5])
            h.clear()
            self.assertEqual(h.get_cumulative_counts(["Mario"]), [0])

    def test_unknown_operator_returns_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            h = HistoryStore(directory=tmp)
            self.assertEqual(h.get_cumulative_counts(["Unknown"]), [0])


class UserConfigTests(unittest.TestCase):
    def test_recent_sessions(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            c.add_recent("/tmp/a.json")
            c.add_recent("/tmp/b.json")
            self.assertEqual(c.recent_sessions[0], os.path.abspath("/tmp/b.json"))
            self.assertEqual(len(c.recent_sessions), 2)

    def test_recent_dedup(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            c.add_recent("/tmp/a.json")
            c.add_recent("/tmp/b.json")
            c.add_recent("/tmp/a.json")
            self.assertEqual(len(c.recent_sessions), 2)
            self.assertEqual(c.recent_sessions[0], os.path.abspath("/tmp/a.json"))

    def test_docx_template_defaults(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            tpl = c.docx_template
            self.assertIn("title", tpl)
            self.assertIn("font_body", tpl)

    def test_docx_template_override(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            c.set_docx_template(title="CUSTOM TITLE")
            c2 = UserConfig(directory=tmp)
            self.assertEqual(c2.docx_template["title"], "CUSTOM TITLE")

    def test_autosave(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            c.save_autosave({"anno": "2026", "test": True})
            data = c.load_autosave()
            self.assertIsNotNone(data)
            self.assertEqual(data["anno"], "2026")
            c.clear_autosave()
            self.assertIsNone(c.load_autosave())

    def test_sites(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = UserConfig(directory=tmp)
            self.assertEqual(c.sites, ["Messina", "Ganzirri"])
            c.set_sites(["Roma", "Milano"])
            c2 = UserConfig(directory=tmp)
            self.assertEqual(c2.sites, ["Roma", "Milano"])


class IcsExportTests(unittest.TestCase):
    def test_parse_week_dates_same_month(self):
        thu, sun = _parse_week_dates("Sett. 08-11 Gen", "2026")
        self.assertIsNotNone(thu)
        self.assertIsNotNone(sun)
        self.assertEqual(thu.month, 1)
        self.assertEqual(thu.day, 8)
        self.assertEqual(sun.day, 11)

    def test_parse_week_dates_cross_month(self):
        thu, sun = _parse_week_dates("Sett. 30 Apr - 03 Mag", "2026")
        self.assertEqual(thu.month, 4)
        self.assertEqual(sun.month, 5)

    def test_build_ics_contains_events(self):
        rows = [
            {"week": "Sett. 08-11 Gen", "audio": "Mario", "video": "Luca",
             "sabato": "Anna", "month": "Gennaio", "busy": ""},
        ]
        ics = build_ics("2026", rows)
        self.assertIn("BEGIN:VCALENDAR", ics)
        self.assertIn("BEGIN:VEVENT", ics)
        self.assertIn("Mario", ics)
        self.assertIn("END:VCALENDAR", ics)

    def test_build_ics_with_site(self):
        rows = [
            {"week": "Sett. 08-11 Gen", "audio": "Mario", "video": "Luca",
             "sabato": "Anna", "month": "Gennaio", "busy": "", "site": "Messina"},
        ]
        ics = build_ics("2026", rows)
        self.assertIn("[Messina]", ics)

    def test_format_whatsapp(self):
        rows = [
            {"week": "Sett. 08-11 Gen", "audio": "Mario", "video": "Luca",
             "sabato": "Anna", "month": "Gennaio", "busy": ""},
        ]
        text = format_whatsapp("2026", rows, ["Mario", "Luca", "Anna"], [2, 1, 1])
        self.assertIn("TURNI AUDIO/VIDEO", text)
        self.assertIn("Mario", text)
        self.assertIn("RIEPILOGO", text)


@unittest.skipUnless(ORTOOLS_OK, "ortools non installato")
class SolverAdvancedTests(unittest.TestCase):
    def _make_weeks(self, n=4, site=""):
        weeks = []
        for i in range(n):
            w = {
                "month": "Gennaio",
                "week": f"Sett.{i+1}",
                "available": [0, 1, 2],
                "busy": [],
            }
            if site:
                w["site"] = site
                w["date_key"] = f"2026-01-{(i+1)*7:02d}"
            weeks.append(w)
        return weeks

    def test_solve_with_roles_restricts_assignments(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = self._make_weeks(2)
        # Mario: solo audio, Luca: solo video, Anna: tutto
        roles = {0: {"audio", "sabato"}, 1: {"video", "sabato"},
                 2: {"audio", "video", "sabato"}}
        solver = TurniSolver(ops, weeks, operator_roles=roles)
        solver.solve()
        self.assertEqual(solver.phase, SolvePhase.SOLVED)
        for r in solver.result_rows:
            # Mario (indice 0) non deve mai apparire in video
            self.assertNotEqual(r["video"], "Mario",
                                "Mario e' abilitato solo audio ma assegnato a video")
            # Luca (indice 1) non deve mai apparire in audio
            self.assertNotEqual(r["audio"], "Luca",
                                "Luca e' abilitato solo video ma assegnato ad audio")

    def test_solve_with_roles_error_on_empty_pool(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = self._make_weeks(1)
        # Nessuno abilitato all'audio
        roles = {0: {"video", "sabato"}, 1: {"video", "sabato"},
                 2: {"video", "sabato"}}
        solver = TurniSolver(ops, weeks, operator_roles=roles)
        result = solver.solve()
        self.assertEqual(solver.phase, SolvePhase.ERROR)
        self.assertIn("audio", result)

    def test_solve_with_historical_counts(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = self._make_weeks(4)
        # Mario ha storicamente 20 turni, gli altri 5
        hist = [20, 5, 5]
        solver = TurniSolver(ops, weeks, historical_counts=hist, history_weight=0.8)
        solver.solve()
        self.assertEqual(solver.phase, SolvePhase.SOLVED)
        # Mario dovrebbe avere meno turni in questa sessione
        mario_count = solver.counts[0]
        max_others = max(solver.counts[1:])
        self.assertLessEqual(mario_count, max_others + 2)

    def test_solve_with_locked_week(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = [
            {
                "month": "Gennaio", "week": "Sett.1",
                "available": [0, 1, 2], "busy": [],
                "locked": True,
                "locked_assignment": {"audio": 0, "video": 1, "sabato": 2},
            },
            {
                "month": "Gennaio", "week": "Sett.2",
                "available": [0, 1, 2], "busy": [],
            },
        ]
        solver = TurniSolver(ops, weeks)
        solver.solve()
        self.assertEqual(solver.phase, SolvePhase.SOLVED)
        # Prima settimana bloccata
        r0 = solver.result_rows[0]
        self.assertEqual(r0["audio"], "Mario")
        self.assertEqual(r0["video"], "Luca")
        self.assertEqual(r0["sabato"], "Anna")

    def test_solve_cross_site_constraints(self):
        ops = ["Mario", "Luca", "Anna", "Giulia"]
        weeks = [
            {
                "month": "Gennaio", "week": "Sett.1",
                "available": [0, 1, 2, 3], "busy": [],
                "site": "Messina", "date_key": "2026-01-01",
            },
            {
                "month": "Gennaio", "week": "Sett.1",
                "available": [0, 1, 2, 3], "busy": [],
                "site": "Ganzirri", "date_key": "2026-01-01",
            },
        ]
        solver = TurniSolver(ops, weeks)
        solver.solve()
        self.assertEqual(solver.phase, SolvePhase.SOLVED)
        r0 = solver.result_rows[0]
        r1 = solver.result_rows[1]
        # Nessun operatore in comune tra le due sedi nella stessa data
        site_a = {r0["audio"], r0["video"], r0["sabato"]}
        site_b = {r1["audio"], r1["video"], r1["sabato"]}
        self.assertEqual(len(site_a & site_b), 0,
                         f"Operatori condivisi tra sedi: {site_a & site_b}")

    def test_solve_csv_with_sites(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = self._make_weeks(1, site="Messina")
        solver = TurniSolver(ops, weeks)
        solver.solve()
        csv_text = solver.format_csv("2026")
        self.assertIn("Sede", csv_text)
        self.assertIn("Messina", csv_text)

    def test_format_text_shows_site_header(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = self._make_weeks(1, site="Messina")
        solver = TurniSolver(ops, weeks)
        solver.solve()
        text = solver._format_text()
        self.assertIn("MESSINA", text)

    def test_locked_week_counts_toward_total(self):
        ops = ["Mario", "Luca", "Anna"]
        weeks = [
            {
                "month": "Gennaio", "week": "Sett.1",
                "available": [0, 1, 2], "busy": [],
                "locked": True,
                "locked_assignment": {"audio": 0, "video": 1, "sabato": 2},
            },
        ]
        solver = TurniSolver(ops, weeks)
        solver.solve()
        self.assertEqual(solver.phase, SolvePhase.SOLVED)
        # Ogni operatore ha 1 turno dalla settimana bloccata
        self.assertEqual(solver.counts, [1, 1, 1])


if __name__ == "__main__":
    unittest.main()
