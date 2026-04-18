# Software Turni Acustica

Applicazione desktop per la **pianificazione turni audio/video** delle sedi di Messina e Ganzirri.

Ottimizza l'assegnazione degli operatori ai turni (giovedì audio/video + domenica) usando il solver **CP-SAT** di Google OR-Tools, garantendo distribuzione equa dei turni e rispettando le indisponibilità settimanali.

---

## Funzionalità

- Inserimento operatori, mesi e anno di pianificazione
- Configurazione settimane con marcatura operatori "busy"
- Ottimizzazione automatica con CP-SAT (equità + penalità sovrapposizione sabato/martedì)
- Annullamento solver in corso con recupero soluzione parziale
- Export risultati in **TXT**, **CSV** (formula-injection safe) e **DOCX** (layout professionale landscape)
- Salvataggio/caricamento sessione in JSON
- Scritture file atomiche con locking distribuito (safe con istanze multiple)

---

## Requisiti

- Python 3.8+
- Tkinter (incluso nella maggior parte delle distribuzioni Python)

```bash
pip install -r requirements.txt
```

---

## Avvio

```bash
python turni_v16.py
```

---

## Struttura del progetto

```
turni_v16.py          entry point
turni/
    constants.py      costanti globali (solver, tema, font)
    helpers.py        funzioni pure di utilità
    validators.py     validazione e canonizzazione dati sessione
    io_utils.py       I/O atomico e locking distribuito
    docx_export.py    generazione documento Word
    solver.py         SolvePhase enum + TurniSolver (CP-SAT)
    ui/
        widgets.py    ScrollableFrame, styled_btn, card
        app.py        TurniApp — wizard a 3 step
tests/
    test_helpers.py   suite unittest
turni_v15.py          versione monolitica originale (riferimento)
```

---

## Test

```bash
python -m unittest discover -s tests -v
```

---

## Licenza

MIT — vedi [LICENSE](LICENSE)
