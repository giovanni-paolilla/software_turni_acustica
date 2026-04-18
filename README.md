# Software Turni Acustica

Applicazione desktop per la **pianificazione turni audio/video** delle sedi di Messina e Ganzirri.

Ottimizza l'assegnazione degli operatori ai turni (giovedi audio/video + domenica) usando il solver **CP-SAT** di Google OR-Tools, garantendo distribuzione equa dei turni e rispettando le indisponibilita' settimanali.

---

## Funzionalita'

- Inserimento operatori, mesi e anno di pianificazione
- **Generazione automatica settimane** da calendario (giovedi-domenica)
- Configurazione settimane con marcatura operatori "busy"
- **Ruoli operatore** configurabili (audio, video, sabato)
- **Pattern indisponibilita' ricorrente** (1a, 2a, 3a... settimana del mese)
- **Gestione multi-sede** con vincoli cross-sede (stesso operatore non assegnabile a due sedi contemporaneamente)
- **Equita' storica inter-sessione** con peso configurabile
- Ottimizzazione automatica con CP-SAT (equita' + penalita' sovrapposizione sabato)
- **Settimane bloccabili** per pianificazione incrementale
- **Modifica manuale post-solver** con validazione vincoli in tempo reale
- Annullamento solver in corso con recupero soluzione parziale
- Export risultati in **TXT**, **CSV**, **DOCX** (layout professionale landscape) e **ICS** (iCalendar)
- **Copia negli appunti** formato WhatsApp/Telegram
- **Template DOCX personalizzabile** (titolo, colori, font)
- Salvataggio/caricamento sessione in JSON con **sessioni recenti**
- **Auto-salvataggio** periodico con recovery automatico
- Scritture file atomiche con locking distribuito

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
turni_v16.py              entry point
turni/
    constants.py          costanti globali (solver, tema, font)
    helpers.py            funzioni pure di utilita'
    validators.py         validazione e canonizzazione dati sessione
    io_utils.py           I/O atomico e locking distribuito
    calendar_utils.py     generazione settimane da calendario
    history.py            memoria storica turni (equita' inter-sessione)
    config.py             configurazione utente e auto-salvataggio
    ics_export.py         export iCalendar e formato WhatsApp
    docx_export.py        generazione documento Word personalizzabile
    solver.py             SolvePhase + TurniSolver (CP-SAT, ruoli, storico, lock, multi-sede)
    ui/
        widgets.py        ScrollableFrame, styled_btn, card
        app.py            TurniApp -- wizard a 3 step
tests/
    test_helpers.py       test moduli core
    test_new_features.py  test nuove funzionalita'
turni_v15.py              versione monolitica originale (riferimento)
```

---

## Test

```bash
python -m unittest discover -s tests -v
```

---

## Configurazione

L'applicazione salva le preferenze in `~/.turni_acustica/`:
- `config.json` -- sessioni recenti, template DOCX, sedi, peso storico
- `history.json` -- conteggi cumulativi turni per equita' a lungo termine
- `autosave.json` -- recovery automatico in caso di crash

---

## Licenza

MIT -- vedi [LICENSE](LICENSE)
