# Software Turni Acustica

Applicazione desktop per la **pianificazione turni audio/video** delle sedi di Messina e Ganzirri.

Ottimizza l'assegnazione degli operatori ai turni (giovedi audio/video + domenica) usando il solver **CP-SAT** di Google OR-Tools, garantendo distribuzione equa dei turni e rispettando le indisponibilita' settimanali.

---

## Installazione su EndeavourOS / Arch Linux

### Metodo rapido (consigliato)

```bash
git clone https://github.com/giovanni-paolilla/software_turni_acustica.git
cd software_turni_acustica
sudo bash install.sh
```

L'applicazione comparira' nel menu come **"Gestione Turni"** e sara' avviabile da terminale con `turni-acustica`.

### Disinstallazione

```bash
sudo bash uninstall.sh
```

### Metodo manuale (senza installazione)

```bash
sudo pacman -S python tk
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python turni_v16.py
```

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
    solver.py             SolvePhase + TurniSolver (CP-SAT)
    ui/
        widgets.py        ScrollableFrame, styled_btn, card (CustomTkinter)
        app.py            TurniApp -- wizard a 3 step
packaging/
    turni-acustica.desktop    desktop entry per il menu applicazioni
    turni-acustica.svg        icona applicazione
    turni-acustica.sh         launcher script
    PKGBUILD                  pacchetto Arch Linux
tests/
    test_helpers.py           test moduli core
    test_new_features.py      test nuove funzionalita'
install.sh                    installazione rapida EndeavourOS/Arch
uninstall.sh                  disinstallazione pulita
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
