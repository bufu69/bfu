# Grind Rings Action Tracker

Dashboard-ul are același aspect ca modelul `past_due_tracker.html` (aceleași grafice, filtre, formular, editare, ștergere cu cod, atașamente), cu numele „Grind Rings Action Tracker”.

**Server complet independent:**
- acțiunile se țin într-o bază de date **SQLite** proprie (`grind_rings.db`), fără Excel, fără `Start-ActionListener.ps1`;
- nu depinde de Past Due, de romhdata03 sau de alt server; atașamentele sunt în folderul propriu `attachments/grindrings/`;
- nu are nevoie de internet: fonturile sunt incluse în `static/fonts/`, nu se mai încarcă nimic din CDN;
- ticker-ul din antet arată indicatorii acestui tracker (Active, Întârziate, Urgent), calculați din baza de date.

Totul rulează dintr-un singur folder, pe portul **8092**. Folderul se poate copia pe orice calculator cu Python 3.

## Pornire
Trebuie doar Python 3 (fără pachete suplimentare).
```
python grind_rings_server.py
```
Pe Windows: dublu-click pe `start_server.bat`. Apoi deschide `http://<calculator>:8092/`.

La prima pornire se creează automat baza de date `grind_rings.db`.

### Pornire automată (Windows)
`install_autostart.bat` (click dreapta → *Run as administrator*) creează o sarcină în Task Scheduler care pornește serverul la fiecare pornire a calculatorului, chiar fără utilizator logat, și deschide portul 8092 în firewall. `uninstall_autostart.bat` o elimină.

## Setări – `config.json`
Copiază `config.example.json` în `config.json` și modifică:

| Cheie | Ce face |
|---|---|
| `port` | 8092 |
| `database` | Fișierul bazei de date (implicit `grind_rings.db`, lângă server) |
| `pins` | Codurile personale: `"cod": "Nume"`. Numele apare în comentarii și în istoric. |
| `lists` | Valori fixe pentru listele din formular (Client, Responsabil, Tip Actiune, Familie Rulment, Tip component). Pe lângă ele apar automat valorile deja folosite. |
| `attachments_dir` | Unde se salvează atașamentele |
| `backup_dir`, `backup_keep_days` | Backup zilnic al bazei de date (înainte de prima modificare din zi); se păstrează ultimele 30 |

## Baza de date
- Tabela `actions`: câte un rând pe acțiune (aceleași câmpuri ca în Excelul vechi + cine/când a creat și modificat).
- Tabela `history`: fiecare adăugare, editare și ștergere, cu utilizatorul, ora și ce s-a schimbat. O acțiune ștearsă din greșeală se poate recupera de aici sau din `backups/`.
- Export: `http://<calculator>:8092/api/export.csv` descarcă toate acțiunile (se deschide în Excel dacă e nevoie de un raport).
- Istoric: `http://<calculator>:8092/api/history` (sau `?no=12` pentru o singură acțiune).

## Import CSV
Pentru a aduce acțiuni dintr-un CSV (de exemplu coada `pending_new_actions.csv` salvată de pagină, sau acțiuni exportate din Excelul vechi ca CSV, cu aceleași nume de coloane):
```
python grind_rings_server.py --import-csv cale\catre\fisier.csv
```
După import, fișierul este redenumit (`.importat_<data>`), ca să nu fie importat de două ori.

## Rute
| Metodă | Rută | Rol |
|---|---|---|
| GET | `/` | dashboard-ul |
| GET | `/api/actions` | toate acțiunile (JSON) |
| GET | `/api/lists` | listele pentru formular |
| GET | `/api/export.csv` | export CSV |
| GET | `/api/history` | istoricul modificărilor |
| GET | `/static/fonts/...` | fonturile paginii |
| POST | `/` | acțiune nouă → `{ok, no}` |
| POST | `/edit` | editare (comentariul nou se adaugă la cele vechi) |
| POST | `/delete` | ștergere |
| POST | `/upload` | atașament (base64) |
| GET | `/list?no=&source=` | lista atașamentelor |
| GET | `/attachments/<source>/<no>/<fișier>` | descărcare atașament |
