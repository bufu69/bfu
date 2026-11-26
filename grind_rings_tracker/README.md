# Grind Rings Action Tracker

`static/index.html` este o copie 1:1 a modelului `past_due_tracker.html`. Diferă doar numele („Grind Rings Action Tracker”, în titlul ferestrei și în antet). Serverul Python o servește pe portul **8092**.

Fiind identică, pagina folosește aceleași adrese ca modelul:
- Excelul: `http://romhdata03/timkentraining/PD_PO_Action_Tracker.xlsm` (foaia `ActionItemsPastDue`)
- ascultătorul pentru adăugare/editare/ștergere/atașamente: `http://10.136.82.208:8791/` (`Start-ActionListener.ps1`)
- atașamentele: `http://romhdata03/timkentraining/attachments/`
- ticker-ul: `http://romhdata03/timkentraining/Dashboard2026.xlsm`

Serverul are în plus propriile rute de scriere în Excel (vezi mai jos), dar pagina nu le folosește cât timp adresele din ea sunt cele ale modelului.
## Pornire
```
pip install openpyxl
python grind_rings_server.py
```
(pe Windows: dublu-click pe `start_server.bat`). Apoi deschide `http://<calculator>:8092/`.

### Pornire automată (Windows)
`install_autostart.bat` (click dreapta → *Run as administrator*) creează o sarcină în Task Scheduler care pornește serverul la fiecare pornire a calculatorului, chiar fără utilizator logat, și deschide portul 8092 în firewall. `uninstall_autostart.bat` o elimină.

Atenție: sarcina rulează ca SYSTEM. Dacă Excelul e pe un share de rețea (cale UNC), contul SYSTEM s-ar putea să nu aibă acces. În cazul ăsta, schimbă în Task Scheduler contul sarcinii cu un utilizator care are drepturi pe share.

La prima pornire, dacă fișierul Excel nu există, se creează `Grind_Rings_Action_Tracker.xlsx` cu foile `ActionItems` și `Lists`.

## Setări – `config.json`
Copiază `config.example.json` în `config.json` și modifică:

| Cheie | Ce face |
|---|---|
| `port` | 8092 |
| `excel_file` | Excelul cu acțiuni (.xlsx sau .xlsm – macro-urile se păstrează). Poate fi cale UNC. |
| `sheet_name` | Foaia cu acțiuni (implicit `ActionItems`) |
| `lists_sheet_name` | Foaia cu listele pentru formular (`Client`, `Responsabil`, `Tip Actiune`, `Familie Rulment`, `Tip component`) |
| `attachments_dir` | Unde se salvează atașamentele |
| `ticker_source` | `Dashboard2026.xlsm` pentru bara ticker (URL sau cale); gol = dezactivat |
| `backup_dir`, `backup_keep_days` | Backup zilnic al Excelului (înainte de prima modificare din zi); se păstrează ultimele 30 |
| `pins` | Codurile personale: `"cod": "Nume"`. Numele apare în comentarii. |

Coloanele sunt recunoscute după nume, la fel ca în dashboard (No, Data adaugare, SAP id Rulment, …, Stadiu, Urgent Client, Comentarii); coloanele lipsă se adaugă automat la final.

## Coada CSV
Dacă la salvare serverul nu răspunde, pagina poate salva acțiunea într-un `pending_new_actions.csv` (în Chrome/Edge). Importul în Excel se face cu:
```
python grind_rings_server.py --import-csv cale\catre\pending_new_actions.csv
```
După import, fișierul este redenumit (`.importat_<data>`), ca să nu fie importat de două ori.

**Important:** când serverul scrie, fișierul Excel nu trebuie să fie deschis în Excel de altcineva.

## Rute
| Metodă | Rută | Rol |
|---|---|---|
| GET | `/` | dashboard-ul |
| GET | `/data/tracker` | Excelul cu acțiuni |
| GET | `/data/ticker` | ticker-ul (proxy) |
| POST | `/` | acțiune nouă → `{ok, no}` |
| POST | `/edit` | editare (comentariul nou se adaugă la cele vechi) |
| POST | `/delete` | ștergere |
| POST | `/upload` | atașament (base64) |
| GET | `/list?no=&source=` | lista atașamentelor |
| GET | `/attachments/<source>/<no>/<fișier>` | descărcare atașament |
