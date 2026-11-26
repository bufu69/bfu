# Aplicații web

Fiecare aplicație stă în folderul ei din `apps/`, iar `server.py` le servește pe toate de pe același server.

```
apps/
  analiza-5c/          -> http://<server>:8080/analiza-5c/
    index.html
    app.json           (opțional: titlu, descriere, iconiță)
  alta-aplicatie/      -> http://<server>:8080/alta-aplicatie/
server.py
porneste-server.bat    (pornire cu dublu-click pe Windows)
```

## Serverul
Are nevoie doar de Python 3.8+ (fără pachete de instalat).

- **Windows:** dublu-click pe `porneste-server.bat`
- **Linux / macOS / Windows din terminal:** `python server.py`

La pornire afișează adresele: `http://localhost:8080/` pe calculatorul serverului și `http://<IP>:8080/` pentru celelalte calculatoare/TV-uri din rețea (Windows poate cere permisiunea în firewall – alege *Rețele private*).

Opțiuni: `--port 9000` (alt port), `--host 127.0.0.1` (doar local, nu din rețea), `--apps <folder>` (alt folder de aplicații), `--quiet` (fără jurnal).

Ce face:
- Pagina principală `/` listează automat toate aplicațiile. Titlul, descrierea și iconița vin din `app.json`; fără el se ia `<title>` din pagină.
- O aplicație nouă = un folder nou în `apps/` cu un `index.html` (sau orice fișier `.html`). Apare imediat, fără repornirea serverului.
- Fișierele modificate (pagini, Excel, JSON, CSV) se văd imediat: browserul le reverifică la fiecare încărcare.
- Folderele și fișierele care încep cu `.` sau `_` sunt ascunse; conținutul folderelor nu se listează; nu se poate ieși din `apps/`.
- `/api/apps` întoarce lista aplicațiilor în JSON.

Exemplu `app.json`:
```json
{ "title": "Dashboard Analiză 5C", "description": "Analiza 5C săptămânală", "icon": "📊" }
```

Notă: toate aplicațiile de pe același server împart `localStorage`-ul browserului, de aceea fiecare aplicație își prefixează cheile (ex. `dashboard5c_…`).

# Dashboard Analiză 5C

`apps/analiza-5c/index.html` este un singur fișier: merge prin server la `/analiza-5c/` sau se poate deschide direct în browser (dublu-click).

**Excel pe server:** copiază `Analiza 5C.xlsx` în `apps/analiza-5c/`, apoi în dashboard, la **Actualizare de pe server**, salvează adresa `Analiza%205C.xlsx` (sau `http://<server>:8080/analiza-5c/Analiza%205C.xlsx`). Cu recitirea automată pornită, înlocuirea fișierului din folder se vede în dashboard în cel mult 30 s.

## Actualizare săptămânală
1. În `Analiza 5C.xlsx` adaugi o foaie nouă la final, cu numele `5C 2026 wNN` (aceeași structură: antetul cu `NR.` pe rândul 4).
2. În dashboard: **Actualizare din fișier Excel** → trage fișierul în chenar sau apasă **Alege fișier**.
3. Dashboard-ul afișează ce s-a schimbat (foi noi, 5C noi, statusuri modificate) și păstrează datele în browser până la următoarea încărcare. **Revino la datele inițiale** șterge datele încărcate.

Opțional, dacă Excelul e publicat pe un server web, poți salva adresa lui în **Actualizare de pe server** (cu recitire automată la 30 s).

## Toate graficele
Tab-ul **Toate graficele** conține peste 70 de grafice generate din fișier, grupate pe secțiuni: evoluție în timp, status, componente și produse, arii, defecte și cauze, responsabili, timp până la prezentare, completarea datelor, registrul RNC lunar, istoric și organizare. Toate respectă filtrele din bara de sus; poți alege 2, 3 sau 4 coloane.

## Ecran complet
Butonul **Ecran complet** pune browserul pe tot ecranul și ascunde antetul, filtrele, panourile de actualizare și meniurile: rămân doar graficele, mărite pe toată suprafața (2 rânduri de grafice pe ecran). În colțul din dreapta-jos rămâne o bară discretă cu tab-ul curent, butoane ‹ › pentru tab-ul anterior/următor, pornire/oprire scroll și **Ieșire** (sau tasta Esc). Se combină cu scroll-ul automat continuu pentru afișare pe TV.

## Scroll automat (afișare pe ecran / TV)
Din bara de sus: **Scroll auto: ON/OFF**, apoi modul și viteza:
- **Continuu** (implicit): pagina coboară lin, fără oprire; la final face o pauză de 4 s, revine sus și o ia de la capăt, în buclă. Viteza: 15, 30 sau 45 sec pentru parcurgerea unui ecran.
- **Pe pași**: la fiecare 15, 30 sau 45 sec sare la secțiunea următoare; la final revine sus.

Cu bifa **tab-uri**, la final trece la tab-ul următor în loc să rămână pe același. Setările sunt ținute minte în browser.

## Ce citește din Excel
| Foaie | Utilizare |
|---|---|
| `5C AAAA wNN` | 5C-urile săptămânale (toate coloanele, inclusiv cauze și contramăsuri) |
| `5C Lun-AAAA` (ex. `5C Sep-2026`) | Registru RNC lunar – tab separat, opțional inclus în analize |
| `Responsabili` (coloana *Associate Name*) | Departament, funcție și manager pentru fiecare responsabil |
| `Graph`, `Sheet1` | Istoric 5C și pivot lună × arie |

## Datele incluse în pagină
Datele afișate la prima deschidere se pot reîmprospăta dintr-un Excel cu:
```
npm i xlsx@0.18.5
node tools/build-data.js "Analiza 5C.xlsx"   # scrie în apps/analiza-5c/index.html
```
