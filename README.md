# Dashboard Analiză 5C

`Dashboard_Analiza_5C.html` este un singur fișier: se deschide direct în browser (dublu-click), fără server.

## Actualizare săptămânală
1. În `Analiza 5C.xlsx` adaugi o foaie nouă la final, cu numele `5C 2026 wNN` (aceeași structură: antetul cu `NR.` pe rândul 4).
2. În dashboard: **Actualizare din fișier Excel** → trage fișierul în chenar sau apasă **Alege fișier**.
3. Dashboard-ul afișează ce s-a schimbat (foi noi, 5C noi, statusuri modificate) și păstrează datele în browser până la următoarea încărcare. **Revino la datele inițiale** șterge datele încărcate.

Opțional, dacă Excelul e publicat pe un server web, poți salva adresa lui în **Actualizare de pe server** (cu recitire automată la 30 s).

## Toate graficele
Tab-ul **Toate graficele** conține peste 70 de grafice generate din fișier, grupate pe secțiuni: evoluție în timp, status, componente și produse, arii, defecte și cauze, responsabili, timp până la prezentare, completarea datelor, registrul RNC lunar, istoric și organizare. Toate respectă filtrele din bara de sus; poți alege 2, 3 sau 4 coloane.

## Scroll automat (afișare pe ecran / TV)
Din bara de sus: **Scroll auto: ON/OFF** pornește sau oprește derularea, iar lista de lângă alege intervalul (15, 30 sau 45 sec). La fiecare pas pagina coboară la secțiunea următoare; la final revine sus. Cu bifa **tab-uri**, la final trece la tab-ul următor. Setările sunt ținute minte în browser.

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
node tools/build-data.js "Analiza 5C.xlsx"
```
