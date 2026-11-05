# Dashboard Analiză 5C

`Dashboard_Analiza_5C.html` este un singur fișier: se deschide direct în browser (dublu-click), fără server.

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
node tools/build-data.js "Analiza 5C.xlsx"
```

---

# NEXUS · Hub Servere

`Hub_Servere.html` este centrul de comandă pentru toate serverele Home Assistant. Tot un singur fișier: îl deschizi în browser (dublu-click) și browserul se conectează direct, prin WebSocket, la fiecare server. Nu există server intermediar.

## Primii pași
1. Deschide `Hub_Servere.html` → **Adaugă server**.
2. Completezi **numele**, **adresa** (cea din bara browserului când deschizi HA: Nabu Casa, Cloudflare Tunnel, Tailscale sau `http://192.168.x.x:8123`) și un **token de acces de lungă durată**. Token-ul îl creezi în HA din Profil → Securitate → Tokenuri de acces de lungă durată → Creează token.
3. **Testează** verifică imediat conexiunea și dacă utilizatorul e administrator. Pentru restart, add-on-uri, registru și backup e nevoie de un utilizator **admin**.

Fără token-uri poți porni **Mod demo**: apar trei servere simulate (unul offline), ca să vezi toate funcțiile.

## Ce poți face
| Secțiune | Ce conține |
|---|---|
| **Panou** | Constelația live: hub-ul în centru, serverele în jur, fiecare eveniment e o particulă; KPI-uri, carduri cu CPU/RAM/disc, latență, evenimente/min, plus lista „Necesită atenție” (offline, actualizări, reparații, notificări, erori, entități indisponibile). Pe un singur server: versiuni Core/OS/Supervisor, uptime, actualizări cu buton de instalare, notificări, reparații. |
| **Entități** | Toate entitățile de pe toate serverele, cu filtre (domeniu, zonă, stare) și control direct: comutatoare, butoane, termostat, jaluzele, yale, media, aspirator, valori, liste. Click pe o entitate deschide detaliile: istoric 24 h, atribute și editare în registru (nume, pictogramă, zonă, ID, ascunsă/dezactivată). |
| **Automatizări** | Automatizări, scripturi și scene: activare/dezactivare, rulare, editare YAML. |
| **Add-on-uri** | Pornire, oprire, restart, actualizare, consum CPU/RAM și **Deschide UI** (VS Code, terminal, Zigbee2MQTT…). |
| **Camere** | Imagini live, reîmprospătate la 10 s. |
| **Loguri** | Jurnalul de sistem al fiecărui server, filtrabil, cu traceback; golire. |
| **Backup** | Lista backup-urilor și **Backup complet acum**. |
| **Consolă** | Șabloane Jinja, acțiuni YAML sau comenzi WebSocket brute, rulate pe unul sau mai multe servere deodată. |
| **Evenimente** | Fluxul live al schimbărilor de stare, al automatizărilor declanșate și al acțiunilor apelate (cu utilizatorul care le-a apelat). |

**Ctrl K** (sau `/`) deschide căutarea globală: scrii „bec bucătărie” și **Enter** îl comută, pe orice server. **Shift+Enter** deschide detaliile. Tot de aici găsești comenzi precum restart, backup sau verificarea configurației.

Acțiunile riscante (restart, repornirea gazdei, oprirea unui add-on, instalarea unei actualizări, descuierea unei yale, redenumirea ID-ului) cer confirmare. Repornirea gazdei cere și să scrii numele serverului.

## Securitate și limitări
- Token-urile se păstrează doar în browserul tău (localStorage). Dacă debifezi „Ține minte”, token-ul rămâne doar până închizi tab-ul. Deschide hub-ul doar pe dispozitive de încredere.
- O pagină deschisă prin `https://` nu se poate conecta la o adresă `http://` (browserul o blochează). Deschide fișierul local sau folosește adresa https a serverului.
- **Salvarea YAML** a automatizărilor folosește API-ul REST, pe care HA îl permite doar de la origini aprobate. Dacă salvarea e blocată, ai trei variante: copiezi `Hub_Servere.html` în `/config/www/` și îl deschizi ca `https://serverul-tau/local/Hub_Servere.html`; adaugi originea la `http: cors_allowed_origins`; sau folosești butonul „Editorul HA”. Restul funcțiilor merg prin WebSocket, fără configurare.
- Add-on-urile, backup-ul prin Supervisor și repornirea gazdei există doar pe Home Assistant OS / Supervised.
