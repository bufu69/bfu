#!/usr/bin/env python3
"""
Safety Dashboard - server (port 8093), baza de date SQLite.

Server independent: Good Catch, Echipamente nonstandard, Actiuni corective, Incidente.
Totul se editeaza din pagina (dupa autentificare). Doar Python 3, fara pachete externe.

  python server.py                         porneste serverul
  python server.py --utilizatori           administrare utilizatori (utilizator + parola)
  python server.py --import-excel F.xlsm   importa datele din "Dashboard safety.xlsm"
"""

import argparse
import base64
import getpass
import hashlib
import json
import mimetypes
import os
import re
import secrets
import sqlite3
import sys
import threading
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote, urlparse

# ============================ SETARI ============================
APP_NAME = "Safety Dashboard"
HOST = "0.0.0.0"
PORT = 8093
DATABASE = "safety.db"
IMAGINI_DIR = "imagini"
BACKUP_DIR = "backups"
BACKUP_ZILE = 30
MAX_IMAGINE_MB = 15
# ===============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def cale(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


DB_PATH = cale(DATABASE)
db_lock = threading.Lock()

# ---------------------------------------------------------------- entitati
# camp: (cheie, eticheta, tip, obligatoriu)   tip: text | textarea | date | bool | select:opt1|opt2
ENTITATI = {
    "gc": {
        "tabel": "good_catch", "nume": "Good Catch",
        "campuri": [
            ("dept", "Dept", "text", True),
            ("descriere", "Descriere", "textarea", True),
            ("gc_id", "ID", "text", False),
            ("data", "Data înregistrării", "date", True),
            ("responsabil", "Assigned To", "text", True),
            ("termen", "Closure Due Date", "date", True),
            ("inchis", "Închis", "bool", False),
        ],
    },
    "eq": {
        "tabel": "echipamente", "nume": "Echipament nonstandard",
        "campuri": [
            ("sectie", "Secție", "text", True),
            ("echipament", "Echipament în condiții nonstandard", "text", True),
            ("motiv", "Motivul", "textarea", True),
            ("status", "Status", "select:in functiune|oprit|in reparatie", False),
            ("inregistrat", "Registered", "date", True),
            ("termen", "Due date rezolvare", "date", True),
            ("rezolvat", "Rezolvat", "bool", False),
        ],
    },
    "ac": {
        "tabel": "actiuni", "nume": "Acțiune corectivă",
        "campuri": [
            ("descriere", "Descriere acțiune", "textarea", True),
            ("responsabil", "Responsabil", "text", True),
            ("termen", "Termen", "date", False),
            ("inchis", "Închisă", "bool", False),
        ],
    },
    "inc": {
        "tabel": "incidente", "nume": "Incident",
        "campuri": [
            ("tip", "Tip incident", "text", True),
            ("descriere", "Descriere incident", "textarea", True),
            ("data", "Data", "date", False),
            ("inchis", "Acțiuni închise", "bool", False),
        ],
    },
}

TEXTE_IMPLICIT = {
    "companie": "TIMKEN",
    "locatie": "Ploiești, România",
    "slogan_dreapta": "MAKE SAFETY|A HABIT",
    "banda": "SAFETY FIRST • SEE IT • OWN IT • FIX IT",
    "stanga_jos": "KEEP|PEOPLE SAFE|EVERYDAY",
    "dreapta_jos": "A SAFER|TOMORROW|TOGETHER",
    "nota_jos": "Editare: apasă DETALII, apoi Adaugă / Modifică. Datele sunt salvate pe server și apar pe toate ecranele.",
}
IMAGINI = {"fundal": "Fundal pagină", "stanga": "Imagine panou stânga (jos)",
           "dreapta": "Imagine panou dreapta (jos)", "logo": "Logo (în locul textului)"}


# ---------------------------------------------------------------- utilitare
def fara_diacritice(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def iso(d):
    if d is None or d == "":
        return None
    if isinstance(d, datetime):
        return d.date().isoformat()
    if isinstance(d, date):
        return d.isoformat()
    if isinstance(d, (int, float)) and 20000 < d < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=float(d))).date().isoformat()
    s = str(d).strip()
    for fmt, txt in (("%Y-%m-%d", s[:10]), ("%d.%m.%Y", s), ("%d/%m/%Y", s), ("%m/%d/%Y", s)):
        try:
            return datetime.strptime(txt, fmt).date().isoformat()
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})\d*", s)      # ex. "15.12.20206" (greseala de tastare)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            pass
    return None


def acum():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def da(v):
    return fara_diacritice(v) in ("1", "da", "true", "yes", "x", "on", "inchis", "rezolvat", "closed")


class Eroare(Exception):
    def __init__(self, mesaj, status=400):
        super().__init__(mesaj)
        self.status = status


# ---------------------------------------------------------------- baza de date
def db():
    c = sqlite3.connect(DB_PATH, timeout=15)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c


def creeaza_db():
    with db() as c:
        for e in ENTITATI.values():
            cols = ", ".join(f"{k} TEXT" for k, *_ in e["campuri"])
            c.execute(f"""CREATE TABLE IF NOT EXISTS {e['tabel']} (
                id INTEGER PRIMARY KEY AUTOINCREMENT, {cols},
                data_inchidere TEXT, creat_de TEXT, creat_la TEXT, modificat_de TEXT, modificat_la TEXT)""")
        c.executescript("""
CREATE TABLE IF NOT EXISTS istoric (id INTEGER PRIMARY KEY AUTOINCREMENT, entitate TEXT, rid INTEGER,
  operatie TEXT, autor TEXT, la TEXT, detalii TEXT);
CREATE TABLE IF NOT EXISTS utilizatori (utilizator TEXT PRIMARY KEY COLLATE NOCASE, nume TEXT, sare TEXT, hash TEXT,
  rol TEXT DEFAULT 'editor', activ INTEGER DEFAULT 1, creat_la TEXT, ultima_logare TEXT);
CREATE TABLE IF NOT EXISTS sesiuni (token_hash TEXT PRIMARY KEY, utilizator TEXT, creata_la TEXT, expira TEXT);
CREATE TABLE IF NOT EXISTS setari (cheie TEXT PRIMARY KEY, valoare TEXT);
""")


def citeste_setare(k, implicit=None):
    with db() as c:
        r = c.execute("SELECT valoare FROM setari WHERE cheie=?", (k,)).fetchone()
    return json.loads(r["valoare"]) if r else implicit


def scrie_setare(k, v):
    with db() as c:
        c.execute("INSERT OR REPLACE INTO setari(cheie, valoare) VALUES (?,?)", (k, json.dumps(v, ensure_ascii=False)))


def backup_zilnic():
    if not os.path.exists(DB_PATH):
        return
    d = cale(BACKUP_DIR)
    os.makedirs(d, exist_ok=True)
    tinta = os.path.join(d, f"safety_{date.today():%Y-%m-%d}.db")
    if not os.path.exists(tinta):
        with db() as s, sqlite3.connect(tinta) as t:
            s.backup(t)
    vechi = sorted(f for f in os.listdir(d) if re.fullmatch(r"safety_\d{4}-\d\d-\d\d\.db", f))
    for f in vechi[:-BACKUP_ZILE]:
        try:
            os.remove(os.path.join(d, f))
        except OSError:
            pass


def jurnal(c, ent, rid, op, autor, detalii):
    c.execute("INSERT INTO istoric(entitate, rid, operatie, autor, la, detalii) VALUES (?,?,?,?,?,?)",
              (ent, rid, op, autor, acum(), json.dumps(detalii, ensure_ascii=False)))


# ---------------------------------------------------------------- utilizatori
SESIUNE_COOKIE = "sd_sesiune"
ROLURI = {"admin": "Administrator", "editor": "Editor"}


def _hash(p, sare):
    return hashlib.pbkdf2_hmac("sha256", str(p).encode(), bytes.fromhex(sare), 200_000).hex()


def numar_utilizatori():
    with db() as c:
        return c.execute("SELECT COUNT(*) FROM utilizatori").fetchone()[0]


def lista_utilizatori():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT utilizator, nume, rol, activ, creat_la, ultima_logare FROM utilizatori ORDER BY utilizator")]


def creeaza_utilizator(ut, nume, parola, rol="editor"):
    ut = str(ut or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9._-]{2,40}", ut):
        raise Eroare("Utilizatorul: 2–40 caractere (litere, cifre, punct, minus, _).")
    if len(str(parola or "")) < 6:
        raise Eroare("Parola trebuie să aibă cel puțin 6 caractere.")
    if rol not in ROLURI:
        raise Eroare("Rol necunoscut.")
    with db() as c:
        if c.execute("SELECT 1 FROM utilizatori WHERE utilizator=?", (ut,)).fetchone():
            raise Eroare(f"Utilizatorul „{ut}” există deja.")
        s = secrets.token_hex(16)
        c.execute("INSERT INTO utilizatori(utilizator, nume, sare, hash, rol, activ, creat_la) VALUES (?,?,?,?,?,1,?)",
                  (ut, str(nume or "").strip() or ut, s, _hash(parola, s), rol, acum()))
    return ut


def schimba_parola(ut, parola):
    if len(str(parola or "")) < 6:
        raise Eroare("Parola trebuie să aibă cel puțin 6 caractere.")
    s = secrets.token_hex(16)
    with db() as c:
        if not c.execute("UPDATE utilizatori SET sare=?, hash=? WHERE utilizator=?", (s, _hash(parola, s), ut)).rowcount:
            raise Eroare("Utilizatorul nu există.", 404)
        c.execute("DELETE FROM sesiuni WHERE utilizator=?", (ut,))


def modifica_utilizator(ut, nume=None, rol=None, activ=None):
    with db() as c:
        r = c.execute("SELECT * FROM utilizatori WHERE utilizator=?", (ut,)).fetchone()
        if not r:
            raise Eroare("Utilizatorul nu există.", 404)
        if r["rol"] == "admin" and r["activ"] and ((rol and rol != "admin") or activ is False):
            if c.execute("SELECT COUNT(*) FROM utilizatori WHERE rol='admin' AND activ=1").fetchone()[0] <= 1:
                raise Eroare("Trebuie să rămână cel puțin un administrator activ.")
        if nume is not None:
            c.execute("UPDATE utilizatori SET nume=? WHERE utilizator=?", (str(nume).strip() or ut, ut))
        if rol is not None:
            if rol not in ROLURI:
                raise Eroare("Rol necunoscut.")
            c.execute("UPDATE utilizatori SET rol=? WHERE utilizator=?", (rol, ut))
        if activ is not None:
            c.execute("UPDATE utilizatori SET activ=? WHERE utilizator=?", (1 if activ else 0, ut))
            if not activ:
                c.execute("DELETE FROM sesiuni WHERE utilizator=?", (ut,))


def sterge_utilizator(ut):
    modifica_utilizator(ut, activ=False)
    with db() as c:
        c.execute("DELETE FROM utilizatori WHERE utilizator=?", (ut,))


def verifica_parola(ut, parola):
    with db() as c:
        r = c.execute("SELECT * FROM utilizatori WHERE utilizator=?", (str(ut or "").strip(),)).fetchone()
    ok = secrets.compare_digest(_hash(parola or "", r["sare"] if r else "00" * 16), r["hash"] if r else "x")
    return dict(r) if (r and ok and r["activ"]) else None


def creeaza_sesiune(ut, tine_minte):
    tok = secrets.token_urlsafe(32)
    dur = timedelta(days=180) if tine_minte else timedelta(hours=12)
    with db() as c:
        c.execute("DELETE FROM sesiuni WHERE expira < ?", (acum(),))
        c.execute("INSERT INTO sesiuni VALUES (?,?,?,?)", (hashlib.sha256(tok.encode()).hexdigest(), ut, acum(),
                                                           (datetime.now() + dur).strftime("%Y-%m-%d %H:%M:%S")))
        c.execute("UPDATE utilizatori SET ultima_logare=? WHERE utilizator=?", (acum(), ut))
    return tok, dur


def din_token(tok):
    if not tok:
        return None
    with db() as c:
        r = c.execute("""SELECT u.* FROM sesiuni s JOIN utilizatori u ON u.utilizator=s.utilizator
                         WHERE s.token_hash=? AND s.expira>? AND u.activ=1""",
                      (hashlib.sha256(tok.encode()).hexdigest(), acum())).fetchone()
    return dict(r) if r else None


def public(u):
    return {"utilizator": u["utilizator"], "nume": u["nume"], "rol": u["rol"], "rol_nume": ROLURI.get(u["rol"])} if u else None


def meniu_utilizatori():
    while True:
        print(f"\n===== {APP_NAME} - UTILIZATORI =====")
        for u in lista_utilizatori():
            print(f"  {u['utilizator']:<18} {u['nume']:<24} {ROLURI.get(u['rol']):<14} {'' if u['activ'] else '(dezactivat)'}")
        print("\n  1 = utilizator nou   2 = schimba parola   3 = sterge   0 = iesire")
        o = input("> ").strip()
        try:
            if o == "1":
                ut = input("Utilizator: ").strip()
                nume = input("Nume afisat: ").strip()
                rol = "admin" if not numar_utilizatori() or input("Administrator? D/N: ").lower().startswith("d") else "editor"
                p = getpass.getpass("Parola (min. 6): ")
                creeaza_utilizator(ut, nume, p, rol)
                print("OK.")
            elif o == "2":
                schimba_parola(input("Utilizator: ").strip(), getpass.getpass("Parola noua: "))
                print("OK.")
            elif o == "3":
                sterge_utilizator(input("Utilizator: ").strip())
                print("Sters.")
            elif o in ("0", ""):
                return
        except Eroare as e:
            print(e)


# ---------------------------------------------------------------- date
def calculeaza(ent, r, azi):
    d = dict(r)
    if ent == "gc":
        t = iso(d["termen"])
        inchis = da(d["inchis"])
        # ca in Excel: termenul de azi conteaza deja ca Overdue
        d["status"] = "Închis" if inchis else ("Overdue" if t and t <= azi else "Open")
        start, sf = iso(d["data"]), (iso(d["data_inchidere"]) if inchis else azi)
        d["zile_deschis"] = (date.fromisoformat(sf) - date.fromisoformat(start)).days if start and sf else None
        d["zile_termen"] = (date.fromisoformat(t) - date.fromisoformat(azi)).days if t else None
    elif ent == "eq":
        t = iso(d["termen"])
        d["depasit"] = (not da(d["rezolvat"])) and bool(t and t < azi)
    elif ent in ("ac",):
        t = iso(d["termen"])
        d["depasit"] = (not da(d["inchis"])) and bool(t and t < azi)
    for k in ("inchis", "rezolvat"):
        if k in d:
            d[k] = da(d[k])
    return d


def toate_datele():
    azi = date.today().isoformat()
    rez = {}
    with db() as c:
        for ent, e in ENTITATI.items():
            rez[ent] = [calculeaza(ent, r, azi) for r in c.execute(f"SELECT * FROM {e['tabel']} ORDER BY id")]
    gc, eq, ac, inc = rez["gc"], rez["eq"], rez["ac"], rez["inc"]
    rez["indicatori"] = {
        "gc_deschise": sum(1 for x in gc if not x["inchis"]),
        "gc_overdue": sum(1 for x in gc if x["status"] == "Overdue"),
        "actiuni": sum(1 for x in ac if not x["inchis"]),
        "echipamente": sum(1 for x in eq if not x["rezolvat"]),
        "echipamente_depasite": sum(1 for x in eq if x["depasit"]),
        "incidente": sum(1 for x in inc if not x["inchis"]),
    }
    rez["azi"] = azi
    rez["actualizat"] = citeste_setare("ultima_modificare")
    return rez


def curata(ent, d, partial=False):
    v = {}
    for k, et, tip, obl in ENTITATI[ent]["campuri"]:
        if k not in d:
            if not partial and obl:
                raise Eroare(f"Completează „{et}”.")
            continue
        x = d[k]
        if tip == "date":
            x = iso(x)
            if x is None and str(d[k] or "").strip():
                raise Eroare(f"Data invalidă la „{et}”.")
        elif tip == "bool":
            x = "1" if (x is True or da(x)) else "0"
        else:
            x = str(x or "").strip()[:4000]
            if tip.startswith("select:") and x and x not in tip[7:].split("|"):
                pass  # se accepta si alte valori (ex. din Excel)
        if obl and not x:
            raise Eroare(f"Completează „{et}”.")
        v[k] = x
    return v


def salveaza(ent, rid, d, autor):
    if ent not in ENTITATI:
        raise Eroare("Tip necunoscut.", 404)
    e = ENTITATI[ent]
    with db_lock:
        backup_zilnic()
        with db() as c:
            if rid is None:
                v = curata(ent, d)
                for k in ("inchis", "rezolvat"):
                    if v.get(k) == "1":
                        v["data_inchidere"] = date.today().isoformat()
                v.update(creat_de=autor, creat_la=acum(), modificat_de=autor, modificat_la=acum())
                cur = c.execute(f"INSERT INTO {e['tabel']} ({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
                rid = cur.lastrowid
                jurnal(c, ent, rid, "adăugare", autor, v)
            else:
                r = c.execute(f"SELECT * FROM {e['tabel']} WHERE id=?", (rid,)).fetchone()
                if not r:
                    raise Eroare("Înregistrarea nu există.", 404)
                v = curata(ent, d, partial=True)
                for k in ("inchis", "rezolvat"):
                    if k in v and v[k] != (r[k] or "0"):
                        v["data_inchidere"] = date.today().isoformat() if v[k] == "1" else None
                sch = {k: {"din": r[k], "in": x} for k, x in v.items() if (r[k] or None) != (x or None)}
                if sch:
                    v.update(modificat_de=autor, modificat_la=acum())
                    c.execute(f"UPDATE {e['tabel']} SET {','.join(f'{k}=?' for k in v)} WHERE id=?", [*v.values(), rid])
                    jurnal(c, ent, rid, "modificare", autor, sch)
    scrie_setare("ultima_modificare", acum())
    return rid


def sterge(ent, rid, autor):
    e = ENTITATI.get(ent)
    if not e:
        raise Eroare("Tip necunoscut.", 404)
    with db_lock:
        backup_zilnic()
        with db() as c:
            r = c.execute(f"SELECT * FROM {e['tabel']} WHERE id=?", (rid,)).fetchone()
            if not r:
                raise Eroare("Înregistrarea nu există.", 404)
            c.execute(f"DELETE FROM {e['tabel']} WHERE id=?", (rid,))
            jurnal(c, ent, rid, "ștergere", autor, dict(r))
    scrie_setare("ultima_modificare", acum())


def istoric(limita=200):
    with db() as c:
        return [{**dict(r), "detalii": json.loads(r["detalii"] or "null")}
                for r in c.execute("SELECT * FROM istoric ORDER BY id DESC LIMIT ?", (limita,))]


# ---------------------------------------------------------------- import Excel (fara pachete externe)
def citeste_xlsx(sursa):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(sursa) as z:
        comune = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                comune.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        foi = {}
        for s in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", ns):
            t = rels[s.get(f"{{{ns['r']}}}id")].lstrip("/")
            t = t if t.startswith("xl/") else "xl/" + t
            rr = []
            for row in ET.fromstring(z.read(t)).iter(f"{{{ns['m']}}}row"):
                cel = {}
                for c in row.findall("m:c", ns):
                    col = 0
                    for ch in re.match(r"[A-Z]+", c.get("r", "A")).group(0):
                        col = col * 26 + ord(ch) - 64
                    tip, v = c.get("t"), c.find("m:v", ns)
                    if tip == "s" and v is not None:
                        val = comune[int(v.text)]
                    elif tip == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(f"{{{ns['m']}}}t"))
                    elif tip in ("str", "e"):
                        val = v.text if v is not None else ""
                    elif v is not None and v.text:
                        f = float(v.text)
                        val = int(f) if f.is_integer() else f
                    else:
                        val = ""
                    cel[col - 1] = val
                rr.append([cel.get(i, "") for i in range(max(cel) + 1)] if cel else [])
            foi[s.get("name").strip()] = rr
        return foi


def _foaie(foi, *nume):
    for n in foi:
        if fara_diacritice(n) in [fara_diacritice(x) for x in nume]:
            return foi[n]
    return []


def _tabel(randuri, cuvant):
    """Gaseste randul de antet (care contine `cuvant`) si intoarce (antet normalizat, randuri date)."""
    for i, r in enumerate(randuri):
        h = [fara_diacritice(x) for x in r]
        if cuvant in h:
            return h, [x for x in randuri[i + 1:] if any(str(v).strip() for v in x)]
    return [], []


def _col(h, r, *nume):
    for n in nume:
        if n in h and h.index(n) < len(r):
            return r[h.index(n)]
    return ""


def importa_excel(sursa, inlocuieste=True, autor="import Excel"):
    foi = citeste_xlsx(sursa)
    date_ = {"gc": [], "eq": [], "ac": [], "inc": []}
    h, rr = _tabel(_foaie(foi, "Status GC deschise"), "dept")
    for r in rr:
        st = fara_diacritice(_col(h, r, "status"))
        date_["gc"].append({"dept": _col(h, r, "dept"), "descriere": _col(h, r, "description"), "gc_id": _col(h, r, "id"),
                            "data": _col(h, r, "dates"), "responsabil": _col(h, r, "assigned to"),
                            "termen": _col(h, r, "closure due date"), "inchis": st in ("closed", "inchis")})
    h, rr = _tabel(_foaie(foi, "Echipamente cond. nonstandard"), "echipament in conditii nonstandard")
    for r in rr:
        date_["eq"].append({"sectie": _col(h, r, "sectie"), "echipament": _col(h, r, "echipament in conditii nonstandard"),
                            "motiv": _col(h, r, "motivul"), "status": _col(h, r, "status"),
                            "inregistrat": _col(h, r, "registred", "registered"), "termen": _col(h, r, "due date rezolvare"),
                            "rezolvat": fara_diacritice(_col(h, r, "status")) in ("rezolvat", "inchis")})
    h, rr = _tabel(_foaie(foi, "Actiuni corective"), "descriere actiune")
    for r in rr:
        date_["ac"].append({"descriere": _col(h, r, "descriere actiune"), "responsabil": _col(h, r, "responsabil"),
                            "termen": _col(h, r, "termen"), "inchis": False})
    h, rr = _tabel(_foaie(foi, "Incidente"), "tip incident")
    for r in rr:
        if not str(_col(h, r, "tip incident")).strip() and not str(_col(h, r, "descriere incident")).strip():
            continue
        date_["inc"].append({"tip": _col(h, r, "tip incident") or "—", "descriere": _col(h, r, "descriere incident") or "—",
                             "inchis": fara_diacritice(_col(h, r, "status actiuni")) in ("inchis", "closed")})
    nr = {}
    with db_lock:
        backup_zilnic()
        with db() as c:
            for ent, lista in date_.items():
                e = ENTITATI[ent]
                if inlocuieste:
                    c.execute(f"DELETE FROM {e['tabel']}")
                n = 0
                for d in lista:
                    try:
                        v = curata(ent, {k: d.get(k, "") for k, *_ in e["campuri"]}, partial=True)
                    except Eroare:
                        continue
                    if not any(v.get(k) for k, _, _, obl in e["campuri"] if obl):
                        continue
                    v.update(creat_de=autor, creat_la=acum(), modificat_de=autor, modificat_la=acum())
                    c.execute(f"INSERT INTO {e['tabel']} ({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
                    n += 1
                nr[ent] = n
            jurnal(c, "import", 0, "import Excel" + (" (înlocuire)" if inlocuieste else ""), autor, nr)
    scrie_setare("ultima_modificare", acum())
    return nr


# ---------------------------------------------------------------- setari pagina (texte + imagini)
def texte():
    t = dict(TEXTE_IMPLICIT)
    t.update(citeste_setare("texte", {}) or {})
    return t


def imagini_setate():
    d = cale(IMAGINI_DIR)
    rez = {}
    for k in IMAGINI:
        f = next((x for x in os.listdir(d) if x.startswith(k + ".")), None) if os.path.isdir(d) else None
        rez[k] = f"imagini/{f}?v={int(os.path.getmtime(os.path.join(d, f)))}" if f else None
    return rez


def salveaza_imagine(cheie, nume, b64):
    if cheie not in IMAGINI:
        raise Eroare("Imagine necunoscută.")
    ext = os.path.splitext(str(nume or ""))[1].lower()
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".svg", ".gif"):
        raise Eroare("Format acceptat: JPG, PNG, WEBP, SVG, GIF.")
    brut = base64.b64decode(b64 or "")
    if not brut or len(brut) > MAX_IMAGINE_MB * 1024 * 1024:
        raise Eroare(f"Imaginea lipsește sau depășește {MAX_IMAGINE_MB} MB.")
    d = cale(IMAGINI_DIR)
    os.makedirs(d, exist_ok=True)
    for x in os.listdir(d):
        if x.startswith(cheie + "."):
            os.remove(os.path.join(d, x))
    with open(os.path.join(d, cheie + ext), "wb") as f:
        f.write(brut)


def sterge_imagine(cheie):
    d = cale(IMAGINI_DIR)
    if os.path.isdir(d):
        for x in os.listdir(d):
            if x.startswith(cheie + "."):
                os.remove(os.path.join(d, x))


# ---------------------------------------------------------------- HTTP
class Server(BaseHTTPRequestHandler):
    server_version = "SafetyDashboard/1.0"

    def log_message(self, fmt, *a):
        sys.stdout.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {self.client_address[0]} {fmt % a}\n")

    def raspuns(self, st, corp, tip, extra=None):
        self.send_response(st)
        self.send_header("Content-Type", tip)
        self.send_header("Content-Length", str(len(corp)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corp)

    def json(self, obj, st=200, extra=None):
        self.raspuns(st, json.dumps(obj, ensure_ascii=False).encode(), "application/json; charset=utf-8", extra)

    def fisier(self, p, tip=None):
        if not os.path.isfile(p):
            return self.json({"ok": False, "eroare": "Nu există."}, 404)
        with open(p, "rb") as f:
            self.raspuns(200, f.read(), tip or mimetypes.guess_type(p)[0] or "application/octet-stream")

    def token(self):
        for b in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = b.strip().partition("=")
            if k == SESIUNE_COOKIE:
                return v
        return None

    def cookie(self, tok, dur=None):
        c = f"{SESIUNE_COOKIE}={tok}; Path=/; HttpOnly; SameSite=Lax"
        return c + (f"; Max-Age={int(dur.total_seconds())}" if dur is not None else "")

    def do_GET(self):
        p = unquote(urlparse(self.path).path)
        try:
            if p in ("/", "/index.html"):
                return self.fisier(cale("index.html"), "text/html; charset=utf-8")
            if p.startswith("/fonts/"):
                f = os.path.basename(p)
                return self.fisier(os.path.join(BASE_DIR, "fonts", f), "text/css; charset=utf-8" if f.endswith(".css") else "font/woff2")
            if p.startswith("/imagini/"):
                f = os.path.basename(p)
                return self.fisier(os.path.join(cale(IMAGINI_DIR), f))
            if p == "/api/date":
                return self.json({"ok": True, **toate_datele()})
            if p == "/api/config":
                u = din_token(self.token())
                return self.json({"ok": True, "eu": public(u), "configurat": numar_utilizatori() > 0,
                                  "texte": texte(), "imagini": imagini_setate(),
                                  "entitati": {k: {"nume": e["nume"], "campuri": e["campuri"]} for k, e in ENTITATI.items()}})
            if p in ("/api/utilizatori", "/api/istoric"):
                u = din_token(self.token())
                if not u or u["rol"] != "admin":
                    raise Eroare("Doar pentru administratori.", 403)
                return self.json({"ok": True, "utilizatori": lista_utilizatori()} if p == "/api/utilizatori"
                                 else {"ok": True, "istoric": istoric()})
            return self.json({"ok": False, "eroare": "Nu există."}, 404)
        except Eroare as e:
            return self.json({"ok": False, "eroare": str(e)}, e.status)
        except Exception as e:  # noqa: BLE001
            return self.json({"ok": False, "eroare": f"{type(e).__name__}: {e}"}, 500)

    def do_POST(self):
        p = urlparse(self.path).path.rstrip("/")
        try:
            orig = self.headers.get("Origin")
            if orig and urlparse(orig).netloc.lower() != (self.headers.get("Host") or "").lower():
                raise Eroare("Cerere respinsă (altă origine).", 403)
            if "application/json" not in (self.headers.get("Content-Type") or ""):
                raise Eroare("Cerere invalidă.")
            n = int(self.headers.get("Content-Length") or 0)
            if n > 40 * 1024 * 1024:
                raise Eroare("Cerere prea mare.", 413)
            d = json.loads(self.rfile.read(n).decode() or "{}")

            if p == "/api/initializare":
                with db_lock:
                    if numar_utilizatori():
                        raise Eroare("Există deja utilizatori.", 403)
                    ut = creeaza_utilizator(d.get("utilizator"), d.get("nume"), d.get("parola"), "admin")
                tok, dur = creeaza_sesiune(ut, True)
                return self.json({"ok": True, "eu": public(din_token(tok))}, extra={"Set-Cookie": self.cookie(tok, dur)})
            if p == "/api/login":
                u = verifica_parola(d.get("utilizator"), d.get("parola"))
                if not u:
                    raise Eroare("Utilizator sau parolă greșită.", 401)
                tok, dur = creeaza_sesiune(u["utilizator"], bool(d.get("tine_minte")))
                return self.json({"ok": True, "eu": public(u)},
                                 extra={"Set-Cookie": self.cookie(tok, dur if d.get("tine_minte") else None)})
            if p == "/api/logout":
                with db() as c:
                    c.execute("DELETE FROM sesiuni WHERE token_hash=?", (hashlib.sha256((self.token() or "").encode()).hexdigest(),))
                return self.json({"ok": True}, extra={"Set-Cookie": self.cookie("", timedelta(0))})

            u = din_token(self.token())
            if not u:
                raise Eroare("Trebuie să te autentifici.", 401)
            autor = u["nume"] or u["utilizator"]

            if p == "/api/parola":
                if not verifica_parola(u["utilizator"], d.get("parola_veche")):
                    raise Eroare("Parola actuală este greșită.", 403)
                schimba_parola(u["utilizator"], d.get("parola_noua"))
                tok, dur = creeaza_sesiune(u["utilizator"], True)
                return self.json({"ok": True}, extra={"Set-Cookie": self.cookie(tok, dur)})

            m = re.fullmatch(r"/api/(gc|eq|ac|inc)(?:/(\d+)(?:/(sterge))?)?", p)
            if m:
                ent, rid, act = m.group(1), m.group(2), m.group(3)
                if act:
                    sterge(ent, int(rid), autor)
                else:
                    rid = salveaza(ent, int(rid) if rid else None, d, autor)
                print(f"  ~ {ent} #{rid} {act or 'salvat'} de {autor}")
                return self.json({"ok": True, "id": rid, **toate_datele()})

            if u["rol"] != "admin":
                raise Eroare("Doar pentru administratori.", 403)
            if p == "/api/import-excel":
                fd = cale("_import.xlsx")
                with open(fd, "wb") as f:
                    f.write(base64.b64decode(d.get("continut") or ""))
                try:
                    nr = importa_excel(fd, bool(d.get("inlocuieste", True)), autor)
                except (zipfile.BadZipFile, KeyError) as e:
                    raise Eroare(f"Fișierul nu pare a fi un Excel valid (.xlsx/.xlsm): {e}")
                finally:
                    os.remove(fd)
                return self.json({"ok": True, "importat": nr, **toate_datele()})
            if p == "/api/texte":
                t = texte()
                for k in TEXTE_IMPLICIT:
                    if k in d:
                        t[k] = str(d[k]).strip()[:300] or TEXTE_IMPLICIT[k]
                scrie_setare("texte", t)
                return self.json({"ok": True, "texte": t})
            if p == "/api/imagine":
                if d.get("sterge"):
                    sterge_imagine(d.get("cheie"))
                else:
                    salveaza_imagine(d.get("cheie"), d.get("nume"), d.get("continut"))
                return self.json({"ok": True, "imagini": imagini_setate()})
            m = re.fullmatch(r"/api/utilizatori(?:/([^/]+)(?:/(parola|sterge))?)?", p)
            if m:
                ut, act = m.group(1), m.group(2)
                if ut is None:
                    creeaza_utilizator(d.get("utilizator"), d.get("nume"), d.get("parola"), d.get("rol") or "editor")
                elif act == "parola":
                    schimba_parola(ut, d.get("parola"))
                elif act == "sterge":
                    if ut.lower() == u["utilizator"].lower():
                        raise Eroare("Nu îți poți șterge propriul cont.")
                    sterge_utilizator(ut)
                else:
                    modifica_utilizator(ut, d.get("nume"), d.get("rol"), d.get("activ"))
                return self.json({"ok": True, "utilizatori": lista_utilizatori()})
            return self.json({"ok": False, "eroare": "Rută necunoscută."}, 404)
        except Eroare as e:
            return self.json({"ok": False, "eroare": str(e)}, e.status)
        except Exception as e:  # noqa: BLE001
            return self.json({"ok": False, "eroare": f"{type(e).__name__}: {e}"}, 500)


def porneste():
    if sys.stdout is None or sys.stderr is None:
        sys.stdout = sys.stderr = open(cale("server.log"), "a", encoding="utf-8", buffering=1)
    srv = ThreadingHTTPServer((HOST, PORT), Server)
    print(f"{APP_NAME} rulează pe http://{HOST}:{PORT}/   (baza de date: {DB_PATH})")
    if not numar_utilizatori():
        print("  Niciun utilizator: deschide pagina în browser și creează contul de administrator.")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("Oprit.")


def main():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--utilizatori", "--editori", dest="ut", action="store_true")
    ap.add_argument("--adauga-utilizator", nargs=3, metavar=("UTILIZATOR", "NUME", "PAROLA"))
    ap.add_argument("--rol", default="editor", choices=list(ROLURI))
    ap.add_argument("--numar-utilizatori", action="store_true")
    ap.add_argument("--import-excel", metavar="FISIER")
    ap.add_argument("--adauga", action="store_true", help="cu --import-excel: adaugă fără să șteargă datele existente")
    a = ap.parse_args()
    creeaza_db()
    if a.ut:
        return meniu_utilizatori()
    if a.adauga_utilizator:
        try:
            creeaza_utilizator(*a.adauga_utilizator, a.rol)
        except Eroare as e:
            raise SystemExit(str(e))
        return print("OK.")
    if a.numar_utilizatori:
        return print(numar_utilizatori())
    if a.import_excel:
        nr = importa_excel(a.import_excel, not a.adauga)
        return print(f"Importat: {nr['gc']} Good Catch, {nr['eq']} echipamente, {nr['ac']} acțiuni corective, {nr['inc']} incidente.")
    porneste()


if __name__ == "__main__":
    main()
