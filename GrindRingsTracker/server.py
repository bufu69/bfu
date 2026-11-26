#!/usr/bin/env python3
"""
Grind Rings Action Tracker - server (port 8092).

Server independent pentru actiunile de la rectificarea inelelor (grind rings).
Date intr-o baza SQLite proprie (grind_rings.db). Doar Python 3, fara pachete externe.

  python server.py                       porneste serverul
  python server.py --editori             administrare editori (cod personal)
  python server.py --import-excel F.xlsx importa actiuni din Excel
  python server.py --trimite-mail        trimite mailurile prin SMTP (mail.ini)
  python server.py --raport-mail         pregateste mailurile pentru Outlook (mail_out/)
"""

import argparse
import base64
import configparser
import csv
import getpass
import hashlib
import html
import io
import json
import mimetypes
import os
import re
import secrets
import smtplib
import sqlite3
import sys
import threading
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

# ============================ SETARI ============================
APP_NAME = "Grind Rings Action Tracker"
HOST = "0.0.0.0"
PORT = 8092
DATABASE = "grind_rings.db"
ATASAMENTE_DIR = "atasamente"
BACKUP_DIR = "backups"
BACKUP_ZILE = 30
MAX_ATASAMENT_MB = 25
MAIL_INI = "mail.ini"
MAIL_OUT_DIR = "mail_out"
# ===============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def cale(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


DB_PATH = cale(DATABASE)

# Campurile unei actiuni: (cheie, eticheta)
CAMPURI = [
    ("nr", "Nr"),
    ("data", "Data"),
    ("masina", "Mașină / Linie"),
    ("reper", "Reper"),
    ("comanda", "Comandă"),
    ("client", "Client"),
    ("problema", "Problemă"),
    ("actiune", "Acțiune"),
    ("responsabil", "Responsabil"),
    ("termen", "Termen"),
    ("stadiu", "Stadiu"),
    ("prioritate", "Prioritate"),
    ("data_inchidere", "Data închidere"),
]
EDITABILE = ["masina", "reper", "comanda", "client", "problema", "actiune",
             "responsabil", "termen", "stadiu", "prioritate"]
DATE = {"data", "termen", "data_inchidere"}
STADII = ["Deschis", "În lucru", "Închis"]
PRIORITATI = ["Normală", "Ridicată", "Urgentă"]
LISTE = ["masina", "client", "responsabil", "reper"]

# recunoasterea coloanelor la importul din Excel (fara diacritice, litere mici)
CUVINTE_CHEIE = {
    "nr": ["nr", "nr.", "no", "numar", "id"],
    "data": ["data", "data adaugare", "data deschidere", "data creare"],
    "masina": ["masina", "masina / linie", "linie", "utilaj", "masina/linie"],
    "reper": ["reper", "cod reper", "cod inel", "inel", "sap"],
    "comanda": ["comanda", "nr comanda", "order"],
    "client": ["client"],
    "problema": ["problema", "defect", "descriere problema"],
    "actiune": ["actiune", "masura", "actiune corectiva"],
    "responsabil": ["responsabil", "owner"],
    "termen": ["termen", "data planificata", "deadline", "data limita"],
    "stadiu": ["stadiu", "status", "stare"],
    "prioritate": ["prioritate", "urgent", "urgenta"],
    "data_inchidere": ["data inchidere", "data finalizare", "inchis la"],
    "comentarii": ["comentarii", "comentariu", "observatii", "note"],
}

db_lock = threading.Lock()


# =============================================================== utilitare
def fara_diacritice(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def iso(d):
    """Orice forma de data -> 'aaaa-ll-zz' (sau None)."""
    if d is None or d == "":
        return None
    if isinstance(d, datetime):
        return d.date().isoformat()
    if isinstance(d, date):
        return d.isoformat()
    if isinstance(d, (int, float)) and 20000 < d < 80000:   # numar serial Excel
        return (datetime(1899, 12, 30) + timedelta(days=float(d))).date().isoformat()
    s = str(d).strip()
    for fmt, txt in (("%Y-%m-%d", s[:10]), ("%d.%m.%Y", s), ("%d/%m/%Y", s), ("%d-%m-%Y", s)):
        try:
            return datetime.strptime(txt, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def ro(d):
    try:
        return datetime.strptime(str(d)[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return ""


def acum():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def stadiu_valid(v):
    s = fara_diacritice(v)
    if re.search(r"inchis|rezolvat|finalizat|done|closed", s):
        return "Închis"
    if re.search(r"lucru|progres|wip", s):
        return "În lucru"
    return "Deschis"


def prioritate_valida(v):
    s = fara_diacritice(v)
    if re.search(r"urgent|^da$|^x$|^1$|critic", s):
        return "Urgentă"
    if re.search(r"ridicat|mare|high", s):
        return "Ridicată"
    return "Normală"


def nume_sigur(s):
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s)).strip().strip(".")
    return s or "fisier"


class Eroare(Exception):
    def __init__(self, mesaj, status=400):
        super().__init__(mesaj)
        self.status = status


# =============================================================== baza de date
def db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def creeaza_db():
    with db() as c:
        c.executescript("""
CREATE TABLE IF NOT EXISTS actiuni (
  nr INTEGER PRIMARY KEY AUTOINCREMENT,
  data TEXT, masina TEXT, reper TEXT, comanda TEXT, client TEXT,
  problema TEXT, actiune TEXT, responsabil TEXT, termen TEXT,
  stadiu TEXT DEFAULT 'Deschis', prioritate TEXT DEFAULT 'Normală', data_inchidere TEXT,
  creat_de TEXT, creat_la TEXT, modificat_de TEXT, modificat_la TEXT
);
CREATE TABLE IF NOT EXISTS comentarii (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nr INTEGER REFERENCES actiuni(nr) ON DELETE CASCADE,
  autor TEXT, la TEXT, text TEXT
);
CREATE TABLE IF NOT EXISTS istoric (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nr INTEGER, operatie TEXT, autor TEXT, la TEXT, detalii TEXT
);
CREATE TABLE IF NOT EXISTS editori (
  nume TEXT PRIMARY KEY, sare TEXT, hash TEXT, creat_la TEXT
);
CREATE TABLE IF NOT EXISTS liste (
  lista TEXT, valoare TEXT, PRIMARY KEY (lista, valoare)
);
CREATE INDEX IF NOT EXISTS ix_com_nr ON comentarii(nr);
CREATE INDEX IF NOT EXISTS ix_ist_nr ON istoric(nr);
""")


def jurnal(c, nr, operatie, autor, detalii):
    c.execute("INSERT INTO istoric(nr, operatie, autor, la, detalii) VALUES (?,?,?,?,?)",
              (nr, operatie, autor, acum(), json.dumps(detalii, ensure_ascii=False)))


def backup_zilnic():
    if not os.path.exists(DB_PATH):
        return
    d = cale(BACKUP_DIR)
    os.makedirs(d, exist_ok=True)
    tinta = os.path.join(d, f"grind_rings_{date.today():%Y-%m-%d}.db")
    if not os.path.exists(tinta):
        with db() as sursa, sqlite3.connect(tinta) as dest:
            sursa.backup(dest)
    vechi = sorted(f for f in os.listdir(d) if re.fullmatch(r"grind_rings_\d{4}-\d\d-\d\d\.db", f))
    for f in vechi[:-BACKUP_ZILE] if BACKUP_ZILE > 0 else []:
        try:
            os.remove(os.path.join(d, f))
        except OSError:
            pass


# =============================================================== editori
def _hash(cod, sare):
    return hashlib.pbkdf2_hmac("sha256", str(cod).encode(), bytes.fromhex(sare), 120_000).hex()


def editor_dupa_cod(cod):
    cod = str(cod or "").strip()
    if not cod:
        return None
    with db() as c:
        for r in c.execute("SELECT nume, sare, hash FROM editori"):
            if secrets.compare_digest(_hash(cod, r["sare"]), r["hash"]):
                return r["nume"]
    return None


def cere_editor(date_cerere):
    nume = editor_dupa_cod(date_cerere.get("cod"))
    if not nume:
        raise Eroare("Cod incorect.", 403)
    return nume


def seteaza_editor(nume, cod):
    nume, cod = nume.strip(), str(cod).strip()
    if not nume or len(cod) < 4:
        raise ValueError("Numele nu poate fi gol, iar codul trebuie sa aiba minim 4 caractere.")
    with db() as c:
        for r in c.execute("SELECT nume, sare, hash FROM editori WHERE nume<>?", (nume,)):
            if _hash(cod, r["sare"]) == r["hash"]:
                raise ValueError(f"Codul este deja folosit de {r['nume']}. Alege alt cod.")
        sare = secrets.token_hex(16)
        c.execute("INSERT OR REPLACE INTO editori(nume, sare, hash, creat_la) VALUES (?,?,?,?)",
                  (nume, sare, _hash(cod, sare), acum()))


def lista_editori():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT nume, creat_la FROM editori ORDER BY nume")]


def meniu_editori():
    while True:
        eds = lista_editori()
        print(f"\n===== {APP_NAME} - EDITORI =====")
        for i, e in enumerate(eds, 1):
            print(f"  {i}. {e['nume']}  (din {e['creat_la'][:10]})")
        if not eds:
            print("  (niciun editor - nimeni nu poate modifica actiunile)")
        print("\n  1 = adauga editor / schimba codul   2 = sterge editor   0 = iesire")
        opt = input("> ").strip()
        if opt == "1":
            nume = input("Nume (apare la comentarii si in istoric): ").strip()
            cod = getpass.getpass("Cod personal (minim 4 caractere, nu se afiseaza): ")
            if cod != getpass.getpass("Repeta codul: "):
                print("Codurile nu coincid.")
                continue
            try:
                seteaza_editor(nume, cod)
                print(f"OK - {nume} poate edita.")
            except ValueError as e:
                print(e)
        elif opt == "2":
            nume = input("Numele de sters: ").strip()
            with db() as c:
                n = c.execute("DELETE FROM editori WHERE nume=?", (nume,)).rowcount
            print("Sters." if n else "Nu exista.")
        elif opt in ("0", ""):
            return


# =============================================================== actiuni
def curata(date_cerere):
    """Pastreaza doar campurile editabile prezente si le normalizeaza."""
    v = {}
    for k in EDITABILE:
        if k not in date_cerere:
            continue
        x = date_cerere.get(k)
        if k == "termen":
            v[k] = iso(x)
        elif k == "stadiu":
            v[k] = stadiu_valid(x)
        elif k == "prioritate":
            v[k] = prioritate_valida(x)
        else:
            v[k] = re.sub(r"\s+", " ", str(x or "")).strip()[:500] if k not in ("problema", "actiune") else str(x or "").strip()[:4000]
    return v


def rand_actiune(r, nr_com=0, nr_at=0):
    d = {k: (r[k] if r[k] is not None else "") for k, _ in CAMPURI}
    d.update(creat_de=r["creat_de"] or "", modificat_la=r["modificat_la"] or "",
             comentarii=nr_com, atasamente=nr_at)
    return d


def toate_actiunile():
    with db() as c:
        com = dict(c.execute("SELECT nr, COUNT(*) FROM comentarii GROUP BY nr").fetchall())
        rows = c.execute("SELECT * FROM actiuni ORDER BY nr DESC").fetchall()
    return [rand_actiune(r, com.get(r["nr"], 0), len(lista_atasamente(r["nr"]))) for r in rows]


def o_actiune(nr):
    with db() as c:
        r = c.execute("SELECT * FROM actiuni WHERE nr=?", (nr,)).fetchone()
        if not r:
            raise Eroare("Acțiunea nu există.", 404)
        com = [dict(x) for x in c.execute("SELECT autor, la, text FROM comentarii WHERE nr=? ORDER BY id", (nr,))]
        ist = [{**dict(x), "detalii": json.loads(x["detalii"] or "null")}
               for x in c.execute("SELECT operatie, autor, la, detalii FROM istoric WHERE nr=? ORDER BY id DESC", (nr,))]
    return {"actiune": rand_actiune(r, len(com), len(lista_atasamente(nr))),
            "comentarii": com, "istoric": ist, "atasamente": lista_atasamente(nr)}


def adauga(date_cerere, autor, c=None, op="adăugare"):
    v = curata({"stadiu": "Deschis", "prioritate": "Normală", **date_cerere})
    for obligatoriu in ("masina", "problema", "actiune", "responsabil"):
        if not v.get(obligatoriu):
            eticheta = dict(CAMPURI)[obligatoriu]
            raise Eroare(f"Completează câmpul „{eticheta}”.")
    v["data"] = iso(date_cerere.get("data")) or date.today().isoformat()
    if v["stadiu"] == "Închis":
        v["data_inchidere"] = iso(date_cerere.get("data_inchidere")) or date.today().isoformat()
    if date_cerere.get("nr"):
        v["nr"] = int(date_cerere["nr"])
    v.update(creat_de=autor, creat_la=acum(), modificat_de=autor, modificat_la=acum())

    def scrie(conn):
        cur = conn.execute(f"INSERT INTO actiuni ({','.join(v)}) VALUES ({','.join('?' * len(v))})", list(v.values()))
        nr = v.get("nr") or cur.lastrowid
        text = str(date_cerere.get("comentariu") or "").strip()
        if text:
            conn.execute("INSERT INTO comentarii(nr, autor, la, text) VALUES (?,?,?,?)", (nr, autor, acum(), text))
        jurnal(conn, nr, op, autor, {k: v[k] for k in v if k in dict(CAMPURI)})
        return nr

    if c is not None:
        return scrie(c)
    with db_lock:
        backup_zilnic()
        with db() as conn:
            return scrie(conn)


def modifica(nr, date_cerere, autor):
    with db_lock:
        backup_zilnic()
        with db() as c:
            r = c.execute("SELECT * FROM actiuni WHERE nr=?", (nr,)).fetchone()
            if not r:
                raise Eroare("Acțiunea nu există.", 404)
            v = curata(date_cerere)
            for obligatoriu in ("masina", "problema", "actiune", "responsabil"):
                if obligatoriu in v and not v[obligatoriu]:
                    raise Eroare(f"Câmpul „{dict(CAMPURI)[obligatoriu]}” nu poate fi gol.")
            if "stadiu" in v:
                if v["stadiu"] == "Închis" and r["stadiu"] != "Închis":
                    v["data_inchidere"] = date.today().isoformat()
                elif v["stadiu"] != "Închis":
                    v["data_inchidere"] = None
            schimbari = {k: {"din": r[k], "in": x} for k, x in v.items() if (r[k] or None) != (x or None)}
            if schimbari:
                v.update(modificat_de=autor, modificat_la=acum())
                c.execute(f"UPDATE actiuni SET {','.join(f'{k}=?' for k in v)} WHERE nr=?", [*v.values(), nr])
                jurnal(c, nr, "modificare", autor, schimbari)
            text = str(date_cerere.get("comentariu") or "").strip()
            if text:
                c.execute("INSERT INTO comentarii(nr, autor, la, text) VALUES (?,?,?,?)", (nr, autor, acum(), text))
                c.execute("UPDATE actiuni SET modificat_de=?, modificat_la=? WHERE nr=?", (autor, acum(), nr))


def comenteaza(nr, text, autor):
    text = str(text or "").strip()
    if not text:
        raise Eroare("Comentariul e gol.")
    with db_lock, db() as c:
        if not c.execute("SELECT 1 FROM actiuni WHERE nr=?", (nr,)).fetchone():
            raise Eroare("Acțiunea nu există.", 404)
        c.execute("INSERT INTO comentarii(nr, autor, la, text) VALUES (?,?,?,?)", (nr, autor, acum(), text[:4000]))
        c.execute("UPDATE actiuni SET modificat_de=?, modificat_la=? WHERE nr=?", (autor, acum(), nr))


def sterge(nr, autor):
    with db_lock:
        backup_zilnic()
        with db() as c:
            r = c.execute("SELECT * FROM actiuni WHERE nr=?", (nr,)).fetchone()
            if not r:
                raise Eroare("Acțiunea nu există.", 404)
            com = [dict(x) for x in c.execute("SELECT autor, la, text FROM comentarii WHERE nr=?", (nr,))]
            c.execute("DELETE FROM actiuni WHERE nr=?", (nr,))
            jurnal(c, nr, "ștergere", autor, {"actiune": dict(r), "comentarii": com})


def liste():
    rez = {}
    with db() as c:
        for k in LISTE:
            vals = [x for (x,) in c.execute("SELECT valoare FROM liste WHERE lista=? ORDER BY valoare", (k,))]
            for (x,) in c.execute(f"SELECT DISTINCT {k} FROM actiuni WHERE {k}<>'' ORDER BY {k}"):
                if x not in vals:
                    vals.append(x)
            rez[k] = vals
    rez["stadiu"], rez["prioritate"] = STADII, PRIORITATI
    return rez


def export_csv():
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow([e for _, e in CAMPURI] + ["Comentarii"])
    with db() as c:
        for r in c.execute("SELECT * FROM actiuni ORDER BY nr"):
            com = " | ".join(f"[{x['la'][:16]} {x['autor']}] {x['text']}"
                             for x in c.execute("SELECT * FROM comentarii WHERE nr=? ORDER BY id", (r["nr"],)))
            w.writerow([ro(r[k]) if k in DATE else (r[k] or "") for k, _ in CAMPURI] + [com])
    return ("﻿" + buf.getvalue()).encode("utf-8")


def istoric_general(limita=300):
    with db() as c:
        return [{**dict(x), "detalii": json.loads(x["detalii"] or "null")}
                for x in c.execute("SELECT * FROM istoric ORDER BY id DESC LIMIT ?", (limita,))]


# =============================================================== atasamente
def dir_atasamente(nr):
    return os.path.join(cale(ATASAMENTE_DIR), str(int(nr)))


def lista_atasamente(nr):
    d = dir_atasamente(nr)
    return sorted(f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f))) if os.path.isdir(d) else []


def salveaza_atasament(nr, nume, continut_b64):
    with db() as c:
        if not c.execute("SELECT 1 FROM actiuni WHERE nr=?", (nr,)).fetchone():
            raise Eroare("Acțiunea nu există.", 404)
    brut = base64.b64decode(continut_b64 or "", validate=False)
    if not brut:
        raise Eroare("Fișier gol.")
    if len(brut) > MAX_ATASAMENT_MB * 1024 * 1024:
        raise Eroare(f"Fișierul depășește {MAX_ATASAMENT_MB} MB.")
    d = dir_atasamente(nr)
    os.makedirs(d, exist_ok=True)
    nume = nume_sigur(os.path.basename(str(nume or "fisier")))
    baza, ext = os.path.splitext(nume)
    tinta, i = os.path.join(d, nume), 1
    while os.path.exists(tinta):
        tinta = os.path.join(d, f"{baza} ({i}){ext}")
        i += 1
    with open(tinta, "wb") as f:
        f.write(brut)
    return os.path.basename(tinta)


# =============================================================== import Excel / CSV
def citeste_xlsx(path):
    """Citeste .xlsx/.xlsm fara pachete externe -> {foaie: [[valori]]}."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        comune = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                comune.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        foi = {}
        for s in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", ns):
            t = rels[s.get(f"{{{ns['r']}}}id")].lstrip("/")
            t = t if t.startswith("xl/") else "xl/" + t
            randuri = []
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
                    elif tip == "b":
                        val = "Da" if v is not None and v.text == "1" else "Nu"
                    elif v is not None and v.text:
                        f = float(v.text)
                        val = int(f) if f.is_integer() else f
                    else:
                        val = ""
                    cel[col - 1] = val
                randuri.append([cel.get(i, "") for i in range(max(cel) + 1)] if cel else [])
            foi[s.get("name")] = randuri
        return foi


def harta_coloane(antet):
    norm = [fara_diacritice(h) for h in antet]
    harta, folosite = {}, set()
    for exact in (True, False):
        for camp, kws in CUVINTE_CHEIE.items():
            if camp in harta:
                continue
            for kw in kws:
                i = next((j for j, h in enumerate(norm) if h and j not in folosite
                          and (h == kw if exact else (len(kw) > 3 and kw in h))), None)
                if i is not None:
                    harta[camp] = i
                    folosite.add(i)
                    break
    return harta


def importa_randuri(antet, randuri, sursa):
    harta = harta_coloane(antet)
    if "actiune" not in harta and "problema" not in harta:
        raise SystemExit("Nu am găsit coloanele Acțiune / Problemă în primul rând al foii.")
    adaugate = sarite = 0
    with db_lock:
        backup_zilnic()
        with db() as c:
            for r in randuri:
                val = lambda k: (r[harta[k]] if k in harta and harta[k] < len(r) else "")
                d = {k: val(k) for k in CUVINTE_CHEIE}
                if not str(d["actiune"]).strip() and not str(d["problema"]).strip():
                    continue
                for k in ("masina", "problema", "actiune", "responsabil"):
                    if not str(d[k]).strip():
                        d[k] = "—"
                try:
                    d["nr"] = int(float(d["nr"])) if str(d["nr"]).strip() else None
                except ValueError:
                    d["nr"] = None
                if d["nr"] and c.execute("SELECT 1 FROM actiuni WHERE nr=?", (d["nr"],)).fetchone():
                    sarite += 1
                    continue
                d["comentariu"] = d.pop("comentarii")
                nr = adauga(d, f"import {sursa}", c=c, op=f"import {sursa}")
                adaugate += 1
                print(f"  + #{nr}: {d['masina']} - {str(d['actiune'])[:60]}")
    return adaugate, sarite


def importa_excel(path, foaie=None):
    foi = citeste_xlsx(path)
    if not foaie:
        foaie = next((n for n, rr in foi.items() if rr and {"actiune", "problema"} & set(harta_coloane(rr[0]))), None)
    if not foaie or foaie not in foi:
        raise SystemExit(f"Nu am găsit foaia cu acțiuni. Foi: {', '.join(foi)}")
    rr = foi[foaie]
    print(f"Foaia „{foaie}”: {max(len(rr) - 1, 0)} rânduri")
    a, s = importa_randuri(rr[0], rr[1:], "Excel")
    print(f"Gata: {a} acțiuni importate, {s} sărite (numărul există deja).")


def importa_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        prob = f.read(4096)
        f.seek(0)
        rr = list(csv.reader(f, csv.Sniffer().sniff(prob, delimiters=",;") if prob else csv.excel))
    if rr:
        a, s = importa_randuri(rr[0], rr[1:], "CSV")
        print(f"Gata: {a} acțiuni importate, {s} sărite.")


# =============================================================== mailuri
def citeste_mail_ini():
    p = cale(MAIL_INI)
    if not os.path.exists(p):
        raise SystemExit("Lipsește mail.ini. Copiază mail-exemplu.ini în mail.ini și completează-l.")
    cp = configparser.ConfigParser(interpolation=None)
    cp.optionxform = str
    cp.read(p, encoding="utf-8")
    return cp


def pregateste_mailuri(cp):
    g = cp["general"] if cp.has_section("general") else {}
    zile = int(g.get("zile_inainte_termen", "2") or 2)
    link = g.get("link_tracker", "").strip()
    subiect = g.get("subiect", "Grind Rings - acțiuni de urmărit").strip()
    cc = g.get("cc", "").strip()
    rezumat = g.get("rezumat_catre", "").strip()
    adrese = {fara_diacritice(k): v.strip() for k, v in (cp["destinatari"].items() if cp.has_section("destinatari") else [])}

    azi = date.today()
    pe_resp, toate = {}, []
    for a in toate_actiunile():
        if a["stadiu"] == "Închis":
            continue
        t = iso(a["termen"])
        zr = (date.fromisoformat(t) - azi).days if t else None
        urg = a["prioritate"] == "Urgentă"
        if not (urg or (zr is not None and zr <= zile)):
            continue
        a = dict(a, _z=zr)
        toate.append(a)
        pe_resp.setdefault(a["responsabil"] or "Nealocat", []).append(a)

    def tabel(items):
        rand = []
        for a in sorted(items, key=lambda x: x["_z"] if x["_z"] is not None else 9999):
            z = a["_z"]
            sit, cul = (("fără termen", "#667085") if z is None else
                        (f"întârziată {-z} zile", "#c0392b") if z < 0 else
                        ("termen azi", "#b7791f") if z == 0 else (f"în {z} zile", "#b7791f"))
            pr = " <b style='color:#c0392b'>URGENT</b>" if a["prioritate"] == "Urgentă" else ""
            rand.append(f"<tr><td>#{a['nr']}</td><td>{html.escape(a['masina'])}</td><td>{html.escape(a['reper'])}</td>"
                        f"<td>{html.escape(a['problema'])}</td><td>{html.escape(a['actiune'])}{pr}</td>"
                        f"<td>{html.escape(a['responsabil'])}</td><td>{ro(a['termen'])}</td>"
                        f"<td style='color:{cul}'>{sit}</td></tr>")
        return ("<table cellpadding='6' cellspacing='0' border='1' style='border-collapse:collapse;border-color:#d0d5dd;"
                "font:13px Segoe UI,Arial,sans-serif'><tr style='background:#f2f4f7'><th>Nr</th><th>Mașină</th><th>Reper</th>"
                "<th>Problemă</th><th>Acțiune</th><th>Responsabil</th><th>Termen</th><th>Situație</th></tr>"
                + "".join(rand) + "</table>")

    def corp(titlu, items):
        l = f"<p><a href='{html.escape(link)}'>Deschide {APP_NAME}</a></p>" if link else ""
        return (f"<div style='font:14px Segoe UI,Arial,sans-serif'><p>{titlu}</p>{tabel(items)}{l}"
                f"<p style='color:#888;font-size:12px'>Mesaj automat {APP_NAME} – {azi:%d.%m.%Y}</p></div>")

    mailuri = []
    for resp, items in sorted(pe_resp.items()):
        catre = adrese.get(fara_diacritice(resp))
        if not catre:
            print(f"  ! {resp}: nu are adresă în [destinatari] – apare doar în rezumat")
            continue
        mailuri.append({"catre": catre, "cc": cc, "subiect": f"{subiect} ({len(items)})",
                        "html": corp(f"Bună ziua, {html.escape(resp)}. Aveți {len(items)} acțiuni care cer atenție:", items)})
    if rezumat and toate:
        mailuri.append({"catre": rezumat, "cc": "", "subiect": f"{subiect} – rezumat ({len(toate)})",
                        "html": corp(f"Rezumat: {len(toate)} acțiuni întârziate, aproape de termen sau urgente.", toate)})
    return mailuri


def raport_mail():
    mailuri = pregateste_mailuri(citeste_mail_ini())
    d = cale(MAIL_OUT_DIR)
    os.makedirs(d, exist_ok=True)
    for f in os.listdir(d):
        if f.endswith(".html") or f == "mailuri.json":
            os.remove(os.path.join(d, f))
    index = []
    for i, m in enumerate(mailuri, 1):
        fn = f"mail_{i:02d}.html"
        with open(os.path.join(d, fn), "w", encoding="utf-8") as f:
            f.write(m["html"])
        index.append({"catre": m["catre"], "cc": m["cc"], "subiect": m["subiect"], "fisier": fn})
    with open(os.path.join(d, "mailuri.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    print(f"{len(mailuri)} mailuri pregătite în {d}")


def trimite_mail():
    cp = citeste_mail_ini()
    if not cp.has_section("smtp"):
        raise SystemExit("Lipsește secțiunea [smtp] din mail.ini (sau folosește outlook-mailuri.bat).")
    s = cp["smtp"]
    mailuri = pregateste_mailuri(cp)
    if not mailuri:
        return print("Nimic de trimis.")
    port = int(s.get("port", "25") or 25)
    sec = fara_diacritice(s.get("securitate", "starttls"))
    srv = smtplib.SMTP_SSL(s["server"], port, timeout=30) if sec == "ssl" else smtplib.SMTP(s["server"], port, timeout=30)
    with srv:
        if sec == "starttls":
            srv.starttls()
        if s.get("utilizator", "").strip():
            srv.login(s["utilizator"].strip(), s.get("parola", ""))
        for m in mailuri:
            msg = MIMEText(m["html"], "html", "utf-8")
            msg["Subject"], msg["From"], msg["To"] = m["subiect"], s["de_la"].strip(), m["catre"]
            if m["cc"]:
                msg["Cc"] = m["cc"]
            dest = [x.strip() for x in re.split(r"[;,]", f"{m['catre']};{m['cc']}") if x.strip()]
            srv.sendmail(msg["From"], dest, msg.as_string())
            print(f"  trimis: {m['catre']} – {m['subiect']}")
    print(f"{len(mailuri)} mailuri trimise.")


# =============================================================== HTTP
PAGINI = {"/": "index.html", "/index.html": "index.html", "/grafice": "grafice.html", "/grafice.html": "grafice.html"}


class Server(BaseHTTPRequestHandler):
    server_version = "GrindRings/1.0"

    def log_message(self, fmt, *args):
        sys.stdout.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {self.client_address[0]} {fmt % args}\n")

    def raspuns(self, status, corp, tip, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", tip)
        self.send_header("Content-Length", str(len(corp)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corp)

    def json(self, obj, status=200):
        self.raspuns(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def fisier(self, path, tip=None, descarcare=None):
        if not os.path.isfile(path):
            return self.json({"ok": False, "eroare": "Nu există."}, 404)
        with open(path, "rb") as f:
            date_f = f.read()
        extra = {"Content-Disposition": f"inline; filename*=UTF-8''{quote(descarcare)}"} if descarcare else None
        self.raspuns(200, date_f, tip or mimetypes.guess_type(path)[0] or "application/octet-stream", extra)

    def corp_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > (MAX_ATASAMENT_MB * 1.4 + 1) * 1024 * 1024:
            raise Eroare("Cerere prea mare.", 413)
        try:
            return json.loads(self.rfile.read(n).decode("utf-8") or "{}")
        except ValueError:
            raise Eroare("JSON invalid.")

    def ruta(self, metoda):
        url = urlparse(self.path)
        p = unquote(url.path).rstrip("/") or "/"
        q = parse_qs(url.query)
        try:
            if metoda == "GET":
                if p in PAGINI:
                    return self.fisier(cale(PAGINI[p]), "text/html; charset=utf-8")
                if p == "/api/actiuni":
                    return self.json({"ok": True, "actiuni": toate_actiunile(), "azi": date.today().isoformat()})
                m = re.fullmatch(r"/api/actiuni/(\d+)", p)
                if m:
                    return self.json({"ok": True, **o_actiune(int(m.group(1)))})
                if p == "/api/liste":
                    return self.json({"ok": True, "liste": liste()})
                if p == "/api/istoric":
                    return self.json({"ok": True, "istoric": istoric_general()})
                if p == "/api/export.csv":
                    return self.raspuns(200, export_csv(), "text/csv; charset=utf-8",
                                        {"Content-Disposition": f"attachment; filename=grind_rings_{date.today():%Y-%m-%d}.csv"})
                m = re.fullmatch(r"/atasamente/(\d+)/([^/]+)", p)
                if m:
                    nume = nume_sigur(m.group(2))
                    return self.fisier(os.path.join(dir_atasamente(int(m.group(1))), nume), descarcare=nume)
                if p == "/sanatate":
                    return self.json({"ok": True, "app": APP_NAME})
                return self.json({"ok": False, "eroare": "Pagina nu există."}, 404)

            d = self.corp_json()
            if p == "/api/cod":
                nume = editor_dupa_cod(d.get("cod"))
                return self.json({"ok": bool(nume), "nume": nume} if nume else {"ok": False, "eroare": "Cod incorect."},
                                 200 if nume else 403)
            autor = cere_editor(d)
            if p == "/api/actiuni":
                nr = adauga(d, autor)
                print(f"  + #{nr} de {autor}")
                return self.json({"ok": True, "nr": nr})
            m = re.fullmatch(r"/api/actiuni/(\d+)(?:/(comentariu|sterge|atasament))?", p)
            if m:
                nr, act = int(m.group(1)), m.group(2)
                if act is None:
                    modifica(nr, d, autor)
                elif act == "comentariu":
                    comenteaza(nr, d.get("text"), autor)
                elif act == "sterge":
                    sterge(nr, autor)
                else:
                    nume = salveaza_atasament(nr, d.get("nume"), d.get("continut"))
                    with db_lock, db() as c:
                        jurnal(c, nr, "atașament", autor, {"fisier": nume})
                print(f"  ~ #{nr} {act or 'modificare'} de {autor}")
                return self.json({"ok": True})
            return self.json({"ok": False, "eroare": "Rută necunoscută."}, 404)
        except Eroare as e:
            return self.json({"ok": False, "eroare": str(e)}, e.status)
        except sqlite3.IntegrityError as e:
            return self.json({"ok": False, "eroare": f"Date invalide: {e}"}, 400)
        except Exception as e:  # noqa: BLE001
            return self.json({"ok": False, "eroare": f"Eroare server: {type(e).__name__}: {e}"}, 500)

    def do_GET(self):
        self.ruta("GET")

    def do_POST(self):
        self.ruta("POST")


def porneste():
    if sys.stdout is None or sys.stderr is None:      # pornit in fundal (pythonw)
        sys.stdout = sys.stderr = open(cale("server.log"), "a", encoding="utf-8", buffering=1)
    os.makedirs(cale(ATASAMENTE_DIR), exist_ok=True)
    srv = ThreadingHTTPServer((HOST, PORT), Server)
    print(f"{APP_NAME} rulează pe http://{HOST}:{PORT}/   (baza de date: {DB_PATH})")
    if not lista_editori():
        print("  ATENȚIE: nu există niciun editor – rulează editori.bat.")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("Oprit.")


def main():
    p = argparse.ArgumentParser(description=APP_NAME)
    p.add_argument("--editori", action="store_true", help="administrare editori")
    p.add_argument("--adauga-editor", nargs=2, metavar=("NUME", "COD"), help="adaugă/schimbă un editor fără meniu")
    p.add_argument("--numar-editori", action="store_true", help="afișează câți editori există")
    p.add_argument("--import-excel", metavar="FISIER")
    p.add_argument("--foaie", metavar="NUME", help="foaia din Excel (implicit se detectează)")
    p.add_argument("--import-csv", metavar="FISIER")
    p.add_argument("--raport-mail", action="store_true", help="pregătește mailurile pentru Outlook")
    p.add_argument("--trimite-mail", action="store_true", help="trimite mailurile prin SMTP")
    a = p.parse_args()
    creeaza_db()
    if a.editori:
        return meniu_editori()
    if a.adauga_editor:
        try:
            seteaza_editor(*a.adauga_editor)
        except ValueError as e:
            raise SystemExit(str(e))
        return print(f"OK - {a.adauga_editor[0]} poate edita.")
    if a.numar_editori:
        return print(len(lista_editori()))
    if a.import_excel:
        return importa_excel(a.import_excel, a.foaie)
    if a.import_csv:
        return importa_csv(a.import_csv)
    if a.raport_mail:
        return raport_mail()
    if a.trimite_mail:
        return trimite_mail()
    porneste()


if __name__ == "__main__":
    main()
