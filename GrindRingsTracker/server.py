#!/usr/bin/env python3
"""
Grind Rings Action Tracker - server (port 8092).

Server independent pentru actiunile de la rectificarea inelelor (grind rings).
Date intr-o baza SQLite proprie (grind_rings.db). Doar Python 3, fara pachete externe.

  python server.py                       porneste serverul
  python server.py --utilizatori         administrare utilizatori (utilizator + parola)
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
CREATE TABLE IF NOT EXISTS utilizatori (
  utilizator TEXT PRIMARY KEY COLLATE NOCASE, nume TEXT, sare TEXT, hash TEXT,
  admin INTEGER DEFAULT 0, activ INTEGER DEFAULT 1, creat_la TEXT, ultima_logare TEXT
);
CREATE TABLE IF NOT EXISTS sesiuni (
  token_hash TEXT PRIMARY KEY, utilizator TEXT, creata_la TEXT, expira TEXT
);
CREATE TABLE IF NOT EXISTS liste (
  lista TEXT, valoare TEXT, PRIMARY KEY (lista, valoare)
);
CREATE TABLE IF NOT EXISTS setari (cheie TEXT PRIMARY KEY, valoare TEXT);
CREATE TABLE IF NOT EXISTS destinatari (responsabil TEXT PRIMARY KEY COLLATE NOCASE, email TEXT);
CREATE TABLE IF NOT EXISTS mail_jurnal (
  id INTEGER PRIMARY KEY AUTOINCREMENT, la TEXT, tip TEXT, catre TEXT, subiect TEXT, ok INTEGER, eroare TEXT
);
CREATE INDEX IF NOT EXISTS ix_com_nr ON comentarii(nr);
CREATE INDEX IF NOT EXISTS ix_ist_nr ON istoric(nr);
""")
        coloane = [r[1] for r in c.execute("PRAGMA table_info(utilizatori)")]
        if "email" not in coloane:
            c.execute("ALTER TABLE utilizatori ADD COLUMN email TEXT DEFAULT ''")
    importa_mail_ini()


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


# =============================================================== utilizatori si sesiuni
SESIUNE_COOKIE = "gr_sesiune"
SESIUNE_ZILE_TINE_MINTE = 180      # "Tine-ma minte": autentificare automata atatea zile
SESIUNE_ORE_FARA = 12              # fara "Tine-ma minte": pana inchizi browserul (max 12 ore)


def _hash(parola, sare):
    return hashlib.pbkdf2_hmac("sha256", str(parola).encode(), bytes.fromhex(sare), 200_000).hex()


def _valideaza(utilizator, parola=None):
    utilizator = str(utilizator or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9._-]{2,40}", utilizator):
        raise Eroare("Utilizatorul trebuie să aibă 2–40 caractere: litere, cifre, punct, minus sau _.")
    if parola is not None and len(str(parola)) < 6:
        raise Eroare("Parola trebuie să aibă cel puțin 6 caractere.")
    return utilizator


def numar_utilizatori():
    with db() as c:
        return c.execute("SELECT COUNT(*) FROM utilizatori").fetchone()[0]


def lista_utilizatori():
    with db() as c:
        return [dict(r) for r in c.execute(
            "SELECT utilizator, nume, email, admin, activ, creat_la, ultima_logare FROM utilizatori ORDER BY utilizator")]


def creeaza_utilizator(utilizator, nume, parola, admin=False, email=""):
    utilizator = _valideaza(utilizator, parola)
    nume = str(nume or "").strip() or utilizator
    email = _adrese_valide(email, "Email")
    with db() as c:
        if c.execute("SELECT 1 FROM utilizatori WHERE utilizator=?", (utilizator,)).fetchone():
            raise Eroare(f"Utilizatorul „{utilizator}” există deja.")
        sare = secrets.token_hex(16)
        c.execute("INSERT INTO utilizatori(utilizator, nume, email, sare, hash, admin, activ, creat_la) VALUES (?,?,?,?,?,?,1,?)",
                  (utilizator, nume, email, sare, _hash(parola, sare), 1 if admin else 0, acum()))
    return utilizator


def schimba_parola(utilizator, parola):
    _valideaza(utilizator, parola)
    sare = secrets.token_hex(16)
    with db() as c:
        n = c.execute("UPDATE utilizatori SET sare=?, hash=? WHERE utilizator=?", (sare, _hash(parola, sare), utilizator)).rowcount
        c.execute("DELETE FROM sesiuni WHERE utilizator=?", (utilizator,))     # delogheaza peste tot
    if not n:
        raise Eroare("Utilizatorul nu există.", 404)


def modifica_utilizator(utilizator, nume=None, admin=None, activ=None, email=None):
    with db() as c:
        r = c.execute("SELECT * FROM utilizatori WHERE utilizator=?", (utilizator,)).fetchone()
        if not r:
            raise Eroare("Utilizatorul nu există.", 404)
        if (admin is False or activ is False) and r["admin"]:
            admini = c.execute("SELECT COUNT(*) FROM utilizatori WHERE admin=1 AND activ=1").fetchone()[0]
            if admini <= 1:
                raise Eroare("Trebuie să rămână cel puțin un administrator activ.")
        if nume is not None:
            c.execute("UPDATE utilizatori SET nume=? WHERE utilizator=?", (str(nume).strip() or utilizator, utilizator))
        if email is not None:
            c.execute("UPDATE utilizatori SET email=? WHERE utilizator=?", (_adrese_valide(email, "Email"), utilizator))
        if admin is not None:
            c.execute("UPDATE utilizatori SET admin=? WHERE utilizator=?", (1 if admin else 0, utilizator))
        if activ is not None:
            c.execute("UPDATE utilizatori SET activ=? WHERE utilizator=?", (1 if activ else 0, utilizator))
            if not activ:
                c.execute("DELETE FROM sesiuni WHERE utilizator=?", (utilizator,))


def sterge_utilizator(utilizator):
    modifica_utilizator(utilizator, activ=False)       # verifica sa ramana un admin
    with db() as c:
        c.execute("DELETE FROM utilizatori WHERE utilizator=?", (utilizator,))
        c.execute("DELETE FROM sesiuni WHERE utilizator=?", (utilizator,))


def verifica_parola(utilizator, parola):
    with db() as c:
        r = c.execute("SELECT * FROM utilizatori WHERE utilizator=?", (str(utilizator or "").strip(),)).fetchone()
    # calculam hash-ul oricum, ca timpul de raspuns sa nu tradeze daca utilizatorul exista
    sare = r["sare"] if r else "00" * 16
    ok = secrets.compare_digest(_hash(parola or "", sare), r["hash"] if r else "x")
    return dict(r) if (r and ok and r["activ"]) else None


def creeaza_sesiune(utilizator, tine_minte):
    token = secrets.token_urlsafe(32)
    durata = timedelta(days=SESIUNE_ZILE_TINE_MINTE) if tine_minte else timedelta(hours=SESIUNE_ORE_FARA)
    with db() as c:
        c.execute("DELETE FROM sesiuni WHERE expira < ?", (acum(),))
        c.execute("INSERT INTO sesiuni(token_hash, utilizator, creata_la, expira) VALUES (?,?,?,?)",
                  (hashlib.sha256(token.encode()).hexdigest(), utilizator, acum(),
                   (datetime.now() + durata).strftime("%Y-%m-%d %H:%M:%S")))
        c.execute("UPDATE utilizatori SET ultima_logare=? WHERE utilizator=?", (acum(), utilizator))
    return token, durata


def utilizator_din_token(token):
    if not token:
        return None
    with db() as c:
        r = c.execute("""SELECT u.* FROM sesiuni s JOIN utilizatori u ON u.utilizator = s.utilizator
                         WHERE s.token_hash=? AND s.expira > ? AND u.activ=1""",
                      (hashlib.sha256(token.encode()).hexdigest(), acum())).fetchone()
    return dict(r) if r else None


def sterge_sesiune(token):
    if token:
        with db() as c:
            c.execute("DELETE FROM sesiuni WHERE token_hash=?", (hashlib.sha256(token.encode()).hexdigest(),))


def public(u):
    return {"utilizator": u["utilizator"], "nume": u["nume"], "email": u.get("email") or "", "admin": bool(u["admin"])} if u else None


def meniu_utilizatori():
    """Administrare din consola (editori.bat) - util si daca ai uitat parola de administrator."""
    def intreaba_parola():
        while True:
            p1 = getpass.getpass("Parola (minim 6 caractere, nu se afiseaza): ")
            if p1 != getpass.getpass("Repeta parola: "):
                print("Parolele nu coincid.")
            elif len(p1) < 6:
                print("Prea scurta.")
            else:
                return p1
    while True:
        us = lista_utilizatori()
        print(f"\n===== {APP_NAME} - UTILIZATORI =====")
        for u in us:
            print(f"  {u['utilizator']:<20} {u['nume']:<25} {'ADMIN' if u['admin'] else '':<6} "
                  f"{'' if u['activ'] else '(dezactivat)'}  ultima logare: {(u['ultima_logare'] or '-')[:16]}")
        if not us:
            print("  (niciun utilizator - creeaza unul, sau deschide tracker-ul in browser)")
        print("\n  1 = utilizator nou   2 = schimba parola   3 = fa / scoate administrator")
        print("  4 = dezactiveaza / activeaza   5 = sterge   0 = iesire")
        opt = input("> ").strip()
        try:
            if opt == "1":
                ut = input("Utilizator (pentru logare, ex. ion.popescu): ").strip()
                nume = input("Nume afisat (ex. Ion Popescu): ").strip()
                adm = not us or input("Administrator? D/N: ").strip().lower().startswith(("d", "y"))
                creeaza_utilizator(ut, nume, intreaba_parola(), adm)
                print(f"OK - {ut} creat{' (administrator)' if adm else ''}.")
            elif opt == "2":
                ut = input("Utilizator: ").strip()
                schimba_parola(ut, intreaba_parola())
                print("Parola schimbata.")
            elif opt == "3":
                ut = input("Utilizator: ").strip()
                u = next((x for x in us if x["utilizator"].lower() == ut.lower()), None)
                if u:
                    modifica_utilizator(u["utilizator"], admin=not u["admin"])
                    print("Acum este administrator." if not u["admin"] else "Nu mai este administrator.")
                else:
                    print("Nu exista.")
            elif opt == "4":
                ut = input("Utilizator: ").strip()
                u = next((x for x in us if x["utilizator"].lower() == ut.lower()), None)
                if u:
                    modifica_utilizator(u["utilizator"], activ=not u["activ"])
                    print("Activat." if not u["activ"] else "Dezactivat.")
                else:
                    print("Nu exista.")
            elif opt == "5":
                ut = input("Utilizator de sters: ").strip()
                if input(f"Sigur stergi {ut}? D/N: ").strip().lower().startswith(("d", "y")):
                    sterge_utilizator(ut)
                    print("Sters.")
            elif opt in ("0", ""):
                return
        except Eroare as e:
            print(e)


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


# =============================================================== setari email
SETARI_EMAIL_IMPLICIT = {
    "smtp": {"server": "", "port": 587, "securitate": "starttls", "utilizator": "", "parola": "", "de_la": ""},
    "general": {"zile_inainte_termen": 2, "subiect": "Grind Rings - acțiuni de urmărit", "link_tracker": "",
                "cc": "", "rezumat_catre": "", "doar_rezumat": False},
    "program": {"activ": False, "zile": [1, 2, 3, 4, 5], "ora": "07:30"},
}


def citeste_setare(cheie, implicit=None):
    with db() as c:
        r = c.execute("SELECT valoare FROM setari WHERE cheie=?", (cheie,)).fetchone()
    return json.loads(r["valoare"]) if r else implicit


def scrie_setare(cheie, valoare):
    with db() as c:
        c.execute("INSERT OR REPLACE INTO setari(cheie, valoare) VALUES (?,?)", (cheie, json.dumps(valoare, ensure_ascii=False)))


def setari_email():
    """Setarile complete (cu parola SMTP) - doar pentru uz intern."""
    s = json.loads(json.dumps(SETARI_EMAIL_IMPLICIT))
    salvat = citeste_setare("email", {}) or {}
    for k in s:
        s[k].update(salvat.get(k, {}))
    return s


def importa_mail_ini():
    """O singura data: preia setarile din mail.ini (versiunea veche) in baza de date."""
    p = cale(MAIL_INI)
    if not os.path.exists(p) or citeste_setare("email") is not None:
        return
    cp = configparser.ConfigParser(interpolation=None)
    cp.optionxform = str
    cp.read(p, encoding="utf-8")
    s = setari_email()
    for sec in ("smtp", "general"):
        if cp.has_section(sec):
            for k, v in cp[sec].items():
                if k in s[sec]:
                    s[sec][k] = int(v) if k in ("port", "zile_inainte_termen") and v.strip().isdigit() else v.strip()
    scrie_setare("email", s)
    if cp.has_section("destinatari"):
        with db() as c:
            for resp, email in cp["destinatari"].items():
                if email.strip():
                    c.execute("INSERT OR REPLACE INTO destinatari(responsabil, email) VALUES (?,?)", (resp.strip(), email.strip()))
    print(f"  Setările de email au fost preluate din {MAIL_INI} (de acum se modifică din pagina Setări).")


def _adrese_valide(text, camp):
    adrese = [x.strip() for x in re.split(r"[;,]", str(text or "")) if x.strip()]
    for a in adrese:
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", a):
            raise Eroare(f"Adresă de email invalidă la „{camp}”: {a}")
    return "; ".join(adrese)


def salveaza_setari_email(d):
    s = setari_email()
    sm, g, pr = d.get("smtp", {}), d.get("general", {}), d.get("program", {})
    s["smtp"].update(server=str(sm.get("server", "")).strip(), utilizator=str(sm.get("utilizator", "")).strip(),
                     de_la=_adrese_valide(sm.get("de_la", ""), "Expeditor"),
                     securitate=sm.get("securitate") if sm.get("securitate") in ("starttls", "ssl", "niciuna") else "starttls")
    try:
        s["smtp"]["port"] = max(1, min(65535, int(sm.get("port") or 25)))
    except (TypeError, ValueError):
        raise Eroare("Portul trebuie să fie un număr.")
    if sm.get("sterge_parola"):
        s["smtp"]["parola"] = ""
    elif sm.get("parola"):
        s["smtp"]["parola"] = str(sm["parola"])
    try:
        s["general"]["zile_inainte_termen"] = max(0, min(60, int(g.get("zile_inainte_termen", 2))))
    except (TypeError, ValueError):
        raise Eroare("„Zile înainte de termen” trebuie să fie un număr.")
    s["general"].update(subiect=str(g.get("subiect", "")).strip() or SETARI_EMAIL_IMPLICIT["general"]["subiect"],
                        link_tracker=str(g.get("link_tracker", "")).strip(),
                        cc=_adrese_valide(g.get("cc"), "Copie (CC)"),
                        rezumat_catre=_adrese_valide(g.get("rezumat_catre"), "Rezumat către"),
                        doar_rezumat=bool(g.get("doar_rezumat")))
    ora = str(pr.get("ora", "07:30")).strip()
    if not re.fullmatch(r"([01]\d|2[0-3]):[0-5]\d", ora):
        raise Eroare("Ora trebuie să fie în forma HH:MM (ex. 07:30).")
    zile = sorted({int(z) for z in pr.get("zile", []) if str(z).isdigit() and 1 <= int(z) <= 7})
    s["program"].update(activ=bool(pr.get("activ")), ora=ora, zile=zile)
    if s["program"]["activ"] and not s["smtp"]["server"]:
        raise Eroare("Pentru trimiterea automată completează serverul SMTP.")
    if s["program"]["activ"] and not zile:
        raise Eroare("Alege cel puțin o zi pentru trimiterea automată.")
    scrie_setare("email", s)
    if "destinatari" in d:
        with db() as c:
            c.execute("DELETE FROM destinatari")
            for x in d["destinatari"]:
                resp, email = str(x.get("responsabil", "")).strip(), str(x.get("email", "")).strip()
                if resp and email:
                    c.execute("INSERT OR REPLACE INTO destinatari(responsabil, email) VALUES (?,?)",
                              (resp, _adrese_valide(email, resp)))


def destinatari():
    """Toti responsabilii (din actiuni) + adresele lor. Daca nu e trecuta o adresa, se foloseste
    emailul contului cu acelasi nume sau utilizator."""
    with db() as c:
        manual = {fara_diacritice(r["responsabil"]): (r["responsabil"], r["email"]) for r in c.execute("SELECT * FROM destinatari")}
        conturi = {}
        for u in c.execute("SELECT utilizator, nume, email FROM utilizatori WHERE activ=1 AND email<>''"):
            conturi[fara_diacritice(u["nume"])] = u["email"]
            conturi[fara_diacritice(u["utilizator"])] = u["email"]
        resp = [r[0] for r in c.execute("SELECT DISTINCT responsabil FROM actiuni WHERE responsabil<>'' ORDER BY responsabil")]
    rez, vazut = [], set()
    for nume in resp + [v[0] for v in manual.values()]:
        k = fara_diacritice(nume)
        if k in vazut:
            continue
        vazut.add(k)
        if k in manual:
            rez.append({"responsabil": nume, "email": manual[k][1], "sursa": "manual"})
        elif k in conturi:
            rez.append({"responsabil": nume, "email": conturi[k], "sursa": "cont"})
        else:
            rez.append({"responsabil": nume, "email": "", "sursa": ""})
    return sorted(rez, key=lambda x: fara_diacritice(x["responsabil"]))


def setari_email_publice():
    s = setari_email()
    s["smtp"]["parola_setata"] = bool(s["smtp"].pop("parola"))
    return {**s, "destinatari": destinatari(), "jurnal": jurnal_mail(), "ultima_auto": citeste_setare("ultima_trimitere_auto")}


def jurnal_mail(limita=40):
    with db() as c:
        return [dict(r) for r in c.execute("SELECT * FROM mail_jurnal ORDER BY id DESC LIMIT ?", (limita,))]


def noteaza_mail(tip, catre, subiect, ok, eroare=""):
    with db() as c:
        c.execute("INSERT INTO mail_jurnal(la, tip, catre, subiect, ok, eroare) VALUES (?,?,?,?,?,?)",
                  (acum(), tip, catre, subiect, 1 if ok else 0, str(eroare)[:500]))
        c.execute("DELETE FROM mail_jurnal WHERE id NOT IN (SELECT id FROM mail_jurnal ORDER BY id DESC LIMIT 1000)")


def pregateste_mailuri(s=None):
    s = s or setari_email()
    g = s["general"]
    zile = int(g.get("zile_inainte_termen", 2))
    link, subiect, cc, rezumat = g["link_tracker"], g["subiect"], g["cc"], g["rezumat_catre"]
    adrese = {fara_diacritice(d["responsabil"]): d["email"] for d in destinatari() if d["email"]}

    azi = date.today()
    pe_resp, toate = {}, []
    for a in toate_actiunile():
        if a["stadiu"] == "Închis":
            continue
        t = iso(a["termen"])
        zr = (date.fromisoformat(t) - azi).days if t else None
        if not (a["prioritate"] == "Urgentă" or (zr is not None and zr <= zile)):
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

    mailuri, fara_adresa = [], []
    if not g.get("doar_rezumat"):
        for resp, items in sorted(pe_resp.items()):
            catre = adrese.get(fara_diacritice(resp))
            if not catre:
                fara_adresa.append(resp)
                continue
            mailuri.append({"catre": catre, "cc": cc, "subiect": f"{subiect} ({len(items)})", "pentru": resp,
                            "html": corp(f"Bună ziua, {html.escape(resp)}. Aveți {len(items)} acțiuni care cer atenție:", items)})
    if rezumat and toate:
        mailuri.append({"catre": rezumat, "cc": "", "subiect": f"{subiect} – rezumat ({len(toate)})", "pentru": "rezumat",
                        "html": corp(f"Rezumat: {len(toate)} acțiuni întârziate, aproape de termen sau urgente.", toate)})
    return mailuri, fara_adresa


def _conexiune_smtp(sm):
    if not sm["server"]:
        raise Eroare("Serverul SMTP nu este completat.")
    if not sm["de_la"]:
        raise Eroare("Adresa expeditorului nu este completată.")
    port = int(sm["port"] or 25)
    srv = (smtplib.SMTP_SSL(sm["server"], port, timeout=30) if sm["securitate"] == "ssl"
           else smtplib.SMTP(sm["server"], port, timeout=30))
    try:
        if sm["securitate"] == "starttls":
            srv.starttls()
        if sm["utilizator"]:
            srv.login(sm["utilizator"], sm["parola"])
    except Exception:
        srv.close()
        raise
    return srv


def _trimite(srv, de_la, m):
    msg = MIMEText(m["html"], "html", "utf-8")
    msg["Subject"], msg["From"], msg["To"] = m["subiect"], de_la.split(";")[0].strip(), m["catre"]
    if m.get("cc"):
        msg["Cc"] = m["cc"]
    dest = [x.strip() for x in re.split(r"[;,]", f"{m['catre']};{m.get('cc', '')}") if x.strip()]
    srv.sendmail(msg["From"], dest, msg.as_string())


def trimite_mailuri(tip="manual"):
    """Trimite prin SMTP. Intoarce {trimise, erori, fara_adresa}."""
    s = setari_email()
    mailuri, fara = pregateste_mailuri(s)
    rez = {"trimise": 0, "erori": [], "fara_adresa": fara}
    if not mailuri:
        noteaza_mail(tip, "-", "Nimic de trimis (nicio acțiune întârziată, apropiată sau urgentă)", True)
        return rez
    try:
        srv = _conexiune_smtp(s["smtp"])
    except Exception as e:  # noqa: BLE001
        noteaza_mail(tip, "-", "Conectare la serverul de mail", False, e)
        raise Eroare(f"Nu mă pot conecta la serverul de mail: {e}")
    with srv:
        for m in mailuri:
            try:
                _trimite(srv, s["smtp"]["de_la"], m)
                rez["trimise"] += 1
                noteaza_mail(tip, m["catre"], m["subiect"], True)
            except Exception as e:  # noqa: BLE001
                rez["erori"].append(f"{m['catre']}: {e}")
                noteaza_mail(tip, m["catre"], m["subiect"], False, e)
    return rez


def trimite_test(catre):
    catre = _adrese_valide(catre, "Destinatar test")
    if not catre:
        raise Eroare("Scrie adresa la care să trimit mailul de test.")
    s = setari_email()
    m = {"catre": catre, "cc": "", "subiect": f"{APP_NAME} – mail de test",
         "html": f"<div style='font:14px Segoe UI,Arial,sans-serif'><p>Acesta este un mail de test trimis din {APP_NAME}.</p>"
                 f"<p>Dacă l-ai primit, setările de email sunt corecte.</p><p style='color:#888;font-size:12px'>{acum()}</p></div>"}
    try:
        with _conexiune_smtp(s["smtp"]) as srv:
            _trimite(srv, s["smtp"]["de_la"], m)
    except Eroare:
        raise
    except Exception as e:  # noqa: BLE001
        noteaza_mail("test", catre, m["subiect"], False, e)
        raise Eroare(f"Trimiterea a eșuat: {e}")
    noteaza_mail("test", catre, m["subiect"], True)


def planificator_mail():
    """Fir de executie: trimite automat mailurile in zilele si la ora din Setari."""
    import time
    while True:
        try:
            pr = setari_email()["program"]
            acum_ = datetime.now()
            if (pr.get("activ") and acum_.isoweekday() in pr.get("zile", [])
                    and acum_.strftime("%H:%M") >= pr.get("ora", "07:30")
                    and citeste_setare("ultima_trimitere_auto", {}).get("data") != acum_.date().isoformat()):
                scrie_setare("ultima_trimitere_auto", {"data": acum_.date().isoformat(), "la": acum(), "rezultat": "în curs"})
                try:
                    r = trimite_mailuri("automat")
                    txt = f"{r['trimise']} trimise" + (f", {len(r['erori'])} erori" if r["erori"] else "")
                except Exception as e:  # noqa: BLE001
                    txt = f"eroare: {e}"
                scrie_setare("ultima_trimitere_auto", {"data": acum_.date().isoformat(), "la": acum(), "rezultat": txt})
                print(f"  @ mail automat: {txt}")
        except Exception as e:  # noqa: BLE001
            print(f"  ! planificator mail: {e}")
        time.sleep(30)


def raport_mail():
    """Pentru outlook-mailuri.bat: scrie mailurile in mail_out/."""
    mailuri, fara = pregateste_mailuri()
    for r in fara:
        print(f"  ! {r}: nu are adresă de email (Setări → Email → Destinatari)")
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
        noteaza_mail("outlook", m["catre"], m["subiect"], True, "pregătit pentru Outlook")
    with open(os.path.join(d, "mailuri.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    print(f"{len(mailuri)} mailuri pregătite în {d}")


def trimite_mail():
    """Pentru mail.bat: trimite acum prin SMTP."""
    try:
        r = trimite_mailuri("manual (mail.bat)")
    except Eroare as e:
        raise SystemExit(str(e))
    for x in r["fara_adresa"]:
        print(f"  ! {x}: nu are adresă de email")
    for x in r["erori"]:
        print(f"  ! {x}")
    print(f"{r['trimise']} mailuri trimise.")


# =============================================================== HTTP
PAGINI = {"/": "index.html", "/index.html": "index.html", "/grafice": "grafice.html", "/grafice.html": "grafice.html",
          "/setari": "setari.html", "/setari.html": "setari.html"}


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

    def json(self, obj, status=200, extra=None):
        self.raspuns(status, json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", extra)

    def fisier(self, path, tip=None, descarcare=None):
        if not os.path.isfile(path):
            return self.json({"ok": False, "eroare": "Nu există."}, 404)
        with open(path, "rb") as f:
            date_f = f.read()
        extra = {"Content-Disposition": f"inline; filename*=UTF-8''{quote(descarcare)}"} if descarcare else None
        self.raspuns(200, date_f, tip or mimetypes.guess_type(path)[0] or "application/octet-stream", extra)

    def token(self):
        for bucata in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = bucata.strip().partition("=")
            if k == SESIUNE_COOKIE:
                return v
        return None

    def utilizator(self):
        return utilizator_din_token(self.token())

    def cookie(self, token, durata=None):
        c = f"{SESIUNE_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax"
        if durata is not None:
            c += f"; Max-Age={int(durata.total_seconds())}"
        return c

    def verifica_origine(self):
        """Blocheaza cererile de scriere trimise de alte site-uri (CSRF)."""
        orig = self.headers.get("Origin")
        if orig and urlparse(orig).netloc.lower() != (self.headers.get("Host") or "").lower():
            raise Eroare("Cerere respinsă (altă origine).", 403)
        if "application/json" not in (self.headers.get("Content-Type") or ""):
            raise Eroare("Cerere invalidă.", 400)

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
                if p == "/api/eu":
                    return self.json({"ok": True, "eu": public(self.utilizator()), "configurat": numar_utilizatori() > 0})
                if p in ("/api/utilizatori", "/api/setari/email", "/api/setari/email/previzualizare"):
                    u = self.utilizator()
                    if not u or not u["admin"]:
                        raise Eroare("Doar pentru administratori.", 403)
                    if p == "/api/utilizatori":
                        return self.json({"ok": True, "utilizatori": lista_utilizatori()})
                    if p == "/api/setari/email":
                        return self.json({"ok": True, "email": setari_email_publice()})
                    mailuri, fara = pregateste_mailuri()
                    return self.json({"ok": True, "mailuri": mailuri, "fara_adresa": fara})
                if p == "/sanatate":
                    return self.json({"ok": True, "app": APP_NAME})
                return self.json({"ok": False, "eroare": "Pagina nu există."}, 404)

            self.verifica_origine()
            d = self.corp_json()

            # ----- autentificare
            if p == "/api/initializare":          # primul cont, doar cat timp nu exista niciunul
                with db_lock:
                    if numar_utilizatori() > 0:
                        raise Eroare("Există deja utilizatori. Autentifică-te.", 403)
                    ut = creeaza_utilizator(d.get("utilizator"), d.get("nume"), d.get("parola"), admin=True)
                token, durata = creeaza_sesiune(ut, True)
                print(f"  * primul administrator creat: {ut}")
                return self.json({"ok": True, "eu": public(utilizator_din_token(token))}, extra={"Set-Cookie": self.cookie(token, durata)})
            if p == "/api/login":
                u = verifica_parola(d.get("utilizator"), d.get("parola"))
                if not u:
                    print(f"  ! logare esuata: {str(d.get('utilizator'))[:40]} de la {self.client_address[0]}")
                    raise Eroare("Utilizator sau parolă greșită.", 401)
                tine = bool(d.get("tine_minte"))
                token, durata = creeaza_sesiune(u["utilizator"], tine)
                return self.json({"ok": True, "eu": public(u)},
                                 extra={"Set-Cookie": self.cookie(token, durata if tine else None)})
            if p == "/api/logout":
                sterge_sesiune(self.token())
                return self.json({"ok": True}, extra={"Set-Cookie": self.cookie("", timedelta(0))})

            u = self.utilizator()
            if not u:
                raise Eroare("Trebuie să te autentifici.", 401)
            autor = u["nume"] or u["utilizator"]

            if p == "/api/parola":                # schimbarea propriei parole
                if not verifica_parola(u["utilizator"], d.get("parola_veche")):
                    raise Eroare("Parola actuală este greșită.", 403)
                schimba_parola(u["utilizator"], d.get("parola_noua"))
                token, durata = creeaza_sesiune(u["utilizator"], True)
                return self.json({"ok": True}, extra={"Set-Cookie": self.cookie(token, durata)})
            m = re.fullmatch(r"/api/utilizatori(?:/([^/]+)(?:/(parola|sterge))?)?", p)
            if m:
                if not u["admin"]:
                    raise Eroare("Doar pentru administratori.", 403)
                ut, act = m.group(1), m.group(2)
                if ut is None:
                    creeaza_utilizator(d.get("utilizator"), d.get("nume"), d.get("parola"), bool(d.get("admin")), d.get("email", ""))
                elif act == "parola":
                    schimba_parola(ut, d.get("parola"))
                elif act == "sterge":
                    if ut.lower() == u["utilizator"].lower():
                        raise Eroare("Nu îți poți șterge propriul cont.")
                    sterge_utilizator(ut)
                else:
                    modifica_utilizator(ut, d.get("nume"), d.get("admin"), d.get("activ"), d.get("email"))
                return self.json({"ok": True, "utilizatori": lista_utilizatori()})
            if p.startswith("/api/setari/"):
                if not u["admin"]:
                    raise Eroare("Doar pentru administratori.", 403)
                if p == "/api/setari/email":
                    salveaza_setari_email(d)
                    print(f"  ~ setari email salvate de {autor}")
                    return self.json({"ok": True, "email": setari_email_publice()})
                if p == "/api/setari/email/test":
                    trimite_test(d.get("catre"))
                    return self.json({"ok": True, "email": setari_email_publice()})
                if p == "/api/setari/email/trimite":
                    r = trimite_mailuri(f"manual ({autor})")
                    return self.json({"ok": True, "rezultat": r, "email": setari_email_publice()})
            if p == "/api/contul-meu":
                modifica_utilizator(u["utilizator"], nume=d.get("nume"), email=d.get("email"))
                return self.json({"ok": True, "eu": public(utilizator_din_token(self.token()))})

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
    threading.Thread(target=planificator_mail, daemon=True, name="planificator-mail").start()
    print(f"{APP_NAME} rulează pe http://{HOST}:{PORT}/   (baza de date: {DB_PATH})")
    if not numar_utilizatori():
        print("  Niciun utilizator încă: deschide tracker-ul în browser și creează contul de administrator.")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("Oprit.")


def main():
    p = argparse.ArgumentParser(description=APP_NAME)
    p.add_argument("--editori", "--utilizatori", dest="editori", action="store_true", help="administrare utilizatori")
    p.add_argument("--adauga-utilizator", nargs=3, metavar=("UTILIZATOR", "NUME", "PAROLA"), help="creează un utilizator fără meniu")
    p.add_argument("--admin", action="store_true", help="cu --adauga-utilizator: îl face administrator")
    p.add_argument("--numar-utilizatori", "--numar-editori", dest="numar", action="store_true", help="câți utilizatori există")
    p.add_argument("--import-excel", metavar="FISIER")
    p.add_argument("--foaie", metavar="NUME", help="foaia din Excel (implicit se detectează)")
    p.add_argument("--import-csv", metavar="FISIER")
    p.add_argument("--raport-mail", action="store_true", help="pregătește mailurile pentru Outlook")
    p.add_argument("--trimite-mail", action="store_true", help="trimite mailurile prin SMTP")
    a = p.parse_args()
    creeaza_db()
    if a.editori:
        return meniu_utilizatori()
    if a.adauga_utilizator:
        try:
            creeaza_utilizator(*a.adauga_utilizator, admin=a.admin)
        except Eroare as e:
            raise SystemExit(str(e))
        return print(f"OK - {a.adauga_utilizator[0]} creat.")
    if a.numar:
        return print(numar_utilizatori())
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
