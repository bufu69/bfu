#!/usr/bin/env python3
"""
Grind Rings Action Tracker - server (port 8092), baza de date SQLite.

Server complet independent: nu are nevoie de Excel, de internet sau de alt server.
Doar Python 3 (fara pachete suplimentare).

Pornire server:            python server.py
Editori (coduri):          python server.py --editori              (sau editori.bat)
Import din Excel:          python server.py --import-excel F.xlsx  (sau import-excel.bat)
Import din CSV:            python server.py --import-csv F.csv
Mailuri prin SMTP:         python server.py --trimite-mail         (sau mail.bat)
Mailuri prin Outlook:      python server.py --raport-mail          (apoi trimite-outlook.ps1;
                                                                     sau outlook-mailuri.bat)

Rute HTTP:
  GET  /                    dashboard-ul (index.html)
  GET  /grafice             pagina cu grafice (grafice.html)
  GET  /api/actions         actiunile (JSON)
  GET  /api/lists           listele pentru formular
  GET  /api/export.csv      export CSV
  GET  /api/history[?no=]   istoricul modificarilor
  POST /  /edit  /delete    adaugare / editare / stergere (cu cod de editor)
  POST /upload              atasament
  GET  /list?no=&source=    lista atasamentelor
  GET  /attachments/<source>/<no>/<fisier>
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
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

# ======================= SETARI =======================
APP_NAME = "Grind Rings Action Tracker"
HOST = "0.0.0.0"
PORT = 8092
DATABASE = "grind_rings.db"          # baza de date (langa server.py)
ATTACHMENTS_DIR = "attachments"      # atasamentele: attachments/<source>/<no>/<fisier>
BACKUP_DIR = "backups"               # backup zilnic al bazei de date
BACKUP_KEEP_DAYS = 30
MAX_UPLOAD_MB = 25
MAIL_INI = "mail.ini"                # setarile pentru mailuri (vezi mail-exemplu.ini)
MAIL_OUT_DIR = "mail_out"            # unde se pregatesc mailurile pentru Outlook
# ======================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def here(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


DB_PATH = here(DATABASE)

# (coloana in baza de date, cheia din formular, antetul trimis paginii / din Excel)
COLUMNS = [
    ("no", "no", "No"),
    ("data_adaugare", "dataAdaugare", "Data adaugare"),
    ("sap_rulment", "sapRulment", "SAP id Rulment"),
    ("desc_rulment", "descRulment", "Descriere Rulment"),
    ("familie_rulment", "familie", "Familie Rulment"),
    ("comanda_montaj", "comandaMontaj", "Comanda Montaj"),
    ("client", "client", "Client"),
    ("sap_component", "sapComponent", "SAP id Component"),
    ("desc_component", "descComponent", "Descriere Component"),
    ("tip_component", "tipComponent", "Tip component"),
    ("comanda_component", "comandaComponent", "Comanda Component"),
    ("qty", "qty", "Qty"),
    ("actiune", "actiune", "Actiune"),
    ("tip_actiune", "tipActiune", "Tip Actiune"),
    ("responsabil", "responsabil", "Responsabil"),
    ("data_planificata", "dataPlanificata", "Data Planificata"),
    ("data_finalizare", "dataActuala", "Data Finalizare"),
    ("stadiu", "stadiu", "Stadiu"),
    ("urgent", "urgent", "Urgent Client"),
    ("comentarii", "comentarii", "Comentarii"),
]
TEXT_FIELDS = ["sapRulment", "descRulment", "familie", "comandaMontaj", "client",
               "sapComponent", "descComponent", "tipComponent", "comandaComponent",
               "qty", "actiune", "tipActiune", "responsabil"]
KEY_TO_COL = {key: col for col, key, _ in COLUMNS}
DATE_COLS = {"data_adaugare", "data_planificata", "data_finalizare"}
LIST_COLUMNS = {"Client": "client", "Responsabil": "responsabil", "Tip Actiune": "tip_actiune",
                "Familie Rulment": "familie_rulment", "Tip component": "tip_component"}

# Recunoasterea coloanelor dupa nume (la fel ca in dashboard): exact, apoi partial.
FIELD_KEYWORDS = {
    "no": ["no", "nr", "nr crt", "numar"],
    "data_adaugare": ["data adaugare", "data creare", "added"],
    "sap_rulment": ["sap id rulment", "sap rulment"],
    "desc_rulment": ["descriere rulment"],
    "familie_rulment": ["familie rulment", "familie"],
    "comanda_montaj": ["comanda montaj"],
    "client": ["client"],
    "sap_component": ["sap id component", "sap component"],
    "desc_component": ["descriere component"],
    "tip_component": ["tip component"],
    "comanda_component": ["comanda component"],
    "qty": ["qty", "cantitate"],
    "tip_actiune": ["tip actiune"],
    "actiune": ["actiune"],
    "responsabil": ["responsabil"],
    "data_planificata": ["data planificata"],
    "data_finalizare": ["data finalizare", "data actuala"],
    "stadiu": ["stadiu", "status"],
    "comentarii": ["comentarii", "observatii", "note"],
    "urgent": ["urgent client", "urgent"],
}

db_lock = threading.Lock()


# ---------------------------------------------------------------- utilitare
def norm(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def to_iso_date(s):
    """'dd.mm.yyyy' / 'yyyy-mm-dd' / 'dd/mm/yyyy' / data / numar serial Excel -> 'yyyy-mm-dd'."""
    if s is None or s == "":
        return None
    if isinstance(s, datetime):
        return s.date().isoformat()
    if isinstance(s, date):
        return s.isoformat()
    if isinstance(s, (int, float)) and 20000 < s < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=float(s))).date().isoformat()
    s = str(s).strip()
    if not s:
        return None
    s10 = s[:10]
    for fmt, txt in (("%d.%m.%Y", s), ("%Y-%m-%d", s10), ("%d/%m/%Y", s), ("%m/%d/%Y", s)):
        try:
            return datetime.strptime(txt, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def fmt_ro(iso):
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except (TypeError, ValueError):
        return str(iso or "")


def is_urgent(v):
    return norm(v) in ("da", "yes", "true", "1", "x")


def normalize_stadiu(v):
    s = norm(v)
    if not s:
        return "Deschis"
    if re.search(r"rezolv|inchis|closed|done|finalizat", s):
        return "Rezolvat"
    if re.search(r"lucru|progress|wip|analiza", s):
        return "În lucru"
    return "Deschis"


def safe_part(s):
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s)).strip().strip(".")
    return s or "_"


def stamp_comment(text, user):
    text = str(text or "").strip()
    if not text:
        return ""
    return f"[{datetime.now():%d.%m.%Y %H:%M} {user}] {text}"


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def detect_columns(headers):
    normed = [norm(h) for h in headers]
    colmap, used = {}, set()
    for exact in (True, False):
        for field, kws in FIELD_KEYWORDS.items():
            if field in colmap:
                continue
            for kw in kws:
                idx = next((i for i, h in enumerate(normed)
                            if h and i not in used and (h == kw if exact else kw in h)), None)
                if idx is not None:
                    colmap[field] = idx
                    used.add(idx)
                    break
    return colmap


# ---------------------------------------------------------------- baza de date
def connect():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    cols = ",\n  ".join(f"{c} TEXT" for c, _, _ in COLUMNS if c != "no")
    with connect() as conn:
        conn.executescript(f"""
CREATE TABLE IF NOT EXISTS actions (
  no INTEGER PRIMARY KEY AUTOINCREMENT,
  {cols},
  created_by TEXT, created_at TEXT, updated_by TEXT, updated_at TEXT
);
CREATE TABLE IF NOT EXISTS history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  action_no INTEGER, operation TEXT, user TEXT, at TEXT, details TEXT
);
CREATE INDEX IF NOT EXISTS ix_history_no ON history(action_no);
CREATE TABLE IF NOT EXISTS editori (
  nume TEXT PRIMARY KEY, salt TEXT, pin_hash TEXT, creat_la TEXT
);
CREATE TABLE IF NOT EXISTS liste (
  lista TEXT, valoare TEXT, PRIMARY KEY (lista, valoare)
);
""")


def log_history(conn, no, operation, user, details):
    conn.execute("INSERT INTO history(action_no, operation, user, at, details) VALUES (?,?,?,?,?)",
                 (no, operation, user, now_str(), json.dumps(details, ensure_ascii=False)))


def daily_backup():
    if not os.path.exists(DB_PATH):
        return
    bdir = here(BACKUP_DIR)
    os.makedirs(bdir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(DB_PATH))
    target = os.path.join(bdir, f"{stem}_{date.today():%Y-%m-%d}{ext}")
    if not os.path.exists(target):
        with connect() as s, sqlite3.connect(target) as d:
            s.backup(d)
    olds = sorted(f for f in os.listdir(bdir) if f.startswith(stem + "_") and f.endswith(ext))
    for f in olds[:-BACKUP_KEEP_DAYS] if BACKUP_KEEP_DAYS > 0 else []:
        try:
            os.remove(os.path.join(bdir, f))
        except OSError:
            pass


# ---------------------------------------------------------------- editori (coduri personale)
def hash_pin(pin, salt):
    return hashlib.pbkdf2_hmac("sha256", str(pin).encode(), bytes.fromhex(salt), 100_000).hex()


def check_pin(pin):
    pin = str(pin or "").strip()
    if not pin:
        return None
    with connect() as conn:
        for r in conn.execute("SELECT nume, salt, pin_hash FROM editori"):
            if secrets.compare_digest(hash_pin(pin, r["salt"]), r["pin_hash"]):
                return r["nume"]
    return None


def set_editor(nume, pin):
    salt = secrets.token_hex(16)
    with connect() as conn:
        for r in conn.execute("SELECT nume, salt, pin_hash FROM editori WHERE nume<>?", (nume,)):
            if hash_pin(pin, r["salt"]) == r["pin_hash"]:
                raise ValueError(f"Codul este deja folosit de {r['nume']}. Alege alt cod.")
        conn.execute("INSERT OR REPLACE INTO editori(nume, salt, pin_hash, creat_la) VALUES (?,?,?,?)",
                     (nume, salt, hash_pin(pin, salt), now_str()))


def list_editori():
    with connect() as conn:
        return [dict(r) for r in conn.execute("SELECT nume, creat_la FROM editori ORDER BY nume")]


def delete_editor(nume):
    with connect() as conn:
        return conn.execute("DELETE FROM editori WHERE nume=?", (nume,)).rowcount > 0


def editori_menu():
    while True:
        eds = list_editori()
        print(f"\n=== {APP_NAME} - editori ===")
        if eds:
            for i, e in enumerate(eds, 1):
                print(f"  {i}. {e['nume']}   (din {e['creat_la'][:10]})")
        else:
            print("  (niciun editor - nimeni nu poate adauga/modifica actiuni)")
        print("\n  1 = Adauga editor / schimba codul     2 = Sterge editor     0 = Iesire")
        opt = input("Alege: ").strip()
        if opt == "1":
            nume = input("Numele editorului (apare in comentarii): ").strip()
            if not nume:
                continue
            pin = getpass.getpass("Codul personal (nu se afiseaza): ").strip()
            pin2 = getpass.getpass("Repeta codul: ").strip()
            if not pin or pin != pin2:
                print("Codurile nu coincid sau sunt goale.")
                continue
            try:
                set_editor(nume, pin)
                print(f"OK - {nume} poate edita.")
            except ValueError as e:
                print(e)
        elif opt == "2":
            nume = input("Numele editorului de sters: ").strip()
            print("Sters." if delete_editor(nume) else "Nu exista acest editor.")
        elif opt == "0" or opt == "":
            return


# ---------------------------------------------------------------- actiuni
def field_values(data):
    vals = {}
    for key in TEXT_FIELDS:
        if key in data:
            vals[KEY_TO_COL[key]] = str(data.get(key) or "").strip()
    if "dataPlanificata" in data:
        vals["data_planificata"] = to_iso_date(data["dataPlanificata"])
    if "urgent" in data:
        vals["urgent"] = "Da" if is_urgent(data["urgent"]) else "Nu"
    return vals


def insert_action(conn, vals, user, op="adaugare"):
    vals = dict(vals)
    vals.update(created_by=user, created_at=now_str(), updated_by=user, updated_at=now_str())
    cur = conn.execute(f"INSERT INTO actions ({','.join(vals)}) VALUES ({','.join('?' * len(vals))})",
                       list(vals.values()))
    no = vals.get("no") or cur.lastrowid
    log_history(conn, no, op, user, vals)
    return no


def add_action(data, user):
    vals = field_values({"urgent": "Nu", **data})
    vals["data_adaugare"] = to_iso_date(data.get("dataAdaugare")) or date.today().isoformat()
    vals["stadiu"] = data.get("stadiu") or "Deschis"
    vals["comentarii"] = stamp_comment(data.get("comentarii"), user)
    with db_lock:
        daily_backup()
        with connect() as conn:
            return insert_action(conn, vals, user)


def get_action(conn, no):
    try:
        no = int(float(str(no).strip()))
    except (TypeError, ValueError):
        return None
    return conn.execute("SELECT * FROM actions WHERE no=?", (no,)).fetchone()


def edit_action(data, user):
    with db_lock:
        daily_backup()
        with connect() as conn:
            row = get_action(conn, data.get("no"))
            if row is None:
                return False
            vals = field_values(data)
            stadiu = data.get("stadiu")
            if stadiu:
                vals["stadiu"] = stadiu
                if stadiu == "Rezolvat" and not row["data_finalizare"]:
                    vals["data_finalizare"] = date.today().isoformat()
                elif stadiu != "Rezolvat" and "rezolv" in norm(row["stadiu"]):
                    vals["data_finalizare"] = None
            new_comment = stamp_comment(data.get("comentarii"), user)
            if new_comment:
                old = (row["comentarii"] or "").strip()
                vals["comentarii"] = f"{old}\n{new_comment}" if old else new_comment
            changes = {k: {"inainte": row[k], "dupa": v} for k, v in vals.items() if row[k] != v}
            vals.update(updated_by=user, updated_at=now_str())
            conn.execute(f"UPDATE actions SET {','.join(f'{k}=?' for k in vals)} WHERE no=?",
                         [*vals.values(), row["no"]])
            log_history(conn, row["no"], "editare", user, changes)
    return True


def delete_action(no, user):
    with db_lock:
        daily_backup()
        with connect() as conn:
            row = get_action(conn, no)
            if row is None:
                return False
            conn.execute("DELETE FROM actions WHERE no=?", (row["no"],))
            log_history(conn, row["no"], "stergere", user, dict(row))
    return True


def all_actions():
    with connect() as conn:
        rows = conn.execute("SELECT * FROM actions ORDER BY no").fetchall()
    return [{h: ("" if r[c] is None else r[c]) for c, _, h in COLUMNS} for r in rows]


def form_lists():
    lists = {}
    with connect() as conn:
        for header, col in LIST_COLUMNS.items():
            values = [v for (v,) in conn.execute("SELECT valoare FROM liste WHERE lista=? ORDER BY valoare", (header,))]
            for (v,) in conn.execute(f"SELECT DISTINCT {col} FROM actions WHERE {col} IS NOT NULL AND {col}<>'' ORDER BY {col}"):
                if v not in values:
                    values.append(v)
            lists[header] = values
    return lists


def export_csv():
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow([h for _, _, h in COLUMNS])
    for item in all_actions():
        w.writerow([item[h] for _, _, h in COLUMNS])
    return ("﻿" + buf.getvalue()).encode("utf-8")


def history(no=None):
    with connect() as conn:
        if no:
            rows = conn.execute("SELECT * FROM history WHERE action_no=? ORDER BY id DESC", (no,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM history ORDER BY id DESC LIMIT 500").fetchall()
    return [{**dict(r), "details": json.loads(r["details"] or "null")} for r in rows]


# ---------------------------------------------------------------- import Excel / CSV
def read_xlsx(path):
    """Citeste un .xlsx/.xlsm fara pachete externe. Intoarce {nume_foaie: [[celule]]}."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels}
        sheets = {}
        for s in wb.find("m:sheets", ns):
            target = targets[s.get(f"{{{ns['r']}}}id")].lstrip("/")
            target = target if target.startswith("xl/") else "xl/" + target
            rows = []
            for row in ET.fromstring(z.read(target)).iter(f"{{{ns['m']}}}row"):
                cells = {}
                for c in row.findall("m:c", ns):
                    ref = re.match(r"([A-Z]+)", c.get("r", "A")).group(1)
                    col = 0
                    for ch in ref:
                        col = col * 26 + ord(ch) - 64
                    t, v = c.get("t"), c.find("m:v", ns)
                    if t == "s" and v is not None:
                        val = shared[int(v.text)]
                    elif t == "inlineStr":
                        val = "".join(x.text or "" for x in c.iter(f"{{{ns['m']}}}t"))
                    elif t in ("str", "e"):
                        val = v.text if v is not None else ""
                    elif t == "b":
                        val = "Da" if v is not None and v.text == "1" else "Nu"
                    elif v is not None and v.text is not None:
                        f = float(v.text)
                        val = int(f) if f.is_integer() else f
                    else:
                        val = ""
                    cells[col - 1] = val
                rows.append([cells.get(i, "") for i in range(max(cells) + 1)] if cells else [])
            sheets[s.get("name")] = rows
        return sheets


def import_rows(header, rows, user, sursa):
    colmap = detect_columns(header)
    if "actiune" not in colmap and "client" not in colmap:
        raise ValueError("Nu am gasit coloanele Actiune/Client in antet.")
    added = skipped = 0
    with db_lock:
        daily_backup()
        with connect() as conn:
            for r in rows:
                cell = lambda f: (r[colmap[f]] if f in colmap and colmap[f] < len(r) else "")
                vals = {}
                for col, _, _ in COLUMNS:
                    v = cell(col)
                    if col in DATE_COLS:
                        v = to_iso_date(v)
                    elif col == "no":
                        try:
                            v = int(float(v)) if str(v).strip() else None
                        except ValueError:
                            v = None
                    else:
                        v = "" if v is None else str(v).strip()
                    vals[col] = v
                if not any(vals[c] for c in ("actiune", "client", "comanda_montaj")):
                    continue
                vals["stadiu"] = normalize_stadiu(vals.get("stadiu"))
                vals["urgent"] = "Da" if is_urgent(vals.get("urgent")) else "Nu"
                if not vals.get("data_adaugare"):
                    vals["data_adaugare"] = date.today().isoformat()
                if vals["no"] is None:
                    vals.pop("no")
                elif get_action(conn, vals["no"]) is not None:
                    skipped += 1
                    continue
                no = insert_action(conn, vals, user, op=f"import {sursa}")
                added += 1
                print(f"  + #{no}: {vals.get('client', '')} - {vals.get('actiune', '')}")
    return added, skipped


def import_excel(path, sheet=None):
    sheets = read_xlsx(path)
    if sheet:
        name = sheet
    else:
        # foaia cu actiuni: prima care are coloana "Actiune" in antet
        name = next((n for n, rows in sheets.items()
                     if rows and "actiune" in detect_columns(rows[0])), None)
    if not name or name not in sheets:
        raise ValueError(f"Nu am gasit foaia cu actiuni. Foi disponibile: {', '.join(sheets)}")
    rows = sheets[name]
    print(f"Foaia '{name}': {len(rows) - 1} randuri")
    added, skipped = import_rows(rows[0], rows[1:], "import Excel", "Excel")
    # foaia "Lists" (optional): valorile pentru listele din formular
    listsheet = next((n for n in sheets if norm(n) in ("lists", "liste")), None)
    nliste = 0
    if listsheet and sheets[listsheet]:
        hdr = sheets[listsheet][0]
        with connect() as conn:
            for i, h in enumerate(hdr):
                key = next((k for k in LIST_COLUMNS if norm(k) == norm(h) or norm(h) == norm(k).split()[0]), None)
                if not key:
                    continue
                for row in sheets[listsheet][1:]:
                    v = str(row[i]).strip() if i < len(row) else ""
                    if v:
                        nliste += conn.execute("INSERT OR IGNORE INTO liste(lista, valoare) VALUES (?,?)", (key, v)).rowcount
    print(f"Gata: {added} actiuni importate, {skipped} sarite (numarul exista deja), {nliste} valori noi in liste.")


def import_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=",;") if sample else csv.excel
        data = list(csv.reader(f, dialect=dialect))
    if not data:
        print("Fisier gol.")
        return
    added, skipped = import_rows(data[0], data[1:], "import CSV", "CSV")
    done = path + f".importat_{datetime.now():%Y%m%d_%H%M%S}"
    os.replace(path, done)
    print(f"Gata: {added} actiuni importate, {skipped} sarite. Fisierul a fost redenumit in: {done}")


# ---------------------------------------------------------------- mailuri
def load_mail_ini():
    path = here(MAIL_INI)
    if not os.path.exists(path):
        raise SystemExit(f"Lipseste {MAIL_INI}. Copiaza mail-exemplu.ini in {MAIL_INI} si completeaza-l.")
    cp = configparser.ConfigParser(interpolation=None)
    cp.optionxform = str
    cp.read(path, encoding="utf-8")
    return cp


def build_mails(cp):
    """Pregateste mailurile: cate unul pe responsabil (actiuni intarziate / aproape de termen / urgente)
    + un rezumat. Intoarce lista de {to, cc, subject, html}."""
    g = cp["general"] if cp.has_section("general") else {}
    zile = int(g.get("zile_inainte_termen", "2") or 2)
    url = g.get("link_dashboard", "").strip()
    subiect = g.get("subiect", "Grind Rings - actiuni de urmarit").strip()
    adrese = {norm(k): v.strip() for k, v in (cp["destinatari"].items() if cp.has_section("destinatari") else [])}
    cc = g.get("cc", "").strip()
    rezumat_catre = g.get("rezumat_catre", "").strip()

    today = date.today()
    per_resp, toate = {}, []
    for a in all_actions():
        if a["Stadiu"] == "Rezolvat":
            continue
        try:
            termen = datetime.strptime(str(a["Data Planificata"])[:10], "%Y-%m-%d").date()
            zile_ramase = (termen - today).days
        except ValueError:
            zile_ramase = None
        urgent = a["Urgent Client"] == "Da"
        if not (urgent or (zile_ramase is not None and zile_ramase <= zile)):
            continue
        a = dict(a, _zile=zile_ramase, _urgent=urgent)
        toate.append(a)
        per_resp.setdefault(a["Responsabil"] or "Nealocat", []).append(a)

    def tabel(items):
        items = sorted(items, key=lambda x: (x["_zile"] if x["_zile"] is not None else 9999))
        rows = []
        for a in items:
            if a["_zile"] is None:
                st, col = "fără termen", "#8d97ac"
            elif a["_zile"] < 0:
                st, col = f"întârziată {-a['_zile']} zile", "#e1596b"
            elif a["_zile"] == 0:
                st, col = "termen azi", "#e8a33d"
            else:
                st, col = f"termen în {a['_zile']} zile", "#e8a33d"
            urg = ' <b style="color:#e1596b">URGENT CLIENT</b>' if a["_urgent"] else ""
            rows.append(
                f"<tr><td>#{a['No']}</td><td>{html.escape(str(a['Client']))}</td>"
                f"<td>{html.escape(str(a['Actiune']))}{urg}</td><td>{html.escape(str(a['Responsabil']))}</td>"
                f"<td>{fmt_ro(a['Data Planificata'])}</td><td style=\"color:{col}\">{st}</td>"
                f"<td>{html.escape(str(a['Stadiu']))}</td></tr>")
        return ('<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;'
                'font-family:Segoe UI,Arial,sans-serif;font-size:13px;border-color:#ddd">'
                '<tr style="background:#f2f4f7"><th>Nr</th><th>Client</th><th>Acțiune</th><th>Responsabil</th>'
                '<th>Termen</th><th>Situație</th><th>Stadiu</th></tr>' + "".join(rows) + "</table>")

    link = f'<p><a href="{html.escape(url)}">Deschide {APP_NAME}</a></p>' if url else ""

    def corp(titlu, items):
        return (f'<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px">'
                f"<p>{titlu}</p>{tabel(items)}{link}"
                f'<p style="color:#888;font-size:12px">Mesaj automat {APP_NAME}, {today:%d.%m.%Y}.</p></div>')

    mails = []
    for resp, items in sorted(per_resp.items()):
        to = adrese.get(norm(resp))
        if not to:
            print(f"  ! {resp}: nu are adresa in [destinatari] - apare doar in rezumat")
            continue
        mails.append({"to": to, "cc": cc, "subject": f"{subiect} ({len(items)})",
                      "html": corp(f"Bună ziua, {html.escape(resp)}. Aveți {len(items)} acțiuni care necesită atenție:", items)})
    if rezumat_catre and toate:
        mails.append({"to": rezumat_catre, "cc": "", "subject": f"{subiect} - rezumat ({len(toate)})",
                      "html": corp(f"Rezumat: {len(toate)} acțiuni întârziate, aproape de termen sau urgente.", toate)})
    return mails


def raport_mail():
    """Scrie mailurile in mail_out/ pentru trimite-outlook.ps1."""
    mails = build_mails(load_mail_ini())
    out = here(MAIL_OUT_DIR)
    os.makedirs(out, exist_ok=True)
    for f in os.listdir(out):
        if f.endswith(".html") or f == "mailuri.json":
            os.remove(os.path.join(out, f))
    index = []
    for i, m in enumerate(mails, 1):
        fn = f"mail_{i:02d}.html"
        with open(os.path.join(out, fn), "w", encoding="utf-8") as f:
            f.write(m["html"])
        index.append({"to": m["to"], "cc": m["cc"], "subject": m["subject"], "file": fn})
    with open(os.path.join(out, "mailuri.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    print(f"{len(mails)} mailuri pregatite in {out}")


def trimite_mail_smtp():
    cp = load_mail_ini()
    if not cp.has_section("smtp"):
        raise SystemExit("Lipseste sectiunea [smtp] din mail.ini (sau foloseste outlook-mailuri.bat).")
    s = cp["smtp"]
    mails = build_mails(cp)
    if not mails:
        print("Nimic de trimis.")
        return
    port = int(s.get("port", "25") or 25)
    secur = norm(s.get("securitate", "starttls"))
    smtp = smtplib.SMTP_SSL(s["server"], port, timeout=30) if secur == "ssl" else smtplib.SMTP(s["server"], port, timeout=30)
    with smtp:
        if secur == "starttls":
            smtp.starttls()
        if s.get("utilizator", "").strip():
            smtp.login(s["utilizator"].strip(), s.get("parola", ""))
        for m in mails:
            msg = MIMEMultipart("alternative")
            msg["Subject"], msg["From"], msg["To"] = m["subject"], s["de_la"].strip(), m["to"]
            if m["cc"]:
                msg["Cc"] = m["cc"]
            msg.attach(MIMEText(m["html"], "html", "utf-8"))
            to_all = [x.strip() for x in re.split(r"[;,]", f"{m['to']};{m['cc']}") if x.strip()]
            smtp.sendmail(msg["From"], to_all, msg.as_string())
            print(f"  trimis: {m['to']} - {m['subject']}")
    print(f"{len(mails)} mailuri trimise.")


# ---------------------------------------------------------------- atasamente
def attach_dir(source, no):
    return os.path.join(here(ATTACHMENTS_DIR), safe_part(source or "grindrings"), safe_part(no))


def list_attachments(source, no):
    d = attach_dir(source, no)
    return sorted(f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f))) if os.path.isdir(d) else []


def save_attachment(source, no, filename, content_b64):
    raw = base64.b64decode(content_b64 or "", validate=False)
    if len(raw) > MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(f"Fișierul depășește {MAX_UPLOAD_MB} MB.")
    d = attach_dir(source, no)
    os.makedirs(d, exist_ok=True)
    name = safe_part(os.path.basename(str(filename or "fisier")))
    stem, ext = os.path.splitext(name)
    target, i = os.path.join(d, name), 1
    while os.path.exists(target):
        target = os.path.join(d, f"{stem} ({i}){ext}")
        i += 1
    with open(target, "wb") as f:
        f.write(raw)
    return os.path.basename(target)


# ---------------------------------------------------------------- HTTP
PAGES = {"/": "index.html", "/index.html": "index.html",
         "/grafice": "grafice.html", "/grafice.html": "grafice.html"}


class Handler(BaseHTTPRequestHandler):
    server_version = "GrindRingsTracker/3.0"

    def log_message(self, fmt, *args):
        sys.stdout.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {self.client_address[0]} {fmt % args}\n")

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def send_json(self, status, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self._cors()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.end_headers()
        self.wfile.write(body)

    def send_bytes(self, data, ctype, extra_headers=None):
        self.send_response(200)
        self._cors()
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        for k, v in (extra_headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path, ctype=None, download_name=None):
        if not os.path.isfile(path):
            return self.send_error(HTTPStatus.NOT_FOUND, "Fisier inexistent")
        ctype = ctype or mimetypes.guess_type(path)[0] or "application/octet-stream"
        with open(path, "rb") as f:
            data = f.read()
        headers = {"Content-Disposition": f"inline; filename*=UTF-8''{quote(download_name)}"} if download_name else {}
        self.send_bytes(data, ctype, headers)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        url = urlparse(self.path)
        path = unquote(url.path)
        qs = parse_qs(url.query)
        try:
            if path in PAGES:
                return self.send_file(here(PAGES[path]), "text/html; charset=utf-8")
            if path.startswith("/fonts/"):
                fname = os.path.basename(path)
                ctype = "text/css; charset=utf-8" if fname.endswith(".css") else "font/woff2"
                return self.send_file(os.path.join(BASE_DIR, "fonts", fname), ctype)
            if path == "/api/actions":
                return self.send_json(200, {"ok": True, "rows": all_actions()})
            if path == "/api/lists":
                return self.send_json(200, {"ok": True, "lists": form_lists()})
            if path == "/api/export.csv":
                name = f"grind_rings_actiuni_{date.today():%Y-%m-%d}.csv"
                return self.send_bytes(export_csv(), "text/csv; charset=utf-8",
                                       {"Content-Disposition": f"attachment; filename={name}"})
            if path == "/api/history":
                no = (qs.get("no") or [""])[0]
                return self.send_json(200, {"ok": True, "history": history(no or None)})
            if path == "/list":
                no = (qs.get("no") or [""])[0]
                source = (qs.get("source") or ["grindrings"])[0]
                return self.send_json(200, {"ok": True, "files": list_attachments(source, no)})
            if path.startswith("/attachments/"):
                parts = [p for p in path[len("/attachments/"):].split("/") if p]
                if len(parts) != 3:
                    return self.send_error(HTTPStatus.NOT_FOUND)
                source, no, fname = parts
                full = os.path.join(attach_dir(source, no), safe_part(fname))
                root = os.path.realpath(here(ATTACHMENTS_DIR))
                if not os.path.realpath(full).startswith(root + os.sep):
                    return self.send_error(HTTPStatus.FORBIDDEN)
                return self.send_file(full, download_name=fname)
            if path == "/health":
                return self.send_json(200, {"ok": True, "app": APP_NAME, "database": DB_PATH})
        except sqlite3.Error as e:
            return self.send_json(500, {"ok": False, "error": f"Baza de date: {e}"})
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        path = urlparse(self.path).path.rstrip("/") or "/"
        try:
            length = int(self.headers.get("Content-Length") or 0)
            data = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except Exception:  # noqa: BLE001
            return self.send_json(400, {"ok": False, "error": "JSON invalid"})
        if path not in ("/", "/edit", "/delete", "/upload"):
            return self.send_json(404, {"ok": False, "error": "Rută necunoscută"})
        user = check_pin(data.get("pin"))
        if not user:
            return self.send_json(403, {"ok": False, "error": "Cod incorect"})
        try:
            if path == "/":
                for f in ("client", "responsabil", "actiune"):
                    if not str(data.get(f) or "").strip():
                        return self.send_json(400, {"ok": False, "error": f"Câmpul „{f}” este obligatoriu"})
                no = add_action(data, user)
                print(f"  + actiune #{no} adaugata de {user}")
                return self.send_json(200, {"ok": True, "no": no})
            if path == "/edit":
                if not edit_action(data, user):
                    return self.send_json(404, {"ok": False, "error": "Acțiunea nu a fost găsită"})
                print(f"  ~ actiune #{data.get('no')} editata de {user}")
                return self.send_json(200, {"ok": True})
            if path == "/delete":
                if not delete_action(data.get("no"), user):
                    return self.send_json(404, {"ok": False, "error": "Acțiunea nu a fost găsită"})
                print(f"  - actiune #{data.get('no')} stearsa de {user}")
                return self.send_json(200, {"ok": True})
            name = save_attachment(data.get("source"), data.get("no"), data.get("filename"), data.get("contentBase64"))
            print(f"  @ atasament '{name}' la #{data.get('no')} de {user}")
            return self.send_json(200, {"ok": True, "filename": name})
        except sqlite3.Error as e:
            return self.send_json(500, {"ok": False, "error": f"Baza de date: {e}"})
        except (RuntimeError, ValueError) as e:
            return self.send_json(500, {"ok": False, "error": str(e)})
        except Exception as e:  # noqa: BLE001
            return self.send_json(500, {"ok": False, "error": f"{type(e).__name__}: {e}"})


def run_server():
    # pornit in fundal (pythonw / Task Scheduler): mesajele merg in server.log
    if sys.stdout is None or sys.stderr is None:
        log = open(here("server.log"), "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = log
    os.makedirs(here(ATTACHMENTS_DIR), exist_ok=True)
    httpd = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"{APP_NAME} ruleaza pe http://{HOST}:{PORT}/")
    print(f"  Baza de date: {DB_PATH}")
    print(f"  Atasamente:   {here(ATTACHMENTS_DIR)}")
    if not list_editori():
        print("  ATENTIE: nu exista niciun editor. Ruleaza editori.bat ca sa adaugi coduri.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nOprit.")


def main():
    p = argparse.ArgumentParser(description=APP_NAME)
    p.add_argument("--editori", action="store_true", help="administrare editori (coduri personale)")
    p.add_argument("--adauga-editor", nargs=2, metavar=("NUME", "COD"), help="adauga/schimba un editor fara meniu")
    p.add_argument("--import-excel", metavar="FISIER", help="importa actiuni dintr-un .xlsx/.xlsm")
    p.add_argument("--foaie", metavar="NUME", help="foaia din Excel (implicit: se detecteaza)")
    p.add_argument("--import-csv", metavar="FISIER", help="importa actiuni dintr-un CSV")
    p.add_argument("--raport-mail", action="store_true", help="pregateste mailurile in mail_out/ (pentru Outlook)")
    p.add_argument("--trimite-mail", action="store_true", help="trimite mailurile prin SMTP (mail.ini)")
    a = p.parse_args()
    init_db()
    if a.editori:
        return editori_menu()
    if a.adauga_editor:
        try:
            set_editor(*a.adauga_editor)
        except ValueError as e:
            raise SystemExit(str(e))
        return print(f"OK - {a.adauga_editor[0]} poate edita.")
    if a.import_excel:
        return import_excel(a.import_excel, a.foaie)
    if a.import_csv:
        return import_csv(a.import_csv)
    if a.raport_mail:
        return raport_mail()
    if a.trimite_mail:
        return trimite_mail_smtp()
    run_server()


if __name__ == "__main__":
    main()
