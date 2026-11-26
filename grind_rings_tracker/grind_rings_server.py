#!/usr/bin/env python3
"""
Grind Rings Action Tracker - server Python (port 8092), cu baza de date SQLite.

Un singur proces, fara Excel si fara Start-ActionListener.ps1:
  * serveste dashboard-ul (static/index.html)            GET  /
  * actiunile din baza de date (JSON)                    GET  /api/actions
  * listele pentru formular (clienti, responsabili...)   GET  /api/lists
  * export CSV al tuturor actiunilor                     GET  /api/export.csv
  * istoricul modificarilor                              GET  /api/history?no=..
  * fonturile paginii (fara internet)                    GET  /static/fonts/...
  * adauga o actiune noua                                POST /
  * editeaza o actiune                                   POST /edit
  * sterge o actiune                                     POST /delete
  * ataseaza un fisier la o actiune                      POST /upload
  * listeaza fisierele atasate                           GET  /list?no=..&source=..
  * serveste fisierele atasate                           GET  /attachments/<source>/<no>/<fisier>

Pornire:  python grind_rings_server.py        (doar Python 3, fara pachete externe)
Setari:   config.json langa acest script (vezi config.example.json).
"""

import argparse
import base64
import csv
import io
import json
import mimetypes
import os
import re
import sqlite3
import sys
import threading
import unicodedata
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

APP_NAME = "Grind Rings Action Tracker"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG = {
    "host": "0.0.0.0",
    "port": 8092,
    # Baza de date SQLite (se creeaza automat la prima pornire).
    "database": "grind_rings.db",
    # Folderul unde se salveaza atasamentele: <folder>/<source>/<no>/<fisier>
    "attachments_dir": "attachments",
    # Codurile personale: "cod": "Nume". Numele apare in comentarii si in istoric.
    "pins": {"1234": "Admin"},
    # Valori fixe pentru listele din formular (pe langa cele deja folosite in actiuni).
    "lists": {
        "Client": [],
        "Responsabil": [],
        "Tip Actiune": [],
        "Familie Rulment": [],
        "Tip component": [],
    },
    "max_upload_mb": 25,
    # Backup zilnic al bazei de date, inainte de prima modificare din zi.
    "backup_dir": "backups",
    "backup_keep_days": 30,
}

# (coloana in baza de date, cheia din formular/payload, antetul trimis paginii)
# Antetele sunt aceleasi ca in Excelul vechi, ca pagina sa le recunoasca la fel.
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
LIST_COLUMNS = {"Client": "client", "Responsabil": "responsabil", "Tip Actiune": "tip_actiune",
                "Familie Rulment": "familie_rulment", "Tip component": "tip_component"}

db_lock = threading.Lock()


# ---------------------------------------------------------------- config
def load_config():
    cfg = dict(DEFAULT_CONFIG)
    path = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            cfg.update(json.load(f))
    for key in ("database", "attachments_dir", "backup_dir"):
        if not os.path.isabs(cfg[key]):
            cfg[key] = os.path.join(BASE_DIR, cfg[key])
    cfg["pins"] = {str(k): str(v) for k, v in cfg.get("pins", {}).items()}
    return cfg


CFG = load_config()


# ---------------------------------------------------------------- helpers
def norm(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def to_iso_date(s):
    """'dd.mm.yyyy' / 'yyyy-mm-dd' / 'dd/mm/yyyy' -> 'yyyy-mm-dd'. Altceva ramane neschimbat."""
    s = str(s or "").strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


def is_urgent(v):
    return norm(v) in ("da", "yes", "true", "1", "x")


def safe_part(s):
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s)).strip().strip(".")
    return s or "_"


def check_pin(pin):
    return CFG["pins"].get(str(pin or "").strip())


def stamp_comment(text, user):
    text = str(text or "").strip()
    if not text:
        return ""
    return f"[{datetime.now():%d.%m.%Y %H:%M} {user}] {text}"


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- baza de date
def connect():
    conn = sqlite3.connect(CFG["database"], timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    os.makedirs(os.path.dirname(CFG["database"]), exist_ok=True)
    cols = ",\n  ".join(f"{c} TEXT" for c, _, _ in COLUMNS if c != "no")
    with connect() as conn:
        conn.executescript(f"""
CREATE TABLE IF NOT EXISTS actions (
  no INTEGER PRIMARY KEY AUTOINCREMENT,
  {cols},
  created_by TEXT,
  created_at TEXT,
  updated_by TEXT,
  updated_at TEXT
);
CREATE TABLE IF NOT EXISTS history (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  action_no INTEGER,
  operation TEXT,
  user TEXT,
  at TEXT,
  details TEXT
);
CREATE INDEX IF NOT EXISTS ix_history_no ON history(action_no);
""")


def log_history(conn, no, operation, user, details):
    conn.execute("INSERT INTO history(action_no, operation, user, at, details) VALUES (?,?,?,?,?)",
                 (no, operation, user, now_str(), json.dumps(details, ensure_ascii=False)))


def daily_backup():
    """Copie a bazei de date o data pe zi, inainte de prima modificare; sterge copiile vechi."""
    src, bdir = CFG["database"], CFG["backup_dir"]
    if not os.path.exists(src):
        return
    os.makedirs(bdir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(src))
    target = os.path.join(bdir, f"{stem}_{date.today():%Y-%m-%d}{ext}")
    if not os.path.exists(target):
        with connect() as s, sqlite3.connect(target) as d:
            s.backup(d)
    keep = int(CFG.get("backup_keep_days") or 0)
    if keep > 0:
        olds = sorted(f for f in os.listdir(bdir) if f.startswith(stem + "_") and f.endswith(ext))
        for f in olds[:-keep]:
            try:
                os.remove(os.path.join(bdir, f))
            except OSError:
                pass


def field_values(data):
    """Valorile din payload care se scriu in baza de date (doar cele prezente)."""
    vals = {}
    for key in TEXT_FIELDS:
        if key in data:
            vals[KEY_TO_COL[key]] = str(data.get(key) or "").strip()
    if "dataPlanificata" in data:
        vals["data_planificata"] = to_iso_date(data["dataPlanificata"])
    if "urgent" in data:
        vals["urgent"] = "Da" if is_urgent(data["urgent"]) else "Nu"
    return vals


def add_action(data, user):
    vals = field_values({"urgent": "Nu", **data})
    vals["data_adaugare"] = to_iso_date(data.get("dataAdaugare")) or date.today().isoformat()
    vals["stadiu"] = data.get("stadiu") or "Deschis"
    vals["comentarii"] = stamp_comment(data.get("comentarii"), user)
    vals.update(created_by=user, created_at=now_str(), updated_by=user, updated_at=now_str())
    with db_lock:
        daily_backup()
        with connect() as conn:
            cur = conn.execute(f"INSERT INTO actions ({','.join(vals)}) VALUES ({','.join('?' * len(vals))})",
                               list(vals.values()))
            no = cur.lastrowid
            log_history(conn, no, "adaugare", user, vals)
    return no


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
    """Actiunile cu antetele din Excelul vechi, ca pagina sa le citeasca la fel."""
    with connect() as conn:
        rows = conn.execute("SELECT * FROM actions ORDER BY no").fetchall()
    out = []
    for r in rows:
        item = {}
        for col, _, header in COLUMNS:
            v = r[col]
            item[header] = "" if v is None else v
        out.append(item)
    return out


def form_lists():
    lists = {}
    with connect() as conn:
        for header, col in LIST_COLUMNS.items():
            values = [str(v).strip() for v in CFG.get("lists", {}).get(header, []) if str(v).strip()]
            for (v,) in conn.execute(f"SELECT DISTINCT {col} FROM actions WHERE {col} IS NOT NULL AND {col} <> '' ORDER BY {col}"):
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


# ---------------------------------------------------------------- atasamente
def attach_dir(source, no):
    return os.path.join(CFG["attachments_dir"], safe_part(source or "grindrings"), safe_part(no))


def list_attachments(source, no):
    d = attach_dir(source, no)
    if not os.path.isdir(d):
        return []
    return sorted(f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f)))


def save_attachment(source, no, filename, content_b64):
    raw = base64.b64decode(content_b64 or "", validate=False)
    if len(raw) > CFG["max_upload_mb"] * 1024 * 1024:
        raise ValueError(f"Fișierul depășește {CFG['max_upload_mb']} MB.")
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
class Handler(BaseHTTPRequestHandler):
    server_version = "GrindRingsTracker/2.0"

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
            self.send_error(HTTPStatus.NOT_FOUND, "Fisier inexistent")
            return
        ctype = ctype or mimetypes.guess_type(path)[0] or "application/octet-stream"
        with open(path, "rb") as f:
            data = f.read()
        headers = {}
        if download_name:
            headers["Content-Disposition"] = f"inline; filename*=UTF-8''{quote(download_name)}"
        self.send_bytes(data, ctype, headers)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    # ---------- GET
    def do_GET(self):
        url = urlparse(self.path)
        path = unquote(url.path)
        qs = parse_qs(url.query)

        try:
            if path in ("/", "/index.html"):
                return self.send_file(os.path.join(BASE_DIR, "static", "index.html"), "text/html; charset=utf-8")

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

            if path.startswith("/static/fonts/"):
                fname = os.path.basename(path)
                full = os.path.join(BASE_DIR, "static", "fonts", fname)
                ctype = "text/css; charset=utf-8" if fname.endswith(".css") else "font/woff2"
                return self.send_file(full, ctype)

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
                root = os.path.realpath(CFG["attachments_dir"])
                if not os.path.realpath(full).startswith(root + os.sep):
                    return self.send_error(HTTPStatus.FORBIDDEN)
                return self.send_file(full, download_name=fname)

            if path == "/health":
                return self.send_json(200, {"ok": True, "app": APP_NAME, "database": CFG["database"]})
        except sqlite3.Error as e:
            return self.send_json(500, {"ok": False, "error": f"Baza de date: {e}"})

        self.send_error(HTTPStatus.NOT_FOUND)

    # ---------- POST
    def do_POST(self):
        path = urlparse(self.path).path.rstrip("/") or "/"
        try:
            length = int(self.headers.get("Content-Length") or 0)
            data = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except Exception:  # noqa: BLE001
            return self.send_json(400, {"ok": False, "error": "JSON invalid"})

        user = check_pin(data.get("pin"))
        if path in ("/", "/edit", "/delete", "/upload") and not user:
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

            if path == "/upload":
                name = save_attachment(data.get("source"), data.get("no"),
                                       data.get("filename"), data.get("contentBase64"))
                print(f"  @ atasament '{name}' la #{data.get('no')} de {user}")
                return self.send_json(200, {"ok": True, "filename": name})
        except sqlite3.Error as e:
            return self.send_json(500, {"ok": False, "error": f"Baza de date: {e}"})
        except (RuntimeError, ValueError) as e:
            return self.send_json(500, {"ok": False, "error": str(e)})
        except Exception as e:  # noqa: BLE001
            return self.send_json(500, {"ok": False, "error": f"{type(e).__name__}: {e}"})

        self.send_json(404, {"ok": False, "error": "Rută necunoscută"})


# ---------------------------------------------------------------- import CSV
def import_pending_csv(csv_path):
    """Importa randurile din pending_new_actions.csv (coada salvata de browser cand serverul nu raspundea)."""
    keys = {norm(h): key for _, key, h in COLUMNS}
    keys.update({"familie": "familie", "urgent": "urgent"})
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample, delimiters=",;") if sample else csv.excel
        rows = list(csv.DictReader(f, dialect=dialect))
    count = 0
    for r in rows:
        data = {keys[norm(k)]: v for k, v in r.items() if k and norm(k) in keys and keys[norm(k)] != "no"}
        if not any(str(v).strip() for v in data.values()):
            continue
        no = add_action(data, "import CSV")
        count += 1
        print(f"  + #{no}: {data.get('client', '')} - {data.get('actiune', '')}")
    done = csv_path + f".importat_{datetime.now():%Y%m%d_%H%M%S}"
    os.replace(csv_path, done)
    print(f"Am importat {count} actiuni. Fisierul a fost redenumit in: {done}")


def main():
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--import-csv", metavar="FISIER",
                        help="importa un CSV cu actiuni (ex. pending_new_actions.csv) in baza de date si iese")
    args = parser.parse_args()
    init_db()
    if args.import_csv:
        import_pending_csv(args.import_csv)
        return
    os.makedirs(CFG["attachments_dir"], exist_ok=True)
    httpd = ThreadingHTTPServer((CFG["host"], int(CFG["port"])), Handler)
    print(f"{APP_NAME} ruleaza pe http://{CFG['host']}:{CFG['port']}/")
    print(f"  Baza de date: {CFG['database']}")
    print(f"  Atasamente:   {CFG['attachments_dir']}")
    if CFG["pins"] == DEFAULT_CONFIG["pins"]:
        print("  ATENTIE: folosesti codul implicit 1234. Seteaza codurile reale in config.json.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nOprit.")


if __name__ == "__main__":
    main()
