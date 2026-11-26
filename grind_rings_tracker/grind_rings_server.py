#!/usr/bin/env python3
"""
Grind Rings Action Tracker - server Python (port 8092).

Inlocuieste combinatia "server web + Start-ActionListener.ps1" a trackerului
Past Due: un singur proces care
  * serveste dashboard-ul (static/index.html)          GET  /
  * serveste fisierul Excel cu actiunile               GET  /data/tracker
  * serveste ticker-ul (Dashboard2026.xlsm) prin proxy GET  /data/ticker
  * adauga o actiune noua in Excel                     POST /
  * editeaza o actiune                                 POST /edit
  * sterge o actiune                                   POST /delete
  * ataseaza un fisier la o actiune                    POST /upload
  * listeaza fisierele atasate                         GET  /list?no=..&source=..
  * serveste fisierele atasate                         GET  /attachments/<source>/<no>/<fisier>

Pornire:  python grind_rings_server.py            (necesita: pip install openpyxl)
Setari:   config.json langa acest script (vezi config.example.json).
"""

import base64
import json
import mimetypes
import os
import re
import shutil
import sys
import tempfile
import threading
import unicodedata
import urllib.request
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

try:
    import openpyxl
except ImportError:
    sys.exit("Lipseste openpyxl. Instaleaza-l cu:  pip install openpyxl")

APP_NAME = "Grind Rings Action Tracker"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG = {
    "host": "0.0.0.0",
    "port": 8092,
    # Fisierul Excel cu actiunile (.xlsx sau .xlsm - macro-urile se pastreaza).
    "excel_file": "Grind_Rings_Action_Tracker.xlsx",
    "sheet_name": "ActionItems",
    "lists_sheet_name": "Lists",
    # Folderul unde se salveaza atasamentele: <folder>/<source>/<no>/<fisier>
    "attachments_dir": "attachments",
    # Fisierul ticker-ului: URL (http://...) sau cale locala / UNC. Gol = fara ticker.
    "ticker_source": "http://romhdata03/timkentraining/Dashboard2026.xlsm",
    # Codurile personale: "cod": "Nume". Numele apare in comentarii.
    "pins": {"1234": "Admin"},
    "max_upload_mb": 25,
}

FIELD_KEYWORDS = {
    "no": ["no", "nr", "nr crt", "numar"],
    "dataAdaugare": ["data adaugare", "data creare", "added"],
    "sapRulment": ["sap id rulment", "sap rulment"],
    "descRulment": ["descriere rulment"],
    "familie": ["familie rulment", "familie"],
    "comandaMontaj": ["comanda montaj"],
    "client": ["client"],
    "sapComponent": ["sap id component", "sap component"],
    "descComponent": ["descriere component"],
    "tipComponent": ["tip component"],
    "comandaComponent": ["comanda component"],
    "qty": ["qty", "cantitate"],
    "tipActiune": ["tip actiune"],
    "actiune": ["actiune"],
    "responsabil": ["responsabil"],
    "dataPlanificata": ["data planificata"],
    "dataActuala": ["data finalizare", "data actuala"],
    "stadiu": ["stadiu", "status"],
    "comentarii": ["comentarii", "observatii", "note"],
    "urgent": ["urgent client", "urgent"],
}

DEFAULT_HEADERS = [
    ("no", "No"), ("dataAdaugare", "Data adaugare"), ("sapRulment", "SAP id Rulment"),
    ("descRulment", "Descriere Rulment"), ("familie", "Familie Rulment"),
    ("comandaMontaj", "Comanda Montaj"), ("client", "Client"),
    ("sapComponent", "SAP id Component"), ("descComponent", "Descriere Component"),
    ("tipComponent", "Tip component"), ("comandaComponent", "Comanda Component"),
    ("qty", "Qty"), ("actiune", "Actiune"), ("tipActiune", "Tip Actiune"),
    ("responsabil", "Responsabil"), ("dataPlanificata", "Data Planificata"),
    ("dataActuala", "Data Finalizare"), ("stadiu", "Stadiu"),
    ("urgent", "Urgent Client"), ("comentarii", "Comentarii"),
]
LIST_HEADERS = ["Client", "Responsabil", "Tip Actiune", "Familie Rulment", "Tip component"]

TEXT_FIELDS = ["sapRulment", "descRulment", "familie", "comandaMontaj", "client",
               "sapComponent", "descComponent", "tipComponent", "comandaComponent",
               "qty", "actiune", "tipActiune", "responsabil"]
DATE_FORMAT = "dd.mm.yyyy"

excel_lock = threading.Lock()


# ---------------------------------------------------------------- config
def load_config():
    cfg = dict(DEFAULT_CONFIG)
    path = os.path.join(BASE_DIR, "config.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            cfg.update(json.load(f))
    for key in ("excel_file", "attachments_dir"):
        if not os.path.isabs(cfg[key]):
            cfg[key] = os.path.join(BASE_DIR, cfg[key])
    cfg["pins"] = {str(k): str(v) for k, v in cfg.get("pins", {}).items()}
    return cfg


CFG = load_config()


# ---------------------------------------------------------------- helpers
def norm(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


def detect_columns(header_cells):
    """Acelasi algoritm ca in dashboard: intai potrivire exacta, apoi partiala."""
    headers = [norm(h) for h in header_cells]
    colmap = {}
    for exact in (True, False):
        for field, kws in FIELD_KEYWORDS.items():
            if field in colmap:
                continue
            for kw in kws:
                idx = next((i for i, h in enumerate(headers)
                            if h and (h == kw if exact else kw in h)
                            and (i + 1) not in colmap.values()), None)
                if idx is not None:
                    colmap[field] = idx + 1  # coloana 1-based
                    break
    return colmap


def parse_date(s):
    s = str(s or "").strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
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


# ---------------------------------------------------------------- Excel
def ensure_workbook():
    path = CFG["excel_file"]
    if os.path.exists(path):
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = CFG["sheet_name"]
    ws.append([h for _, h in DEFAULT_HEADERS])
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "A2"
    lists = wb.create_sheet(CFG["lists_sheet_name"])
    lists.append(LIST_HEADERS)
    wb.save(path)
    print(f"Am creat fisierul Excel nou: {path}")


def open_sheet():
    path = CFG["excel_file"]
    wb = openpyxl.load_workbook(path, keep_vba=path.lower().endswith(".xlsm"))
    ws = wb[CFG["sheet_name"]] if CFG["sheet_name"] in wb.sheetnames else wb.worksheets[0]
    header = [c.value for c in ws[1]]
    colmap = detect_columns(header)
    # adauga coloanele lipsa la final, ca sa nu se piarda date
    for field, title in DEFAULT_HEADERS:
        if field not in colmap:
            col = ws.max_column + 1
            ws.cell(row=1, column=col, value=title)
            colmap[field] = col
    return wb, ws, colmap


def save_workbook(wb):
    path = CFG["excel_file"]
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1], dir=os.path.dirname(path))
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except PermissionError:
        os.remove(tmp)
        raise RuntimeError("Fișierul Excel este deschis/blocat de alt program. Închide-l și încearcă din nou.")
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def last_data_row(ws, colmap):
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in colmap.values()):
            return r
    return 1


def find_row(ws, colmap, no):
    key = str(no).strip()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=colmap["no"]).value
        if v is None:
            continue
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        if str(v).strip() == key:
            return r
    return None


def set_cell(ws, row, col, value):
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, (datetime, date)):
        cell.number_format = DATE_FORMAT


def write_fields(ws, row, colmap, data):
    for f in TEXT_FIELDS:
        if f in data:
            val = str(data.get(f) or "").strip()
            if f == "qty" and re.fullmatch(r"-?\d+(\.\d+)?", val):
                val = float(val) if "." in val else int(val)
            set_cell(ws, row, colmap[f], val)
    if "dataPlanificata" in data:
        set_cell(ws, row, colmap["dataPlanificata"], parse_date(data["dataPlanificata"]))
    if "urgent" in data:
        set_cell(ws, row, colmap["urgent"], "Da" if is_urgent(data["urgent"]) else "Nu")


def add_action(data, user):
    with excel_lock:
        wb, ws, colmap = open_sheet()
        row = last_data_row(ws, colmap) + 1
        nums = []
        for r in range(2, row):
            v = ws.cell(row=r, column=colmap["no"]).value
            try:
                nums.append(int(float(v)))
            except (TypeError, ValueError):
                pass
        no = max(nums, default=0) + 1
        set_cell(ws, row, colmap["no"], no)
        set_cell(ws, row, colmap["dataAdaugare"], parse_date(data.get("dataAdaugare")) or datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))
        write_fields(ws, row, colmap, data)
        set_cell(ws, row, colmap["stadiu"], data.get("stadiu") or "Deschis")
        set_cell(ws, row, colmap["comentarii"], stamp_comment(data.get("comentarii"), user))
        save_workbook(wb)
        return no


def edit_action(data, user):
    with excel_lock:
        wb, ws, colmap = open_sheet()
        row = find_row(ws, colmap, data.get("no"))
        if row is None:
            return False
        write_fields(ws, row, colmap, data)
        stadiu = data.get("stadiu")
        if stadiu:
            old = norm(ws.cell(row=row, column=colmap["stadiu"]).value)
            set_cell(ws, row, colmap["stadiu"], stadiu)
            fin = ws.cell(row=row, column=colmap["dataActuala"])
            if stadiu == "Rezolvat" and fin.value in (None, ""):
                set_cell(ws, row, colmap["dataActuala"], datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))
            elif stadiu != "Rezolvat" and "rezolv" in old:
                fin.value = None
        new_comment = stamp_comment(data.get("comentarii"), user)
        if new_comment:
            cell = ws.cell(row=row, column=colmap["comentarii"])
            old = str(cell.value or "").strip()
            cell.value = f"{old}\n{new_comment}" if old else new_comment
        save_workbook(wb)
        return True


def delete_action(no):
    with excel_lock:
        wb, ws, colmap = open_sheet()
        row = find_row(ws, colmap, no)
        if row is None:
            return False
        ws.delete_rows(row, 1)
        save_workbook(wb)
        return True


# ---------------------------------------------------------------- attachments
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
    server_version = "GrindRingsTracker/1.0"

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
            headers["Content-Disposition"] = f"inline; filename*=UTF-8''{urllib.request.quote(download_name)}"
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

        if path in ("/", "/index.html"):
            return self.send_file(os.path.join(BASE_DIR, "static", "index.html"), "text/html; charset=utf-8")

        if path == "/data/tracker":
            with excel_lock:
                return self.send_file(CFG["excel_file"], "application/octet-stream")

        if path == "/data/ticker":
            return self.serve_ticker()

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

        if path == "/xlsx.full.min.js":
            return self.send_file(os.path.join(BASE_DIR, "static", "xlsx.full.min.js"), "application/javascript")

        if path == "/health":
            return self.send_json(200, {"ok": True, "app": APP_NAME, "excel": CFG["excel_file"]})

        self.send_error(HTTPStatus.NOT_FOUND)

    def serve_ticker(self):
        src = (CFG.get("ticker_source") or "").strip()
        if not src:
            return self.send_error(HTTPStatus.NOT_FOUND, "Ticker neconfigurat")
        try:
            if re.match(r"^https?://", src, re.I):
                with urllib.request.urlopen(src, timeout=10) as r:
                    data = r.read()
            else:
                with open(src, "rb") as f:
                    data = f.read()
        except Exception as e:  # noqa: BLE001
            return self.send_error(HTTPStatus.BAD_GATEWAY, f"Ticker indisponibil: {e}")
        self.send_bytes(data, "application/octet-stream")

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
                if not delete_action(data.get("no")):
                    return self.send_json(404, {"ok": False, "error": "Acțiunea nu a fost găsită"})
                print(f"  - actiune #{data.get('no')} stearsa de {user}")
                return self.send_json(200, {"ok": True})

            if path == "/upload":
                name = save_attachment(data.get("source"), data.get("no"),
                                       data.get("filename"), data.get("contentBase64"))
                print(f"  @ atasament '{name}' la #{data.get('no')} de {user}")
                return self.send_json(200, {"ok": True, "filename": name})
        except (RuntimeError, ValueError) as e:
            return self.send_json(500, {"ok": False, "error": str(e)})
        except Exception as e:  # noqa: BLE001
            return self.send_json(500, {"ok": False, "error": f"{type(e).__name__}: {e}"})

        self.send_json(404, {"ok": False, "error": "Rută necunoscută"})


def main():
    ensure_workbook()
    os.makedirs(CFG["attachments_dir"], exist_ok=True)
    httpd = ThreadingHTTPServer((CFG["host"], int(CFG["port"])), Handler)
    print(f"{APP_NAME} ruleaza pe http://{CFG['host']}:{CFG['port']}/")
    print(f"  Excel:       {CFG['excel_file']}")
    print(f"  Atasamente:  {CFG['attachments_dir']}")
    print(f"  Ticker:      {CFG.get('ticker_source') or '(dezactivat)'}")
    if CFG["pins"] == DEFAULT_CONFIG["pins"]:
        print("  ATENTIE: folosesti codul implicit 1234. Seteaza codurile reale in config.json.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nOprit.")


if __name__ == "__main__":
    main()
