#!/usr/bin/env python3
"""Server pentru mai multe aplicații web, fiecare în folderul ei din `apps/`.

    apps/
      analiza-5c/        ->  http://<server>:8080/analiza-5c/
        index.html
        app.json         (opțional: titlu, descriere, iconiță)
      alta-aplicatie/    ->  http://<server>:8080/alta-aplicatie/

Pagina principală (/) listează automat toate aplicațiile. Un folder nou apare
imediat, fără repornirea serverului. Doar biblioteca standard Python (3.8+).

Pornire:  python server.py                 (port 8080, accesibil din rețea)
          python server.py --port 9000 --host 127.0.0.1
"""
import argparse
import html
import json
import mimetypes
import os
import re
import socket
import sys
import time
from datetime import datetime
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote, urlsplit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Tipuri de fișiere pe care browserul trebuie să le reverifice la fiecare cerere
# (date care se actualizează des, ex. Excelul citit de dashboard).
for ext, mime in {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".csv": "text/csv; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".mjs": "text/javascript",
    ".webmanifest": "application/manifest+json",
    ".svg": "image/svg+xml",
    ".woff2": "font/woff2",
}.items():
    mimetypes.add_type(mime, ext)

APP_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
TITLE_RE = re.compile(r"<title[^>]*>(.*?)</title>", re.I | re.S)


def is_hidden(name):
    return name.startswith((".", "_"))


def read_app_info(apps_dir, name):
    """Citește titlul/descrierea unei aplicații din app.json sau din <title>."""
    folder = os.path.join(apps_dir, name)
    info = {"name": name, "title": name, "description": "", "icon": "🧩", "url": f"/{quote(name)}/"}
    meta_path = os.path.join(folder, "app.json")
    if os.path.isfile(meta_path):
        try:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
            for key in ("title", "description", "icon"):
                if isinstance(meta.get(key), str) and meta[key].strip():
                    info[key] = meta[key].strip()
        except (OSError, ValueError) as e:
            print(f"[avertisment] {meta_path} invalid: {e}", file=sys.stderr)
    entry = find_entry_page(folder)
    if entry is None:
        info["missing_index"] = True
    elif info["title"] == name:
        try:
            with open(os.path.join(folder, entry), encoding="utf-8", errors="ignore") as f:
                m = TITLE_RE.search(f.read(20000))
            if m and m.group(1).strip():
                info["title"] = html.unescape(m.group(1).strip())
        except OSError:
            pass
    newest = max((os.path.getmtime(os.path.join(r, fn))
                  for r, ds, fs in os.walk(folder)
                  for fn in fs if not is_hidden(fn)), default=os.path.getmtime(folder))
    info["updated"] = datetime.fromtimestamp(newest).strftime("%d.%m.%Y %H:%M")
    return info


def find_entry_page(folder):
    """index.html, altfel singurul/primul fișier .html din folder."""
    if os.path.isfile(os.path.join(folder, "index.html")):
        return "index.html"
    pages = sorted(f for f in os.listdir(folder) if f.lower().endswith((".html", ".htm")) and not is_hidden(f))
    return pages[0] if pages else None


def list_apps(apps_dir):
    try:
        names = sorted(os.listdir(apps_dir), key=str.lower)
    except FileNotFoundError:
        return []
    return [read_app_info(apps_dir, n) for n in names
            if not is_hidden(n) and APP_NAME_RE.match(n) and os.path.isdir(os.path.join(apps_dir, n))]


INDEX_TEMPLATE = """<!doctype html>
<html lang="ro"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Aplicații</title>
<style>
:root{--bg:#f3f5f9;--panel:#fff;--line:#dde3ec;--text:#0b1220;--muted:#56627a;--accent:#0b8f83}
@media (prefers-color-scheme:dark){:root{--bg:#07101f;--panel:#0d1930;--line:#22375b;--text:#eef5ff;--muted:#91a5c5;--accent:#14d7c4}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:16px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
main{max-width:1100px;margin:0 auto;padding:32px 16px}
h1{margin:0 0 4px;font-size:28px}.sub{color:var(--muted);margin:0 0 28px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px}
a.card{display:flex;flex-direction:column;gap:6px;padding:20px;background:var(--panel);border:1px solid var(--line);border-radius:14px;color:inherit;text-decoration:none;transition:border-color .15s,transform .15s}
a.card:hover,a.card:focus-visible{border-color:var(--accent);transform:translateY(-2px);outline:none}
.icon{font-size:34px;line-height:1}.title{font-weight:650;font-size:18px}
.desc{color:var(--muted);font-size:14px;flex:1}.meta{color:var(--muted);font-size:12px;display:flex;justify-content:space-between;gap:8px}
.warn{color:#d03b3b;font-size:13px}.empty{color:var(--muted);padding:40px;text-align:center;border:1px dashed var(--line);border-radius:14px}
code{background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:1px 6px}
</style></head><body><main>
<h1>Aplicații</h1>
<p class="sub">__COUNT__ · actualizat automat din folderul <code>apps/</code></p>
__BODY__
</main></body></html>"""


def render_index(apps):
    if not apps:
        body = ('<div class="empty">Nicio aplicație încă. Creează un folder în <code>apps/</code> '
                'cu un fișier <code>index.html</code> și reîncarcă pagina.</div>')
    else:
        cards = []
        for a in apps:
            warn = ('<span class="warn">Lipsește index.html</span>' if a.get("missing_index") else "")
            cards.append(
                f'<a class="card" href="{a["url"]}"><span class="icon">{html.escape(a["icon"])}</span>'
                f'<span class="title">{html.escape(a["title"])}</span>'
                f'<span class="desc">{html.escape(a["description"])}</span>{warn}'
                f'<span class="meta"><span>/{html.escape(a["name"])}/</span><span>{a["updated"]}</span></span></a>')
        body = '<div class="grid">' + "".join(cards) + "</div>"
    n = len(apps)
    count = "o aplicație" if n == 1 else f"{n} aplicații"
    return INDEX_TEMPLATE.replace("__COUNT__", count).replace("__BODY__", body)


class AppsHandler(SimpleHTTPRequestHandler):
    server_version = "AppsServer/1.0"

    def __init__(self, *args, apps_dir, listing, **kwargs):
        self.apps_dir = apps_dir
        self.listing = listing
        super().__init__(*args, directory=apps_dir, **kwargs)

    # --- rutare -------------------------------------------------------------
    def do_GET(self):
        if self._route(head=False):
            super().do_GET()

    def do_HEAD(self):
        if self._route(head=True):
            super().do_HEAD()

    def _route(self, head):
        """Tratează căile speciale. Întoarce True dacă cererea e un fișier static."""
        path = unquote(urlsplit(self.path).path)
        parts = [p for p in path.split("/") if p]
        if not parts:
            return self._send_text(render_index(list_apps(self.apps_dir)), "text/html; charset=utf-8", head)
        if parts == ["api", "apps"]:
            return self._send_text(json.dumps(list_apps(self.apps_dir), ensure_ascii=False, indent=2),
                                   "application/json; charset=utf-8", head)
        if parts == ["favicon.ico"]:
            self.send_response(HTTPStatus.NO_CONTENT)
            self.end_headers()
            return False
        if any(is_hidden(p) or p == ".." for p in parts):
            return self._not_found(head)
        app_dir = os.path.join(self.apps_dir, parts[0])
        if not (APP_NAME_RE.match(parts[0]) and os.path.isdir(app_dir)):
            return self._not_found(head)
        if len(parts) == 1 and not path.endswith("/"):
            return self._redirect(f"/{quote(parts[0])}/")
        if path.endswith("/"):
            folder = os.path.join(self.apps_dir, *parts)
            if os.path.isdir(folder) and not os.path.isfile(os.path.join(folder, "index.html")):
                entry = find_entry_page(folder)
                if entry is None:
                    return self._not_found(head)
                return self._redirect(path + quote(entry))
        return True

    def list_directory(self, path):  # fără listarea conținutului folderelor
        self.send_error(HTTPStatus.NOT_FOUND, "Pagina nu există")
        return None

    # --- antete -------------------------------------------------------------
    def end_headers(self):
        # Browserul reverifică fișierele la fiecare cerere (304 dacă nu s-au schimbat),
        # astfel o aplicație sau un Excel actualizat se vede imediat.
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    def guess_type(self, path):
        t = super().guess_type(path)
        if t in ("text/html", "text/css", "text/javascript", "application/javascript", "text/plain"):
            t += "; charset=utf-8"
        return t

    # --- utilitare ----------------------------------------------------------
    def _send_text(self, text, ctype, head):
        data = text.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if not head:
            self.wfile.write(data)
        return False

    def _redirect(self, location):
        self.send_response(HTTPStatus.MOVED_PERMANENTLY)
        self.send_header("Location", location)
        self.send_header("Content-Length", "0")
        self.end_headers()
        return False

    def _not_found(self, head):
        body = ('<!doctype html><meta charset="utf-8"><title>Negăsit</title>'
                '<body style="font-family:system-ui;padding:40px">'
                '<h1>404 – Pagina nu există</h1><p><a href="/">Înapoi la lista de aplicații</a></p>')
        data = body.encode("utf-8")
        self.send_response(HTTPStatus.NOT_FOUND)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if not head:
            self.wfile.write(data)
        return False

    def log_message(self, fmt, *args):
        if self.listing.quiet:
            return
        sys.stderr.write(f"{time.strftime('%H:%M:%S')} {self.address_string()} {fmt % args}\n")


def local_addresses():
    addrs = set()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        addrs.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            addrs.add(info[4][0])
    except OSError:
        pass
    return sorted(a for a in addrs if not a.startswith("127."))


def main():
    p = argparse.ArgumentParser(description="Server pentru aplicațiile web din folderul apps/.")
    p.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8080)), help="portul (implicit 8080)")
    p.add_argument("--host", default=os.environ.get("HOST", "0.0.0.0"),
                   help="adresa de ascultare (implicit 0.0.0.0 = accesibil din rețea; 127.0.0.1 = doar local)")
    p.add_argument("--apps", default=os.path.join(BASE_DIR, "apps"), help="folderul cu aplicații")
    p.add_argument("--quiet", action="store_true", help="fără jurnal pentru fiecare cerere")
    args = p.parse_args()

    apps_dir = os.path.abspath(args.apps)
    os.makedirs(apps_dir, exist_ok=True)
    handler = partial(AppsHandler, apps_dir=apps_dir, listing=args)
    try:
        httpd = ThreadingHTTPServer((args.host, args.port), handler)
    except OSError as e:
        sys.exit(f"Nu pot porni pe {args.host}:{args.port}: {e}\nÎncearcă alt port: python server.py --port 9000")
    httpd.daemon_threads = True

    apps = list_apps(apps_dir)
    print(f"Server pornit. Folder aplicații: {apps_dir}")
    print(f"  Local:   http://localhost:{args.port}/")
    if args.host in ("0.0.0.0", ""):
        for a in local_addresses():
            print(f"  Rețea:   http://{a}:{args.port}/")
    for a in apps:
        print(f"    • {a['title']}  ->  {a['url']}")
    print("Oprire: Ctrl+C")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer oprit.")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
