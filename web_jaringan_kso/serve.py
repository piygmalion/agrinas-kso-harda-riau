# -*- coding: utf-8 -*-
"""Server lokal untuk webpage interaktif jaringan KSO Agrinas.

Usage:
  python serve.py
  lalu buka http://localhost:8765
"""

from __future__ import annotations

import functools
import http.server
import os
import socketserver
import webbrowser
from pathlib import Path

DIR = Path(__file__).resolve().parent
PORT = 8765


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DIR), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, fmt, *args):
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))


def main():
    # Pastikan graph-data.json mutakhir dari GEXF
    build = DIR / "build_from_gexf.py"
    if build.exists():
        os.system(f'python "{build}"')

    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
        url = f"http://127.0.0.1:{PORT}/"
        print("=" * 60)
        print("Peta Jaringan KSO Agrinas — Unit II Harda")
        print(f"Serving: {DIR}")
        print(f"URL    : {url}")
        print("Tekan Ctrl+C untuk berhenti.")
        print("=" * 60)
        try:
            webbrowser.open(url)
        except Exception:
            pass
        httpd.serve_forever()


if __name__ == "__main__":
    main()
