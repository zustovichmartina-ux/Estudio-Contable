# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))

from motor import parsear_pdf

BASE_PDF = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\SUNNY BEACH SA SAMSARA"
    r"\Balances\Balances\Balance al 30-04-2026\Banco Galicia"
)
MUESTRAS = {
    "042026": BASE_PDF / "BANCO GALICIA 042026.pdf",
    "032026": BASE_PDF / "BANCO GALICIA 032026.pdf",
    "062025": BASE_PDF / "BANCO GALICIA 062025.pdf",
    "082025": BASE_PDF / "BANCO GALICIA 082025.pdf",
}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(HERE), **kwargs)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _json(self, payload: dict, code: int = 200) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/parse-sample":
            qs = parse_qs(parsed.query)
            clave = (qs.get("id") or ["062025"])[0]
            path = MUESTRAS.get(clave)
            if not path or not path.exists():
                self._json({"error": f"No esta {clave}"}, 404)
                return
            print(f"sample {path.name} ocr-if-needed...", flush=True)
            out = parsear_pdf(path.read_bytes(), path.name)
            self._json(out)
            return
        super().do_GET()

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/parse":
            self.send_error(404)
            return
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length)
        nombre = self.headers.get("X-Filename") or "extracto.pdf"
        ctype = self.headers.get("Content-Type") or ""
        data = body
        if "multipart/form-data" in ctype:
            data, nombre = _extraer_archivo(body, ctype, nombre)
        if not data:
            self._json({"error": "PDF vacio"}, 400)
            return
        print(f"parse {nombre} ({len(data)} bytes)", flush=True)
        out = parsear_pdf(data, nombre)
        self._json(out)


def _extraer_archivo(body: bytes, ctype: str, fallback: str) -> tuple[bytes, str]:
    marker = b"filename=\""
    i = body.find(marker)
    nombre = fallback
    if i >= 0:
        j = body.find(b"\"", i + len(marker))
        if j > i:
            nombre = body[i + len(marker) : j].decode("utf-8", errors="replace")
    sep = b"\r\n\r\n"
    k = body.find(sep)
    if k < 0:
        return body, nombre
    start = k + 4
    end = body.rfind(b"\r\n--")
    if end <= start:
        return body[start:], nombre
    return body[start:end], nombre


def main() -> None:
    port = 8765
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Demo OCR en http://127.0.0.1:{port}/", flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
