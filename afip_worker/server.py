# -*- coding: utf-8 -*-
"""API HTTP local: Streamlit Cloud encola acá. Token obligatorio."""
from __future__ import annotations

import base64
import hmac
import json
import logging
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from .auth import admin_mark_cuit_ready
from .jobs import (
    create_job,
    enqueue_job,
    list_jobs,
    uploads_dir,
)
from .registry import (
    badge_label,
    ensure_cuit_registered,
    list_cuits,
    mark_needs_admin,
)

LOG = logging.getLogger("afip_worker.api")
MAX_BODY = 8 * 1024 * 1024


def _json_bytes(payload: Any, status: int = 200) -> tuple[int, bytes]:
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    return status, raw


class WorkerApiHandler(BaseHTTPRequestHandler):
    server_version = "AfipWorker/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        LOG.info("%s - " + fmt, self.address_string(), *args)

    def _send(self, status: int, body: bytes, content_type: str = "application/json") -> None:
        self.send_response(status)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type, X-AFIP-Token")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def do_OPTIONS(self) -> None:  # noqa: N802
        self._send(204, b"")

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path.rstrip("/") or "/"
        if path == "/health":
            status, body = _json_bytes({"ok": True, "service": "afip_worker"})
            self._send(status, body)
            return
        if not self._authorized():
            self._send(*_json_bytes({"ok": False, "error": "unauthorized"}, 401))
            return
        if path == "/v1/status":
            self._send(*_json_bytes({"ok": True, "service": "afip_worker"}))
            return
        if path == "/v1/jobs":
            jobs = [j.to_dict() for j in list_jobs()]
            self._send(*_json_bytes({"ok": True, "jobs": jobs}))
            return
        if path == "/v1/cuits":
            rows = []
            for e in list_cuits():
                d = e.to_dict()
                d["acceso"] = badge_label(e.status)
                rows.append(d)
            self._send(*_json_bytes({"ok": True, "cuits": rows}))
            return
        self._send(*_json_bytes({"ok": False, "error": "not found"}, 404))

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path.rstrip("/") or "/"
        if not self._authorized():
            self._send(*_json_bytes({"ok": False, "error": "unauthorized"}, 401))
            return
        try:
            payload = self._read_json()
        except ValueError as exc:
            self._send(*_json_bytes({"ok": False, "error": str(exc)}, 400))
            return
        try:
            if path == "/v1/jobs":
                self._send(*self._post_job(payload))
                return
            if path == "/v1/cuits/ensure":
                entry = ensure_cuit_registered(
                    str(payload.get("cuit") or ""),
                    str(payload.get("razon_social") or ""),
                )
                d = entry.to_dict()
                d["acceso"] = badge_label(entry.status)
                self._send(*_json_bytes({"ok": True, "cuit": d}))
                return
            if path == "/v1/cuits/ready":
                admin_mark_cuit_ready(
                    str(payload.get("cuit") or ""),
                    str(payload.get("razon_social") or ""),
                )
                self._send(*_json_bytes({"ok": True}))
                return
            if path == "/v1/cuits/needs_admin":
                mark_needs_admin(
                    str(payload.get("cuit") or ""),
                    razon_social=str(payload.get("razon_social") or ""),
                )
                self._send(*_json_bytes({"ok": True}))
                return
        except ValueError as exc:
            self._send(*_json_bytes({"ok": False, "error": str(exc)}, 400))
            return
        except Exception as exc:
            LOG.exception("api error")
            self._send(*_json_bytes({"ok": False, "error": str(exc)}, 500))
            return
        self._send(*_json_bytes({"ok": False, "error": "not found"}, 404))

    def _authorized(self) -> bool:
        expected = getattr(self.server, "worker_token", "")  # type: ignore[attr-defined]
        if not expected:
            return False
        auth = self.headers.get("Authorization") or ""
        token = ""
        if auth.lower().startswith("bearer "):
            token = auth[7:].strip()
        if not token:
            token = (self.headers.get("X-AFIP-Token") or "").strip()
        return hmac.compare_digest(token, expected)

    def _read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            raise ValueError("cuerpo demasiado grande")
        raw = self.rfile.read(length) if length else b"{}"
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("JSON debe ser un objeto")
        return data

    def _post_job(self, payload: dict[str, Any]) -> tuple[int, bytes]:
        params = dict(payload.get("params") or {})
        b64 = str(payload.get("plantilla_b64") or "").strip()
        name = str(payload.get("plantilla_name") or "plantilla.xlsx")
        if b64:
            dest = uploads_dir() / Path(name).name
            dest.write_bytes(base64.b64decode(b64))
            params["plantilla_excel"] = str(dest)
        job = create_job(
            cuit=str(payload.get("cuit") or ""),
            razon_social=str(payload.get("razon_social") or ""),
            action=str(payload.get("action") or ""),  # type: ignore[arg-type]
            params=params,
            requested_by=str(payload.get("requested_by") or "cloud"),
        )
        ensure_cuit_registered(job.cuit, job.razon_social)
        path = enqueue_job(job)
        LOG.info("enqueued via api %s → %s", job.id, path.name)
        return _json_bytes({"ok": True, "job": job.to_dict(), "file": path.name})


def start_api_thread(*, host: str, port: int, token: str) -> ThreadingHTTPServer:
    httpd = ThreadingHTTPServer((host, port), WorkerApiHandler)
    httpd.worker_token = token  # type: ignore[attr-defined]
    thread = threading.Thread(target=httpd.serve_forever, name="afip-api", daemon=True)
    thread.start()
    LOG.info("API listening http://%s:%s", host, port)
    return httpd
