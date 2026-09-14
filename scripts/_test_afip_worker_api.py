# -*- coding: utf-8 -*-
"""Smoke test API local (sin túnel)."""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from afip_worker.jobs import ensure_job_dirs
from afip_worker.registry import mark_ready
from afip_worker.server import start_api_thread


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="afip_api_"))
    ensure_job_dirs(tmp)
    os.environ["AFIP_JOBS_ROOT"] = str(tmp)
    os.environ["AFIP_CUIT_REGISTRY"] = str(tmp / "cuit_registry.json")
    mark_ready("27-42043034-0", razon_social="Camila Rocio Albarello", note="test")
    token = "test-token-32chars______________"
    httpd = start_api_thread(host="127.0.0.1", port=18765, token=token)
    time.sleep(0.3)

    def _post(payload: dict) -> tuple[int, dict]:
        req = urllib.request.Request(
            "http://127.0.0.1:18765/v1/jobs",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=5) as resp:
                return resp.status, json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="replace")
            try:
                parsed = json.loads(raw) if raw else {}
            except json.JSONDecodeError:
                parsed = {"error": raw}
            return exc.code, parsed

    try:
        health = urllib.request.urlopen("http://127.0.0.1:18765/health", timeout=5)
        assert health.status == 200
        status, body = _post(
            {
                "cuit": "27-42043034-0",
                "razon_social": "Camila Rocio Albarello",
                "action": "bajar_comprobantes",
                "params": {
                    "ruta_destino": str(tmp / "out"),
                    "periodo_desde": "2026-01-01",
                    "periodo_hasta": "2026-01-31",
                },
                "requested_by": "test",
            }
        )
        assert status == 200
        assert body.get("ok") is True
        assert body["job"]["status"] == "pending"
        print("PASS api enqueue", body["job"]["id"])

        status, body = _post(
            {
                "cuit": "27-42043034-0",
                "razon_social": "Camila Rocio Albarello",
                "action": "bajar_portal_iva",
                "params": {
                    "ruta_destino": str(tmp / "portal"),
                    "periodo_desde": "2026-08-01",
                    "periodo_hasta": "2026-08-31",
                },
                "requested_by": "test",
            }
        )
        assert status == 200, body
        assert body["job"]["action"] == "bajar_portal_iva"
        print("PASS api enqueue portal_iva", body["job"]["id"])

        status, body = _post(
            {
                "cuit": "27-42043034-0",
                "razon_social": "Camila Rocio Albarello",
                "action": "no_existe",
                "params": {"ruta_destino": str(tmp / "x")},
                "requested_by": "test",
            }
        )
        assert status == 400
        assert body.get("ok") is False
        print("PASS api rechaza action inválida")
        return 0
    finally:
        httpd.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
