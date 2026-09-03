# -*- coding: utf-8 -*-
"""Smoke test API local (sin túnel)."""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
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
    try:
        health = urllib.request.urlopen("http://127.0.0.1:18765/health", timeout=5)
        assert health.status == 200
        req = urllib.request.Request(
            "http://127.0.0.1:18765/v1/jobs",
            data=json.dumps(
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
            ).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        assert body.get("ok") is True
        assert body["job"]["status"] == "pending"
        print("PASS api enqueue", body["job"]["id"])
        return 0
    finally:
        httpd.shutdown()


if __name__ == "__main__":
    raise SystemExit(main())
