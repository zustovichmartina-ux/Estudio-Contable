# -*- coding: utf-8 -*-
"""Cliente HTTP para Streamlit Cloud → worker local (túnel)."""
from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import Any


class RemoteError(RuntimeError):
    def __init__(self, message: str, status: int = 0) -> None:
        super().__init__(message)
        self.status = status


class RemoteWorker:
    def __init__(self, base_url: str, token: str, timeout: float = 25.0) -> None:
        self.base_url = (base_url or "").rstrip("/")
        self.token = token
        self.timeout = timeout

    def _request(self, method: str, path: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        url = f"{self.base_url}{path}"
        data = None
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/json",
        }
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                raw = resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise RemoteError(f"HTTP {exc.code}: {body or exc.reason}", exc.code) from exc
        except urllib.error.URLError as exc:
            raise RemoteError(
                "No se llega al worker de RECEPCION. "
                "¿Está abierto iniciar_afip_worker.bat y el túnel? "
                f"({exc.reason})"
            ) from exc
        if not raw:
            return {}
        parsed = json.loads(raw)
        if not isinstance(parsed, dict):
            raise RemoteError("respuesta inválida del worker")
        if parsed.get("ok") is False:
            raise RemoteError(str(parsed.get("error") or "error del worker"))
        return parsed

    def health(self) -> bool:
        url = f"{self.base_url}/health"
        req = urllib.request.Request(url, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=8) as resp:
                return resp.status == 200
        except Exception:
            return False

    def list_jobs(self) -> list[dict[str, Any]]:
        return list(self._request("GET", "/v1/jobs").get("jobs") or [])

    def list_cuits(self) -> list[dict[str, Any]]:
        return list(self._request("GET", "/v1/cuits").get("cuits") or [])

    def ensure_cuit(self, cuit: str, razon_social: str = "") -> dict[str, Any]:
        data = self._request(
            "POST",
            "/v1/cuits/ensure",
            {"cuit": cuit, "razon_social": razon_social},
        )
        return dict(data.get("cuit") or {})

    def mark_ready(self, cuit: str, razon_social: str = "") -> None:
        self._request("POST", "/v1/cuits/ready", {"cuit": cuit, "razon_social": razon_social})

    def mark_needs_admin(self, cuit: str, razon_social: str = "") -> None:
        self._request(
            "POST",
            "/v1/cuits/needs_admin",
            {"cuit": cuit, "razon_social": razon_social},
        )

    def enqueue(
        self,
        *,
        cuit: str,
        razon_social: str,
        action: str,
        params: dict[str, Any],
        requested_by: str = "cloud",
        plantilla_b64: str = "",
        plantilla_name: str = "",
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "cuit": cuit,
            "razon_social": razon_social,
            "action": action,
            "params": params,
            "requested_by": requested_by,
        }
        if plantilla_b64:
            payload["plantilla_b64"] = plantilla_b64
            payload["plantilla_name"] = plantilla_name or "plantilla.xlsx"
        return dict(self._request("POST", "/v1/jobs", payload).get("job") or {})
