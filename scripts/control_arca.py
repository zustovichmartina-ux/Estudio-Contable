# -*- coding: utf-8 -*-
"""Control ARCA desde tu PC: mira si el ejecutor está vivo. No abre AFIP."""
from __future__ import annotations

import json
import socket
import sys
import urllib.request
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from afip_worker.tunnel import DISCOVERY_URLS


def _get_json(url: str, timeout: float = 8.0) -> dict:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "EstudioContable-ARCA/1.0", "Accept": "application/json,text/plain"},
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return {"raw": raw.strip()}
    return parsed if isinstance(parsed, dict) else {"raw": raw.strip()}


def _health(base: str) -> dict:
    return _get_json(base.rstrip("/") + "/health", timeout=5.0)


def main() -> int:
    yo = socket.gethostname()
    print(f"PC de control: {yo}")
    print("Web: https://estudiocontablemdp.streamlit.app/  ->  ARCA")
    print()

    local_ok = False
    try:
        info = _health("http://127.0.0.1:8765")
        if info.get("ok"):
            local_ok = True
            host = str(info.get("host") or yo)
            print(f"Ejecutor local: SI  (esta misma PC se llama {host})")
            print("Si queres que EJECUTE OTRA maquina, cerra ejecutor_arca.bat aca.")
    except Exception:
        print("Ejecutor en esta PC: no (bien, si la otra maquina es la que ejecuta).")

    print()
    remoto = None
    for src in DISCOVERY_URLS:
        try:
            listing = urllib.request.Request(
                src,
                headers={"User-Agent": "EstudioContable-ARCA/1.0", "Cache-Control": "no-cache"},
            )
            with urllib.request.urlopen(listing, timeout=8) as resp:
                url = resp.read().decode("utf-8", errors="replace").strip().split()[0]
        except Exception:
            continue
        if not url.startswith("https://"):
            continue
        try:
            info = _health(url)
        except Exception:
            continue
        if info.get("ok"):
            remoto = (url, str(info.get("host") or "?"))
            break

    if remoto:
        _u, host = remoto
        print("Ejecutor remoto: SI  - maquina {host}".format(host=host))
        print("Desde esta PC encolas. La otra ejecuta.")
        return 0

    if local_ok:
        print("La web puede usar esta PC como ejecutor mientras este prendida.")
        return 0

    print("Ejecutor: NO se llega.")
    print("En la maquina que ejecuta: deja abierto ejecutor_arca.bat")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
