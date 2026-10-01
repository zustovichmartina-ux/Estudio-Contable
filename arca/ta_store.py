"""Ticket WSAA durable: la base del estudio manda; el archivo es solo la caché de obtener_ta."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os

from arca import wsaa
from arca.persistencia import guardar_ta, leer_ta


def hash_cert(cert_path: str) -> str:
    with open(cert_path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()[:12]


def _volcar_si_vigente(cache_dir: str, env: str, servicio: str, cert_path: str, margen_min: int) -> None:
    guardado = leer_ta(env, servicio, hash_cert(cert_path))
    if not guardado or not guardado.get("expiration"):
        return
    try:
        exp = wsaa._parse_fecha(guardado["expiration"])
    except (ValueError, TypeError):
        return
    ahora = dt.datetime.now(wsaa.TZ)
    if exp - dt.timedelta(minutes=margen_min) <= ahora:
        return
    cf = wsaa._cache_file(cache_dir, env, servicio, cert_path)
    if os.path.exists(cf):
        return
    payload = {k: guardado.get(k) for k in ("token", "sign", "expiration", "generation", "destination") if guardado.get(k)}
    fd = os.open(cf, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(payload, f)
    os.chmod(cf, 0o600)


def obtener_ta_durable(env, cert_path, key_path, servicio="wsfe", cache_dir="cache", margen_min=10):
    """Igual que wsaa.obtener_ta, pero reutiliza el TA guardado en la base (~12 h)."""
    os.makedirs(cache_dir, mode=0o700, exist_ok=True)
    _volcar_si_vigente(cache_dir, env, servicio, cert_path, margen_min)
    ta = wsaa.obtener_ta(env, cert_path, key_path, servicio, cache_dir, margen_min)
    if ta.get("source") == "wsaa":
        guardar_ta(env, servicio, hash_cert(cert_path), ta)
    return ta
