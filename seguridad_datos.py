"""Cifrado at-rest (Fernet) para archivos sensibles en Cloud / disco local.

Clave: st.secrets["DATA_ENCRYPTION_KEY"] (o variable de entorno homónima).
Generar: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
"""

from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Optional

BASE_DIR = Path(__file__).resolve().parent
SECURE_ROOT = BASE_DIR / "data" / "secure"
ENC_SUFFIX = ".enc"


def _leer_secreto(clave: str) -> str | None:
    """Lee un secret de Streamlit o del entorno (sin fallar fuera de Streamlit)."""
    env = str(os.environ.get(clave) or "").strip()
    if env:
        return env
    try:
        import streamlit as st

        secrets = getattr(st, "secrets", None)
        if secrets is None:
            return None
        val = secrets.get(clave)
        if val is None:
            return None
        return str(val).strip() or None
    except Exception:
        return None


def tiene_clave_cifrado() -> bool:
    return bool(_leer_secreto("DATA_ENCRYPTION_KEY"))


def obtener_fernet():
    """Devuelve Fernet o None si no hay clave configurada."""
    raw = _leer_secreto("DATA_ENCRYPTION_KEY")
    if not raw:
        return None
    try:
        from cryptography.fernet import Fernet

        key = raw.encode("utf-8") if isinstance(raw, str) else raw
        return Fernet(key)
    except Exception:
        return None


def instrucciones_clave_cifrado() -> str:
    return (
        "Falta `DATA_ENCRYPTION_KEY` en Secrets. Generá una con:\n\n"
        "```python\n"
        "from cryptography.fernet import Fernet\n"
        "print(Fernet.generate_key().decode())\n"
        "```\n\n"
        "Pegala en Manage app → Settings → Secrets como:\n\n"
        'DATA_ENCRYPTION_KEY = "PEGAR_CLAVE_AQUI"'
    )


def slug_usuario(usuario: str | None) -> str:
    slug = re.sub(r"[^a-z0-9]+", "_", str(usuario or "").strip().lower()).strip("_")
    return (slug or "_anon")[:40]


def directorio_seguro(usuario: str | None, *subdirs: str) -> Path:
    """Carpeta escribible namespaced por usuario: data/secure/<user>/..."""
    ruta = SECURE_ROOT.joinpath(slug_usuario(usuario), *[str(s) for s in subdirs if s])
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def cifrar_bytes(data: bytes) -> bytes:
    f = obtener_fernet()
    if f is None:
        raise RuntimeError("DATA_ENCRYPTION_KEY no configurada o inválida.")
    return f.encrypt(bytes(data))


def descifrar_bytes(data: bytes) -> bytes:
    f = obtener_fernet()
    if f is None:
        raise RuntimeError("DATA_ENCRYPTION_KEY no configurada o inválida.")
    return f.decrypt(bytes(data))


def escribir_cifrado(ruta: Path, data: bytes) -> Path:
    """Escribe bytes cifrados. Si la ruta no termina en .enc, se agrega."""
    destino = Path(ruta)
    if destino.suffix.lower() != ".enc" and not str(destino).endswith(ENC_SUFFIX):
        destino = Path(str(destino) + ENC_SUFFIX)
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(cifrar_bytes(data))
    return destino


def leer_cifrado(ruta: Path) -> bytes:
    """Lee y descifra. Si el archivo no parece cifrado, devuelve plaintext (compat)."""
    bruto = Path(ruta).read_bytes()
    if str(ruta).endswith(ENC_SUFFIX) or bruto.startswith(b"gAAAAA"):
        return descifrar_bytes(bruto)
    return bruto


def materializar_descifrado(ruta: Path, *, suffix: str | None = None) -> Path:
    """Descifra a un temporal en disco (para APIs que exigen Path)."""
    plain = leer_cifrado(ruta)
    suf = suffix
    if suf is None:
        name = Path(ruta).name
        if name.endswith(ENC_SUFFIX):
            name = name[: -len(ENC_SUFFIX)]
        suf = Path(name).suffix or ".bin"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suf)
    tmp.write(plain)
    tmp.close()
    return Path(tmp.name)


def guardar_plan_cifrado(usuario: str | None, cuit: str, archivo_bytes: bytes) -> Path:
    """Persiste plan de cuentas cifrado bajo el namespace del usuario."""
    dest = directorio_seguro(usuario, "planes") / f"plan_{cuit}.xlsx{ENC_SUFFIX}"
    return escribir_cifrado(dest, archivo_bytes)


def ruta_plan_cifrado(usuario: str | None, cuit: str) -> Path:
    return directorio_seguro(usuario, "planes") / f"plan_{cuit}.xlsx{ENC_SUFFIX}"


def guardar_plan_cifrado_por_cliente(
    usuario: str | None,
    cliente_id: int | str,
    archivo_bytes: bytes,
) -> Path:
    """Plan cifrado asociado al id de cliente (CUIT placeholder / temporal no importa)."""
    dest = (
        directorio_seguro(usuario, "planes")
        / f"plan_id_{int(cliente_id)}.xlsx{ENC_SUFFIX}"
    )
    return escribir_cifrado(dest, archivo_bytes)


def ruta_plan_cifrado_por_cliente(usuario: str | None, cliente_id: int | str) -> Path:
    return (
        directorio_seguro(usuario, "planes")
        / f"plan_id_{int(cliente_id)}.xlsx{ENC_SUFFIX}"
    )

def guardar_balance_cifrado(
    usuario: str | None,
    sociedad_id: int | str,
    nombre_original: str,
    archivo_bytes: bytes,
) -> Path:
    safe = re.sub(r"[^a-zA-Z0-9._-]+", "_", Path(nombre_original).name)[:80] or "balance.xlsx"
    dest = directorio_seguro(usuario, "balances") / f"soc_{sociedad_id}_{safe}{ENC_SUFFIX}"
    return escribir_cifrado(dest, archivo_bytes)


def guardar_upload_cifrado(
    usuario: str | None,
    categoria: str,
    nombre_original: str,
    archivo_bytes: bytes,
) -> Path:
    safe = re.sub(r"[^a-zA-Z0-9._-]+", "_", Path(nombre_original).name)[:80] or "archivo.bin"
    dest = directorio_seguro(usuario, categoria) / f"{safe}{ENC_SUFFIX}"
    return escribir_cifrado(dest, archivo_bytes)


def estado_cifrado_ui() -> dict:
    """Resumen para mostrar en login / Acerca de."""
    tiene = tiene_clave_cifrado()
    fernet_ok = obtener_fernet() is not None
    return {
        "tiene_clave": tiene,
        "fernet_ok": fernet_ok,
        "root": str(SECURE_ROOT),
    }
