"""Certificado y clave del estudio. Salen de st.secrets o del entorno, nunca del repo."""
from __future__ import annotations

import os
from pathlib import Path

try:
    import streamlit as st
except Exception:
    st = None

CUIT_ESTUDIO = "23422843434"


class CredencialesError(Exception):
    pass


def _pem(valor) -> str:
    texto = str(valor or "").strip()
    if not texto or "PEGAR_" in texto:
        return ""
    if "\\n" in texto and "BEGIN" in texto:
        texto = texto.replace("\\n", "\n")
    if not texto.endswith("\n"):
        texto += "\n"
    return texto


def _desde_secrets(env: str) -> tuple[str, str]:
    if st is None:
        return "", ""
    try:
        bloque = st.secrets.get("arca", None)
    except Exception:
        return "", ""
    if not bloque:
        return "", ""
    cert_key = "cert_homo" if env == "homo" else "cert_prod"
    try:
        cert = _pem(bloque.get(cert_key, ""))
        key = _pem(bloque.get("key", ""))
    except Exception:
        return "", ""
    return cert, key


def _desde_entorno(env: str) -> tuple[str, str]:
    cert_var = "ARCA_CERT_HOMO" if env == "homo" else "ARCA_CERT_PROD"
    return _pem(os.environ.get(cert_var, "")), _pem(os.environ.get("ARCA_KEY", ""))


def pem_credenciales(env: str) -> tuple[str, str]:
    cert, key = _desde_secrets(env)
    c2, k2 = _desde_entorno(env)
    cert = cert or c2
    key = key or k2
    ambiente = "homologación" if env == "homo" else "producción"
    cual = "cert_homo" if env == "homo" else "cert_prod"
    if not cert or not key:
        raise CredencialesError(
            f"Falta el certificado de {ambiente} o la clave privada. "
            f"En Streamlit Cloud → Manage app → Settings → Secrets, completá "
            f"[arca] {cual} y [arca] key con el PEM (incluyendo BEGIN/END). "
            "No se commitean. El ejemplo está en .streamlit/secrets.toml.example y el README."
        )
    if "BEGIN CERTIFICATE" not in cert:
        raise CredencialesError("El certificado no parece un PEM: tiene que incluir BEGIN CERTIFICATE.")
    if "PRIVATE KEY" not in key:
        raise CredencialesError("La clave no parece un PEM: tiene que incluir BEGIN PRIVATE KEY.")
    return cert, key


def _escribir_600(path: Path, texto: str) -> None:
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as handle:
        handle.write(texto)
    os.chmod(path, 0o600)


def materializar(env: str, dest) -> tuple[str, str]:
    """Escribe cert y clave en `dest` (permisos 600) y devuelve las rutas."""
    cert, key = pem_credenciales(env)
    carpeta = Path(dest)
    carpeta.mkdir(parents=True, mode=0o700, exist_ok=True)
    nombre = "estudiozg-homo.crt" if env == "homo" else "estudiozg-prod.crt"
    cert_path = carpeta / nombre
    key_path = carpeta / "estudiozg.key"
    _escribir_600(cert_path, cert)
    _escribir_600(key_path, key)
    return str(cert_path), str(key_path)
