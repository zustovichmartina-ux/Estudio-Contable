# -*- coding: utf-8 -*-
"""Lee claves AFIP ya guardadas en Edge/Chrome (DPAPI local). Nunca loguea el valor."""
from __future__ import annotations

import base64
import ctypes
import json
import logging
import os
import re
import shutil
import sqlite3
import tempfile
from ctypes import wintypes
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

LOG = logging.getLogger("afip_worker.password_vault")


class DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _crypt_unprotect(data: bytes) -> bytes:
    blob_in = DATA_BLOB(len(data), ctypes.create_string_buffer(data, len(data)))
    blob_out = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out)
    ):
        raise OSError("CryptUnprotectData failed")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def _aes_key_from_local_state(local_state: Path) -> bytes:
    ls = json.loads(local_state.read_text(encoding="utf-8"))
    enc = base64.b64decode((ls.get("os_crypt") or {})["encrypted_key"])
    if not enc.startswith(b"DPAPI"):
        raise ValueError("os_crypt key sin prefijo DPAPI")
    return _crypt_unprotect(enc[5:])


def _decrypt_password(aes_key: bytes, blob: bytes) -> str:
    if blob.startswith(b"v10") or blob.startswith(b"v11"):
        nonce, ct = blob[3:15], blob[15:]
        return AESGCM(aes_key).decrypt(nonce, ct, None).decode("utf-8")
    if blob.startswith(b"v20"):
        raise ValueError("password v20 app-bound no soportado")
    return _crypt_unprotect(blob).decode("utf-8", errors="replace")


def _digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def _candidate_stores() -> list[tuple[Path, Path]]:
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    out: list[tuple[Path, Path]] = []
    edge = local / "Microsoft" / "Edge" / "User Data"
    if (edge / "Default" / "Login Data").exists():
        out.append((edge / "Default" / "Login Data", edge / "Local State"))
    root = Path(__file__).resolve().parents[1] / "jobs" / "chrome_afip"
    if (root / "Default" / "Login Data").exists() and (root / "Local State").exists():
        out.append((root / "Default" / "Login Data", root / "Local State"))
    chrome = local / "Google" / "Chrome" / "User Data"
    for prof in ("Profile 3", "Default", "Profile 4"):
        ld = chrome / prof / "Login Data"
        if ld.exists():
            out.append((ld, chrome / "Local State"))
    return out


def lookup_afip_password(cuit: str) -> str | None:
    """Devuelve la clave fiscal para el CUIT si esta en autofill local."""
    want = _digits(cuit)
    if len(want) < 11:
        return None
    for login_data, local_state in _candidate_stores():
        tmp = tempfile.mktemp(suffix=".db")
        try:
            shutil.copy2(login_data, tmp)
            aes_key = _aes_key_from_local_state(local_state)
            con = sqlite3.connect(tmp)
            rows = con.execute(
                "select username_value, password_value, origin_url from logins"
            ).fetchall()
            con.close()
            for user, pwblob, origin in rows:
                ou = (origin or "").lower()
                if "afip" not in ou and "arca" not in ou:
                    continue
                if _digits(user or "") != want:
                    continue
                if not pwblob:
                    continue
                try:
                    secret = _decrypt_password(aes_key, pwblob)
                except Exception as exc:
                    LOG.warning(
                        "no pude desencriptar clave de %s (%s)",
                        login_data,
                        type(exc).__name__,
                    )
                    continue
                if secret:
                    label = login_data.parent.parent.name
                    LOG.info("clave AFIP encontrada en autofill (%s)", label)
                    return secret
        except Exception as exc:
            LOG.warning("vault skip %s: %s", login_data, type(exc).__name__)
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
    return None
