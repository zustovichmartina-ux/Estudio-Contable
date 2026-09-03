# -*- coding: utf-8 -*-
"""Registry de CUITs AFIP: ready / needs_admin. Nunca guarda claves."""
from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Literal
from zoneinfo import ZoneInfo

from .jobs import _repo_root, normalizar_cuit

TZ = ZoneInfo("America/Argentina/Buenos_Aires")
CuitAccess = Literal["unknown", "ready", "needs_admin", "failed"]


@dataclass
class CuitEntry:
    cuit: str
    razon_social: str = ""
    status: CuitAccess = "unknown"
    note: str = ""
    updated_at: str = ""
    chrome_profile: str = ""  # solo etiqueta del perfil, sin secretos

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "CuitEntry":
        status = str(data.get("status") or "unknown")
        if status not in {"unknown", "ready", "needs_admin", "failed"}:
            status = "unknown"
        return cls(
            cuit=normalizar_cuit(str(data.get("cuit") or "")),
            razon_social=str(data.get("razon_social") or ""),
            status=status,  # type: ignore[arg-type]
            note=str(data.get("note") or ""),
            updated_at=str(data.get("updated_at") or ""),
            chrome_profile=str(data.get("chrome_profile") or ""),
        )


def registry_path() -> Path:
    raw = (os.environ.get("AFIP_CUIT_REGISTRY") or "").strip()
    if raw:
        return Path(raw)
    return _repo_root() / "jobs" / "cuit_registry.json"


def load_registry(path: Path | None = None) -> dict[str, CuitEntry]:
    p = path or registry_path()
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}
    items = data.get("cuits") if isinstance(data, dict) else data
    out: dict[str, CuitEntry] = {}
    if isinstance(items, list):
        for raw in items:
            if not isinstance(raw, dict):
                continue
            entry = CuitEntry.from_dict(raw)
            if entry.cuit:
                out[_key(entry.cuit)] = entry
    elif isinstance(items, dict):
        for raw in items.values():
            if not isinstance(raw, dict):
                continue
            entry = CuitEntry.from_dict(raw)
            if entry.cuit:
                out[_key(entry.cuit)] = entry
    return out


def save_registry(entries: dict[str, CuitEntry], path: Path | None = None) -> Path:
    p = path or registry_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "updated_at": datetime.now(TZ).isoformat(timespec="seconds"),
        "note": "Solo estado de acceso. Nunca claves ni tokens.",
        "cuits": [e.to_dict() for e in sorted(entries.values(), key=lambda x: x.cuit)],
    }
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(p)
    return p


def _key(cuit: str) -> str:
    return re.sub(r"\D", "", normalizar_cuit(cuit))


def get_cuit(cuit: str, path: Path | None = None) -> CuitEntry | None:
    return load_registry(path).get(_key(cuit))


def upsert_cuit(
    *,
    cuit: str,
    razon_social: str = "",
    status: CuitAccess | None = None,
    note: str = "",
    chrome_profile: str = "",
    path: Path | None = None,
) -> CuitEntry:
    entries = load_registry(path)
    key = _key(cuit)
    prev = entries.get(key)
    entry = CuitEntry(
        cuit=normalizar_cuit(cuit),
        razon_social=(razon_social or (prev.razon_social if prev else "")).strip(),
        status=status or (prev.status if prev else "unknown"),
        note=note if note else (prev.note if prev else ""),
        updated_at=datetime.now(TZ).isoformat(timespec="seconds"),
        chrome_profile=chrome_profile or (prev.chrome_profile if prev else ""),
    )
    entries[key] = entry
    save_registry(entries, path)
    return entry


def mark_ready(cuit: str, *, razon_social: str = "", note: str = "Chrome autofill listo", path: Path | None = None) -> CuitEntry:
    return upsert_cuit(
        cuit=cuit,
        razon_social=razon_social,
        status="ready",
        note=note,
        path=path,
    )


def mark_needs_admin(
    cuit: str,
    *,
    razon_social: str = "",
    note: str = "CUIT nuevo o sesión caída: handoff 2FA al admin",
    path: Path | None = None,
) -> CuitEntry:
    return upsert_cuit(
        cuit=cuit,
        razon_social=razon_social,
        status="needs_admin",
        note=note,
        path=path,
    )


def ensure_cuit_registered(cuit: str, razon_social: str = "", path: Path | None = None) -> CuitEntry:
    """Alta de CUIT: si no existía → needs_admin (nunca inventar claves)."""
    existing = get_cuit(cuit, path)
    if existing:
        if razon_social and not existing.razon_social:
            return upsert_cuit(cuit=cuit, razon_social=razon_social, path=path)
        return existing
    return mark_needs_admin(
        cuit,
        razon_social=razon_social,
        note="CUIT nuevo en registry: Pedir acceso / 2FA una vez en PC worker",
        path=path,
    )


def badge_label(status: CuitAccess | str) -> str:
    return {
        "ready": "Listo",
        "needs_admin": "Pedir acceso",
        "failed": "Falló auth",
        "unknown": "Desconocido",
    }.get(str(status), str(status))


def list_cuits(path: Path | None = None) -> list[CuitEntry]:
    return sorted(load_registry(path).values(), key=lambda e: e.cuit)
