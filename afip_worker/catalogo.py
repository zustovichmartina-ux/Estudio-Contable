# -*- coding: utf-8 -*-
"""Catálogo de CUITs con acceso conocido (sin claves). Va al repo para la web en la nube."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .registry import badge_label

_CATALOGO = Path(__file__).resolve().parent / "cuit_catalogo.json"


def load_catalogo() -> list[dict[str, Any]]:
    if not _CATALOGO.exists():
        return []
    try:
        data = json.loads(_CATALOGO.read_text(encoding="utf-8"))
    except Exception:
        return []
    items = data.get("cuits") if isinstance(data, dict) else data
    out: list[dict[str, Any]] = []
    if not isinstance(items, list):
        return []
    for raw in items:
        if not isinstance(raw, dict):
            continue
        cuit = str(raw.get("cuit") or "").strip()
        if not cuit:
            continue
        status = str(raw.get("status") or "ready")
        row = {
            "cuit": cuit,
            "razon_social": str(raw.get("razon_social") or "").strip(),
            "status": status,
            "acceso": badge_label(status),
            "note": str(raw.get("note") or "Acceso conocido en RECEPCION"),
            "updated_at": str(raw.get("updated_at") or ""),
            "chrome_profile": "",
        }
        out.append(row)
    out.sort(key=lambda r: r["cuit"])
    return out


def catalogo_lookup(cuit: str) -> dict[str, Any] | None:
    digits = "".join(ch for ch in (cuit or "") if ch.isdigit())
    for row in load_catalogo():
        other = "".join(ch for ch in str(row.get("cuit") or "") if ch.isdigit())
        if other and other == digits:
            return row
    return None
