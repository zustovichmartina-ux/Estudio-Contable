"""Escalas salariales de CCT (básicos por categoría).

Fuente Comercio julio 2026: Circular FAECYS Acuerdo 04/2026 (CCT 130/75).
"""

from __future__ import annotations

import json
import re
import unicodedata
from typing import Any

# Básicos remunerativos julio 2026 (suma NR absorbida al básico).
ESCALA_COMERCIO_130_75_JULIO_2026: dict[str, float] = {
    "Maestranza y Servicios A": 1233585.0,
    "Maestranza y Servicios B": 1236794.0,
    "Maestranza y Servicios C": 1248038.0,
    "Administrativo A": 1245631.0,
    "Administrativo B": 1250454.0,
    "Administrativo C": 1255270.0,
    "Administrativo D": 1269729.0,
    "Administrativo E": 1281775.0,
    "Administrativo F": 1299445.0,
    "Cajero A": 1249646.0,
    "Cajero B": 1255270.0,
    "Cajero C": 1262499.0,
    "Personal Auxiliar A": 1249646.0,
    "Personal Auxiliar B": 1257677.0,
    "Personal Auxiliar C": 1284184.0,
    "Auxiliar Especializado A": 1259287.0,
    "Auxiliar Especializado B": 1273743.0,
    "Vendedor A": 1249646.0,
    "Vendedor B": 1273746.0,
    "Vendedor C": 1281775.0,
    "Vendedor D": 1299445.0,
}

# Alias frecuentes (Tango / legajos) → clave canónica
_ALIAS_CATEGORIA: dict[str, str] = {
    "maestranza a": "Maestranza y Servicios A",
    "maestranza y servicios a": "Maestranza y Servicios A",
    "maestranza b": "Maestranza y Servicios B",
    "maestranza y servicios b": "Maestranza y Servicios B",
    "maestranza c": "Maestranza y Servicios C",
    "maestranza y servicios c": "Maestranza y Servicios C",
    "administrativo a": "Administrativo A",
    "admin a": "Administrativo A",
    "administrativo b": "Administrativo B",
    "admin b": "Administrativo B",
    "administrativo c": "Administrativo C",
    "admin c": "Administrativo C",
    "administrativo d": "Administrativo D",
    "admin d": "Administrativo D",
    "administrativo e": "Administrativo E",
    "admin e": "Administrativo E",
    "administrativo f": "Administrativo F",
    "admin f": "Administrativo F",
    "cajero a": "Cajero A",
    "cajeros a": "Cajero A",
    "cajero b": "Cajero B",
    "cajeros b": "Cajero B",
    "cajero c": "Cajero C",
    "cajeros c": "Cajero C",
    "personal auxiliar a": "Personal Auxiliar A",
    "auxiliar a": "Personal Auxiliar A",
    "personal auxiliar b": "Personal Auxiliar B",
    "auxiliar b": "Personal Auxiliar B",
    "personal auxiliar c": "Personal Auxiliar C",
    "auxiliar c": "Personal Auxiliar C",
    "auxiliar especializado a": "Auxiliar Especializado A",
    "especializado a": "Auxiliar Especializado A",
    "auxiliar especializado b": "Auxiliar Especializado B",
    "especializado b": "Auxiliar Especializado B",
    "vendedor a": "Vendedor A",
    "vendedores a": "Vendedor A",
    "vendedor b": "Vendedor B",
    "vendedores b": "Vendedor B",
    "vendedor c": "Vendedor C",
    "vendedores c": "Vendedor C",
    "vendedor d": "Vendedor D",
    "vendedores d": "Vendedor D",
}


def _fold(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.lower().strip()
    t = re.sub(r"[^a-z0-9\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def normalizar_categoria(categoria: str) -> str | None:
    """Devuelve la clave canónica de la escala o None."""
    key = _fold(categoria)
    if not key:
        return None
    if key in _ALIAS_CATEGORIA:
        return _ALIAS_CATEGORIA[key]
    # Match exacto contra nombres canónicos
    for canon in ESCALA_COMERCIO_130_75_JULIO_2026:
        if _fold(canon) == key:
            return canon
    # Contiene letra de categoría al final: "administrativo  a"
    m = re.search(r"\b([a-f])\b$", key)
    letra = m.group(1).upper() if m else ""
    if "maestranza" in key and letra in "ABC":
        return f"Maestranza y Servicios {letra}"
    if "admin" in key and letra in "ABCDEF":
        return f"Administrativo {letra}"
    if "cajer" in key and letra in "ABC":
        return f"Cajero {letra}"
    if "especializ" in key and letra in "AB":
        return f"Auxiliar Especializado {letra}"
    if "auxiliar" in key and letra in "ABC":
        return f"Personal Auxiliar {letra}"
    if "vendedor" in key and letra in "ABCD":
        return f"Vendedor {letra}"
    return None


def obtener_escala_desde_reglas(reglas: dict[str, Any] | None) -> dict[str, float]:
    """Lee escalasJson/escalas del CCT; si no hay, usa Comercio julio 2026."""
    if not reglas:
        return dict(ESCALA_COMERCIO_130_75_JULIO_2026)
    raw = reglas.get("escalas") or reglas.get("escalaBasicos") or {}
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            raw = {}
    if isinstance(raw, dict) and raw:
        out: dict[str, float] = {}
        for k, v in raw.items():
            try:
                out[str(k)] = float(v)
            except (TypeError, ValueError):
                continue
        if out:
            return out
    return dict(ESCALA_COMERCIO_130_75_JULIO_2026)


def resolver_basico_cct(
    categoria: str,
    reglas: dict[str, Any] | None = None,
) -> tuple[float | None, str | None]:
    """
    Resuelve sueldo básico de convenio por categoría.
    Retorna (importe, categoria_canonica) o (None, None).
    """
    canon = normalizar_categoria(categoria)
    if not canon:
        return None, None
    escala = obtener_escala_desde_reglas(reglas)
    # Match directo
    if canon in escala:
        return float(escala[canon]), canon
    # Match fold
    fold_map = {_fold(k): (k, float(v)) for k, v in escala.items()}
    hit = fold_map.get(_fold(canon))
    if hit:
        return hit[1], hit[0]
    return None, canon


def reglas_comercio_julio_2026() -> dict[str, Any]:
    """Reglas + escala FAECYS julio 2026 para sembrar COMERCIO_130_75."""
    return {
        "antiguedadPorAnioPct": 0.01,  # CCT Comercio: 1% por año
        "presentismoDivisor": 12,  # art. 40 = 1/12 ≈ 8,33%
        "horasMensuales": 200,
        "horasExtras50Multiplicador": 1.5,
        "diasMes": 30,
        "jubilacionPct": 0.11,
        "pamiPct": 0.03,
        "obraSocialPct": 0.03,
        "aporteSindicalPct": 0.02,
        "escalaVigencia": "2026-07",
        "escalaFuente": "FAECYS Acuerdo 04/2026 — Circular escalas abr-jul 2026",
        "escalas": dict(ESCALA_COMERCIO_130_75_JULIO_2026),
    }


def listar_categorias_cct(reglas: dict[str, Any] | None = None) -> list[str]:
    return sorted(obtener_escala_desde_reglas(reglas).keys())
