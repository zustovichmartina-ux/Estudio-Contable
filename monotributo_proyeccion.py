"""Proyección de monotributo: semestre fijo (AFIP) vs ventana rodante (planificación)."""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parent
CATEGORIAS_PATH = ROOT / "data" / "monotributo_categorias.json"

# Fuente: ARCA categorías vigentes desde 01/08/2026. Verificar en arca.gob.ar/monotributo/categorias.asp
_TOPES_DEFAULT = {
    "A": 12009410.45,
    "B": 17595182.74,
    "C": 24670494.31,
    "D": 30628651.43,
    "E": 36028231.33,
    "F": 45151659.41,
    "G": 53995798.87,
    "H": 81924660.37,
    "I": 91699761.90,
    "J": 105012519.20,
    "K": 126610838.75,
}
_FUENTE_DEFAULT = (
    "ARCA — Montos y categorías vigentes (aplicación desde el 01/08/2026). "
    "Verificar en https://www.arca.gob.ar/monotributo/categorias.asp"
)


def cargar_topes_categorias() -> dict[str, Any]:
    topes = dict(_TOPES_DEFAULT)
    fuente = _FUENTE_DEFAULT
    vigencia = "01/08/2026"
    if CATEGORIAS_PATH.exists():
        data = json.loads(CATEGORIAS_PATH.read_text(encoding="utf-8"))
        extra = {
            str(k).upper(): float(v)
            for k, v in (data.get("topes_ingresos") or {}).items()
        }
        if extra:
            topes.update(extra)
        fuente = str(data.get("fuente") or fuente)
        vigencia = str(data.get("vigencia") or vigencia)
    return {
        "fuente": fuente,
        "vigencia": vigencia,
        "topes_ingresos": topes,
    }


def tope_categoria(categoria: str, topes: dict[str, float] | None = None) -> float | None:
    cat = str(categoria or "").strip().upper()
    tabla = topes if topes is not None else cargar_topes_categorias()["topes_ingresos"]
    if cat not in tabla:
        return None
    return float(tabla[cat])


def periodo_recategorizacion_fijo(hoy: date | None = None) -> tuple[date, date, str]:
    """Último semestre de recategorización cerrado (el que se compara contra AFIP)."""
    hoy = hoy or date.today()
    if hoy.month <= 6:
        # Recategorización de enero: 01/01–31/12 del año anterior
        return date(hoy.year - 1, 1, 1), date(hoy.year - 1, 12, 31), "enero"
    # Recategorización de julio: 01/07 año anterior – 30/06 año en curso
    return date(hoy.year - 1, 7, 1), date(hoy.year, 6, 30), "julio"


def ventana_rodante_12m(hoy: date | None = None) -> tuple[date, date]:
    """Últimos 12 meses terminados en el mes en curso (incluye el mes actual)."""
    hoy = hoy or date.today()
    mes = hoy.month - 11
    anio = hoy.year
    while mes <= 0:
        mes += 12
        anio -= 1
    return date(anio, mes, 1), hoy


def _parsear_fecha_col(val) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, date) and not isinstance(val, pd.Timestamp):
        return val
    ts = pd.to_datetime(val, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def facturado_en_rango(
    df: pd.DataFrame,
    desde: date,
    hasta: date,
    col_fecha: str = "Período Desde",
    col_importe: str = "Importe Total",
) -> float:
    if df is None or df.empty or col_fecha not in df.columns or col_importe not in df.columns:
        return 0.0
    total = 0.0
    for _, row in df.iterrows():
        f = _parsear_fecha_col(row.get(col_fecha))
        if f is None or f < desde or f > hasta:
            continue
        try:
            total += float(row.get(col_importe) or 0)
        except (TypeError, ValueError):
            continue
    return round(total, 2)


def proyectar_monotributo(
    df: pd.DataFrame,
    categoria: str,
    *,
    hoy: date | None = None,
) -> dict[str, Any]:
    """
    Dos totales separados:
    - fijo: semestre/año de recategorización (control vs AFIP)
    - rodante: últimos 12 meses (alimenta «máximo a facturar este mes»)
    """
    hoy = hoy or date.today()
    cat = str(categoria or "").strip().upper()
    meta = cargar_topes_categorias()
    tope = tope_categoria(cat, meta["topes_ingresos"])
    d_fijo, h_fijo, recat = periodo_recategorizacion_fijo(hoy)
    d_rod, h_rod = ventana_rodante_12m(hoy)
    fact_fijo = facturado_en_rango(df, d_fijo, h_fijo)
    fact_rod = facturado_en_rango(df, d_rod, h_rod)
    max_mes = None if tope is None else round(tope - fact_rod, 2)
    return {
        "categoria": cat,
        "tope": tope,
        "fuente_tope": meta.get("fuente") or "",
        "vigencia_tope": meta.get("vigencia") or "",
        "recategorizacion": recat,
        "fijo_desde": d_fijo,
        "fijo_hasta": h_fijo,
        "facturado_fijo": fact_fijo,
        "rodante_desde": d_rod,
        "rodante_hasta": h_rod,
        "facturado_rodante": fact_rod,
        "max_facturar_mes": max_mes,
        "supera_tope": bool(tope is not None and fact_rod > tope),
    }
