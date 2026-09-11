"""Capa de revisión contador ↔ sistema: no se archiva ni exporta a ciegas.

La IA y el motor proponen. El contador confirma. Si no cierra o no hay cuenta
del plan, se bloquea. Si el extracto no alcanza, queda pendiente.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Any

import pandas as pd

SCORE_AUTO_OK = 80.0
SCORE_MIN = 68.0
TOL_PARTIDA_DOBLE = 0.05
CODIGO_REVISAR = "99999"


def _norm(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def totales_debe_haber(rows: list[dict] | None) -> tuple[float, float, float]:
    filas = rows or []
    debe = round(sum(float(r.get("Debe") or r.get("debe") or 0) for r in filas), 2)
    haber = round(sum(float(r.get("Haber") or r.get("haber") or 0) for r in filas), 2)
    return debe, haber, round(debe - haber, 2)


def partida_doble_ok(diferencia: float, *, tol: float = TOL_PARTIDA_DOBLE) -> bool:
    return abs(float(diferencia or 0)) <= tol


def codigo_fila(row: dict) -> str:
    for k in ("Código", "Codigo", "codigo", "cuenta_sugerida"):
        if k in row and str(row.get(k) or "").strip():
            return str(row.get(k) or "").strip()
    return ""


def filas_sin_cuenta(rows: list[dict] | None) -> list[int]:
    malas = []
    for i, row in enumerate(rows or []):
        cod = codigo_fila(row)
        if not cod or cod == CODIGO_REVISAR:
            malas.append(i)
    return malas


def confianza_clasificacion(
    *,
    fuente: str,
    score: float,
    identificado: bool,
) -> str:
    """alta = se puede dejar OK; media = confirmar; baja = pendiente."""
    if fuente == "conceptos_bancos" and identificado and float(score or 0) >= SCORE_AUTO_OK:
        return "alta"
    if fuente == "conceptos_bancos" and float(score or 0) >= SCORE_MIN:
        return "media"
    if fuente == "regla_local":
        return "media"
    return "baja"


def estado_desde_confianza(tipo: str, confianza: str) -> str:
    """DEBITO_REVISAR y match débil no se dan por buenos."""
    if tipo == "DEBITO_REVISAR" or confianza == "baja":
        return "PENDIENTE"
    if confianza == "media" and tipo not in {
        "DEBITO_PROVEEDOR",
        "DEBITO_VEP",
    }:
        return "PENDIENTE"
    return "OK"


def gate_asiento(
    rows: list[dict] | None,
    *,
    bloqueantes_tango: list | None = None,
    plan_vacio: bool = False,
    tol: float = TOL_PARTIDA_DOBLE,
) -> dict[str, Any]:
    """Checklist antes de biblioteca / Excel Tango."""
    bloqueantes: list[str] = []
    advertencias: list[str] = []
    filas = list(rows or [])
    if not filas:
        bloqueantes.append("No hay líneas en la grilla para archivar o exportar.")
        return {
            "ok": False,
            "bloqueantes": bloqueantes,
            "advertencias": advertencias,
            "n_sin_cuenta": 0,
            "diferencia": 0.0,
        }

    n_sin = len(filas_sin_cuenta(filas))
    if n_sin:
        bloqueantes.append(
            f"Hay {n_sin} línea(s) sin cuenta del plan (vacía o 99999). "
            "Imputá cada una antes de seguir."
        )

    _debe, _haber, diferencia = totales_debe_haber(filas)
    if not partida_doble_ok(diferencia, tol=tol):
        bloqueantes.append(
            f"El asiento no cierra: diferencia $ {diferencia:,.2f}. "
            "Corregí la grilla hasta que Debe = Haber."
        )

    for item in bloqueantes_tango or []:
        if isinstance(item, dict):
            cod = item.get("codigo") or item.get("Código") or ""
            motivo = item.get("motivo") or item.get("Detalle") or "cuenta no imputable"
            bloqueantes.append(f"Tango rechaza {cod}: {motivo}".strip())
        else:
            bloqueantes.append(str(item))

    if plan_vacio:
        bloqueantes.append(
            "No hay plan de cuentas de esta sociedad. "
            "Vinculalo en el Editor de Clientes antes de archivar o exportar."
        )

    return {
        "ok": not bloqueantes,
        "bloqueantes": bloqueantes,
        "advertencias": advertencias,
        "n_sin_cuenta": n_sin,
        "diferencia": diferencia,
    }


def _col_plan(plan_df: pd.DataFrame, nombres: tuple[str, ...]) -> str | None:
    cols = {str(c).strip().lower(): c for c in plan_df.columns}
    for n in nombres:
        if n.lower() in cols:
            return cols[n.lower()]
    for orig in plan_df.columns:
        if _norm(str(orig)) in {_norm(n) for n in nombres}:
            return orig
    return None


def resolver_codigo_plan(
    nombre_cuenta: str,
    plan_df: pd.DataFrame | None,
    *,
    hints: dict[str, str] | None = None,
    score_min: float = 78.0,
) -> tuple[str, str, float]:
    """Cuenta del plan del cliente. Si no hay match claro → 99999 (no inventar)."""
    nombre = str(nombre_cuenta or "").strip()
    if not nombre or "identificar" in _norm(nombre):
        return CODIGO_REVISAR, nombre, 0.0

    if plan_df is None or plan_df.empty:
        hint = (hints or {}).get(nombre) or ""
        return (hint or CODIGO_REVISAR), nombre, 40.0 if hint else 0.0

    c_cod = _col_plan(plan_df, ("codigo", "código", "cuenta"))
    c_desc = _col_plan(plan_df, ("descripcion", "descripción", "nombre"))
    if not c_cod or not c_desc:
        hint = (hints or {}).get(nombre) or ""
        return (hint or CODIGO_REVISAR), nombre, 0.0

    objetivo = _norm(nombre)
    mejor_cod = CODIGO_REVISAR
    mejor_desc = nombre
    mejor = 0.0
    for _, row in plan_df.iterrows():
        desc = str(row.get(c_desc) or "").strip()
        if not desc:
            continue
        n = _norm(desc)
        if n == objetivo:
            return str(row.get(c_cod) or "").strip() or CODIGO_REVISAR, desc, 100.0
        if objetivo and (objetivo in n or n in objetivo):
            score = 92.0 if objetivo in n else 88.0
            if score > mejor:
                mejor = score
                mejor_cod = str(row.get(c_cod) or "").strip() or CODIGO_REVISAR
                mejor_desc = desc

    if mejor >= score_min:
        return mejor_cod, mejor_desc, mejor

    hint = (hints or {}).get(nombre) or ""
    if hint and hint != CODIGO_REVISAR:
        # Hint genérico del estudio: solo si ese código existe en ESTE plan.
        cods = {str(v).strip() for v in plan_df[c_cod].tolist()}
        if hint in cods:
            return hint, nombre, 55.0
    return CODIGO_REVISAR, nombre, mejor


def resumen_revision_motor(movimientos: list[dict] | None) -> dict[str, Any]:
    movs = list(movimientos or [])
    n_ok = sum(1 for m in movs if m.get("estado") == "OK")
    n_conc = sum(1 for m in movs if m.get("estado") == "CONCILIADO")
    n_pend = sum(1 for m in movs if m.get("estado") == "PENDIENTE")
    n_debiles = sum(
        1
        for m in movs
        if str(m.get("confianza") or "") == "media" or (
            str(m.get("estado")) == "PENDIENTE"
            and "confirmar" in str(m.get("match_detalle") or "").lower()
        )
    )
    n_regla_local = sum(1 for m in movs if str(m.get("fuente") or "") == "regla_local")
    return {
        "total": len(movs),
        "ok": n_ok,
        "conciliados": n_conc,
        "pendientes": n_pend,
        "debiles": n_debiles,
        "regla_local": n_regla_local,
        "requiere_confirmacion": n_pend > 0 or n_ok > 0,
        "pisar_peligro": True,
    }
