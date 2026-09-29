# -*- coding: utf-8 -*-
"""Listado por Imputación de Tango (Proveedores / Clientes) → facturas pendientes
para matchear contra el banco en Conciliación.

El estudio exporta de Tango un Excel "Listado por Imputación" filtrado a una
cuenta (21101 - Proveedores, 11301 - Deudores por Ventas, etc.), con una fila
por comprobante: Cuenta, Fecha, Tipo comprobante, Numero Comprobante,
Cliente / Proveedor (CUIT), Razón social, Debe, Haber, Saldo acumulado.

Este módulo lee ese Excel, se queda solo con la cuenta pedida (Proveedores o
Clientes) y neta el Debe (pagos/notas de crédito ya registrados en Tango)
contra el Haber (facturas) por CUIT, más antigua primero (PEPS) — igual que la
posición de Inversiones. Lo que sigue pendiente es lo que hay que cancelar
contra el banco.
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime

import pandas as pd

_ALIAS_CUENTA: dict[str, tuple[str, ...]] = {
    "proveedor": ("provee",),
    "cliente": ("cliente", "deudor"),
}


def _norm_col(c) -> str:
    return re.sub(r"\s+", " ", str(c or "").strip().lower())


def _hoja_es_listado(df: pd.DataFrame) -> bool:
    cols = {_norm_col(c) for c in df.columns}
    return {"cuenta", "fecha", "debe", "haber"}.issubset(cols)


def leer_listado_imputacion(nombre: str, data: bytes) -> pd.DataFrame:
    """Concatena todas las hojas "Listado por Imputación…" del Excel de Tango.

    Tango pagina el listado en varias hojas (Conta_01, Conta_02, …) cuando es
    largo; se leen todas las que tengan el encabezado esperado y se descartan
    las demás (por ejemplo "Datos de la Empresa").
    """
    xls = pd.ExcelFile(io.BytesIO(data))
    partes: list[pd.DataFrame] = []
    for hoja in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=hoja, dtype=object)
        except Exception:
            continue
        if df is None or df.empty or not _hoja_es_listado(df):
            continue
        df = df.rename(columns={c: _norm_col(c) for c in df.columns})
        partes.append(df)
    if not partes:
        raise ValueError(
            f"«{nombre}» no tiene una hoja «Listado por Imputación» reconocible "
            "(se esperan columnas Cuenta, Fecha, Debe, Haber)."
        )
    return pd.concat(partes, ignore_index=True)


def _cuit_limpio(v) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    return d


def _money(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        s = str(v).strip().replace(".", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0


def _fecha_iso(v) -> str:
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v or "").strip()
    if not s or s.lower() == "nan":
        return ""
    try:
        return pd.to_datetime(s, dayfirst=True).date().isoformat()
    except Exception:
        return s


def facturas_pendientes(df: pd.DataFrame, *, tipo: str = "proveedor") -> list[dict]:
    """Filtra por cuenta (Proveedores/Clientes) y neta Debe contra Haber por CUIT.

    PEPS por CUIT: el pago o nota de crédito ya registrado en Tango cancela
    primero la factura más antigua de ese mismo proveedor/cliente. Devuelve
    solo las facturas que siguen con saldo pendiente, con su importe ya neto.
    """
    if "cuenta" not in df.columns:
        return []
    alias = _ALIAS_CUENTA.get(tipo, _ALIAS_CUENTA["proveedor"])
    mask = df["cuenta"].astype(str).str.lower().apply(lambda s: any(a in s for a in alias))
    sub = df[mask].copy()
    if sub.empty:
        return []

    filas = []
    for _, row in sub.iterrows():
        filas.append(
            {
                "cuit": _cuit_limpio(row.get("cliente / proveedor")),
                "razon_social": str(row.get("razón social") or "").strip(),
                "fecha": _fecha_iso(row.get("fecha")),
                "tipo_comp": str(row.get("tipo comprobante") or "").strip(),
                "num_comp": str(row.get("numero comprobante") or "").strip(),
                "debe": _money(row.get("debe")),
                "haber": _money(row.get("haber")),
            }
        )
    filas.sort(key=lambda f: (f["cuit"], f["fecha"], f["num_comp"]))

    por_cuit: dict[str, list[dict]] = {}
    for f in filas:
        por_cuit.setdefault(f["cuit"], []).append(f)

    pendientes: list[dict] = []
    for _cuit, items in por_cuit.items():
        facturas = [dict(f, importe=round(f["haber"], 2)) for f in items if f["haber"] > 0.005]
        pagos = [f["debe"] for f in items if f["debe"] > 0.005]
        for monto_pago in pagos:
            restante = monto_pago
            for fac in facturas:
                if restante <= 0.005:
                    break
                if fac["importe"] <= 0.005:
                    continue
                aplica = min(fac["importe"], restante)
                fac["importe"] = round(fac["importe"] - aplica, 2)
                restante = round(restante - aplica, 2)
        for fac in facturas:
            if fac["importe"] > 0.005:
                pendientes.append(fac)
    return pendientes


def resumen_por_tercero(pendientes: list[dict]) -> pd.DataFrame:
    """Vista rápida: cuánto se debe (o cobrar) por proveedor/cliente."""
    if not pendientes:
        return pd.DataFrame(columns=["CUIT", "Razón social", "Facturas", "Importe pendiente"])
    df = pd.DataFrame(pendientes)
    g = (
        df.groupby(["cuit", "razon_social"], dropna=False)
        .agg(facturas=("importe", "count"), importe=("importe", "sum"))
        .reset_index()
        .rename(
            columns={
                "cuit": "CUIT",
                "razon_social": "Razón social",
                "facturas": "Facturas",
                "importe": "Importe pendiente",
            }
        )
    )
    g["Importe pendiente"] = g["Importe pendiente"].round(2)
    return g.sort_values("Importe pendiente", ascending=False, ignore_index=True)
