# -*- coding: utf-8 -*-
"""Reglas de nombre PDF para VEPs, comprobantes y FCC emitidas."""
from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path

_INVALID = re.compile(r'[\\/:*?"<>|]+')
_RE_MES_CARPETA = re.compile(r"^\d{2}-\d{4}$")

# Códigos AFIP → convención FCC del estudio (compras / control diario).
_AFIP_FCC = {
    1: ("FCC", "A"),
    2: ("NDC", "A"),
    3: ("NCC", "A"),
    6: ("FCC", "B"),
    7: ("NDC", "B"),
    8: ("NCC", "B"),
    11: ("FCC", "C"),
    12: ("NDC", "C"),
    13: ("NCC", "C"),
    51: ("FCC", "M"),
    201: ("FCC", "A"),
    203: ("NCC", "A"),
}


def safe_name(text: str, max_len: int = 120) -> str:
    t = _INVALID.sub(" ", str(text or "")).strip()
    t = re.sub(r"\s+", " ", t)
    return t[:max_len].strip(" .")


def nombre_fcc(*, cliente: str, pv: int | str, nro: int | str) -> str:
    """{Cliente} {0001-00000xxx}.pdf"""
    pv_s = f"{int(pv):04d}"
    nro_s = f"{int(nro):08d}"
    return f"{safe_name(cliente)} {pv_s}-{nro_s}.pdf"


def nombre_vep(*, nro: str | int, concepto: str, importe: str | float, periodo_aaaa_mm: str) -> str:
    """VEP_{nro}_{concepto}_{importe}_pago_{AAAA-MM}.pdf"""
    imp = str(importe).replace(" ", "")
    return (
        f"VEP_{safe_name(str(nro), 40)}_{safe_name(concepto, 40)}_"
        f"{safe_name(imp, 24)}_pago_{safe_name(periodo_aaaa_mm, 7)}.pdf"
    )


def nombre_comprobante(
    *,
    pv: int | str,
    nro: int | str,
    emision: str,
    desde: str,
    hasta: str,
) -> str:
    """{PV}-{NRO}_{emisión}_{desde}_a_{hasta}.pdf"""
    pv_s = f"{int(pv):04d}" if str(pv).isdigit() else safe_name(str(pv), 8)
    nro_s = f"{int(nro):08d}" if str(nro).isdigit() else safe_name(str(nro), 12)
    return (
        f"{pv_s}-{nro_s}_{safe_name(emision, 12)}_"
        f"{safe_name(desde, 12)}_a_{safe_name(hasta, 12)}.pdf"
    )


def join_destino(ruta_destino: str, filename: str) -> Path:
    return Path(ruta_destino) / filename


def iso_to_ddmmyyyy(iso: str) -> str:
    raw = (iso or "").strip()[:10]
    if len(raw) == 10 and raw[4] == "-" and raw[7] == "-":
        year, month, day = raw.split("-")
        return f"{day}/{month}/{year}"
    return raw


def nombre_mis_comprobantes(
    *,
    cliente: str,
    tipo: str,
    desde: str,
    hasta: str,
    ext: str,
) -> str:
    """Martina Zustovich MisComprobantes Emitidos 2026-01-01_a_2026-09-04.xlsx"""
    suf = (ext or "xlsx").lstrip(".")
    return (
        f"{safe_name(cliente or 'Cliente', 40)} MisComprobantes "
        f"{safe_name(tipo, 16)} {safe_name(desde, 12)}_a_{safe_name(hasta, 12)}.{suf}"
    )


def meses_del_periodo(desde: str, hasta: str) -> list[tuple[str, str]]:
    """[(YYYY-MM-DD inicio, YYYY-MM-DD fin), ...] recortado al rango pedido."""
    d0 = date.fromisoformat(desde[:10])
    d1 = date.fromisoformat(hasta[:10])
    out: list[tuple[str, str]] = []
    year, month = d0.year, d0.month
    while (year, month) <= (d1.year, d1.month):
        last = calendar.monthrange(year, month)[1]
        start = date(year, month, 1)
        end = date(year, month, last)
        if start < d0:
            start = d0
        if end > d1:
            end = d1
        out.append((start.isoformat(), end.isoformat()))
        if month == 12:
            year, month = year + 1, 1
        else:
            month += 1
    return out


def etiqueta_mm_yyyy(iso: str) -> str:
    """2026-08-01 → 08-2026 (carpeta habitual Impuestos/Portal IVA)."""
    raw = (iso or "").strip()[:10]
    if len(raw) >= 7 and raw[4] == "-":
        return f"{raw[5:7]}-{raw[:4]}"
    return safe_name(raw, 7)


def dest_portal_iva(ruta_destino: str, periodo_iso: str) -> Path:
    """Si ruta_destino ya es MM-YYYY, usarla; si no, crear subcarpeta del mes.

    Respeta UNC/Windows (`\\servidor\\...`) aunque el worker se teste en Linux.
    """
    raw = (ruta_destino or "").rstrip("\\/")
    tag = etiqueta_mm_yyyy(periodo_iso)
    tail = re.split(r"[\\/]+", raw)[-1] if raw else ""
    if _RE_MES_CARPETA.match(tail):
        return Path(ruta_destino)
    if "\\" in (ruta_destino or ""):
        return Path(f"{raw}\\{tag}")
    return Path(raw) / tag


def nombre_portal_iva_csv(lado: str) -> str:
    """Compras.csv / Ventas.csv — convención de carpeta Portal IVA."""
    return "Ventas.csv" if _lado_es_ventas(lado) else "Compras.csv"


def nombre_portal_iva_libro(lado: str, periodo_mm_yyyy: str, ext: str = "pdf") -> str:
    tag = "Ventas" if _lado_es_ventas(lado) else "Compras"
    suf = (ext or "pdf").lstrip(".")
    return f"LibroIVA_{tag}_{safe_name(periodo_mm_yyyy, 7)}.{suf}"


def codigo_afip_a_fcc(codigo: int | str) -> tuple[str, str]:
    digits = re.sub(r"\D", "", str(codigo or ""))
    try:
        n = int(digits) if digits else 0
    except ValueError:
        n = 0
    return _AFIP_FCC.get(n, ("FCC", "A"))


def nombre_fcc_compra(
    *,
    tipo: str,
    letra: str,
    pv: int | str,
    nro: int | str,
    proveedor: str,
) -> str:
    """FCCA13-157 PROVEEDOR.pdf — control diario de compras."""
    return (
        f"{safe_name(tipo, 3)}{safe_name(letra, 1)}"
        f"{int(pv)}-{int(nro)} {safe_name(proveedor, 60)}.pdf"
    )


def _lado_es_ventas(lado: str) -> bool:
    key = (lado or "").strip().lower()
    return key.startswith("vent") or "emitid" in key
