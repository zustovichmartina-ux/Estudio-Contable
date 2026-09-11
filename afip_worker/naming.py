# -*- coding: utf-8 -*-
"""Reglas de nombre PDF para VEPs, comprobantes y FCC emitidas."""
from __future__ import annotations

import re
from pathlib import Path

_INVALID = re.compile(r'[\\/:*?"<>|]+')


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
