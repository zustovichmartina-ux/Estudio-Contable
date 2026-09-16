# -*- coding: utf-8 -*-
"""Parser Galicia (Y + últimos 2 importes) con OCR del estudio si el PDF es escaneo."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from extracto_layout import lineas_por_y, parsear_lineas, score_cadena
from procesador import _paginas_texto_extracto_pdf

CACHE = ROOT / "_cache_extractos_sunny_galicia"

RULES = [
    ("inter-cta", "FCI - Suscripcion", re.compile(r"SUSCRIPCION FIMA|SUSCRIPCION FCI", re.I)),
    ("inter-cta", "FCI - Rescate", re.compile(r"RESCATE FIMA|RESCATE FCI", re.I)),
    ("inter-cta", "Transferencia propia", re.compile(r"TRANSF\. CTAS PROPIAS|TRANSFERENCIA DE CUENTA PROPIA", re.I)),
    ("retencion", "Ret. IIBB - ARBA", re.compile(r"ARBA AUTOM|PAGO.*ARBA|ARBA IIBB", re.I)),
    ("retencion", "ICDB Ley 25.413", re.compile(r"LEY 25\.413|IMP\. DEB\. LEY 25413|IMP\. CRE\. LEY 25413|IMP\.LEY 25413|DEB\.LEY", re.I)),
    ("retencion", "Percepcion IVA", re.compile(r"PERCEP\. IVA|PERCEPCION IVA|PERC\.IVA", re.I)),
    ("impuesto", "IVA banco", re.compile(r"(^|\s)IVA(\s|$)", re.I)),
    ("impuesto", "Pago AFIP", re.compile(r"TRANSF\. AFIP|\bVEP\b|PAGO DE SERVICIOS.*AFIP", re.I)),
    ("egreso", "Comisiones y gastos bancarios", re.compile(r"COMISION SERVICIO|COMISION BANCO|COM\. GESTION|COM\. DEPOSITO", re.I)),
    ("impuesto", "IIBB", re.compile(r"ING\. BRUTOS|IIBB", re.I)),
    ("ingreso", "Transferencia de tercero", re.compile(r"TRANSFERENCIA DE TERCEROS|CREDITO TRANSFERENCIA", re.I)),
    ("egreso", "Pago tarjeta VISA", re.compile(r"PAGO TARJETA VISA|PAGO VISA", re.I)),
    ("egreso", "Cheque / ECHEQ", re.compile(r"ECHEQ|DEPOSITO DE CHEQUE|G\.DE ECHEQ", re.I)),
    ("egreso", "Transferencia a tercero", re.compile(r"TRANSFERENCIA A TERCEROS|TRF INMED PROVEED|TRANSFERENCIAS CASH|SERVICIO PAGO A PROVEEDORES", re.I)),
    ("egreso", "Pago de servicios", re.compile(r"PAGO DE SERVICIOS", re.I)),
    ("ingreso", "Anulacion / devolucion", re.compile(r"ANULACION|DEV\.IMP", re.I)),
]


def classify(desc: str) -> tuple[str, str]:
    for cat, sub, rx in RULES:
        if rx.search(desc or ""):
            return cat, sub
    return "sin-cat", "-"


def cache_ocr(data: bytes) -> Path:
    h = hashlib.sha1(data).hexdigest()[:16]
    CACHE.mkdir(parents=True, exist_ok=True)
    return CACHE / f"ocr_{h}.json"


def lineas_ocr(data: bytes) -> list[str]:
    cp = cache_ocr(data)
    if cp.exists():
        payload = json.loads(cp.read_text(encoding="utf-8"))
        return list(payload.get("lineas") or [])
    paginas = _paginas_texto_extracto_pdf(data, dpi_ocr=170, forzar_ocr=True)
    lineas: list[str] = []
    for _n, txt in paginas:
        for ln in str(txt).splitlines():
            s = ln.strip()
            if s:
                lineas.append(s)
    cp.write_text(json.dumps({"lineas": lineas}, ensure_ascii=False), encoding="utf-8")
    return lineas


def parsear_pdf(data: bytes, nombre: str = "") -> dict:
    lineas_nat, chars = lineas_por_y(data)
    uso_ocr = chars < 50
    lineas = lineas_ocr(data) if uso_ocr else lineas_nat
    movs = parsear_lineas(lineas)
    for m in movs:
        cat, sub = classify(f"{m.get('descripcion') or ''} {m.get('detalle') or ''}")
        m["categoria"] = cat
        m["sub"] = sub
    cred = round(sum(m["credito"] for m in movs), 2)
    deb = round(sum(m["debito"] for m in movs), 2)
    cadena = score_cadena(movs)
    return {
        "archivo": nombre,
        "ocr": uso_ocr,
        "chars_nativos": chars,
        "movimientos": movs,
        "n": len(movs),
        "cadena_ok": cadena,
        "cadena_total": max(len(movs) - 1, 0),
        "creditos": cred,
        "debitos": deb,
        "saldo_final": movs[-1]["saldo"] if movs else None,
    }
