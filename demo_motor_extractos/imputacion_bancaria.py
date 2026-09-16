# -*- coding: utf-8 -*-
"""Imputación de movimientos bancarios: fijas vs deudores/proveedores."""
from __future__ import annotations

import re
from typing import Any

# Movimientos generales: siempre la misma cuenta. No se tocan a mano
# salvo que el usuario fuerce (flag clasificado_manual más adelante).
IMPUTACIONES_FIJAS: list[tuple[str, str, str, str]] = [
    # regex, imputacion, cuenta placeholder, categoria AE-Studio
    (r"IMP\.?\s*DEB\.?\s*LEY\s*25\.?413|IMP\.?\s*CRE\.?\s*LEY\s*25\.?413|DEV\.IMP\.DEB", "Impuesto Ley 25.413", "532001", "retencion"),
    (r"PERCEP\.?\s*\.?\s*IVA|PERCEPCION IVA", "Percepción IVA", "114002", "retencion"),
    (r"(^|\s)IVA(\s|$)|IVA SOBRE", "IVA comisiones bancarias", "532002", "impositivo"),
    (r"COMISION SERVICIO DE CUENTA|COM\.?\s*DEPOSITO|COM\.?\s*GESTION|GASTOS BANCO", "Gastos bancarios", "521001", "egreso"),
    (r"ARBA", "ARBA IIBB", "211003", "impositivo"),
    (r"ING\.?\s*BRUTOS|IIBB|LEY 12837", "IIBB", "532003", "impositivo"),
    (r"TRANSF\.?\s*AFIP|\bVEP\b", "Pagos AFIP", "211001", "impositivo"),
    (r"PAGO VISA|PAGO TARJETA VISA|TARJETA EMPRESA", "Tarjeta corporativa", "211002", "egreso"),
    (r"SUSCRIPCION FIMA|SUSCRIPCION FCI", "FCI suscripción (inter-cta)", "125001", "inter-cta"),
    (r"RESCATE FIMA|RESCATE FCI", "FCI rescate (inter-cta)", "125001", "inter-cta"),
    (r"TRANSF\.?\s*CTAS PROPIAS|CUENTA PROPIA", "Transferencia entre cuentas propias", "111099", "inter-cta"),
]

RE_CUIT = re.compile(r"\b((?:20|23|24|27|30|33|34)\d{8}\d)\b")


def _norm(texto: str) -> str:
    t = (texto or "").upper()
    for a, b in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")):
        t = t.replace(a, b)
    return re.sub(r"\s+", " ", t).strip()


def _blob(mov: dict[str, Any]) -> str:
    return _norm(f"{mov.get('descripcion') or ''} {mov.get('detalle') or ''}")


def _cuit(mov: dict[str, Any]) -> str:
    m = RE_CUIT.search(_blob(mov).replace("-", "").replace(" ", ""))
    if m:
        return m.group(1)
    m = RE_CUIT.search(re.sub(r"\D", "", _blob(mov)))
    return m.group(1) if m else ""


def buscar_padron(cuit: str, nombre: str, padron: dict[str, Any]) -> dict[str, Any] | None:
    cuit_n = re.sub(r"\D", "", cuit or "")
    nom = _norm(nombre)
    for tipo in ("deudores", "proveedores"):
        for row in padron.get(tipo) or []:
            c = re.sub(r"\D", "", str(row.get("cuit") or ""))
            if cuit_n and c == cuit_n:
                return {**row, "tipo_padron": tipo}
            rn = _norm(str(row.get("nombre") or ""))
            if rn and len(rn) >= 5 and rn in nom:
                return {**row, "tipo_padron": tipo}
    return None


def imputar_movimiento(mov: dict[str, Any], padron: dict[str, Any] | None = None) -> dict[str, Any]:
    """Devuelve el movimiento + imputacion, origen (fija|sugerida|pendiente) y contraparte."""
    padron = padron or {}
    out = dict(mov)
    blob = _blob(mov)
    credito = float(mov.get("credito") or 0)
    debito = float(mov.get("debito") or 0)

    for rx, imputacion, cuenta, cat in IMPUTACIONES_FIJAS:
        if re.search(rx, blob, flags=re.I):
            out.update(
                {
                    "imputacion": imputacion,
                    "cuenta": cuenta,
                    "origen_imputacion": "fija",
                    "contraparte": "",
                    "tipo_contraparte": "",
                    "categoria": cat,
                    "editable": False,
                }
            )
            return out

    cuit = _cuit(mov)
    hit = buscar_padron(cuit, blob, padron)
    if hit:
        tipo = hit.get("tipo_padron") or ""
        nombre = str(hit.get("nombre") or "")
        if tipo == "deudores":
            imputacion = hit.get("imputacion") or f"Deudores por ventas · {nombre}"
            cuenta = str(hit.get("cuenta") or "113001")
        else:
            imputacion = hit.get("imputacion") or f"Proveedores · {nombre}"
            cuenta = str(hit.get("cuenta") or "211010")
        out.update(
            {
                "imputacion": imputacion,
                "cuenta": cuenta,
                "origen_imputacion": "sugerida",
                "contraparte": nombre,
                "tipo_contraparte": "Deudor" if tipo == "deudores" else "Proveedor",
                "cuit_contraparte": re.sub(r"\D", "", str(hit.get("cuit") or cuit)),
                "categoria": "ingreso" if credito > 0 else "egreso",
                "editable": True,
            }
        )
        return out

    # Sin padron: si hay CUIT/nombre de transferencia, queda pendiente de imputar
    if cuit or re.search(r"TRANSFERENCIA|TRF INMED|CASH PROVEED|PAGO A PROVEED|ECHEQ|ANULACION", blob):
        lado = "Deudor (cobranza)" if credito > 0 else "Proveedor (pago)"
        out.update(
            {
                "imputacion": f"A imputar · {lado}",
                "cuenta": "",
                "origen_imputacion": "pendiente",
                "contraparte": cuit or "",
                "tipo_contraparte": "Deudor" if credito > 0 else "Proveedor",
                "cuit_contraparte": cuit,
                "categoria": "ingreso" if credito > 0 else "egreso",
                "editable": True,
            }
        )
        return out

    out.update(
        {
            "imputacion": "A imputar",
            "cuenta": "",
            "origen_imputacion": "pendiente",
            "contraparte": "",
            "tipo_contraparte": "",
            "categoria": "ingreso" if credito > 0 else "egreso",
            "editable": True,
        }
    )
    return out


def imputar_extracto(movimientos: list[dict[str, Any]], padron: dict[str, Any] | None = None) -> dict[str, Any]:
    filas = [imputar_movimiento(m, padron) for m in movimientos]
    n_fija = sum(1 for f in filas if f.get("origen_imputacion") == "fija")
    n_sug = sum(1 for f in filas if f.get("origen_imputacion") == "sugerida")
    n_pend = sum(1 for f in filas if f.get("origen_imputacion") == "pendiente")
    return {
        "movimientos": filas,
        "resumen": {
            "total": len(filas),
            "fijas": n_fija,
            "sugeridas": n_sug,
            "pendientes": n_pend,
        },
    }
