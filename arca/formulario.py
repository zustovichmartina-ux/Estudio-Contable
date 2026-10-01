"""Reglas del formulario de un comprobante, sin Streamlit.

Autocompletar la ficha del emisor, avisar una CUIT mal escrita y armar
una nota de crédito a partir de un comprobante ya guardado.
"""
from __future__ import annotations

import json

from arca.codigos import ALICUOTAS, ASOC_PERMITIDOS, DOC_TIPOS, ETIQUETA_CBTE, TIPOS_CBTE
from arca.comprobante import COND_IVA_EMISOR, a_fecha, cuit_valido, solo_digitos

TIPOS_MONOTRIBUTO = ("C", "NC C", "ND C")
COND_MONOTRIBUTO = "Responsable Monotributo"


def mensaje_cuit(texto: str) -> str | None:
    """None si está vacía o es válida. Aviso si está mal (largo o dígito verificador)."""
    digits = solo_digitos(texto)
    if not digits:
        return None
    if len(digits) < 11:
        return "La CUIT tiene que tener 11 dígitos."
    if len(digits) > 11:
        return "La CUIT tiene más de 11 dígitos."
    if not cuit_valido(digits):
        return "El dígito verificador de la CUIT no cierra."
    return None


def tipos_para_condicion(cond: str) -> list[str]:
    if (cond or "").strip() == COND_MONOTRIBUTO:
        return list(TIPOS_MONOTRIBUTO)
    return list(TIPOS_CBTE)


def campos_desde_emisor(emisor: dict | None) -> dict:
    """Datos para completar el formulario. Vacío si no hay ficha."""
    if not emisor:
        return {}
    cond = str(emisor.get("condicion_iva") or "").strip()
    if cond not in COND_IVA_EMISOR:
        cond = ""
    tipos = tipos_para_condicion(cond)
    tipo = str(emisor.get("tipo") or "").strip()
    if tipo not in tipos:
        tipo = "C" if cond == COND_MONOTRIBUTO else ""
    inicio = str(emisor.get("inicio_actividades") or "").strip()
    if inicio:
        try:
            inicio = a_fecha(inicio).strftime("%d/%m/%Y")
        except ValueError:
            pass
    out = {
        "razon": str(emisor.get("razon_social") or ""),
        "domicilio": str(emisor.get("domicilio") or ""),
        "iibb": str(emisor.get("iibb") or ""),
        "inicio": inicio,
        "cond": cond,
    }
    pto = emisor.get("pto_vta")
    if pto not in (None, ""):
        try:
            n = int(pto)
        except (TypeError, ValueError):
            n = 0
        if 1 <= n <= 99999:
            out["pto"] = n
    if tipo:
        out["tipo"] = tipo
    return out


def tipos_asociables(tipo_nc: str) -> set[str]:
    """Etiquetas de factura que ARCA deja asociar a esta NC/ND (misma letra)."""
    if tipo_nc not in TIPOS_CBTE:
        return set()
    cod = TIPOS_CBTE[tipo_nc][0]
    permitidos = ASOC_PERMITIDOS.get(cod, set())
    return {
        etiq for etiq, (c, _clase, _fce) in TIPOS_CBTE.items()
        if c in permitidos and not etiq.startswith(("NC", "ND"))
    }


def candidatos_asociados(filas: list[dict], tipo_nc: str) -> list[dict]:
    """Comprobantes APROBADOS de la misma letra, del más nuevo al más viejo."""
    permitidos = tipos_asociables(tipo_nc)
    salida = []
    for row in filas:
        if not str(row.get("estado") or "").startswith("APROBADO"):
            continue
        if str(row.get("tipo") or "") not in permitidos:
            continue
        if not str(row.get("numero") or "").strip():
            continue
        salida.append(row)
    return salida


def _payload(row: dict) -> dict:
    payload = row.get("payload_json")
    if isinstance(payload, str) and payload.strip():
        try:
            payload = json.loads(payload)
        except json.JSONDecodeError:
            return {}
    return payload if isinstance(payload, dict) else {}


def _fecha_txt(valor) -> str:
    if valor in (None, ""):
        return ""
    try:
        return a_fecha(valor).strftime("%d/%m/%Y")
    except ValueError:
        return str(valor)


def _importe(payload: dict, row: dict):
    items = payload.get("items") or []
    neto = 0.0
    hubo = False
    for it in items:
        if it.get("neto") not in (None, ""):
            neto += float(it["neto"])
            hubo = True
    if hubo:
        return round(neto, 2)
    f0 = payload.get("f0") or {}
    if f0.get("neto") not in (None, ""):
        return round(float(f0["neto"]), 2)
    if row.get("total") not in (None, ""):
        return round(float(row["total"]), 2)
    return None


def etiqueta_asociado(row: dict) -> str:
    payload = _payload(row)
    f0 = payload.get("f0") or {}
    tipo = str(row.get("tipo") or "")
    pto = str(row.get("pto_vta") or "")
    nro = str(row.get("numero") or "")
    fecha = _fecha_txt(row.get("fecha_cbte") or f0.get("fecha"))
    nombre = str(f0.get("nombre_rec") or "")
    total = row.get("total")
    partes = [f"#{row.get('id') or ''} {tipo} {pto}-{nro}".strip()]
    if fecha:
        partes.append(fecha)
    if nombre:
        partes.append(nombre)
    if total not in (None, ""):
        partes.append(f"${float(total):,.2f}")
    return " · ".join(p for p in partes if p)


def campos_desde_asociado(row: dict) -> dict:
    """Receptor, importe y referencia. No manda ID y tipo a la vez (ARCA pide uno)."""
    payload = _payload(row)
    f0 = payload.get("f0") or {}
    doc_tipo = str(f0.get("doc_tipo") or "").strip()
    if doc_tipo not in DOC_TIPOS:
        doc_tipo = ""
    alicuota = str(f0.get("alicuota") or "").strip()
    if alicuota not in ALICUOTAS:
        alicuota = ""
    cond = str(f0.get("cond_iva") or "").strip()
    tipo = str(row.get("tipo") or f0.get("tipo") or "").strip()
    if tipo not in TIPOS_CBTE:
        tipo = ETIQUETA_CBTE.get(row.get("cbte_tipo"), "")
    try:
        pto = int(row.get("pto_vta") or f0.get("pto_vta") or 0)
    except (TypeError, ValueError):
        pto = 0
    try:
        nro = int(row.get("numero") or 0)
    except (TypeError, ValueError):
        nro = 0
    return {
        "doc_tipo": doc_tipo,
        "doc_nro": str(f0.get("doc_nro") or ""),
        "nombre_rec": str(f0.get("nombre_rec") or ""),
        "dom_rec": str(f0.get("dom_rec") or ""),
        "cond_rec": cond,
        "detalle": str(f0.get("detalle") or ""),
        "neto": _importe(payload, row),
        "alicuota": alicuota,
        "asoc_tipo": tipo,
        "asoc_pto": pto,
        "asoc_nro": nro,
        "asoc_fecha": _fecha_txt(row.get("fecha_cbte") or f0.get("fecha")),
        "asoc_id": "",
    }
