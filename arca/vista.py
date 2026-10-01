"""Vista mínima para regenerar el PDF (la misma forma que usa arca.pdf.generar_pdf)."""
from __future__ import annotations

import datetime as dt

from arca.codigos import ETIQUETA_CBTE, NOMBRE_CBTE, TIPOS_CBTE


def _fecha(valor):
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    texto = str(valor or "").strip()
    if not texto:
        return dt.date.today()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(texto[:10] if fmt == "%Y-%m-%d" else texto, fmt).date()
        except ValueError:
            continue
    return dt.date.today()


def _num(valor, default=0):
    if valor in (None, ""):
        return default
    return float(str(valor).replace(",", "."))


def _entero(valor, default=0):
    try:
        return int(float(valor))
    except (TypeError, ValueError):
        return default


def _lista(valor):
    if valor is None:
        return []
    if isinstance(valor, list):
        return valor
    return [valor]


class VistaComprobante:
    def __init__(self):
        self.f0 = {}
        self.det = {}
        self.items = []
        self.clase = "C"
        self.pto_vta = 0
        self.cbte_tipo = 11
        self.cuit = ""
        self.fecha = dt.date.today()
        self.es_fce = False
        self.asoc_desc = None

    @classmethod
    def from_payload(cls, payload: dict) -> "VistaComprobante":
        v = cls()
        v.f0 = dict(payload.get("f0") or {})
        v.det = dict(payload.get("det") or {})
        v.items = list(payload.get("items") or [])
        v.clase = payload.get("clase") or "C"
        v.pto_vta = _entero(payload.get("pto_vta"))
        v.cbte_tipo = _entero(payload.get("cbte_tipo"))
        v.cuit = str(payload.get("cuit") or "")
        v.fecha = _fecha(payload.get("fecha") or v.det.get("CbteFch"))
        v.es_fce = bool(payload.get("es_fce"))
        v.asoc_desc = payload.get("asoc_desc")
        v.det["DocTipo"] = _entero(v.det.get("DocTipo"), 99)
        v.det["DocNro"] = _entero(v.det.get("DocNro"))
        v.det["CondicionIVAReceptorId"] = _entero(v.det.get("CondicionIVAReceptorId"), 5)
        if v.det.get("Iva"):
            for alic in v.det["Iva"]:
                if isinstance(alic, dict) and "Id" in alic:
                    alic["Id"] = _entero(alic["Id"])
        return v

    @classmethod
    def from_consulta(cls, cons: dict, emisor: dict | None, cuit: str, pto_vta: int, cbte_tipo: int) -> "VistaComprobante":
        v = cls()
        emisor = emisor or {}
        etiqueta = ETIQUETA_CBTE.get(int(cbte_tipo), "")
        meta = TIPOS_CBTE.get(etiqueta, (cbte_tipo, "C", False))
        v.cbte_tipo = int(cbte_tipo)
        v.clase = meta[1]
        v.es_fce = bool(meta[2])
        v.pto_vta = int(pto_vta)
        v.cuit = str(cuit)
        v.fecha = _fecha(cons.get("CbteFch"))
        ivas = []
        bloque = cons.get("Iva")
        for b in _lista(bloque):
            if isinstance(b, dict):
                ivas.extend(a for a in _lista(b.get("AlicIva")) if isinstance(a, dict))
        iva_norm = []
        for a in ivas:
            iva_norm.append({
                "Id": _entero(a.get("Id")),
                "BaseImp": _num(a.get("BaseImp")),
                "Importe": _num(a.get("Importe")),
            })
        doc_tipo = _entero(cons.get("DocTipo"), 99)
        doc_nro = _entero(cons.get("DocNro"))
        v.det = {
            "CbteFch": v.fecha.strftime("%Y%m%d"),
            "ImpTotal": _num(cons.get("ImpTotal")),
            "ImpNeto": _num(cons.get("ImpNeto")),
            "ImpIVA": _num(cons.get("ImpIVA")),
            "ImpOpEx": _num(cons.get("ImpOpEx")),
            "ImpTotConc": _num(cons.get("ImpTotConc")),
            "DocTipo": doc_tipo,
            "DocNro": doc_nro,
            "CondicionIVAReceptorId": _entero(cons.get("CondicionIVAReceptorId"), 5),
            "FchServDesde": cons.get("FchServDesde") or None,
            "FchServHasta": cons.get("FchServHasta") or None,
            "FchVtoPago": cons.get("FchVtoPago") or None,
            "Iva": iva_norm or None,
            "Opcionales": None,
        }
        v.f0 = {
            "razon_emisor": emisor.get("razon_social") or "",
            "dom_emisor": emisor.get("domicilio") or "",
            "cond_iva_emisor": emisor.get("condicion_iva") or "",
            "iibb_emisor": emisor.get("iibb") or "",
            "inicio_act": emisor.get("inicio_actividades") or "",
            "nombre_rec": "",
            "dom_rec": "",
            "cond_venta": "",
        }
        if iva_norm and v.clase == "A":
            v.items = [{
                "detalle": "Según comprobante autorizado en ARCA",
                "neto": a["BaseImp"],
                "alic": {5: "21%", 4: "10,5%", 6: "27%", 3: "0%"}.get(a["Id"], ""),
                "iva": a["Importe"],
                "total": a["BaseImp"] + a["Importe"],
            } for a in iva_norm]
        else:
            v.items = [{
                "detalle": "Según comprobante autorizado en ARCA",
                "neto": v.det["ImpTotal"],
                "alic": "",
                "iva": 0,
                "total": v.det["ImpTotal"],
            }]
        asocs = []
        bloque_asoc = cons.get("CbtesAsoc")
        for b in _lista(bloque_asoc):
            if isinstance(b, dict):
                asocs.extend(c for c in _lista(b.get("CbteAsoc")) if isinstance(c, dict))
        if asocs:
            c0 = asocs[0]
            et = ETIQUETA_CBTE.get(_entero(c0.get("Tipo")), str(c0.get("Tipo")))
            nombre = NOMBRE_CBTE.get(_entero(c0.get("Tipo")), "Comprobante")
            v.asoc_desc = f"{nombre} {et} {_entero(c0.get('PtoVta')):05d}-{_entero(c0.get('Nro')):08d}"
        return v
