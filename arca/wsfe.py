"""Cliente mínimo de WSFEv1 (SOAP 1.1 a mano, sin dependencias de WSDL)."""
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .codigos import URLS
from .net import soap_post

NS = "http://ar.gov.afip.dif.FEV1/"
_Q = "{%s}" % NS

# Orden de elementos según WSDL de WSFEv1 (manual v4.7)
ORDEN_DET = ["Concepto", "DocTipo", "DocNro", "CbteDesde", "CbteHasta", "CbteFch",
             "ImpTotal", "ImpTotConc", "ImpNeto", "ImpOpEx", "ImpTrib", "ImpIVA",
             "FchServDesde", "FchServHasta", "FchVtoPago", "MonId", "MonCotiz",
             "CanMisMonExt", "CondicionIVAReceptorId", "CbtesAsoc", "Tributos", "Iva",
             "Opcionales", "Compradores", "PeriodoAsoc", "Actividades"]


class WSFEError(Exception):
    pass


def _fmt(v):
    if isinstance(v, float):
        return f"{v:.2f}"
    return escape(str(v))


def _xml_det(det):
    out = []
    for k in ORDEN_DET:
        if k not in det or det[k] in (None, "", []):
            continue
        v = det[k]
        if k == "Iva":
            out.append("<Iva>" + "".join(
                f"<AlicIva><Id>{a['Id']}</Id><BaseImp>{a['BaseImp']:.2f}</BaseImp>"
                f"<Importe>{a['Importe']:.2f}</Importe></AlicIva>" for a in v) + "</Iva>")
        elif k == "Opcionales":
            out.append("<Opcionales>" + "".join(
                f"<Opcional><Id>{escape(str(o['Id']))}</Id><Valor>{escape(str(o['Valor']))}</Valor></Opcional>"
                for o in v) + "</Opcionales>")
        elif k == "Tributos":
            out.append("<Tributos>" + "".join(
                f"<Tributo><Id>{t['Id']}</Id><Desc>{escape(t.get('Desc', ''))}</Desc>"
                f"<BaseImp>{t['BaseImp']:.2f}</BaseImp><Alic>{t['Alic']:.2f}</Alic>"
                f"<Importe>{t['Importe']:.2f}</Importe></Tributo>" for t in v) + "</Tributos>")
        elif k == "MonCotiz":
            out.append(f"<MonCotiz>{float(v):.6f}</MonCotiz>")
        elif k == "CbtesAsoc":
            out.append("<CbtesAsoc>" + "".join(
                "<CbteAsoc>" + f"<Tipo>{c['Tipo']}</Tipo><PtoVta>{c['PtoVta']}</PtoVta><Nro>{c['Nro']}</Nro>"
                + (f"<Cuit>{c['Cuit']}</Cuit>" if c.get("Cuit") else "")
                + (f"<CbteFch>{c['CbteFch']}</CbteFch>" if c.get("CbteFch") else "")
                + "</CbteAsoc>" for c in v) + "</CbtesAsoc>")
        elif k in ("Compradores", "PeriodoAsoc", "Actividades"):
            raise NotImplementedError(k)
        else:
            out.append(f"<{k}>{_fmt(v)}</{k}>")
    return "<FECAEDetRequest>" + "".join(out) + "</FECAEDetRequest>"


def xml_fecaesolicitar(auth, pto_vta, cbte_tipo, det):
    return (f'<FECAESolicitar xmlns="{NS}">{_xml_auth(auth)}'
            f"<FeCAEReq><FeCabReq><CantReg>1</CantReg><PtoVta>{pto_vta}</PtoVta>"
            f"<CbteTipo>{cbte_tipo}</CbteTipo></FeCabReq>"
            f"<FeDetReq>{_xml_det(det)}</FeDetReq></FeCAEReq></FECAESolicitar>")


def _xml_auth(auth):
    return (f"<Auth><Token>{escape(auth['token'])}</Token><Sign>{escape(auth['sign'])}</Sign>"
            f"<Cuit>{int(auth['cuit'])}</Cuit></Auth>")


def _arbol(el):
    """Elemento XML → dict. Las etiquetas repetidas quedan en lista."""
    hijos = list(el)
    if not hijos:
        return (el.text or "").strip()
    out = {}
    for h in hijos:
        k = h.tag.replace(_Q, "")
        v = _arbol(h)
        if k in out:
            cur = out[k]
            if not isinstance(cur, list):
                out[k] = [cur]
            out[k].append(v)
        else:
            out[k] = v
    return out


def _errs(node):
    res = []
    for tag in ("Errors", "Observaciones", "Events"):
        for e in node.iter(_Q + tag):
            for item in e:
                res.append((tag, item.findtext(_Q + "Code"), item.findtext(_Q + "Msg")))
    return res


class WSFE:
    def __init__(self, env):
        self.env = env
        self.url = URLS[env]["wsfe"]

    def _call(self, op, inner_xml):
        r = soap_post(self.url, f'<{op} xmlns="{NS}">{inner_xml}</{op}>', NS + op)
        try:
            root = ET.fromstring(r.content)
        except ET.ParseError:
            raise WSFEError(f"{op}: HTTP {r.status_code} respuesta no XML: {r.text[:300]}")
        fault = root.find(".//{http://schemas.xmlsoap.org/soap/envelope/}Fault")
        if fault is not None:
            raise WSFEError(f"{op}: SOAP Fault {fault.findtext('faultstring')}")
        res = root.find(f".//{_Q}{op}Result")
        if res is None:
            raise WSFEError(f"{op}: respuesta inesperada HTTP {r.status_code}")
        return res

    def dummy(self):
        res = self._call("FEDummy", "")
        return {k: res.findtext(_Q + k) for k in ("AppServer", "DbServer", "AuthServer")}

    def ultimo_autorizado(self, auth, pto_vta, cbte_tipo):
        res = self._call("FECompUltimoAutorizado",
                         f"{_xml_auth(auth)}<PtoVta>{pto_vta}</PtoVta><CbteTipo>{cbte_tipo}</CbteTipo>")
        errs = [e for e in _errs(res) if e[0] == "Errors"]
        if errs:
            raise WSFEError("; ".join(f"{c}: {m}" for _, c, m in errs))
        return int(res.findtext(_Q + "CbteNro"))

    def consultar(self, auth, pto_vta, cbte_tipo, nro):
        res = self._call("FECompConsultar",
                         f"{_xml_auth(auth)}<FeCompConsReq><CbteTipo>{cbte_tipo}</CbteTipo>"
                         f"<CbteNro>{nro}</CbteNro><PtoVta>{pto_vta}</PtoVta></FeCompConsReq>")
        g = res.find(_Q + "ResultGet")
        if g is None:
            return None, _errs(res)
        return {c.tag.replace(_Q, ""): c.text for c in g if len(c) == 0}, _errs(res)

    def consultar_completo(self, auth, pto_vta, cbte_tipo, nro):
        """Igual que consultar, pero conserva IVA, tributos y comprobantes asociados."""
        res = self._call("FECompConsultar",
                         f"{_xml_auth(auth)}<FeCompConsReq><CbteTipo>{cbte_tipo}</CbteTipo>"
                         f"<CbteNro>{nro}</CbteNro><PtoVta>{pto_vta}</PtoVta></FeCompConsReq>")
        g = res.find(_Q + "ResultGet")
        if g is None:
            return None, _errs(res)
        return _arbol(g), _errs(res)

    def solicitar_cae(self, auth, pto_vta, cbte_tipo, det):
        inner = xml_fecaesolicitar(auth, pto_vta, cbte_tipo, det)
        r = soap_post(self.url, inner, NS + "FECAESolicitar")
        root = ET.fromstring(r.content)
        fault = root.find(".//{http://schemas.xmlsoap.org/soap/envelope/}Fault")
        if fault is not None:
            raise WSFEError(f"FECAESolicitar: SOAP Fault {fault.findtext('faultstring')}")
        res = root.find(f".//{_Q}FECAESolicitarResult")
        if res is None:
            raise WSFEError(f"FECAESolicitar: respuesta inesperada HTTP {r.status_code}")
        cab = res.find(_Q + "FeCabResp")
        detr = res.find(f".//{_Q}FECAEDetResponse")
        out = {
            "Resultado": (detr.findtext(_Q + "Resultado") if detr is not None else None)
                         or (cab.findtext(_Q + "Resultado") if cab is not None else None),
            "CAE": detr.findtext(_Q + "CAE") if detr is not None else None,
            "CAEFchVto": detr.findtext(_Q + "CAEFchVto") if detr is not None else None,
            "CbteDesde": detr.findtext(_Q + "CbteDesde") if detr is not None else None,
            "mensajes": _errs(res),
        }
        return out

    def param(self, auth, metodo, extra=""):
        res = self._call(metodo, _xml_auth(auth) + extra)
        items = []
        g = res.find(_Q + "ResultGet")
        if g is not None:
            for it in g:
                items.append({c.tag.replace(_Q, ""): c.text for c in it})
        return items, _errs(res)
