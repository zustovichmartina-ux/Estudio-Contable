"""Categoría de monotributo vía ws_sr_constancia_inscripcion (padrón A5).

Si el computador fiscal del estudio no tiene ese servicio autorizado, se
devuelve un mensaje claro. La facturación (wsfe) no depende de esto.
"""
from __future__ import annotations

import os
import tempfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from arca.net import soap_post
from arca.secretos import CUIT_ESTUDIO, materializar
from arca.ta_store import obtener_ta_durable
from arca.wsaa import WSAAError

URLS = {
    "homo": "https://awshomo.afip.gov.ar/sr-padron/webservices/personaServiceA5",
    "prod": "https://aws.afip.gov.ar/sr-padron/webservices/personaServiceA5",
}
SERVICIO = "ws_sr_constancia_inscripcion"
_MENSAJE_NO_AUTORIZADO = (
    "El certificado del estudio no está autorizado al servicio Constancia de Inscripción "
    "(ws_sr_constancia_inscripcion). La facturación electrónica no lo necesita. "
    "Para habilitarlo: Administrador de Relaciones de Clave Fiscal → nueva relación del "
    "servicio al computador fiscal del estudio, y aceptar la designación si ARCA lo pide."
)


def xml_get_persona(token: str, sign: str, cuit_representada: str, id_persona: str) -> str:
    return (
        '<ns1:getPersona_v2 xmlns:ns1="http://a5.soap.ws.server.puc.sr/">'
        f"<token>{escape(token)}</token><sign>{escape(sign)}</sign>"
        f"<cuitRepresentada>{int(cuit_representada)}</cuitRepresentada>"
        f"<idPersona>{int(id_persona)}</idPersona>"
        "</ns1:getPersona_v2>"
    )


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def interpretar_constancia(xml_bytes: bytes) -> dict:
    """Parsea la respuesta SOAP (o un Fault) sin red."""
    root = ET.fromstring(xml_bytes)
    fault = root.find(".//{http://schemas.xmlsoap.org/soap/envelope/}Fault")
    if fault is not None:
        code = (fault.findtext("faultcode") or "").strip()
        msg = (fault.findtext("faultstring") or "").strip()
        crudo = f"{code} {msg}"
        if any(x in crudo.lower() for x in ("notauthorized", "no autorizado", "no esta autorizado", "no está autorizado")):
            return {"ok": False, "no_autorizado": True, "mensaje": _MENSAJE_NO_AUTORIZADO, "detalle": crudo}
        return {"ok": False, "no_autorizado": False, "mensaje": f"Padrón ARCA: {crudo}", "detalle": crudo}

    errores = []
    categoria = ""
    id_categoria = ""
    razon = ""
    apellido = ""
    nombre = ""
    for el in root.iter():
        tag = _local(el.tag)
        texto = (el.text or "").strip()
        if tag in ("error", "mensaje") and texto:
            errores.append(texto)
        elif tag == "descripcionCategoria" and texto and not categoria:
            categoria = texto
        elif tag == "idCategoria" and texto and not id_categoria:
            id_categoria = texto
        elif tag == "razonSocial" and texto and not razon:
            razon = texto
        elif tag == "apellido" and texto and not apellido:
            apellido = texto
        elif tag == "nombre" and texto and not nombre:
            nombre = texto
    if not razon:
        razon = " ".join(p for p in (apellido, nombre) if p)
    if errores and not categoria:
        return {"ok": False, "no_autorizado": False, "mensaje": " | ".join(errores), "categoria": "", "razon": razon}
    if not categoria:
        return {
            "ok": True,
            "no_autorizado": False,
            "mensaje": "ARCA no informó categoría de monotributo para esa CUIT (puede ser responsable inscripto, exento o no monotributista).",
            "categoria": "",
            "razon": razon,
        }
    return {
        "ok": True,
        "no_autorizado": False,
        "mensaje": f"Categoría de monotributo: {categoria}",
        "categoria": categoria,
        "id_categoria": id_categoria,
        "razon": razon,
    }


def mensaje_no_autorizado(texto: str) -> str | None:
    crudo = (texto or "").lower()
    if any(x in crudo for x in ("notauthorized", "no autorizado", "not authorized", "coe.notauthorized")):
        return _MENSAJE_NO_AUTORIZADO
    return None


def consultar_categoria(env: str, cuit_consulta: str, cuit_representada: str | None = None) -> dict:
    """Consulta el padrón. `cuit_consulta` es la persona; la representada es el estudio."""
    representada = cuit_representada or CUIT_ESTUDIO
    tmp = tempfile.TemporaryDirectory(prefix="arca_padron_")
    try:
        cert, key = materializar(env, tmp.name)
        cache = os.path.join(tmp.name, "cache")
        try:
            ta = obtener_ta_durable(env, cert, key, SERVICIO, cache)
        except WSAAError as ex:
            aviso = mensaje_no_autorizado(str(ex))
            if aviso:
                return {"ok": False, "no_autorizado": True, "mensaje": aviso, "detalle": str(ex)}
            return {"ok": False, "no_autorizado": False, "mensaje": f"No se pudo obtener el ticket del padrón: {ex}"}
        body = xml_get_persona(ta["token"], ta["sign"], representada, cuit_consulta)
        try:
            r = soap_post(URLS[env], body, "http://a5.soap.ws.server.puc.sr/getPersona_v2")
        except Exception as ex:
            return {"ok": False, "no_autorizado": False, "mensaje": f"No se pudo consultar el padrón: {ex}"}
        try:
            return interpretar_constancia(r.content)
        except ET.ParseError:
            return {"ok": False, "no_autorizado": False, "mensaje": f"Respuesta inesperada del padrón (HTTP {r.status_code})."}
    finally:
        tmp.cleanup()
