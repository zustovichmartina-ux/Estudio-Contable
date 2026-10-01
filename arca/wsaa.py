"""WSAA: obtención y cache del Ticket de Acceso (TA).

- Firma el LoginTicketRequest (TRA) en CMS/PKCS#7 (contenido incluido, DER, base64)
  con el certificado X.509 de ARCA + clave privada del estudio.
- Cachea el TA en disco (permisos 600) por ambiente+servicio+certificado y lo reutiliza
  hasta que falten < 10 min para su vencimiento (~12 h). Esto es obligatorio en la
  práctica: ARCA responde coe.alreadyAuthenticated si se pide otro TA vigente.
"""
import base64
import datetime as dt
import hashlib
import json
import os
import random
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs7

from .codigos import URLS
from .net import soap_post

TZ = dt.timezone(dt.timedelta(hours=-3))


class WSAAError(Exception):
    pass


def crear_tra(servicio, ahora=None):
    ahora = ahora or dt.datetime.now(TZ)
    gen = (ahora - dt.timedelta(minutes=10)).replace(microsecond=0)
    exp = (ahora + dt.timedelta(minutes=10)).replace(microsecond=0)
    return ('<?xml version="1.0" encoding="UTF-8"?>'
            '<loginTicketRequest version="1.0"><header>'
            f'<uniqueId>{random.randint(1, 2**31 - 1)}</uniqueId>'
            f'<generationTime>{gen.isoformat()}</generationTime>'
            f'<expirationTime>{exp.isoformat()}</expirationTime>'
            f'</header><service>{servicio}</service></loginTicketRequest>')


def firmar_tra(tra_xml, cert_path, key_path):
    with open(cert_path, "rb") as f:
        cert = x509.load_pem_x509_certificate(f.read())
    with open(key_path, "rb") as f:
        key = serialization.load_pem_private_key(f.read(), password=None)
    cms_der = (pkcs7.PKCS7SignatureBuilder()
               .set_data(tra_xml.encode("utf-8"))
               .add_signer(cert, key, hashes.SHA256())
               .sign(serialization.Encoding.DER, [pkcs7.PKCS7Options.Binary]))
    return base64.b64encode(cms_der).decode("ascii")


def _cache_file(cache_dir, env, servicio, cert_path):
    with open(cert_path, "rb") as f:
        h = hashlib.sha256(f.read()).hexdigest()[:12]
    return os.path.join(cache_dir, f"TA_{env}_{servicio}_{h}.json")


def _parse_fecha(s):
    return dt.datetime.fromisoformat(s.strip())


def obtener_ta(env, cert_path, key_path, servicio="wsfe", cache_dir="cache", margen_min=10):
    """Devuelve dict {token, sign, expiration, source}. Usa cache si está vigente."""
    os.makedirs(cache_dir, mode=0o700, exist_ok=True)
    cf = _cache_file(cache_dir, env, servicio, cert_path)
    ahora = dt.datetime.now(TZ)
    if os.path.exists(cf):
        with open(cf) as f:
            ta = json.load(f)
        if _parse_fecha(ta["expiration"]) - dt.timedelta(minutes=margen_min) > ahora:
            ta["source"] = "cache"
            return ta

    cms = firmar_tra(crear_tra(servicio), cert_path, key_path)
    body = ('<wsaa:loginCms xmlns:wsaa="http://wsaa.view.sua.dvadac.desein.afip.gov">'
            f'<wsaa:in0>{cms}</wsaa:in0></wsaa:loginCms>')
    r = soap_post(URLS[env]["wsaa"], body, "")
    root = ET.fromstring(r.content)
    fault = root.find(".//{http://schemas.xmlsoap.org/soap/envelope/}Fault")
    if fault is not None:
        code = (fault.findtext("faultcode") or "").strip()
        msg = (fault.findtext("faultstring") or "").strip()
        if "alreadyAuthenticated" in code + msg:
            msg += (" | Ya existe un TA vigente para este certificado y servicio que no está en el"
                    " cache local (¿se pidió desde otro equipo/proceso?). Esperar a que venza (~12 h)"
                    " o usar el mismo directorio de cache.")
        raise WSAAError(f"WSAA {code}: {msg}")
    ret = root.find(".//{http://wsaa.view.sua.dvadac.desein.afip.gov}loginCmsReturn")
    if ret is None:
        raise WSAAError(f"Respuesta WSAA inesperada (HTTP {r.status_code}): {r.text[:500]}")
    tr = ET.fromstring(ret.text.encode("utf-8"))
    ta = {
        "token": tr.findtext(".//token"),
        "sign": tr.findtext(".//sign"),
        "expiration": tr.findtext(".//expirationTime"),
        "generation": tr.findtext(".//generationTime"),
        "destination": tr.findtext(".//destination"),
    }
    fd = os.open(cf, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(ta, f)
    os.chmod(cf, 0o600)
    ta["source"] = "wsaa"
    return ta
