"""Sesión HTTP para los servidores de ARCA.

Los servidores de ARCA negocian parámetros DH "chicos" que OpenSSL 3 rechaza con
SECLEVEL=2 (error DH_KEY_TOO_SMALL). Se baja el nivel sólo para estas conexiones;
la verificación del certificado del servidor se mantiene activa."""
import ssl
import requests
from requests.adapters import HTTPAdapter


class _ArcaAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)


def session():
    s = requests.Session()
    s.mount("https://", _ArcaAdapter())
    s.headers["User-Agent"] = "EstudioZG-arca_fe/1.0"
    return s


def soap_post(url, body_xml, soap_action, timeout=60):
    env = ('<?xml version="1.0" encoding="utf-8"?>'
           '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
           '<soap:Body>' + body_xml + '</soap:Body></soap:Envelope>')
    r = session().post(url, data=env.encode("utf-8"), timeout=timeout,
                       headers={"Content-Type": "text/xml; charset=utf-8",
                                "SOAPAction": f'"{soap_action}"' if soap_action else '""'})
    return r
