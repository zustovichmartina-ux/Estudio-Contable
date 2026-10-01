"""Facturación web: validación, persistencia y ticket, sin llamar a ARCA."""
import datetime as dt
import os
import subprocess
from argparse import Namespace

import pytest

import database
from arca import wsaa
from arca.constancia import interpretar_constancia, mensaje_no_autorizado
from arca.pdf import generar_pdf
from arca.persistencia import guardar_emisor, guardar_ta, insertar_emision, cargar_emisor
from arca.secretos import materializar
from arca.servicio import emitir_excel, excel_desde_filas, validar_excel
from arca.ta_store import hash_cert, obtener_ta_durable
from arca.vista import VistaComprobante
from emitir import EmisionAbortada, correr


def _db(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "estudio.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))


def _fila_c(**kw):
    base = dict(
        cuit_emisor="27301234568",
        razon_emisor="Mono Ejemplo",
        dom_emisor="Güemes 1",
        iibb_emisor="27301234568",
        inicio_act="01/03/2019",
        cond_iva_emisor="Responsable Monotributo",
        pto_vta=3,
        tipo="C",
        concepto="Productos",
        fecha=dt.date.today(),
        cond_venta="Contado",
        doc_tipo="DNI",
        doc_nro="30111222",
        nombre_rec="Juan Pérez",
        cond_iva="Consumidor Final",
        detalle="Honorarios",
        neto=1500,
        alicuota="No corresponde (C)",
    )
    base.update(kw)
    return base


def test_borrador_manual_sin_red(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    r = validar_excel(excel_desde_filas([_fila_c()]), "homo")
    assert r["resumen"].get("OK-DRYRUN") == 1
    assert r["filas"][0]["Estado"] == "OK-DRYRUN"
    assert r["pdfs"] and r["pdfs"][0][1].startswith(b"%PDF")
    assert r["excel"][:2] == b"PK"


def test_ejemplos_no_salen_si_no_se_piden(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    modelo = os.path.join(os.path.dirname(os.path.dirname(__file__)), "plantillas", "Facturas_a_emitir_MODELO.xlsx")
    r = validar_excel(open(modelo, "rb").read(), "homo", incluir_ejemplos=False)
    assert r["filas"] == []


def test_huella_aprobada_no_se_reenvia(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    blob = excel_desde_filas([_fila_c()])
    primero = validar_excel(blob, "homo")
    huella = primero["filas"][0]["Huella"]
    insertar_emision([
        "01/10/2026 10:00:00", "homo", "ENVÍO", "2", None, "27301234568", 3, "C",
        9, "71234567890123", "20261011", "APROBADO", "", 1500, None, huella,
        dt.date.today().strftime("%Y%m%d"), None,
    ])
    segundo = validar_excel(blob, "homo")
    assert segundo["filas"][0]["Estado"] == "YA EMITIDO"


def test_nc_referencia_factura_guardada(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    insertar_emision([
        "01/10/2026 10:00:00", "homo", "ENVÍO", "4", "F1", "27301234568", 3, "C",
        9, "71234567890123", "20261011", "APROBADO", "", 1500, None, "huella-previa",
        dt.date.today().strftime("%Y%m%d"), None,
    ])
    blob = excel_desde_filas([_fila_c(tipo="NC C", neto=400, asoc_id="F1", id=None, detalle="Bonificación")])
    r = validar_excel(blob, "homo")
    assert r["filas"][0]["Estado"] == "OK-DRYRUN", r["filas"][0]
    pdf = tmp_path / "nc.pdf"
    pdf.write_bytes(r["pdfs"][0][1])
    texto = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True).stdout
    assert "NOTA DE CRÉDITO" in texto


def test_emisor_vacio_se_completa(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    guardar_emisor("27301234568", "Estudio Guardado", "Colón 1", "Responsable Monotributo", "IIBB1", "01/03/2019")
    fila = _fila_c()
    fila["razon_emisor"] = None
    fila["dom_emisor"] = None
    r = validar_excel(excel_desde_filas([fila]), "homo")
    assert r["filas"][0]["Estado"] == "OK-DRYRUN", r["filas"]
    pdf = tmp_path / "em.pdf"
    pdf.write_bytes(r["pdfs"][0][1])
    texto = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True).stdout
    assert "Estudio Guardado" in texto


def test_ta_vigente_no_pide_otro(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    cert = tmp_path / "c.crt"
    key = tmp_path / "k.key"
    cert.write_bytes(b"certificado-de-prueba")
    key.write_bytes(b"clave")
    exp = (dt.datetime.now(wsaa.TZ) + dt.timedelta(hours=11)).isoformat()
    guardar_ta("homo", "wsfe", hash_cert(str(cert)), {
        "token": "TK", "sign": "SG", "expiration": exp,
    })
    monkeypatch.setattr(wsaa, "soap_post", lambda *a, **k: pytest.fail("no debía pedir otro TA"))
    ta = obtener_ta_durable("homo", str(cert), str(key), "wsfe", str(tmp_path / "cache"))
    assert ta["source"] == "cache" and ta["token"] == "TK"


def test_constancia_categoria_y_no_autorizado():
    xml = """<?xml version="1.0"?>
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
      <soapenv:Body>
        <personaReturn>
          <datosGenerales><razonSocial>ACME SA</razonSocial></datosGenerales>
          <datosMonotributo><categoriaMonotributo>
            <descripcionCategoria>A</descripcionCategoria>
            <idCategoria>8</idCategoria>
          </categoriaMonotributo></datosMonotributo>
        </personaReturn>
      </soapenv:Body>
    </soapenv:Envelope>""".encode()
    out = interpretar_constancia(xml)
    assert out["ok"] and out["categoria"] == "A" and out["razon"] == "ACME SA"
    fault = """<?xml version="1.0"?>
    <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
      <soapenv:Body><soapenv:Fault>
        <faultcode>ns1:coe.notAuthorized</faultcode>
        <faultstring>Computador no autorizado</faultstring>
      </soapenv:Fault></soapenv:Body>
    </soapenv:Envelope>""".encode()
    neg = interpretar_constancia(fault)
    assert neg["no_autorizado"] and "Constancia de Inscripción" in neg["mensaje"]
    assert mensaje_no_autorizado("coe.notAuthorized")


def test_materializar_pem_desde_entorno(tmp_path, monkeypatch):
    monkeypatch.setenv("ARCA_CERT_HOMO", "-----BEGIN CERTIFICATE-----\nAAA\n-----END CERTIFICATE-----")
    monkeypatch.setenv("ARCA_KEY", "-----BEGIN PRIVATE KEY-----\nBBB\n-----END PRIVATE KEY-----")
    cert, key = materializar("homo", tmp_path)
    assert "BEGIN CERTIFICATE" in open(cert).read()
    assert oct(os.stat(key).st_mode & 0o777) == "0o600"
    assert oct(os.stat(cert).st_mode & 0o777) == "0o600"


def test_produccion_sin_confirmacion_y_ejemplos():
    with pytest.raises(EmisionAbortada, match="confirmar-produccion"):
        correr(Namespace(env="prod", dry_run=False, confirmar_produccion=False, incluir_ejemplos=False, dummy=False))
    with pytest.raises(EmisionAbortada, match="EJEMPLO"):
        correr(Namespace(env="prod", dry_run=False, confirmar_produccion=True, incluir_ejemplos=True, dummy=False))


def test_emitir_produccion_sin_flag_no_llama_red(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    with pytest.raises(EmisionAbortada):
        emitir_excel(excel_desde_filas([_fila_c()]), "prod", confirmar_produccion=False)


def test_pdf_desde_consulta(tmp_path):
    cons = {
        "CbteFch": "20261001", "ImpTotal": "1500.00", "ImpNeto": "1500.00", "ImpIVA": "0",
        "ImpOpEx": "0", "ImpTotConc": "0", "DocTipo": "96", "DocNro": "30111222",
        "CondicionIVAReceptorId": "5", "CodAutorizacion": "71234567890123", "FchVto": "20261011",
    }
    vista = VistaComprobante.from_consulta(
        cons,
        {"razon_social": "Mono", "domicilio": "Güemes 1", "condicion_iva": "Responsable Monotributo",
         "iibb": "1", "inicio_actividades": "01/03/2019"},
        "27301234568", 3, 11,
    )
    destino = tmp_path / "c.pdf"
    generar_pdf(str(destino), vista, 9, "71234567890123", "20261011", marca_agua="HOMOLOGACIÓN - SIN VALIDEZ")
    assert destino.read_bytes().startswith(b"%PDF")


def test_solapa_facturacion_conectada():
    texto = open(os.path.join(os.path.dirname(os.path.dirname(__file__)), "afip_worker", "ui_streamlit.py"), encoding="utf-8").read()
    assert "Facturación" in texto and "render_facturacion_arca" in texto
    assert "Consulta" in texto and "Próximamente" in texto
    import ui_arca_facturacion
    assert callable(ui_arca_facturacion.render_facturacion_arca)


def test_modulo_arca_sin_ejecutor():
    """La UI de ARCA no muestra ni consulta la máquina ejecutor."""
    raiz = os.path.dirname(os.path.dirname(__file__))
    texto = open(os.path.join(raiz, "afip_worker", "ui_streamlit.py"), encoding="utf-8").read()
    for marca in (
        "ejecutor_arca.bat",
        "máquina ejecutor",
        "Encolar",
        "CUITs",
        "RemoteWorker",
        "AFIP_WORKER",
        "túnel",
        "_remote(",
    ):
        assert marca not in texto, marca
    assert 'st.tabs(["Facturación", "Consulta"])' in texto
