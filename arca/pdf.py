"""PDF simple del comprobante con datos fiscales, CAE, vencimiento y QR de ARCA."""
import base64
import datetime as dt
import io
import json
from decimal import Decimal

import qrcode
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

from .codigos import COND_IVA_RECEPTOR, NOMBRE_CBTE, NOMBRE_DOC, QR_URL

W, H = A4


def url_qr(fecha, cuit, pto_vta, tipo, nro, importe, doc_tipo, doc_nro, cae,
           moneda="PES", ctz=1, tipo_cod_aut="E"):
    """Texto del QR según 'Especificaciones del QR' de ARCA (versión 1)."""
    num = lambda x, n: int(x) if float(x).is_integer() else round(float(x), n)  # noqa: E731
    datos = {"ver": 1, "fecha": fecha.strftime("%Y-%m-%d"), "cuit": int(cuit), "ptoVta": int(pto_vta),
             "tipoCmp": int(tipo), "nroCmp": int(nro), "importe": num(importe, 2), "moneda": moneda,
             "ctz": num(ctz, 6)}
    if doc_tipo and int(doc_tipo) != 99:
        datos["tipoDocRec"] = int(doc_tipo)
        datos["nroDocRec"] = int(doc_nro)
    datos.update({"tipoCodAut": tipo_cod_aut, "codAut": int(cae)})
    js = json.dumps(datos, separators=(",", ":"))
    return QR_URL + "?p=" + base64.b64encode(js.encode()).decode(), datos


def _m(x):
    x = Decimal(str(x))
    s = f"{x:,.2f}"
    return "$ " + s.replace(",", "X").replace(".", ",").replace("X", ".")


def _f(s):
    if not s:
        return ""
    if isinstance(s, (dt.date, dt.datetime)):
        return s.strftime("%d/%m/%Y")
    s = str(s)
    if len(s) == 8 and s.isdigit():
        return f"{s[6:]}/{s[4:6]}/{s[:4]}"
    return s


def _cuit(c):
    c = str(c)
    return f"{c[:2]}-{c[2:10]}-{c[10:]}" if len(c) == 11 else c


def generar_pdf(path, cbte, nro, cae, cae_vto, marca_agua=None):
    """cbte: arca.comprobante.Comprobante validado. marca_agua: texto diagonal (EJEMPLO/BORRADOR)."""
    f0, det = cbte.f0, cbte.det
    clase = cbte.clase
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"{NOMBRE_CBTE[cbte.cbte_tipo]} {clase} {cbte.pto_vta:05d}-{nro:08d}")
    if marca_agua:
        c.saveState()
        from reportlab.pdfbase.pdfmetrics import stringWidth
        tam = min(60, 60 * (H * 0.72) / max(stringWidth(marca_agua, "Helvetica-Bold", 60), 1))
        c.setFont("Helvetica-Bold", tam)
        c.setFillColorRGB(0.85, 0.85, 0.85)
        c.translate(W / 2, H / 2)
        c.rotate(40)
        c.drawCentredString(0, 0, marca_agua)
        c.restoreState()

    m = 12 * mm
    top = H - m
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(W / 2, top - 4 * mm, "ORIGINAL")
    c.rect(m, top - 52 * mm, W - 2 * m, 46 * mm)
    # Letra
    bx = W / 2 - 9 * mm
    c.rect(bx, top - 24 * mm, 18 * mm, 16 * mm)
    c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(W / 2, top - 19 * mm, clase)
    c.setFont("Helvetica", 7)
    c.drawCentredString(W / 2, top - 23 * mm, f"COD. {cbte.cbte_tipo:03d}")
    c.line(W / 2, top - 24 * mm, W / 2, top - 52 * mm)
    # Emisor
    y = top - 14 * mm
    from reportlab.pdfbase.pdfmetrics import stringWidth
    razon = str(f0.get("razon_emisor") or "")[:60]
    tam = 13
    while tam > 7 and stringWidth(razon, "Helvetica-Bold", tam) > (W / 2 - 11 * mm) - (m + 3 * mm):
        tam -= 0.5
    c.setFont("Helvetica-Bold", tam)
    c.drawString(m + 3 * mm, y, razon)
    c.setFont("Helvetica", 8.5)
    cond_emisor = (f0.get("cond_iva_emisor")
                   or ("Responsable Monotributo" if clase == "C" else "IVA Responsable Inscripto"))
    for i, t in enumerate([f"Domicilio comercial: {f0.get('dom_emisor') or ''}",
                           f"Condición frente al IVA: {cond_emisor}"]):
        c.drawString(m + 3 * mm, y - (8 + 5 * i) * mm, t[:70])
    # Comprobante
    x2 = W / 2 + 12 * mm
    c.setFont("Helvetica-Bold", 12)
    titulo = NOMBRE_CBTE[cbte.cbte_tipo]
    if cbte.es_fce or len(titulo) > 18:
        c.setFont("Helvetica-Bold", 8.5)
    c.drawString(x2, y, titulo)
    c.setFont("Helvetica", 8.5)
    lineas = [f"Punto de venta: {cbte.pto_vta:05d}    Comp. Nro: {nro:08d}",
              f"Fecha de emisión: {_f(det['CbteFch'])}",
              f"CUIT: {_cuit(cbte.cuit)}",
              f"Ingresos Brutos: {f0.get('iibb_emisor') or ''}",
              f"Inicio de actividades: {_f(f0.get('inicio_act'))}"]
    for i, t in enumerate(lineas):
        c.drawString(x2, y - (7 + 5 * i) * mm, t)

    # Comprobante asociado (NC / ND)
    if getattr(cbte, "asoc_desc", None):
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(m + 3 * mm, top - 50 * mm, f"Comprobante asociado: {cbte.asoc_desc}")
    # Período de servicios
    y = top - 60 * mm
    c.rect(m, y - 2 * mm, W - 2 * m, 8 * mm)
    c.setFont("Helvetica", 8.5)
    if det.get("FchServDesde"):
        c.drawString(m + 3 * mm, y + 1 * mm, f"Período facturado desde: {_f(det['FchServDesde'])}   "
                     f"hasta: {_f(det['FchServHasta'])}   Fecha de vto. para el pago: {_f(det['FchVtoPago'])}")
    else:
        c.drawString(m + 3 * mm, y + 1 * mm, "Concepto: Productos" +
                     (f"   Fecha de vto. para el pago: {_f(det['FchVtoPago'])}" if det.get("FchVtoPago") else ""))

    # Receptor
    y -= 26 * mm
    c.rect(m, y, W - 2 * m, 22 * mm)
    cond = COND_IVA_RECEPTOR.get(det["CondicionIVAReceptorId"], ("", None))[0]
    doc = NOMBRE_DOC.get(det["DocTipo"], str(det["DocTipo"]))
    docnro = _cuit(det["DocNro"]) if det["DocTipo"] in (80, 86) else (str(det["DocNro"]) if det["DocTipo"] != 99 else "")
    c.setFont("Helvetica", 8.5)
    nombre = f0.get("nombre_rec") or ""
    c.drawString(m + 3 * mm, y + 16 * mm, f"{doc}: {docnro}" if det["DocTipo"] != 99 else "Documento: (no requerido)")
    c.drawString(m + 95 * mm, y + 16 * mm, f"Razón social: {str(nombre)[:50]}")
    c.drawString(m + 3 * mm, y + 10 * mm, f"Condición frente al IVA: {cond}")
    c.drawString(m + 95 * mm, y + 10 * mm, f"Domicilio: {str(f0.get('dom_rec') or '')[:55]}")
    c.drawString(m + 3 * mm, y + 4 * mm, f"Condición de venta: {f0.get('cond_venta') or ''}")
    if cond == "Consumidor Final":
        c.drawString(m + 95 * mm, y + 4 * mm, "A CONSUMIDOR FINAL")

    # Ítems
    y -= 8 * mm
    c.setFillColorRGB(0.9, 0.9, 0.9)
    c.rect(m, y - 1 * mm, W - 2 * m, 6 * mm, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 8)
    cols = [m + 2 * mm, W - m - 75 * mm, W - m - 50 * mm, W - m - 25 * mm, W - m - 2 * mm]
    c.drawString(cols[0], y + 1 * mm, "Descripción")
    if clase in ("A",):
        c.drawRightString(cols[2], y + 1 * mm, "Neto")
        c.drawRightString(cols[3], y + 1 * mm, "Alícuota")
        c.drawRightString(cols[4], y + 1 * mm, "Subtotal c/IVA")
    else:
        c.drawRightString(cols[4], y + 1 * mm, "Importe")
    c.setFont("Helvetica", 8.5)
    for it in cbte.items:
        y -= 6 * mm
        c.drawString(cols[0], y + 1 * mm, str(it["detalle"] or "")[:70])
        if clase == "A":
            c.drawRightString(cols[2], y + 1 * mm, _m(it["neto"]))
            c.drawRightString(cols[3], y + 1 * mm, it["alic"])
            c.drawRightString(cols[4], y + 1 * mm, _m(it["total"]))
        else:  # B y C: importe final (IVA incluido en B)
            c.drawRightString(cols[4], y + 1 * mm, _m(it["total"]))

    # Totales
    y = 85 * mm
    c.rect(m, y - 22 * mm, W - 2 * m, 30 * mm)
    c.setFont("Helvetica", 9)
    filas = []
    if clase == "A":
        filas.append(("Importe neto gravado:", det["ImpNeto"]))
        for a in det.get("Iva") or []:
            nombre_a = {5: "21%", 4: "10,5%", 6: "27%", 3: "0%", 8: "5%", 9: "2,5%"}.get(a["Id"], a["Id"])
            filas.append((f"IVA {nombre_a}:", a["Importe"]))
        if det["ImpOpEx"]:
            filas.append(("Importe exento:", det["ImpOpEx"]))
        if det["ImpTotConc"]:
            filas.append(("Importe no gravado:", det["ImpTotConc"]))
    else:
        filas.append(("Subtotal:", det["ImpTotal"]))
    filas.append(("IMPORTE TOTAL:", det["ImpTotal"]))
    yy = y + 4 * mm
    for i, (t, v) in enumerate(filas):
        if i == len(filas) - 1:
            c.setFont("Helvetica-Bold", 10)
        c.drawRightString(W - m - 40 * mm, yy, t)
        c.drawRightString(W - m - 3 * mm, yy, _m(v))
        yy -= 5 * mm
    if clase == "B":
        c.setFont("Helvetica", 7.5)
        c.drawString(m + 3 * mm, y - 18 * mm, "Régimen de Transparencia Fiscal al Consumidor (Ley 27.743)")
        c.drawString(m + 3 * mm, y - 14 * mm + 0.5 * mm, f"IVA contenido: {_m(det['ImpIVA'])}")
    if cbte.es_fce:
        c.setFont("Helvetica", 7.5)
        opc = {o["Id"]: o["Valor"] for o in det.get("Opcionales") or []}
        c.drawString(m + 3 * mm, y + 4 * mm, f"CBU emisor: {opc.get('2101', '')}   Transferencia: {opc.get('27', '')}")

    # Pie: QR + CAE
    qr_txt, _ = url_qr(cbte.fecha, cbte.cuit, cbte.pto_vta, cbte.cbte_tipo, nro, det["ImpTotal"],
                       det["DocTipo"], det["DocNro"], cae or 0)
    img = qrcode.make(qr_txt, box_size=4, border=1)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), m, 15 * mm, 40 * mm, 40 * mm)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(m + 45 * mm, 45 * mm, "ARCA")
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(m + 45 * mm, 39 * mm, "Comprobante Autorizado")
    c.setFont("Helvetica", 6.5)
    c.drawString(m + 45 * mm, 34 * mm, "Esta Agencia no se responsabiliza por los datos ingresados en el detalle de la operación")
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(W - m, 45 * mm, f"CAE N°: {cae or ''}")
    c.drawRightString(W - m, 39 * mm, f"Fecha de Vto. de CAE: {_f(cae_vto)}")
    c.setFont("Helvetica", 6)
    c.drawString(m, 10 * mm, qr_txt[:160])
    c.showPage()
    c.save()
    return qr_txt
