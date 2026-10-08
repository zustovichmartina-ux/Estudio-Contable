"""Control de comprobantes: cruza lo que entrega el cliente contra el CSV de
"Mis Comprobantes" de ARCA, arma el pedido de faltantes (PDF) y corrige
percepciones mal informadas.

REGLA CENTRAL: el CSV original se respeta. Cada línea se guarda cruda; al
exportar, las filas sin cambios salen idénticas y en las modificadas solo se
reemplazan los campos tocados (no se usa pandas.to_csv).
"""
from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

import pandas as pd
import streamlit as st

NAVY = "#1F3557"
GOLD = "#B08D48"
MAIL_DEFAULT = "hernant@etrujillo.com.ar"

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
    "septiembre", "octubre", "noviembre", "diciembre",
]

TIPOS_COMPROBANTE = {
    1: "Factura A", 2: "Nota de Débito A", 3: "Nota de Crédito A", 4: "Recibo A",
    6: "Factura B", 7: "Nota de Débito B", 8: "Nota de Crédito B", 9: "Recibo B",
    11: "Factura C", 12: "Nota de Débito C", 13: "Nota de Crédito C", 15: "Recibo C",
    19: "Factura E", 51: "Factura M", 52: "Nota de Débito M", 53: "Nota de Crédito M",
    81: "Tique Factura A", 82: "Tique Factura B", 83: "Tique", 111: "Tique Factura C",
    201: "Factura de Crédito MiPyME A", 202: "Nota de Débito MiPyME A",
    203: "Nota de Crédito MiPyME A", 206: "Factura de Crédito MiPyME B",
    211: "Factura de Crédito MiPyME C",
}

# claves internas de los seis tributos y su etiqueta
TRIBUTOS = ("otros_nac", "iibb", "mun", "iva", "int", "otros_trib")
ETIQUETA_TRIB = {
    "otros_nac": "Otros Imp. Nac.", "iibb": "Perc. IIBB", "mun": "Municipales",
    "iva": "Perc. IVA", "int": "Internos", "otros_trib": "Otros Tributos",
}
TRIB_REVISAR = ("otros_nac", "mun", "int", "otros_trib")

DESTINOS_FILA = ("Por defecto", "No gravado", "Exento")


def _stretch() -> dict:
    """Ancho completo, compatible con versiones viejas y nuevas de Streamlit."""
    try:
        mayor, menor = (int(x) for x in st.__version__.split(".")[:2])
        if (mayor, menor) >= (1, 50):
            return {"width": "stretch"}
    except Exception:
        pass
    return {"use_container_width": True}


_STRETCH = _stretch()


# --- utilidades de texto / números ---
def _norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _split_campos(linea: str) -> list[str]:
    """Separa por ';' respetando comillas. Devuelve los campos CRUDOS."""
    campos, cur, comillas = [], [], False
    for ch in linea:
        if ch == '"':
            comillas = not comillas
            cur.append(ch)
        elif ch == ";" and not comillas:
            campos.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    campos.append("".join(cur))
    return campos


def _valor(tok: str) -> str:
    t = tok.strip()
    if len(t) >= 2 and t[0] == '"' and t[-1] == '"':
        t = t[1:-1].replace('""', '"')
    return t


def _dec(tok: str) -> Decimal:
    t = _valor(tok).replace(" ", "")
    if not t:
        return Decimal(0)
    try:
        return Decimal(t.replace(",", "."))
    except InvalidOperation:
        return Decimal(0)


def _fmt_campo(d: Decimal) -> str:
    """Formato del CSV de ARCA: 2 decimales, coma, sin miles."""
    d = d.quantize(Decimal("0.01"), ROUND_HALF_UP)
    if d == 0:
        return "0,00"
    return f"{d:.2f}".replace(".", ",")


def parse_importe(txt) -> Decimal | None:
    """Acepta '1.596,66', '1596,66' y '1596.66'. None si no se entiende."""
    t = str(txt if txt is not None else "").strip().replace("$", "").replace(" ", "")
    if t == "":
        return Decimal(0)
    neg = t.startswith("-")
    t = t.lstrip("+-")
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        if t.count(",") > 1:
            return None
        t = t.replace(",", ".")
    elif "." in t:
        if t.count(".") > 1 or re.fullmatch(r"\d{1,3}\.\d{3}", t):
            t = t.replace(".", "")
    if not re.fullmatch(r"\d+(\.\d+)?", t):
        return None
    d = Decimal(t)
    return -d if neg else d


def fmt_ar(d: Decimal, signo: bool = True) -> str:
    """$ 1.234,56 (negativos: -$ 123,45)."""
    d = Decimal(d).quantize(Decimal("0.01"), ROUND_HALF_UP)
    neg = d < 0
    s = f"{abs(d):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if neg else ''}$ {s}" if signo else f"{'-' if neg else ''}{s}"


def tipo_en_palabras(n: int) -> str:
    return TIPOS_COMPROBANTE.get(n, f"Comprobante tipo {n}")


def fmt_numero(pv: int, num: int) -> str:
    return f"{pv:05d}-{num:08d}"


# --- lectura del CSV ---
def _decodificar(raw: bytes) -> tuple[str, str, bool]:
    bom = raw.startswith(b"\xef\xbb\xbf")
    body = raw[3:] if bom else raw
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            texto = body.decode(enc)
        except UnicodeDecodeError:
            continue
        if "�" in texto:
            continue
        return texto, enc, bom
    return body.decode("latin-1"), "latin-1", bom


def _ubicar_columnas(hn: list[str]) -> dict:
    def primero(pred):
        for i, h in enumerate(hn):
            if pred(h):
                return i
        return None

    cols = {
        "fecha": primero(lambda h: h.startswith("fecha de emision") or h == "fecha"),
        "tipo": primero(lambda h: h.startswith("tipo de comprobante") or h == "tipo"),
        "pv": primero(lambda h: h.startswith("punto de venta")),
        "num": primero(lambda h: h in ("numero de comprobante", "numero desde")),
        "denom": primero(lambda h: h.startswith("denominacion")),
        "doc": primero(lambda h: h.startswith("nro. doc")),
        "total": primero(lambda h: h in ("importe total", "imp. total")),
        "ng": primero(lambda h: h == "importe no gravado"),
        "ex": primero(lambda h: h == "importe exento"),
        "iva": primero(lambda h: "percepciones o pagos a cuenta de iva" in h),
        "iibb": primero(lambda h: "ingresos brutos" in h),
        "mun": primero(lambda h: "municipales" in h),
        "int": primero(lambda h: "impuestos internos" in h),
        "otros_trib": primero(lambda h: "otros tributos" in h),
        "otros_nac": primero(lambda h: h.startswith("importe de per. o pagos a cta")),
    }
    faltan = [k for k in ("fecha", "tipo", "pv", "num", "denom", "doc", "total") if cols[k] is None]
    if faltan:
        raise ValueError(
            "No encontré estas columnas en el CSV: " + ", ".join(faltan)
            + ". ¿Es el CSV de Mis Comprobantes de ARCA?"
        )
    return cols


def _entero(tok: str) -> int:
    try:
        return int(re.sub(r"\D", "", _valor(tok)) or 0)
    except ValueError:
        return 0


def parsear_csv(raw: bytes) -> dict:
    texto, enc, bom = _decodificar(raw)
    lineas = re.split(r"\r\n|\n|\r", texto)
    while lineas and lineas[-1].strip() == "":
        lineas.pop()
    if len(lineas) < 2:
        raise ValueError("El CSV no tiene comprobantes.")
    header_raw = lineas[0]
    hn = [_norm(_valor(t)) for t in _split_campos(header_raw)]
    cols = _ubicar_columnas(hn)
    es_ventas = any(h.startswith("denominacion comprador") for h in hn)
    rows = []
    for n, linea in enumerate(lineas[1:]):
        if linea.strip() == "":
            continue
        tok = _split_campos(linea)
        if len(tok) < len(hn):
            tok = tok + [""] * (len(hn) - len(tok))
        fecha_txt = _valor(tok[cols["fecha"]])
        try:
            fecha = dt.datetime.strptime(fecha_txt[:10], "%Y-%m-%d").date()
        except ValueError:
            fecha = None
        trib = {
            k: (_dec(tok[cols[k]]) if cols[k] is not None else Decimal(0))
            for k in TRIBUTOS
        }
        rows.append({
            "i": len(rows),
            "raw": linea,
            "tok": tok,
            "fecha": fecha,
            "tipo": _entero(tok[cols["tipo"]]),
            "pv": _entero(tok[cols["pv"]]),
            "num": _entero(tok[cols["num"]]),
            "nombre": _valor(tok[cols["denom"]]),
            "cuit": _valor(tok[cols["doc"]]),
            "total": _dec(tok[cols["total"]]),
            "trib": trib,
        })
    return {
        "header_raw": header_raw, "rows": rows, "cols": cols, "enc": enc,
        "bom": bom, "es_ventas": es_ventas,
    }


def detectar_periodo(nombre: str) -> str:
    m = re.search(r"periodo[_\- ]?(\d{4})(\d{2})", nombre, re.I)
    if not m:
        return ""
    anio, mes = int(m.group(1)), int(m.group(2))
    if not 1 <= mes <= 12:
        return ""
    tipo = ""
    low = nombre.lower()
    if "compras" in low:
        tipo = " — Compras"
    elif "ventas" in low:
        tipo = " — Ventas"
    return f"{MESES[mes - 1].capitalize()} {anio}{tipo}"


# --- percepciones ---
def tiene_tributos(r: dict) -> bool:
    return any(r["trib"][k] != 0 for k in TRIBUTOS)


def _perc_default(r: dict) -> dict:
    return {
        "iva": _fmt_campo(r["trib"]["iva"]),
        "iibb": _fmt_campo(r["trib"]["iibb"]),
        "destino": "Por defecto",
        "revisado": False,
    }


def evaluar_perc(r: dict, p: dict, destino_global: str) -> dict:
    """Calcula pozo, resto y estado (OK / Revisar / Error) de una fila."""
    pozo = sum((r["trib"][k] for k in TRIBUTOS), Decimal(0))
    iva = parse_importe(p["iva"])
    iibb = parse_importe(p["iibb"])
    destino = destino_global if p["destino"] == "Por defecto" else p["destino"]
    res = {"pozo": pozo, "iva": iva, "iibb": iibb, "resto": None,
           "destino": destino, "estado": "OK", "msg": ""}
    if iva is None or iibb is None:
        res.update(estado="Error", msg="Importe inválido")
        return res
    resto = pozo - iva - iibb
    res["resto"] = resto
    if abs(iva) + abs(iibb) > abs(pozo) or (pozo != 0 and (iva * pozo < 0 or iibb * pozo < 0)):
        res.update(estado="Error", msg="IVA + IIBB supera el total de tributos o tiene otro signo")
        return res
    d = _perc_default(r)
    editado = (
        iva != _dec(d["iva"]) or iibb != _dec(d["iibb"]) or p["destino"] != "Por defecto"
    )
    hay_dudosos = any(r["trib"][k] != 0 for k in TRIB_REVISAR)
    if hay_dudosos and not editado and not p["revisado"]:
        res.update(estado="Revisar", msg="Tiene montos para revisar")
    return res


def _etiquetas_perc(r: dict, estado: str) -> str:
    t = r["trib"]
    et = []
    if t["iva"] != 0:
        et.append("IVA")
    if t["iibb"] != 0:
        et.append("IIBB")
    if any(t[k] != 0 for k in TRIB_REVISAR):
        et.append("Revisar" if estado == "Revisar" else "Otros")
    return " · ".join(et)


# --- exportación fiel del CSV ---
def _linea_exportada(data: dict, r: dict, perc: dict, destino_global: str) -> tuple[str | None, str]:
    """(línea, error). Si no hay cambios, devuelve la línea cruda original."""
    if not tiene_tributos(r):
        return r["raw"], ""
    p = perc.get(r["i"]) or _perc_default(r)
    ev = evaluar_perc(r, p, destino_global)
    if ev["estado"] == "Error":
        return None, ev["msg"]
    cols = data["cols"]
    tok = list(r["tok"])
    nuevo = {
        "iva": ev["iva"], "iibb": ev["iibb"], "otros_nac": Decimal(0), "mun": Decimal(0),
        "int": Decimal(0), "otros_trib": Decimal(0),
    }
    destino_col = "ng" if ev["destino"] == "No gravado" else "ex"
    if ev["resto"] != 0 and cols.get(destino_col) is None:
        return None, "El CSV no tiene la columna de destino del resto"
    for k, v in nuevo.items():
        if cols.get(k) is None:
            continue
        if v != r["trib"][k]:
            tok[cols[k]] = _fmt_campo(v)
    if ev["resto"] != 0 and cols.get(destino_col) is not None:
        orig = _dec(r["tok"][cols[destino_col]])
        tok[cols[destino_col]] = _fmt_campo(orig + ev["resto"])
    if tok == r["tok"]:
        return r["raw"], ""
    return ";".join(tok), ""


def exportar_csv(data: dict, indices: list[int], perc: dict, destino_global: str) -> tuple[bytes | None, list[str]]:
    errores = []
    lineas = [data["header_raw"]]
    for i in indices:
        r = data["rows"][i]
        linea, err = _linea_exportada(data, r, perc, destino_global)
        if linea is None:
            errores.append(f"{fmt_numero(r['pv'], r['num'])} {r['nombre']}: {err}")
        else:
            lineas.append(linea)
    if errores:
        return None, errores
    texto = "\r\n".join(lineas) + "\r\n"
    enc = data["enc"]
    out = texto.encode(enc)
    if data["bom"]:
        out = b"\xef\xbb\xbf" + out
    return out, []


# --- PDF ---
def fecha_larga(hoy: dt.date | None = None) -> str:
    if hoy is None:
        try:
            from zoneinfo import ZoneInfo
            hoy = dt.datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()
        except Exception:
            hoy = dt.date.today()
    return f"Mar del Plata, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}"


def nombre_pdf(cliente: str, periodo: str) -> str:
    base = f"Faltantes - {cliente.strip() or 'Cliente'} - {periodo.strip() or 'Periodo'}.pdf"
    return re.sub(r'[\\/:*?"<>|]', "-", base)


def generar_pdf(rows: list[dict], cliente: str, periodo: str, mail: str, hoy: dt.date | None = None) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.platypus import Flowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    navy, gold = colors.HexColor(NAVY), colors.HexColor(GOLD)
    gris = colors.HexColor("#6B7280")
    ancho = A4[0] - 36 * mm

    class Membrete(Flowable):
        def __init__(self):
            super().__init__()
            self.width, self.height = ancho, 24 * mm

        def wrap(self, aw, ah):
            return self.width, self.height

        def draw(self):
            c = self.canv
            top = self.height
            c.setFillColor(navy)
            c.setFont("Times-Bold", 24)
            c.drawString(0, top - 8 * mm, "H. TRUJILLO")
            c.setFillColor(gold)
            t = c.beginText(0, top - 13.5 * mm)
            t.setFont("Helvetica-Bold", 8.5)
            t.setCharSpace(1.6)
            t.textOut("CONTADORES PÚBLICOS S.A.")
            t.setCharSpace(0)
            c.drawText(t)
            c.setFillColor(gris)
            c.setFont("Helvetica", 8)
            c.drawString(0, top - 18 * mm, "Estudio Profesional en Cs. Económicas")
            c.setFont("Helvetica", 8.5)
            for n, txt in enumerate(("Alvear 3278, Mar del Plata — CP 7600", "(223) 154-568372", mail or "")):
                c.drawRightString(self.width, top - (6 + 4.2 * n) * mm, txt)
            c.setStrokeColor(gold)
            c.setLineWidth(1.2)
            c.line(0, 1 * mm, self.width, 1 * mm)

    class NumberedCanvas(rl_canvas.Canvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for st_ in self._saved:
                self.__dict__.update(st_)
                self._pie(total)
                super().showPage()
            super().save()

        def _pie(self, total):
            self.setStrokeColor(gold)
            self.setLineWidth(0.6)
            self.line(18 * mm, 15 * mm, A4[0] - 18 * mm, 15 * mm)
            self.setFillColor(gris)
            self.setFont("Helvetica", 8)
            self.drawString(18 * mm, 10.5 * mm, "H. Trujillo Contadores Públicos S.A.")
            self.drawRightString(A4[0] - 18 * mm, 10.5 * mm, f"Página {self._pageNumber} de {total}")

    base = ParagraphStyle("b", fontName="Helvetica", fontSize=10, leading=14, textColor=colors.HexColor("#222222"))
    titulo = ParagraphStyle("t", parent=base, fontName="Times-Bold", fontSize=16, leading=20, textColor=navy)
    der = ParagraphStyle("d", parent=base, alignment=TA_RIGHT, textColor=gris)
    celda = ParagraphStyle("c", parent=base, fontSize=9, leading=11)
    celda_der = ParagraphStyle("cd", parent=celda, alignment=TA_RIGHT)
    celda_h = ParagraphStyle("ch", parent=celda, fontName="Helvetica-Bold", textColor=colors.white)
    celda_h_der = ParagraphStyle("chd", parent=celda_h, alignment=TA_RIGHT)
    prov_st = ParagraphStyle("p", parent=celda, fontName="Helvetica-Bold", textColor=navy)

    story = [Membrete(), Spacer(1, 6 * mm)]
    cab = Table(
        [[Paragraph("Comprobantes pendientes de entrega", titulo), Paragraph(escape(fecha_larga(hoy)), der)]],
        colWidths=[ancho * 0.6, ancho * 0.4],
    )
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [cab, Spacer(1, 4 * mm)]
    story.append(Paragraph(f"<b>Cliente:</b> {escape(cliente or '—')}", base))
    story.append(Paragraph(f"<b>Período:</b> {escape(periodo or '—')}", base))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "Para completar la registración del período les solicitamos que nos envíen los "
        "siguientes comprobantes, que figuran en ARCA y todavía no recibimos:", base))
    story.append(Spacer(1, 4 * mm))

    anchos = [24 * mm, 54 * mm, 46 * mm, ancho - 124 * mm]
    filas = [[Paragraph("Fecha", celda_h), Paragraph("Comprobante", celda_h),
              Paragraph("Número", celda_h), Paragraph("Importe total", celda_h_der)]]
    estilos = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 1), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
    ]
    grupos: dict[tuple, list[dict]] = {}
    for r in rows:
        grupos.setdefault((r["nombre"], r["cuit"]), []).append(r)
    total = Decimal(0)
    for (nombre, cuit) in sorted(grupos, key=lambda k: (k[0].casefold(), k[1])):
        n = len(filas)
        filas.append([Paragraph(f"{escape(nombre.upper())} · CUIT {escape(cuit)}", prov_st), "", "", ""])
        estilos += [("SPAN", (0, n), (-1, n)), ("BACKGROUND", (0, n), (-1, n), colors.HexColor("#E9E4D8"))]
        for r in sorted(grupos[(nombre, cuit)], key=lambda x: (x["fecha"] or dt.date.min, x["pv"], x["num"])):
            f = r["fecha"].strftime("%d/%m/%Y") if r["fecha"] else ""
            filas.append([
                Paragraph(f, celda), Paragraph(escape(tipo_en_palabras(r["tipo"])), celda),
                Paragraph(fmt_numero(r["pv"], r["num"]), celda), Paragraph(fmt_ar(r["total"]), celda_der),
            ])
            total += r["total"]
    tabla = Table(filas, colWidths=anchos, repeatRows=1)
    tabla.setStyle(TableStyle(estilos))
    story.append(tabla)
    story.append(Spacer(1, 3 * mm))
    n_comp = len(rows)
    pie_tabla = Table(
        [[Paragraph(f"<b>{n_comp} comprobante{'s' if n_comp != 1 else ''}</b>", base),
          Paragraph(f"<b>Total: {fmt_ar(total)}</b>", ParagraphStyle("tt", parent=base, alignment=TA_RIGHT))]],
        colWidths=[ancho / 2, ancho / 2])
    pie_tabla.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, 0), 1.2, gold), ("LINEBELOW", (0, 0), (-1, 0), 1.2, gold),
                                   ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    story.append(pie_tabla)
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(
        f"Pueden enviarlos por correo a {escape(mail or MAIL_DEFAULT)} o acercarlos al estudio. "
        "Ante cualquier consulta, quedamos a disposición.", base))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=18 * mm, bottomMargin=22 * mm,
                            title="Comprobantes pendientes de entrega", author="H. Trujillo Contadores Públicos S.A.")
    doc.build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()


# --- estado ---
def _ss():
    return st.session_state


def _rows() -> list[dict]:
    return _ss()["cc_data"]["rows"]


def _desc(r: dict) -> str:
    f = r["fecha"].strftime("%d/%m/%Y") if r["fecha"] else "s/f"
    return f"{r['pv']}-{r['num']} · {r['nombre']} · {f} · {fmt_ar(r['total'])}"


def _cargar_archivo(raw: bytes, nombre: str) -> None:
    data = parsear_csv(raw)
    s = _ss()
    s["cc_data"] = data
    s["cc_hash"] = hashlib.sha256(raw).hexdigest()
    s["cc_nombre"] = nombre
    s["cc_tengo"] = set()
    s["cc_nopedir"] = set()
    s["cc_perc"] = {r["i"]: _perc_default(r) for r in data["rows"] if tiene_tributos(r)}
    s["cc_opciones"] = []
    s["cc_msg"] = None
    s["cc_ver"] = 0
    s["cc_periodo"] = detectar_periodo(nombre)
    s["cc_destino_def"] = "No gravado"
    s.pop("cc_destino_w", None)


def _marcar(i: int) -> None:
    s = _ss()
    r = _rows()[i]
    s["cc_opciones"] = []
    if i in s["cc_tengo"]:
        s["cc_msg"] = ("warning", f"Ya estaba marcado: {_desc(r)}")
    else:
        s["cc_tengo"].add(i)
        s["cc_msg"] = ("success", f"Marcado como Tengo: {_desc(r)}")


def _buscar() -> None:
    s = _ss()
    val = str(s.get("cc_busq", "")).strip()
    s["cc_busq"] = ""
    s["cc_opciones"] = []
    if not val:
        return
    m = re.fullmatch(r"(\d+)\s*[-\s]\s*(\d+)", val)
    if m:
        pv, num = int(m.group(1)), int(m.group(2))
    elif val.isdigit():
        pv, num = None, int(val)
    else:
        s["cc_msg"] = ("error", "Escribí solo el número (o punto de venta y número, ej.: 2-31145).")
        return
    coinc = [r for r in _rows() if r["num"] == num and (pv is None or r["pv"] == pv)]
    if not coinc:
        s["cc_msg"] = ("error", f"No encontré el comprobante {val} en el CSV")
    elif len(coinc) == 1:
        _marcar(coinc[0]["i"])
    else:
        s["cc_opciones"] = [r["i"] for r in coinc]
        s["cc_msg"] = ("info", f"El número {num} coincide con {len(coinc)} comprobantes. Elegí cuál tenés:")


def _marcar_todos() -> None:
    s = _ss()
    s["cc_tengo"] = {r["i"] for r in _rows()}
    s["cc_msg"] = ("success", f"Marcados los {len(_rows())} comprobantes como Tengo. Desmarcá los que falten.")
    s["cc_ver"] += 1


def _desmarcar_todos() -> None:
    s = _ss()
    s["cc_tengo"] = set()
    s["cc_msg"] = ("info", "Quedaron todos como Falta.")
    s["cc_conf_desm"] = False
    s["cc_ver"] += 1


def _faltantes() -> list[int]:
    t = _ss()["cc_tengo"]
    return [r["i"] for r in _rows() if r["i"] not in t]


def _pedidos() -> list[int]:
    n = _ss()["cc_nopedir"]
    return [i for i in _faltantes() if i not in n]


def _pedir_todos(valor: bool) -> None:
    s = _ss()
    if valor:
        s["cc_nopedir"] -= set(_faltantes())
    else:
        s["cc_nopedir"] |= set(_faltantes())
    s["cc_ver"] += 1


def _cb_pedir_todos() -> None:
    _pedir_todos(bool(_ss().get("cc_pedir_todos")))


# --- persistencia ---
def _avance_json() -> bytes:
    s = _ss()
    obj = {
        "version": 1,
        "hash": s["cc_hash"],
        "archivo": s["cc_nombre"],
        "tengo": sorted(s["cc_tengo"]),
        "no_pedir": sorted(s["cc_nopedir"]),
        "percepciones": {str(i): p for i, p in s["cc_perc"].items()},
        "destino_defecto": s.get("cc_destino_def", "No gravado"),
        "cliente": s.get("cc_cliente", ""),
        "periodo": s.get("cc_periodo", ""),
        "mail": s.get("cc_mail", MAIL_DEFAULT),
    }
    return json.dumps(obj, ensure_ascii=False, indent=1).encode("utf-8")


def _aplicar_avance(raw: bytes) -> tuple[bool, str]:
    s = _ss()
    try:
        obj = json.loads(raw.decode("utf-8"))
    except Exception:
        return False, "El archivo de avance no se puede leer."
    if obj.get("hash") != s["cc_hash"]:
        return False, "Ese avance es de otro CSV (el hash no coincide). Subí el mismo CSV con el que lo guardaste."
    n = len(_rows())
    s["cc_tengo"] = {int(i) for i in obj.get("tengo", []) if 0 <= int(i) < n}
    s["cc_nopedir"] = {int(i) for i in obj.get("no_pedir", []) if 0 <= int(i) < n}
    for k, p in (obj.get("percepciones") or {}).items():
        k = int(k)
        if k in s["cc_perc"]:
            s["cc_perc"][k] = {
                "iva": str(p.get("iva", s["cc_perc"][k]["iva"])),
                "iibb": str(p.get("iibb", s["cc_perc"][k]["iibb"])),
                "destino": p.get("destino", "Por defecto") if p.get("destino") in DESTINOS_FILA else "Por defecto",
                "revisado": bool(p.get("revisado", False)),
            }
    s["cc_destino_def"] = obj.get("destino_defecto") if obj.get("destino_defecto") in ("No gravado", "Exento") else "No gravado"
    s.pop("cc_destino_w", None)
    s["cc_cliente"] = obj.get("cliente", "")
    s["cc_periodo"] = obj.get("periodo", s.get("cc_periodo", ""))
    s["cc_mail"] = obj.get("mail", MAIL_DEFAULT)
    s["cc_ver"] += 1
    return True, "Avance retomado."


def _cb_aplicar_avance() -> None:
    s = _ss()
    av = s.get("cc_up_avance")
    if av is None:
        return
    ok, txt = _aplicar_avance(av.getvalue())
    s["cc_msg_avance"] = (ok, txt)


# --- pantalla ---
_CSS = f"""
<style>
.st-key-cc_busq input {{ font-size: 1.7rem; padding: .7rem .9rem; font-weight: 600; color: {NAVY}; }}
.cc-titulo {{ color: {NAVY}; font-weight: 700; }}
.cc-cont {{ font-size: 1.15rem; font-weight: 700; color: {NAVY}; }}
.cc-cont span {{ color: {GOLD}; }}
</style>
"""


def _estilo_filas(df: pd.DataFrame, colores: dict):
    def f(row):
        c = colores.get(row.name, "")
        return [c] * len(row)
    return df.style.apply(f, axis=1)


def _editor(df: pd.DataFrame, colores: dict, **kw):
    try:
        return st.data_editor(_estilo_filas(df, colores), **kw)
    except Exception:
        return st.data_editor(df, **kw)


def _tabla_comprobantes() -> None:
    s = _ss()
    rows = _rows()
    tengo, nopedir = s["cc_tengo"], s["cc_nopedir"]
    destino_g = s.get("cc_destino_def", "No gravado")

    filtro = st.radio("Mostrar", ["Todos", "Tengo", "Faltan", "Con percepciones"],
                      horizontal=True, key="cc_filtro")

    falt = _faltantes()
    ped = _pedidos()
    s["cc_pedir_todos"] = bool(falt) and len(ped) == len(falt)
    c1, c2 = st.columns([2, 3])
    c1.checkbox("Pedir todos los faltantes", key="cc_pedir_todos", on_change=_cb_pedir_todos,
                disabled=not falt)
    c2.markdown(f"Se van a pedir **{len(ped)}** de **{len(falt)}** faltantes")
    if filtro == "Faltan":
        b1, b2, _ = st.columns([1, 1, 4])
        b1.button("Pedir todos", key="cc_btn_pedir", on_click=_pedir_todos, args=(True,))
        b2.button("No pedir ninguno", key="cc_btn_nopedir", on_click=_pedir_todos, args=(False,))

    datos, colores = [], {}
    for r in rows:
        i = r["i"]
        es_tengo = i in tengo
        if filtro == "Tengo" and not es_tengo:
            continue
        if filtro == "Faltan" and es_tengo:
            continue
        if filtro == "Con percepciones" and not tiene_tributos(r):
            continue
        pedir = (not es_tengo) and (i not in nopedir)
        estado_perc = ""
        if tiene_tributos(r):
            ev = evaluar_perc(r, s["cc_perc"].get(i) or _perc_default(r), destino_g)
            estado_perc = _etiquetas_perc(r, ev["estado"])
        datos.append({
            "i": i,
            "Fecha": r["fecha"].strftime("%d/%m/%Y") if r["fecha"] else "",
            "Tipo": tipo_en_palabras(r["tipo"]),
            "Pto. venta": f"{r['pv']:05d}",
            "Número": f"{r['num']:08d}",
            "Proveedor" if not s["cc_data"]["es_ventas"] else "Cliente": r["nombre"],
            "CUIT": r["cuit"],
            "Importe total": fmt_ar(r["total"]),
            "Percepciones": estado_perc,
            "Estado": "Tengo" if es_tengo else "Falta",
            "Tengo": es_tengo,
            "Pedir": pedir,
        })
        colores[i] = "background-color: #E3F3E6" if es_tengo else (
            "" if pedir else "background-color: #EDEDED; color: #7A7A7A")
    if not datos:
        st.info("No hay comprobantes para mostrar con este filtro.")
        return
    df = pd.DataFrame(datos).set_index("i")
    df.index.name = None
    nombre_contra = "Proveedor" if not s["cc_data"]["es_ventas"] else "Cliente"
    editado = _editor(
        df, colores,
        key=f"cc_editor_{filtro}_{s['cc_ver']}_{hash(tuple(df.index))}",
        **_STRETCH, hide_index=True,
        height=min(640, 38 + 35 * len(df)),
        disabled=["Fecha", "Tipo", "Pto. venta", "Número", nombre_contra, "CUIT",
                  "Importe total", "Percepciones", "Estado"],
        column_config={
            "Tengo": st.column_config.CheckboxColumn("Tengo", width="small"),
            "Pedir": st.column_config.CheckboxColumn("Pedir", width="small",
                                                      help="Solo se pide si es un faltante"),
            "Estado": st.column_config.TextColumn("Estado", width="small"),
            "Pto. venta": st.column_config.TextColumn("Pto. venta", width="small"),
            "Fecha": st.column_config.TextColumn("Fecha", width="small"),
        },
    )
    cambio = False
    for i in editado.index:
        t_old, t_new = bool(df.at[i, "Tengo"]), bool(editado.at[i, "Tengo"])
        p_old, p_new = bool(df.at[i, "Pedir"]), bool(editado.at[i, "Pedir"])
        if t_new != t_old:
            (tengo.add if t_new else tengo.discard)(i)
            cambio = True
        if p_new != p_old:
            cambio = True
            if i not in tengo:
                (nopedir.discard if p_new else nopedir.add)(i)
    if cambio:
        # Sin cambiar la key del editor: así la tabla conserva el scroll y no vuelve arriba.
        st.rerun()


def _tabla_percepciones() -> None:
    s = _ss()
    rows = [r for r in _rows() if tiene_tributos(r)]
    cols = s["cc_data"]["cols"]
    if not rows:
        st.info("Ningún comprobante trae tributos distintos de cero.")
        return
    faltan_cols = [ETIQUETA_TRIB[k] for k in TRIBUTOS if cols.get(k) is None]
    if faltan_cols or cols.get("ng") is None or cols.get("ex") is None:
        st.warning("Este CSV no tiene todas las columnas de tributos / no gravado / exento: "
                   "las correcciones de percepciones pueden no aplicarse completas.")
    st.caption(
        "ARCA a veces informa las percepciones en la columna equivocada. Cargá la Perc. IVA y la Perc. IIBB "
        "reales; lo que sobre del total de tributos va a No gravado o Exento. El Importe Total no se toca."
    )
    def _sync_destino():
        s["cc_destino_def"] = s["cc_destino_w"]

    st.radio("Destino del resto por defecto", ["No gravado", "Exento"], horizontal=True,
             index=0 if s.get("cc_destino_def", "No gravado") == "No gravado" else 1,
             key="cc_destino_w", on_change=_sync_destino)
    destino_g = s.get("cc_destino_def", "No gravado")
    solo_rev = st.checkbox("Mostrar solo los que hay que revisar o tienen error", key="cc_solo_rev")

    datos, colores = [], {}
    for r in rows:
        i = r["i"]
        p = s["cc_perc"].setdefault(i, _perc_default(r))
        ev = evaluar_perc(r, p, destino_g)
        if solo_rev and ev["estado"] == "OK":
            continue
        t = r["trib"]
        datos.append({
            "i": i,
            "Fecha": r["fecha"].strftime("%d/%m/%Y") if r["fecha"] else "",
            "Proveedor": r["nombre"],
            "Comprobante": f"{r['pv']:05d}-{r['num']:08d}",
            **{ETIQUETA_TRIB[k]: float(t[k]) for k in TRIBUTOS},
            "Perc. IVA real": p["iva"],
            "Perc. IIBB real": p["iibb"],
            "Resto": float(ev["resto"]) if ev["resto"] is not None else None,
            "Destino": p["destino"],
            "Revisado": bool(p["revisado"]),
            "Estado": ev["estado"] if not ev["msg"] else f"{ev['estado']}: {ev['msg']}",
        })
        colores[i] = {"Error": "background-color: #F8D7DA", "Revisar": "background-color: #FFF3CD"}.get(ev["estado"], "")
    if not datos:
        st.success("No queda nada para revisar.")
        return
    df = pd.DataFrame(datos).set_index("i")
    df.index.name = None
    fmt_num = "%.2f"
    editado = _editor(
        df, colores,
        key=f"cc_perc_editor_{s['cc_ver']}_{int(solo_rev)}_{destino_g}",
        **_STRETCH, hide_index=True,
        height=min(640, 38 + 35 * len(df)),
        disabled=["Fecha", "Proveedor", "Comprobante", "Resto", "Estado"] + [ETIQUETA_TRIB[k] for k in TRIBUTOS],
        column_config={
            **{ETIQUETA_TRIB[k]: st.column_config.NumberColumn(ETIQUETA_TRIB[k], format=fmt_num)
               for k in TRIBUTOS},
            "Resto": st.column_config.NumberColumn("Resto", format=fmt_num),
            "Perc. IVA real": st.column_config.TextColumn("Perc. IVA real", help="Ej.: 1.596,66 o 1596.66"),
            "Perc. IIBB real": st.column_config.TextColumn("Perc. IIBB real"),
            "Destino": st.column_config.SelectboxColumn("Destino", options=list(DESTINOS_FILA)),
            "Revisado": st.column_config.CheckboxColumn("Revisado"),
        },
    )
    cambio = False
    for i in editado.index:
        p = s["cc_perc"][i]
        nuevo = {
            "iva": str(editado.at[i, "Perc. IVA real"] if editado.at[i, "Perc. IVA real"] is not None else ""),
            "iibb": str(editado.at[i, "Perc. IIBB real"] if editado.at[i, "Perc. IIBB real"] is not None else ""),
            "destino": editado.at[i, "Destino"] or "Por defecto",
            "revisado": bool(editado.at[i, "Revisado"]),
        }
        if nuevo != {k: p[k] for k in nuevo}:
            s["cc_perc"][i] = nuevo
            cambio = True
    if cambio:
        s["cc_ver"] += 1
        st.rerun()


def _pendientes_revisar() -> int:
    s = _ss()
    n = 0
    for r in _rows():
        if tiene_tributos(r):
            ev = evaluar_perc(r, s["cc_perc"].get(r["i"]) or _perc_default(r), s.get("cc_destino_def", "No gravado"))
            if ev["estado"] == "Revisar":
                n += 1
    return n


def _descargas() -> None:
    s = _ss()
    data = s["cc_data"]
    base = re.sub(r"\.csv$", "", s["cc_nombre"], flags=re.I)
    tengo_idx = [r["i"] for r in _rows() if r["i"] in s["cc_tengo"]]
    falt_idx = _faltantes()
    pend = _pendientes_revisar()
    if pend:
        st.warning(f"Hay {pend} comprobante(s) con tributos para revisar en la solapa Percepciones. "
                   "Podés descargar igual, pero se van a mover a No gravado / Exento tal como están.")
    c1, c2, c3 = st.columns(3)
    for col, titulo, idxs, suf in (
        (c1, "CSV depurado (Tengo)", tengo_idx, "_depurado.csv"),
        (c2, "Faltantes en CSV", falt_idx, "_faltantes.csv"),
    ):
        out, errs = exportar_csv(data, idxs, s["cc_perc"], s.get("cc_destino_def", "No gravado"))
        if errs:
            col.error("No se puede descargar: hay errores en percepciones.")
            col.caption("; ".join(errs[:3]) + (" …" if len(errs) > 3 else ""))
        else:
            col.download_button(titulo, out, file_name=base + suf, mime="text/csv",
                                key=f"cc_dl_{suf}", disabled=len(idxs) == 0, **_STRETCH)
    ped = _pedidos()
    pdf = b""
    if ped:
        pdf = generar_pdf([data["rows"][i] for i in ped], s.get("cc_cliente", ""), s.get("cc_periodo", ""),
                          s.get("cc_mail", MAIL_DEFAULT))
    c3.download_button("Pedido de faltantes (PDF)", pdf,
                       file_name=nombre_pdf(s.get("cc_cliente", ""), s.get("cc_periodo", "")),
                       mime="application/pdf", key="cc_dl_pdf", disabled=not ped, **_STRETCH)
    if not ped:
        c3.caption("No hay faltantes tildados en Pedir.")


def render_control_comprobantes() -> None:
    s = _ss()
    st.markdown(_CSS, unsafe_allow_html=True)
    st.markdown("<h2 class='cc-titulo'>Control de comprobantes</h2>", unsafe_allow_html=True)
    st.caption("Subí el CSV de Mis Comprobantes de ARCA, marcá lo que entregó el cliente y pedí lo que falta.")

    up = st.file_uploader("CSV de Mis Comprobantes (ARCA)", type=["csv"], key="cc_up")
    if up is None:
        if "cc_data" in s:
            st.info("Hay un CSV cargado en esta sesión. Si cerrás la pestaña se pierde: usá Guardar avance.")
        else:
            return
    else:
        raw = up.getvalue()
        if hashlib.sha256(raw).hexdigest() != s.get("cc_hash"):
            try:
                _cargar_archivo(raw, up.name)
            except ValueError as e:
                st.error(str(e))
                return
    if "cc_data" not in s:
        return

    es_ventas = s["cc_data"]["es_ventas"]
    st.success(f"{len(_rows())} comprobantes de {'ventas' if es_ventas else 'compras'} · {s['cc_nombre']}")

    # datos del informe
    d1, d2, d3 = st.columns(3)
    s.setdefault("cc_mail", MAIL_DEFAULT)
    s.setdefault("cc_cliente", "")
    d1.text_input("Cliente (razón social)", key="cc_cliente")
    d2.text_input("Período", key="cc_periodo")
    d3.text_input("Mail que figura en el informe", key="cc_mail")

    # búsqueda
    st.text_input("Número de comprobante", key="cc_busq", on_change=_buscar,
                  placeholder="Tipeá el número y apretá Enter (ej.: 31145 o 2-31145)")
    msg = s.get("cc_msg")
    if msg:
        {"success": st.success, "warning": st.warning, "error": st.error, "info": st.info}[msg[0]](msg[1])
    for i in s.get("cc_opciones", []):
        st.button(_desc(_rows()[i]), key=f"cc_op_{i}", on_click=_marcar, args=(i,))

    # contadores + masivos
    n_t, n_f = len(s["cc_tengo"]), len(_rows()) - len(s["cc_tengo"])
    k1, k2, k3, k4 = st.columns([2, 2, 2, 2])
    k1.markdown(f"<div class='cc-cont'><span>{n_t}</span> tengo / <span>{n_f}</span> faltan</div>",
                unsafe_allow_html=True)
    k2.button("Marcar todos como Tengo", key="cc_btn_marcar_todos", on_click=_marcar_todos,
              **_STRETCH)
    k3.checkbox("Confirmo desmarcar todos", key="cc_conf_desm")
    k4.button("Desmarcar todos", key="cc_btn_desmarcar", on_click=_desmarcar_todos,
              disabled=not s.get("cc_conf_desm"), **_STRETCH)

    pend = _pendientes_revisar()
    tab_c, tab_p = st.tabs(["Comprobantes", f"Percepciones ({pend} por revisar)" if pend else "Percepciones"])
    with tab_c:
        _tabla_comprobantes()
    with tab_p:
        _tabla_percepciones()

    st.divider()
    st.subheader("Descargas")
    _descargas()

    st.divider()
    st.subheader("Guardar / retomar avance")
    g1, g2 = st.columns(2)
    g1.download_button("Guardar avance", _avance_json(),
                       file_name=re.sub(r"\.csv$", "", s["cc_nombre"], flags=re.I) + "_avance.json",
                       mime="application/json", key="cc_dl_avance", **_STRETCH)
    with g2:
        av = st.file_uploader("Retomar avance (JSON)", type=["json"], key="cc_up_avance")
        st.button("Aplicar avance", key="cc_btn_avance", on_click=_cb_aplicar_avance, disabled=av is None)
        m = s.get("cc_msg_avance")
        if m:
            (st.success if m[0] else st.error)(m[1])
