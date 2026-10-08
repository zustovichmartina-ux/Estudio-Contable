"""Lectura de comprobantes en PDF (facturas electrónicas) para el Control de comprobantes.

Cada PDF se identifica por el código QR de ARCA (CUIT emisor, tipo, punto de venta, número,
importe). Si no tiene QR legible, se intenta leer el texto del PDF. Después se cruza con
las filas del CSV de Mis Comprobantes.
"""
from __future__ import annotations

import base64
import io
import json
import re
import zipfile
from decimal import Decimal, InvalidOperation
from urllib.parse import parse_qs, urlparse

MAX_PDFS = 600
MAX_BYTES_ZIP_ITEM = 25 * 1024 * 1024
TOLERANCIA_IMPORTE = Decimal("1.00")


# ---------------------------------------------------------------- lectura
def _a_decimal(txt) -> Decimal | None:
    try:
        return Decimal(str(txt))
    except (InvalidOperation, ValueError):
        return None


def _num_ar(txt: str) -> Decimal | None:
    t = txt.strip().replace("$", "").replace(" ", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    return _a_decimal(t)


def parsear_qr(contenido: str) -> dict | None:
    """Contenido del QR de ARCA: https://www.afip.gob.ar/fe/qr/?p=<base64 de un JSON>."""
    try:
        q = parse_qs(urlparse(contenido.strip()).query)
        p = (q.get("p") or [""])[0]
        if not p:
            return None
        p = p.replace(" ", "+")
        p += "=" * (-len(p) % 4)
        try:
            raw = base64.b64decode(p)
        except Exception:
            raw = base64.urlsafe_b64decode(p)
        j = json.loads(raw.decode("utf-8"))
        importe = _a_decimal(j.get("importe"))
        ctz = _a_decimal(j.get("ctz")) or Decimal(1)
        moneda = str(j.get("moneda") or "PES")
        total = importe * ctz if (importe is not None and moneda != "PES") else importe
        return {
            "origen": "QR",
            "cuit": re.sub(r"\D", "", str(j.get("cuit", ""))),
            "tipo": int(j.get("tipoCmp")),
            "pv": int(j.get("ptoVta")),
            "num": int(j.get("nroCmp")),
            "importe": total,
            "fecha": str(j.get("fecha", "")),
        }
    except Exception:
        return None


def _decodificar_qr_imagenes(doc) -> dict | None:
    try:
        import cv2
        import numpy as np
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
    except Exception:
        return None
    det = cv2.QRCodeDetector()

    def intentar(img) -> dict | None:
        try:
            ok, textos, _, _ = det.detectAndDecodeMulti(img)
            cand = list(textos) if ok else []
        except Exception:
            cand = []
        if not cand:
            try:
                t, _, _ = det.detectAndDecode(img)
                cand = [t]
            except Exception:
                cand = []
        for t in cand:
            if t and "afip" in t.lower():
                r = parsear_qr(t)
                if r:
                    return r
        return None

    for pn in range(min(len(doc), 3)):
        page = doc[pn]
        for zoom, parte in ((3.0, None), (4.5, "abajo"), (4.5, "arriba")):
            rect = page.rect
            clip = None
            if parte == "abajo":
                clip = fitz.Rect(rect.x0, rect.y0 + rect.height * 0.55, rect.x1, rect.y1)
            elif parte == "arriba":
                clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.45)
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False,
                                  colorspace=fitz.csGRAY)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
            r = intentar(img)
            if r:
                return r
    return None


_TIPO_TXT = (
    ("nota de credito", {"A": 3, "B": 8, "C": 13, "M": 53}),
    ("nota de debito", {"A": 2, "B": 7, "C": 12, "M": 52}),
    ("factura", {"A": 1, "B": 6, "C": 11, "M": 51, "E": 19}),
)


def _norm(t: str) -> str:
    t = t.lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        t = t.replace(a, b)
    return t


def _desde_texto(doc) -> dict | None:
    texto = "\n".join(doc[i].get_text() for i in range(min(len(doc), 2)))
    n = _norm(texto)
    m_pv = re.search(r"punto de venta[:\s]*0*(\d{1,5})", n)
    m_num = re.search(r"comp\.?\s*(?:nro|n[°o])\.?[:\s]*0*(\d{1,8})", n)
    if not (m_pv and m_num):
        m = re.search(r"\b0*(\d{1,5})\s*-\s*0*(\d{1,8})\b", n)
        if not m:
            return None
        pv, num = int(m.group(1)), int(m.group(2))
    else:
        pv, num = int(m_pv.group(1)), int(m_num.group(1))
    cuits = re.findall(r"cuit[:\s]*(\d{2}-?\d{8}-?\d)", n)
    cuit = re.sub(r"\D", "", cuits[0]) if cuits else ""
    tipo = None
    m_cod = re.search(r"cod\.?\s*0*(\d{1,3})\b", n)
    if m_cod:
        tipo = int(m_cod.group(1))
    if tipo is None:
        for nombre, letras in _TIPO_TXT:
            if nombre in n:
                ml = re.search(nombre + r"\s*\n?\s*([abcme])\b", n)
                if ml:
                    tipo = letras.get(ml.group(1).upper())
                break
    importe = None
    m_imp = re.search(r"importe total[:\s]*\$?\s*([\d.,]+)", n)
    if m_imp:
        importe = _num_ar(m_imp.group(1))
    if tipo is None:
        return None
    return {"origen": "Texto", "cuit": cuit, "tipo": tipo, "pv": pv, "num": num,
            "importe": importe, "fecha": ""}


def leer_pdf(data: bytes) -> dict | None:
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception:
        return None
    try:
        r = _decodificar_qr_imagenes(doc)
        if r:
            return r
        return _desde_texto(doc)
    except Exception:
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def expandir_archivos(archivos: list[tuple[str, bytes]]) -> list[tuple[str, bytes]]:
    """PDF sueltos y/o .zip -> lista plana de (nombre, bytes) de PDF."""
    out: list[tuple[str, bytes]] = []
    for nombre, data in archivos:
        if nombre.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        if info.is_dir() or not info.filename.lower().endswith(".pdf"):
                            continue
                        if info.file_size > MAX_BYTES_ZIP_ITEM:
                            continue
                        out.append((f"{nombre}/{info.filename}", z.read(info)))
            except zipfile.BadZipFile:
                out.append((nombre, b""))
        else:
            out.append((nombre, data))
    return out[:MAX_PDFS]


# ---------------------------------------------------------------- cruce
def _solo_digitos(t) -> str:
    return re.sub(r"\D", "", str(t or ""))


def cruzar(rows: list[dict], archivos: list[tuple[str, bytes]], es_ventas: bool,
           es_b) -> dict:
    """Lee los PDF y los cruza con el CSV.

    Devuelve {"marcar": set[int], "detalle": list[dict]}. Estados:
    OK, Importe distinto, Ya leído (duplicado), No figura en el CSV,
    Ambiguo, Comprobante B (se desestima), No se pudo leer.
    """
    pdfs = expandir_archivos(archivos)
    por_pv_num: dict[tuple[int, int], list[dict]] = {}
    for r in rows:
        por_pv_num.setdefault((r["pv"], r["num"]), []).append(r)
    marcar: set[int] = set()
    vistos: dict[int, str] = {}
    detalle: list[dict] = []
    for nombre, data in pdfs:
        corto = nombre.split("/")[-1]
        info = leer_pdf(data) if data else None
        if not info:
            detalle.append({"Archivo": corto, "Estado": "No se pudo leer", "Comprobante": "", "Detalle":
                            "Marcalo a mano con el buscador"})
            continue
        etiqueta = f"{info['pv']}-{info['num']}"
        cands = por_pv_num.get((info["pv"], info["num"]), [])
        if not es_ventas and info["cuit"]:
            c2 = [r for r in cands if _solo_digitos(r["cuit"]) == info["cuit"]]
            cands = c2
        if len(cands) > 1:
            c3 = [r for r in cands if r["tipo"] == info["tipo"]]
            if c3:
                cands = c3
        if not cands:
            detalle.append({"Archivo": corto, "Estado": "No figura en el CSV", "Comprobante": etiqueta,
                            "Detalle": "Puede ser de otro período o no estar en Mis Comprobantes"})
            continue
        if len(cands) > 1:
            detalle.append({"Archivo": corto, "Estado": "Ambiguo", "Comprobante": etiqueta,
                            "Detalle": f"Coincide con {len(cands)} filas; marcalo a mano"})
            continue
        r = cands[0]
        if es_b(r):
            detalle.append({"Archivo": corto, "Estado": "Comprobante B (se desestima)", "Comprobante": etiqueta,
                            "Detalle": r["nombre"]})
            continue
        if r["i"] in vistos:
            detalle.append({"Archivo": corto, "Estado": "Ya leído (duplicado)", "Comprobante": etiqueta,
                            "Detalle": f"Mismo comprobante que {vistos[r['i']]}"})
            continue
        vistos[r["i"]] = corto
        imp = info.get("importe")
        if imp is not None and abs(abs(imp) - abs(r["total"])) > TOLERANCIA_IMPORTE:
            detalle.append({"Archivo": corto, "Estado": "Importe distinto", "Comprobante": etiqueta,
                            "Detalle": f"PDF {imp} vs CSV {r['total']} · {r['nombre']}"})
            continue
        marcar.add(r["i"])
        detalle.append({"Archivo": corto, "Estado": "OK", "Comprobante": etiqueta,
                        "Detalle": f"{r['nombre']} (leído por {info['origen']})"})
    return {"marcar": marcar, "detalle": detalle}
