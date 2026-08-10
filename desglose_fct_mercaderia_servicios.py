# -*- coding: utf-8 -*-
"""
Desglose de facturas de venta: ítems mercadería vs servicios.

Entrada: PDFs de FCT + lista de conceptos (mercadería / servicios).
Salida Excel (formato estudio) con columnas:
  Detalle de Items Vendidos | Fecha | Comprobante | Total Importe | Total IVA | Total Bruto
"""
from __future__ import annotations

import io
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Sequence

import pandas as pd

from excel_formato_estudio import exportar_informe_excel, guardar_informe_excel

COLUMNAS = [
    "Detalle de Items Vendidos",
    "Fecha",
    "Comprobante",
    "Total Importe",
    "Total IVA",
    "Total Bruto",
]

RE_MONEY = re.compile(
    r"(?<![\d.,])(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+(?:,\d{2}))(?![\d])"
)


def _norm(texto: str) -> str:
    if not texto:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper()
    t = re.sub(r"[^A-Z0-9\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def parse_ar_money(s: Any) -> float:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return 0.0
    if isinstance(s, (int, float)):
        return round(float(s), 2)
    raw = str(s).strip().replace("$", "").replace(" ", "")
    if not raw or raw.lower() in {"nan", "none", "-"}:
        return 0.0
    neg = raw.startswith("-")
    raw = raw.lstrip("-")
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    try:
        v = float(raw)
    except ValueError:
        return 0.0
    return round(-v if neg else v, 2)


def detectar_tipo_factura(texto: str) -> str:
    t = _norm(texto)
    # COD. 01 / 011 / letra grande
    if re.search(r"\bCOD\.?\s*0*1\b", t) or re.search(r"\bFACTURA\s+A\b", t) or re.search(r"\bCOD\.?\s*01\b", t):
        if "COD 011" in t or "COD. 011" in t or re.search(r"\bCOD\.?\s*011\b", t):
            return "Factura C"
        return "Factura A"
    if re.search(r"\bCOD\.?\s*0*6\b", t) or re.search(r"\bFACTURA\s+B\b", t):
        return "Factura B"
    if re.search(r"\bCOD\.?\s*011\b", t) or re.search(r"\bFACTURA\s+C\b", t):
        return "Factura C"
    # Letra suelta en encabezado típico
    if re.search(r"\bA\b.*\bCOD", texto[:800], re.I) or re.search(r"ORIGINAL[\s\S]{0,80}\bA\b", texto[:500], re.I):
        return "Factura A"
    if re.search(r"\bB\b.*\bCOD", texto[:800], re.I):
        return "Factura B"
    return "Factura"


def extraer_fecha(texto: str) -> str:
    m = re.search(
        r"Fecha\s*(?:de\s*)?Emisi[oó]n\s*:?\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.I,
    )
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto[:1500])
    return m.group(1) if m else ""


def extraer_comprobante(texto: str, tipo: str) -> str:
    pv = nro = ""
    m = re.search(
        r"Punto\s*de\s*Venta\s*:?\s*(\d+)\s*.{0,40}?Comp\.?\s*Nro\.?\s*:?\s*(\d+)",
        texto,
        re.I | re.S,
    )
    if m:
        pv, nro = m.group(1), m.group(2)
    else:
        m = re.search(r"\b(\d{4,5})\s*[-/]\s*(\d{6,8})\b", texto[:2000])
        if m:
            pv, nro = m.group(1), m.group(2)
    letra = "A" if "A" in tipo.upper() else ("B" if "B" in tipo.upper() else ("C" if "C" in tipo.upper() else ""))
    if pv and nro:
        try:
            base = f"{int(pv):05d}-{int(nro):08d}"
        except ValueError:
            base = f"{pv}-{nro}"
        return f"{letra} {base}".strip() if letra else base
    return letra or ""


def totales_documento(texto: str) -> tuple[float, float, float]:
    """Retorna (neto, iva, bruto) a nivel documento."""
    neto = iva = bruto = 0.0

    def _buscar(*labels: str) -> float:
        for lab in labels:
            m = re.search(
                lab + r"\s*:?\s*\$?\s*(" + RE_MONEY.pattern + r")",
                texto,
                re.I,
            )
            if m:
                return parse_ar_money(m.group(1))
        return 0.0

    neto = _buscar(
        r"Importe\s+Neto\s+Gravado",
        r"Neto\s+Gravado",
        r"Subtotal",
    )
    iva = _buscar(
        r"IVA\s*21\s*%",
        r"Importe\s+IVA",
        r"IVA",
    )
    bruto = _buscar(
        r"Importe\s+Total",
        r"Total\s+a\s+Pagar",
        r"Total",
    )
    if bruto <= 0 and neto > 0:
        bruto = round(neto + iva, 2)
    if neto <= 0 and bruto > 0 and iva > 0:
        neto = round(bruto - iva, 2)
    if iva <= 0 and neto > 0 and bruto > neto:
        iva = round(bruto - neto, 2)
    return neto, iva, bruto


def _parece_linea_item(concepto: str) -> bool:
    c = _norm(concepto)
    if len(c) < 3:
        return False
    bloqueados = (
        "SUBTOTAL", "IMPORTE NETO", "IMPORTE TOTAL", "IVA ", "CAE", "PAGINA",
        "CODIGO PRODUCTO", "PRODUCTO SERVICIO", "CANTIDAD", "PRECIO UNIT",
        "FECHA DE", "CUIT", "RAZON SOCIAL", "CONDICION FRENTE", "PERIODO FACTURADO",
        "OTROS TRIBUTOS", "BONIF", "COMPROBANTE AUTORIZADO",
    )
    return not any(b in c for b in bloqueados)


def extraer_items_lineas(texto: str) -> list[dict]:
    """
    Extrae ítems con subtotal de línea.
    Heurística: líneas con descripción + monto al final (tabla AFIP).
    """
    items: list[dict] = []
    # Cortar pie (CAE / totales globales) para no mezclar
    cuerpo = texto
    for corte in ("CAE N", "C.A.E", "Importe Neto Gravado", "Importe Total", "Comprobante Autorizado"):
        idx = re.search(corte, cuerpo, re.I)
        if idx and idx.start() > 200:
            cuerpo = cuerpo[: idx.start()]
            break

    for raw in cuerpo.splitlines():
        line = raw.strip()
        if not line or len(line) < 5:
            continue
        montos = RE_MONEY.findall(line)
        if not montos:
            continue
        # Último monto = subtotal de la línea
        subtotal = parse_ar_money(montos[-1])
        if subtotal <= 0:
            continue
        # Quitar montos del final para dejar el concepto
        concepto = line
        for mon in reversed(montos[-3:]):
            pos = concepto.rfind(mon)
            if pos >= 0:
                concepto = concepto[:pos]
        concepto = re.sub(r"\s+", " ", concepto).strip(" -|\t")
        # Sacar código numérico inicial típico (002, 001, etc.)
        concepto = re.sub(r"^\d{1,6}\s+", "", concepto).strip()
        if not _parece_linea_item(concepto):
            continue
        # Evitar líneas que son solo un monto
        if len(_norm(concepto)) < 4:
            continue
        items.append({"concepto": concepto, "subtotal": subtotal})

    # Dedup consecutivos iguales
    out: list[dict] = []
    for it in items:
        if out and out[-1]["concepto"] == it["concepto"] and abs(out[-1]["subtotal"] - it["subtotal"]) < 0.02:
            continue
        out.append(it)
    return out


def _leer_conceptos_de_archivo(fuente: Any) -> list[str]:
    """Lee una lista de conceptos desde Excel/CSV/TXT (primera columna útil)."""
    if fuente is None:
        return []

    nombre = str(getattr(fuente, "name", "") or "").lower()
    data = fuente.read() if hasattr(fuente, "read") else fuente
    if hasattr(fuente, "seek"):
        try:
            fuente.seek(0)
        except Exception:
            pass

    if isinstance(data, str):
        data = data.encode("utf-8")
    if not isinstance(data, (bytes, bytearray)):
        return []

    # TXT: una línea = un concepto (ignora prefijos M:/S: si vienen)
    if nombre.endswith(".txt") or (not nombre.endswith((".xlsx", ".xls", ".csv")) and b"\x00" not in data[:20]):
        text = data.decode("utf-8", errors="replace")
        out: list[str] = []
        for line in text.splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            s = re.sub(r"^(m:|s:|mercader[ií]a|servicios?)\s*:?\s*", "", s, flags=re.I).strip()
            if s:
                out.append(s)
        return out

    bio = io.BytesIO(bytes(data))
    if nombre.endswith(".csv"):
        df = pd.read_csv(bio)
    else:
        df = pd.read_excel(bio)

    cols = {_norm(c): c for c in df.columns}
    # Preferir columna con nombre explícito; si no, la primera
    prefer = next(
        (
            cols[k]
            for k in cols
            if k in {"CONCEPTO", "ITEM", "DETALLE", "PRODUCTO", "NOMBRE", "SERVICIO", "SERVICIOS", "MERCADERIA", "MERCADERIA"}
            or "CONCEPTO" in k
            or "PRODUCTO" in k
            or "SERVIC" in k
            or "MERC" in k
            or "DETALLE" in k
        ),
        None,
    )
    col = prefer or df.columns[0]
    return [str(x).strip() for x in df[col].dropna().tolist() if str(x).strip() and str(x).strip().lower() != "nan"]


def cargar_lista_clasificacion(
    fuente: Any = None,
    *,
    fuente_mercaderia: Any = None,
    fuente_servicios: Any = None,
) -> tuple[list[str], list[str]]:
    """
    Dos archivos separados (recomendado) o uno combinado (legacy).

    - fuente_mercaderia / fuente_servicios: Excel/CSV/TXT con conceptos
    - fuente: Excel con columnas Mercadería|Servicios, Concepto+Tipo, o TXT M:/S:
    """
    merc: list[str] = []
    serv: list[str] = []

    if fuente_mercaderia is not None:
        merc = _leer_conceptos_de_archivo(fuente_mercaderia)
    if fuente_servicios is not None:
        serv = _leer_conceptos_de_archivo(fuente_servicios)

    if merc or serv:
        return merc, serv

    if fuente is None:
        return merc, serv

    nombre = str(getattr(fuente, "name", "") or "").lower()
    data = fuente.read() if hasattr(fuente, "read") else fuente
    if hasattr(fuente, "seek"):
        try:
            fuente.seek(0)
        except Exception:
            pass

    if isinstance(data, str):
        data = data.encode("utf-8")

    bio = io.BytesIO(data if isinstance(data, (bytes, bytearray)) else bytes(data))

    # TXT combinado
    if nombre.endswith(".txt") or (not nombre.endswith((".xlsx", ".xls", ".csv")) and b"\x00" not in data[:20]):
        try:
            text = data.decode("utf-8", errors="replace") if isinstance(data, (bytes, bytearray)) else str(data)
        except Exception:
            text = str(data)
        for line in text.splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            low = s.lower()
            if low.startswith(("m:", "merc", "mercaderia", "mercadería")):
                merc.append(re.sub(r"^(m:|mercader[ií]a)\s*:?\s*", "", s, flags=re.I).strip())
            elif low.startswith(("s:", "serv", "servicio")):
                serv.append(re.sub(r"^(s:|servicios?)\s*:?\s*", "", s, flags=re.I).strip())
        return [x for x in merc if x], [x for x in serv if x]

    # Excel / CSV combinado
    bio.seek(0)
    if nombre.endswith(".csv"):
        df = pd.read_csv(bio)
    else:
        df = pd.read_excel(bio)

    cols = {_norm(c): c for c in df.columns}
    col_m = next((cols[k] for k in cols if "MERC" in k), None)
    col_s = next((cols[k] for k in cols if "SERV" in k), None)
    col_tipo = next((cols[k] for k in cols if k in {"TIPO", "CLASE", "RUBRO"}), None)
    col_conc = next((cols[k] for k in cols if k in {"CONCEPTO", "ITEM", "DETALLE", "PRODUCTO", "NOMBRE"}), None)

    if col_m or col_s:
        if col_m:
            merc = [str(x).strip() for x in df[col_m].dropna().tolist() if str(x).strip()]
        if col_s:
            serv = [str(x).strip() for x in df[col_s].dropna().tolist() if str(x).strip()]
    elif col_tipo and col_conc:
        for _, row in df.iterrows():
            tip = _norm(str(row[col_tipo]))
            conc = str(row[col_conc]).strip()
            if not conc:
                continue
            if "MERC" in tip:
                merc.append(conc)
            elif "SERV" in tip:
                serv.append(conc)
    else:
        if len(df.columns) >= 2:
            c0, c1 = df.columns[0], df.columns[1]
            for _, row in df.iterrows():
                conc = str(row[c0]).strip()
                tip = _norm(str(row[c1]))
                if not conc or conc.lower() == "nan":
                    continue
                if "MERC" in tip:
                    merc.append(conc)
                elif "SERV" in tip:
                    serv.append(conc)

    return merc, serv


def clasificar_concepto(concepto: str, mercaderia: Sequence[str], servicios: Sequence[str]) -> str:
    """Devuelve Mercadería | Servicio | Sin clasificar (fuzzy simple)."""
    try:
        from rapidfuzz import fuzz
    except Exception:
        fuzz = None

    nc = _norm(concepto)
    best_m, score_m = "", 0.0
    best_s, score_s = "", 0.0

    for m in mercaderia:
        nm = _norm(m)
        if not nm:
            continue
        if nm in nc or nc in nm:
            sc = 100.0
        elif fuzz:
            sc = float(fuzz.token_set_ratio(nc, nm))
        else:
            sc = 0.0
        if sc > score_m:
            score_m, best_m = sc, m

    for s in servicios:
        ns = _norm(s)
        if not ns:
            continue
        if ns in nc or nc in ns:
            sc = 100.0
        elif fuzz:
            sc = float(fuzz.token_set_ratio(nc, ns))
        else:
            sc = 0.0
        if sc > score_s:
            score_s, best_s = sc, s

    if score_m >= 60 and score_m >= score_s:
        return "Mercadería"
    if score_s >= 60 and score_s > score_m:
        return "Servicio"
    # keywords genéricos
    if any(k in nc for k in ("ABONO", "HONORARIO", "SERVICIO", "ASESOR", "MANTENIMIENTO", "ALQUILER")):
        return "Servicio"
    if any(k in nc for k in ("PRODUCTO", "MERCADER", "ARTICULO", "UNIDAD", "REPUESTO")):
        return "Mercadería"
    return "Sin clasificar"


def procesar_pdf_factura(
    nombre: str,
    data: bytes,
    *,
    mercaderia: Sequence[str],
    servicios: Sequence[str],
    usar_ocr: bool = True,
) -> tuple[list[dict], dict]:
    from cruce_facturas_arca import extraer_texto_comprobante

    texto, metodo = extraer_texto_comprobante(nombre, data, usar_ocr=usar_ocr)
    tipo = detectar_tipo_factura(texto)
    fecha = extraer_fecha(texto)
    comprobante = extraer_comprobante(texto, tipo)
    neto_doc, iva_doc, bruto_doc = totales_documento(texto)
    items = extraer_items_lineas(texto)

    filas: list[dict] = []
    suma_sub = sum(i["subtotal"] for i in items) or 0.0

    for it in items:
        sub = float(it["subtotal"])
        # Prorrateo de IVA documento según peso del ítem
        if suma_sub > 0 and iva_doc > 0:
            iva_linea = round(iva_doc * (sub / suma_sub), 2)
        elif tipo == "Factura A" and sub > 0:
            iva_linea = round(sub * 0.21, 2)
        else:
            iva_linea = 0.0

        if tipo == "Factura B" and iva_doc <= 0 and bruto_doc > 0 and suma_sub > 0:
            # B sin IVA discriminado: el subtotal de línea se toma como bruto
            bruto_l = sub
            # si el doc tiene solo total, neto=bruto iva=0
            neto_l = sub
            iva_l = 0.0
        else:
            neto_l = sub
            iva_l = iva_linea
            bruto_l = round(neto_l + iva_l, 2)

        clase = clasificar_concepto(it["concepto"], mercaderia, servicios)
        filas.append(
            {
                "Detalle de Items Vendidos": it["concepto"],
                "Fecha": fecha,
                "Comprobante": comprobante,
                "Total Importe": neto_l,
                "Total IVA": iva_l,
                "Total Bruto": bruto_l,
                "Clase": clase,
                "Tipo FCT": tipo,
                "Archivo": nombre,
            }
        )

    meta = {
        "archivo": nombre,
        "metodo": metodo,
        "tipo": tipo,
        "fecha": fecha,
        "comprobante": comprobante,
        "items": len(filas),
        "neto_doc": neto_doc,
        "iva_doc": iva_doc,
        "bruto_doc": bruto_doc,
    }
    return filas, meta


def procesar_lote(
    archivos: Iterable[Any],
    lista_clasificacion: Any = None,
    *,
    lista_mercaderia: Any = None,
    lista_servicios: Any = None,
    usar_ocr: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, list[dict]]:
    merc, serv = cargar_lista_clasificacion(
        lista_clasificacion,
        fuente_mercaderia=lista_mercaderia,
        fuente_servicios=lista_servicios,
    )
    todas: list[dict] = []
    metas: list[dict] = []
    errores: list[dict] = []

    for up in archivos or []:
        nombre = str(getattr(up, "name", "factura.pdf"))
        data = up.read() if hasattr(up, "read") else up
        if hasattr(up, "seek"):
            try:
                up.seek(0)
            except Exception:
                pass
        if not isinstance(data, (bytes, bytearray)):
            continue
        try:
            filas, meta = procesar_pdf_factura(
                nombre, bytes(data), mercaderia=merc, servicios=serv, usar_ocr=usar_ocr
            )
            metas.append(meta)
            if not filas:
                errores.append({"archivo": nombre, "motivo": "No se detectaron ítems de detalle"})
            todas.extend(filas)
        except Exception as exc:
            errores.append({"archivo": nombre, "motivo": str(exc)})

    df = pd.DataFrame(todas)
    if df.empty:
        vacio = pd.DataFrame(columns=COLUMNAS)
        return vacio, vacio.copy(), vacio.copy(), vacio.copy(), errores

    detalle = df[COLUMNAS].copy()
    merc_df = df.loc[df["Clase"] == "Mercadería", COLUMNAS].copy()
    serv_df = df.loc[df["Clase"] == "Servicio", COLUMNAS].copy()
    sin_df = df.loc[df["Clase"] == "Sin clasificar", COLUMNAS].copy()
    return detalle, merc_df, serv_df, sin_df, errores


def exportar_excel_bytes(
    detalle: pd.DataFrame,
    mercaderia: pd.DataFrame,
    servicios: pd.DataFrame,
    sin_clasificar: pd.DataFrame,
    *,
    titulo: str = "Desglose FCT — Mercadería y Servicios",
    subtitulo: str = "",
) -> bytes:
    resumen_rows = []
    for nombre, frame in (
        ("Mercadería", mercaderia),
        ("Servicios", servicios),
        ("Sin clasificar", sin_clasificar),
        ("Total", detalle),
    ):
        if frame is None or frame.empty:
            resumen_rows.append(
                {"Rubro": nombre, "Ítems": 0, "Total Importe": 0.0, "Total IVA": 0.0, "Total Bruto": 0.0}
            )
        else:
            resumen_rows.append(
                {
                    "Rubro": nombre,
                    "Ítems": len(frame),
                    "Total Importe": float(frame["Total Importe"].sum()),
                    "Total IVA": float(frame["Total IVA"].sum()),
                    "Total Bruto": float(frame["Total Bruto"].sum()),
                }
            )
    resumen = pd.DataFrame(resumen_rows)
    return exportar_informe_excel(
        titulo=titulo,
        subtitulo=subtitulo or "Ítems de facturas de venta clasificados",
        periodo=date.today().strftime("%d/%m/%Y"),
        kpis=[
            ("Ítems", len(detalle)),
            ("Bruto total", round(float(detalle["Total Bruto"].sum()) if not detalle.empty else 0, 2)),
            ("Mercadería", len(mercaderia)),
            ("Servicios", len(servicios)),
        ],
        resumenes=[("Totales por rubro", resumen)],
        detalle=detalle,
        hoja_detalle="Lista",
        hojas_adicionales=[
            ("Mercadería", mercaderia if not mercaderia.empty else pd.DataFrame(columns=COLUMNAS)),
            ("Servicios", servicios if not servicios.empty else pd.DataFrame(columns=COLUMNAS)),
            ("Sin clasificar", sin_clasificar if not sin_clasificar.empty else pd.DataFrame(columns=COLUMNAS)),
        ],
        col_moneda=["Total Importe", "Total IVA", "Total Bruto"],
        col_fecha=["Fecha"],
        total_col="Total Bruto",
    )


def guardar_excel(
    ruta: str | Path,
    detalle: pd.DataFrame,
    mercaderia: pd.DataFrame,
    servicios: pd.DataFrame,
    sin_clasificar: pd.DataFrame,
    **kwargs: Any,
) -> Path:
    data = exportar_excel_bytes(detalle, mercaderia, servicios, sin_clasificar, **kwargs)
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_bytes(data)
    return ruta
