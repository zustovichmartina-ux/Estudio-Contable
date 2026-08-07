"""
Cruce de facturas (PDF/fotos) vs listado Mis Comprobantes ARCA (portal IVA).

Resultado estilo papel de trabajo:
  - Resumen
  - Matcheadas
  - A revisar (en facturas, no en ARCA / receptor distinto / sin datos)
  - Faltantes (en ARCA, sin factura subida)
  - Diferencias (mismo comprobante, importe distinto)
"""

from __future__ import annotations

import hashlib
import io
import re
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, BinaryIO, Iterable, Sequence

import pandas as pd

TOL_IMPORTE = 0.05  # centavos / redondeo

COLUMNAS_MATCH = [
    "Proveedor",
    "Tipo",
    "Nº Comprobante",
    "Fecha",
    "Total",
    "CAE",
    "CUIT Emisor",
    "Archivo",
    "Estado",
]
COLUMNAS_REVISAR = [
    "Proveedor",
    "Comprobante",
    "Fecha",
    "Total",
    "Motivo / Observación",
    "CUIT Emisor",
    "CUIT Receptor",
    "Archivo",
]
COLUMNAS_FALTANTES = [
    "Fecha",
    "Proveedor",
    "Tipo",
    "Comprobante",
    "Total",
    "CUIT Emisor",
    "CAE",
]
COLUMNAS_DIF = [
    "Proveedor",
    "Tipo",
    "Nº Comprobante",
    "Fecha",
    "Total Factura",
    "Total ARCA",
    "Diferencia",
    "CAE",
    "Archivo",
]


# ── Normalización ─────────────────────────────────────────────────────────────

def normalizar_cuit(valor: Any) -> str:
    digitos = re.sub(r"\D", "", str(valor or ""))
    return digitos if len(digitos) == 11 else digitos


def normalizar_cae(valor: Any) -> str:
    digitos = re.sub(r"\D", "", str(valor or ""))
    if not digitos or digitos in {"0", "00"}:
        return ""
    return digitos


def _solo_digitos(valor: Any) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def formatear_comprobante(pv: Any, nro: Any) -> str:
    """Formato canónico PV(5)-NRO(8)."""
    pv_d = _solo_digitos(pv)
    nro_d = _solo_digitos(nro)
    if not pv_d or not nro_d:
        return ""
    try:
        return f"{int(pv_d):05d}-{int(nro_d):08d}"
    except ValueError:
        return ""


def parsear_comprobante_texto(texto: Any) -> tuple[str, str, str]:
    """
    Extrae (pv, nro, comprobante_fmt) desde textos varios:
      00014-00082908 | 14-82908 | IPE-7133-A-00031153 | 7133-00031285
    """
    raw = str(texto or "").strip()
    if not raw or raw.lower() in {"nan", "none"}:
        return "", "", ""

    # Prefijos tipo IPE-7133-A-00031153
    m = re.search(
        r"(?:^|[^\d])(\d{1,5})\s*[-/]\s*[A-Z]?\s*[-/]?\s*(\d{1,8})\b",
        raw,
        flags=re.IGNORECASE,
    )
    if not m:
        m = re.search(r"\b(\d{1,5})\s*[-/]\s*(\d{1,8})\b", raw)
    if m:
        cmpte = formatear_comprobante(m.group(1), m.group(2))
        return m.group(1), m.group(2), cmpte

    # Columna ARCA nueva: "00014 00082908" o pegados
    m2 = re.search(r"\b(\d{1,5})\s+(\d{4,8})\b", raw)
    if m2:
        cmpte = formatear_comprobante(m2.group(1), m2.group(2))
        return m2.group(1), m2.group(2), cmpte

    return "", "", ""


def _parse_monto(valor: Any) -> float:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    from procesador import _limpiar_monto

    return float(_limpiar_monto(valor) or 0.0)


def _fmt_fecha(valor: Any) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    if hasattr(valor, "to_pydatetime"):
        try:
            return valor.to_pydatetime().strftime("%d/%m/%Y")
        except Exception:
            pass
    txt = str(valor).strip()
    if not txt or txt.lower() in {"nan", "nat", "none"}:
        return ""
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(txt[:10], fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    try:
        dt = pd.to_datetime(txt, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return txt


def _tipo_legible(tipo_raw: Any) -> str:
    t = str(tipo_raw or "").strip()
    if not t or t.lower() in {"nan", "none"}:
        return ""
    # "1 - Factura A" → "Factura A"
    if " - " in t:
        t = t.split(" - ", 1)[1].strip()
    return re.sub(r"\s+", " ", t)


def _norm_nombre(s: Any) -> str:
    t = str(s or "").upper()
    t = re.sub(r"[ÁÀÄÂ]", "A", t)
    t = re.sub(r"[ÉÈËÊ]", "E", t)
    t = re.sub(r"[ÍÌÏÎ]", "I", t)
    t = re.sub(r"[ÓÒÖÔ]", "O", t)
    t = re.sub(r"[ÚÙÜÛ]", "U", t)
    t = re.sub(r"[^A-Z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _clave_match(
    *,
    cae: str = "",
    cuit_emisor: str = "",
    comprobante: str = "",
) -> str:
    cae_n = normalizar_cae(cae)
    if cae_n and len(cae_n) >= 10:
        return f"cae:{cae_n}"
    cmpte = str(comprobante or "").strip()
    cuit = normalizar_cuit(cuit_emisor)
    if cmpte and cuit:
        return f"cuit+cmp:{cuit}|{cmpte}"
    if cmpte:
        return f"cmp:{cmpte}"
    return ""


# ── Lectura listado ARCA ──────────────────────────────────────────────────────

def _buscar_col(cols: Iterable[str], *needles: str) -> str | None:
    lowered = {str(c): re.sub(r"\s+", " ", str(c).lower().strip()) for c in cols}
    for needle in needles:
        n = needle.lower()
        for original, low in lowered.items():
            if n in low:
                return original
    return None


def _detectar_header_arca(raw: pd.DataFrame) -> int:
    """Índice de fila con encabezados reales (Fecha / Tipo / Imp. Total, etc.)."""
    for i in range(min(25, len(raw))):
        vals = [str(v).strip().lower() for v in raw.iloc[i].tolist() if pd.notna(v)]
        joined = " | ".join(vals)
        tiene_fecha = "fecha" in joined
        tiene_tipo = "tipo" in joined
        tiene_total = "imp. total" in joined or "importe total" in joined or "total" in joined
        tiene_pv = "punto de venta" in joined or "punto venta" in joined
        tiene_nro = "número" in joined or "numero" in joined or "nro" in joined
        if tiene_fecha and (tiene_tipo or tiene_pv) and (tiene_total or tiene_nro):
            return i
    return 0


def leer_listado_arca(
    fuente: str | Path | BinaryIO | bytes,
    nombre: str | None = None,
) -> tuple[pd.DataFrame, str]:
    """
    Lee Mis Comprobantes Recibidos (xlsx/xls/csv) del portal ARCA/AFIP.

    Retorna (df_normalizado, cuit_contribuyente_detectado).
    Columnas normalizadas:
      Fecha, Tipo, Proveedor, CUIT Emisor, CUIT Receptor, Punto Venta, Número,
      Comprobante, CAE, Total
    """
    data: bytes
    fname = nombre or ""
    if isinstance(fuente, (str, Path)):
        path = Path(fuente)
        data = path.read_bytes()
        fname = fname or path.name
    elif isinstance(fuente, (bytes, bytearray)):
        data = bytes(fuente)
    else:
        if hasattr(fuente, "seek"):
            try:
                fuente.seek(0)
            except Exception:
                pass
        raw_get = getattr(fuente, "getvalue", None)
        if callable(raw_get):
            data = raw_get()
        else:
            data = fuente.read()  # type: ignore[union-attr]
        fname = fname or str(getattr(fuente, "name", "arca.xlsx"))

    cuit_titulo = ""
    m_cuit = re.search(r"CUIT\s*([0-9\- ]{11,15})", fname, flags=re.IGNORECASE)
    if m_cuit:
        cuit_titulo = normalizar_cuit(m_cuit.group(1))

    suf = Path(fname).suffix.lower()
    bio = io.BytesIO(data)

    if suf == ".csv":
        raw = None
        for sep in (";", ",", "\t"):
            bio.seek(0)
            try:
                candidate = pd.read_csv(bio, header=None, dtype=object, sep=sep, engine="python")
                if candidate.shape[1] >= 5:
                    raw = candidate
                    break
            except Exception:
                continue
        if raw is None:
            bio.seek(0)
            raw = pd.read_csv(bio, header=None, dtype=object, engine="python")
    else:
        bio.seek(0)
        try:
            raw = pd.read_excel(bio, header=None, dtype=object)
        except Exception:
            # .xls legacy a veces falla con openpyxl; reintentar con xlrd vía pandas
            bio.seek(0)
            raw = pd.read_excel(bio, header=None, dtype=object, engine="xlrd")

    if raw is None or raw.empty:
        return pd.DataFrame(columns=[
            "Fecha", "Tipo", "Proveedor", "CUIT Emisor", "CUIT Receptor",
            "Punto Venta", "Número", "Comprobante", "CAE", "Total",
        ]), cuit_titulo

    # CUIT en fila título
    for i in range(min(3, len(raw))):
        fila_txt = " ".join(str(v) for v in raw.iloc[i].tolist() if pd.notna(v))
        m = re.search(r"CUIT\s*([0-9\- ]{11,15})", fila_txt, flags=re.IGNORECASE)
        if m:
            cuit_titulo = normalizar_cuit(m.group(1)) or cuit_titulo
            break

    hdr = _detectar_header_arca(raw)
    cols = []
    seen: dict[str, int] = {}
    for j, c in enumerate(raw.iloc[hdr].tolist()):
        name = str(c).strip() if pd.notna(c) else f"col_{j}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)

    df = raw.iloc[hdr + 1 :].copy()
    df.columns = cols
    df = df.dropna(how="all")

    col_fecha = _buscar_col(df.columns, "fecha")
    col_tipo = _buscar_col(df.columns, "tipo")
    col_pv = _buscar_col(df.columns, "punto de venta", "punto venta")
    col_nro = _buscar_col(df.columns, "número desde", "numero desde", "nro. desde", "número", "numero")
    # Formato nuevo: PV y número en una sola columna
    col_pv_nro = _buscar_col(df.columns, "punto de venta-número", "punto de venta - número", "nro. de comprobante")
    col_cae = _buscar_col(df.columns, "autoriz", "cae", "cai")
    col_cuit_e = _buscar_col(df.columns, "nro. doc. emisor", "cuit emisor", "doc. emisor")
    col_denom = _buscar_col(df.columns, "denominación emisor", "denominacion emisor", "razón social", "razon social")
    col_cuit_r = _buscar_col(df.columns, "nro. doc. receptor", "cuit receptor", "doc. receptor")
    col_total = _buscar_col(df.columns, "imp. total", "importe total", "total")

    filas: list[dict] = []
    for _, row in df.iterrows():
        fecha = _fmt_fecha(row[col_fecha]) if col_fecha else ""
        tipo = _tipo_legible(row[col_tipo]) if col_tipo else ""
        proveedor = str(row[col_denom]).strip() if col_denom and pd.notna(row[col_denom]) else ""
        cuit_e = normalizar_cuit(row[col_cuit_e]) if col_cuit_e else ""
        cuit_r = normalizar_cuit(row[col_cuit_r]) if col_cuit_r else ""
        cae = normalizar_cae(row[col_cae]) if col_cae else ""
        total = _parse_monto(row[col_total]) if col_total else 0.0

        pv, nro, cmpte = "", "", ""
        if col_pv and col_nro:
            pv = _solo_digitos(row[col_pv])
            nro = _solo_digitos(row[col_nro])
            cmpte = formatear_comprobante(pv, nro)
        elif col_pv_nro:
            pv, nro, cmpte = parsear_comprobante_texto(row[col_pv_nro])
        elif col_pv:
            # a veces todo viene en PV
            pv, nro, cmpte = parsear_comprobante_texto(row[col_pv])

        # Filas basura / títulos
        if not fecha and not cmpte and not cae and not proveedor:
            continue
        if proveedor.lower() in {"denominación emisor", "denominacion emisor"}:
            continue

        filas.append({
            "Fecha": fecha,
            "Tipo": tipo,
            "Proveedor": proveedor,
            "CUIT Emisor": cuit_e,
            "CUIT Receptor": cuit_r,
            "Punto Venta": pv,
            "Número": nro,
            "Comprobante": cmpte,
            "CAE": cae,
            "Total": round(total, 2),
        })

    out = pd.DataFrame(filas)
    if not cuit_titulo and not out.empty and "CUIT Receptor" in out.columns:
        moda = (
            out["CUIT Receptor"]
            .astype(str)
            .map(normalizar_cuit)
            .replace("", pd.NA)
            .dropna()
            .mode()
        )
        if len(moda):
            cuit_titulo = str(moda.iloc[0])

    return out, cuit_titulo


# ── Extracción facturas PDF / imágenes ────────────────────────────────────────

_IMG_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}


def _ocr_imagen_bytes(img_bytes: bytes) -> str:
    """OCR lazy con easyocr del proyecto (procesador)."""
    from PIL import Image
    import numpy as np
    from procesador import _lector_ocr_streamlit_cached as _obtener_lector_ocr, _lector_ocr_run_lock

    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    arr = np.array(img)
    lector = _obtener_lector_ocr()
    if lector is None:
        return ""
    with _lector_ocr_run_lock:
        res = lector.readtext(arr, detail=0, paragraph=False)
    return "\n".join(str(x) for x in (res or []))


def _ocr_pdf_bytes(pdf_bytes: bytes, max_paginas: int = 2) -> str:
    import fitz
    import os
    from PIL import Image
    import numpy as np
    from procesador import _lector_ocr_streamlit_cached as _obtener_lector_ocr, _lector_ocr_run_lock

    # Cloud: una pagina a DPI mas bajo para no tumbar el free tier.
    cloud = any(
        str(os.environ.get(k) or "").strip().lower() in {"1", "true", "yes"}
        for k in ("STREAMLIT_SHARING_MODE", "STREAMLIT_CLOUD", "IS_STREAMLIT_CLOUD")
    ) or Path("/mount/src").is_dir()
    if cloud:
        max_paginas = min(max_paginas, 1)

    partes: list[str] = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        lector = _obtener_lector_ocr()
        if lector is None:
            return ""
        scale = 1.5 if cloud else 2
        for i in range(min(max_paginas, len(doc))):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            arr = np.array(img)
            with _lector_ocr_run_lock:
                res = lector.readtext(arr, detail=0, paragraph=False)
            partes.append("\n".join(str(x) for x in (res or [])))
    finally:
        doc.close()
    return "\n".join(partes)


def extraer_texto_comprobante(nombre: str, data: bytes, *, usar_ocr: bool = True) -> tuple[str, str]:
    """
    Extrae texto de PDF o imagen.
    Siempre intenta texto embebido (pdfplumber/pymupdf) antes que OCR.
    Retorna (texto, metodo) donde metodo es pdfplumber|ocr_pdf|ocr_img|vacio.
    """
    from procesador import extraer_texto_factura_afip

    lower = nombre.lower()
    ext = Path(lower).suffix

    if ext == ".pdf" or lower.endswith(".pdf"):
        try:
            texto = extraer_texto_factura_afip(data)
        except Exception:
            texto = ""
        if texto and len(texto.strip()) >= 40:
            return texto, "pdfplumber"
        if usar_ocr:
            try:
                ocr = _ocr_pdf_bytes(data)
                if ocr.strip():
                    return ocr, "ocr_pdf"
            except Exception as exc:
                return texto or "", f"error_ocr:{exc}"
        return texto or "", "vacio" if not (texto or "").strip() else "pdfplumber"

    if ext in _IMG_EXT:
        if not usar_ocr:
            return "", "ocr_omitido"
        try:
            out = _ocr_imagen_bytes(data)
            if not out.strip():
                return "", "ocr_vacio"
            return out, "ocr_img"
        except Exception as exc:
            return "", f"error_ocr:{exc}"

    return "", "formato_no_soportado"


def _extraer_partes_factura(texto: str) -> dict[str, Any]:
    """Campos extra para cruce (emisor/receptor) sobre el parser AFIP base."""
    from procesador import (
        detectar_tipo_comprobante_afip,
        formatear_comprobante_tango,
        _extraer_cae_afip,
        _limpiar_monto,
    )

    bloque = re.sub(r"[ \t]+", " ", texto or "")
    bloque = re.sub(r"\n{2,}", "\n", bloque)

    tipo, codigo, signo = detectar_tipo_comprobante_afip(texto)
    cae = _extraer_cae_afip(bloque)

    fecha = ""
    m_f = re.search(
        r"Fecha\s+de\s+Emisi[oó]n[:\s]*(\d{2}/\d{2}/\d{4})",
        bloque,
        flags=re.IGNORECASE,
    )
    if m_f:
        fecha = m_f.group(1)
    else:
        m_f2 = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", bloque)
        if m_f2:
            fecha = m_f2.group(1)

    pv = ""
    m_pv = re.search(r"Punto\s+de\s+Venta[:\s]*(\d{1,5})", bloque, flags=re.IGNORECASE)
    if m_pv:
        pv = m_pv.group(1)

    nro = ""
    for pat in (
        r"Comp(?:\.|\s)*Nro[:\s\.]*(\d+)",
        r"Nro\.?\s*(?:de\s+)?Comprobante[:\s]*(\d+)",
        r"N[uú]mero\s*(?:de\s+)?[Cc]omp(?:robante)?[:\s]*(\d+)",
        r"\b(\d{5})\s*[-–]\s*(\d{8})\b",
    ):
        m = re.search(pat, bloque, flags=re.IGNORECASE)
        if m:
            if m.lastindex == 2:
                pv = pv or m.group(1)
                nro = m.group(2)
            else:
                nro = m.group(1)
            break

    # También buscar patrón PV-NRO suelto tipo 00014-00082908
    if not nro:
        _, nro2, cmp_tmp = parsear_comprobante_texto(bloque)
        if cmp_tmp:
            pv2, _, _ = parsear_comprobante_texto(cmp_tmp)
            pv = pv or pv2
            nro = nro2

    importe = 0.0
    m_imp = re.search(
        r"Importe\s+Total[:\s]*\$?\s*([\d][\d.,]*)",
        bloque,
        flags=re.IGNORECASE,
    )
    if m_imp:
        importe = _limpiar_monto(m_imp.group(1))
    if importe <= 0:
        m_imp2 = re.search(
            r"(?:^|\n)\s*Total[:\s]*\$?\s*([\d][\d.,]*)",
            bloque,
            flags=re.IGNORECASE,
        )
        if m_imp2:
            importe = _limpiar_monto(m_imp2.group(1))

    # CUIT emisor / receptor: en comprobantes AFIP suele haber dos bloques.
    cuits = re.findall(
        r"(?:CUIT|C\.?U\.?I\.?T\.?)[:\s]*([0-9]{2}[- ]?[0-9]{8}[- ]?[0-9])",
        bloque,
        flags=re.IGNORECASE,
    )
    cuits_n = [normalizar_cuit(c) for c in cuits if normalizar_cuit(c)]
    # Dedup preservando orden
    seen: set[str] = set()
    cuits_u: list[str] = []
    for c in cuits_n:
        if c not in seen:
            seen.add(c)
            cuits_u.append(c)

    cuit_emisor = cuits_u[0] if cuits_u else ""
    cuit_receptor = cuits_u[1] if len(cuits_u) > 1 else ""

    # Razón social emisor: heurística cerca de "Razón Social" / primera línea comercial
    proveedor = ""
    m_rs = re.search(
        r"(?:Raz[oó]n\s+Social|Apellido\s+y\s+Nombre)[:\s]*([^\n\r]{3,80})",
        bloque,
        flags=re.IGNORECASE,
    )
    if m_rs:
        proveedor = re.sub(r"\s+", " ", m_rs.group(1)).strip(" :-")
    # A veces hay dos "Razón Social": emisor arriba, receptor abajo — tomar la primera no vacía
    razones = re.findall(
        r"(?:Raz[oó]n\s+Social|Apellido\s+y\s+Nombre)[:\s]*([^\n\r]{3,80})",
        bloque,
        flags=re.IGNORECASE,
    )
    if razones:
        proveedor = re.sub(r"\s+", " ", razones[0]).strip(" :-")

    receptor_nombre = ""
    if len(razones) > 1:
        receptor_nombre = re.sub(r"\s+", " ", razones[1]).strip(" :-")

    comprobante = ""
    if pv and nro:
        try:
            comprobante = formatear_comprobante_tango(pv, nro)
        except Exception:
            comprobante = formatear_comprobante(pv, nro)

    importe_firmado = round(abs(float(importe or 0)) * (signo or 1), 2)

    return {
        "Proveedor": proveedor,
        "Tipo": tipo,
        "Código AFIP": codigo,
        "Nº Comprobante": comprobante,
        "Fecha": fecha,
        "Total": importe_firmado,
        "CAE": cae or "",
        "CUIT Emisor": cuit_emisor,
        "CUIT Receptor": cuit_receptor,
        "Receptor Nombre": receptor_nombre,
        "Punto Venta": pv,
        "Número": nro,
    }


def iter_archivos_factura(uploads: Sequence[Any]) -> list[tuple[str, bytes]]:
    """PDFs/imágenes sueltos o dentro de ZIP (memoria)."""
    salida: list[tuple[str, bytes]] = []
    vistos: set[str] = set()
    exts_ok = {".pdf"} | _IMG_EXT

    for uploaded in uploads or []:
        nombre = str(getattr(uploaded, "name", "archivo"))
        if hasattr(uploaded, "getvalue"):
            data = uploaded.getvalue()
        elif isinstance(uploaded, (bytes, bytearray)):
            data = bytes(uploaded)
        elif isinstance(uploaded, (str, Path)):
            path = Path(uploaded)
            data = path.read_bytes()
            nombre = path.name
        else:
            data = uploaded.read()  # type: ignore[union-attr]

        lower = nombre.lower()
        if lower.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    ext = Path(info.filename).suffix.lower()
                    if ext not in exts_ok:
                        continue
                    raw = zf.read(info.filename)
                    digest = hashlib.sha1(raw).hexdigest()
                    if digest in vistos:
                        continue
                    vistos.add(digest)
                    interno = Path(info.filename).name or info.filename
                    salida.append((f"{nombre}::{interno}", raw))
        else:
            ext = Path(lower).suffix
            if ext not in exts_ok:
                continue
            digest = hashlib.sha1(data).hexdigest()
            if digest in vistos:
                continue
            vistos.add(digest)
            salida.append((nombre, data))
    return salida


def extraer_facturas(
    uploads: Sequence[Any],
    *,
    usar_ocr: bool = True,
) -> tuple[pd.DataFrame, list[dict]]:
    """Extrae comprobantes desde PDF/fotos. Retorna (df, errores)."""
    filas: list[dict] = []
    errores: list[dict] = []

    for nombre, data in iter_archivos_factura(uploads):
        try:
            texto, metodo = extraer_texto_comprobante(nombre, data, usar_ocr=usar_ocr)
            if not texto.strip():
                errores.append({
                    "archivo": nombre,
                    "motivo": f"Sin texto legible ({metodo}).",
                })
                continue
            parsed = _extraer_partes_factura(texto)
            parsed["Archivo"] = nombre
            parsed["Metodo"] = metodo
            if not parsed.get("Nº Comprobante") and not parsed.get("CAE") and not parsed.get("Total"):
                errores.append({
                    "archivo": nombre,
                    "motivo": "No se detectaron comprobante/CAE/importe.",
                })
                continue
            # Absoluto para cruce (NC también se cruzan por número; el signo se refleja en Total)
            filas.append(parsed)
        except Exception as exc:
            errores.append({"archivo": nombre, "motivo": str(exc)})

    if not filas:
        return pd.DataFrame(columns=COLUMNAS_MATCH), errores
    return pd.DataFrame(filas), errores


# ── Cruce ─────────────────────────────────────────────────────────────────────

def cruzar_facturas_vs_arca(
    df_facturas: pd.DataFrame,
    df_arca: pd.DataFrame,
    *,
    cuit_contribuyente: str = "",
    tol_importe: float = TOL_IMPORTE,
) -> dict[str, pd.DataFrame]:
    """
    Cruza facturas extraídas vs listado ARCA.

    Retorna dict con DataFrames: matcheadas, a_revisar, faltantes, diferencias.
    """
    cuit_esp = normalizar_cuit(cuit_contribuyente)
    fact = df_facturas.copy() if df_facturas is not None else pd.DataFrame()
    arca = df_arca.copy() if df_arca is not None else pd.DataFrame()

    # Indexar ARCA
    arca_rows: list[dict] = arca.to_dict(orient="records") if not arca.empty else []
    usados_arca: set[int] = set()

    def _arca_clave(row: dict) -> str:
        return _clave_match(
            cae=str(row.get("CAE") or ""),
            cuit_emisor=str(row.get("CUIT Emisor") or ""),
            comprobante=str(row.get("Comprobante") or ""),
        )

    indice: dict[str, list[int]] = {}
    for i, row in enumerate(arca_rows):
        k = _arca_clave(row)
        if k:
            indice.setdefault(k, []).append(i)
        # también indexar solo por comprobante por si la factura no trae CUIT
        cmpte = str(row.get("Comprobante") or "")
        if cmpte:
            indice.setdefault(f"cmp:{cmpte}", []).append(i)
        cae = normalizar_cae(row.get("CAE"))
        if cae:
            indice.setdefault(f"cae:{cae}", []).append(i)

    matcheadas: list[dict] = []
    a_revisar: list[dict] = []
    diferencias: list[dict] = []

    for _, f in fact.iterrows() if not fact.empty else []:
        proveedor = str(f.get("Proveedor") or "").strip()
        tipo = str(f.get("Tipo") or "").strip()
        cmpte = str(f.get("Nº Comprobante") or f.get("Comprobante") or "").strip()
        if not cmpte:
            _, _, cmpte = parsear_comprobante_texto(
                f"{f.get('Punto Venta')}-{f.get('Número')}"
            )
        fecha = _fmt_fecha(f.get("Fecha"))
        total_f = round(abs(float(f.get("Total") or 0)), 2)
        cae = normalizar_cae(f.get("CAE"))
        cuit_e = normalizar_cuit(f.get("CUIT Emisor"))
        cuit_r = normalizar_cuit(f.get("CUIT Receptor"))
        archivo = str(f.get("Archivo") or "")

        # Receptor distinto al contribuyente del listado
        if cuit_esp and cuit_r and cuit_r != cuit_esp:
            a_revisar.append({
                "Proveedor": proveedor,
                "Comprobante": f"{tipo} {cmpte}".strip(),
                "Fecha": fecha,
                "Total": total_f,
                "Motivo / Observación": (
                    f"Facturada a CUIT {cuit_r}"
                    + (f" ({f.get('Receptor Nombre')})" if f.get("Receptor Nombre") else "")
                    + f", no al contribuyente del listado ({cuit_esp}). "
                    "No corresponde que esté en el listado ARCA de esa sociedad."
                ),
                "CUIT Emisor": cuit_e,
                "CUIT Receptor": cuit_r,
                "Archivo": archivo,
            })
            continue

        candidatos_idx: list[int] = []
        for k in (
            _clave_match(cae=cae, cuit_emisor=cuit_e, comprobante=cmpte),
            f"cae:{cae}" if cae else "",
            f"cuit+cmp:{cuit_e}|{cmpte}" if (cuit_e and cmpte) else "",
            f"cmp:{cmpte}" if cmpte else "",
        ):
            if not k:
                continue
            for idx in indice.get(k, []):
                if idx not in usados_arca and idx not in candidatos_idx:
                    candidatos_idx.append(idx)

        # Soft: mismo CUIT emisor + misma fecha + importe cercano
        if not candidatos_idx and cuit_e and fecha:
            for i, row in enumerate(arca_rows):
                if i in usados_arca:
                    continue
                if normalizar_cuit(row.get("CUIT Emisor")) != cuit_e:
                    continue
                if _fmt_fecha(row.get("Fecha")) != fecha:
                    continue
                if abs(float(row.get("Total") or 0) - total_f) <= max(tol_importe, total_f * 0.02):
                    candidatos_idx.append(i)
                    break

        if not candidatos_idx:
            motivo = "Factura subida que NO figura en el listado ARCA. Revisar: puede faltar en ARCA o ser documento no fiscal / otro CUIT."
            if not cuit_r and cuit_esp:
                motivo = (
                    "Campo cliente/CUIT receptor vacío o no legible. "
                    "No se puede confirmar que sea del contribuyente. No figura en ARCA."
                )
            a_revisar.append({
                "Proveedor": proveedor,
                "Comprobante": f"{tipo} {cmpte}".strip(),
                "Fecha": fecha,
                "Total": total_f,
                "Motivo / Observación": motivo,
                "CUIT Emisor": cuit_e,
                "CUIT Receptor": cuit_r,
                "Archivo": archivo,
            })
            continue

        idx = candidatos_idx[0]
        usados_arca.add(idx)
        arow = arca_rows[idx]
        total_a = round(abs(float(arow.get("Total") or 0)), 2)
        dif = round(total_f - total_a, 2)

        proveedor_out = proveedor or str(arow.get("Proveedor") or "")
        tipo_out = tipo or str(arow.get("Tipo") or "")
        cmpte_out = cmpte or str(arow.get("Comprobante") or "")
        fecha_out = fecha or _fmt_fecha(arow.get("Fecha"))
        cae_out = cae or normalizar_cae(arow.get("CAE")) or "-"

        if abs(dif) > tol_importe:
            diferencias.append({
                "Proveedor": proveedor_out,
                "Tipo": tipo_out,
                "Nº Comprobante": cmpte_out,
                "Fecha": fecha_out,
                "Total Factura": total_f,
                "Total ARCA": total_a,
                "Diferencia": dif,
                "CAE": cae_out,
                "Archivo": archivo,
            })
        else:
            matcheadas.append({
                "Proveedor": proveedor_out,
                "Tipo": tipo_out,
                "Nº Comprobante": cmpte_out,
                "Fecha": fecha_out,
                "Total": total_a if total_a else total_f,
                "CAE": cae_out if cae_out else "-",
                "CUIT Emisor": cuit_e or normalizar_cuit(arow.get("CUIT Emisor")),
                "Archivo": archivo,
                "Estado": "OK",
            })

    faltantes: list[dict] = []
    for i, row in enumerate(arca_rows):
        if i in usados_arca:
            continue
        faltantes.append({
            "Fecha": _fmt_fecha(row.get("Fecha")),
            "Proveedor": str(row.get("Proveedor") or ""),
            "Tipo": str(row.get("Tipo") or ""),
            "Comprobante": str(row.get("Comprobante") or ""),
            "Total": round(abs(float(row.get("Total") or 0)), 2),
            "CUIT Emisor": normalizar_cuit(row.get("CUIT Emisor")),
            "CAE": normalizar_cae(row.get("CAE")) or "",
        })

    def _df(cols: list[str], rows: list[dict]) -> pd.DataFrame:
        if not rows:
            return pd.DataFrame(columns=cols)
        df = pd.DataFrame(rows)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df[cols]

    return {
        "matcheadas": _df(COLUMNAS_MATCH, matcheadas),
        "a_revisar": _df(COLUMNAS_REVISAR, a_revisar),
        "faltantes": _df(COLUMNAS_FALTANTES, faltantes),
        "diferencias": _df(COLUMNAS_DIF, diferencias),
    }


# ── Export ────────────────────────────────────────────────────────────────────

def exportar_cruce_excel(
    resultado: dict[str, pd.DataFrame],
    *,
    titulo: str = "Cruce Facturas vs ARCA",
    subtitulo: str = "Mis Comprobantes · Estudio Contable",
    periodo: str = "",
    cuit: str = "",
) -> bytes:
    from excel_formato_estudio import exportar_informe_excel

    m = resultado.get("matcheadas")
    r = resultado.get("a_revisar")
    f = resultado.get("faltantes")
    d = resultado.get("diferencias")
    if m is None:
        m = pd.DataFrame(columns=COLUMNAS_MATCH)
    if r is None:
        r = pd.DataFrame(columns=COLUMNAS_REVISAR)
    if f is None:
        f = pd.DataFrame(columns=COLUMNAS_FALTANTES)
    if d is None:
        d = pd.DataFrame(columns=COLUMNAS_DIF)

    resumen = pd.DataFrame(
        [
            {"Concepto": "Matcheadas (OK)", "Cantidad": len(m), "Importe": round(float(pd.to_numeric(m["Total"], errors="coerce").fillna(0).sum()) if not m.empty else 0, 2)},
            {"Concepto": "A revisar (en factura, no en ARCA)", "Cantidad": len(r), "Importe": round(float(pd.to_numeric(r["Total"], errors="coerce").fillna(0).sum()) if not r.empty else 0, 2)},
            {"Concepto": "Faltantes (en ARCA, sin factura)", "Cantidad": len(f), "Importe": round(float(pd.to_numeric(f["Total"], errors="coerce").fillna(0).sum()) if not f.empty else 0, 2)},
            {"Concepto": "Diferencias de importe", "Cantidad": len(d), "Importe": round(float(pd.to_numeric(d.get("Diferencia", pd.Series(dtype=float)), errors="coerce").fillna(0).abs().sum()) if not d.empty else 0, 2)},
        ]
    )

    sub = subtitulo
    if cuit:
        sub = f"{subtitulo} · CUIT {cuit}"

    return exportar_informe_excel(
        titulo=titulo,
        subtitulo=sub,
        periodo=periodo,
        kpis=[
            ("Matcheadas", len(m), "int"),
            ("A revisar", len(r), "int"),
            ("Faltantes", len(f), "int"),
            ("Diferencias", len(d), "int"),
        ],
        resumenes=[("Resumen del cruce", resumen)],
        detalle=m,
        hoja_detalle="Matcheadas",
        hojas_adicionales=[
            ("A revisar", r),
            ("Faltantes", f),
            ("Diferencias", d),
        ],
        col_moneda=[
            "Total", "Importe", "Total Factura", "Total ARCA", "Diferencia",
        ],
        col_fecha=["Fecha"],
        total_col="Total",
    )


def guardar_cruce_excel(
    ruta: str | Path,
    resultado: dict[str, pd.DataFrame],
    **kwargs: Any,
) -> Path:
    from excel_formato_estudio import guardar_informe_excel

    m = resultado.get("matcheadas", pd.DataFrame())
    r = resultado.get("a_revisar", pd.DataFrame())
    f = resultado.get("faltantes", pd.DataFrame())
    d = resultado.get("diferencias", pd.DataFrame())

    resumen = pd.DataFrame(
        [
            {"Concepto": "Matcheadas (OK)", "Cantidad": len(m), "Importe": round(float(pd.to_numeric(m["Total"], errors="coerce").fillna(0).sum()) if not m.empty and "Total" in m.columns else 0, 2)},
            {"Concepto": "A revisar", "Cantidad": len(r), "Importe": round(float(pd.to_numeric(r["Total"], errors="coerce").fillna(0).sum()) if not r.empty and "Total" in r.columns else 0, 2)},
            {"Concepto": "Faltantes", "Cantidad": len(f), "Importe": round(float(pd.to_numeric(f["Total"], errors="coerce").fillna(0).sum()) if not f.empty and "Total" in f.columns else 0, 2)},
            {"Concepto": "Diferencias", "Cantidad": len(d), "Importe": 0.0},
        ]
    )
    cuit = str(kwargs.get("cuit") or "")
    subtitulo = str(kwargs.get("subtitulo") or "Mis Comprobantes · Estudio Contable")
    if cuit:
        subtitulo = f"{subtitulo} · CUIT {cuit}"

    return guardar_informe_excel(
        ruta,
        titulo=str(kwargs.get("titulo") or "Cruce Facturas vs ARCA"),
        subtitulo=subtitulo,
        periodo=str(kwargs.get("periodo") or ""),
        kpis=[
            ("Matcheadas", len(m), "int"),
            ("A revisar", len(r), "int"),
            ("Faltantes", len(f), "int"),
            ("Diferencias", len(d), "int"),
        ],
        resumenes=[("Resumen del cruce", resumen)],
        detalle=m if m is not None else pd.DataFrame(),
        hoja_detalle="Matcheadas",
        hojas_adicionales=[
            ("A revisar", r if r is not None else pd.DataFrame()),
            ("Faltantes", f if f is not None else pd.DataFrame()),
            ("Diferencias", d if d is not None else pd.DataFrame()),
        ],
        col_moneda=["Total", "Importe", "Total Factura", "Total ARCA", "Diferencia"],
        col_fecha=["Fecha"],
        total_col="Total",
    )


def procesar_cruce_facturas_arca(
    facturas_uploads: Sequence[Any],
    arca_upload: Any,
    *,
    usar_ocr: bool = True,
    cuit_contribuyente: str = "",
    nombre_arca: str | None = None,
) -> tuple[dict[str, pd.DataFrame], list[dict], str, bytes]:
    """
    Pipeline completo para Streamlit.

    Retorna (resultado_dfs, errores_facturas, cuit_detectado, excel_bytes).
    """
    nombre = nombre_arca or str(getattr(arca_upload, "name", "arca.xlsx"))
    if hasattr(arca_upload, "getvalue"):
        arca_bytes = arca_upload.getvalue()
    elif isinstance(arca_upload, (str, Path)):
        arca_bytes = Path(arca_upload).read_bytes()
        nombre = Path(arca_upload).name
    else:
        arca_bytes = arca_upload  # type: ignore[assignment]

    df_arca, cuit_det = leer_listado_arca(arca_bytes, nombre=nombre)
    cuit = normalizar_cuit(cuit_contribuyente) or cuit_det

    df_fact, errores = extraer_facturas(facturas_uploads, usar_ocr=usar_ocr)
    resultado = cruzar_facturas_vs_arca(df_fact, df_arca, cuit_contribuyente=cuit)
    xlsx = exportar_cruce_excel(resultado, cuit=cuit)
    return resultado, errores, cuit, xlsx
