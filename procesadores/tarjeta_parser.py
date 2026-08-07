"""
Parser de liquidaciones de tarjeta por plantillas.
Soporta Prisma/Visa, Mercado Pago y First Data/Posnet.
"""

from __future__ import annotations

import io
import re
import unicodedata
from pathlib import Path
from typing import Any

import pdfplumber

# ---------------------------------------------------------------------------
# Plantillas: keywords de detección + etiquetas / regex por concepto fiscal
# ---------------------------------------------------------------------------

PLANTILLAS_TARJETAS: dict[str, dict[str, Any]] = {
    "CABAL": {
        "keywords": ["CABAL", "LIQUIDACION PARA PAGO A COMERCIOS"],
        "campos": {},
    },
    "FIRST DATA (Posnet)": {
        "keywords": [
            "FIRST DATA", "FIRSTDATA", "FISERV", "POSNET",
            "TOTAL LIQ. TARJ. CREDITO", "CENTRO DE ATENCIÓN TELEFÓNICA A ESTABLECIMIENTOS",
        ],
        "campos": {},
    },
    "PRISMA (Visa/Mastercard)": {
        "keywords": [
            "PRISMA", "VISA", "MASTERCARD", "DEBIN PRISMA", "PAYWAY",
            "RESUMEN MENSUAL DE LIQUIDACIONES", "DESGLOSE DE DESCUENTOS",
        ],
        "campos": {},  # extracción dedicada en _extraer_prisma_galicia
    },
    "MERCADO PAGO": {
        "keywords": ["MERCADO PAGO", "MERCADOPAGO"],
        "campos": {
            "Neto_Gravado": {
                "regex": (
                    r"(?:Tarifa\s+por\s+procesamiento|Costo\s+de\s+servicio|Comisi[oó]n)"
                    r"\s*[:\-]?\s*\$?\s*([\d\.,]+)"
                ),
                "etiquetas": [
                    "Tarifa por procesamiento",
                    "Costo de servicio",
                    "Comision",
                    "Comisión",
                ],
            },
            "IVA_21": {
                "regex": r"IVA\s*(?:21\s*%?)?\s*[:\-]?\s*\$?\s*([\d\.,]+)",
                "etiquetas": ["IVA 21%", "IVA"],
            },
            "Retencion_IVA": {
                "regex": r"Retenci[oó]n(?:es)?\s+(?:de\s+)?IVA\s*[:\-]?\s*\$?\s*([\d\.,]+)",
                "etiquetas": ["Retencion IVA", "Retención IVA"],
            },
            "Retencion_IIBB": {
                "regex": r"Retenci[oó]n(?:es)?\s+(?:de\s+)?IIBB\s*[:\-]?\s*\$?\s*([\d\.,]+)",
                "etiquetas": ["Retencion de IIBB", "Retención de IIBB", "IIBB"],
            },
            "Percepcion_IVA": {
                "regex": r"Percepci[oó]n(?:es)?\s+(?:de\s+)?IVA\s*[:\-]?\s*\$?\s*([\d\.,]+)",
                "etiquetas": ["Percepcion IVA", "Percepción IVA"],
            },
            "Percepcion_IIBB": {
                "regex": r"Percepci[oó]n(?:es)?\s+(?:de\s+)?IIBB\s*[:\-]?\s*\$?\s*([\d\.,]+)",
                "etiquetas": ["Percepcion IIBB", "Percepción IIBB"],
            },
        },
    },
    "NARANJA": {
        "keywords": ["TARJETA NARANJA", "NARANJA S.A.U.", "NARANJA SAU"],
        "campos": {},
    },
    "AJUSTES / CONTRACARGOS": {
        "keywords": [
            "DETALLE DE AJUSTES",
            "CONTRACARGOS",
            "PROMOCIONES",
            "DTO.PROMO.COM",
            "PER.DTO.PR.COM",
        ],
        "campos": {},
    },
}

CAMPOS_MONTO = (
    "Neto_Gravado",
    "IVA_21",
    "Percepcion_IVA",
    "Retencion_IVA",
    "Retencion_IIBB",
    "Percepcion_IIBB",
)


def _normalizar(texto: str) -> str:
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()


def _limpiar_monto(valor: object) -> float:
    if valor is None:
        return 0.0
    txt = str(valor).replace("$", "").replace("\xa0", "").replace(" ", "").strip()
    if not txt or txt in ("-", "—", "–"):
        return 0.0
    if re.search(r"\d,\d{2}$", txt):
        txt = txt.replace(".", "").replace(",", ".")
    else:
        txt = txt.replace(",", "")
    try:
        return round(float(txt), 2)
    except ValueError:
        return 0.0


def _buscar_monto_regex(patron: str, texto: str) -> float:
    match = re.search(patron, texto or "", re.IGNORECASE)
    if not match:
        return 0.0
    raw = next((g for g in match.groups() if g), None)
    return _limpiar_monto(raw) if raw else 0.0


def _buscar_monto_cerca_etiqueta(texto: str, etiquetas: list[str]) -> float:
    lineas = (texto or "").splitlines()
    etiquetas_n = [_normalizar(e) for e in etiquetas]
    for i, linea in enumerate(lineas):
        low = _normalizar(linea)
        if not any(e in low for e in etiquetas_n):
            continue
        montos = re.findall(r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})", linea)
        if not montos and i + 1 < len(lineas):
            montos = re.findall(
                r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})",
                lineas[i + 1],
            )
        if montos:
            return _limpiar_monto(montos[-1])
    return 0.0


def _extraer_seccion(texto: str, inicio: str, fin: str | list[str] | None = None) -> str:
    """Devuelve el texto entre un título de inicio y el primer título de fin (o fin de texto)."""
    lineas = (texto or "").splitlines()
    rx_ini = re.compile(re.escape(inicio), re.IGNORECASE)
    fines = [fin] if isinstance(fin, str) else list(fin or [])
    rx_fines = [re.compile(re.escape(f), re.IGNORECASE) for f in fines]
    bloque: list[str] = []
    activo = False
    for linea in lineas:
        if not activo:
            if rx_ini.search(linea):
                activo = True
                bloque.append(linea)
            continue
        if any(rx.search(linea) for rx in rx_fines):
            break
        bloque.append(linea)
    return "\n".join(bloque)


def _sumar_montos_lineas(texto: str, patron_linea: str) -> float:
    """Suma el último monto de las líneas que matchean el patrón (o la siguiente)."""
    total = 0.0
    rx = re.compile(patron_linea, re.IGNORECASE)
    lineas = (texto or "").splitlines()
    for i, linea in enumerate(lineas):
        if not rx.search(linea):
            continue
        montos = re.findall(r"[-−]?\s*\$?\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*[-−]?", linea)
        if not montos and i + 1 < len(lineas):
            montos = re.findall(
                r"[-−]?\s*\$?\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*[-−]?",
                lineas[i + 1],
            )
        if not montos:
            continue
        total += abs(_limpiar_monto(montos[-1]))
    return round(total, 2)


def _monto_ultima_linea(texto: str, patron_linea: str, firmado: bool = False) -> float:
    """Devuelve el último monto encontrado en la última línea que matchea el patrón."""
    rx = re.compile(patron_linea, re.IGNORECASE | re.MULTILINE)
    lineas = (texto or "").splitlines()
    ultimo = 0.0
    for i, linea in enumerate(lineas):
        if not rx.search(linea):
            continue
        montos = re.findall(r"([-−]?)\s*\$?\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})", linea)
        if not montos and i + 1 < len(lineas):
            montos = re.findall(
                r"([-−]?)\s*\$?\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})",
                lineas[i + 1],
            )
        if not montos:
            continue
        signo, raw = montos[-1]
        valor = abs(_limpiar_monto(raw))
        if firmado and (signo in ("-", "−") or "−" in linea or linea.strip().startswith("-")):
            valor = -valor
        ultimo = valor
    return round(ultimo, 2)


def _completar_percepciones_retenciones(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    """
    Completa Percepción/Retención IVA e IIBB con etiquetas reales de liquidaciones AR
    (Naranja, Prisma/Payway/Galicia, resúmenes Visa, etc.).
    """
    if not float(datos.get("Percepcion_IVA") or 0):
        datos["Percepcion_IVA"] = (
            _sumar_montos_lineas(texto_pdf, r"Percepci[oó]n(?:es)?\s+(?:de\s+)?IVA")
            or _sumar_montos_lineas(texto_pdf, r"PERCEP\.?\s*IVA")
            or _sumar_montos_lineas(texto_pdf, r"Percep\./\s*Retenc\.?\s*AFIP")
            or _sumar_montos_lineas(texto_pdf, r"Percep\./Retenc\.AFIP")
            or _sumar_montos_lineas(texto_pdf, r"Percepci[oó]n\s+IVA\s+RG")
        )
    if not float(datos.get("Percepcion_IIBB") or 0):
        datos["Percepcion_IIBB"] = (
            _sumar_montos_lineas(
                texto_pdf,
                r"Percepci[oó]n(?:es)?\s+(?:de\s+)?(?:Ingresos?\s+Brutos|IIBB|IB\b)",
            )
            or _sumar_montos_lineas(texto_pdf, r"Perc\.?\s*IB\b")
            or _sumar_montos_lineas(texto_pdf, r"IIBB\s+PERCEP")
            or _sumar_montos_lineas(texto_pdf, r"PERCEPCION\s+ING\.?\s*BRUTOS")
            or _sumar_montos_lineas(texto_pdf, r"SIRCREB")
        )
    if not float(datos.get("Retencion_IIBB") or 0):
        datos["Retencion_IIBB"] = (
            _sumar_montos_lineas(
                texto_pdf,
                r"Retenci[oó]n(?:es)?\s+(?:de\s+)?(?:Ingresos?\s+Brutos|IIBB|Ing\.?\s*Brutos)",
            )
            or _sumar_montos_lineas(texto_pdf, r"Ret\.?\s*IB\s*SIRTAC")
            or _sumar_montos_lineas(texto_pdf, r"\bSIRTAC\b")
            or _sumar_montos_lineas(texto_pdf, r"RETENCION\s+IIBB")
            or _sumar_montos_lineas(texto_pdf, r"Total de Retenciones(?:\s+Impositivas)?")
        )
    if not float(datos.get("Retencion_IVA") or 0):
        datos["Retencion_IVA"] = (
            _sumar_montos_lineas(texto_pdf, r"Retenci[oó]n(?:es)?\s+(?:de\s+)?IVA")
            or _sumar_montos_lineas(texto_pdf, r"Ret\.?\s*IVA\b")
        )
    return datos


def _calcular_total_descontado(datos: dict[str, Any], extra: float = 0.0) -> float:
    total = round(sum(float(datos.get(c) or 0) for c in CAMPOS_MONTO) + float(extra or 0), 2)
    datos["Total_Descontado"] = total
    return total


def _extraer_cabal(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    datos["Entidad"] = "CABAL"
    datos["Neto_Gravado"] = round(
        _sumar_montos_lineas(texto_pdf, r"^ARANCEL DE DESCUENTO\s+(?!\d+,\d+%)")
        + _sumar_montos_lineas(texto_pdf, r"^COSTO FINANCIERO TOTAL")
        + _sumar_montos_lineas(texto_pdf, r"^COMISION ADMINISTRACION"),
        2,
    )
    # Solo IVA 21 % explícito; evitar la línea combinada que duplica arancel + costo financiero.
    iva21 = _sumar_montos_lineas(texto_pdf, r"^IVA S/ARANCEL DE DESCUENTO")
    iva21 += _sumar_montos_lineas(texto_pdf, r"^[−-]?IVA\s+21,00%")
    datos["IVA_21"] = round(iva21, 2)
    iva105 = _sumar_montos_lineas(texto_pdf, r"^IVA S/COSTO FINANCIERO")
    datos["Retencion_IIBB"] = _sumar_montos_lineas(texto_pdf, r"RETENCION IIBB")
    _completar_percepciones_retenciones(texto_pdf, datos)
    _calcular_total_descontado(datos, extra=iva105)
    return datos


def _extraer_first_data(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    datos["Entidad"] = "FIRST DATA (Posnet)"
    datos["Neto_Gravado"] = _sumar_montos_lineas(texto_pdf, r"^-\s*ARANCEL\b")
    datos["IVA_21"] = _sumar_montos_lineas(texto_pdf, r"IVA CRED\.?FISC\.?COMERCIO|IVA.*ARANC")
    datos["Retencion_IIBB"] = _sumar_montos_lineas(
        texto_pdf, r"RETENCION ING\.?BRUTOS|RETENCION IIBB"
    )
    _completar_percepciones_retenciones(texto_pdf, datos)
    _calcular_total_descontado(datos)
    return datos


def _extraer_naranja(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    """
    Soporta:
    - Constancia de retención provincial (solo IIBB)
    - Liquidación comercial completa (arancel, IVA, percepciones, SIRTAC)
    """
    datos["Entidad"] = "NARANJA"
    facturacion = _extraer_seccion(
        texto_pdf,
        "Detalles de facturación",
        ["Detalle Retenciones Impositivas", "Neto Liquidado"],
    )
    retenciones = _extraer_seccion(
        texto_pdf,
        "Detalle Retenciones Impositivas",
        ["Neto Liquidado", "A Pagar"],
    )

    if facturacion.strip():
        arancel = _sumar_montos_lineas(facturacion, r"^Arancel\b")
        intereses = _sumar_montos_lineas(facturacion, r"^Inter[eé]s\s+Plan\b")
        bonificacion = _sumar_montos_lineas(facturacion, r"^Bonificaci[oó]n\b")
        datos["Neto_Gravado"] = round(arancel + intereses - bonificacion, 2)
        datos["IVA_21"] = _sumar_montos_lineas(facturacion, r"^IVA\s+21")
        datos["Percepcion_IVA"] = _sumar_montos_lineas(
            facturacion, r"Percepci[oó]n\s+(?:de\s+)?IVA"
        )
        datos["Percepcion_IIBB"] = _sumar_montos_lineas(
            facturacion, r"Percepci[oó]n\s+(?:de\s+)?Ingresos?\s+Brutos"
        )
        total_fact = _sumar_montos_lineas(facturacion, r"^Total de Facturaci[oó]n\b")
    else:
        total_fact = 0.0

    datos["Retencion_IIBB"] = (
        _sumar_montos_lineas(retenciones, r"\bSirtac\b")
        or _sumar_montos_lineas(retenciones, r"^Total de Retenciones(?:\s+Impositivas)?")
        or _sumar_montos_lineas(texto_pdf, r"Total de Retenciones(?:\s+Impositivas)?")
        or _monto_ultima_linea(texto_pdf, r"Importe de la Retenci[oó]n Ingresos Brutos")
    )
    _completar_percepciones_retenciones(texto_pdf, datos)

    if total_fact > 0:
        datos["Total_Descontado"] = round(total_fact + float(datos.get("Retencion_IIBB") or 0), 2)
    else:
        _calcular_total_descontado(datos)
    return datos


def _extraer_prisma_galicia(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    """Resumen mensual Prisma/Payway/Galicia (liquidación a comercios)."""
    datos["Entidad"] = "PRISMA (Visa/Mastercard)"
    desglose = _extraer_seccion(
        texto_pdf,
        "DESGLOSE DE DESCUENTOS",
        [
            "Continúa en la página siguiente",
            "Continua en la pagina siguiente",
            "SR. COMERCIANTE",
            "(*) Servicio prestado por PAYWAY",
        ],
    )
    fuente = desglose if desglose.strip() else texto_pdf

    # Base imponible del desglose mensual (coincide con Tasa 21 % + Tasa 10,5 %).
    base21 = _sumar_montos_lineas(fuente, r"^Tasa\s+21")
    base105 = _sumar_montos_lineas(fuente, r"^Tasa\s+10[,.]50")
    neto = round(base21 + base105, 2)
    if not neto:
        neto = round(
            _sumar_montos_lineas(fuente, r"^Arancel\s+Tj\.")
            + _sumar_montos_lineas(fuente, r"Pago Expreso de Cupones")
            + _sumar_montos_lineas(fuente, r"Ventas?\s+en\s+\d+\s+cuotas"),
            2,
        )
    datos["Neto_Gravado"] = neto

    iva21 = _sumar_montos_lineas(fuente, r"^IVA\s+21")
    iva105 = _sumar_montos_lineas(fuente, r"^IVA\s+10[,.]50")
    datos["IVA_21"] = round(iva21, 2)
    datos["Percepcion_IIBB"] = _sumar_montos_lineas(fuente, r"Perc\.?\s*IB\b")
    datos["Retencion_IIBB"] = _sumar_montos_lineas(fuente, r"Ret\.?\s*IB\s*SIRTAC")
    datos["Percepcion_IVA"] = _sumar_montos_lineas(
        fuente, r"Percep\./Retenc\.AFIP|Percep\./\s*Retenc\.?\s*AFIP"
    )
    _completar_percepciones_retenciones(fuente, datos)
    sellos = _sumar_montos_lineas(fuente, r"Sellos\s+PEX")
    _calcular_total_descontado(datos, extra=iva105 + sellos)
    return datos


def _extraer_ajustes_contracargos(texto_pdf: str, datos: dict[str, Any]) -> dict[str, Any]:
    datos["Entidad"] = "AJUSTES / CONTRACARGOS"
    total = _monto_ultima_linea(
        texto_pdf,
        r"^[\s\-−]*\d{1,3}(?:\.\d{3})*,\d{2}\s*$|^[\s\-−]*\d+,\d{2}\s*$",
        firmado=True,
    )
    datos["Neto_Gravado"] = total
    datos["Total_Descontado"] = total
    return datos


def detectar_entidad_por_texto(texto_pdf: str) -> str:
    """Devuelve la clave de PLANTILLAS_TARJETAS más probable, o 'Otra / No detectada'."""
    texto_u = (texto_pdf or "").upper()
    texto_n = _normalizar(texto_pdf or "")
    # Mercado Pago: exigir ambas palabras para no confundir con "MP" suelto
    if "mercado pago" in texto_n or "mercadopago" in texto_n:
        return "MERCADO PAGO"
    mejor = "Otra / No detectada"
    mejor_score = 0
    for entidad, plantilla in PLANTILLAS_TARJETAS.items():
        if entidad == "MERCADO PAGO":
            continue
        score = sum(1 for kw in plantilla.get("keywords") or [] if kw.upper() in texto_u)
        if score > mejor_score:
            mejor_score = score
            mejor = entidad
    return mejor if mejor_score > 0 else "Otra / No detectada"


def extraer_con_plantilla(texto_pdf: str, entidad: str) -> dict:
    """
    Extrae conceptos fiscales usando la plantilla de la entidad confirmada.
    Si la entidad es 'Otra / No detectada', intenta heurística genérica.
    """
    datos: dict[str, Any] = {
        "Fecha": "No detectada",
        "Entidad": entidad if entidad != "Otra / No detectada" else "Desconocida",
        "Nro_Liquidacion": "S/D",
        "Neto_Gravado": 0.0,
        "IVA_21": 0.0,
        "Percepcion_IVA": 0.0,
        "Retencion_IVA": 0.0,
        "Retencion_IIBB": 0.0,
        "Percepcion_IIBB": 0.0,
        "Total_Descontado": 0.0,
    }

    nro = re.search(
        r"(?:liquidaci[oó]n\s*nro\.?|nro\.?\s*liq(?:uidaci[oó]n)?|lote|n[º°o]\.?\s*liquidaci[oó]n|liquidaci[oó]n)"
        r"\s*[:\-\.]?\s*(\d{4,})",
        texto_pdf or "",
        re.IGNORECASE,
    )
    if nro:
        datos["Nro_Liquidacion"] = nro.group(1)

    fecha_match = re.search(r"\b(\d{2}/\d{2}/(\d{2}|\d{4}))\b", texto_pdf or "")
    if fecha_match:
        fecha_txt = fecha_match.group(1)
        if re.fullmatch(r"\d{2}/\d{2}/\d{2}", fecha_txt):
            d, m, y = fecha_txt.split("/")
            fecha_txt = f"{d}/{m}/20{y}"
        datos["Fecha"] = fecha_txt

    if entidad == "CABAL":
        return _extraer_cabal(texto_pdf, datos)
    if entidad == "FIRST DATA (Posnet)":
        return _extraer_first_data(texto_pdf, datos)
    if entidad == "NARANJA":
        return _extraer_naranja(texto_pdf, datos)
    if entidad == "PRISMA (Visa/Mastercard)":
        return _extraer_prisma_galicia(texto_pdf, datos)
    if entidad == "AJUSTES / CONTRACARGOS":
        return _extraer_ajustes_contracargos(texto_pdf, datos)

    plantilla = PLANTILLAS_TARJETAS.get(entidad)
    if plantilla:
        for campo, cfg in (plantilla.get("campos") or {}).items():
            monto = _buscar_monto_regex(cfg.get("regex") or "", texto_pdf)
            if monto <= 0:
                monto = _buscar_monto_cerca_etiqueta(texto_pdf, list(cfg.get("etiquetas") or []))
            datos[campo] = monto
    else:
        # Heurística genérica si el usuario eligió "Otra"
        datos["Neto_Gravado"] = _buscar_monto_cerca_etiqueta(
            texto_pdf, ["Arancel", "Comision", "Comisión", "Costo Financiero", "Tarifa"]
        )
        datos["IVA_21"] = _buscar_monto_cerca_etiqueta(texto_pdf, ["IVA 21%", "IVA 21", "IVA"])
        datos["Retencion_IVA"] = _buscar_monto_cerca_etiqueta(
            texto_pdf, ["Retencion IVA", "Retención IVA"]
        )
        datos["Retencion_IIBB"] = _buscar_monto_cerca_etiqueta(
            texto_pdf, ["Retencion IIBB", "Retención IIBB", "IIBB"]
        )
        datos["Percepcion_IVA"] = _buscar_monto_cerca_etiqueta(
            texto_pdf, ["Percepcion IVA", "Percepción IVA"]
        )
        datos["Percepcion_IIBB"] = _buscar_monto_cerca_etiqueta(
            texto_pdf, ["Percepcion IIBB", "Percepción IIBB"]
        )

    _completar_percepciones_retenciones(texto_pdf, datos)
    _calcular_total_descontado(datos)
    return datos


def extraer_texto_liquidacion_pdf(file_buffer) -> str:
    """Extrae texto nativo del PDF; si viene vacío, intenta OCR vía procesador."""
    if hasattr(file_buffer, "getvalue"):
        data = file_buffer.getvalue()
        fuente = io.BytesIO(data)
    elif isinstance(file_buffer, (bytes, bytearray)):
        data = bytes(file_buffer)
        fuente = io.BytesIO(data)
    elif isinstance(file_buffer, (str, Path)):
        data = Path(file_buffer).read_bytes()
        fuente = io.BytesIO(data)
    else:
        data = b""
        fuente = file_buffer

    texto = ""
    with pdfplumber.open(fuente) as pdf:
        texto = "\n".join((page.extract_text() or "") for page in pdf.pages)

    if len(texto.strip()) >= 40:
        return texto

    # OCR fallback (lazy import para no circularizar al cargar el módulo)
    try:
        from procesador import _paginas_texto_extracto_pdf  # noqa: WPS433

        paginas = _paginas_texto_extracto_pdf(data, dpi_ocr=160, forzar_ocr=True)
        return "\n".join(t for _, t in paginas)
    except Exception:
        return texto
