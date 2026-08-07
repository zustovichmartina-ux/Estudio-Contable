"""
Analizador de inversiones — PDF/Excel → clasificación por especie → FIFO/PEPS.

Flujo:
1. Normalizar movimientos (AlyC/broker PDF o Excel).
2. Clasificar por grupo de especie (bonos, FCI, MEP/USD, acciones, otros).
3. Sembrar saldo inicial desde DDJJ (PDF) o Excel manual.
4. Aplicar FIFO por especie y exportar Excel de trabajo.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, BinaryIO

import pandas as pd

# ---------------------------------------------------------------------------
# Constantes / catálogos
# ---------------------------------------------------------------------------

GRUPOS_ESPECIE = (
    "Bonos / Títulos públicos",
    "FCI",
    "Dólar / MEP",
    "Acciones / Cedears",
    "Otros",
)

COLUMNAS_MOV = [
    "Fecha",
    "Especie",
    "Grupo",
    "Tipo_Operacion",
    "Cantidad",
    "Precio",
    "Monto_Total",
    "Moneda",
    "Descripcion",
    "Archivo origen",
    "Nueva_Clasificacion",
]

TITULOS_PUBLICOS = {
    "AL30", "AL30D", "GD30", "GD30D", "AL29", "GD29", "AL35", "GD35",
    "AL41", "GD41", "AE38", "GE38", "TX26", "TX28", "T2X4", "T4X4",
    "BONAR", "BOPREAL", "LECAP", "LETRA", "S31L", "S30O",
}

_ALIAS_COLS: dict[str, list[str]] = {
    "Fecha": [
        "fecha", "date", "fecha concertacion", "fecha liquidacion",
        "fecha operacion", "fecha operación", "fecha boleto",
    ],
    "Especie": [
        "especie", "ticker", "simbolo", "símbolo", "instrumento", "papel",
        "codigo", "código", "activo", "fondo", "descripcion especie",
    ],
    "Cantidad": [
        "cantidad", "cant", "nominal", "qty", "cuotapartes", "cuotas",
    ],
    "Moneda": ["moneda", "currency", "ccy"],
    "Monto_Total": [
        "monto total", "importe", "neto", "total", "monto", "bruto",
        "importe neto", "monto operado", "valor",
    ],
    "Precio": ["precio", "px", "cotizacion", "cotización", "precio unitario"],
    "Tipo_Operacion": [
        "tipo de operacion", "tipo de operación", "tipo", "operacion",
        "operación", "movimiento", "sentido", "side",
    ],
    "Descripcion": ["descripcion", "descripción", "detalle", "concepto", "obs"],
}

_COMPRA_KW = (
    "compra", "suscrip", "suscripción", "aporte", "ingreso", "alta",
    "acredita", "credito", "crédito", "recibid",
)
_VENTA_KW = (
    "venta", "rescate", "egreso", "baja", "debito", "débito", "retiro",
    "transferencia enviada", "pago",
)


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _norm(s: object) -> str:
    t = str(s or "").replace("\n", " ").strip().lower()
    t = (
        t.replace("á", "a").replace("é", "e").replace("í", "i")
        .replace("ó", "o").replace("ú", "u").replace("ñ", "n")
    )
    return re.sub(r"\s+", " ", t)


def _to_float(val: object) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    t = str(val).strip().replace("$", "").replace("U$S", "").replace("USD", "").replace(" ", "")
    if not t or t.lower() in {"nan", "none", "-"}:
        return None
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    if re.search(r"\d\.\d{3}", t) and "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t and "." not in t:
        t = t.replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    try:
        v = float(t)
        return -v if neg else v
    except ValueError:
        return None


def _parse_fecha(val: object) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    ts = pd.to_datetime(str(val), dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.date()


def _leer_bytes(uploaded) -> tuple[str, bytes]:
    nombre = getattr(uploaded, "name", None) or "archivo"
    if hasattr(uploaded, "getvalue"):
        return nombre, uploaded.getvalue()
    if hasattr(uploaded, "read"):
        data = uploaded.read()
        if hasattr(uploaded, "seek"):
            uploaded.seek(0)
        return nombre, data if isinstance(data, (bytes, bytearray)) else bytes(data)
    if isinstance(uploaded, (bytes, bytearray)):
        return nombre, bytes(uploaded)
    raise TypeError(f"No se pudo leer: {type(uploaded)}")


def _texto_pdf(data: bytes) -> str:
    partes: list[str] = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                partes.append(page.extract_text() or "")
    except Exception:
        try:
            import fitz
            with fitz.open(stream=data, filetype="pdf") as doc:
                for page in doc:
                    partes.append(page.get_text() or "")
        except Exception:
            return ""
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Clasificación por especie
# ---------------------------------------------------------------------------

def clasificar_grupo_especie(especie: str, descripcion: str = "", tipo_op: str = "") -> str:
    """Asigna grupo de trabajo (bonos / FCI / MEP-USD / acciones / otros)."""
    texto = f"{especie} {descripcion} {tipo_op}"
    n = _norm(texto)
    esp = re.sub(r"[^A-Za-z0-9]", "", str(especie or "")).upper()

    if any(k in n for k in ("dolar mep", "dólar mep", "compra mep", "venta mep", " billete", "u$s", "usd")):
        if "fci" not in n and "fondo" not in n:
            # AL30/GD30 usados para MEP → bonos; USD resultante → dólar
            if esp in TITULOS_PUBLICOS or any(t in esp for t in TITULOS_PUBLICOS):
                if "mep" in n or "dolar" in n or "dólar" in n:
                    return "Dólar / MEP"
            if esp in {"USD", "U$S", "DOLAR", "DOLARES"} or n.strip() in {"usd", "u$s", "dolar"}:
                return "Dólar / MEP"
            if "mep" in n and ("dolar" in n or "dólar" in n or "usd" in n):
                return "Dólar / MEP"

    for t in TITULOS_PUBLICOS:
        if esp == t or esp.startswith(t) or t in esp:
            return "Bonos / Títulos públicos"
    if any(k in n for k in ("bono", "titulo publico", "título público", "lecap", "letra del tesoro", "bopreal")):
        return "Bonos / Títulos públicos"

    if any(k in n for k in ("fci", "fondo comun", "fondo común", "cuotaparte", "money market", "fima", "premium clase")):
        return "FCI"
    if "fima" in n or re.search(r"\bfondo\b", n):
        return "FCI"

    if any(k in n for k in ("cedear", "accion", "acción", "equity", "adr")):
        return "Acciones / Cedears"
    # Tickers cortos tipo acciones (heurística suave)
    if re.fullmatch(r"[A-Z]{3,5}", esp) and esp not in TITULOS_PUBLICOS:
        if not any(k in n for k in ("fci", "fondo", "on ", "obligacion")):
            return "Acciones / Cedears"

    if any(k in n for k in ("obligacion negociable", "obligación negociable", "\bon\b")):
        return "Otros"

    return "Otros"


def normalizar_tipo_operacion(raw: str, cantidad: float | None = None, monto: float | None = None) -> str:
    n = _norm(raw)
    if any(k in n for k in _COMPRA_KW) and not any(k in n for k in ("venta", "rescate")):
        if "rescate" in n:
            return "Rescate"
        if "suscrip" in n:
            return "Suscripcion"
        return "Compra"
    if any(k in n for k in _VENTA_KW):
        if "rescate" in n:
            return "Rescate"
        return "Venta"
    if "dividendo" in n or "cupon" in n or "cupón" in n or "renta" in n:
        return "Renta"
    if "amortiz" in n:
        return "Amortizacion"
    # Por signo de cantidad/monto
    if cantidad is not None and cantidad < 0:
        return "Venta"
    if monto is not None and monto < 0:
        return "Venta"
    if cantidad is not None and cantidad > 0:
        return "Compra"
    return (raw or "Movimiento").strip() or "Movimiento"


# ---------------------------------------------------------------------------
# Lectura Excel / PDF → movimientos
# ---------------------------------------------------------------------------

def _mapear_columnas(df: pd.DataFrame) -> pd.DataFrame:
    cols = {_norm(c): c for c in df.columns}
    ren: dict[str, str] = {}
    for canon, aliases in _ALIAS_COLS.items():
        for a in aliases:
            if a in cols:
                ren[cols[a]] = canon
                break
    return df.rename(columns=ren)


def _df_a_movimientos(df: pd.DataFrame, archivo: str) -> list[dict]:
    if df is None or df.empty:
        return []
    mapped = _mapear_columnas(df)
    out: list[dict] = []
    for _, row in mapped.iterrows():
        fecha = _parse_fecha(row.get("Fecha"))
        especie = str(row.get("Especie") or "").strip()
        if not especie and "Descripcion" in mapped.columns:
            especie = str(row.get("Descripcion") or "").strip()[:80]
        if not fecha and not especie:
            continue
        cant = _to_float(row.get("Cantidad"))
        monto = _to_float(row.get("Monto_Total"))
        precio = _to_float(row.get("Precio"))
        tipo_raw = str(row.get("Tipo_Operacion") or "")
        tipo = normalizar_tipo_operacion(tipo_raw, cant, monto)
        desc = str(row.get("Descripcion") or "")
        # Cantidad siempre positiva; el sentido está en Tipo_Operacion
        cant_abs = abs(cant) if cant is not None else None
        monto_abs = abs(monto) if monto is not None else None
        if precio is None and cant_abs and monto_abs and cant_abs > 0:
            precio = round(monto_abs / cant_abs, 6)
        if monto_abs is None and cant_abs and precio:
            monto_abs = round(cant_abs * precio, 2)
        grupo = clasificar_grupo_especie(especie, desc, tipo)
        out.append({
            "Fecha": fecha.strftime("%d/%m/%Y") if fecha else "",
            "_fecha": fecha or date.min,
            "Especie": especie or "Sin especie",
            "Grupo": grupo,
            "Tipo_Operacion": tipo,
            "Cantidad": cant_abs,
            "Precio": precio,
            "Monto_Total": monto_abs,
            "Moneda": str(row.get("Moneda") or "ARS").strip() or "ARS",
            "Descripcion": desc,
            "Archivo origen": archivo,
            "Nueva_Clasificacion": None,
        })
    return out


def _movimientos_desde_excel(data: bytes, archivo: str) -> list[dict]:
    movs: list[dict] = []
    try:
        xl = pd.ExcelFile(io.BytesIO(data))
    except Exception:
        return movs
    for sh in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sh, dtype=object)
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        movs.extend(_df_a_movimientos(df, f"{archivo}::{sh}"))
    return movs


_RE_LINEA_MOV = re.compile(
    r"(?P<fecha>\d{1,2}[/-]\d{1,2}[/-]\d{2,4}).{0,40}?"
    r"(?P<tipo>compra|venta|suscrip\w*|rescate|dividendo|cup[oó]n|amortiz\w*)?"
    r".{0,60}?"
    r"(?P<especie>[A-Z]{2,6}\d{0,2}[A-Z]?|[A-Za-z][A-Za-z0-9 .\-]{2,40})",
    re.IGNORECASE,
)

_RE_GALICIA_USD_MOV = re.compile(
    r"^(?P<fecha>\d{1,2}/\d{1,2}/\d{2,4})\s+"
    r"(?P<desc>.+?)\s+"
    r"(?P<importe>-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2})\s+"
    r"(?P<saldo>-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2})\s*$"
)

_RE_GALICIA_FCI_MOV = re.compile(
    r"^(?P<fecha>\d{1,2}/\d{1,2}/\d{2,4})\s+"
    r"(?P<tipo>SUSCRIPCION|RESCATE|SUSCRIPCIÓN)\s+"
    r"(?P<cant>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+)\s+"
    r"(?:\$|USD|U\s*SD)\s*(?P<px>\d{1,3}(?:\.\d{3})*,\d+|\d+,\d+|\d+\.\d+)\s+"
    r"(?:\$|USD|U\s*SD)\s*(?P<monto>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s+"
    r"(?P<liq>\d{1,2}/\d{1,2}/\d{2,4})?",
    re.IGNORECASE,
)


def _es_pdf_galicia_caja_usd(texto: str) -> bool:
    n = _norm(texto[:800])
    return "caja de ahorro en dolares" in n or "caja de ahorro en dólares" in n


def _es_pdf_galicia_inversiones(texto: str) -> bool:
    n = _norm(texto[:1200])
    return (
        "cuenta comitente" in n
        or "fima-fondos" in n
        or "movimientos / operaciones" in n
        or "fondos comunes de inversion" in n
    )


def _movimientos_galicia_caja_usd(texto: str, archivo: str) -> list[dict]:
    """Extracto Galicia CA en dólares: Fecha Desc Importe(+/-) Saldo."""
    movs: list[dict] = []
    if "sin movimientos" in _norm(texto):
        return movs
    for raw in (texto or "").splitlines():
        ln = re.sub(r"[ \t]+", " ", raw).strip()
        m = _RE_GALICIA_USD_MOV.match(ln)
        if not m:
            continue
        fecha = _parse_fecha(m.group("fecha"))
        if not fecha:
            continue
        desc = m.group("desc").strip()
        # Quitar códigos de origen cortos al final del desc (ej. 0340)
        desc = re.sub(r"\s+\d{3,6}$", "", desc).strip()
        imp = _to_float(m.group("importe"))
        if imp is None or abs(imp) < 0.0001:
            continue
        tipo = "Compra" if imp > 0 else "Venta"
        # Suscripción FIMA desde CA USD = egreso de dólares (sigue siendo salida de caja USD)
        if "suscrip" in _norm(desc):
            tipo = "Venta"
        elif "rescate" in _norm(desc):
            tipo = "Compra"
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": "USD",
            "Grupo": "Dólar / MEP",
            "Tipo_Operacion": tipo,
            "Cantidad": abs(imp),
            "Precio": 1.0,
            "Monto_Total": abs(imp),
            "Moneda": "USD",
            "Descripcion": desc,
            "Archivo origen": archivo,
            "Nueva_Clasificacion": None,
        })
    return movs


def _movimientos_galicia_resumen_inversiones(texto: str, archivo: str) -> list[dict]:
    """Resumen mensual Galicia Inversiones: SUSCRIPCION/RESCATE FCI."""
    movs: list[dict] = []
    fondo_actual = ""
    for raw in (texto or "").splitlines():
        ln = re.sub(r"[ \t]+", " ", raw).strip()
        # Cabecera de fondo
        m_fondo = re.match(r"(?i)^FONDO\s*-\s*(.+)$", ln)
        if m_fondo:
            fondo_actual = m_fondo.group(1).strip()
            continue
        m = _RE_GALICIA_FCI_MOV.match(ln)
        if not m:
            continue
        fecha = _parse_fecha(m.group("fecha"))
        if not fecha:
            continue
        tipo_raw = m.group("tipo")
        tipo = "Suscripcion" if "suscrip" in _norm(tipo_raw) else "Rescate"
        cant = _to_float(m.group("cant"))
        px = _to_float(m.group("px"))
        monto = _to_float(m.group("monto"))
        if not cant or cant <= 0:
            continue
        especie = fondo_actual or "FCI"
        mon = "USD" if re.search(r"(?i)\bUSD\b|U\s*SD", ln) else "ARS"
        if "dolar" in _norm(especie) or "dolares" in _norm(especie):
            mon = "USD"
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": especie,
            "Grupo": "FCI",
            "Tipo_Operacion": tipo,
            "Cantidad": abs(cant),
            "Precio": px,
            "Monto_Total": abs(monto) if monto else (round(abs(cant) * float(px or 0), 2) if px else None),
            "Moneda": mon,
            "Descripcion": f"{tipo_raw} {especie}",
            "Archivo origen": archivo,
            "Nueva_Clasificacion": None,
        })
    return movs


def _movimientos_desde_pdf_texto(texto: str, archivo: str) -> list[dict]:
    """Despacha parsers Galicia / heurística genérica."""
    if _es_pdf_galicia_caja_usd(texto):
        got = _movimientos_galicia_caja_usd(texto, archivo)
        if got or "sin movimientos" in _norm(texto):
            return got
    if _es_pdf_galicia_inversiones(texto):
        got = _movimientos_galicia_resumen_inversiones(texto, archivo)
        if got:
            return got

    movs: list[dict] = []
    for raw in (texto or "").splitlines():
        ln = re.sub(r"[ \t]+", " ", raw).strip()
        if len(ln) < 10:
            continue
        m_fecha = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", ln)
        if not m_fecha:
            continue
        fecha = _parse_fecha(m_fecha.group(1))
        if not fecha:
            continue
        low = _norm(ln)
        # Detectar especie
        especie = ""
        for t in sorted(TITULOS_PUBLICOS, key=len, reverse=True):
            if re.search(rf"\b{re.escape(t)}\b", ln, re.I):
                especie = t
                break
        if not especie:
            m_esp = re.search(
                r"\b(FIMA[A-Z0-9 ]{0,30}|CEDEAR\s+\w+|[A-Z]{3,5})\b",
                ln,
            )
            if m_esp:
                especie = m_esp.group(1).strip()
        if not especie:
            continue
        nums = []
        for m in re.finditer(
            r"(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2}|\d+)(?!\d)",
            ln[m_fecha.end():],
        ):
            v = _to_float(m.group(1))
            if v is None or v <= 0:
                continue
            if 1900 <= v <= 2100 and "," not in m.group(1) and "." not in m.group(1):
                continue
            nums.append(v)
        if not nums:
            continue
        cant = nums[0]
        monto = nums[-1] if len(nums) >= 2 else None
        precio = None
        if cant and monto and cant > 0:
            if len(nums) >= 3:
                precio = nums[1]
            elif monto > cant * 1.5:
                precio = round(monto / cant, 6)
            else:
                precio, monto = monto, round(cant * monto, 2)
        tipo = normalizar_tipo_operacion(ln, cant, monto)
        grupo = clasificar_grupo_especie(especie, ln, tipo)
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": especie,
            "Grupo": grupo,
            "Tipo_Operacion": tipo,
            "Cantidad": abs(cant) if cant else None,
            "Precio": precio,
            "Monto_Total": abs(monto) if monto else None,
            "Moneda": "USD" if ("usd" in low or "u$s" in low or "mep" in low) else "ARS",
            "Descripcion": ln[:180],
            "Archivo origen": archivo,
            "Nueva_Clasificacion": None,
        })
    return movs


def procesar_archivos_inversiones(archivos) -> tuple[pd.DataFrame, list[dict]]:
    """Paso 1: PDF/Excel → DataFrame de movimientos normalizados."""
    errores: list[dict] = []
    filas: list[dict] = []
    for uploaded in archivos or []:
        try:
            nombre, data = _leer_bytes(uploaded)
        except Exception as e:
            errores.append({"archivo": str(uploaded), "motivo": str(e)})
            continue
        low = nombre.lower()
        try:
            if low.endswith((".xlsx", ".xls", ".xlsm", ".csv")):
                if low.endswith(".csv"):
                    df = pd.read_csv(io.BytesIO(data), dtype=object)
                    filas.extend(_df_a_movimientos(df, nombre))
                else:
                    filas.extend(_movimientos_desde_excel(data, nombre))
            elif low.endswith(".pdf"):
                texto = _texto_pdf(data)
                if not texto.strip():
                    errores.append({"archivo": nombre, "motivo": "PDF sin texto extractable (¿escaneado?)"})
                    continue
                got = _movimientos_desde_pdf_texto(texto, nombre)
                if not got:
                    errores.append({
                        "archivo": nombre,
                        "motivo": "No se detectaron movimientos. Preferí Excel del AlyC si el PDF no es tabular.",
                    })
                filas.extend(got)
            else:
                errores.append({"archivo": nombre, "motivo": "Formato no soportado (PDF/Excel/CSV)."})
        except Exception as e:
            errores.append({"archivo": nombre, "motivo": str(e)})

    if not filas:
        return pd.DataFrame(columns=COLUMNAS_MOV), errores
    df = pd.DataFrame(filas)
    df = df.sort_values(["_fecha", "Especie"], kind="stable").reset_index(drop=True)
    # Reaplicar clasificación (permite Nueva_Clasificacion futura)
    df["Grupo"] = [
        clasificar_grupo_especie(e, d, t)
        for e, d, t in zip(df["Especie"], df["Descripcion"], df["Tipo_Operacion"])
    ]
    return df, errores


def reclasificar_movimientos(df: pd.DataFrame) -> pd.DataFrame:
    """Paso 2: asegura Grupo; respeta Nueva_Clasificacion si el usuario la completó."""
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_MOV)
    out = df.copy()
    grupos = []
    for _, row in out.iterrows():
        nueva = str(row.get("Nueva_Clasificacion") or "").strip()
        if nueva and nueva in GRUPOS_ESPECIE:
            grupos.append(nueva)
        else:
            grupos.append(
                clasificar_grupo_especie(
                    str(row.get("Especie") or ""),
                    str(row.get("Descripcion") or ""),
                    str(row.get("Tipo_Operacion") or ""),
                )
            )
    out["Grupo"] = grupos
    return out


# ---------------------------------------------------------------------------
# Saldo inicial desde DDJJ
# ---------------------------------------------------------------------------

_RE_TENENCIA = re.compile(
    r"(?P<esp>AL30D?|GD30D?|AL29|GD29|AL35|GD35|AL41|GD41|AE38|GE38|TX26|TX28|"
    r"FIMA[A-Z0-9 ]{0,24}|CEDEAR\s+\w+|[A-Z]{3,5})"
    r".{0,40}?"
    r"(?P<cant>\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?)",
    re.IGNORECASE,
)


def extraer_saldo_inicial_ddjj_pdf(data: bytes, archivo: str = "ddjj.pdf") -> tuple[pd.DataFrame, list[str]]:
    """
    Heurística sobre PDF de DDJJ / papeles BIENES (F.711):
    FCI (Fima), depósitos USD y tickers sueltos.
    """
    avisos: list[str] = []
    texto = _texto_pdf(data)
    if not texto.strip():
        avisos.append("DDJJ sin texto extractable.")
        return pd.DataFrame(
            columns=["Especie", "Grupo", "Cantidad", "Costo_Unitario", "Costo_Total", "Moneda", "Origen"]
        ), avisos

    filas: list[dict] = []
    vistos: set[str] = set()

    # FCI en papeles BIENES (importe suele ir en línea siguiente al rótulo)
    bloques = re.split(r"(?i)Fondos comunes de Inversi.", texto)
    for bloque in bloques[1:]:
        imps = re.findall(
            r"Importe al 31/12/(20\d{2}).{0,180}?\$\s*([\d.]+,\d{2})",
            bloque,
            flags=re.I | re.S,
        )
        if not imps:
            continue
        imps_sorted = sorted(imps, key=lambda x: x[0], reverse=True)
        costo = _to_float(imps_sorted[0][1])
        m_desc = re.search(r"(?i)(Fima[^\n]{0,60}|FIMA[^\n]{0,60})", bloque)
        if not m_desc or not costo or costo < 10:
            continue
        esp = re.sub(r"\s+", " ", m_desc.group(1)).strip(" -")
        esp = re.sub(r"\s+\d+\s*$", "", esp).strip()
        # Normalizar nombre
        if "premium" in _norm(esp):
            esp = "FIMA PREMIUM CLASE A"
        elif "ahorro" in _norm(esp):
            esp = "FIMA AHORRO PESOS CLASE A"
        elif "renta fija" in _norm(esp) and "dolar" in _norm(esp):
            esp = "FIMA RENTA FIJA DOLARES CLASE A"
        key = f"FCI|{esp.upper()}"
        if key in vistos:
            continue
        vistos.add(key)
        # Evitar "Cantidad" basura del formulario AFIP (IDs tipo 25, 0.01)
        cant = None
        for m_cant in re.finditer(
            r"(?i)(?<!ID Clase / Tipo )Cantidad(?:\s+nominal)?[^\d]{0,20}([\d.]+(?:,\d+)?)",
            bloque,
        ):
            c = _to_float(m_cant.group(1))
            # Cuotapartes reales suelen ser >100 o con decimales; IDs son enteros chicos
            if c and (c >= 100 or (c >= 1 and not float(c).is_integer())):
                cant = c
                break
        if cant is None:
            cant = 1.0
            cu = costo
            avisos.append(
                f"{esp}: valuación ${costo:,.2f} sin cuotas claras; "
                "lote inicial = 1 @ costo total (revisar / estimar cuotas)."
            )
        else:
            cu = round(costo / cant, 6)
        filas.append({
            "Especie": esp,
            "Grupo": "FCI",
            "Cantidad": cant,
            "Costo_Unitario": cu,
            "Costo_Total": costo,
            "Moneda": "ARS",
            "Origen": archivo,
            "Fecha": "01/01/1900",
        })

    # Depósitos USD (cantidad nominal razonable)
    for m in re.finditer(
        r"(?is)Cantidad nominal de moneda\s*([\d.]+(?:,\d+)?)\s*"
        r"Pa[ií]s\s*CBU[^\n]*\nArgentina[^\n]*\n"
        r".{0,120}?D[OÓ]LAR",
        texto,
    ):
        cant = _to_float(m.group(1))
        if not cant or cant <= 0:
            continue
        ventana = texto[m.start(): m.start() + 500]
        imps = re.findall(
            r"Importe al 31/12/(20\d{2})\s*\$?\s*([\d.]+,\d{2})",
            ventana,
            flags=re.I,
        )
        costo_ars = None
        if imps:
            costo_ars = _to_float(sorted(imps, key=lambda x: x[0], reverse=True)[0][1])
        cu = round(costo_ars / cant, 6) if costo_ars and cant else 1.0
        key = f"USD|{cant}"
        if key in vistos:
            continue
        vistos.add(key)
        filas.append({
            "Especie": "USD",
            "Grupo": "Dólar / MEP",
            "Cantidad": cant,
            "Costo_Unitario": cu if costo_ars else 1.0,
            "Costo_Total": costo_ars if costo_ars else cant,
            "Moneda": "USD",
            "Origen": archivo,
            "Fecha": "01/01/1900",
        })

    # Fallback tickers SOLO si el archivo parece un listado de títulos (no papeles BIENES ruidosos)
    es_bienes = "papelestrabajobienes" in _norm(texto) or "bienes inmuebles" in _norm(texto)
    if not es_bienes:
        for m in _RE_TENENCIA.finditer(texto):
            esp = re.sub(r"\s+", " ", m.group("esp").strip().upper())
            if esp in {"IVA", "CUIT", "AFIP", "ARCA", "PDF", "TOTAL", "ANEXO"}:
                continue
            if len(esp) < 3:
                continue
            cant = _to_float(m.group("cant"))
            if not cant or cant <= 0:
                continue
            key = f"{esp}|{cant}"
            if key in vistos:
                continue
            vistos.add(key)
            filas.append({
                "Especie": esp,
                "Grupo": clasificar_grupo_especie(esp),
                "Cantidad": cant,
                "Costo_Unitario": None,
                "Costo_Total": None,
                "Moneda": "ARS",
                "Origen": archivo,
                "Fecha": "01/01/1900",
            })

    if not filas:
        avisos.append(
            "No se detectaron tenencias automáticamente. "
            "Usá BIENES.pdf o un Excel de saldo inicial."
        )
    else:
        avisos.append(f"Se detectaron {len(filas)} tenencia(s) tentativas (revisar).")
    return pd.DataFrame(filas), avisos


def leer_saldo_inicial_excel(data: bytes, archivo: str = "saldo.xlsx") -> pd.DataFrame:
    """Excel manual: Especie, Cantidad, Costo_Unitario o Costo_Total, Moneda opcional."""
    try:
        df = pd.read_excel(io.BytesIO(data), dtype=object)
    except Exception:
        return pd.DataFrame()
    df = _mapear_columnas(df)
    # Alias costo
    cols = {_norm(c): c for c in df.columns}
    if "Costo_Unitario" not in df.columns:
        for a in ("costo unitario", "costo", "precio", "px costo"):
            if a in cols:
                df = df.rename(columns={cols[a]: "Costo_Unitario"})
                break
    if "Costo_Total" not in df.columns:
        for a in ("costo total", "importe", "monto", "valor"):
            if a in cols and cols[a] != "Costo_Unitario":
                df = df.rename(columns={cols[a]: "Costo_Total"})
                break
    out = []
    for _, row in df.iterrows():
        esp = str(row.get("Especie") or "").strip()
        if not esp:
            continue
        cant = _to_float(row.get("Cantidad"))
        if not cant or cant <= 0:
            continue
        cu = _to_float(row.get("Costo_Unitario"))
        ct = _to_float(row.get("Costo_Total"))
        if cu is None and ct is not None:
            cu = round(ct / cant, 6)
        if ct is None and cu is not None:
            ct = round(cant * cu, 2)
        out.append({
            "Especie": esp,
            "Grupo": clasificar_grupo_especie(esp),
            "Cantidad": cant,
            "Costo_Unitario": cu,
            "Costo_Total": ct,
            "Moneda": str(row.get("Moneda") or "ARS"),
            "Origen": archivo,
            "Fecha": "01/01/1900",
        })
    return pd.DataFrame(out)


# ---------------------------------------------------------------------------
# Motor FIFO / PEPS
# ---------------------------------------------------------------------------

@dataclass
class LoteFifo:
    fecha: date
    cantidad: float
    costo_unitario: float
    origen: str = ""


@dataclass
class ResultadoFifo:
    aplicaciones: list[dict] = field(default_factory=list)
    saldos: list[dict] = field(default_factory=list)
    movimientos: list[dict] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


def _clave_especie(especie: str, moneda: str = "ARS") -> str:
    return f"{str(especie).strip().upper()}|{str(moneda).strip().upper()}"


def aplicar_fifo(
    df_mov: pd.DataFrame,
    df_inicial: pd.DataFrame | None = None,
) -> ResultadoFifo:
    """
    FIFO por especie (+ moneda):
    - Compra / Suscripcion → entra lote
    - Venta / Rescate → sale de los lotes más viejos
    """
    res = ResultadoFifo()
    colas: dict[str, deque[LoteFifo]] = defaultdict(deque)

    # Sembrar saldo inicial
    if df_inicial is not None and not df_inicial.empty:
        for _, row in df_inicial.iterrows():
            esp = str(row.get("Especie") or "").strip()
            if not esp:
                continue
            cant = float(row.get("Cantidad") or 0)
            if cant <= 0:
                continue
            mon = str(row.get("Moneda") or "ARS")
            cu = row.get("Costo_Unitario")
            ct = row.get("Costo_Total")
            if cu is None or (isinstance(cu, float) and pd.isna(cu)):
                cu = (float(ct) / cant) if ct not in (None, 0) and not (isinstance(ct, float) and pd.isna(ct)) else 0.0
            else:
                cu = float(cu)
            key = _clave_especie(esp, mon)
            colas[key].append(LoteFifo(date(1900, 1, 1), cant, cu, "Saldo inicial"))
            res.movimientos.append({
                "Fecha": "Saldo inicial",
                "Especie": esp,
                "Grupo": row.get("Grupo") or clasificar_grupo_especie(esp),
                "Tipo_Operacion": "Saldo inicial",
                "Cantidad": cant,
                "Precio": cu,
                "Monto_Total": round(cant * cu, 2),
                "Moneda": mon,
                "Costo_Aplicado": round(cant * cu, 2),
                "Resultado": 0.0,
                "Detalle_FIFO": "Semilla DDJJ/manual",
            })

    work = reclasificar_movimientos(df_mov) if df_mov is not None else pd.DataFrame()
    if not work.empty:
        work = work.copy()
        if "_fecha" not in work.columns:
            work["_fecha"] = work["Fecha"].map(lambda x: _parse_fecha(x) or date.min)
        work = work.sort_values(["_fecha", "Especie"], kind="stable")

    for _, row in work.iterrows():
        esp = str(row.get("Especie") or "").strip()
        if not esp:
            continue
        mon = str(row.get("Moneda") or "ARS")
        key = _clave_especie(esp, mon)
        tipo = str(row.get("Tipo_Operacion") or "")
        cant = float(row.get("Cantidad") or 0)
        if cant <= 0:
            continue
        precio = row.get("Precio")
        monto = row.get("Monto_Total")
        if precio is None or (isinstance(precio, float) and pd.isna(precio)):
            precio = (float(monto) / cant) if monto not in (None, 0) and not pd.isna(monto) else 0.0
        else:
            precio = float(precio)
        if monto is None or (isinstance(monto, float) and pd.isna(monto)):
            monto = round(cant * precio, 2)
        else:
            monto = float(monto)
        fecha = row.get("_fecha") or _parse_fecha(row.get("Fecha")) or date.min
        fecha_txt = fecha.strftime("%d/%m/%Y") if isinstance(fecha, date) and fecha.year > 1900 else str(row.get("Fecha") or "")

        es_entrada = tipo.lower() in {
            "compra", "suscripcion", "suscripción", "aporte", "ingreso", "saldo inicial",
        }
        es_salida = tipo.lower() in {"venta", "rescate", "egreso", "amortizacion", "amortización"}

        if es_entrada or (not es_salida and tipo.lower() not in {"renta", "dividendo"}):
            if not es_salida:
                colas[key].append(LoteFifo(fecha if isinstance(fecha, date) else date.min, cant, precio, tipo))
                res.movimientos.append({
                    "Fecha": fecha_txt,
                    "Especie": esp,
                    "Grupo": row.get("Grupo"),
                    "Tipo_Operacion": tipo,
                    "Cantidad": cant,
                    "Precio": precio,
                    "Monto_Total": monto,
                    "Moneda": mon,
                    "Costo_Aplicado": round(cant * precio, 2),
                    "Resultado": 0.0,
                    "Detalle_FIFO": "Alta de lote",
                })
                continue

        # Salida FIFO
        restante = cant
        costo_total = 0.0
        detalle_partes: list[str] = []
        while restante > 1e-12:
            if not colas[key]:
                # Sin stock: costo = precio de salida (resultado 0 sobre faltante)
                costo_total += restante * precio
                detalle_partes.append(f"SIN_STOCK:{restante:.4f}@{precio:.6f}")
                res.avisos.append(
                    f"{fecha_txt} {esp}: venta/rescate sin stock suficiente "
                    f"(faltan {restante:.4f})."
                )
                restante = 0.0
                break
            lote = colas[key][0]
            toma = min(lote.cantidad, restante)
            costo_total += toma * lote.costo_unitario
            detalle_partes.append(
                f"{lote.fecha.strftime('%d/%m/%Y') if lote.fecha.year > 1900 else 'INI'}:"
                f"{toma:.4f}@{lote.costo_unitario:.6f}"
            )
            res.aplicaciones.append({
                "Fecha_salida": fecha_txt,
                "Especie": esp,
                "Grupo": row.get("Grupo"),
                "Cantidad": round(toma, 6),
                "Costo_unitario_lote": lote.costo_unitario,
                "Costo_parcial": round(toma * lote.costo_unitario, 2),
                "Precio_salida": precio,
                "Fecha_lote": lote.fecha.strftime("%d/%m/%Y") if lote.fecha.year > 1900 else "Saldo inicial",
                "Origen_lote": lote.origen,
                "Moneda": mon,
            })
            lote.cantidad -= toma
            restante -= toma
            if lote.cantidad <= 1e-12:
                colas[key].popleft()

        resultado = round(float(monto) - costo_total, 2)
        res.movimientos.append({
            "Fecha": fecha_txt,
            "Especie": esp,
            "Grupo": row.get("Grupo"),
            "Tipo_Operacion": tipo,
            "Cantidad": cant,
            "Precio": precio,
            "Monto_Total": monto,
            "Moneda": mon,
            "Costo_Aplicado": round(costo_total, 2),
            "Resultado": resultado,
            "Detalle_FIFO": " | ".join(detalle_partes),
        })

    # Saldos de cierre
    for key, cola in colas.items():
        esp, mon = key.split("|", 1)
        cant = sum(l.cantidad for l in cola)
        if cant <= 1e-12:
            continue
        costo = sum(l.cantidad * l.costo_unitario for l in cola)
        res.saldos.append({
            "Especie": esp,
            "Grupo": clasificar_grupo_especie(esp),
            "Cantidad": round(cant, 6),
            "Costo_Total": round(costo, 2),
            "Costo_Unitario_Promedio": round(costo / cant, 6) if cant else 0.0,
            "Moneda": mon,
            "Lotes": len(cola),
        })
    res.saldos = sorted(res.saldos, key=lambda r: (r["Grupo"], r["Especie"]))
    return res


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

def exportar_inversiones_excel(
    df_mov: pd.DataFrame,
    resultado: ResultadoFifo,
    df_inicial: pd.DataFrame | None = None,
    meta: dict | None = None,
) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils.dataframe import dataframe_to_rows

    from excel_formato_estudio import COLOR_PRIMARIO, HDR_FONT

    meta = meta or {}
    header_font = HDR_FONT
    body_font = Font(name="Calibri", size=11)
    fill_h = PatternFill("solid", fgColor=COLOR_PRIMARIO)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    def _style(ws, money_cols=()):
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = fill_h
            cell.border = thin
        for row in ws.iter_rows(min_row=2, max_row=max(2, ws.max_row), max_col=ws.max_column):
            for cell in row:
                cell.font = body_font
                cell.border = thin
                if cell.column in money_cols and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00;[Red]-#,##0.00'
                    cell.alignment = Alignment(horizontal="right")

    wb = Workbook()
    # Resumen
    ws = wb.active
    ws.title = "Resumen"
    resumen = pd.DataFrame([
        {"Campo": "Cliente / nota", "Valor": meta.get("nota") or ""},
        {"Campo": "Movimientos", "Valor": len(df_mov) if df_mov is not None else 0},
        {"Campo": "Especies con saldo", "Valor": len(resultado.saldos)},
        {"Campo": "Avisos FIFO", "Valor": len(resultado.avisos)},
        {"Campo": "Metodo", "Valor": "FIFO / PEPS por especie"},
        {"Campo": "Generado", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M")},
    ])
    for r in dataframe_to_rows(resumen, index=False, header=True):
        ws.append(r)
    _style(ws)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 60

    # Movimientos normalizados
    ws2 = wb.create_sheet("Movimientos")
    cols = [c for c in COLUMNAS_MOV if c in (df_mov.columns if df_mov is not None else [])]
    df_m = df_mov[cols].copy() if df_mov is not None and not df_mov.empty else pd.DataFrame(columns=COLUMNAS_MOV)
    for r in dataframe_to_rows(df_m, index=False, header=True):
        ws2.append(r)
    _style(ws2, money_cols={5, 6, 7})
    for i, w in enumerate([12, 22, 24, 14, 12, 12, 14, 10, 40, 22, 18], start=1):
        from openpyxl.utils import get_column_letter
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    # Saldo inicial
    ws_i = wb.create_sheet("Saldo_inicial")
    df_i = df_inicial if df_inicial is not None else pd.DataFrame()
    if df_i.empty:
        ws_i.append(["Especie", "Grupo", "Cantidad", "Costo_Unitario", "Costo_Total", "Moneda", "Origen"])
    else:
        for r in dataframe_to_rows(df_i, index=False, header=True):
            ws_i.append(r)
    _style(ws_i, money_cols={3, 4, 5})

    # FIFO aplicaciones + movimientos valuados
    ws_f = wb.create_sheet("FIFO_aplicaciones")
    df_ap = pd.DataFrame(resultado.aplicaciones)
    if df_ap.empty:
        ws_f.append(["Fecha_salida", "Especie", "Cantidad", "Costo_unitario_lote", "Costo_parcial", "Precio_salida"])
    else:
        for r in dataframe_to_rows(df_ap, index=False, header=True):
            ws_f.append(r)
    _style(ws_f, money_cols=set(range(3, 12)))

    ws_v = wb.create_sheet("FIFO_movimientos")
    df_v = pd.DataFrame(resultado.movimientos)
    if df_v.empty:
        ws_v.append(["Fecha", "Especie", "Tipo_Operacion", "Cantidad", "Costo_Aplicado", "Resultado"])
    else:
        for r in dataframe_to_rows(df_v, index=False, header=True):
            ws_v.append(r)
    _style(ws_v, money_cols=set(range(4, 12)))

    # Saldos cierre
    ws_s = wb.create_sheet("Saldos_cierre")
    df_s = pd.DataFrame(resultado.saldos)
    if df_s.empty:
        ws_s.append(["Especie", "Grupo", "Cantidad", "Costo_Total", "Costo_Unitario_Promedio", "Moneda"])
    else:
        for r in dataframe_to_rows(df_s, index=False, header=True):
            ws_s.append(r)
    _style(ws_s, money_cols={3, 4, 5})

    # Una hoja por grupo
    if df_mov is not None and not df_mov.empty and "Grupo" in df_mov.columns:
        usados: set[str] = set()
        for grupo in sorted(df_mov["Grupo"].dropna().unique()):
            nombre = re.sub(r'[\\/*?:\[\]]', "-", str(grupo))[:31].strip() or "Grupo"
            base = nombre
            n = 2
            while nombre.lower() in usados:
                suf = f"_{n}"
                nombre = (base[: 31 - len(suf)] + suf)
                n += 1
            usados.add(nombre.lower())
            ws_g = wb.create_sheet(nombre)
            sub = df_mov[df_mov["Grupo"] == grupo]
            cols_g = [c for c in COLUMNAS_MOV if c in sub.columns]
            for r in dataframe_to_rows(sub[cols_g], index=False, header=True):
                ws_g.append(r)
            _style(ws_g, money_cols={5, 6, 7})

    # Avisos
    ws_a = wb.create_sheet("Avisos")
    ws_a.append(["Aviso"])
    for a in resultado.avisos or ["Sin avisos"]:
        ws_a.append([a])
    _style(ws_a)
    ws_a.column_dimensions["A"].width = 100

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def plantilla_saldo_inicial_excel() -> bytes:
    """Excel vacío para cargar tenencias iniciales a mano."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Saldo_inicial"
    ws.append(["Especie", "Cantidad", "Costo_Unitario", "Costo_Total", "Moneda"])
    ws.append(["AL30", 1000, 50.25, None, "ARS"])
    ws.append(["FIMA PREMIUM CLASE A", 1500.5, None, 2500000, "ARS"])
    ws.append(["USD", 5000, 1000, None, "USD"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
