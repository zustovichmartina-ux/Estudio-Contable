"""
Analizador de inversiones — identificar tipo → reglas por tipo → FIFO.

Flujo:
1. Normalizar movimientos (AlyC/broker PDF o Excel).
2. Identificar especie (FCI, bono, acción, USD/MEP, otro).
3. Sembrar saldo inicial desde DDJJ (PDF) o Excel manual.
4. Aplicar reglas del tipo + FIFO y exportar Excel del estudio.
   FCI puro: formato extracto (origen por fórmula + intereses + tenencia con
   VC del último extracto de fondos del mes) vía inversiones_fci_formato.
"""

from __future__ import annotations

import io
import re
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

import pandas as pd

from excel_formato_estudio import exportar_informe_excel
from inversiones_catalogo import (
    CONF_ALTA,
    CONF_REVISAR,
    EVENTO_AMORT,
    EVENTO_ENTRADA,
    EVENTO_INGRESO,
    EVENTO_OMITIR,
    EVENTO_OMITIR_USD,
    EVENTO_REVISAR,
    EVENTO_SALIDA,
    GRUPOS_ESPECIE,
    TIPO_A_GRUPO,
    TIPO_FCI,
    TIPO_LABEL,
    TIPO_OTRO,
    TIPO_USD_MEP,
    TIPOS_INVERSION,
    TITULOS_PUBLICOS,
    IdentificacionEspecie,
    canon_fima,
    evento_para,
    identificar_especie,
    orden_evento,
    tipo_desde_grupo,
)

# ---------------------------------------------------------------------------
# Constantes / catálogos
# ---------------------------------------------------------------------------

COLUMNAS_MOV = [
    "Fecha",
    "Especie",
    "Especie_canonica",
    "Tipo_inversion",
    "Grupo",
    "Confianza",
    "Motivo_id",
    "Tipo_Operacion",
    "Cantidad",
    "Precio",
    "Monto_Total",
    "Moneda",
    "Descripcion",
    "Archivo origen",
    "Nueva_Clasificacion",
]

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
# Clasificación / identificación por especie
# ---------------------------------------------------------------------------

def clasificar_grupo_especie(especie: str, descripcion: str = "", tipo_op: str = "") -> str:
    """Grupo de trabajo a partir del identificador (compatibilidad)."""
    return identificar_especie(especie, descripcion, tipo_op).grupo


def _id_desde_override(row: pd.Series) -> IdentificacionEspecie | None:
    """Solo si el usuario forzó grupo (Nueva_Clasificacion) o tipo manual."""
    nueva = str(row.get("Nueva_Clasificacion") or "").strip()
    tipo = tipo_desde_grupo(nueva) if nueva else None
    forzado = str(row.get("Tipo_forzado") or "").strip().lower()
    if tipo is None and forzado in TIPOS_INVERSION:
        tipo = forzado
    if tipo is None:
        return None
    orig = str(row.get("Especie") or "").strip()
    canon = str(row.get("Especie_canonica") or "").strip() or orig
    if tipo == TIPO_FCI:
        canon = canon_fima(f"{canon} {orig}") or canon
    if tipo == TIPO_USD_MEP:
        canon = "USD"
    return IdentificacionEspecie(
        orig,
        canon,
        tipo,
        TIPO_A_GRUPO[tipo],
        CONF_ALTA,
        "Manual",
    )


def aplicar_edicion_identificacion(df_mov: pd.DataFrame, df_id: pd.DataFrame) -> pd.DataFrame:
    """Aplica canónica/tipo editados en el cuadro Identificación a cada movimiento."""
    if df_mov is None or df_mov.empty or df_id is None or df_id.empty:
        return identificar_movimientos(df_mov)
    out = df_mov.copy()
    mapa: dict[str, tuple[str, str]] = {}
    for _, row in df_id.iterrows():
        orig = str(row.get("Especie") or "").strip()
        if not orig:
            continue
        canon = str(row.get("Especie_canonica") or "").strip() or orig
        tipo = str(row.get("Tipo_inversion") or "").strip().lower()
        if tipo not in TIPOS_INVERSION:
            tipo = tipo_desde_grupo(str(row.get("Grupo") or "")) or ""
        if tipo:
            mapa[orig] = (canon, tipo)
    if "Tipo_forzado" not in out.columns:
        out["Tipo_forzado"] = ""
    for idx, row in out.iterrows():
        orig = str(row.get("Especie") or "").strip()
        if orig not in mapa:
            continue
        canon, tipo = mapa[orig]
        out.at[idx, "Especie_canonica"] = canon
        out.at[idx, "Tipo_forzado"] = tipo
        out.at[idx, "Nueva_Clasificacion"] = TIPO_A_GRUPO.get(tipo, row.get("Nueva_Clasificacion"))
    return identificar_movimientos(out)


def identificar_movimientos(df: pd.DataFrame) -> pd.DataFrame:
    """Completa Especie_canonica / Tipo_inversion / Grupo / Confianza / Motivo_id."""
    if df is None or df.empty:
        out = pd.DataFrame(columns=COLUMNAS_MOV)
        return out
    out = df.copy()
    ids: list[IdentificacionEspecie] = []
    for _, row in out.iterrows():
        forced = _id_desde_override(row)
        if forced is not None:
            ids.append(forced)
            continue
        ids.append(
            identificar_especie(
                str(row.get("Especie") or ""),
                str(row.get("Descripcion") or ""),
                str(row.get("Tipo_Operacion") or ""),
                str(row.get("Moneda") or ""),
            )
        )
    out["Especie_canonica"] = [i.especie_canonica for i in ids]
    out["Tipo_inversion"] = [i.tipo_inversion for i in ids]
    out["Grupo"] = [i.grupo for i in ids]
    out["Confianza"] = [i.confianza for i in ids]
    out["Motivo_id"] = [i.motivo for i in ids]
    return out


def cuadro_identificacion(df: pd.DataFrame) -> pd.DataFrame:
    """Una fila por especie canónica para revisar antes del FIFO."""
    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
                "Especie", "Especie_canonica", "Tipo_inversion", "Tipo",
                "Grupo", "Confianza", "Motivo_id", "Movimientos",
            ]
        )
    work = identificar_movimientos(df)
    g = (
        work.groupby(
            ["Especie", "Especie_canonica", "Tipo_inversion", "Grupo", "Confianza", "Motivo_id"],
            dropna=False,
        )
        .size()
        .reset_index(name="Movimientos")
    )
    g["Tipo"] = g["Tipo_inversion"].map(lambda t: TIPO_LABEL.get(t, t))
    return g[
        ["Especie", "Especie_canonica", "Tipo_inversion", "Tipo", "Grupo", "Confianza", "Motivo_id", "Movimientos"]
    ].sort_values(["Confianza", "Tipo_inversion", "Especie_canonica"], kind="stable")


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
    if "dividendo" in n:
        return "Dividendo"
    if "cupon" in n or "cupón" in n or "renta" in n:
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
        ident = identificar_especie(especie or "Sin especie", desc, tipo)
        out.append({
            "Fecha": fecha.strftime("%d/%m/%Y") if fecha else "",
            "_fecha": fecha or date.min,
            "Especie": especie or "Sin especie",
            "Especie_canonica": ident.especie_canonica,
            "Tipo_inversion": ident.tipo_inversion,
            "Grupo": ident.grupo,
            "Confianza": ident.confianza,
            "Motivo_id": ident.motivo,
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
        ident = identificar_especie("USD", desc, tipo, "USD")
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": "USD",
            "Especie_canonica": ident.especie_canonica,
            "Tipo_inversion": ident.tipo_inversion,
            "Grupo": ident.grupo,
            "Confianza": ident.confianza,
            "Motivo_id": ident.motivo,
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
        ident = identificar_especie(especie, f"{tipo_raw} {especie}", tipo, mon)
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": especie,
            "Especie_canonica": ident.especie_canonica,
            "Tipo_inversion": ident.tipo_inversion,
            "Grupo": ident.grupo,
            "Confianza": ident.confianza,
            "Motivo_id": ident.motivo,
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
        mon = "USD" if ("usd" in low or "u$s" in low or "mep" in low) else "ARS"
        ident = identificar_especie(especie, ln, tipo, mon)
        movs.append({
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "_fecha": fecha,
            "Especie": especie,
            "Especie_canonica": ident.especie_canonica,
            "Tipo_inversion": ident.tipo_inversion,
            "Grupo": ident.grupo,
            "Confianza": ident.confianza,
            "Motivo_id": ident.motivo,
            "Tipo_Operacion": tipo,
            "Cantidad": abs(cant) if cant else None,
            "Precio": precio,
            "Monto_Total": abs(monto) if monto else None,
            "Moneda": mon,
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
    return identificar_movimientos(df), errores


def reclasificar_movimientos(df: pd.DataFrame) -> pd.DataFrame:
    """Paso 2: reidentifica; respeta Nueva_Clasificacion / Tipo_forzado."""
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_MOV)
    return identificar_movimientos(df)


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
        ident = identificar_especie(esp, esp, "Saldo inicial")
        esp = ident.especie_canonica
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
            avisos.append(
                f"{esp}: valuación ${costo:,.2f} sin cuotapartes. "
                "Completá la cantidad en el Excel de saldo inicial; no se siembra lote ficticio."
            )
            filas.append({
                "Especie": esp,
                "Especie_canonica": esp,
                "Tipo_inversion": TIPO_FCI,
                "Grupo": ident.grupo,
                "Cantidad": None,
                "Costo_Unitario": None,
                "Costo_Total": costo,
                "Moneda": "ARS",
                "Origen": archivo,
                "Fecha": "01/01/1900",
                "Revisar": "SI — falta cantidad de cuotapartes",
                "Confianza": CONF_REVISAR,
            })
            continue
        cu = round(costo / cant, 6)
        filas.append({
            "Especie": esp,
            "Especie_canonica": esp,
            "Tipo_inversion": TIPO_FCI,
            "Grupo": ident.grupo,
            "Cantidad": cant,
            "Costo_Unitario": cu,
            "Costo_Total": costo,
            "Moneda": "ARS",
            "Origen": archivo,
            "Fecha": "01/01/1900",
            "Revisar": "",
            "Confianza": CONF_ALTA,
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
            "Especie_canonica": "USD",
            "Tipo_inversion": TIPO_USD_MEP,
            "Grupo": TIPO_A_GRUPO[TIPO_USD_MEP],
            "Cantidad": cant,
            "Costo_Unitario": cu if costo_ars else 1.0,
            "Costo_Total": costo_ars if costo_ars else cant,
            "Moneda": "USD",
            "Origen": archivo,
            "Fecha": "01/01/1900",
            "Revisar": "USD — analizar en Caja USD",
            "Confianza": CONF_ALTA,
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
            ident = identificar_especie(esp)
            filas.append({
                "Especie": ident.especie_canonica,
                "Especie_canonica": ident.especie_canonica,
                "Tipo_inversion": ident.tipo_inversion,
                "Grupo": ident.grupo,
                "Cantidad": cant,
                "Costo_Unitario": None,
                "Costo_Total": None,
                "Moneda": "ARS",
                "Origen": archivo,
                "Fecha": "01/01/1900",
                "Revisar": "" if ident.confianza == CONF_ALTA else "SI — confirmar especie",
                "Confianza": ident.confianza,
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
        ident = identificar_especie(esp, "", "Saldo inicial", str(row.get("Moneda") or "ARS"))
        out.append({
            "Especie": ident.especie_canonica,
            "Especie_canonica": ident.especie_canonica,
            "Tipo_inversion": ident.tipo_inversion,
            "Grupo": ident.grupo,
            "Cantidad": cant,
            "Costo_Unitario": cu,
            "Costo_Total": ct,
            "Moneda": str(row.get("Moneda") or "ARS"),
            "Origen": archivo,
            "Fecha": "01/01/1900",
            "Revisar": "",
            "Confianza": ident.confianza,
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
    tipo_inversion: str = ""


@dataclass
class ResultadoFifo:
    aplicaciones: list[dict] = field(default_factory=list)
    saldos: list[dict] = field(default_factory=list)
    movimientos: list[dict] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    lotes_abiertos: list[dict] = field(default_factory=list)


def _clave_especie(especie: str, moneda: str = "ARS") -> str:
    return f"{str(especie).strip().upper()}|{str(moneda).strip().upper()}"


def _num(val: object, default: float = 0.0) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _fila_valuada(
    *,
    fecha_txt: str,
    esp: str,
    tipo_inv: str,
    grupo: object,
    tipo_op: str,
    cant: float,
    precio: float,
    monto: float,
    mon: str,
    costo: float,
    resultado: float,
    detalle: str,
    estado: str,
    cant_aplicada: float | None = None,
    cant_sin_stock: float = 0.0,
) -> dict:
    return {
        "Fecha": fecha_txt,
        "Especie": esp,
        "Tipo_inversion": tipo_inv,
        "Grupo": grupo,
        "Tipo_Operacion": tipo_op,
        "Cantidad": cant,
        "Cantidad_aplicada": round(cant_aplicada if cant_aplicada is not None else cant, 6),
        "Cantidad_sin_stock": round(cant_sin_stock, 6),
        "Precio": precio,
        "Monto_Total": monto,
        "Moneda": mon,
        "Costo_Aplicado": round(costo, 2),
        "Resultado": round(resultado, 2),
        "Detalle_FIFO": detalle,
        "Estado": estado,
    }


def aplicar_fifo(
    df_mov: pd.DataFrame,
    df_inicial: pd.DataFrame | None = None,
) -> ResultadoFifo:
    """
    Identifica el tipo y recién ahí mueve lotes:
    FCI suscrip/rescate · bono compra/venta/amort/cupón · acción compra/venta/div.
    USD/MEP y 'otro' no entran al FIFO.
    """
    res = ResultadoFifo()
    colas: dict[str, deque[LoteFifo]] = defaultdict(deque)

    if df_inicial is not None and not df_inicial.empty:
        for _, row in df_inicial.iterrows():
            ident = identificar_especie(
                str(row.get("Especie_canonica") or row.get("Especie") or ""),
                str(row.get("Especie") or ""),
                "Saldo inicial",
                str(row.get("Moneda") or "ARS"),
            )
            tipo_inv = str(row.get("Tipo_inversion") or ident.tipo_inversion)
            if tipo_inv == TIPO_USD_MEP:
                res.avisos.append(
                    f"Saldo inicial {ident.especie_canonica}: USD/MEP no se FIFO-ea acá — usá Caja USD."
                )
                continue
            if tipo_inv == TIPO_OTRO:
                res.avisos.append(
                    f"Saldo inicial {ident.especie_canonica}: tipo Otro — confirmar antes de sembrar lote."
                )
                continue
            cant = _num(row.get("Cantidad"))
            if cant <= 0:
                nota = str(row.get("Revisar") or "sin cantidad")
                res.avisos.append(
                    f"Saldo inicial {ident.especie_canonica}: no se sembró lote ({nota})."
                )
                continue
            mon = str(row.get("Moneda") or "ARS")
            cu = row.get("Costo_Unitario")
            ct = row.get("Costo_Total")
            if cu is None or (isinstance(cu, float) and pd.isna(cu)):
                cu = (float(ct) / cant) if ct not in (None, 0) and not (isinstance(ct, float) and pd.isna(ct)) else 0.0
            else:
                cu = float(cu)
            key = _clave_especie(ident.especie_canonica, mon)
            colas[key].append(LoteFifo(date(1900, 1, 1), cant, cu, "Saldo inicial", tipo_inv))
            res.movimientos.append(_fila_valuada(
                fecha_txt="Saldo inicial",
                esp=ident.especie_canonica,
                tipo_inv=tipo_inv,
                grupo=row.get("Grupo") or ident.grupo,
                tipo_op="Saldo inicial",
                cant=cant,
                precio=cu,
                monto=round(cant * cu, 2),
                mon=mon,
                costo=round(cant * cu, 2),
                resultado=0.0,
                detalle="Semilla DDJJ/manual",
                estado="ENTRADA",
            ))

    work = reclasificar_movimientos(df_mov) if df_mov is not None else pd.DataFrame()
    if not work.empty:
        work = work.copy()
        if "_fecha" not in work.columns:
            work["_fecha"] = work["Fecha"].map(lambda x: _parse_fecha(x) or date.min)
        work["_evento"] = [
            evento_para(str(r.get("Tipo_inversion") or ""), str(r.get("Tipo_Operacion") or ""))
            for _, r in work.iterrows()
        ]
        # FCI: respetar el orden del extracto (FIFO ya come lo más viejo).
        # Bonos/acciones: el mismo día, entradas antes que salidas.
        work["_ord_ev"] = [
            0 if str(r.get("Tipo_inversion")) == TIPO_FCI else orden_evento(str(r.get("_evento")))
            for _, r in work.iterrows()
        ]
        work["_idx"] = range(len(work))
        work = work.sort_values(["_fecha", "_ord_ev", "_idx"], kind="stable")

    for _, row in work.iterrows():
        ident_tipo = str(row.get("Tipo_inversion") or TIPO_OTRO)
        esp = str(row.get("Especie_canonica") or row.get("Especie") or "").strip()
        if not esp:
            continue
        mon = str(row.get("Moneda") or "ARS")
        key = _clave_especie(esp, mon)
        tipo_op = str(row.get("Tipo_Operacion") or "")
        evento = str(row.get("_evento") or evento_para(ident_tipo, tipo_op))
        cant = _num(row.get("Cantidad"))
        precio = row.get("Precio")
        monto = row.get("Monto_Total")
        if cant > 0:
            if precio is None or (isinstance(precio, float) and pd.isna(precio)):
                precio = (float(monto) / cant) if monto not in (None, 0) and not pd.isna(monto) else 0.0
            else:
                precio = float(precio)
            if monto is None or (isinstance(monto, float) and pd.isna(monto)):
                monto = round(cant * precio, 2)
            else:
                monto = float(monto)
        else:
            precio = _num(precio)
            monto = _num(monto)
        fecha = row.get("_fecha") or _parse_fecha(row.get("Fecha")) or date.min
        fecha_txt = (
            fecha.strftime("%d/%m/%Y")
            if isinstance(fecha, date) and fecha.year > 1900
            else str(row.get("Fecha") or "")
        )
        grupo = row.get("Grupo")

        if evento == EVENTO_OMITIR_USD:
            res.avisos.append(
                f"{fecha_txt} {esp}: USD/MEP no se analiza acá — usá la herramienta Caja USD."
            )
            res.movimientos.append(_fila_valuada(
                fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
                tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
                costo=0.0, resultado=0.0,
                detalle="Omitido — Caja USD", estado="OMITIDO_USD",
                cant_aplicada=0.0,
            ))
            continue

        if evento == EVENTO_REVISAR:
            res.avisos.append(
                f"{fecha_txt} {esp}: tipo Otro — confirmar identificación antes de FIFO."
            )
            res.movimientos.append(_fila_valuada(
                fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
                tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
                costo=0.0, resultado=0.0,
                detalle="Revisar tipo de inversión", estado="REVISAR",
                cant_aplicada=0.0,
            ))
            continue

        if evento == EVENTO_OMITIR:
            res.avisos.append(
                f"{fecha_txt} {esp}: operación '{tipo_op}' no aplica al tipo {ident_tipo}."
            )
            res.movimientos.append(_fila_valuada(
                fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
                tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
                costo=0.0, resultado=0.0,
                detalle="Operación no aplica al tipo", estado="OMITIDO",
                cant_aplicada=0.0,
            ))
            continue

        if evento == EVENTO_INGRESO:
            res.movimientos.append(_fila_valuada(
                fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
                tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
                costo=0.0, resultado=monto,
                detalle="No toca stock (renta / dividendo / cupón)", estado="INGRESO",
                cant_aplicada=0.0,
            ))
            continue

        if evento == EVENTO_ENTRADA:
            if cant <= 0:
                continue
            colas[key].append(
                LoteFifo(fecha if isinstance(fecha, date) else date.min, cant, precio, tipo_op, ident_tipo)
            )
            res.movimientos.append(_fila_valuada(
                fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
                tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
                costo=round(cant * precio, 2), resultado=0.0,
                detalle="Alta de lote", estado="ENTRADA",
            ))
            continue

        if evento not in {EVENTO_SALIDA, EVENTO_AMORT} or cant <= 0:
            continue

        # Salida / amortización / rescate FCI: FIFO (first in, first out)
        restante = cant
        costo_total = 0.0
        interes_fci = 0.0
        aplicado = 0.0
        detalle_partes: list[str] = []
        es_fci = ident_tipo == TIPO_FCI
        while restante > 1e-12:
            if not colas[key]:
                detalle_partes.append(f"SIN_STOCK:{restante:.4f}")
                res.avisos.append(
                    f"{fecha_txt} {esp}: {'amortización' if evento == EVENTO_AMORT else 'venta/rescate'} "
                    f"sin stock suficiente (faltan {restante:.4f})."
                )
                break
            lote = colas[key][0]
            toma = min(lote.cantidad, restante)
            costo_parte = toma * lote.costo_unitario
            costo_total += costo_parte
            aplicado += toma
            vc_lote = lote.costo_unitario
            interes_parte = (precio - vc_lote) * toma if es_fci else 0.0
            interes_fci += interes_parte
            lote_txt = lote.fecha.strftime("%d/%m/%Y") if lote.fecha.year > 1900 else "INI"
            if es_fci:
                detalle_partes.append(
                    f"({precio:.6f}-{vc_lote:.6f})×{toma:.4f}"
                )
            else:
                detalle_partes.append(f"{lote_txt}:{toma:.4f}@{vc_lote:.6f}")
            res.aplicaciones.append({
                "Fecha_salida": fecha_txt,
                "Especie": esp,
                "Tipo_inversion": ident_tipo,
                "Grupo": grupo,
                "Evento": "Amortizacion" if evento == EVENTO_AMORT else tipo_op,
                "Cantidad": round(toma, 6),
                "Valor_CC_lote": vc_lote,
                "Valor_CC_salida": precio,
                "Costo_unitario_lote": vc_lote,
                "Costo_parcial": round(costo_parte, 2),
                "Interes": round(interes_parte, 2) if es_fci else None,
                "Precio_salida": precio,
                "Fecha_lote": lote_txt if lote_txt != "INI" else "Saldo inicial",
                "Origen_lote": lote.origen,
                "Moneda": mon,
            })
            lote.cantidad -= toma
            restante -= toma
            if lote.cantidad <= 1e-12:
                colas[key].popleft()

        sin_stock = max(restante, 0.0)
        if es_fci:
            # Interés = (VC rescate − VC lote) × cuotas de ese lote
            resultado = round(interes_fci, 2) if aplicado > 0 else 0.0
            estado = "INTERES" if sin_stock <= 1e-12 else "SIN_STOCK"
        else:
            if cant > 0 and aplicado + 1e-12 < cant:
                monto_aplicado = round(monto * (aplicado / cant), 2)
            else:
                monto_aplicado = monto
            resultado = round(float(monto_aplicado) - costo_total, 2) if aplicado > 0 else 0.0
            estado = "AMORTIZACION" if evento == EVENTO_AMORT else "SALIDA"
            if sin_stock > 1e-12:
                estado = "SIN_STOCK"
        res.movimientos.append(_fila_valuada(
            fecha_txt=fecha_txt, esp=esp, tipo_inv=ident_tipo, grupo=grupo,
            tipo_op=tipo_op, cant=cant, precio=precio, monto=monto, mon=mon,
            costo=costo_total, resultado=resultado,
            detalle=" | ".join(detalle_partes) or "SIN_STOCK",
            estado=estado, cant_aplicada=aplicado, cant_sin_stock=sin_stock,
        ))

    for key, cola in colas.items():
        esp, mon = key.split("|", 1)
        cant = sum(l.cantidad for l in cola)
        if cant <= 1e-12:
            continue
        costo = sum(l.cantidad * l.costo_unitario for l in cola)
        tipo_inv = next((l.tipo_inversion for l in cola if l.tipo_inversion), "")
        ident = identificar_especie(esp)
        res.saldos.append({
            "Especie": esp,
            "Tipo_inversion": tipo_inv or ident.tipo_inversion,
            "Grupo": TIPO_A_GRUPO.get(tipo_inv, ident.grupo),
            "Cantidad": round(cant, 6),
            "Costo_Total": round(costo, 2),
            "Costo_Unitario_Promedio": round(costo / cant, 6) if cant else 0.0,
            "Moneda": mon,
            "Lotes": len(cola),
        })
        for lote in cola:
            if lote.cantidad <= 1e-12:
                continue
            fecha_lote = lote.fecha if isinstance(lote.fecha, date) and lote.fecha.year > 1900 else None
            res.lotes_abiertos.append({
                "Especie": esp,
                "Fecha": fecha_lote,
                "Fecha_txt": fecha_lote.strftime("%d/%m/%Y") if fecha_lote else "Saldo inicial",
                "Origen": lote.origen or "",
                "Cantidad": round(lote.cantidad, 6),
                "Costo_Unitario": lote.costo_unitario,
                "Costo_Total": round(lote.cantidad * lote.costo_unitario, 2),
                "Moneda": mon,
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
    meta = meta or {}
    df_m = identificar_movimientos(df_mov) if df_mov is not None and not df_mov.empty else pd.DataFrame()
    df_v = pd.DataFrame(resultado.movimientos)
    df_ap = pd.DataFrame(resultado.aplicaciones)
    df_s = pd.DataFrame(resultado.saldos)
    df_i = df_inicial if df_inicial is not None else pd.DataFrame()
    df_id = cuadro_identificacion(df_m) if not df_m.empty else pd.DataFrame()

    por_tipo = []
    if not df_v.empty and "Tipo_inversion" in df_v.columns:
        tmp = df_v.copy()
        tmp["Resultado_fifo"] = tmp.apply(
            lambda r: r["Resultado"] if str(r.get("Estado")) in {
                "SALIDA", "AMORTIZACION", "INTERES", "SIN_STOCK",
            } else 0.0,
            axis=1,
        )
        g = (
            tmp.groupby("Tipo_inversion", dropna=False)
            .agg(Movimientos=("Especie", "size"), Resultado=("Resultado_fifo", "sum"))
            .reset_index()
        )
        g["Tipo"] = g["Tipo_inversion"].map(lambda t: TIPO_LABEL.get(t, t))
        por_tipo = g[["Tipo", "Movimientos", "Resultado"]]

    kpis: list[tuple[str, Any]] = [
        ("Movimientos", int(len(df_m))),
        ("Especies con saldo", int(len(df_s))),
        ("Avisos", int(len(resultado.avisos))),
        (
            "Resultado FIFO (ventas/rescates)",
            float(df_v.loc[df_v["Estado"].isin(["SALIDA", "AMORTIZACION", "INTERES", "SIN_STOCK"]), "Resultado"].sum())
            if not df_v.empty and "Estado" in df_v.columns
            else 0.0,
        ),
    ]

    cols_mov = [c for c in COLUMNAS_MOV if c in df_m.columns] if not df_m.empty else COLUMNAS_MOV
    avisos_df = pd.DataFrame({"Aviso": resultado.avisos or ["Sin avisos"]})

    hojas: list[tuple[str, pd.DataFrame]] = [
        ("Identificacion", df_id),
        ("Saldo_inicial", df_i if not df_i.empty else pd.DataFrame(columns=["Especie", "Cantidad", "Costo_Total"])),
        ("FIFO", df_ap),
        ("Saldos_cierre", df_s),
        ("Avisos", avisos_df),
    ]
    if not df_m.empty and "Grupo" in df_m.columns:
        for grupo in sorted(df_m["Grupo"].dropna().unique()):
            sub = df_m[df_m["Grupo"] == grupo]
            hojas.append((str(grupo)[:31], sub[cols_mov]))

    return exportar_informe_excel(
        titulo="Analizador de inversiones FIFO",
        subtitulo=str(meta.get("nota") or "Identificar tipo → reglas por tipo → FIFO"),
        periodo=datetime.now().strftime("%d/%m/%Y %H:%M"),
        kpis=kpis,
        resumenes=[
            ("Por tipo de inversión", pd.DataFrame(por_tipo) if len(por_tipo) else pd.DataFrame()),
        ],
        detalle=df_v if not df_v.empty else df_m[cols_mov] if not df_m.empty else pd.DataFrame(),
        hoja_detalle="FIFO_movimientos" if not df_v.empty else "Movimientos",
        hojas_adicionales=hojas + ([("Movimientos", df_m[cols_mov])] if not df_v.empty and not df_m.empty else []),
        col_moneda=[
            "Importe", "Monto_Total", "Precio", "Costo_Aplicado", "Resultado",
            "Costo_Total", "Costo_Unitario", "Costo_unitario_lote", "Costo_parcial",
            "Precio_salida", "Costo_Unitario_Promedio",
        ],
        col_fecha=["Fecha", "Fecha_salida", "Fecha_lote"],
        total_col="Resultado" if not df_v.empty and "Resultado" in df_v.columns else None,
    )


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
