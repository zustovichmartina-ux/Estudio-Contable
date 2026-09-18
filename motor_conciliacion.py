# -*- coding: utf-8 -*-
"""Motor de conciliación bancaria: clasificación por reglas, match proveedores/VEPs."""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from itertools import combinations
from typing import Any, Iterable

import pandas as pd

from conceptos_bancos import (
    _ALIAS_BANCO,
    _banco_desde_texto,
    cargar_instructivo,
    clasificar_movimiento_instructivo,
    tipo_desde_cuenta,
)
from capa_revision import (
    SCORE_AUTO_OK,
    confianza_clasificacion,
    resolver_codigo_plan,
)

# Seed del prompt (orden = prioridad; primera match gana)
REGLAS_SEED: list[tuple[str, str, str]] = [
    # INGRESOS
    ("CREDITO TRANSFERENCIA COELSA", "Deudores por Ventas - Cobranza", "INGRESO"),
    ("TRANSFERENCIA DE TERCEROS", "Deudores por Ventas - Cobranza", "INGRESO"),
    ("TRANSFERENCIAS CASH PROVEEDORES", "Deudores por Ventas - Cobranza", "INGRESO"),
    ("SERVICIO PAGO A PROVEEDORES", "Deudores por Ventas - Cobranza", "INGRESO"),
    ("G.DE ECHEQ", "Deudores por Ventas - Cobranza (echeq)", "INGRESO"),
    ("RESCATE FIMA", "Rescate FCI", "INGRESO"),
    # Ley 25413
    ("IMP. CRE. LEY 25413", "Imp. Déb/Cred. Ley 25413 (sobre créditos)", "DEBITO_IMPUESTO"),
    ("IMP. DEB. LEY 25413", "Imp. Déb/Cred. Ley 25413 (sobre débitos)", "DEBITO_IMPUESTO"),
    ("IMPUESTO DEB.LEY 25413", "Imp. Déb/Cred. Ley 25413 (sobre débitos)", "DEBITO_IMPUESTO"),
    # IIBB
    ("ING. BRUTOS S/ CRED", "Percepción IIBB (SIRCREB / Tucumán)", "DEBITO_IMPUESTO"),
    ("IMP. ING. BRUTOS", "Percepción IIBB Pcia Bs As", "DEBITO_IMPUESTO"),
    ("REG.RECAU.SIRCREB", "Percepción IIBB (SIRCREB)", "DEBITO_IMPUESTO"),
    # IVA
    ("PERCEP. IVA", "Percepción IVA", "DEBITO_IMPUESTO"),
    ("IVA", "IVA (débito fiscal / gasto bancario)", "DEBITO_IMPUESTO"),
    # AFIP / VEPs
    ("TRANSF. AFIP", "Pago AFIP - VEP (identificar por importe)", "DEBITO_VEP"),
    ("DEB. AUTOM. DE SERV. AFIP", "Plan de pago AFIP (débito automático)", "DEBITO_VEP"),
    ("PLANRG5321", "Plan de pago AFIP", "DEBITO_VEP"),
    # Sueldos / cargas
    ("SERVICIO ACREDITAMIENTO DE HABERES", "Pago de Haberes", "DEBITO_FIJO"),
    ("PAGO DE SERVICIOS AUTONOMOS", "Aportes SS (Autónomos)", "DEBITO_FIJO"),
    ("PAGO DE SERVICIOS MONOTR", "Monotributo / Aportes SS", "DEBITO_FIJO"),
    ("SANCOR SEGURO", "Seguros Pagados", "DEBITO_FIJO"),
    ("DEB. AUTOM. DE SERV. PREVENCION SALUD", "Deb aut. Obra social", "DEBITO_FIJO"),
    ("PROVINCIASEGUROS", "Seguros Pagados", "DEBITO_FIJO"),
    ("FAECYS", "Sindicatos sec (FAECYS)", "DEBITO_FIJO"),
    ("SINDICATO EMPLEADOS", "Sindicatos sec / Osecac", "DEBITO_FIJO"),
    ("INACAP", "Pago INACAP", "DEBITO_FIJO"),
    # Servicios
    ("CLARO", "Gastos de Telefonia e Internet", "DEBITO_FIJO"),
    ("SOLUCIONES ONLIN", "Otros gastos (Soluciones Online)", "DEBITO_FIJO"),
    ("IGSABEGA", "Gastos de Energia Electrica / Servicios", "DEBITO_FIJO"),
    # Tesorería
    ("SUSCRIPCION FIMA", "Suscripción Fondos de Inversión", "DEBITO_FIJO"),
    ("COMISION POR CUSTODIA DE TITULOS", "Comis. y Gtos Bcarios.", "DEBITO_FIJO"),
    ("COMP. TITULOS", "Compra de Titulos/Bonos", "DEBITO_FIJO"),
    ("COMISION Y DERECHOS DE MERCADO", "Comis. y Gtos Bcarios.", "DEBITO_FIJO"),
    # Comisiones
    ("COMISION SERVICIO DE CUENTA", "Comis. y Gtos Bcarios.", "DEBITO_FIJO"),
    ("COMISION EXTRACCION EN EFECTIVO", "Comis. y Gtos Bcarios.", "DEBITO_FIJO"),
    ("EXTRACCION EN AUTOSERVICIO", "Caja (extracción efectivo)", "DEBITO_FIJO"),
    # Cuentas propias (después de ingresos genéricos)
    ("TRANSF INMED CP", "TRANSFERENCIAS REALIZADAS e/ctas (propia)", "DEBITO_FIJO"),
    # Proveedores
    ("TRF INMED PROVEED", "Proveedores varios (a conciliar por importe)", "DEBITO_PROVEEDOR"),
]

# Mapeo mínimo categoría → hint de cuenta (fase 1b; 99999 = revisar)
CATEGORIA_A_CUENTA_HINT: dict[str, str] = {
    "Deudores por Ventas - Cobranza": "11301",
    "Deudores por Ventas - Cobranza (echeq)": "11301",
    "Rescate FCI": "12501",
    "Imp. Déb/Cred. Ley 25413 (sobre créditos)": "51201",
    "Imp. Déb/Cred. Ley 25413 (sobre débitos)": "51201",
    "Percepción IIBB (SIRCREB / Tucumán)": "11402",
    "Percepción IIBB Pcia Bs As": "11402",
    "Percepción IIBB (SIRCREB)": "11402",
    "Percepción IVA": "11401",
    "Pago de Haberes": "21102",
    "Comis. y Gtos Bcarios.": "52201",
    "Proveedores varios (a conciliar por importe)": "21101",
    "TRANSFERENCIAS REALIZADAS e/ctas (propia)": "11101",
    "Caja (extracción efectivo)": "11101",
}

_MONEY = Decimal("0.01")
TOL_IMPORTE = Decimal("1.00")
SIMILITUD_MIN = 55.0
RE_VEP = re.compile(r"VEP\s*(\d+)", re.I)


def money(v: Any) -> Decimal:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return Decimal("0.00")
    if isinstance(v, Decimal):
        return v.quantize(_MONEY, rounding=ROUND_HALF_UP)
    s = str(v).strip().replace(" ", "")
    if not s or s.lower() in {"nan", "none"}:
        return Decimal("0.00")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s).quantize(_MONEY, rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0.00")


def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper()
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _parse_fecha(val: Any) -> date | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.notna(ts):
            return ts.date()
    except Exception:
        pass
    return None


_TOKENS_INTER = (
    "FIMA",
    "FCI",
    "FONDOS DE INVERSION",
    "FONDOS COMUNES",
    "CTAS PROPIAS",
    "CUENTA PROPIA",
    "E/CTAS",
    "ENTRE CUENTAS",
    "MONEDA EXTRANJ",
)
_TOKENS_RETENCION = (
    "25413",
    "PERCEPCION",
    "RETENCION",
    "SIRCREB",
    "IIBB",
    "ING BRUTOS",
    "ING. BRUTOS",
    "ARBA",
    "DEBITOS Y CREDITOS",
    "DEB/CRED",
    "LEY 25",
)
_TOKENS_DEDUCCION = (
    "AUTONOM",
    "MONOTRIBUTO",
    "OBRA SOCIAL",
    "PREPAGA",
    "SEGURO",
)
_LABEL_EXTRACTO_A_VISTA = {
    "pago arba": "retencion",
    "sircreb": "retencion",
    "ingresos brutos tucuman": "retencion",
    "impuestos a los debitos y creditos": "retencion",
    "impuesto a los sellos": "retencion",
    "percepcion iva": "retencion",
    "iva": "retencion",
    "iibb": "retencion",
    "rescate fima": "inter-cta",
    "suscripcion fci": "inter-cta",
    "transferencias recibidas": "ingreso",
    "pagos recibidos": "ingreso",
    "acreditaciones comercios": "ingreso",
    "depositos en efvo": "ingreso",
    "cheques recibidos": "ingreso",
    "intereses": "ingreso",
    "transferencias emitidas": "egreso",
    "cheques emitidos": "egreso",
    "gastos bancarios": "egreso",
    "pago de haberes": "egreso",
    "pagos afip": "egreso",
    "pagos tarjeta corporativa": "egreso",
    "compras": "egreso",
    "pago de servicios": "egreso",
    "inversiones": "inter-cta",
}


def bucket_ae(mov: dict, *, extracto_label: str = "") -> str:
    """Vista AE-Studio (Ingresos/Egresos/…) sobre la clasificación que ya hace la web."""
    tipo = str(mov.get("tipo") or "")
    cat = normalizar_texto(str(mov.get("categoria") or ""))
    desc = normalizar_texto(str(mov.get("descripcion") or ""))
    label = normalizar_texto(extracto_label)
    blob = f"{cat} {desc} {label}"
    credito = money(mov.get("credito"))
    debito = money(mov.get("debito"))

    if tipo == "INGRESO_O_DEBITO_PROPIO" or any(t in blob for t in _TOKENS_INTER):
        return "inter-cta"
    if tipo == "DEBITO_IMPUESTO" or any(t in blob for t in _TOKENS_RETENCION):
        return "retencion"
    if "IVA" in blob and ("PERCEP" in blob or "TASA GENERAL" in blob or "ALQUILER" in blob):
        return "retencion"
    if any(t in blob for t in _TOKENS_DEDUCCION):
        return "deduccion"
    mapped = _LABEL_EXTRACTO_A_VISTA.get(label.lower())
    if mapped:
        if mapped == "ingreso" and debito > credito:
            return "egreso"
        return mapped
    if tipo == "INGRESO" or credito > debito:
        return "ingreso"
    if "IDENTIFICAR" in cat or tipo == "DEBITO_REVISAR":
        return "sin-cat"
    return "egreso"


def _resultado_instructivo(hit: dict, debito: float, credito: float) -> dict[str, Any]:
    cuenta = str(hit.get("cuenta") or "Movimientos a identificar")
    score = float(hit.get("score") or 0)
    identificado = bool(hit.get("identificado"))
    tipo = tipo_desde_cuenta(cuenta, debito, credito)
    if not identificado:
        tipo = "DEBITO_REVISAR"
    return {
        "categoria": cuenta,
        "tipo": tipo,
        "citar": str(hit.get("citar") or ""),
        "concepto_instructivo": str(hit.get("concepto") or ""),
        "fuente": "conceptos_bancos",
        "score": score,
        "confianza": confianza_clasificacion(
            fuente="conceptos_bancos", score=score, identificado=identificado,
        ),
    }


def _resultado_regla_local(regla: dict) -> dict[str, Any]:
    patron = normalizar_texto(str(regla.get("patron") or ""))
    return {
        "categoria": str(regla.get("categoria") or ""),
        "tipo": str(regla.get("tipo") or "DEBITO_REVISAR"),
        "citar": f"Regla local «{patron}» (Configuración).",
        "concepto_instructivo": "",
        "fuente": "regla_local",
        "score": 75.0,
        "confianza": "media",
    }


def _buscar_regla_local(descripcion: str, reglas: list[dict] | None) -> dict | None:
    texto = normalizar_texto(descripcion)
    lista = reglas
    if lista is None:
        lista = [
            {"patron": p, "categoria": c, "tipo": t, "orden": i, "activo": 1}
            for i, (p, c, t) in enumerate(REGLAS_SEED)
        ]
    for r in sorted(lista, key=lambda x: int(x.get("orden") or 0)):
        if not r.get("activo", 1):
            continue
        patron = normalizar_texto(str(r.get("patron") or ""))
        if patron and patron in texto:
            return r
    return None


def clasificar(
    descripcion: str,
    reglas: list[dict] | None = None,
    *,
    banco: str = "",
    debito: float = 0.0,
    credito: float = 0.0,
    instructivo: dict | None = None,
) -> dict[str, Any]:
    """Instructivo si identifica; si no, reglas de Configuración; si no, a revisar."""
    hit_inst: dict[str, Any] | None = None
    if instructivo and banco:
        hit_inst = clasificar_movimiento_instructivo(
            descripcion,
            banco,
            debito=debito,
            credito=credito,
            instructivo=instructivo,
        )
        if hit_inst.get("identificado"):
            return _resultado_instructivo(hit_inst, debito, credito)

    regla = _buscar_regla_local(descripcion, reglas)
    if regla is not None:
        return _resultado_regla_local(regla)

    if hit_inst is not None:
        return _resultado_instructivo(hit_inst, debito, credito)

    return {
        "categoria": "Movimientos a identificar",
        "tipo": "DEBITO_REVISAR",
        "citar": "Sin fila citable en Conceptos Bancos ni en reglas locales.",
        "concepto_instructivo": "",
        "fuente": "",
        "score": 0.0,
        "confianza": "baja",
    }


def extraer_nombre_proveedor(descripcion: str) -> str:
    """TRF INMED PROVEED / Nombre / CUIT / Banco → segundo campo."""
    partes = [p.strip() for p in str(descripcion or "").split("/")]
    if len(partes) >= 2:
        return partes[1].strip()
    return str(descripcion or "").strip()


def extraer_numero_vep(descripcion: str) -> str:
    m = RE_VEP.search(str(descripcion or ""))
    return m.group(1) if m else ""


def validar_saldos_corridos(filas: list[dict]) -> tuple[bool, str]:
    """Valida saldo línea a línea: saldo_prev + credito - debito ≈ saldo.

    Si el extracto no trae saldo (Excel/CSV sin columna, todo en 0) no se
    trata como inconsistente: no hay cadena que controlar.
    """
    if not filas:
        return True, ""
    if all(money(f.get("saldo")) == 0 for f in filas):
        return True, ""
    prev: Decimal | None = None
    for i, f in enumerate(filas):
        saldo = money(f.get("saldo"))
        credito = money(f.get("credito"))
        debito = money(f.get("debito"))
        if prev is None:
            prev = saldo
            continue
        esperado = (prev + credito - debito).quantize(_MONEY, rounding=ROUND_HALF_UP)
        if abs(esperado - saldo) > Decimal("0.05"):
            return False, (
                f"Saldo inconsistente en fila {i + 1}: "
                f"esperado {esperado}, informado {saldo}"
            )
        prev = saldo
    return True, ""


_TOKENS_NCC = ("NCC", "NOTA DE CREDITO", "NOTA CREDITO", "NOTA DE CRÉDITO")
_TOKENS_NDD = ("NDD", "NOTA DE DEBITO", "NOTA DEBITO", "NOTA DE DÉBITO")


def _lado_debito_credito(
    descripcion: str,
    debito: Decimal,
    credito: Decimal,
    importe: Decimal,
    tipo_mov: str,
) -> tuple[Decimal, Decimal]:
    """Débito/crédito del extracto. NCC acredita; el signo del Importe manda si no hay D/C."""
    if debito != 0 or credito != 0:
        return debito, credito
    blob = normalizar_texto(descripcion)
    tipo_u = normalizar_texto(tipo_mov)
    es_ncc = any(t in blob for t in _TOKENS_NCC) or "NCC" in blob
    es_ndd = any(t in blob for t in _TOKENS_NDD) or "NDD" in blob
    if es_ncc and not es_ndd:
        return Decimal("0.00"), abs(importe)
    if es_ndd and not es_ncc:
        return abs(importe), Decimal("0.00")
    if "CRED" in tipo_u and "DEB" not in tipo_u:
        return Decimal("0.00"), abs(importe)
    if "DEB" in tipo_u:
        return abs(importe), Decimal("0.00")
    if importe < 0:
        return abs(importe), Decimal("0.00")
    if importe > 0:
        return Decimal("0.00"), abs(importe)
    return Decimal("0.00"), Decimal("0.00")


_RE_SALDO_MOV = re.compile(
    r"SALDO\s+(ANTERIOR|INICIAL|FINAL|TOTAL|DE CUENTA|EN CUENTA)|SALDO\s+AL\s+\d"
)


def es_fila_saldo_bancario(desc: str, tipo_fila: str = "") -> bool:
    """Saldo anterior / inicial / final: no es movimiento a clasificar."""
    tipo = str(tipo_fila or "").strip().lower()
    if tipo in {"saldo inicial", "saldo final", "saldo anterior"}:
        return True
    n = re.sub(r"\bNAN\b", " ", normalizar_texto(desc))
    n = re.sub(r"\s+", " ", n).strip()
    if n in {"SALDO", "SALDO TOTAL", "SALDO DE CUENTA", "SALDO ANTERIOR"}:
        return True
    return bool(_RE_SALDO_MOV.search(n))


def _limpiar_desc_extracto(desc: str) -> str:
    t = str(desc or "").strip()
    t = re.sub(r"\s+\bnan\b", "", t, flags=re.I)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def df_extracto_a_filas(df: pd.DataFrame) -> list[dict]:
    """Normaliza DF unificado de procesador → filas del motor."""
    if df is None or df.empty:
        return []
    out: list[dict] = []
    for _, row in df.iterrows():
        desc = str(row.get("Descripcion") or row.get("Detalle") or row.get("Concepto unificado") or "")
        detalle = str(row.get("Detalle") or "").strip()
        if detalle.lower() in {"nan", "none", "null"}:
            detalle = ""
        if detalle and detalle not in desc:
            desc = f"{desc} {detalle}".strip()
        desc = _limpiar_desc_extracto(desc)
        if es_fila_saldo_bancario(desc, str(row.get("Tipo fila") or "")):
            continue
        debito, credito = _lado_debito_credito(
            desc,
            money(row.get("Debito")),
            money(row.get("Credito")),
            money(row.get("Importe")),
            str(row.get("Tipo Movimiento") or ""),
        )
        if debito == 0 and credito == 0:
            continue
        out.append(
            {
                "fecha": _parse_fecha(row.get("Fecha")),
                "descripcion": desc,
                "credito": credito,
                "debito": debito,
                "saldo": money(row.get("Saldo")),
                "archivo": str(row.get("Archivo origen") or ""),
                "banco": str(row.get("Banco") or ""),
            }
        )
    return out


def _fuzzy_ratio(a: str, b: str) -> float:
    try:
        from rapidfuzz import fuzz

        return float(fuzz.token_set_ratio(normalizar_texto(a), normalizar_texto(b)))
    except Exception:
        na, nb = normalizar_texto(a), normalizar_texto(b)
        if not na or not nb:
            return 0.0
        if na == nb:
            return 100.0
        if na in nb or nb in na:
            return 70.0
        return 0.0


def _fecha_factura(f: dict) -> date | None:
    fecha_f = f.get("fecha")
    if isinstance(fecha_f, str):
        return _parse_fecha(fecha_f)
    if isinstance(fecha_f, datetime):
        return fecha_f.date()
    if isinstance(fecha_f, date):
        return fecha_f
    return None


def match_proveedor(
    mov: dict,
    pendientes: list[dict],
) -> dict | None:
    """1:1 (importe ±1) y, si no hay, 1:N (un pago cubre varias facturas del mismo proveedor)."""
    hit = match_proveedor_1a1(mov, pendientes)
    if hit:
        return hit
    return match_proveedor_1n(mov, pendientes)


def match_proveedor_1a1(
    mov: dict,
    pendientes: list[dict],
) -> dict | None:
    """Mejor factura pendiente: importe ±1, fecha ok, fuzzy ≥55."""
    importe = money(mov.get("debito") or mov.get("importe"))
    if importe <= 0:
        return None
    fecha_pago = mov.get("fecha")
    if isinstance(fecha_pago, str):
        fecha_pago = _parse_fecha(fecha_pago)
    nombre = extraer_nombre_proveedor(str(mov.get("descripcion") or ""))
    mejor = None
    mejor_score = -1.0
    for f in pendientes:
        if f.get("usado"):
            continue
        imp_f = money(f.get("importe"))
        if abs(imp_f - importe) > TOL_IMPORTE:
            continue
        fecha_f = _fecha_factura(f)
        if fecha_pago and fecha_f and fecha_f > fecha_pago + timedelta(days=2):
            continue
        score = _fuzzy_ratio(nombre, str(f.get("razon_social") or ""))
        if score > mejor_score:
            mejor_score = score
            mejor = f
    if mejor is None or mejor_score < SIMILITUD_MIN:
        return None
    return {
        "factura": mejor,
        "facturas": [mejor],
        "similitud": mejor_score,
        "detalle": (
            f"{mejor.get('razon_social')} | {mejor.get('tipo_comp') or ''} "
            f"{mejor.get('num_comp') or ''} | similitud {mejor_score:.0f}%"
        ),
    }


def match_proveedor_1n(
    mov: dict,
    pendientes: list[dict],
    *,
    max_facturas: int = 4,
    max_candidatos: int = 12,
) -> dict | None:
    """Un débito cubre 2–4 facturas del mismo proveedor (suma ± tolerancia)."""
    importe = money(mov.get("debito") or mov.get("importe"))
    if importe <= 0:
        return None
    fecha_pago = mov.get("fecha")
    if isinstance(fecha_pago, str):
        fecha_pago = _parse_fecha(fecha_pago)
    nombre = extraer_nombre_proveedor(str(mov.get("descripcion") or ""))
    candidatos: list[tuple[float, dict]] = []
    for f in pendientes:
        if f.get("usado"):
            continue
        score = _fuzzy_ratio(nombre, str(f.get("razon_social") or f.get("proveedor") or ""))
        if score < SIMILITUD_MIN:
            continue
        fecha_f = _fecha_factura(f)
        if fecha_pago and fecha_f and fecha_f > fecha_pago + timedelta(days=2):
            continue
        if money(f.get("importe")) <= 0:
            continue
        candidatos.append((score, f))
    if len(candidatos) < 2:
        return None
    candidatos.sort(key=lambda x: -x[0])
    candidatos = candidatos[:max_candidatos]
    mejor = None
    for k in range(2, min(max_facturas, len(candidatos)) + 1):
        for combo in combinations(candidatos, k):
            facturas = [c[1] for c in combo]
            suma = sum((money(f.get("importe")) for f in facturas), Decimal("0.00"))
            dif = abs(suma - importe)
            if dif > TOL_IMPORTE:
                continue
            sim = min(c[0] for c in combo)
            cand = (dif, -sim, k, facturas, sim, suma)
            if mejor is None or cand[:3] < mejor[:3]:
                mejor = cand
    if mejor is None:
        return None
    _dif, _ns, _k, facturas, sim, suma = mejor
    nums = " + ".join(
        f"{money(f.get('importe'))} ({(f.get('tipo_comp') or '')} {(f.get('num_comp') or '')})".strip()
        for f in facturas
    )
    return {
        "factura": facturas[0],
        "facturas": facturas,
        "similitud": sim,
        "detalle": (
            f"1:N {facturas[0].get('razon_social') or ''} | {nums} = {suma} "
            f"≈ {importe} | similitud {sim:.0f}%"
        ),
    }


def match_vep(
    mov: dict,
    veps: list[dict],
) -> dict | None:
    importe = money(mov.get("debito") or mov.get("importe"))
    if importe <= 0:
        return None
    fecha_pago = mov.get("fecha")
    nro = extraer_numero_vep(str(mov.get("descripcion") or ""))
    candidatos = []
    for v in veps:
        if nro and str(v.get("numero_vep") or "") == nro:
            candidatos.append(v)
            continue
        if abs(money(v.get("importe")) - importe) > TOL_IMPORTE:
            continue
        fecha_v = v.get("fecha")
        if isinstance(fecha_v, str):
            fecha_v = _parse_fecha(fecha_v)
        if fecha_pago and fecha_v and abs((fecha_pago - fecha_v).days) > 3:
            continue
        candidatos.append(v)
    if not candidatos:
        return None
    # Preferir match por número VEP; si no, por importe exacto
    if nro:
        for v in candidatos:
            if str(v.get("numero_vep") or "") == nro:
                return v
    return min(candidatos, key=lambda v: abs(money(v.get("importe")) - importe))


def correr_motor(
    filas_extracto: list[dict],
    reglas: list[dict],
    proveedores: list[dict],
    veps: list[dict],
    *,
    cliente_id: int,
    banco: str,
    periodo: date | None,
    saldo_ok: bool = True,
) -> list[dict]:
    """
    Clasifica y matchea. Devuelve lista de bank_transactions listas para persistir.
    Si saldo_ok=False, no auto-concilia (todo queda revisable).
    """
    # Copia local de usados
    usados: set[Any] = {p.get("id") for p in proveedores if p.get("usado")}
    resultados: list[dict] = []
    instructivo = None
    try:
        instructivo = cargar_instructivo()
    except Exception:
        instructivo = None

    for f in filas_extracto:
        desc = str(f.get("descripcion") or "")
        credito = money(f.get("credito"))
        debito = money(f.get("debito"))
        banco_fila = str(f.get("banco") or "").strip() or str(banco or "")
        clf = clasificar(
            desc,
            reglas,
            banco=banco_fila,
            debito=float(debito),
            credito=float(credito),
            instructivo=instructivo,
        )
        categoria = clf["categoria"]
        tipo = clf["tipo"]
        citar = str(clf.get("citar") or "")
        confianza = str(clf.get("confianza") or "baja")
        fuente = str(clf.get("fuente") or "")
        score = float(clf.get("score") or 0)
        estado = "OK"
        match_detalle = None
        match_ref_id = None

        if not saldo_ok:
            estado = "PENDIENTE"
            match_detalle = "Extracto con saldos inconsistentes — revisar parseo/OCR"
        elif tipo == "DEBITO_PROVEEDOR" and debito > 0:
            pend = [p for p in proveedores if p.get("id") not in usados]
            m = match_proveedor(
                {"fecha": f.get("fecha"), "descripcion": desc, "debito": debito},
                pend,
            )
            if m:
                estado = "CONCILIADO"
                match_detalle = m["detalle"]
                match_ref_id = m["factura"].get("id")
                for fac in m.get("facturas") or [m["factura"]]:
                    fid = fac.get("id")
                    if fid is None:
                        continue
                    usados.add(fid)
                    for p in proveedores:
                        if p.get("id") == fid:
                            p["usado"] = True
            else:
                estado = "PENDIENTE"
                match_detalle = (
                    "No matchea con ninguna factura del ERP — revisar o imputar a gasto directo"
                )
        elif tipo == "DEBITO_VEP" and debito > 0:
            nro = extraer_numero_vep(desc)
            if not veps:
                estado = "PENDIENTE"
                match_detalle = f"Sin padrón VEPs cargado" + (f" — VEP {nro}" if nro else "")
            else:
                v = match_vep(
                    {"fecha": f.get("fecha"), "descripcion": desc, "debito": debito},
                    veps,
                )
                if v:
                    estado = "CONCILIADO"
                    impuesto = str(v.get("impuesto") or "AFIP")
                    categoria = f"Pago AFIP - {impuesto}"
                    match_detalle = (
                        f"VEP {v.get('numero_vep') or nro} | {impuesto} | "
                        f"período {v.get('periodo_fiscal') or ''}"
                    )
                    match_ref_id = v.get("id")
                else:
                    estado = "PENDIENTE"
                    match_detalle = f"VEP sin match en padrón" + (f" — VEP {nro}" if nro else "")
        elif tipo == "DEBITO_REVISAR":
            estado = "PENDIENTE"
            match_detalle = citar or "Sin regla de clasificación — revisar"
        elif tipo in ("INGRESO", "DEBITO_IMPUESTO", "DEBITO_FIJO", "INGRESO_O_DEBITO_PROPIO"):
            estado = "OK"
            if citar:
                match_detalle = citar

        if estado == "OK" and confianza in {"media", "baja"}:
            estado = "PENDIENTE"
            extra = f" — confirmar (score {score:.0f}"
            if fuente == "regla_local":
                extra = " — confirmar regla local"
            elif score:
                extra = f" — confirmar (score {score:.0f} < {SCORE_AUTO_OK:.0f})"
            else:
                extra = " — confirmar"
            match_detalle = f"{citar or match_detalle or 'Clasificación automática'}{extra}"

        resultados.append(
            {
                "cliente_id": cliente_id,
                "banco": banco_fila or f.get("banco") or "",
                "periodo": periodo,
                "fecha": f.get("fecha"),
                "descripcion": desc,
                "credito": str(credito),
                "debito": str(debito),
                "saldo": str(money(f.get("saldo"))),
                "categoria": categoria,
                "tipo": tipo,
                "estado": estado,
                "match_detalle": match_detalle,
                "match_ref_id": str(match_ref_id) if match_ref_id is not None else None,
                "fuente": fuente,
                "score": score,
                "confianza": confianza,
            }
        )
    return resultados


def resumen_por_categoria(movimientos: list[dict]) -> pd.DataFrame:
    if not movimientos:
        return pd.DataFrame(
            columns=["categoria", "cantidad", "creditos", "debitos", "neto"]
        )
    rows = []
    for m in movimientos:
        rows.append(
            {
                "categoria": m.get("categoria") or "(sin categoría)",
                "credito": money(m.get("credito")),
                "debito": money(m.get("debito")),
            }
        )
    df = pd.DataFrame(rows)
    g = (
        df.groupby("categoria", dropna=False)
        .agg(cantidad=("credito", "count"), creditos=("credito", "sum"), debitos=("debito", "sum"))
        .reset_index()
    )
    g["neto"] = g["creditos"] - g["debitos"]
    for col in ("creditos", "debitos", "neto"):
        g[col] = g[col].map(lambda x: float(money(x)))
    total = pd.DataFrame(
        [
            {
                "categoria": "TOTAL",
                "cantidad": int(g["cantidad"].sum()),
                "creditos": float(money(g["creditos"].sum())),
                "debitos": float(money(g["debitos"].sum())),
                "neto": float(money(g["creditos"].sum() - g["debitos"].sum())),
            }
        ]
    )
    return pd.concat([g, total], ignore_index=True)


def origen_linea_extracto(mov: dict) -> str:
    """regla = instructivo/seed; sugerido = sentido común; a_clasificar = sin cuenta."""
    fuente = str(mov.get("fuente") or "")
    categoria = str(mov.get("categoria") or "").lower()
    codigo = str(mov.get("cuenta_codigo") or mov.get("cuenta_sugerida") or "").strip()
    if "identificar" in categoria:
        if codigo and codigo != "99999":
            return "sugerido"
        return "a_clasificar"
    if fuente in {"conceptos_bancos", "regla_local"}:
        return "regla"
    if codigo and codigo not in {"", "99999"}:
        return "sugerido"
    return "a_clasificar"


def cuenta_banco_del_plan(
    plan_df: pd.DataFrame | None,
    nombre_banco: str,
) -> tuple[str, str]:
    """Si hay una sola cuenta del banco en el plan, la usa; si hay varias, no adivina."""
    if plan_df is None or getattr(plan_df, "empty", True) or not nombre_banco:
        return "99999", str(nombre_banco or "Banco")

    clave = _banco_desde_texto(nombre_banco)
    aliases = list(_ALIAS_BANCO.get(clave, (clave,)))
    col_cod = "codigo" if "codigo" in plan_df.columns else plan_df.columns[0]
    col_desc = "descripcion" if "descripcion" in plan_df.columns else plan_df.columns[1]
    hits: list[tuple[str, str]] = []
    for _, row in plan_df.iterrows():
        desc = str(row.get(col_desc) or "").strip()
        dn = normalizar_texto(desc).lower()
        if not any(a in dn for a in aliases):
            continue
        if not any(t in dn for t in ("banco", "bco", "cta", "cuenta", "cc ")):
            continue
        cod = str(row.get(col_cod) or "").strip()
        if cod:
            hits.append((cod, desc))
    if len(hits) == 1:
        return hits[0]
    cte = [
        h for h in hits
        if "cte" in normalizar_texto(h[1]).lower()
        or "corriente" in normalizar_texto(h[1]).lower()
    ]
    if len(cte) == 1:
        return cte[0]
    return "99999", str(nombre_banco or "Banco")


def renglones_asiento_banco_mes(
    movimientos: list[dict],
    *,
    codigo_banco: str,
    descripcion_banco: str,
    periodo: str = "",
    fecha_str: str = "",
) -> list[dict]:
    """Agrupa el extracto por cuenta de contrapartida y cierra contra el banco."""
    por_cuenta: dict[str, dict] = {}
    banco_debe = 0.0
    banco_haber = 0.0
    for m in movimientos:
        debito = float(money(m.get("debito")))
        credito = float(money(m.get("credito")))
        cod = str(m.get("cuenta_codigo") or m.get("cuenta_sugerida") or "99999").strip() or "99999"
        desc = str(m.get("cuenta_plan") or m.get("categoria") or "").strip() or cod
        slot = por_cuenta.setdefault(cod, {"debe": 0.0, "haber": 0.0, "desc": desc})
        slot["desc"] = desc or slot["desc"]
        if debito > 0.005:
            slot["debe"] = round(slot["debe"] + debito, 2)
            banco_haber = round(banco_haber + debito, 2)
        if credito > 0.005:
            slot["haber"] = round(slot["haber"] + credito, 2)
            banco_debe = round(banco_debe + credito, 2)

    rows: list[dict] = []

    def _fila(codigo: str, descripcion: str, debe: float, haber: float) -> dict:
        return {
            "Período": periodo,
            "Fecha": fecha_str,
            "Código": codigo,
            "Descripción": descripcion,
            "Debe": round(float(debe or 0), 2),
            "Haber": round(float(haber or 0), 2),
            "Estado": "Ingresado",
        }

    for cod, slot in sorted(por_cuenta.items(), key=lambda kv: kv[0]):
        neto = round(slot["debe"] - slot["haber"], 2)
        if abs(neto) < 0.005:
            continue
        if neto > 0:
            rows.append(_fila(cod, slot["desc"], neto, 0.0))
        else:
            rows.append(_fila(cod, slot["desc"], 0.0, abs(neto)))

    neto_banco = round(banco_debe - banco_haber, 2)
    if abs(neto_banco) >= 0.005:
        if neto_banco > 0:
            rows.append(_fila(codigo_banco, descripcion_banco, neto_banco, 0.0))
        else:
            rows.append(_fila(codigo_banco, descripcion_banco, 0.0, abs(neto_banco)))
    return rows


def movimientos_a_filas_grilla_tango(
    movimientos: list[dict],
    plan_cuentas: pd.DataFrame | None = None,
    *,
    incluir_pendientes: bool = True,
) -> list[dict]:
    """Puente a grilla: categoría del instructivo → código del plan del cliente.

    Pendientes sin cuenta clara van como 99999 (sin match, no se inventa).
    No se omiten: el asiento tiene que mostrar lo que falta confirmar.
    """
    filas = []
    for m in movimientos:
        estado = str(m.get("estado") or "")
        if estado == "PENDIENTE" and not incluir_pendientes:
            continue
        cat = str(m.get("categoria") or "")
        codigo, desc_plan, score_plan = resolver_codigo_plan(
            cat, plan_cuentas, hints=CATEGORIA_A_CUENTA_HINT,
        )
        if estado == "PENDIENTE" and (
            not cat or "identificar" in cat.lower() or codigo == "99999"
        ):
            codigo = "99999"
        credito = money(m.get("credito"))
        debito = money(m.get("debito"))
        filas.append(
            {
                "fecha": m.get("fecha"),
                "descripcion": m.get("descripcion"),
                "categoria": cat,
                "cuenta_plan": desc_plan,
                "tipo": m.get("tipo"),
                "estado": estado,
                "cuenta_sugerida": codigo,
                "debe": float(debito) if debito > 0 else 0.0,
                "haber": float(credito) if credito > 0 else 0.0,
                "match_detalle": m.get("match_detalle") or "",
                "origen": m.get("fuente") or "",
                "score_plan": score_plan,
            }
        )
    return filas


_MESES_ES = (
    "", "ene", "feb", "mar", "abr", "may", "jun",
    "jul", "ago", "sep", "oct", "nov", "dic",
)


def armar_papel_mes(
    movs: list[dict],
    *,
    saldo_inicial_arrastre: Decimal | float | None = None,
) -> dict[str, Any]:
    """Papel de un mes: cierre = apertura + créditos − débitos. No fuerza diferencia a 0."""
    if not movs:
        return {
            "anio": None,
            "mes": None,
            "periodo": "",
            "apertura": 0.0,
            "creditos": 0.0,
            "debitos": 0.0,
            "cierre": 0.0,
            "segun_resumen": None,
            "diferencia": None,
            "origen_apertura": "sin_datos",
            "movimientos": 0,
        }

    def _ord(m: dict):
        return (_parse_fecha(m.get("fecha")) or date.min, str(m.get("id") or ""))

    ordenados = sorted(movs, key=_ord)
    creditos = sum((money(m.get("credito")) for m in ordenados), Decimal("0.00"))
    debitos = sum((money(m.get("debito")) for m in ordenados), Decimal("0.00"))
    first = ordenados[0]
    last = ordenados[-1]
    if saldo_inicial_arrastre is not None:
        apertura = money(saldo_inicial_arrastre)
        origen = "arrastre"
    else:
        s0 = money(first.get("saldo"))
        apertura = (s0 - money(first.get("credito")) + money(first.get("debito"))).quantize(
            _MONEY, rounding=ROUND_HALF_UP
        )
        origen = "extracto"
    cierre = (apertura + creditos - debitos).quantize(_MONEY, rounding=ROUND_HALF_UP)
    hay_saldo = any(money(m.get("saldo")) != 0 for m in ordenados)
    if hay_saldo:
        segun = money(last.get("saldo"))
        diferencia = (segun - cierre).quantize(_MONEY, rounding=ROUND_HALF_UP)
    else:
        segun = None
        diferencia = None
    fecha0 = _parse_fecha(first.get("fecha")) or _parse_fecha(first.get("periodo"))
    periodo = ""
    anio = mes = None
    if fecha0:
        anio, mes = fecha0.year, fecha0.month
        periodo = f"{_MESES_ES[mes]}-{str(anio)[2:]}"
    return {
        "anio": anio,
        "mes": mes,
        "periodo": periodo,
        "apertura": float(apertura),
        "creditos": float(creditos),
        "debitos": float(debitos),
        "cierre": float(cierre),
        "segun_resumen": float(segun) if segun is not None else None,
        "diferencia": float(diferencia) if diferencia is not None else None,
        "origen_apertura": origen,
        "movimientos": len(ordenados),
    }


def _meses_consecutivos(a: tuple[int, int], b: tuple[int, int]) -> bool:
    y1, m1 = a
    y2, m2 = b
    if m1 == 12:
        return (y2, m2) == (y1 + 1, 1)
    return (y2, m2) == (y1, m1 + 1)


def papeles_por_mes(movimientos: list[dict]) -> list[dict]:
    """Cadena mensual. Primer mes: apertura del extracto. Siguientes: arrastre del cierre.

    No inventa meses sin movimientos. Si hay un hueco, el mes que retoma abre
    con el extracto. La diferencia (según resumen − cierre) se informa y no
    se corrige a 0.
    """
    grupos: dict[tuple[int, int], list[dict]] = {}
    for m in movimientos:
        d = _parse_fecha(m.get("fecha")) or _parse_fecha(m.get("periodo"))
        if d is None:
            continue
        grupos.setdefault((d.year, d.month), []).append(m)
    out: list[dict] = []
    arrastre: Decimal | None = None
    prev_clave: tuple[int, int] | None = None
    for clave in sorted(grupos):
        if prev_clave is not None and not _meses_consecutivos(prev_clave, clave):
            arrastre = None
        papel = armar_papel_mes(
            grupos[clave],
            saldo_inicial_arrastre=arrastre,
        )
        out.append(papel)
        arrastre = money(papel["cierre"])
        prev_clave = clave
    return out
