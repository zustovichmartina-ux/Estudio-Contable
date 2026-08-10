# -*- coding: utf-8 -*-
"""Motor de conciliación bancaria: clasificación por reglas, match proveedores/VEPs."""
from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Iterable

import pandas as pd

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


def clasificar(
    descripcion: str,
    reglas: list[dict] | None = None,
) -> dict[str, str]:
    """Primera regla cuyo patrón aparece en el texto normalizado."""
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
            return {
                "categoria": str(r.get("categoria") or ""),
                "tipo": str(r.get("tipo") or "DEBITO_REVISAR"),
            }
    return {
        "categoria": "Otros gastos (sin clasificar)",
        "tipo": "DEBITO_REVISAR",
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
    """Valida saldo línea a línea: saldo_prev + credito - debito ≈ saldo."""
    if not filas:
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


def df_extracto_a_filas(df: pd.DataFrame) -> list[dict]:
    """Normaliza DF unificado de procesador → filas del motor."""
    if df is None or df.empty:
        return []
    out: list[dict] = []
    for _, row in df.iterrows():
        desc = str(row.get("Descripcion") or row.get("Detalle") or row.get("Concepto unificado") or "")
        detalle = str(row.get("Detalle") or "")
        if detalle and detalle not in desc:
            desc = f"{desc} {detalle}".strip()
        debito = money(row.get("Debito"))
        credito = money(row.get("Credito"))
        if debito == 0 and credito == 0:
            imp = money(row.get("Importe"))
            tipo_m = str(row.get("Tipo Movimiento") or "").upper()
            if "CRED" in tipo_m or imp > 0 and "DEB" not in tipo_m:
                credito = abs(imp)
            else:
                debito = abs(imp)
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


def match_proveedor(
    mov: dict,
    pendientes: list[dict],
) -> dict | None:
    """Mejor factura pendiente: importe ±1, fecha ok, fuzzy ≥55."""
    importe = money(mov.get("debito") or mov.get("importe"))
    if importe <= 0:
        return None
    fecha_pago = mov.get("fecha")
    nombre = extraer_nombre_proveedor(str(mov.get("descripcion") or ""))
    mejor = None
    mejor_score = -1.0
    for f in pendientes:
        if f.get("usado"):
            continue
        imp_f = money(f.get("importe"))
        if abs(imp_f - importe) > TOL_IMPORTE:
            continue
        fecha_f = f.get("fecha")
        if isinstance(fecha_f, str):
            fecha_f = _parse_fecha(fecha_f)
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
        "similitud": mejor_score,
        "detalle": (
            f"{mejor.get('razon_social')} | {mejor.get('tipo_comp') or ''} "
            f"{mejor.get('num_comp') or ''} | similitud {mejor_score:.0f}%"
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

    for f in filas_extracto:
        desc = str(f.get("descripcion") or "")
        clf = clasificar(desc, reglas)
        categoria = clf["categoria"]
        tipo = clf["tipo"]
        credito = money(f.get("credito"))
        debito = money(f.get("debito"))
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
                usados.add(match_ref_id)
                for p in proveedores:
                    if p.get("id") == match_ref_id:
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
            match_detalle = "Sin regla de clasificación — revisar"
        elif tipo in ("INGRESO", "DEBITO_IMPUESTO", "DEBITO_FIJO", "INGRESO_O_DEBITO_PROPIO"):
            estado = "OK"

        resultados.append(
            {
                "cliente_id": cliente_id,
                "banco": banco or f.get("banco") or "",
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


def movimientos_a_filas_grilla_tango(movimientos: list[dict]) -> list[dict]:
    """Puente mínimo hacia grilla: categoría → hint de cuenta."""
    filas = []
    for m in movimientos:
        if m.get("estado") == "PENDIENTE":
            continue
        cat = str(m.get("categoria") or "")
        codigo = CATEGORIA_A_CUENTA_HINT.get(cat, "99999")
        credito = money(m.get("credito"))
        debito = money(m.get("debito"))
        filas.append(
            {
                "fecha": m.get("fecha"),
                "descripcion": m.get("descripcion"),
                "categoria": cat,
                "tipo": m.get("tipo"),
                "estado": m.get("estado"),
                "cuenta_sugerida": codigo,
                "debe": float(debito) if debito > 0 else 0.0,
                "haber": float(credito) if credito > 0 else 0.0,
                "match_detalle": m.get("match_detalle") or "",
            }
        )
    return filas
