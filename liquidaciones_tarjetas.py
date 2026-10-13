"""Liquidaciones de tarjetas: lectura de los PDF de Payway, Fiserv (First Data) y Naranja,
control contra lo acreditado y armado de una factura de gastos por procesadora."""
from __future__ import annotations

import hashlib
import io
import re
import zipfile
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

import pandas as pd
import streamlit as st

NAVY = "#1F3557"
GOLD = "#B08D48"

PROCESADORAS = ("Payway", "Fiserv", "Naranja", "Favacard", "Nave")
# Nombre con que el estudio carga cada procesadora (editable en pantalla). Payway se separa además por banco.
PV_DEFECTO = {"Payway": "FRANCES - 00001", "Naranja": "NARANJA - 00002",
              "Favacard": "FAVACARD - 00003", "Fiserv": "FIRST DATA - 00004", "Nave": "NAVE"}
ORDEN_PROC = ("Payway", "Favacard", "Naranja", "Fiserv", "Nave")

# columnas numéricas de cada liquidación (orden de la planilla del estudio)
COLS = ("Total presentado", "BI 21%", "BI 10,5%", "IVA 21%", "IVA 10,5%", "P IIBB", "R IIBB", "P IVA", "R IVA",
        "R Gcias", "No gravado", "Acreditado PDF")
CAMPOS = ("presentado", "bi", "bi105", "iva", "iva105", "p_iibb", "r_iibb", "p_iva", "r_iva", "r_gcias",
          "no_grav", "neto")
CAMPO_A_COL = dict(zip(CAMPOS, COLS))


def _txt(x) -> str:
    return "" if x is None or (isinstance(x, float) and x != x) else str(x).strip()


def grupo_de(proc: str, banco: str = "") -> str:
    return f"{proc} · {banco}" if _txt(banco) else proc


def nombre_defecto(proc: str, banco: str = "") -> str:
    if proc == "Payway" and _txt(banco):
        return f"PAYWAY {_txt(banco)}"
    return PV_DEFECTO.get(proc, proc)

_RE_IMPORTE = re.compile(r"^-?\s*\$?\s*-?\s*\d{1,3}(?:\.\d{3})*,\d{2}\s*-?$")
Q2 = Decimal("0.01")


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


# ---------------------------------------------------------------- números
def a_dec(txt) -> Decimal | None:
    """'1.234,56' / '5.794,50-' / '-   32.714,69' / '$ -53.999,00' -> Decimal."""
    t = str(txt or "").strip()
    if not t:
        return None
    neg = "-" in t
    t = re.sub(r"[^\d,.]", "", t)
    if not t:
        return None
    t = t.replace(".", "").replace(",", ".")
    try:
        d = Decimal(t)
    except InvalidOperation:
        return None
    return -d if neg else d


def es_importe(tok: str) -> bool:
    return bool(_RE_IMPORTE.match(tok.strip()))


def fmt_ar(d: Decimal | float | None) -> str:
    if d is None:
        return ""
    d = Decimal(str(d)).quantize(Q2, ROUND_HALF_UP)
    s = f"{abs(d):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if d < 0 else ''}$ {s}"


def _vacio() -> dict:
    return {c: Decimal(0) for c in CAMPOS}


def _tokens(doc) -> list[str]:
    out = []
    for p in doc:
        for l in p.get_text("text").split("\n"):
            l = l.strip()
            if l:
                out.append(l)
    return out


def _abrir(data: bytes):
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        return fitz.open(stream=data, filetype="pdf")
    except Exception:
        return None


# ---------------------------------------------------------------- Payway
def _marca_por_nombre(nombre: str) -> str:
    n = nombre.lower()
    for clave, marca in (("amex", "AMEX"), ("american", "AMEX"), ("cabal", "CABAL"), ("master", "MASTER"),
                         ("visa", "VISA"), ("naranja", "NARANJA"), ("maestro", "MAESTRO")):
        if clave in n:
            return marca
    return ""


def _banco_pagador(tok: list[str]) -> str:
    """Banco que paga el resumen Payway ('007 BANCO GALICIA' -> 'GALICIA')."""
    for t in tok[:80]:
        m = re.match(r"^\d{3}\s+(BANCO\s+.+)$", t.strip(), re.I)
        if m:
            n = re.sub(r"^BANCO\s+", "", m.group(1), flags=re.I)
            n = re.sub(r"^(DE LA|DEL|DE)\s+", "", n, flags=re.I)
            n = re.sub(r"\s+S\.?A\.?U?\.?$", "", n, flags=re.I)
            n = re.sub(r"\s+ARGENTINA$", "", n, flags=re.I)
            return n.upper().strip() or m.group(1).upper()
    return ""


def parsear_payway(tok: list[str], nombre: str) -> dict:
    texto = "\n".join(tok)
    r = _vacio()
    notas: list[str] = []
    banco = _banco_pagador(tok)
    # totales del encabezado: presentado, descuento, saldo, y tres ceros de U$S
    m = re.search(r"([\d.]+,\d\d)\n([\d.]+,\d\d)\n([\d.]+,\d\d)\n0,00\n0,00\n0,00", texto)
    descuento = None
    if m:
        r["presentado"] = a_dec(m.group(1))
        descuento = a_dec(m.group(2))
        r["neto"] = a_dec(m.group(3))
    else:
        notas.append("No pude leer los totales del encabezado.")
    try:
        i0 = max(k for k, t in enumerate(tok) if t.startswith("DESGLOSE DE DESCUENTOS"))
    except ValueError:
        return {"datos": r, "marca": _marca_por_nombre(nombre), "banco": banco,
                "notas": notas + ["No encontré el desglose de descuentos."], "fecha": ""}
    fin = len(tok)
    for k in range(i0, len(tok)):
        if tok[k].startswith("(*) Servicio") or tok[k].startswith("SR. COMERCIANTE"):
            fin = k
            break
    bloque = tok[i0 + 1:fin]
    previos: list[str] = []
    cargos = Decimal(0)
    bases: dict[str, Decimal] = {}
    en_base = False
    for t in bloque:
        if t.startswith("Base Imponible"):
            en_base = True
            previos = []
            continue
        if t.startswith("Deducciones"):
            en_base = False
            previos = [t]
            continue
        if re.match(r"^\$\s", t) and es_importe(t):
            monto = a_dec(t)
            etiqueta = " ".join(x for x in previos if re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", x)).lower()
            previos = []
            if en_base:
                mt = re.search(r"tasa\s+(\d+)[,.](\d+)", etiqueta)
                if mt:
                    k = f"{int(mt.group(1))},{mt.group(2)}"
                    bases[k] = bases.get(k, Decimal(0)) + monto
                continue
            if "arancel" in etiqueta or "servicio" in etiqueta or "cargo por serv" in etiqueta \
                    or "bonif" in etiqueta:
                cargos += monto
            elif re.search(r"\biva\b", etiqueta) and "ret" not in etiqueta and "perc" not in etiqueta:
                if re.search(r"10[,.]5", etiqueta):
                    r["iva105"] += monto
                else:
                    r["iva"] += monto
            elif "sellos" in etiqueta:
                r["no_grav"] += monto
            elif "perc" in etiqueta and ("ib" in etiqueta or "bruto" in etiqueta) and "afip" not in etiqueta:
                r["p_iibb"] += monto
            elif ("ret" in etiqueta and "ib" in etiqueta) or "sirtac" in etiqueta:
                r["r_iibb"] += monto
            elif "afip" in etiqueta or "dgi" in etiqueta:
                r["p_iva"] += monto
                notas.append("'Percep./Retenc. AFIP-DGI' se cargó como Percepción de IVA (como en la planilla del estudio).")
            elif "ganancia" in etiqueta:
                r["r_gcias"] += monto
            elif "ret" in etiqueta and "iva" in etiqueta:
                r["r_iva"] += monto
            else:
                r["no_grav"] += monto
                notas.append(f"Concepto sin clasificar ({etiqueta.strip() or 's/nombre'}): {fmt_ar(monto)} quedó en No gravado.")
        else:
            previos.append(t)
            if len(previos) > 6:
                previos = previos[-6:]
    if not bases:
        r["bi"] = cargos
    else:
        r["bi"] = bases.get("21,00", Decimal(0))
        r["bi105"] = bases.get("10,50", Decimal(0))
        for k, v in bases.items():
            if k not in ("21,00", "10,50"):
                notas.append(f"Base imponible con alícuota {k}% no contemplada ({fmt_ar(v)}): revisar.")
        if (sum(bases.values(), Decimal(0)) - cargos).copy_abs() > Decimal("0.05"):
            notas.append("Los cargos (arancel/servicio) no coinciden con la base imponible informada.")
    if (r["iva"] - r["bi"] * Decimal("0.21")).copy_abs() > Decimal("1.5"):
        notas.append("El IVA 21% no coincide con el 21% de la base imponible.")
    if (r["iva105"] - r["bi105"] * Decimal("0.105")).copy_abs() > Decimal("1.5"):
        notas.append("El IVA 10,5% no coincide con el 10,5% de la base imponible.")
    if descuento is not None:
        suma = (r["bi"] + r["bi105"] + r["iva"] + r["iva105"] + r["p_iibb"] + r["r_iibb"] + r["p_iva"] + r["r_iva"]
                + r["r_gcias"] + r["no_grav"] - (r["bi"] + r["bi105"]) + cargos)
        if (suma - descuento).copy_abs() > Decimal("0.05"):
            notas.append(f"Los descuentos del desglose ({fmt_ar(suma)}) no coinciden con el total del encabezado ({fmt_ar(descuento)}).")
    fecha = ""
    mf = re.search(r"(\d{2}/\d{2}/\d{4})\n000\d+", texto)
    if mf:
        fecha = mf.group(1)
    return {"datos": r, "marca": _marca_por_nombre(nombre), "banco": banco, "notas": sorted(set(notas)), "fecha": fecha}


# ---------------------------------------------------------------- Fiserv (First Data)
def parsear_fiserv(tok: list[str], nombre: str) -> dict:
    r = _vacio()
    notas: list[str] = []
    suma_netos = Decimal(0)
    n_netos = 0
    for i, t in enumerate(tok):
        if t.startswith("IMPORTE NETO DE PAGOS"):
            for j in (i + 1, i + 2):
                if j < len(tok) and es_importe(tok[j]):
                    suma_netos += a_dec(tok[j])
                    n_netos += 1
                    break
            continue
        if t in ("+", "-") and i >= 1 and i + 2 < len(tok) and tok[i + 1] == "$" and es_importe(tok[i + 2]):
            etiqueta = tok[i - 1].strip().upper()
            monto = a_dec(tok[i + 2]).copy_abs()
            if etiqueta.startswith("VENTAS") or "ACREDITACIONES" in etiqueta:
                r["presentado"] += monto if t == "+" else -monto
            elif etiqueta.startswith("ARANCEL"):
                r["bi"] += monto
            elif "IVA CRED" in etiqueta or etiqueta.startswith("IVA "):
                r["iva"] += monto
            elif "PERC IVA" in etiqueta or "PERCEP" in etiqueta and "IVA" in etiqueta:
                r["p_iva"] += monto
            elif "PERC" in etiqueta and ("IIBB" in etiqueta or "ING" in etiqueta or "IB" in etiqueta):
                r["p_iibb"] += monto
            elif "RET" in etiqueta and ("IB" in etiqueta or "BRUTOS" in etiqueta or "SIRTAC" in etiqueta):
                r["r_iibb"] += monto
            elif "RET" in etiqueta and "IVA" in etiqueta:
                r["r_iva"] += monto
            elif "RET" in etiqueta and "GAN" in etiqueta:
                r["r_gcias"] += monto
            else:
                r["no_grav"] += monto if t == "-" else -monto
                notas.append(f"Concepto sin clasificar '{etiqueta}': {fmt_ar(monto)} quedó en No gravado.")
    texto = "\n".join(tok)
    m = re.search(r"IMPORTE :\n(-?\s*[\d.]+,\d\d-?)", texto)
    total_pdf = a_dec(m.group(1)) if m else None
    r["neto"] = total_pdf if total_pdf is not None else suma_netos
    if total_pdf is not None and (suma_netos - total_pdf).copy_abs() > Decimal("0.05"):
        notas.append(f"La suma de 'Importe neto de pagos' ({fmt_ar(suma_netos)}) no coincide con el total del PDF ({fmt_ar(total_pdf)}).")
    marca = _marca_por_nombre(nombre)
    if "credito" in nombre.lower().replace("é", "e").replace("_", "e"):
        marca = (marca + " Crédito").strip()
    elif "debit" in nombre.lower():
        marca = (marca + " Débito").strip()
    return {"datos": r, "marca": marca, "notas": sorted(set(notas)), "fecha": ""}


# ---------------------------------------------------------------- Naranja
def parsear_naranja(tok: list[str], nombre: str) -> dict:
    r = _vacio()
    notas: list[str] = []
    texto = "\n".join(tok)
    # total vendido: el importe que sigue al último "Totales"
    idx = [k for k, t in enumerate(tok) if t == "Totales"]
    if idx:
        for j in range(idx[-1] + 1, min(idx[-1] + 4, len(tok))):
            if es_importe(tok[j].replace("$", "").strip()) or es_importe(tok[j]):
                r["presentado"] = a_dec(tok[j])
                break
    try:
        a = next(k for k, t in enumerate(tok) if t.startswith("Detalles de facturaci"))
        b = next(k for k in range(a, len(tok)) if tok[k].startswith("Total de Facturaci"))
    except StopIteration:
        return {"datos": r, "marca": "NARANJA", "notas": ["No encontré el detalle de facturación."], "fecha": ""}
    bloque = tok[a + 1:b]
    previos: list[str] = []
    for t in bloque:
        if re.match(r"^\$\s*-?[\d.]+,\d\d$", t):
            monto = a_dec(t)
            etiqueta = " ".join(x for x in previos if re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", x)).lower()
            previos = []
            if "arancel" in etiqueta or "inter" in etiqueta:
                r["bi"] += monto
            elif "iva" in etiqueta and "perc" not in etiqueta:
                r["iva"] += monto
            elif "perc" in etiqueta and "iva" in etiqueta:
                r["p_iva"] += monto
            elif "perc" in etiqueta and ("ib" in etiqueta or "bruto" in etiqueta):
                r["p_iibb"] += monto
            else:
                r["no_grav"] += monto
                notas.append(f"Concepto sin clasificar ({etiqueta.strip() or 's/nombre'}): {fmt_ar(monto)} quedó en No gravado.")
        else:
            previos.append(t)
    # retenciones
    try:
        c = next(k for k, t in enumerate(tok) if t.startswith("Detalle Retenciones"))
        d = next(k for k in range(c, len(tok)) if tok[k].startswith("Total de Retenciones"))
        previos = []
        for t in tok[c + 1:d]:
            if re.match(r"^\$\s*-?[\d.]+,\d\d$", t):
                monto = a_dec(t)
                etiqueta = " ".join(previos).lower()
                previos = []
                if "iva" in etiqueta:
                    r["r_iva"] += monto
                elif "gan" in etiqueta:
                    r["r_gcias"] += monto
                else:
                    r["r_iibb"] += monto
            else:
                previos.append(t)
    except StopIteration:
        pass
    # neto liquidado: "Importe | $ x | Neto Liquidado"
    for k, t in enumerate(tok):
        if t.startswith("Neto Liquidado") and k >= 1 and re.match(r"^\$\s*-?[\d.]+,\d\d$", tok[k - 1]):
            r["neto"] = a_dec(tok[k - 1])
            break
    fecha = ""
    mf = re.search(r"Fecha de Pago\s+(\d{2}/\d{2}/\d{4})", texto)
    if mf:
        fecha = mf.group(1)
    return {"datos": r, "marca": "NARANJA", "notas": sorted(set(notas)), "fecha": fecha}


# ---------------------------------------------------------------- detección
def parsear_favacard(tok: list[str], nombre: str) -> dict:
    """Sólo se lee 'PROXIMOS PAGOS DE ESTA PRESENTACION' (el resto de la liquidación no se toma)."""
    r = _vacio()
    notas: list[str] = []
    try:
        a = next(k for k, t in enumerate(tok) if t.startswith("PROXIMOS PAGOS"))
    except StopIteration:
        return {"datos": r, "marca": "FAVACARD", "notas": ["No encontré 'Próximos pagos de esta presentación'."], "fecha": ""}
    k = a + 2  # salta el encabezado de columnas
    fechas: list[str] = []
    while k + 5 < len(tok) and re.match(r"^\d{2}/\d{2}/\d{2}$", tok[k]):
        v = [a_dec(x) for x in tok[k + 1:k + 6]]
        if None in v:
            break
        pres, cargos, iva_perc, p_ib, neto = v
        iva = min(iva_perc, (cargos * Decimal("0.21")).quantize(Q2, ROUND_HALF_UP))
        r["presentado"] += pres
        r["bi"] += cargos
        r["iva"] += iva
        r["p_iva"] += iva_perc - iva
        r["p_iibb"] += p_ib
        r["neto"] += neto
        if (pres - cargos - iva_perc - p_ib - neto).copy_abs() > Decimal("0.05"):
            notas.append(f"La fila del {tok[k]} no cierra (presentado - cargos - IVA/perc - perc IIBB <> neto).")
        d, m, y = tok[k].split("/")
        fechas.append(f"{d}/{m}/20{y}")
        k += 6
    if not fechas:
        notas.append("No pude leer ninguna fila de próximos pagos.")
    notas.append("La retención IIBB SIRTAC no se carga (se toma de ARCA); el neto es el de la presentación, antes de retenciones.")
    return {"datos": r, "marca": "FAVACARD", "notas": sorted(set(notas)), "fecha": fechas[0] if fechas else ""}


def _n(x) -> str:
    return _sin_acentos(_txt(x).lower())


def _num(x) -> Decimal:
    if x is None or (isinstance(x, float) and x != x):
        return Decimal(0)
    if isinstance(x, (int, float)):
        return Decimal(str(x))
    t = str(x).strip().replace("$", "").replace(" ", "")
    if not t or t == "-":
        return Decimal(0)
    if "," in t:
        v = a_dec(t)
        return v if v is not None else Decimal(0)
    try:
        return Decimal(t)
    except InvalidOperation:
        return Decimal(0)


def _iva_credito(r: dict, notas: list[str], monto: Decimal, costo: Decimal) -> None:
    """IVA del costo: crédito fiscal al 21% o al 10,5% según corresponda."""
    if costo and (monto / costo - Decimal("0.105")).copy_abs() < Decimal("0.003"):
        r["iva105"] += monto
        r["bi105"] += costo
        return
    r["iva"] += monto
    r["bi"] += costo
    if costo and (monto / costo - Decimal("0.21")).copy_abs() >= Decimal("0.003"):
        notas.append(f"El IVA ({fmt_ar(monto)}) no es el 21% ni el 10,5% del costo ({fmt_ar(costo)}): puede incluir "
                     "retención de IVA, revisar.")


def _tipo_impuesto(h: str) -> str | None:
    iva = bool(re.search(r"\biva\b|i\.v\.a", h))
    if "reten" in h:
        return "r_iva" if iva else ("r_gcias" if "ganancia" in h else "r_iibb")
    if "percep" in h:
        return "p_iva" if iva else "p_iibb"
    return "iva" if iva else None


def parsear_nave_excel(data: bytes, nombre: str) -> dict | None:
    """Excel 'Detalle de operaciones' de Nave. Devuelve None si no tiene ese formato."""
    try:
        raw = pd.read_excel(io.BytesIO(data), header=None, dtype=object)
    except Exception:
        return None
    filas = raw.values.tolist()
    ih = next((i for i, f in enumerate(filas) if any(_n(c) == "monto bruto" for c in f)), None)
    if ih is None:
        return None
    hdr = [_n(c) for c in filas[ih]]
    idx = lambda pred: next((i for i, h in enumerate(hdr) if h and pred(h)), None)
    cb, cn = idx(lambda h: h == "monto bruto"), idx(lambda h: h == "monto neto")
    cc = idx(lambda h: "iva" not in h and h.startswith(("costo", "comision", "cargo", "arancel", "tarifa")))
    ce = idx(lambda h: h == "estado")
    cf = idx(lambda h: h.startswith("fecha de operacion")) or 0
    if cb is None or cn is None:
        return None
    imp = [(i, _tipo_impuesto(h)) for i, h in enumerate(hdr) if h and i not in (cb, cn, cc) and _tipo_impuesto(h)]
    r = _vacio()
    notas: list[str] = []
    n_aj, aj, n_raro, n_pend, pend = 0, Decimal(0), 0, 0, Decimal(0)
    n_ops = 0
    for f in filas[ih + 1:]:
        c0 = f[cf]
        if not re.match(r"^\d{2}/\d{2}/\d{4}", _txt(c0)) and not hasattr(c0, "year"):
            continue
        n_ops += 1
        bruto, neto = _num(f[cb]), _num(f[cn])
        costo = _num(f[cc]) if cc is not None else Decimal(0)
        t = {}
        for i, tipo in imp:
            t[tipo] = t.get(tipo, Decimal(0)) + _num(f[i])
        ivaf = t.pop("iva", Decimal(0))
        if cc is None and ivaf:
            costo = (ivaf / Decimal("0.21")).quantize(Q2, ROUND_HALF_UP)
            notas.append("No encontré la columna de comisiones: la base imponible se calculó desde el IVA (IVA / 21%); verificar.")
        resid = bruto - costo - ivaf - sum(t.values(), Decimal(0)) - neto
        if resid.copy_abs() > Decimal("0.011"):
            t["r_iibb"] = t.get("r_iibb", Decimal(0)) + resid
            n_aj += 1
            aj += resid
        if costo or ivaf:
            if costo and (ivaf / costo - Decimal("0.105")).copy_abs() < Decimal("0.003"):
                r["bi105"] += costo
                r["iva105"] += ivaf
            else:
                r["bi"] += costo
                r["iva"] += ivaf
                if costo and (ivaf / costo - Decimal("0.21")).copy_abs() >= Decimal("0.003"):
                    n_raro += 1
        for k, v in t.items():
            r[k] += v
        r["presentado"] += bruto
        r["neto"] += neto
        if ce is not None and "acreditar" in _n(f[ce]):
            n_pend += 1
            pend += neto
    if not n_ops:
        return None
    if n_aj:
        notas.append(f"{n_aj} operación(es) con diferencia entre bruto, costos y neto ({fmt_ar(aj)}), por ejemplo "
                     "devoluciones que reintegran retenciones: se cargó como ajuste de R IIBB.")
    if n_raro:
        notas.append(f"En {n_raro} operación(es) el IVA no es el 21% ni el 10,5% del costo: puede incluir retención "
                     "de IVA, revisar.")
    if n_pend:
        notas.append(f"Incluye {n_pend} operación(es) por acreditar (neto {fmt_ar(pend)}).")
    fecha, cuit = "", ""
    for f in filas[:ih]:
        for c in f:
            mm = re.search(r"Detalles del .* al (\d{2}/\d{2}/\d{4})", _txt(c))
            if mm:
                fecha = mm.group(1)
            mc = re.match(r"^CUIT:\s*(\d+)", _txt(c))
            if mc:
                cuit = mc.group(1)
    tot = next((f for f in filas[ih + 1:] if _n(f[0]) == "total"), None)
    if tot is not None and abs(_num(tot[cn]) - r["neto"]) > Decimal("0.05"):
        notas.append("El neto de las operaciones no coincide con la fila Total del Excel.")
    return {"datos": r, "marca": "NAVE", "banco": "", "notas": sorted(set(notas)), "fecha": fecha,
            "clave_nave": f"{cuit}|{fecha}"}


def _monto_sig(tok: list[str], i: int) -> Decimal | None:
    for j in range(i + 1, min(i + 3, len(tok))):
        if es_importe(tok[j]):
            return a_dec(tok[j])
    return None


def parsear_nave_pdf(tok: list[str], nombre: str) -> dict:
    r = _vacio()
    notas: list[str] = []
    texto = "\n".join(tok)
    mf = re.search(r"Detalles del .* al (\d{2}/\d{2}/\d{4})", texto)
    fecha = mf.group(1) if mf else ""
    mc = re.search(r"CUIT:\s*(\d+)", texto)
    clave = f"{mc.group(1) if mc else ''}|{fecha}"
    try:
        i_fin = max(k for k, t in enumerate(tok) if t == "Total bruto")
    except ValueError:
        return {"datos": r, "marca": "NAVE", "banco": "", "notas": ["No encontré el resumen de totales."],
                "fecha": fecha, "clave_nave": clave}

    def val(etq: str) -> Decimal:
        for i in range(i_fin, len(tok)):
            if tok[i].startswith(etq):
                m = _monto_sig(tok, i)
                if m is not None:
                    return m
        return Decimal(0)

    r["presentado"] = val("Total bruto")
    r["neto"] = val("Total neto")
    costo = val("Costo del servicio")
    try:
        a = next(k for k in range(i_fin, len(tok)) if tok[k].startswith("Impuestos sobre costos"))
        b = next(k for k in range(a, len(tok)) if tok[k].startswith("Total debitado por impuestos"))
    except StopIteration:
        a = b = None
    if costo and a is None:
        r["bi"] = costo
    en_costos, previos, hubo_iva = True, [], False
    for t in (tok[a + 1:b] if a is not None else []):
        if t.startswith("Impuestos sobre operaciones"):
            en_costos, previos = False, []
            continue
        if re.match(r"^\$\s*-?[\d.]+,\d\d$", t):
            monto = a_dec(t)
            e = _n(" ".join(previos))
            previos = []
            if re.search(r"\biva\b|i\.v\.a", e) and "reten" not in e and "percep" not in e:
                if en_costos:
                    _iva_credito(r, notas, monto, costo)
                    hubo_iva = True
                else:
                    r["r_iva"] += monto
                    notas.append("IVA bajo 'impuestos sobre operaciones': se cargó como retención de IVA, revisar.")
            elif re.search(r"\biva\b|i\.v\.a", e) and "reten" in e:
                r["r_iva"] += monto
            elif re.search(r"\biva\b|i\.v\.a", e) and "percep" in e:
                r["p_iva"] += monto
            elif "ganancia" in e:
                r["r_gcias"] += monto
            elif "percep" in e:
                r["p_iibb"] += monto
            elif "reten" in e or "sirtac" in e or "iibb" in e:
                r["r_iibb"] += monto
            else:
                r["no_grav"] += monto
                notas.append(f"Concepto sin clasificar ({e or 's/nombre'}): {fmt_ar(monto)} quedó en No gravado.")
        else:
            previos.append(t)
    if costo and not hubo_iva:
        r["bi"] += costo
    notas.append("El PDF de Nave es un resumen: si el mes tiene devoluciones puede quedar una diferencia; "
                 "el Excel de detalle es más preciso.")
    return {"datos": r, "marca": "NAVE", "banco": "", "notas": sorted(set(notas)), "fecha": fecha,
            "clave_nave": clave}


def clasificar(tok: list[str]) -> str:
    t = "\n".join(tok[:400])
    if "FAVACARD" in t.upper() and any(x.startswith("PROXIMOS PAGOS") for x in tok):
        return "Favacard"
    if any(x.startswith("Costos por cobro") for x in tok) and any(
            x.startswith("Total debitado por impuestos") for x in tok):
        return "Nave"
    if "Resumen Mensual de Liquidaciones" in t or "Servicio PAYWAY" in "\n".join(tok):
        return "Payway"
    if "FIRST DATA" in t.upper() and "Entidad Pagadora" in t:
        return "Fiserv"
    if "Tarjeta Naranja" in t and "Detalle de Cupones Liquidados" in "\n".join(tok):
        return "Naranja"
    if "Detalle de ajustes" in t:
        return "Ajustes"
    return ""


def _completar_bases(base: dict) -> None:
    """Si hay IVA sin base imponible, calcula la base desde el IVA y lo avisa."""
    d = base["datos"]
    for iv, bi, tasa, lbl in (("iva", "bi", Decimal("0.21"), "21%"), ("iva105", "bi105", Decimal("0.105"), "10,5%")):
        if d[iv] and not d[bi]:
            d[bi] = (d[iv] / tasa).quantize(Q2, ROUND_HALF_UP)
            base["notas"] = list(base["notas"]) + [
                f"IVA {lbl} sin base imponible en el archivo: se calculó la base ({fmt_ar(d[bi])}) a partir del "
                "IVA; verificar."]


def leer_archivo(nombre: str, data: bytes) -> dict:
    base = {"archivo": nombre.split("/")[-1], "proc": "", "marca": "", "banco": "", "datos": _vacio(),
            "notas": [], "fecha": ""}
    if nombre.lower().endswith((".xlsx", ".xls", ".xlsm")):
        res = parsear_nave_excel(data, nombre)
        if res is None:
            base["notas"] = ["No reconocí el Excel (se espera el 'Detalle de operaciones' de Nave)."]
            return base
        base["proc"] = "Nave"
        base.update(res)
        _completar_bases(base)
        return base
    doc = _abrir(data)
    if doc is None:
        base["notas"] = ["No se pudo abrir el PDF."]
        return base
    try:
        tok = _tokens(doc)
    finally:
        doc.close()
    proc = clasificar(tok)
    base["proc"] = proc
    if proc == "Payway":
        res = parsear_payway(tok, nombre)
    elif proc == "Fiserv":
        res = parsear_fiserv(tok, nombre)
    elif proc == "Naranja":
        res = parsear_naranja(tok, nombre)
    elif proc == "Favacard":
        res = parsear_favacard(tok, nombre)
    elif proc == "Nave":
        res = parsear_nave_pdf(tok, nombre)
    elif proc == "Ajustes":
        base["notas"] = ["Detalle de ajustes/promociones: no se carga."]
        return base
    else:
        base["notas"] = ["No reconocí de qué procesadora es este PDF."]
        return base
    base.update(res)
    _completar_bases(base)
    return base


def expandir(archivos: list[tuple[str, bytes]]) -> list[tuple[str, bytes]]:
    out = []
    for nombre, data in archivos:
        if nombre.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        if not info.is_dir() and info.filename.lower().endswith((".pdf", ".xlsx", ".xls")) \
                                and info.file_size < 40_000_000:
                            out.append((f"{nombre}/{info.filename}", z.read(info)))
            except zipfile.BadZipFile:
                pass
        else:
            out.append((nombre, data))
    return out


def leer_archivos(archivos: list[tuple[str, bytes]]) -> list[dict]:
    docs = [leer_archivo(n, d) for n, d in expandir(archivos)]
    # Nave: si llegan el PDF y el Excel del mismo período, se usa el Excel (más detallado)
    por_clave: dict[str, list[dict]] = {}
    for d in docs:
        if d["proc"] == "Nave" and d.get("clave_nave"):
            por_clave.setdefault(d["clave_nave"], []).append(d)
    for grupo in por_clave.values():
        if len(grupo) > 1:
            xl = [d for d in grupo if d["archivo"].lower().endswith((".xlsx", ".xls", ".xlsm"))]
            keep = xl[0] if xl else grupo[0]
            for d in grupo:
                if d is not keep:
                    d["proc"] = ""
                    d["notas"] = [f"Mismo período que {keep['archivo']}: se usó ese archivo."]
    orden = {p: i for i, p in enumerate(ORDEN_PROC)}
    docs.sort(key=lambda d: (orden.get(d["proc"], 9), d.get("banco", ""), d["fecha"][6:] + d["fecha"][3:5] + d["fecha"][:2],
                             d["archivo"]))
    return docs


# ---------------------------------------------------------------- tabla y cálculos
def a_tabla(docs: list[dict]) -> pd.DataFrame:
    filas = []
    for d in docs:
        if not d["proc"] or d["proc"] == "Ajustes":
            continue
        fila = {"Procesadora": d["proc"], "Banco": d.get("banco", ""), "Detalle": d["archivo"], "Tarjeta": d["marca"]}
        for c in CAMPOS:
            fila[CAMPO_A_COL[c]] = float(d["datos"][c])
        filas.append(fila)
    cols = ["Procesadora", "Banco", "Detalle", "Tarjeta"] + list(COLS)
    return pd.DataFrame(filas, columns=cols)


def _dec(x) -> Decimal:
    try:
        if x is None or (isinstance(x, float) and x != x):
            return Decimal(0)
        return Decimal(str(x))
    except InvalidOperation:
        return Decimal(0)


def diferencia(fila) -> Decimal:
    """Presentado - descuentos - retenciones - acreditado. Debe dar 0."""
    g = lambda c: _dec(fila.get(c))
    return (g("Total presentado") - g("BI 21%") - g("BI 10,5%") - g("IVA 21%") - g("IVA 10,5%") - g("P IIBB")
            - g("R IIBB") - g("P IVA") - g("R IVA") - g("R Gcias") - g("No gravado")
            - g("Acreditado PDF")).quantize(Q2, ROUND_HALF_UP)


def grupos(df: pd.DataFrame) -> list[tuple[str, str, pd.DataFrame, str]]:
    """(procesadora, banco, filas, etiqueta) en el orden de la planilla: una factura por cada uno."""
    d = df.copy()
    if "Banco" not in d.columns:
        d["Banco"] = ""
    d["_b"] = d["Banco"].map(_txt)
    orden = {p: i for i, p in enumerate(ORDEN_PROC)}
    claves = sorted({(p, b) for p, b in zip(d["Procesadora"], d["_b"]) if isinstance(p, str) and p},
                    key=lambda k: (orden.get(k[0], 9), k[1]))
    return [(p, b, d[(d["Procesadora"] == p) & (d["_b"] == b)].drop(columns=["_b"]), grupo_de(p, b)) for p, b in claves]


def resumen_facturas(df: pd.DataFrame, pv: dict[str, str]) -> pd.DataFrame:
    filas = []
    for proc, banco, sub, etq in grupos(df):
        s = lambda c: sum((_dec(v) for v in sub[c]), Decimal(0)).quantize(Q2, ROUND_HALF_UP)
        total = s("BI 21%") + s("BI 10,5%") + s("IVA 21%") + s("IVA 10,5%") + s("P IVA") + s("P IIBB")
        filas.append({
            "Procesadora": etq, "Comprobante": pv.get(etq, nombre_defecto(proc, banco)),
            "Gasto BI 21%": float(s("BI 21%")), "Gasto BI 10,5%": float(s("BI 10,5%")),
            "No gravado": float(s("No gravado")), "IVA 21%": float(s("IVA 21%")), "IVA 10,5%": float(s("IVA 10,5%")),
            "P IVA": float(s("P IVA")), "P IIBB": float(s("P IIBB")), "Total factura": float(total),
            "R IIBB": float(s("R IIBB")), "R IVA": float(s("R IVA")), "R Gcias": float(s("R Gcias")),
        })
    return pd.DataFrame(filas)


# ---------------------------------------------------------------- asiento
# Nombres de cuenta por defecto: editables en pantalla. Deben coincidir con la hoja CUENTAS del estudio.
CUENTAS_DEFECTO = {"gasto": "Gastos por comisiones tarjetas", "iva": "IVA Crédito Fiscal",
                   "gasto105": "Gastos por comisiones tarjetas 10,5%", "iva105": "IVA Crédito Fiscal 10,5%", "p_iva": "Percepciones IVA", "p_iibb": "Percepciones IIBB", "deudores": "Deudores por ventas"}
SIN_CUENTA = "(elegir cuenta)"

# patrones para sugerir cuenta del plan; sólo se sugiere si hay UNA sola coincidencia
_PATRONES = {
    "gasto": (r"comisi\w*.*tarjeta|tarjeta.*comisi", r"gastos?.*tarjeta"),
    "gasto105": (r"comisi\w*.*10[,.]?5",),
    "iva": (r"iva.*cr[eé]dito.*21|iva.*21.*cr[eé]dito", r"iva.*cr[eé]dito"),
    "iva105": (r"iva.*cr[eé]dito.*10[,.]?5|10[,.]?5.*iva.*cr[eé]dito|iva.*10[,.]?5",),
    "p_iva": (r"percepci\w*.*iva|iva.*percepci",),
    "p_iibb": (r"percepci\w*.*(iibb|ingresos brutos|ii\.?bb)", r"(iibb|ingresos brutos).*percepci"),
    "deudores": (r"deudores por ventas",),
}


def _sin_acentos(t: str) -> str:
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        t = t.replace(a, b)
    return t


def sugerir_cuenta(plan: list[str], clave: str) -> str:
    for pat in _PATRONES.get(clave, ()):
        rx = re.compile(_sin_acentos(pat))
        hallados = [c for c in plan if rx.search(_sin_acentos(c.lower()))]
        if len(hallados) == 1:
            return hallados[0]
    return ""


@st.cache_data(ttl=300, show_spinner=False)
def plan_del_cliente(nombre: str) -> list[str]:
    """Cuentas imputables del plan del cliente ('codigo - descripcion'), o [] si no se encuentra."""
    try:
        import database as db
        from pathlib import Path
        from procesador import cargar_plan_cuentas, plan_cuentas_desde_csv
        cli = next((c for c in db.listar_clientes() if str(c.get("nombre") or "").strip() == nombre), None)
        if not cli:
            return []
        df = None
        csv_txt = db.plan_cuentas_csv_cliente(int(cli["id"]))
        if csv_txt:
            df = plan_cuentas_desde_csv(csv_txt)
        if df is None or df.empty:
            ruta = str(cli.get("plan_cuentas_path") or "").strip()
            if ruta and Path(ruta).is_file():
                df = cargar_plan_cuentas(ruta)
        if df is None or df.empty:
            return []
        if "imputable" in df.columns:
            imp = df[df["imputable"].astype(str).str.lower().isin(["true", "1", "si", "sí", "s"])]
            if not imp.empty:
                df = imp
        out, vistos = [], set()
        for _, r in df.iterrows():
            cod = str(r.get("codigo", "")).strip()
            desc = str(r.get("descripcion", "")).strip()
            if not cod or cod.lower() in ("nan", "none", "codigo", "código") or cod in vistos:
                continue
            vistos.add(cod)
            out.append(f"{cod} - {desc}" if desc and desc.lower() not in ("nan", "none") else cod)
        return out
    except Exception:
        return []


def cuenta_haber(cu: dict) -> str:
    return cu.get("deudores") or SIN_CUENTA


def armar_asiento(res: pd.DataFrame, cuentas: dict | None = None) -> pd.DataFrame:
    """Factura de gastos: Debe gasto, IVA y percepciones / Haber deudores por ventas (tarjetas)."""
    cu = {**CUENTAS_DEFECTO, **(cuentas or {})}
    filas = []
    for _, r in res.iterrows():
        proc = r["Procesadora"]
        deb = [(cu["gasto"], _dec(r["Gasto BI 21%"])), (cu["gasto105"], _dec(r["Gasto BI 10,5%"])),
               (cu["iva"], _dec(r["IVA 21%"])), (cu["iva105"], _dec(r["IVA 10,5%"])),
               (cu["p_iva"], _dec(r["P IVA"])), (cu["p_iibb"], _dec(r["P IIBB"]))]
        for cta, imp in deb:
            if imp:
                filas.append({"Procesadora": proc, "Cuenta": cta, "Debe": float(imp), "Haber": 0.0})
        filas.append({"Procesadora": proc, "Cuenta": cuenta_haber(cu), "Debe": 0.0,
                      "Haber": float(_dec(r["Total factura"]))})
    return pd.DataFrame(filas, columns=["Procesadora", "Cuenta", "Debe", "Haber"])


# ---------------------------------------------------------------- Excel
def generar_excel(df: pd.DataFrame, cliente: str, periodo: str, pv: dict[str, str],
                  cuentas: dict | None = None) -> bytes:
    import xlsxwriter
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})
    f = lambda **k: wb.add_format({"font_name": "Arial", "font_size": 10, **k})
    CONT = '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'
    tit = f(bold=True, font_size=12, font_color=NAVY)
    sub = f(bold=True, font_color=NAVY)
    head = f(bold=True, bg_color=NAVY, font_color="#FFFFFF", align="center", valign="vcenter", text_wrap=True, border=1)
    txt = f()
    num = f(num_format=CONT)
    num_in = f(num_format=CONT, font_color="#0000FF")
    tot = f(num_format=CONT, bold=True, top=1)
    tot_t = f(bold=True, top=1)
    mes = f(num_format="mmm-yy", bold=True, align="left")
    ok = f(num_format=CONT, bold=True, bg_color="#E3F3E6")
    fecha_mes = _fecha_periodo(periodo)
    gl = grupos(df)

    def titulo(w):
        w.write("B1", cliente or "", tit)
        if fecha_mes:
            w.write_datetime("B2", fecha_mes, mes)
        else:
            w.write("B2", periodo or "", sub)

    ws = wb.add_worksheet("LIQUIDACIONES")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 46)
    ws.set_column("C:C", 18)
    ws.set_column("D:Q", 15)
    titulo(ws)
    fila = 3  # 0-index -> fila 4
    ref_tot: dict[str, int] = {}   # etiqueta -> fila de totales (1-index)
    usa105: dict[str, bool] = {}
    cab = ["Detalle", "Tarjeta", "TOTAL", "BI 21", "BI 10,5", "IVA 21%", "IVA 10,5%", "P IIBB", "R IIBB", "P IVA",
           "R IVA", "R GCIAS", "NO GRAVADO", "NETO", "ACREDITADO PDF", "DIFERENCIA"]
    for proc, banco, sub_df, etq in gl:
        ws.write(fila, 1, pv.get(etq, nombre_defecto(proc, banco)), sub)
        fila += 1
        for j, c in enumerate(cab):
            ws.write(fila, 1 + j, c, head)
        ws.set_row(fila, 30)
        fila += 1
        ini = fila + 1
        for _, r in sub_df.iterrows():
            fx = fila + 1
            ws.write(fila, 1, str(r["Detalle"]), txt)
            ws.write(fila, 2, str(r["Tarjeta"] or ""), txt)
            for k, c in enumerate(COLS[:-1]):
                ws.write_number(fila, 3 + k, float(_dec(r[c])), num_in)
            # NETO = TOTAL - descuentos, percepciones y retenciones
            ws.write_formula(fila, 14, f"=D{fx}-SUM(E{fx}:N{fx})", num)
            ws.write_number(fila, 15, float(_dec(r["Acreditado PDF"])), num_in)
            ws.write_formula(fila, 16, f"=O{fx}-P{fx}", num)
            fila += 1
        fin = fila
        ws.write(fila, 1, "TOTAL", tot_t)
        ws.write(fila, 2, "", tot_t)
        for k in range(3, 17):
            col = chr(ord("A") + k)
            ws.write_formula(fila, k, f"=SUM({col}{ini}:{col}{fin})", tot)
        ref_tot[etq] = fila + 1
        usa105[etq] = any(_dec(v) for c in ("BI 10,5%", "IVA 10,5%") for v in sub_df[c])
        fila += 3
    # hoja de facturas
    wf = wb.add_worksheet("FACTURAS")
    wf.set_column("A:A", 3)
    wf.set_column("B:B", 28)
    wf.set_column("C:M", 16)
    titulo(wf)
    cab2 = ["Comprobante", "Gasto (BI 21%)", "Gasto (BI 10,5%)", "No gravado (excluido)", "IVA 21%", "IVA 10,5%", "P IVA",
            "P IIBB", "TOTAL FACTURA", "R IIBB", "R IVA", "R GCIAS"]
    for j, c in enumerate(cab2):
        wf.write(3, 1 + j, c, head)
    wf.set_row(3, 30)
    fr = 4
    fila_fact: dict[str, int] = {}
    # (columna de FACTURAS, columna de LIQUIDACIONES)
    enlaces = ((2, "E"), (3, "F"), (4, "N"), (5, "G"), (6, "H"), (7, "K"), (8, "I"), (10, "J"), (11, "L"), (12, "M"))
    for proc, banco, sub_df, etq in gl:
        t = ref_tot[etq]
        wf.write(fr, 1, pv.get(etq, nombre_defecto(proc, banco)), txt)
        x = fr + 1
        fila_fact[etq] = x
        for cf, cl in enlaces:
            wf.write_formula(fr, cf, f"=LIQUIDACIONES!{cl}{t}", num)
        wf.write_formula(fr, 9, f"=C{x}+D{x}+F{x}+G{x}+H{x}+I{x}", ok)
        fr += 1
    if fr > 4:
        wf.write(fr, 1, "TOTAL", tot_t)
        for k in range(2, 13):
            col = chr(ord("A") + k)
            wf.write_formula(fr, k, f"=SUM({col}5:{col}{fr})", tot)
    # hoja de asientos (factura de gastos contra deudores por ventas)
    cu = {**CUENTAS_DEFECTO, **(cuentas or {})}
    wa = wb.add_worksheet("ASIENTO")
    wa.set_column("A:A", 3)
    wa.set_column("B:B", 52)
    wa.set_column("C:D", 18)
    titulo(wa)
    ar = 3
    for proc, banco, sub_df, etq in gl:
        x = fila_fact[etq]
        wa.write(ar, 1, pv.get(etq, nombre_defecto(proc, banco)) + " · factura de gastos", sub)
        ar += 1
        for j, c in enumerate(["Cuenta", "Debe", "Haber"]):
            wa.write(ar, 1 + j, c, head)
        ar += 1
        ini = ar + 1
        lineas = [(cu["gasto"], f"=FACTURAS!C{x}")]
        if usa105[etq]:
            lineas.append((cu["gasto105"], f"=FACTURAS!D{x}"))
        lineas.append((cu["iva"], f"=FACTURAS!F{x}"))
        if usa105[etq]:
            lineas.append((cu["iva105"], f"=FACTURAS!G{x}"))
        lineas += [(cu["p_iva"], f"=FACTURAS!H{x}"), (cu["p_iibb"], f"=FACTURAS!I{x}")]
        for cta, fd in lineas:
            wa.write(ar, 1, cta, txt)
            wa.write_formula(ar, 2, fd, num)
            wa.write_blank(ar, 3, None, num)
            ar += 1
        wa.write(ar, 1, cuenta_haber(cu), txt)
        wa.write_blank(ar, 2, None, num)
        wa.write_formula(ar, 3, f"=FACTURAS!J{x}", num)
        ar += 1
        wa.write(ar, 1, "TOTAL", tot_t)
        wa.write_formula(ar, 2, f"=SUM(C{ini}:C{ar})", tot)
        wa.write_formula(ar, 3, f"=SUM(D{ini}:D{ar})", tot)
        ar += 1
        wa.write(ar, 1, "Diferencia Debe - Haber", txt)
        wa.write_formula(ar, 2, f"=C{ar}-D{ar}", ok)
        ar += 1
        wa.write(ar, 1, "No gravado (no incluido en la factura)", txt)
        wa.write_formula(ar, 2, f"=FACTURAS!E{x}", num)
        ar += 3
    wb.close()
    return bio.getvalue()


_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
          "noviembre", "diciembre"]


def _fecha_periodo(periodo: str):
    import calendar
    import datetime as dt
    m = re.search(r"([A-Za-zñÑáéíóú]+)\s+(\d{4})", periodo or "")
    if not m or m.group(1).lower() not in _MESES:
        return None
    mes, anio = _MESES.index(m.group(1).lower()) + 1, int(m.group(2))
    return dt.datetime(anio, mes, calendar.monthrange(anio, mes)[1])


def periodo_desde_docs(docs: list[dict]) -> str:
    from collections import Counter
    c = Counter()
    for d in docs:
        m = re.match(r"(\d{2})/(\d{2})/(\d{4})", d.get("fecha") or "")
        if m and d["proc"] == "Payway":
            c[(int(m.group(3)), int(m.group(2)))] += 1
    if not c:
        for d in docs:
            m = re.match(r"(\d{2})/(\d{2})/(\d{4})", d.get("fecha") or "")
            if m:
                c[(int(m.group(3)), int(m.group(2)))] += 1
    if not c:
        return ""
    (a, m), _ = c.most_common(1)[0]
    return f"{_MESES[m - 1].capitalize()} {a}"


# ---------------------------------------------------------------- pantalla
@st.cache_data(show_spinner="Leyendo los PDF…", max_entries=8)
def _leer_cacheado(clave: str, archivos: tuple) -> list[dict]:
    return leer_archivos(list(archivos))


def render_liquidaciones_tarjetas() -> None:
    s = st.session_state
    st.markdown(f"<h2 style='color:{NAVY};font-weight:700'>Liquidaciones de tarjetas</h2>", unsafe_allow_html=True)
    st.caption("Subí los PDF o Excel del mes (Payway, Fiserv, Naranja, Favacard, Nave). Se leen solos, se controlan "
               "contra lo acreditado y se arma una factura de gastos por procesadora (Payway, una por banco).")
    try:
        from control_comprobantes import _nombres_clientes
        nombres = _nombres_clientes()
    except Exception:
        nombres = []
    c1, c2 = st.columns(2)
    if nombres:
        if "lt_cliente" not in s:
            activo = str(s.get("nombre_activo") or "").strip()
            s["lt_cliente"] = activo if activo in nombres else None
        cliente = c1.selectbox("Cliente (razón social)", nombres, key="lt_cliente",
                               placeholder="Elegí o buscá el cliente")
    else:
        cliente = c1.text_input("Cliente (razón social)", key="lt_cliente_txt")
    subidos = st.file_uploader("Liquidaciones (PDF, Excel o un .zip)", type=["pdf", "zip", "xlsx", "xls"],
                               accept_multiple_files=True,
                               key="lt_pdfs")
    if not subidos:
        st.info("Subí los PDF de la carpeta del mes (Fiserv, Favacard, Naranja y Payway).")
        return
    archivos = tuple((f.name, f.getvalue()) for f in subidos)
    clave = hashlib.sha256(b"".join(n.encode() + hashlib.sha256(d).digest() for n, d in archivos)).hexdigest()
    docs = _leer_cacheado(clave, archivos)
    if s.get("lt_clave") != clave:
        s["lt_clave"] = clave
        s["lt_df"] = a_tabla(docs)
        s["lt_periodo"] = periodo_desde_docs(docs)
        s["lt_ver"] = s.get("lt_ver", 0) + 1
    periodo = c2.text_input("Período", key="lt_periodo")

    ignorados = [d for d in docs if not d["proc"] or d["proc"] == "Ajustes"]
    for d in ignorados:
        st.caption(f"{d['archivo']}: " + (d["notas"][0] if d["notas"] else "no se carga."))
    avisos = [(d["archivo"], n) for d in docs if d["proc"] and d["proc"] != "Ajustes" for n in d["notas"]]
    if avisos:
        with st.expander(f"Avisos de lectura ({len(avisos)})", expanded=False):
            for a, n in avisos:
                st.write(f"**{a}** · {n}")

    st.subheader("Liquidaciones leídas")
    st.caption("Podés corregir cualquier importe o agregar una fila (por ejemplo Favacard, que se carga a mano).")
    cfg = {c: st.column_config.NumberColumn(c, format="%.2f") for c in COLS}
    cfg["Procesadora"] = st.column_config.SelectboxColumn("Procesadora", options=list(PROCESADORAS), required=True)
    base = s["lt_df"].copy()
    if "Banco" not in base.columns:
        base.insert(1, "Banco", "")
    editado = st.data_editor(base, num_rows="dynamic", hide_index=True, key=f"lt_editor_{s['lt_ver']}",
                             column_config=cfg, **_STRETCH)
    df = editado.copy()
    df = df[df["Procesadora"].notna()].reset_index(drop=True)
    if df.empty:
        st.warning("No hay liquidaciones para mostrar.")
        return
    df["Diferencia"] = [float(diferencia(r)) for _, r in df.iterrows()]
    malas = df[df["Diferencia"].abs() > 0.05]
    if len(malas):
        st.warning("Hay liquidaciones donde Total − descuentos − acreditado no da cero:")
        st.dataframe(malas[["Procesadora", "Banco", "Detalle", "Tarjeta", "Diferencia"]], hide_index=True, **_STRETCH)
    else:
        st.success("Todas las liquidaciones cierran contra lo acreditado.")

    st.subheader("Factura de gastos por procesadora")
    gl = grupos(df)
    with st.expander("Nombre y punto de venta de cada comprobante", expanded=False):
        for proc, banco, _, etq in gl:
            st.text_input(etq, value=nombre_defecto(proc, banco), key=f"lt_pv_{etq}")
    pv = {etq: s.get(f"lt_pv_{etq}", nombre_defecto(proc, banco)) for proc, banco, _, etq in gl}
    res = resumen_facturas(df, pv)
    if res.empty:
        return
    st.dataframe(res, hide_index=True, **_STRETCH,
                 column_config={c: st.column_config.NumberColumn(c, format="%.2f")
                                for c in res.columns if c not in ("Procesadora", "Comprobante")})
    st.caption("Las retenciones (R IIBB, R IVA, R Gcias) no forman parte de la factura: se toman directamente de ARCA. "
               "Las percepciones sí.")
    ng = res["No gravado"].sum()
    if ng:
        st.caption(f"No gravado (por ejemplo sellos): {fmt_ar(ng)}. No se incluye en la factura ni en el asiento.")
    st.subheader("Asiento: factura de gastos contra deudores por ventas")
    plan = plan_del_cliente(cliente) if cliente else []
    cuentas = {}
    usa105 = res["IVA 10,5%"].abs().sum() + res["Gasto BI 10,5%"].abs().sum() > 0
    conceptos = [("gasto", "Gasto al 21% (comisiones)")]
    if usa105:
        conceptos.append(("gasto105", "Gasto al 10,5% (comisiones)"))
    conceptos.append(("iva", "IVA crédito fiscal 21%"))
    if usa105:
        conceptos.append(("iva105", "IVA crédito fiscal 10,5%"))
    conceptos += [("p_iva", "Percepción de IVA"), ("p_iibb", "Percepción de IIBB"),
                  ("deudores", "Deudores por ventas (Haber)")]
    with st.expander("Cuentas del asiento", expanded=bool(plan)):
        if plan:
            st.caption(f"Plan de cuentas de **{cliente}** ({len(plan)} cuentas). Se sugiere una cuenta sólo cuando hay "
                       "una única coincidencia; si no, elegila vos.")
            for k, rot in conceptos:
                sug = sugerir_cuenta(plan, k)
                op = [SIN_CUENTA] + plan
                cuentas[k] = st.selectbox(rot, op, index=op.index(sug) if sug else 0, key=f"lt_cta_{k}_{cliente}")
        else:
            st.caption("No encontré el plan de cuentas de este cliente: escribí los nombres tal cual figuran en su plan.")
            for k, rot in conceptos:
                cuentas[k] = st.text_input(rot, value=CUENTAS_DEFECTO[k], key=f"lt_cta_{k}_{cliente or 'x'}")
    faltan = [rot for k, rot in conceptos if cuentas.get(k) in (None, "", SIN_CUENTA)]
    if faltan:
        st.warning("Falta elegir cuenta para: " + ", ".join(faltan))
    asi = armar_asiento(res, cuentas)
    st.dataframe(asi, hide_index=True, **_STRETCH,
                 column_config={c: st.column_config.NumberColumn(c, format="%.2f") for c in ("Debe", "Haber")})
    st.download_button("Descargar Excel", generar_excel(df.drop(columns=["Diferencia"]), cliente or "", periodo, pv, cuentas),
                       file_name=f"LIQUIDACIONES TARJETAS {periodo or ''}.xlsx".replace("  ", " ").strip(),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="lt_dl_xlsx")
