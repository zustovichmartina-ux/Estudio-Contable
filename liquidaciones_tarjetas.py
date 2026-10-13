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

PROCESADORAS = ("Payway", "Fiserv", "Naranja", "Favacard")
# Nombre y punto de venta con que el estudio carga cada procesadora (editable en pantalla).
PV_DEFECTO = {"Payway": "FRANCES - 00001", "Naranja": "NARANJA - 00002",
              "Favacard": "FAVACARD - 00003", "Fiserv": "FIRST DATA - 00004"}
ORDEN_PROC = ("Payway", "Favacard", "Naranja", "Fiserv")

# columnas numéricas de cada liquidación (orden de la planilla del estudio)
COLS = ("Total presentado", "BI 21%", "IVA 21%", "P IIBB", "R IIBB", "P IVA", "R IVA", "R Gcias",
        "No gravado", "Acreditado PDF")
CAMPOS = ("presentado", "bi", "iva", "p_iibb", "r_iibb", "p_iva", "r_iva", "r_gcias", "no_grav", "neto")
CAMPO_A_COL = dict(zip(CAMPOS, COLS))

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


def parsear_payway(tok: list[str], nombre: str) -> dict:
    texto = "\n".join(tok)
    r = _vacio()
    notas: list[str] = []
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
        return {"datos": r, "marca": _marca_por_nombre(nombre), "notas": notas + ["No encontré el desglose de descuentos."],
                "fecha": ""}
    fin = len(tok)
    for k in range(i0, len(tok)):
        if tok[k].startswith("(*) Servicio") or tok[k].startswith("SR. COMERCIANTE"):
            fin = k
            break
    bloque = tok[i0 + 1:fin]
    previos: list[str] = []
    arancel_serv = Decimal(0)
    for t in bloque:
        if re.match(r"^\$\s", t) and es_importe(t):
            monto = a_dec(t)
            etiqueta = " ".join(x for x in previos if re.search(r"[A-Za-zÁÉÍÓÚáéíóú]", x)).lower()
            previos = []
            if "base imponible" in etiqueta:
                continue
            if "arancel" in etiqueta or "servicio" in etiqueta or "cargo por serv" in etiqueta \
                    or "bonif" in etiqueta:
                arancel_serv += monto
            elif re.search(r"\biva\b", etiqueta) and "ret" not in etiqueta and "perc" not in etiqueta:
                r["iva"] += monto
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
    r["bi"] = arancel_serv
    if descuento is not None:
        suma = r["bi"] + r["iva"] + r["p_iibb"] + r["r_iibb"] + r["p_iva"] + r["r_iva"] + r["r_gcias"] + r["no_grav"]
        if (suma - descuento).copy_abs() > Decimal("0.05"):
            notas.append(f"Los descuentos del desglose ({fmt_ar(suma)}) no coinciden con el total del encabezado ({fmt_ar(descuento)}).")
    fecha = ""
    mf = re.search(r"(\d{2}/\d{2}/\d{4})\n000\d+", texto)
    if mf:
        fecha = mf.group(1)
    return {"datos": r, "marca": _marca_por_nombre(nombre), "notas": sorted(set(notas)), "fecha": fecha}


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


def clasificar(tok: list[str]) -> str:
    t = "\n".join(tok[:400])
    if "FAVACARD" in t.upper() and any(x.startswith("PROXIMOS PAGOS") for x in tok):
        return "Favacard"
    if "Resumen Mensual de Liquidaciones" in t or "Servicio PAYWAY" in "\n".join(tok):
        return "Payway"
    if "FIRST DATA" in t.upper() and "Entidad Pagadora" in t:
        return "Fiserv"
    if "Tarjeta Naranja" in t and "Detalle de Cupones Liquidados" in "\n".join(tok):
        return "Naranja"
    if "Detalle de ajustes" in t:
        return "Ajustes"
    return ""


def leer_archivo(nombre: str, data: bytes) -> dict:
    doc = _abrir(data)
    base = {"archivo": nombre.split("/")[-1], "proc": "", "marca": "", "datos": _vacio(), "notas": [], "fecha": ""}
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
    elif proc == "Ajustes":
        base["notas"] = ["Detalle de ajustes/promociones: no se carga."]
        return base
    else:
        base["notas"] = ["No reconocí de qué procesadora es este PDF."]
        return base
    base.update(res)
    return base


def expandir(archivos: list[tuple[str, bytes]]) -> list[tuple[str, bytes]]:
    out = []
    for nombre, data in archivos:
        if nombre.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        if not info.is_dir() and info.filename.lower().endswith(".pdf") and info.file_size < 40_000_000:
                            out.append((f"{nombre}/{info.filename}", z.read(info)))
            except zipfile.BadZipFile:
                pass
        else:
            out.append((nombre, data))
    return out


def leer_archivos(archivos: list[tuple[str, bytes]]) -> list[dict]:
    docs = [leer_archivo(n, d) for n, d in expandir(archivos)]
    orden = {p: i for i, p in enumerate(ORDEN_PROC)}
    docs.sort(key=lambda d: (orden.get(d["proc"], 9), d["fecha"][6:] + d["fecha"][3:5] + d["fecha"][:2], d["archivo"]))
    return docs


# ---------------------------------------------------------------- tabla y cálculos
def a_tabla(docs: list[dict]) -> pd.DataFrame:
    filas = []
    for d in docs:
        if not d["proc"] or d["proc"] == "Ajustes":
            continue
        fila = {"Procesadora": d["proc"], "Detalle": d["archivo"], "Tarjeta": d["marca"]}
        for c in CAMPOS:
            fila[CAMPO_A_COL[c]] = float(d["datos"][c])
        filas.append(fila)
    cols = ["Procesadora", "Detalle", "Tarjeta"] + list(COLS)
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
    return (g("Total presentado") - g("BI 21%") - g("IVA 21%") - g("P IIBB") - g("R IIBB") - g("P IVA")
            - g("R IVA") - g("R Gcias") - g("No gravado") - g("Acreditado PDF")).quantize(Q2, ROUND_HALF_UP)


def resumen_facturas(df: pd.DataFrame, pv: dict[str, str]) -> pd.DataFrame:
    filas = []
    for proc in ORDEN_PROC:
        sub = df[df["Procesadora"] == proc]
        if sub.empty:
            continue
        s = lambda c: sum((_dec(v) for v in sub[c]), Decimal(0)).quantize(Q2, ROUND_HALF_UP)
        gasto = s("BI 21%") + s("No gravado")
        total = gasto + s("IVA 21%") + s("P IVA") + s("P IIBB")
        filas.append({
            "Procesadora": proc, "Comprobante": pv.get(proc, PV_DEFECTO.get(proc, proc)),
            "Gasto (BI 21%)": float(s("BI 21%")), "No gravado": float(s("No gravado")),
            "IVA 21%": float(s("IVA 21%")),
            "P IVA": float(s("P IVA")), "P IIBB": float(s("P IIBB")), "Total factura": float(total),
            "R IIBB": float(s("R IIBB")), "R IVA": float(s("R IVA")), "R Gcias": float(s("R Gcias")),
        })
    return pd.DataFrame(filas)


# ---------------------------------------------------------------- asiento
# Nombres de cuenta por defecto: editables en pantalla. Deben coincidir con la hoja CUENTAS del estudio.
CUENTAS_DEFECTO = {"gasto": "Gastos por comisiones tarjetas", "iva": "IVA Crédito Fiscal",
                   "p_iva": "Percepciones IVA", "p_iibb": "Percepciones IIBB", "deudores": "Deudores por ventas"}
SIN_CUENTA = "(elegir cuenta)"

# patrones para sugerir cuenta del plan; sólo se sugiere si hay UNA sola coincidencia
_PATRONES = {
    "gasto": (r"comisi\w*.*tarjeta|tarjeta.*comisi", r"gastos?.*tarjeta"),
    "iva": (r"iva.*cr[eé]dito.*21|iva.*21.*cr[eé]dito", r"iva.*cr[eé]dito"),
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
        deb = [(cu["gasto"], _dec(r["Gasto (BI 21%)"]) + _dec(r["No gravado"])), (cu["iva"], _dec(r["IVA 21%"])),
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

    ws = wb.add_worksheet("LIQUIDACIONES")
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 46)
    ws.set_column("C:C", 18)
    ws.set_column("D:N", 15)
    ws.write("B1", cliente or "", tit)
    fecha_mes = _fecha_periodo(periodo)
    if fecha_mes:
        ws.write_datetime("B2", fecha_mes, mes)
    else:
        ws.write("B2", periodo or "", sub)
    fila = 3  # 0-index -> fila 4
    ref_tot = {}  # proc -> fila de totales (1-index)
    cab = ["Detalle", "Tarjeta", "TOTAL", "BI 21", "IVA 21%", "P IIBB", "R IIBB", "P IVA", "R IVA", "R GCIAS",
           "NO GRAVADO", "NETO", "ACREDITADO PDF", "DIFERENCIA"]
    for proc in ORDEN_PROC:
        sub_df = df[df["Procesadora"] == proc]
        if sub_df.empty:
            continue
        ws.write(fila, 1, pv.get(proc, PV_DEFECTO.get(proc, proc)), sub)
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
            # NETO = TOTAL - descuentos y retenciones
            ws.write_formula(fila, 12, f"=D{fx}-SUM(E{fx}:L{fx})", num)
            ws.write_number(fila, 13, float(_dec(r["Acreditado PDF"])), num_in)
            ws.write_formula(fila, 14, f"=M{fx}-N{fx}", num)
            fila += 1
        fin = fila
        ws.write(fila, 1, "TOTAL", tot_t)
        ws.write(fila, 2, "", tot_t)
        for k in range(3, 15):
            col = chr(ord("A") + k)
            ws.write_formula(fila, k, f"=SUM({col}{ini}:{col}{fin})", tot)
        ref_tot[proc] = fila + 1
        fila += 3
    # hoja de facturas
    wf = wb.add_worksheet("FACTURAS")
    wf.set_column("A:A", 3)
    wf.set_column("B:B", 24)
    wf.set_column("C:L", 16)
    wf.write("B1", cliente or "", tit)
    if fecha_mes:
        wf.write_datetime("B2", fecha_mes, mes)
    else:
        wf.write("B2", periodo or "", sub)
    cab2 = ["Comprobante", "Gasto (BI 21%)", "No gravado", "IVA 21%", "P IVA", "P IIBB", "TOTAL FACTURA",
            "R IIBB", "R IVA", "R GCIAS"]
    for j, c in enumerate(cab2):
        wf.write(3, 1 + j, c, head)
    wf.set_row(3, 30)
    fr = 4
    fila_fact: dict[str, int] = {}
    for proc in ORDEN_PROC:
        if proc not in ref_tot:
            continue
        t = ref_tot[proc]
        wf.write(fr, 1, pv.get(proc, PV_DEFECTO.get(proc, proc)), txt)
        x = fr + 1
        fila_fact[proc] = x
        wf.write_formula(fr, 2, f"=LIQUIDACIONES!E{t}", num)
        wf.write_formula(fr, 3, f"=LIQUIDACIONES!L{t}", num)
        wf.write_formula(fr, 4, f"=LIQUIDACIONES!F{t}", num)
        wf.write_formula(fr, 5, f"=LIQUIDACIONES!I{t}", num)
        wf.write_formula(fr, 6, f"=LIQUIDACIONES!G{t}", num)
        wf.write_formula(fr, 7, f"=SUM(C{x}:G{x})", ok)
        wf.write_formula(fr, 8, f"=LIQUIDACIONES!H{t}", num)
        wf.write_formula(fr, 9, f"=LIQUIDACIONES!J{t}", num)
        wf.write_formula(fr, 10, f"=LIQUIDACIONES!K{t}", num)
        fr += 1
    if fr > 4:
        wf.write(fr, 1, "TOTAL", tot_t)
        for k in range(2, 11):
            col = chr(ord("A") + k)
            wf.write_formula(fr, k, f"=SUM({col}5:{col}{fr})", tot)
    # hoja de asientos (factura de gastos contra proveedores)
    cu = {**CUENTAS_DEFECTO, **(cuentas or {})}
    wa = wb.add_worksheet("ASIENTO")
    wa.set_column("A:A", 3)
    wa.set_column("B:B", 46)
    wa.set_column("C:D", 18)
    wa.write("B1", cliente or "", tit)
    if fecha_mes:
        wa.write_datetime("B2", fecha_mes, mes)
    else:
        wa.write("B2", periodo or "", sub)
    ar = 3
    for proc in ORDEN_PROC:
        if proc not in fila_fact:
            continue
        x = fila_fact[proc]
        wa.write(ar, 1, pv.get(proc, PV_DEFECTO.get(proc, proc)) + " · factura de gastos", sub)
        ar += 1
        for j, c in enumerate(["Cuenta", "Debe", "Haber"]):
            wa.write(ar, 1 + j, c, head)
        ar += 1
        ini = ar + 1
        for cta, fd in ((cu["gasto"], f"=FACTURAS!C{x}+FACTURAS!D{x}"), (cu["iva"], f"=FACTURAS!E{x}"),
                        (cu["p_iva"], f"=FACTURAS!F{x}"), (cu["p_iibb"], f"=FACTURAS!G{x}")):
            wa.write(ar, 1, cta, txt)
            wa.write_formula(ar, 2, fd, num)
            wa.write_blank(ar, 3, None, num)
            ar += 1
        wa.write(ar, 1, cuenta_haber(cu), txt)
        wa.write_blank(ar, 2, None, num)
        wa.write_formula(ar, 3, f"=FACTURAS!H{x}", num)
        ar += 1
        wa.write(ar, 1, "TOTAL", tot_t)
        wa.write_formula(ar, 2, f"=SUM(C{ini}:C{ar})", tot)
        wa.write_formula(ar, 3, f"=SUM(D{ini}:D{ar})", tot)
        ar += 1
        wa.write(ar, 1, "Diferencia Debe - Haber", txt)
        wa.write_formula(ar, 2, f"=C{ar}-D{ar}", ok)
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
    st.caption("Subí los PDF del mes de Payway, Fiserv (First Data) y Naranja. Se leen solos, se controlan contra "
               "lo acreditado y se arma una factura de gastos por procesadora.")
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
    subidos = st.file_uploader("PDF de liquidaciones (o un .zip)", type=["pdf", "zip"], accept_multiple_files=True,
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
        st.dataframe(malas[["Procesadora", "Detalle", "Tarjeta", "Diferencia"]], hide_index=True, **_STRETCH)
    else:
        st.success("Todas las liquidaciones cierran contra lo acreditado.")

    st.subheader("Factura de gastos por procesadora")
    with st.expander("Nombre y punto de venta de cada comprobante", expanded=False):
        pv = {}
        for p in ORDEN_PROC:
            pv[p] = st.text_input(p, value=PV_DEFECTO[p], key=f"lt_pv_{p}")
    pv = {p: s.get(f"lt_pv_{p}", PV_DEFECTO[p]) for p in ORDEN_PROC}
    res = resumen_facturas(df, pv)
    if res.empty:
        return
    st.dataframe(res, hide_index=True, **_STRETCH,
                 column_config={c: st.column_config.NumberColumn(c, format="%.2f")
                                for c in res.columns if c not in ("Procesadora", "Comprobante")})
    st.caption("Las retenciones (R IIBB, R IVA, R Gcias) no forman parte de la factura: se toman directamente de ARCA. "
               "Las percepciones sí.")
    st.subheader("Asiento: factura de gastos contra deudores por ventas")
    plan = plan_del_cliente(cliente) if cliente else []
    cuentas = {}
    conceptos = (("gasto", "Gasto (BI 21% + no gravado)"), ("iva", "IVA 21%"), ("p_iva", "Percepción de IVA"),
                 ("p_iibb", "Percepción de IIBB"), ("deudores", "Deudores por ventas (Haber)"))
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
