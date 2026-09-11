# -*- coding: utf-8 -*-
"""Conciliación Growth: deudores (cobranzas) y proveedores (pagos + Visa cuota×N)."""
from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_formato_estudio import guardar_informe_excel
from procesador import (
    _normalizar_texto,
    cargar_facturas_proveedores_excel,
    ejecutar_match_debitos_proveedores,
    matchear_debitos_con_facturas,
    procesar_extractos_bancarios_pdfs,
)

BANCOS = Path(
    r"T:\CLIENTES\GROWTH MARKETING SA\Cierre 2025 01-07-2025 al 30-06-2026"
    r"\Papeles de Trabajo\Bancos"
)
MAYOR = Path(
    r"T:\CLIENTES\GROWTH MARKETING SA\Balance\2026\deudores y proveedores.xlsx"
)
OUT_T = BANCOS / "Conciliacion_Deudores_Proveedores_Growth_2026.xlsx"
OUT_DESK = Path(r"C:\Users\recep\Desktop\Conciliacion_Deudores_Proveedores_Growth_2026.xlsx")

RE_VISA = re.compile(
    r"^(?P<fecha>\d{2}-\d{2}-\d{2})\s+(?P<marca>[*KF])?\s*"
    r"(?P<ref>.+?)\s+"
    r"(?:(?P<cuota>\d{2})/(?P<plazo>\d{2})\s+)?"
    r"(?P<comp>\d{5,})?\s*"
    r"(?P<pesos>\d{1,3}(?:\.\d{3})*,\d{2})\s*$",
    re.I,
)
RE_VISA_USD = re.compile(
    r"^(?P<fecha>\d{2}-\d{2}-\d{2})\s+(?P<marca>[*KF])?\s*"
    r"(?P<ref>.+?)\s+"
    r"(?P<usd>\d{1,3}(?:\.\d{3})*,\d{2})\s+\d+\s+"
    r"(?P<pesos>\d{1,3}(?:\.\d{3})*,\d{2})\s*$",
    re.I,
)
SKIP_VISA = (
    "IMPUESTO", "COMISION", "IVA", "PERCEP", "IIBB", "SELLOS", "TOTAL",
    "SU PAGO", "DEV.IMP", "PAGINA", "RESUMEN", "TARJETA", "SALDO",
)
MP_PAGO = ("pago", "retiro de dinero")
MP_COBRO = ("cobro", "ingreso de dinero")


def _money(s) -> float:
    t = str(s or "").strip().replace("$", "").replace(" ", "")
    if not t:
        return 0.0
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return round(float(t), 2)
    except ValueError:
        return 0.0


def _fecha_visa(raw: str):
    try:
        return datetime.strptime(raw, "%d-%m-%y").date()
    except ValueError:
        return None


def pdfs_galicia() -> tuple[list[Path], list[dict]]:
    """Digitales (~300 KB). Los escaneados pesados se dejan para OCR aparte."""
    out = []
    avisos = []
    for p in sorted(BANCOS.glob("*.pdf")):
        n = p.name.upper()
        if "INVERSION" in n or "USD" in n or "VISA" in n or "TARJETA" in n:
            continue
        if "MERCADO" in n:
            continue
        if p.stat().st_size > 800_000:
            avisos.append(
                {
                    "archivo": p.name,
                    "motivo": "PDF pesado/escaneado — no se OCR en este lote (sep-25 / nov-25)",
                }
            )
            continue
        out.append(p)
    return out, avisos


def parsear_visa(proveedores_norm: list[str]) -> tuple[pd.DataFrame, list[dict]]:
    filas = []
    ocr = []
    for p in sorted(BANCOS.glob("*.pdf")):
        n = p.name.upper()
        if "VISA" not in n and "TARJETA" not in n:
            continue
        try:
            import pdfplumber

            with pdfplumber.open(p) as pdf:
                texto = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
        except Exception as exc:
            ocr.append({"archivo": p.name, "motivo": str(exc)})
            continue
        if len(texto.strip()) < 200:
            ocr.append({"archivo": p.name, "motivo": "PDF sin texto (escaneado)"})
            continue
        for line in texto.splitlines():
            s = line.strip()
            if not s or any(k in s.upper() for k in SKIP_VISA):
                continue
            m = RE_VISA.match(s) or RE_VISA_USD.match(s)
            if not m:
                continue
            fecha = _fecha_visa(m.group("fecha"))
            if not fecha:
                continue
            ref = re.sub(r"\s+", " ", m.group("ref")).strip()
            pesos = _money(m.group("pesos"))
            if pesos <= 0 or not ref:
                continue
            cuota = m.groupdict().get("cuota")
            plazo = m.groupdict().get("plazo")
            n_cuotas = int(plazo) if plazo else 1
            n_esta = int(cuota) if cuota else 1
            nombre_norm = _normalizar_texto(ref)
            es_prov = any(
                nombre_norm and (nombre_norm in pn or pn in nombre_norm or _fuzzy(nombre_norm, pn) >= 70)
                for pn in proveedores_norm
                if pn
            )
            total_plan = round(pesos * n_cuotas, 2) if (es_prov and n_cuotas > 1) else pesos
            filas.append(
                {
                    "Fecha": fecha,
                    "Descripcion": f"TRF INMED PROVEED {ref}" if es_prov else f"PAGO TARJETA {ref}",
                    "Detalle": f"Visa {n_esta}/{n_cuotas}" if n_cuotas > 1 else "Visa",
                    "Importe": -pesos,
                    "Importe match": -total_plan,
                    "Cuotas": n_cuotas,
                    "Cuota nro": n_esta,
                    "Es proveedor (nombre)": "SI" if es_prov else "",
                    "Origen": "Visa Galicia",
                    "Archivo": p.name,
                }
            )
    return pd.DataFrame(filas), ocr


def _fuzzy(a: str, b: str) -> float:
    try:
        from rapidfuzz import fuzz

        return float(fuzz.token_set_ratio(a, b))
    except Exception:
        return 100.0 if a and b and (a in b or b in a) else 0.0


def parsear_mp() -> pd.DataFrame:
    filas = []
    vistos = set()
    for p in sorted(BANCOS.glob("*.xlsx")):
        if "mercado" not in p.name.lower() and "pago" not in p.name.lower():
            continue
        if p.name.lower().startswith("conciliacion"):
            continue
        if "galicia" in p.name.lower():
            continue
        try:
            df = pd.read_excel(p, header=0)
        except Exception:
            continue
        cols = {_normalizar_texto(c): c for c in df.columns}
        c_fecha = cols.get("fecha de pago") or cols.get("fecha")
        c_tipo = cols.get("tipo de operacion") or cols.get("tipo")
        c_imp = cols.get("importe")
        c_nro = cols.get("numero de movimiento")
        if not (c_fecha and c_tipo and c_imp):
            continue
        for _, row in df.iterrows():
            tipo = str(row[c_tipo] or "")
            tn = _normalizar_texto(tipo)
            imp = pd.to_numeric(row[c_imp], errors="coerce")
            if pd.isna(imp) or abs(float(imp)) < 0.01:
                continue
            fecha = pd.to_datetime(row[c_fecha], errors="coerce")
            if pd.isna(fecha):
                continue
            nro = str(row[c_nro]) if c_nro else ""
            clave = (str(fecha.date()), tn, round(float(imp), 2), nro)
            if clave in vistos:
                continue
            vistos.add(clave)
            es_pago = any(k in tn for k in MP_PAGO) and float(imp) < 0
            es_cobro = any(k in tn for k in MP_COBRO) and float(imp) > 0
            if not es_pago and not es_cobro:
                continue
            if "retiro" in tn:
                continue
            desc = f"TRF INMED PROVEED {tipo}" if es_pago else f"TRANSFERENCIA {tipo}"
            filas.append(
                {
                    "Fecha": fecha.date(),
                    "Descripcion": desc,
                    "Detalle": nro,
                    "Importe": round(float(imp), 2),
                    "Importe match": round(float(imp), 2),
                    "Cuotas": 1,
                    "Cuota nro": 1,
                    "Es proveedor (nombre)": "",
                    "Origen": "Mercado Pago",
                    "Archivo": p.name,
                }
            )
    return pd.DataFrame(filas)


def normalizar_extracto_galicia(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    work = df.copy()
    if "Fecha" not in work.columns:
        return pd.DataFrame()
    work["Fecha"] = pd.to_datetime(work["Fecha"], dayfirst=True, errors="coerce").dt.date
    if "Importe" in work.columns:
        work["Importe"] = pd.to_numeric(work["Importe"], errors="coerce")
    elif "Debito" in work.columns:
        deb = pd.to_numeric(work.get("Debito"), errors="coerce").fillna(0)
        cre = pd.to_numeric(work.get("Credito"), errors="coerce").fillna(0)
        work["Importe"] = cre - deb
    else:
        return pd.DataFrame()
    desc = work["Descripcion"].astype(str) if "Descripcion" in work.columns else ""
    det = work["Detalle"].astype(str) if "Detalle" in work.columns else ""
    work["Descripcion"] = (desc + " " + det).str.strip()
    work["Detalle"] = det if "Detalle" in work.columns else ""
    work["Importe match"] = work["Importe"]
    work["Cuotas"] = 1
    work["Cuota nro"] = 1
    work["Es proveedor (nombre)"] = ""
    work["Origen"] = "Galicia"
    if "Archivo origen" in work.columns:
        work["Archivo"] = work["Archivo origen"]
    elif "Archivo" not in work.columns:
        work["Archivo"] = "Galicia PDF"
    return work[
        [
            "Fecha", "Descripcion", "Detalle", "Importe", "Importe match",
            "Cuotas", "Cuota nro", "Es proveedor (nombre)", "Origen", "Archivo",
        ]
    ].dropna(subset=["Fecha"])


def cargar_deudores(path: Path) -> pd.DataFrame:
    hojas = pd.read_excel(path, sheet_name=None)
    df = None
    for frame in hojas.values():
        cols = {_normalizar_texto(str(c)) for c in frame.columns}
        if "razon social" in cols or any("razon" in c for c in cols):
            df = frame
            break
    if df is None:
        df = list(hojas.values())[-1]
    mapeo = {}
    for col in df.columns:
        cn = _normalizar_texto(str(col))
        if "fecha" in cn:
            mapeo[col] = "fecha"
        elif "razon" in cn or "social" in cn:
            mapeo[col] = "proveedor"
        elif "numero" in cn and "comprob" in cn:
            mapeo[col] = "comprobante"
        elif "tipo comprob" in cn:
            mapeo[col] = "tipo"
        elif cn == "debe":
            mapeo[col] = "debe"
        elif cn == "haber":
            mapeo[col] = "haber"
        elif "cuenta" in cn:
            mapeo[col] = "cuenta"
    work = df.rename(columns=mapeo).copy()
    work["fecha"] = pd.to_datetime(work["fecha"], errors="coerce").dt.date
    work["debe"] = pd.to_numeric(work.get("debe"), errors="coerce").fillna(0)
    work["haber"] = pd.to_numeric(work.get("haber"), errors="coerce").fillna(0)
    work = work[work["cuenta"].astype(str).str.contains("11301", na=False)].copy()
    work = work[work["debe"] > 0.009].copy()
    work["importe"] = work["debe"].round(2)
    work["proveedor"] = work["proveedor"].astype(str).str.strip()
    work["proveedor_norm"] = work["proveedor"].map(_normalizar_texto)
    work["comprobante"] = work.get("comprobante", "").astype(str)
    work["tipo"] = work.get("tipo", "").astype(str)
    work = work[work["proveedor_norm"].str.len() > 1].copy().reset_index(drop=True)
    work["factura_id"] = work.index.astype(int)
    return work[
        ["factura_id", "fecha", "proveedor", "proveedor_norm", "importe", "comprobante", "tipo", "debe", "haber"]
    ]


def df_a_debitos_match(movs: pd.DataFrame, *, solo_egresos: bool) -> pd.DataFrame:
    if movs.empty:
        return pd.DataFrame()
    work = movs.copy()
    work["importe_raw"] = pd.to_numeric(work["Importe match"], errors="coerce").fillna(0)
    if solo_egresos:
        work = work[work["importe_raw"] < -0.009].copy()
        work["importe"] = work["importe_raw"].abs()
    else:
        work = work[work["importe_raw"] > 0.009].copy()
        work["importe"] = work["importe_raw"]
    work["fecha"] = work["Fecha"]
    work["descripcion"] = (
        work["Descripcion"].astype(str) + " " + work["Detalle"].astype(str)
    )
    work["descripcion_norm"] = work["descripcion"].map(_normalizar_texto)
    work = work.reset_index(drop=True)
    work["debito_id"] = work.index.astype(int)
    work["comprobante"] = work.get("Archivo", "")
    return work


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    print("1) Mayor deudores/proveedores…", flush=True)
    fact_prov = cargar_facturas_proveedores_excel(MAYOR)
    fact_deu = cargar_deudores(MAYOR)
    print(f"   proveedores FCC {len(fact_prov)} · deudores FCV {len(fact_deu)}")

    print("2) Galicia PDFs…")
    g_pdfs, avisos_scan = pdfs_galicia()
    print("  ", len(g_pdfs), "archivos digitales ·", len(avisos_scan), "escaneados omitidos")

    class _PdfUp:
        def __init__(self, path: Path) -> None:
            self.name = path.name
            self._data = path.read_bytes()

        def getvalue(self) -> bytes:
            return self._data

    df_g, meta_g, err_g = procesar_extractos_bancarios_pdfs([_PdfUp(p) for p in g_pdfs])
    print("   filas", 0 if df_g is None else len(df_g), "errores", err_g)
    gal = normalizar_extracto_galicia(df_g)

    print("3) Mercado Pago…")
    mp = parsear_mp()
    print("   filas", len(mp))

    print("4) Visa (cuota×N si nombre=proveedor)…")
    visa, ocr_visa = parsear_visa(fact_prov["proveedor_norm"].tolist())
    print("   filas", len(visa), "ocr", ocr_visa)
    if not visa.empty:
        n_exp = int((visa["Cuotas"] > 1).sum())
        print("   cuotas expandidas", n_exp)

    movs = pd.concat([gal, mp, visa], ignore_index=True)
    movs = movs.dropna(subset=["Fecha"]).sort_values("Fecha").reset_index(drop=True)
    print("5) Movimientos unificados", len(movs))

    print("6) Match proveedores…")
    debitos = df_a_debitos_match(movs, solo_egresos=True)
    res_prov = matchear_debitos_con_facturas(debitos, fact_prov)

    print("7) Match deudores (cobranzas)…")
    creditos = df_a_debitos_match(movs, solo_egresos=False)
    res_deu = matchear_debitos_con_facturas(creditos, fact_deu)

    def _kpi(res, label):
        cal = res.get("calzados")
        n = 0 if cal is None or cal.empty else len(cal)
        imp = 0.0 if cal is None or cal.empty else float(
            pd.to_numeric(cal.get("Debito banco", cal.get("importe", 0)), errors="coerce").fillna(0).sum()
        )
        print(f"   {label}: calzados {n}  ${imp:,.2f}")
        return n, imp

    n_p, i_p = _kpi(res_prov, "Proveedores")
    n_d, i_d = _kpi(res_deu, "Deudores")

    resumen = pd.DataFrame(
        [
            {"Concepto": "Facturas proveedores (FCC)", "Cantidad": len(fact_prov), "Importe": float(fact_prov["importe"].sum())},
            {"Concepto": "Facturas deudores (FCV)", "Cantidad": len(fact_deu), "Importe": float(fact_deu["importe"].sum())},
            {"Concepto": "Movimientos Galicia", "Cantidad": len(gal), "Importe": float(gal["Importe"].sum()) if not gal.empty else 0},
            {"Concepto": "Movimientos Mercado Pago (pago/cobro)", "Cantidad": len(mp), "Importe": float(mp["Importe"].sum()) if not mp.empty else 0},
            {"Concepto": "Consumos Visa", "Cantidad": len(visa), "Importe": float(visa["Importe"].abs().sum()) if not visa.empty else 0},
            {"Concepto": "Calzados proveedores", "Cantidad": n_p, "Importe": i_p},
            {"Concepto": "Calzados deudores", "Cantidad": n_d, "Importe": i_d},
        ]
    )

    hojas = [
        ("Calzados proveedores", res_prov.get("calzados") if res_prov.get("calzados") is not None else pd.DataFrame()),
        ("Pagos sin factura", res_prov.get("pagos_sin_factura") if res_prov.get("pagos_sin_factura") is not None else pd.DataFrame()),
        ("FCC impagas", res_prov.get("facturas_impagas") if res_prov.get("facturas_impagas") is not None else pd.DataFrame()),
        ("Calzados deudores", res_deu.get("calzados") if res_deu.get("calzados") is not None else pd.DataFrame()),
        ("Cobranzas sin factura", res_deu.get("pagos_sin_factura") if res_deu.get("pagos_sin_factura") is not None else pd.DataFrame()),
        ("FCV impagas", res_deu.get("facturas_impagas") if res_deu.get("facturas_impagas") is not None else pd.DataFrame()),
        ("Movimientos", movs),
    ]
    if ocr_visa:
        hojas.append(("Visa OCR", pd.DataFrame(ocr_visa)))
    avisos_g = list(err_g or []) + avisos_scan
    if avisos_g:
        hojas.append(("Galicia avisos", pd.DataFrame(avisos_g)))

    for ruta in (OUT_T, OUT_DESK):
        guardar_informe_excel(
            ruta,
            titulo="Conciliación deudores y proveedores",
            subtitulo="GROWTH STARTUP AGENCY SA · Galicia + Mercado Pago + Visa",
            periodo="01/07/2025 al 30/06/2026",
            kpis=[
                ("Calzados proveedores", n_p, "int"),
                ("Calzados deudores", n_d, "int"),
                ("Movimientos", len(movs), "int"),
            ],
            resumenes=[("Resumen", resumen)],
            detalle=movs,
            hoja_detalle="Movimientos",
            hojas_adicionales=hojas,
            col_moneda=["Importe", "Importe match", "Debito banco", "Importe factura"],
            col_fecha=["Fecha", "Fecha banco", "Fecha factura"],
            total_col="Importe",
        )
        print("OK", ruta)


if __name__ == "__main__":
    main()
