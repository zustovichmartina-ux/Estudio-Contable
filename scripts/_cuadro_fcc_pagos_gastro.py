"""Cuadro por proveedor: FCC/NCC/NDC Tango 21101 vs pagos en extracto banco."""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_formato_estudio import (
    BODY_FONT,
    BOLD_FONT,
    COLOR_PRIMARIO,
    DATE_FMT,
    HDR_FILL,
    HDR_FONT,
    MONEY_FMT,
    SECTION_FONT,
    SUB_FONT,
    TITLE_FONT,
    ZEBRA,
    guardar_informe_excel,
)
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from procesador import (
    _corregir_filas_extracto_por_saldos,
    _normalizar_texto,
    es_debito_pago_o_transferencia,
)

TANGO = Path(r"C:\Users\recep\Desktop\Gastr.xlsx")
CACHE = Path(r"C:\Users\recep\Desktop\Estudio Contable\_cache_extractos_gastro")
BANCOS = Path(
    r"\\TANGOSRV\Compartido\CLIENTES"
    r"\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
    r"\Balances\31-08-2026\Bancos\Cuenta Corriente Nº 067-0850315\2026"
)
OUT = Path(r"C:\Users\recep\Desktop\Mayor_Proveedores_Gastro.xlsx")

# OCR a veces pega un "60 millones" al importe. Caso D Amico 30/10/2025:
# el extracto / FCC 200002105 es $272.250, no $60.272.250.
CORRECCIONES_OCR_DEBITO: list[tuple[str, str, float, float]] = [
    ("2025-10-30", "27209561684", 60_272_250.0, 272_250.0),
]

# Tokens extra cuando el extracto no trae razón social (solo CUIT / nombre cortado)
ALIAS_EXTRA: dict[str, tuple[str, ...]] = {
    "CIRUGIA J.F. S.A.": ("cirugia", "30709314978", "709314978"),
    "H.TRUJILLO CONTADORES PUBLICOS S.A.": ("trujillo", "h trujillo"),
    "CENTRO MEDICO DE MAR DEL PLATA": ("centro medico", "30541190518"),
    "PENSAR DIGITAL S.R.L.": ("pensar digital", "30715823663"),
    "OMNIASALUD S.R.L.": ("omniasalud", "30714747831"),
    "AVIANI SILVANA NOEMI": ("aviani", "27173834972"),
    "LISTORTI ANTONELLA ANA": ("listorti", "23160237384"),
    "LOZZI ADRIANA VALERIA": ("lozzi adriana", "23328483424"),
    "LOZZI RUBEN DARIO": ("lozzi ruben", "20146714510"),
    "PASTORINO MARTIN": ("pastorino", "20286081283"),
    "D AMICO MARIANA": ("d amico", "damico", "27209561684"),
    "ARCE ENRIQUE MIGUEL ANGEL": ("arce enrique", "20302079030"),
    "MEDRANO LOPEZ CARLOS ANDRES": ("medrano", "20955912102"),
    "GARCIA FLORENCIA NADIA": ("garcia florencia", "27352474830"),
    "MARTINEZ SONIA MABEL": ("martinez sonia", "27161485409"),
}


def _cuit_digits(v: object) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    return d if len(d) == 11 else ""


def _aplicar_correcciones_ocr(ext: pd.DataFrame) -> pd.DataFrame:
    blob_digits = (
        ext.get("Descripcion", pd.Series("", index=ext.index)).fillna("").astype(str)
        + " "
        + ext.get("Detalle", pd.Series("", index=ext.index)).fillna("").astype(str)
    ).map(lambda s: re.sub(r"\D", "", s))
    fechas = pd.to_datetime(ext["Fecha"], errors="coerce").dt.normalize()
    for fecha, cuit, mal, ok in CORRECCIONES_OCR_DEBITO:
        mask = (
            (fechas == pd.Timestamp(fecha))
            & blob_digits.str.contains(cuit, na=False)
            & (ext["Debito"].round(2) == round(mal, 2))
        )
        if mask.any():
            ext.loc[mask, "Debito"] = ok
            ext.loc[mask, "Importe"] = -ok
    return ext


def cargar_extracto() -> pd.DataFrame:
    frames = [pd.read_pickle(p) for p in sorted(CACHE.glob("*.pkl"))]
    if not frames:
        raise SystemExit("No hay cache de extractos Gastro")
    ext = pd.concat(frames, ignore_index=True)
    ext = pd.DataFrame(_corregir_filas_extracto_por_saldos(ext.to_dict("records")))
    ext["Fecha"] = pd.to_datetime(ext["Fecha"], dayfirst=True, errors="coerce")
    ext["Debito"] = pd.to_numeric(ext.get("Debito"), errors="coerce")
    ext["Importe"] = pd.to_numeric(ext.get("Importe"), errors="coerce")
    ext = _aplicar_correcciones_ocr(ext)
    ext["blob"] = (
        ext.get("Descripcion", pd.Series("", index=ext.index)).fillna("").astype(str)
        + " "
        + ext.get("Detalle", pd.Series("", index=ext.index)).fillna("").astype(str)
    )
    ext["blob_norm"] = ext["blob"].map(_normalizar_texto)
    # débito efectivo
    deb = ext["Debito"].fillna(0)
    imp = ext["Importe"].fillna(0)
    ext["pago"] = deb.where(deb > 0.009, imp.where(imp < -0.009, 0).abs())
    ext = ext[ext["pago"] > 0.009].copy()
    # solo pagos/transf (no impuestos/comisiones genéricos) salvo que tenga CUIT de proveedor
    mask_pago = ext["blob"].map(es_debito_pago_o_transferencia)
    ext = ext[mask_pago].copy().reset_index(drop=True)
    return ext


def comprobantes_21101() -> pd.DataFrame:
    tango = pd.read_excel(TANGO, sheet_name=1)
    p211 = tango[tango["Cuenta"].astype(str).str.contains("21101", na=False)].copy()
    p211["Fecha"] = pd.to_datetime(p211["Fecha"], errors="coerce")
    p211["Debe"] = pd.to_numeric(p211["Debe"], errors="coerce").fillna(0)
    p211["Haber"] = pd.to_numeric(p211["Haber"], errors="coerce").fillna(0)
    p211["CUIT"] = p211["Cliente / Proveedor"].map(_cuit_digits)
    rows = []
    for _, r in p211.iterrows():
        tipo = str(r["Tipo comprobante"]).strip().upper()
        haber, debe = float(r["Haber"]), float(r["Debe"])
        if tipo in {"FCC", "NDC"}:
            importe = haber if haber > 0.009 else -debe
            signo = "A pagar"
        elif tipo in {"NCC", "NCV"}:
            importe = -debe if debe > 0.009 else haber
            signo = "Nota crédito (baja deuda)"
        else:
            importe = haber - debe
            signo = "Otro"
        rows.append(
            {
                "Proveedor": str(r["Razón social"]).strip(),
                "CUIT": r["CUIT"],
                "Fecha": r["Fecha"],
                "Tipo": tipo,
                "Nro comprobante": str(r["Numero Comprobante"]),
                "Importe": round(importe, 2),
                "Sentido": signo,
                "Descripción Tango": str(r.get("Descripción") or ""),
            }
        )
    return pd.DataFrame(rows)


NOMBRES_PILA = {
    "gustavo", "gabriel", "martin", "mariana", "carlos", "andres", "enrique",
    "miguel", "angel", "florencia", "nadia", "sonia", "mabel", "adriana",
    "valeria", "antonella", "silvana", "silvina", "noemi", "juan", "claudio",
    "ariel", "ruben", "dario", "ana", "karina", "agustin", "sebastian",
    "alfredo", "lopez", "garcia", "perez",
}


def tokens_proveedor(nombre: str, cuit: str) -> list[str]:
    """CUIT + alias + apellido/razón (no nombres de pila)."""
    toks: list[str] = []
    if cuit:
        toks.append(cuit)
    toks.extend(ALIAS_EXTRA.get(nombre, ()))
    n = _normalizar_texto(nombre)
    stop = {
        "sa", "srl", "sas", "soc", "de", "del", "la", "las", "los", "y",
        "sociedad", "anonima", "limitada", "trabajo", "cooperativa", "h",
    }
    words = [w for w in n.split() if w not in stop and len(w) >= 4]
    if words:
        apellido = words[0]
        if apellido not in NOMBRES_PILA:
            toks.append(apellido)
    seen: set[str] = set()
    out: list[str] = []
    for t in toks:
        t = _normalizar_texto(t)
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def asignar_pagos(ext: pd.DataFrame, comps: pd.DataFrame) -> pd.DataFrame:
    proveedores = (
        comps.groupby("Proveedor", as_index=False)
        .agg(CUIT=("CUIT", lambda s: next((x for x in s if x), "")))
    )
    # más específico primero (CUIT 11 dígitos)
    catalog = []
    for _, p in proveedores.iterrows():
        toks = tokens_proveedor(p["Proveedor"], p["CUIT"])
        catalog.append((p["Proveedor"], p["CUIT"], toks))

    usados: set[int] = set()
    filas = []
    # 1) CUIT exacto
    for nombre, cuit, toks in catalog:
        if not cuit:
            continue
        for i, row in ext.iterrows():
            if i in usados:
                continue
            blob = str(row["blob"])
            if cuit in re.sub(r"\D", "", blob) or cuit in row["blob_norm"]:
                usados.add(i)
                filas.append(_fila_pago(nombre, cuit, row, "CUIT en extracto"))
    # 2) alias / apellido fuerte
    for nombre, cuit, toks in catalog:
        fuertes = [
            t for t in toks
            if not t.isdigit() and len(t) >= 5 and t not in NOMBRES_PILA
        ]
        if not fuertes:
            continue
        for i, row in ext.iterrows():
            if i in usados:
                continue
            bn = row["blob_norm"]
            if any(t in bn for t in fuertes):
                usados.add(i)
                filas.append(_fila_pago(nombre, cuit, row, "Nombre en extracto"))
    return pd.DataFrame(filas)


def _fila_pago(nombre: str, cuit: str, row: pd.Series, criterio: str) -> dict:
    return {
        "Proveedor": nombre,
        "CUIT": cuit,
        "Fecha banco": row["Fecha"],
        "Descripción banco": str(row.get("Descripcion") or ""),
        "Detalle banco": str(row.get("Detalle") or ""),
        "Importe pago": round(float(row["pago"]), 2),
        "Comprobante banco": str(row.get("Comprobante") or ""),
        "Criterio": criterio,
    }


MES_FILL = PatternFill("solid", fgColor="D6E3F0")
MES_FONT = Font(name="Calibri", bold=True, size=11, color=COLOR_PRIMARIO)
TOTAL_FILL = PatternFill("solid", fgColor="1F4E79")
TOTAL_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")

MESES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def fmt_cuit(d: str) -> str:
    d = re.sub(r"\D", "", str(d or ""))
    if len(d) == 11:
        return f"{d[:2]}-{d[2:10]}-{d[10]}"
    return d


def sheet_name(nombre: str, usados: set[str]) -> str:
    raw = re.sub(r"[\[\]:*?/\\]", " ", nombre).strip()
    base = raw[:31] or "Proveedor"
    cand = base
    n = 2
    while cand.lower() in {x.lower() for x in usados}:
        suf = f" ({n})"
        cand = (base[: 31 - len(suf)] + suf)[:31]
        n += 1
    usados.add(cand)
    return cand


def etiqueta_mes(dt) -> str:
    if pd.isna(dt):
        return ""
    ts = pd.to_datetime(dt)
    return f"{MESES_ES[int(ts.month)]} {ts.year}"


def clave_mes(dt) -> str:
    if pd.isna(dt):
        return "9999-99"
    ts = pd.to_datetime(dt)
    return f"{ts.year:04d}-{ts.month:02d}"


def obs_proveedor(prov: str, saldo: float, tot_pagos: float, ncc_n: int, ndc_n: int) -> str:
    if "CIRUGIA" in prov.upper():
        return (
            "Mayor auxiliar: FCC de maquinaria en Haber; cuotas bancarias en Debe. "
            "NCC/NDC = diferencia de tipo de cambio (no son cuotas)."
        )
    if abs(saldo) < 1 and tot_pagos > 0:
        return "Saldo final $0: neto Tango = pagos banco."
    if tot_pagos == 0:
        return "Sin pagos identificados en esta cuenta (CC 067-0850315)."
    if tot_pagos > 0 and (ncc_n or ndc_n):
        return "Incluye NC/ND de Tango (ajustes) y pagos bancarios."
    if tot_pagos > 0 and abs(saldo) >= 1:
        return "Queda saldo (Haber - Debe) al cierre del período."
    return ""


def armar_asientos(comps: pd.DataFrame, pagos: pd.DataFrame) -> pd.DataFrame:
    """Mayor de proveedores: FCC/NDC = Haber; NCC y pagos banco = Debe."""
    rows: list[dict] = []
    for _, r in comps.iterrows():
        tipo = str(r["Tipo"]).upper()
        imp = float(r["Importe"] or 0)
        nro = str(r["Nro comprobante"])
        desc = str(r.get("Descripción Tango") or "")
        prov = str(r["Proveedor"])
        if tipo in {"NCC", "NCV"}:
            concepto = f"{tipo} {nro}"
            if "CIRUGIA" in prov.upper():
                concepto += " — ajuste tipo de cambio"
            elif desc:
                concepto += f" — {desc}"
            debe, haber = abs(imp), 0.0
        else:
            concepto = f"{tipo} {nro}"
            if desc:
                concepto += f" — {desc}"
            debe, haber = 0.0, abs(imp)
        rows.append(
            {
                "Fecha": pd.to_datetime(r["Fecha"]),
                "Proveedor": prov,
                "CUIT": r.get("CUIT") or "",
                "Concepto": concepto,
                "Comprobante": nro,
                "Origen": "Tango 21101",
                "Debe": round(debe, 2),
                "Haber": round(haber, 2),
            }
        )
    if pagos is not None and len(pagos):
        for _, r in pagos.iterrows():
            det = str(r.get("Detalle banco") or "").replace("\n", " ").strip()
            desc = str(r.get("Descripción banco") or "").strip()
            concepto = "Pago banco"
            if desc:
                concepto += f" — {desc}"
            rows.append(
                {
                    "Fecha": pd.to_datetime(r["Fecha banco"]),
                    "Proveedor": str(r["Proveedor"]),
                    "CUIT": r.get("CUIT") or "",
                    "Concepto": concepto[:90],
                    "Comprobante": str(r.get("Comprobante banco") or ""),
                    "Origen": "Banco Provincia",
                    "Debe": round(float(r["Importe pago"] or 0), 2),
                    "Haber": 0.0,
                    "Detalle": det[:80],
                }
            )
    if not rows:
        return pd.DataFrame(
            columns=[
                "Fecha",
                "Proveedor",
                "CUIT",
                "Concepto",
                "Comprobante",
                "Origen",
                "Debe",
                "Haber",
            ]
        )
    df = pd.DataFrame(rows)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Mes"] = df["Fecha"].map(etiqueta_mes)
    df["Mes clave"] = df["Fecha"].map(clave_mes)
    df = df.sort_values(["Fecha", "Proveedor", "Haber", "Debe"]).reset_index(drop=True)
    return df


def escribir_mayor(
    ws,
    asientos: pd.DataFrame,
    *,
    start_row: int = 8,
    incluir_proveedor: bool = False,
    saldo_inicial: float = 0.0,
) -> int:
    """Escribe mayor con corte mensual y saldo acumulado. Devuelve última fila."""
    cols = ["Fecha", "Concepto", "Comprobante", "Origen", "Debe", "Haber", "Saldo"]
    if incluir_proveedor:
        cols = ["Fecha", "Proveedor", "Concepto", "Comprobante", "Origen", "Debe", "Haber", "Saldo"]

    for c, h in enumerate(cols, 1):
        cell = ws.cell(start_row, c, h)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    if asientos is None or asientos.empty:
        ws.cell(start_row + 1, 1, "(sin movimientos)").font = SUB_FONT
        return start_row + 2

    work = asientos.sort_values(["Mes clave", "Fecha"]).copy()
    saldo = round(float(saldo_inicial), 2)
    r = start_row + 1
    mes_actual = None
    debe_mes = haber_mes = 0.0

    def _fila_corte(texto: str, d: float, h: float, s: float, fill, font) -> None:
        nonlocal r
        n = len(cols)
        for c in range(1, n + 1):
            ws.cell(r, c).fill = fill
            ws.cell(r, c).font = font
        ws.cell(r, 1, texto)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n - 3)
        ws.cell(r, n - 2, d).number_format = MONEY_FMT
        ws.cell(r, n - 1, h).number_format = MONEY_FMT
        ws.cell(r, n, s).number_format = MONEY_FMT
        ws.cell(r, n - 2).font = font
        ws.cell(r, n - 1).font = font
        ws.cell(r, n).font = font
        r += 1

    if abs(saldo) >= 0.009:
        _fila_corte("SALDO INICIAL", 0.0, 0.0, saldo, MES_FILL, MES_FONT)

    for _, row in work.iterrows():
        mes = row["Mes clave"]
        if mes_actual is not None and mes != mes_actual:
            _fila_corte(
                f"TOTAL {etiqueta_mes(pd.Timestamp(f'{mes_actual}-01'))}",
                debe_mes,
                haber_mes,
                saldo,
                MES_FILL,
                MES_FONT,
            )
            debe_mes = haber_mes = 0.0
        if mes != mes_actual:
            mes_actual = mes
            n = len(cols)
            ws.cell(r, 1, row["Mes"]).font = HDR_FONT
            ws.cell(r, 1).fill = HDR_FILL
            for c in range(1, n + 1):
                ws.cell(r, c).fill = HDR_FILL
                ws.cell(r, c).font = HDR_FONT
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n)
            r += 1

        debe = float(row["Debe"] or 0)
        haber = float(row["Haber"] or 0)
        saldo = round(saldo + haber - debe, 2)
        debe_mes += debe
        haber_mes += haber

        valores = {
            "Fecha": row["Fecha"].date() if hasattr(row["Fecha"], "date") else row["Fecha"],
            "Proveedor": row.get("Proveedor"),
            "Concepto": row["Concepto"],
            "Comprobante": str(row["Comprobante"]),
            "Origen": row["Origen"],
            "Debe": debe if debe else None,
            "Haber": haber if haber else None,
            "Saldo": saldo,
        }
        for c, h in enumerate(cols, 1):
            cell = ws.cell(r, c, valores[h])
            cell.font = BODY_FONT
            if h == "Fecha":
                cell.number_format = DATE_FMT
            elif h in {"Debe", "Haber", "Saldo"}:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
            elif h == "Comprobante":
                cell.number_format = "@"
            zebra_i = r % 2
            if zebra_i == 0 and h not in {"Debe", "Haber", "Saldo"}:
                cell.fill = ZEBRA
        r += 1

    if mes_actual is not None:
        _fila_corte(
            f"TOTAL {etiqueta_mes(pd.Timestamp(f'{mes_actual}-01'))}",
            debe_mes,
            haber_mes,
            saldo,
            MES_FILL,
            MES_FONT,
        )
    tot_d = float(work["Debe"].sum())
    tot_h = float(work["Haber"].sum())
    _fila_corte("SALDO FINAL DEL MAYOR", tot_d, tot_h, saldo, TOTAL_FILL, TOTAL_FONT)

    widths = {
        "Fecha": 12,
        "Proveedor": 32,
        "Concepto": 48,
        "Comprobante": 16,
        "Origen": 18,
        "Debe": 16,
        "Haber": 16,
        "Saldo": 16,
    }
    for c, h in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 14)
    return r


def encabezado_mayor(ws, titulo: str, sub: str, cuit: str = "", obs: str = "") -> None:
    ws.sheet_properties.tabColor = COLOR_PRIMARIO
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws["A1"] = titulo
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = sub
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:H2")
    linea = "Cuenta 21101 Proveedores  |  Debe = pagos y NCC  |  Haber = FCC y NDC  |  Saldo acreedor = deuda"
    if cuit:
        linea = f"CUIT {fmt_cuit(cuit)}  |  " + linea
    ws["A3"] = linea
    ws["A3"].font = SUB_FONT
    ws.merge_cells("A3:H3")
    if obs:
        ws["A4"] = obs
        ws["A4"].font = Font(name="Calibri", size=11, italic=True, color="1F4E79")
        ws["A4"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells("A4:H4")
        ws.row_dimensions[4].height = 28
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def main() -> None:
    print("Comprobantes 21101...")
    comps = comprobantes_21101()
    print("Extracto...")
    ext = cargar_extracto()
    print("Pagos...")
    pagos = asignar_pagos(ext, comps)
    asientos = armar_asientos(comps, pagos)

    resumen_rows = []
    usados: set[str] = set()
    hojas_map: dict[str, str] = {}
    for prov, g in comps.groupby("Proveedor"):
        pagos_p = pagos[pagos["Proveedor"] == prov] if len(pagos) else pd.DataFrame()
        a_p = asientos[asientos["Proveedor"] == prov]
        debe = float(a_p["Debe"].sum()) if len(a_p) else 0.0
        haber = float(a_p["Haber"].sum()) if len(a_p) else 0.0
        saldo = round(haber - debe, 2)
        hojas_map[prov] = sheet_name(prov, usados)
        resumen_rows.append(
            {
                "Proveedor": prov,
                "Hoja": hojas_map[prov],
                "CUIT": next((x for x in g["CUIT"] if str(x).strip()), ""),
                "Debe (pagos + NCC)": round(debe, 2),
                "Haber (FCC + NDC)": round(haber, 2),
                "Saldo acreedor": saldo,
                "Movimientos": len(a_p),
            }
        )
    resumen = pd.DataFrame(resumen_rows).sort_values("Haber (FCC + NDC)", ascending=False)

    mensual_rows = []
    for clave, gm in asientos.groupby("Mes clave", sort=True):
        debe = float(gm["Debe"].sum())
        haber = float(gm["Haber"].sum())
        mensual_rows.append(
            {
                "Mes": gm["Mes"].iloc[0],
                "Mes clave": clave,
                "Debe": round(debe, 2),
                "Haber": round(haber, 2),
                "Saldo del mes (H-D)": round(haber - debe, 2),
                "Movimientos": len(gm),
            }
        )
    mensual = pd.DataFrame(mensual_rows)
    if len(mensual):
        mensual["Saldo acumulado"] = mensual["Saldo del mes (H-D)"].cumsum().round(2)

    guardar_informe_excel(
        OUT,
        titulo="Gastro MDP — Mayor de proveedores",
        subtitulo=(
            "Mayor contable 21101 por mes: Debe = pagos banco y notas de crédito; "
            "Haber = facturas y notas de débito. Una hoja por mes y una por proveedor."
        ),
        periodo="sep-2025 a jul-2026",
        kpis=[
            ("Proveedores", len(resumen), "int"),
            ("Asientos", len(asientos), "int"),
            ("Debe", float(asientos["Debe"].sum()) if len(asientos) else 0.0, "money"),
            ("Haber", float(asientos["Haber"].sum()) if len(asientos) else 0.0, "money"),
        ],
        resumenes=[
            ("Cierre por proveedor", resumen),
            ("Cierre por mes", mensual.drop(columns=["Mes clave"]) if len(mensual) else mensual),
        ],
        detalle=None,
        col_moneda=[
            "Debe",
            "Haber",
            "Debe (pagos + NCC)",
            "Haber (FCC + NDC)",
            "Saldo acreedor",
            "Saldo del mes (H-D)",
            "Saldo acumulado",
        ],
        col_texto=["CUIT", "Hoja"],
    )

    wb = load_workbook(OUT)

    ws_g = wb.create_sheet("Mayor general", 1)
    encabezado_mayor(
        ws_g,
        "Mayor general de proveedores",
        "Todos los proveedores · cortes mensuales · saldo acreedor acumulado",
    )
    escribir_mayor(ws_g, asientos, incluir_proveedor=True)

    saldo_acum = 0.0
    for clave, gm in asientos.groupby("Mes clave", sort=True):
        nom = f"Mes {clave}"
        ws_m = wb.create_sheet(nom[:31])
        encabezado_mayor(
            ws_m,
            gm["Mes"].iloc[0],
            f"Mayor del mes {gm['Mes'].iloc[0]} — todos los proveedores (con saldo inicial)",
        )
        escribir_mayor(ws_m, gm, incluir_proveedor=True, saldo_inicial=saldo_acum)
        saldo_acum = round(
            saldo_acum + float(gm["Haber"].sum()) - float(gm["Debe"].sum()),
            2,
        )

    for prov in resumen["Proveedor"]:
        g = comps[comps["Proveedor"] == prov]
        pagos_p = pagos[pagos["Proveedor"] == prov] if len(pagos) else pd.DataFrame()
        a_p = asientos[asientos["Proveedor"] == prov]
        cuit = next((x for x in g["CUIT"] if str(x).strip()), "")
        debe = float(a_p["Debe"].sum()) if len(a_p) else 0.0
        haber = float(a_p["Haber"].sum()) if len(a_p) else 0.0
        ws_p = wb.create_sheet(hojas_map[prov])
        encabezado_mayor(
            ws_p,
            prov,
            "Mayor auxiliar de proveedores (por mes)",
            cuit=str(cuit),
            obs=obs_proveedor(
                prov,
                round(haber - debe, 2),
                float(pagos_p["Importe pago"].sum()) if len(pagos_p) else 0.0,
                int((g["Tipo"] == "NCC").sum()),
                int((g["Tipo"] == "NDC").sum()),
            ),
        )
        escribir_mayor(ws_p, a_p, incluir_proveedor=False)

    wb.save(OUT)
    print("OK", OUT)
    print("hojas", len(wb.sheetnames), "asientos", len(asientos))


if __name__ == "__main__":
    main()
