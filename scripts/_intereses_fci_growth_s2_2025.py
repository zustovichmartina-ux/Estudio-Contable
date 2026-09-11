# -*- coding: utf-8 -*-
"""Growth FCI: saldo 30/06/2025 + intereses FIFO jul-dic 2025 (fórmula explícita)."""
from __future__ import annotations

import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_formato_estudio import construir_informe_excel
from extraccion_fci import extraer_texto_pdf, parsear_pdf_fci
from inversiones import aplicar_fifo, identificar_movimientos

INV = Path(
    r"T:\CLIENTES\GROWTH MARKETING SA\Cierre 2025 01-07-2025 al 30-06-2026"
    r"\Papeles de Trabajo\Inversiones"
)
OUT_T = INV / "Intereses_FCI_Growth_jul_dic_2025.xlsx"
OUT_DESK = Path(r"C:\Users\recep\Desktop\Intereses_FCI_Growth_jul_dic_2025.xlsx")

FONDO = "FIMA PREMIUM CLASE A"
CIERRE = date(2025, 6, 30)
DESDE = date(2025, 7, 1)
HASTA = date(2025, 12, 31)

RE_POS_FECHA = re.compile(r"Posicion al\s+(?P<fecha>\d{2}/\d{2}/\d{4})", re.I)
RE_POS_LINEA = re.compile(
    r"(?P<fondo>FIMA PREMIUM CLASE A)\s+"
    r"(?P<cuotas>[\d\s.]+,\s*[\d\s]+)\s+\$\s*"
    r"(?P<vc>[\d.]+,\d+)\s+\$\s*"
    r"(?P<saldo>[\d.]+,\d{2})",
    re.I,
)
MESES = {
    7: "2025-07 Julio",
    8: "2025-08 Agosto",
    9: "2025-09 Septiembre",
    10: "2025-10 Octubre",
    11: "2025-11 Noviembre",
    12: "2025-12 Diciembre",
}


def _num_ar(raw: str) -> float:
    s = re.sub(r"\s+", "", str(raw or "")).replace("$", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def extraer_posicion(pdf: Path) -> dict:
    texto = extraer_texto_pdf(pdf.read_bytes())
    m_f = RE_POS_FECHA.search(texto)
    m = RE_POS_LINEA.search(texto)
    if not m_f or not m:
        raise ValueError(f"No se leyó posición en {pdf.name}")
    fondo = re.sub(r"\s+", " ", m.group("fondo")).strip()
    return {
        "Fecha": datetime.strptime(m_f.group("fecha"), "%d/%m/%Y").date(),
        "Fondo": fondo,
        "Cantidad": _num_ar(m.group("cuotas")),
        "Valor_CC": _num_ar(m.group("vc")),
        "Saldo": _num_ar(m.group("saldo")),
        "Archivo": pdf.name,
    }


def pdfs_semestre() -> list[Path]:
    out = []
    for p in sorted(INV.glob("Extracto_Inversiones_Galicia_2025_*.pdf")):
        # 06 = solo posición; 07-12 = movimientos del semestre
        if any(f"2025_{mm}" in p.name for mm in ("07", "08", "09", "10", "11", "12")):
            out.append(p)
    return out


def movimientos_jul_dic() -> pd.DataFrame:
    filas = []
    for p in pdfs_semestre():
        res = parsear_pdf_fci(p.read_bytes(), p.name)
        for f in res.get("filas") or []:
            fe = f.get("Fecha")
            if not isinstance(fe, date):
                continue
            if fe < DESDE or fe > HASTA:
                continue
            filas.append(f)
    df = pd.DataFrame(filas)
    if df.empty:
        return df
    df["_fecha"] = df["Fecha"]
    df["Especie"] = df.get("Fondo", FONDO)
    df["Tipo_Operacion"] = df["Descripcion"].map(
        lambda x: "Rescate" if "RESCATE" in str(x).upper() else "Suscripcion"
    )
    df["Cantidad"] = df["Cantidad de Cuotas"]
    df["Precio"] = df["Valor CC"]
    df["Monto_Total"] = df["Total"]
    df["Moneda"] = "ARS"
    df["Nueva_Clasificacion"] = None
    df = df.sort_values(["_fecha"], kind="stable").reset_index(drop=True)
    return identificar_movimientos(df)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    jun = INV / "Extracto_Inversiones_Galicia_2025_06_30.pdf"
    pos = extraer_posicion(jun)
    print("POSICION", pos)
    dic = extraer_posicion(INV / "Extracto_Inversiones_Galicia_2025_12_30.pdf")
    print("CIERRE DIC", dic)

    df_mov = movimientos_jul_dic()
    print("movimientos", len(df_mov))
    print(df_mov.groupby(df_mov["_fecha"].map(lambda d: d.strftime("%Y-%m")))["Tipo_Operacion"].value_counts().to_string())

    df_ini = pd.DataFrame([{
        "Especie": FONDO,
        "Especie_canonica": FONDO,
        "Tipo_inversion": "fci",
        "Grupo": "FCI",
        "Cantidad": pos["Cantidad"],
        "Costo_Unitario": pos["Valor_CC"],
        "Costo_Total": pos["Saldo"],
        "Moneda": "ARS",
        "Origen": f"Posición al {pos['Fecha'].strftime('%d/%m/%Y')}",
        "Fecha": pos["Fecha"].strftime("%d/%m/%Y"),
    }])
    res = aplicar_fifo(df_mov, df_ini)
    print("avisos", res.avisos)
    print("saldos", res.saldos)

    apps = []
    for a in res.aplicaciones:
        fe = datetime.strptime(a["Fecha_salida"], "%d/%m/%Y").date()
        lote = a.get("Fecha_lote") or ""
        if lote in {"Saldo inicial", "INI"}:
            lote = "30/06/2025"
        vc_out = float(a["Valor_CC_salida"])
        vc_lot = float(a["Valor_CC_lote"])
        qty = float(a["Cantidad"])
        apps.append({
            "Fecha": fe,
            "Mes": MESES[fe.month],
            "Fondo": a["Especie"],
            "Tipo": "RESCATE",
            "Cuotas de este lote": qty,
            "Valor CC rescate": vc_out,
            "Valor CC lote": vc_lot,
            "Fecha lote": lote,
            "Interes": round((vc_out - vc_lot) * qty, 2),
            "Formula": f"=({vc_out:.6f}-{vc_lot:.6f})*{qty:.2f}",
        })
    df_int = pd.DataFrame(apps)

    por_mes = (
        df_int.groupby("Mes", as_index=False)
        .agg(Rescates_partidos=("Interes", "size"), Interes=("Interes", "sum"))
        .sort_values("Mes")
    )
    print(por_mes.to_string(index=False))
    print("TOTAL INTERES", round(float(df_int["Interes"].sum()), 2))

    stock = pd.DataFrame(res.saldos)
    if not stock.empty:
        stock["Posicion mercado 31/12"] = dic["Cantidad"]
        stock["Valor CC 31/12"] = dic["Valor_CC"]
        stock["Saldo mercado 31/12"] = dic["Saldo"]
        stock["Dif cuotas vs extracto"] = stock["Cantidad"] - dic["Cantidad"]

    mov_out = df_mov[[
        "Fecha", "Descripcion", "Cantidad de Cuotas", "Valor CC", "Total", "Fondo", "Archivo",
    ]].copy() if not df_mov.empty else pd.DataFrame()

    kpis = [
        ("Saldo inicial 30/06/2025 (cuotas)", pos["Cantidad"]),
        ("Valor cuota 30/06/2025", pos["Valor_CC"]),
        ("Valorizado 30/06/2025", pos["Saldo"]),
        ("Intereses jul-dic 2025", float(df_int["Interes"].sum()) if not df_int.empty else 0.0),
        ("Cuotas FIFO al 31/12", float(stock["Cantidad"].sum()) if not stock.empty else 0.0),
        ("Cuotas extracto 31/12", dic["Cantidad"]),
    ]

    wb = construir_informe_excel(
        titulo="Intereses FCI — GROWTH MARKETING SA",
        subtitulo="FIMA PREMIUM CLASE A — FIFO: (VC rescate − VC lote) × cuotas de ese lote",
        periodo="Ejercicio 01/07/2025 al 30/06/2026 · intereses 07 a 12/2025",
        kpis=kpis,
        resumenes=[
            ("Saldo inicial = posición al 30/06/2025", pd.DataFrame([{
                "Fecha": pos["Fecha"],
                "Fondo": pos["Fondo"],
                "Cuotas": pos["Cantidad"],
                "Valor CC": pos["Valor_CC"],
                "Saldo valorizado": pos["Saldo"],
                "Origen": pos["Archivo"],
            }])),
            ("Intereses por mes", por_mes),
        ],
        detalle=df_int,
        hoja_detalle="Intereses",
        hojas_adicionales=[
            ("Movimientos", mov_out),
            ("Saldos_3112", stock),
            ("Avisos", pd.DataFrame({"Aviso": res.avisos or ["Sin avisos"]})),
        ],
        col_moneda=["Interes", "Valorizado 30/06/2025", "Saldo valorizado", "Total", "Saldo mercado 31/12", "Costo_Total"],
        col_fecha=["Fecha", "Fecha lote"],
        total_col="Interes",
    )

    # Fórmula Excel real en Interés: =(Valor CC rescate − Valor CC lote) × cuotas
    ws = wb["Intereses"]
    headers = {str(c.value).strip(): c.column for c in ws[1] if c.value}
    c_int = headers["Interes"]
    c_qty = headers["Cuotas de este lote"]
    c_out = headers["Valor CC rescate"]
    c_lot = headers["Valor CC lote"]
    last = ws.max_row
    if ws.cell(last, 1).value == "TOTAL":
        last -= 1
    for r in range(2, last + 1):
        ws.cell(r, c_int).value = (
            f"={get_column_letter(c_out)}{r}-{get_column_letter(c_lot)}{r}"
            f"*{get_column_letter(c_qty)}{r}"
        )
        ws.cell(r, c_int).number_format = '#,##0.00;[Red]-#,##0.00'

    for dest in (OUT_T, OUT_DESK):
        wb.save(dest)
        print("OK", dest)


if __name__ == "__main__":
    main()
