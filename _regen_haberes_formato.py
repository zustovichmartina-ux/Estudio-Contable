"""Regenera haberes.xlsx desde checkpoint con formato Estudio."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _extraer_haberes_rele import _parsear_lineas_haberes
from excel_formato_estudio import guardar_informe_excel

d = json.loads(Path("logs/rele_haberes_ocr_checkpoint.json").read_text(encoding="utf-8"))
items = []
for num_s, lineas in sorted(d["paginas"].items(), key=lambda x: int(x[0])):
    items.extend(_parsear_lineas_haberes(int(num_s), lineas))
df = pd.DataFrame(items).drop_duplicates(subset=["Fecha", "Comprobante", "Debito"], keep="first")
df["_f"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
df = df.sort_values(["_f", "Pagina"]).drop(columns=["_f"]).reset_index(drop=True)
df = df.rename(columns={"Debito": "Débito", "Pagina": "Página", "Movimiento": "Concepto"})
fechas = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
resumen = (
    df.assign(Mes=fechas.dt.to_period("M").astype(str))
    .groupby("Mes", dropna=False)["Débito"]
    .agg(Cantidad="count", Total="sum")
    .reset_index()
)
resumen["Total"] = resumen["Total"].round(2)
fmin, fmax = fechas.min(), fechas.max()
out = Path("logs/haberes_movimientos.xlsx")
guardar_informe_excel(
    out,
    titulo="Débitos de haberes",
    subtitulo="Extracto bancario · Pago haberes",
    periodo=f"Período: {fmin.strftime('%d/%m/%Y')} a {fmax.strftime('%d/%m/%Y')}",
    kpis=[
        ("Total débitos", float(df["Débito"].sum()), "money"),
        ("Cantidad de movimientos", len(df), "int"),
    ],
    resumenes=[("Por mes", resumen)],
    detalle=df,
    hoja_detalle="Movimientos",
    col_moneda=["Débito", "Total"],
    col_fecha=["Fecha"],
    total_col="Débito",
)
print(out, len(df), round(float(df["Débito"].sum()), 2))
