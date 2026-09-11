# -*- coding: utf-8 -*-
"""Percepciones IVA/IIBB Rele ene-jun 2026 + listado REVISAR."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE.parent))
sys.path.insert(0, str(_HERE))

from excel_formato_estudio import guardar_informe_excel
from _percepciones_rele_julio import datos_archivo, extraer_percepciones, texto_archivo
from _renombrar_fcts_rele_2026 import ROOT, SKIP_MESES

OK_NAME = re.compile(r"^(FCC|NCC|NDC)[ABC]\d+", re.I)
OUT = Path(r"C:\Users\recep\Desktop\Percepciones_IVA_IIBB_Rele_01a06-2026.xlsx")
OUT_REV = Path(r"C:\Users\recep\Desktop\Revisar_facturas_Rele_01a06-2026.xlsx")


def main() -> None:
    filas: list[dict] = []
    revisar: list[dict] = []
    meses = sorted(
        p for p in ROOT.iterdir() if p.is_dir() and p.name not in SKIP_MESES
    )
    for mes in meses:
        for p in sorted(mes.iterdir(), key=lambda x: x.name.lower()):
            if not p.is_file():
                continue
            if p.suffix.lower() not in {".pdf", ".jpeg", ".jpg", ".png"}:
                continue
            if p.name.startswith("Percepciones") or p.name.lower() == "thumbs.db":
                continue
            incompleto = (
                "PROVEEDOR" in p.name.upper()
                or not OK_NAME.match(p.name)
                or re.search(r"adobe scan|materiales|liquidacion|resumen|retenciones", p.name, re.I)
            )
            if incompleto:
                revisar.append(
                    {
                        "Mes": mes.name,
                        "Archivo": p.name,
                        "Tipo": p.suffix.lower(),
                    }
                )
            if p.suffix.lower() != ".pdf":
                continue
            texto = texto_archivo(p)
            percs = extraer_percepciones(texto)
            meta = datos_archivo(p.name)
            for perc in percs:
                filas.append({"Mes": mes.name, **meta, "Archivo": p.name, **perc})

    df = pd.DataFrame(filas)
    if df.empty:
        df = pd.DataFrame(
            columns=[
                "Mes",
                "Comprobante",
                "Proveedor",
                "Tipo percepción",
                "Jurisdicción",
                "Importe",
                "Archivo",
            ]
        )
    por_tipo = (
        df.groupby(["Mes", "Tipo percepción", "Jurisdicción"], dropna=False)["Importe"]
        .sum()
        .reset_index()
        if not df.empty
        else pd.DataFrame()
    )
    por_fct = (
        df.groupby(["Mes", "Comprobante", "Proveedor"], dropna=False)
        .agg(
            Cant_percepciones=("Importe", "size"),
            Total_percepciones=("Importe", "sum"),
        )
        .reset_index()
        if not df.empty
        else pd.DataFrame()
    )
    kpis = [
        ("Comprobantes con percepción", int(df["Comprobante"].nunique()) if not df.empty else 0, "int"),
        ("Líneas de percepción", len(df), "int"),
        ("Total percepciones", float(df["Importe"].sum()) if not df.empty else 0.0, "money"),
        ("Archivos a revisar", len(revisar), "int"),
    ]
    guardar_informe_excel(
        OUT,
        titulo="Percepciones IVA e IIBB — Rele",
        subtitulo="Oftalmología Rele Mar del Plata SRL — compras ene-jun 2026",
        periodo="01 a 06/2026",
        kpis=kpis,
        resumenes=[
            ("Por mes, tipo y jurisdicción", por_tipo),
            ("Por comprobante", por_fct),
        ],
        detalle=df,
        hoja_detalle="Detalle percepciones",
        col_moneda=["Importe", "Total_percepciones"],
        total_col="Importe",
    )
    df_rev = pd.DataFrame(revisar)
    guardar_informe_excel(
        OUT_REV,
        titulo="Facturas Rele a revisar",
        subtitulo="Fotos WhatsApp / liquidaciones / sin PV-número",
        periodo="01 a 06/2026",
        kpis=[("Archivos", len(revisar), "int")],
        resumenes=[("Por mes", df_rev.groupby("Mes").size().reset_index(name="Cantidad") if not df_rev.empty else pd.DataFrame())],
        detalle=df_rev,
        hoja_detalle="Revisar",
    )
    print("PERC", OUT, "filas", len(df))
    if not df.empty:
        print(df.groupby("Mes")["Importe"].sum().to_string())
        print(df[["Mes", "Comprobante", "Proveedor", "Tipo percepción", "Jurisdicción", "Importe"]].to_string(index=False))
    print("REVISAR", OUT_REV, len(revisar))


if __name__ == "__main__":
    main()
