"""Prueba match débitos extracto 2026-03 vs Proveedores.xlsx."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from procesador import (
    _parsear_movimientos_santander_paginas,
    cargar_debitos_desde_extracto_df,
    cargar_facturas_proveedores_excel,
    exportar_match_proveedores_excel,
    matchear_debitos_con_facturas,
)

CACHE = Path("_tmp_santander_2026-03_ocr.json")
PROV = Path(r"T:\Estudio Contable\Proveedores.xlsx")
OUT = Path(r"C:\Users\recep\Desktop\Match_Debitos_Proveedores_2026-03.xlsx")


def main() -> None:
    pages = json.loads(CACHE.read_text(encoding="utf-8"))
    paginas = [(i + 1, "\n".join(p)) for i, p in enumerate(pages)]
    movs, meta = _parsear_movimientos_santander_paginas(paginas, "2026-03.pdf")
    df_ext = pd.DataFrame(movs)
    debitos = cargar_debitos_desde_extracto_df(df_ext)
    facturas = cargar_facturas_proveedores_excel(PROV)
    print("debitos", len(debitos), "total", round(debitos["importe"].sum(), 2))
    print("facturas", len(facturas), "total", round(facturas["importe"].sum(), 2))
    print("facturas rango", facturas["fecha"].min(), facturas["fecha"].max())

    # Filtrar facturas al mes del extracto (mar-2026) + margen, o todo el ejercicio
    res = matchear_debitos_con_facturas(debitos, facturas)
    print(res["resumen"].to_string(index=False))
    print("--- calzados sample ---")
    if len(res["calzados"]):
        print(res["calzados"].head(12).to_string(index=False))
    else:
        print("(ninguno)")
    xlsx = exportar_match_proveedores_excel(
        res,
        {
            "cliente": meta.get("cliente"),
            "origen_extracto": "2026-03.pdf (OCR)",
            "origen_facturas": str(PROV),
        },
    )
    OUT.write_bytes(xlsx)
    print("saved", OUT, len(xlsx))


if __name__ == "__main__":
    main()
