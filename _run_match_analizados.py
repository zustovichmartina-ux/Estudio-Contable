"""Match débitos de extractos ya analizados vs proveedores desde Tango SQL."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from procesador import (
    _parsear_movimientos_santander_paginas,
    cargar_debitos_desde_extracto_df,
    cargar_facturas_proveedores_tango_sql,
    exportar_match_proveedores_excel,
    matchear_debitos_con_facturas,
)

CACHE_OCR = Path("_tmp_santander_2026-03_ocr.json")
EXCELS = [
    Path(r"C:\Users\recep\Downloads\Extracto_Santander_unificado_3071802274_20260710.xlsx"),
    Path(r"C:\Users\recep\Downloads\Extracto_Santander_2026-03_OCR.xlsx"),
    Path(r"C:\Users\recep\Desktop\Extracto_Santander_Oftalmologia_RELE_2026-03.xlsx"),
]
OUTS = [
    Path(r"C:\Users\recep\Desktop\Match_Debitos_Proveedores_RELE.xlsx"),
    Path(r"C:\Users\recep\Downloads\Match_Debitos_Proveedores_RELE.xlsx"),
    Path(r"T:\Estudio Contable\Match_Debitos_Proveedores_RELE.xlsx"),
]


def _debitos_desde_excel(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Movimientos")
    except Exception:
        df = pd.read_excel(path)
    deb = cargar_debitos_desde_extracto_df(df)
    if not deb.empty:
        deb["Archivo origen"] = path.name
    return deb


def main() -> None:
    partes: list[pd.DataFrame] = []
    origenes: list[str] = []

    if CACHE_OCR.exists():
        pages = json.loads(CACHE_OCR.read_text(encoding="utf-8"))
        paginas = [(i + 1, "\n".join(p)) for i, p in enumerate(pages)]
        movs, _meta = _parsear_movimientos_santander_paginas(paginas, "2026-03.pdf")
        deb = cargar_debitos_desde_extracto_df(pd.DataFrame(movs))
        if not deb.empty:
            deb["Archivo origen"] = "2026-03.pdf (OCR)"
            partes.append(deb)
            origenes.append("2026-03.pdf OCR")
            print(f"OCR 2026-03: {len(deb)} débitos", flush=True)

    vistos_excel: set[str] = set()
    for path in EXCELS:
        if not path.exists():
            continue
        key = f"{path.stat().st_size}"
        if key in vistos_excel and "unificado" not in path.name.lower():
            continue
        deb = _debitos_desde_excel(path)
        if deb.empty:
            print(f"Sin débitos: {path.name}", flush=True)
            continue
        vistos_excel.add(key)
        partes.append(deb)
        origenes.append(path.name)
        print(f"Excel {path.name}: {len(deb)} débitos", flush=True)

    if not partes:
        raise SystemExit("No hay débitos para matchear.")

    debitos = pd.concat(partes, ignore_index=True)
    debitos["_k"] = (
        debitos["fecha"].astype(str)
        + "|"
        + debitos["importe"].round(2).astype(str)
        + "|"
        + debitos["descripcion_norm"].str[:60]
    )
    antes = len(debitos)
    debitos = debitos.drop_duplicates(subset=["_k"], keep="first").drop(columns=["_k"])
    debitos = debitos.reset_index(drop=True)
    debitos["debito_id"] = debitos.index.astype(int)
    print(f"Debitos consolidados: {antes} -> {len(debitos)} (sin dup)", flush=True)

    facturas = cargar_facturas_proveedores_tango_sql(
        nombre_empresa="OFTALMOLOGIA RELE MAR DEL PLATA S.R.L.",
        cuit="30-71802274-2",
    )
    origen_prov = facturas.attrs.get("origen") or "Tango SQL"
    print(f"Facturas SQL: {len(facturas)} · {origen_prov}", flush=True)

    res = matchear_debitos_con_facturas(debitos, facturas)
    print(res["resumen"].to_string(index=False), flush=True)

    xlsx = exportar_match_proveedores_excel(
        res,
        {
            "cliente": "OFTALMOLOGIA RELE MAR DEL PLATA S.R.L.",
            "cuit": "30-71802274-2",
            "origen_extracto": " | ".join(origenes),
            "origen_facturas": origen_prov,
        },
    )
    for out in OUTS:
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(xlsx)
            print(f"OK {out}", flush=True)
        except Exception as exc:
            print(f"FAIL {out}: {exc}", flush=True)


if __name__ == "__main__":
    main()
