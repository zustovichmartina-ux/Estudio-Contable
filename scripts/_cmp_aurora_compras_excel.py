"""Compara carpeta Facturas Aurora vs Compras.csv y exporta Excel del estudio."""

from __future__ import annotations

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from excel_formato_estudio import guardar_informe_excel

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Cierre 2026\202607\Facturas"
)
CSV_PATH = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Impuestos\Portal IVA\07-2026\Compras.csv"
)
OUT = Path(r"C:\Users\recep\Desktop\Aurora_Facturas_vs_PortalIVA_07-2026.xlsx")

PAT = re.compile(
    r"(?i)^(FC(?:-REMITO)?|NC|REC)\s*(?:C\s+)?(\d+)\s*-\s*[A-Z]?(\d+)"
)


def norm_num(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"^0+", "", s) or "0"
    return s


def col(row: dict, *parts: str) -> str:
    for k, v in row.items():
        kl = k.lower()
        if all(p.lower() in kl for p in parts):
            return str(v).strip()
    return ""


def parse_money(s: str) -> float | None:
    t = str(s).strip().replace(" ", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def main() -> None:
    with open(CSV_PATH, "r", encoding="latin-1", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))

    csv_keys: dict[tuple[str, str], dict] = {}
    for r in rows:
        pv = col(r, "punto", "venta")
        nro = col(r, "numero", "comprobante") or col(r, "mero", "comprobante")
        key = (norm_num(pv), norm_num(nro))
        csv_keys[key] = {
            "Fecha": col(r, "fecha"),
            "Tipo comprobante": col(r, "tipo", "comprobante"),
            "Punto de venta": pv.zfill(5) if pv.isdigit() else pv,
            "Nro comprobante": nro.zfill(8) if nro.isdigit() else nro,
            "Proveedor": col(r, "denominaci"),
            "Importe total": parse_money(col(r, "importe", "total")),
            "Moneda": col(r, "moneda"),
        }

    by_key: dict[tuple[str, str], list] = defaultdict(list)
    skipped = []
    for p in sorted(FOLDER.iterdir()):
        if not p.is_file():
            continue
        name = p.name
        low = name.lower()
        if low.endswith((".xlsx", ".xls", ".jpeg", ".jpg", ".png", ".db")) or low.startswith(
            "whatsapp"
        ):
            skipped.append(name)
            continue
        m = PAT.search(name)
        if not m:
            skipped.append(name)
            continue
        tipo_f, pv, nro = m.group(1).upper(), m.group(2), m.group(3)
        key = (norm_num(pv), norm_num(nro))
        prov = name
        if " - " in name:
            prov = name.split(" - ", 1)[1].rsplit(".", 1)[0]
            prov = re.sub(r"\s*\(\d+\)\s*$", "", prov).strip()
        by_key[key].append(
            {
                "Tipo archivo": tipo_f,
                "Punto de venta": pv.zfill(5),
                "Nro comprobante": nro.zfill(8),
                "Proveedor (archivo)": prov,
                "Archivo": name,
            }
        )

    folder_keys = set(by_key)
    csv_only = []
    for key, v in sorted(
        csv_keys.items(),
        key=lambda x: (
            int(x[0][0]) if x[0][0].isdigit() else 0,
            int(x[0][1]) if x[0][1].isdigit() else 0,
        ),
    ):
        if key not in folder_keys:
            csv_only.append({**v, "Estado": "Falta PDF en carpeta"})

    carpeta_only = []
    for key, items in sorted(
        by_key.items(),
        key=lambda x: (
            int(x[0][0]) if x[0][0].isdigit() else 0,
            int(x[0][1]) if x[0][1].isdigit() else 0,
        ),
    ):
        if key not in csv_keys:
            for it in items:
                carpeta_only.append({**it, "Estado": "En carpeta, no en Portal IVA"})

    match_rows = []
    for key in sorted(
        folder_keys & set(csv_keys),
        key=lambda x: (
            int(x[0]) if x[0].isdigit() else 0,
            int(x[1]) if x[1].isdigit() else 0,
        ),
    ):
        v = csv_keys[key]
        files = " | ".join(i["Archivo"] for i in by_key[key])
        match_rows.append(
            {
                **v,
                "Archivo(s)": files,
                "Estado": "OK — match",
            }
        )

    dups = []
    for key, items in by_key.items():
        if len(items) > 1:
            for it in items:
                dups.append(it)

    df_faltan = pd.DataFrame(csv_only)
    df_extra = pd.DataFrame(carpeta_only)
    df_ok = pd.DataFrame(match_rows)
    df_dups = pd.DataFrame(dups) if dups else pd.DataFrame(
        columns=["Tipo archivo", "Punto de venta", "Nro comprobante", "Proveedor (archivo)", "Archivo"]
    )
    df_omit = pd.DataFrame({"Archivo omitido": skipped})

    total_faltan = float(df_faltan["Importe total"].sum()) if len(df_faltan) else 0.0

    guardar_informe_excel(
        OUT,
        titulo="Aurora de las Sierras SRL — Facturas vs Portal IVA",
        subtitulo="Cruce carpeta Cierre 2026/202607/Facturas vs Impuestos/Portal IVA/07-2026/Compras.csv",
        periodo="07/2026",
        kpis=[
            ("Comprobantes Portal IVA", len(csv_keys)),
            ("Comprobantes en carpeta", len(by_key)),
            ("Match OK", len(match_rows)),
            ("Faltan en carpeta", len(csv_only)),
            ("Solo en carpeta", len(carpeta_only)),
            ("Total $ faltantes (CSV)", total_faltan),
        ],
        resumenes=[
            (
                "Faltan en carpeta (por proveedor)",
                df_faltan.groupby("Proveedor", as_index=False)
                .agg(**{"Cantidad": ("Nro comprobante", "count"), "Importe total": ("Importe total", "sum")})
                .sort_values("Importe total", ascending=False)
                if len(df_faltan)
                else pd.DataFrame(columns=["Proveedor", "Cantidad", "Importe total"]),
            ),
        ],
        detalle=df_faltan if len(df_faltan) else pd.DataFrame(
            columns=[
                "Fecha",
                "Tipo comprobante",
                "Punto de venta",
                "Nro comprobante",
                "Proveedor",
                "Importe total",
                "Moneda",
                "Estado",
            ]
        ),
        hoja_detalle="Faltan en carpeta",
        hojas_adicionales=[
            ("Solo en carpeta", df_extra if len(df_extra) else pd.DataFrame(
                columns=[
                    "Tipo archivo",
                    "Punto de venta",
                    "Nro comprobante",
                    "Proveedor (archivo)",
                    "Archivo",
                    "Estado",
                ]
            )),
            ("Match OK", df_ok),
            ("Duplicados carpeta", df_dups),
            ("Omitidos", df_omit),
        ],
        col_moneda=["Importe total"],
        col_texto=[
            "Punto de venta",
            "Nro comprobante",
            "Tipo comprobante",
            "Tipo archivo",
        ],
        total_col="Importe total",
    )
    print(OUT)


if __name__ == "__main__":
    main()
