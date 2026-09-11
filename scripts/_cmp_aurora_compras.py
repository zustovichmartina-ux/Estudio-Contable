"""Compara facturas carpeta Aurora vs Compras.csv Portal IVA."""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from pathlib import Path

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Cierre 2026\202607\Facturas"
)
CSV_PATH = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Impuestos\Portal IVA\07-2026\Compras.csv"
)

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


def main() -> None:
    with open(CSV_PATH, "r", encoding="latin-1", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))

    csv_keys: dict[tuple[str, str], dict] = {}
    for r in rows:
        pv = col(r, "punto", "venta")
        nro = col(r, "numero", "comprobante")
        if not nro:
            nro = col(r, "mero", "comprobante")
        key = (norm_num(pv), norm_num(nro))
        csv_keys[key] = {
            "tipo": col(r, "tipo", "comprobante"),
            "fecha": col(r, "fecha"),
            "den": col(r, "denominaci"),
            "total": col(r, "importe", "total"),
            "pv": pv,
            "nro": nro,
        }

    folder_items = []
    skipped = []
    for p in sorted(FOLDER.iterdir()):
        if not p.is_file():
            continue
        name = p.name
        low = name.lower()
        if low.endswith((".xlsx", ".xls", ".jpeg", ".jpg", ".png")) or low.startswith(
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
        folder_items.append(
            {
                "file": name,
                "tipo": tipo_f,
                "key": key,
                "pv": pv,
                "nro": nro,
                "prov": prov,
            }
        )

    by_key: dict[tuple[str, str], list] = defaultdict(list)
    for it in folder_items:
        by_key[it["key"]].append(it)

    faltan_csv = []
    for key, items in sorted(
        by_key.items(),
        key=lambda x: (
            int(x[0][0]) if x[0][0].isdigit() else 0,
            int(x[0][1]) if x[0][1].isdigit() else 0,
        ),
    ):
        if key not in csv_keys:
            faltan_csv.append((key, items))

    faltan_carpeta = []
    folder_keys = set(by_key)
    for key, v in sorted(
        csv_keys.items(),
        key=lambda x: (
            int(x[0][0]) if x[0][0].isdigit() else 0,
            int(x[0][1]) if x[0][1].isdigit() else 0,
        ),
    ):
        if key not in folder_keys:
            faltan_carpeta.append((key, v))

    print("=" * 70)
    print("EN CARPETA pero NO estan en Compras.csv (Portal IVA)")
    print("=" * 70)
    if not faltan_csv:
        print("(ninguna)")
    for key, items in faltan_csv:
        it = items[0]
        files = " | ".join(i["file"] for i in items)
        print(
            f"- {it['tipo']:10} PV {it['pv']:>6}-{it['nro']:<12} {it['prov']}"
        )
        print(f"  archivo: {files}")

    print()
    print("=" * 70)
    print("EN Compras.csv pero NO estan en la carpeta Facturas")
    print("=" * 70)
    if not faltan_carpeta:
        print("(ninguna)")
    for key, v in faltan_carpeta:
        print(
            f"- {v['fecha']}  Tipo {v['tipo']:>2}  "
            f"PV {v['pv']:>6}-{v['nro']:<12}  {v['den']}  Total {v['total']}"
        )

    print()
    print("=" * 70)
    print("RESUMEN")
    print("=" * 70)
    print(f"Filas CSV Portal IVA:           {len(csv_keys)}")
    print(f"Comprobantes unicos en carpeta: {len(by_key)}")
    print(f"Archivos parseados:             {len(folder_items)}")
    print(f"Omitidos (xlsx/jpg/etc):        {len(skipped)}")
    for s in skipped:
        print(f"  - {s}")
    print(f"Match (misma PV-Nro):           {len(folder_keys & set(csv_keys))}")
    print(f"Faltan en CSV (hay PDF):        {len(faltan_csv)}")
    print(f"Faltan en carpeta (hay CSV):    {len(faltan_carpeta)}")

    # Duplicados carpeta
    dups = {k: v for k, v in by_key.items() if len(v) > 1}
    if dups:
        print()
        print("Duplicados en carpeta (misma PV-Nro, varios archivos):")
        for k, items in dups.items():
            print(f"  {k}:")
            for i in items:
                print(f"    - {i['file']}")


if __name__ == "__main__":
    main()
