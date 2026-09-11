# -*- coding: utf-8 -*-
import csv
import re
from collections import defaultdict
from pathlib import Path

csv_path = Path(r"C:\Users\recep\Desktop\RELE.csv")
folder = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026\07-2026"
)

TIPOS = {
    "1": "Factura A",
    "2": "ND A",
    "3": "NC A",
    "6": "Factura B",
    "11": "Factura C",
    "15": "Recibo C",
    "39": "Otros A",
    "63": "Liquidación",
    "81": "Tique Factura A",
}

arca = []
with csv_path.open(encoding="latin-1") as f:
    for row in csv.DictReader(f, delimiter=";"):
        pv = int(str(row["Punto de Venta"]).replace(".", ""))
        nro = int(str(row["Número de Comprobante"]).replace(".", ""))
        tipo = str(row["Tipo de Comprobante"]).strip()
        arca.append(
            {
                "fecha": row["Fecha de Emisión"],
                "tipo": tipo,
                "tipo_txt": TIPOS.get(tipo, tipo),
                "pv": pv,
                "nro": nro,
                "prov": row["Denominación Vendedor"].strip(),
                "total": row["Importe Total"],
                "key": (pv, nro),
            }
        )

pat = re.compile(
    r"^(FCC|NCC|NDC)([ABC])(?:(\d+)-)?(\d+)\s+(.+)\.(pdf|jpeg|jpg)$",
    re.I,
)
uniq = {}
for p in folder.iterdir():
    if p.suffix.lower() not in {".pdf", ".jpeg", ".jpg"}:
        continue
    m = pat.match(p.name)
    if not m:
        continue
    tipo, letra, pv, nro, prov, _ext = m.groups()
    pv_i = int(pv) if pv else None
    nro_i = int(nro)
    prov = re.sub(r"\s+HOJA\s+\d+", "", prov, flags=re.I).strip()
    k2 = (pv_i, nro_i)
    uniq.setdefault(k2, {"tipo": tipo, "letra": letra, "pv": pv_i, "nro": nro_i, "prov": prov, "file": p.name})

arca_by_nro = defaultdict(list)
for a in arca:
    arca_by_nro[a["nro"]].append(a["key"])

matched_arca = set()
matched_folder = set()

for k, fx in uniq.items():
    pv, nro = k
    if pv is not None and (pv, nro) in {a["key"] for a in arca}:
        matched_folder.add(k)
        matched_arca.add((pv, nro))
        continue
    cands = arca_by_nro.get(nro, [])
    if pv is None and cands:
        matched_folder.add(k)
        matched_arca.add(cands[0])
        continue

falta = [a for a in arca if a["key"] not in matched_arca]
extra = [uniq[k] for k in uniq if k not in matched_folder]

print(f"CSV ARCA: {len(arca)}  |  Carpeta (únicos): {len(uniq)}  |  Matcheados: {len(matched_arca)}")
print()
print("=== EN EL CSV / EXCEL Y NO TENEMOS EN LA CARPETA ===")
for a in falta:
    print(
        f"{a['fecha']}  {a['tipo_txt']:18}  {a['pv']}-{a['nro']}  {a['prov']}  total {a['total']}"
    )
print(f"Total faltantes: {len(falta)}")
print()
print("=== TENEMOS EN LA CARPETA Y NO ESTAN EN EL CSV ===")
for x in extra:
    pv = x["pv"] if x["pv"] is not None else "?"
    print(f"{x['tipo']}{x['letra']}{pv}-{x['nro']}  {x['prov']}")
print(f"Total de más: {len(extra)}")

import sys
from datetime import datetime

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from excel_formato_estudio import guardar_informe_excel


def _fecha(txt: str):
    try:
        return datetime.strptime(str(txt).strip(), "%d/%m/%Y").date()
    except Exception:
        return txt


def _money(txt) -> float:
    s = str(txt or "0").strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


df_falta = pd.DataFrame(
    [
        {
            "Fecha": _fecha(a["fecha"]),
            "Tipo": a["tipo_txt"],
            "Punto de venta": a["pv"],
            "Número": a["nro"],
            "Proveedor": a["prov"],
            "Importe": _money(a["total"]),
        }
        for a in falta
    ]
)
df_extra = pd.DataFrame(
    [
        {
            "Comprobante": f"{x['tipo']}{x['letra']}{x['pv'] or ''}-{x['nro']}".replace("--", "-"),
            "Proveedor": x["prov"],
            "Archivo": x["file"],
        }
        for x in extra
    ]
)
arca_map = {a["key"]: a for a in arca}
df_ok = pd.DataFrame(
    [
        {
            "Fecha": _fecha(arca_map[k]["fecha"]),
            "Tipo ARCA": arca_map[k]["tipo_txt"],
            "PV-Nro": f"{arca_map[k]['pv']}-{arca_map[k]['nro']}",
            "Proveedor ARCA": arca_map[k]["prov"],
            "Archivo carpeta": next(
                (
                    uniq[fk]["file"]
                    for fk in matched_folder
                    if fk[1] == k[1] and (fk[0] == k[0] or fk[0] is None)
                ),
                "",
            ),
            "Importe": _money(arca_map[k]["total"]),
        }
        for k in sorted(matched_arca, key=lambda x: (x[0], x[1]))
    ]
)

out = Path(r"C:\Users\recep\Desktop\Cruce_RELE_ARCA_vs_carpeta_07-2026.xlsx")
guardar_informe_excel(
    out,
    titulo="Cruce ARCA vs carpeta de facturas — Rele",
    subtitulo="Oftalmología Rele Mar del Plata SRL",
    periodo="07/2026",
    kpis=[
        ("Comprobantes ARCA", len(arca), "int"),
        ("Únicos en carpeta", len(uniq), "int"),
        ("Matcheados", len(matched_arca), "int"),
        ("En ARCA y no en carpeta", len(falta), "int"),
        ("En carpeta y no en ARCA", len(extra), "int"),
    ],
    resumenes=[
        ("En ARCA y no tenemos", df_falta),
        ("Tenemos y no están en ARCA", df_extra),
    ],
    detalle=df_ok,
    hoja_detalle="Matcheados",
    col_moneda=["Importe"],
    col_fecha=["Fecha"],
    total_col="Importe",
)
print("EXCEL", out)
