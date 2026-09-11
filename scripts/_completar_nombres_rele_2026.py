# -*- coding: utf-8 -*-
"""Segunda pasada: PDFs/OCR sobre archivos incompletos (ene-jun 2026)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_HERE))
from cruce_facturas_arca import extraer_texto_comprobante

from _renombrar_fcts_rele_2026 import (
    ROOT,
    SKIP_MESES,
    SKIP_NAMES,
    _desde_nombre,
    _desde_texto,
    armar_nombre,
    unique_target,
)

INCOMPLETO = re.compile(
    r"PROVEEDOR|adobe scan|materiales",
    re.I,
)
SIN_NRO = re.compile(r"^(FCC|NCC|NDC)[ABC](?!\d)", re.I)
CUIT_COMO_PV = re.compile(
    r"^(FCC|NCC|NDC)[ABC](20|23|27|30|33)-(\d{8,})\b",
    re.I,
)


def es_incompleto(nombre: str) -> bool:
    if SKIP_NAMES.search(nombre):
        return False
    if INCOMPLETO.search(nombre):
        return True
    if SIN_NRO.match(nombre):
        return True
    if CUIT_COMO_PV.match(nombre) and "PROVEEDOR" in nombre.upper():
        return True
    if re.search(r"CENTRO MEDICO", nombre, re.I) and not re.search(
        r"(FCC|NCC|NDC)[ABC]\d+-\d+", nombre, re.I
    ):
        return True
    return False


def completar(path: Path) -> dict:
    info = _desde_nombre(path.name)
    data = path.read_bytes()
    texto, metodo = extraer_texto_comprobante(path.name, data, usar_ocr=True)
    extra = _desde_texto(texto or "")
    for k, v in extra.items():
        if not info.get(k):
            info[k] = v
    nuevo = armar_nombre(info, path.name)
    return {
        "orig": path.name,
        "nuevo": nuevo,
        "info": info,
        "metodo": metodo,
        "completo": bool(info.get("nro") and info.get("prov")),
        "texto_len": len(texto or ""),
    }


def main() -> None:
    meses = sorted(
        p for p in ROOT.iterdir() if p.is_dir() and p.name not in SKIP_MESES
    )
    for mes in meses:
        files = [
            p
            for p in mes.iterdir()
            if p.is_file()
            and p.suffix.lower() in {".pdf", ".jpeg", ".jpg", ".png"}
            and es_incompleto(p.name)
        ]
        print(f"==== {mes.name} incompletos={len(files)}")
        used = {p.name.lower() for p in mes.iterdir() if p.is_file()}
        for p in sorted(files, key=lambda x: x.name.lower()):
            try:
                row = completar(p)
            except Exception as e:
                print(f"  ERR {p.name}: {e}")
                continue
            dest = unique_target(mes, row["nuevo"], used)
            flag = "OK" if row["completo"] else "REV"
            print(
                f"  [{flag}] {row['metodo']:10} {p.name[:48]} -> {dest[:60]}"
            )
            if p.name.lower() == dest.lower():
                continue
            try:
                p.rename(mes / dest)
            except Exception as e:
                print(f"    rename fail: {e}")


if __name__ == "__main__":
    main()
