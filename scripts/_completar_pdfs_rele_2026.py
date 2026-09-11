# -*- coding: utf-8 -*-
"""PDF-only: completar nro/proveedor en archivos incompletos (sin OCR)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE.parent))
sys.path.insert(0, str(_HERE))

from _renombrar_fcts_rele_2026 import (
    ROOT,
    SKIP_MESES,
    _desde_nombre,
    _desde_texto,
    _texto_pdf,
    armar_nombre,
    unique_target,
)

PDF_INCOMP = re.compile(
    r"PROVEEDOR|"
    r"^FCCA CENTRO MEDICO|"
    r"^FCCA COLON SA|"
    r"^FCCC |"
    r"^FCCA MARINO|"
    r"^FCCA TRUJILLO|"
    r"30-71802274|30-70066760",
    re.I,
)


def main() -> None:
    meses = sorted(
        p for p in ROOT.iterdir() if p.is_dir() and p.name not in SKIP_MESES
    )
    for mes in meses:
        files = [
            p
            for p in mes.iterdir()
            if p.is_file()
            and p.suffix.lower() == ".pdf"
            and PDF_INCOMP.search(p.name)
        ]
        print("====", mes.name, len(files))
        used = {p.name.lower() for p in mes.iterdir() if p.is_file()}
        for p in sorted(files, key=lambda x: x.name.lower()):
            info = _desde_nombre(p.name)
            if re.search(r"(FCC|NCC|NDC)[ABC](20|23|27|30|33)-(\d{8,})", p.name, re.I):
                info = {k: v for k, v in info.items() if k == "prov"}
            texto = _texto_pdf(p)
            extra = _desde_texto(texto)
            for k, v in extra.items():
                if not info.get(k):
                    info[k] = v
            nuevo = armar_nombre(info, p.name)
            ok = bool(info.get("nro") and info.get("prov") and info.get("prov") != "PROVEEDOR")
            print(f"  {'OK' if ok else 'REV'} {p.name[:50]}")
            print(f"     -> {nuevo[:70]} nro={info.get('nro')} pv={info.get('pv')} {info.get('prov')}")
            print("     txt:", (texto or "").replace("\n", " | ")[:180])
            if not ok:
                continue
            dest = unique_target(mes, nuevo, used, current=p.name)
            if p.name.lower() == dest.lower():
                continue
            p.rename(mes / dest)
            print("     REN", dest)


if __name__ == "__main__":
    main()
