"""Exporta Excel del extracto 2026-03 desde cache OCR."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from procesador import _parsear_movimientos_santander_paginas, exportar_extracto_santander_excel

CACHE = Path(__file__).resolve().parent / "_tmp_santander_2026-03_ocr.json"
OUTS = [
    Path(r"C:\Users\recep\Downloads\Extracto_Santander_2026-03_OCR.xlsx"),
    Path(r"C:\Users\recep\Desktop\Extracto_Santander_Oftalmologia_RELE_2026-03.xlsx"),
    Path(
        r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
        r"\Balances\2026\Banco Santander\Extracto_Santander_2026-03_OCR.xlsx"
    ),
]


def main() -> None:
    pages = json.loads(CACHE.read_text(encoding="utf-8"))
    paginas = [(i + 1, "\n".join(page)) for i, page in enumerate(pages)]
    movs, meta = _parsear_movimientos_santander_paginas(paginas, "2026-03.pdf")
    df = pd.DataFrame(movs)
    n_mov = int((df["Tipo fila"] == "Movimiento").sum())
    saldo = df["Saldo"].dropna().iloc[-1]
    print(f"movs={n_mov} filas={len(df)} saldo_final={saldo}", flush=True)
    xlsx = exportar_extracto_santander_excel(df, meta)
    for out in OUTS:
        try:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(xlsx)
            print(f"OK {out}", flush=True)
        except Exception as exc:
            print(f"FAIL {out}: {exc}", flush=True)


if __name__ == "__main__":
    main()
