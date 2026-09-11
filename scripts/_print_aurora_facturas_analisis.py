"""Analiza PDFs Aurora: páginas y duplicados."""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import fitz

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Cierre 2026\202607\Facturas"
)
PAT = re.compile(
    r"(?i)^(FC(?:-REMITO)?|NC|REC)\s*(?:C\s+)?(\d+)\s*-\s*[A-Z]?(\d+)"
)


def key_of(name: str) -> tuple:
    m = PAT.search(name)
    if not m:
        return ("RAW", name.lower())
    return (m.group(1).upper(), str(int(m.group(2))), str(int(m.group(3))))


def main() -> None:
    groups: dict[tuple, list] = defaultdict(list)
    for p in sorted(FOLDER.glob("*.pdf")):
        doc = fitz.open(p)
        pages = doc.page_count
        doc.close()
        groups[key_of(p.name)].append((p.name, pages, p.stat().st_size))

    print("=== POR COMPROBANTE ===")
    multi_n = dup_n = 0
    for k, items in groups.items():
        flags = []
        if len(items) > 1:
            flags.append("DUP")
            dup_n += 1
        if any(pg > 1 for _, pg, _ in items):
            flags.append("MULTI")
            multi_n += 1
        flag = " ".join(flags) or "ok"
        print(f"{k}: {flag}")
        for name, pages, size in items:
            print(f"  {pages}p  {size:>7}  {name}")
    print("grupos", len(groups), "archivos", sum(len(v) for v in groups.values()))
    print("con duplicado", dup_n, "con multipagina", multi_n)


if __name__ == "__main__":
    main()
