# -*- coding: utf-8 -*-
"""Extrae hojas de Héctor Román de los PDFs de recibos Gastro."""
from __future__ import annotations

from pathlib import Path

import fitz

OUT = Path(r"C:\Users\recep\Desktop\Recibos_Hector_Roman_Gastro.pdf")
CUIL = "20162925971"
NEEDLES = ("ROMAN , HECTOR", "ROMAN, HECTOR", "HECTOR OSVALDO")

PDFS = [
    Path(
        r"\\TANGOSRV\Compartido\CLIENTES\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
        rf"\Sueldos\Recibos\{rel}"
    )
    for rel in (
        r"2025\202512 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2025\202512 SAC Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202601 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202602 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202603 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202604 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202605 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202606 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202606 Recibos SAC Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
        r"2026\202607 Recibos Gastroenterología y  Endoscopia Digestiva MdP S.A..pdf",
    )
]


def decode_pua(texto: str) -> str:
    out: list[str] = []
    for ch in texto:
        o = ord(ch)
        if 0xF000 <= o <= 0xF0FF:
            out.append(chr(o - 0xF000))
        else:
            out.append(ch)
    return "".join(out)


def es_roman(texto: str) -> bool:
    n = decode_pua(texto).upper()
    compact = "".join(c for c in n if c.isdigit())
    if CUIL in compact:
        return True
    return any(needle in n for needle in NEEDLES)


def main() -> None:
    dst = fitz.open()
    encontrados: list[str] = []
    faltantes: list[str] = []
    for path in PDFS:
        if not path.exists():
            print("NO EXISTE", path)
            faltantes.append(path.name)
            continue
        src = fitz.open(path)
        hits: list[int] = []
        for i, page in enumerate(src):
            if es_roman(page.get_text("text") or ""):
                hits.append(i)
                dst.insert_pdf(src, from_page=i, to_page=i)
        src.close()
        if hits:
            pags = ", ".join(str(i + 1) for i in hits)
            encontrados.append(f"{path.name}  hoja {pags}")
            print("OK", path.name, "hojas", [i + 1 for i in hits])
        else:
            faltantes.append(path.name)
            print("SIN ROMAN", path.name, "paginas", src.page_count if False else "")

    if dst.page_count == 0:
        dst.close()
        raise SystemExit("No se encontró ninguna hoja de Héctor Román")
    dst.save(OUT)
    dst.close()
    copia = Path(
        r"\\TANGOSRV\Compartido\CLIENTES\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
        r"\Sueldos\Recibos\Recibos_Hector_Roman_dic2025_a_jul2026.pdf"
    )
    copia.write_bytes(OUT.read_bytes())
    print("HOJAS", len(encontrados), "PDF paginas", fitz.open(OUT).page_count)
    print("OUT", OUT)
    print("COPIA", copia)
    for x in encontrados:
        print(" ", x)
    if faltantes:
        print("SIN HOJA:")
        for x in faltantes:
            print(" ", x)


if __name__ == "__main__":
    main()
