"""Imprime 1ª hoja de cada comprobante único (carpeta Aurora) a HP M203."""

from __future__ import annotations

import re
import time
import subprocess
from collections import defaultdict
from pathlib import Path

import fitz
import win32print

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\AURORA DE LAS SIERRAS SRL"
    r"\Cierre 2026\202607\Facturas"
)
TMP = Path(r"C:\Users\recep\Desktop\Estudio Contable\_tmp_print_aurora")
LOG = TMP / "print_log.txt"
ACROBAT = Path(r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe")
PRINTER = "NPIF7747C (HP LaserJet M203dw)"
DRIVER = "HP LaserJet M203-M206 PCL-6 (V4)"
PORT = "WSD-ab3844e9-09e2-4a22-a6c7-ddb9f859e9b2"

PAT = re.compile(
    r"(?i)^(FC(?:-REMITO)?|NC|REC)\s*(?:C\s+)?(\d+)\s*-\s*[A-Z]?(\d+)"
)


def key_of(name: str) -> tuple:
    m = PAT.search(name)
    if not m:
        return ("RAW", name.lower())
    return (m.group(1).upper(), str(int(m.group(2))), str(int(m.group(3))))


def score(path: Path) -> tuple:
    """Menor score = preferido (sin (2), sin guia, menos páginas, más chico)."""
    name = path.name.lower()
    doc = fitz.open(path)
    pages = doc.page_count
    doc.close()
    return (
        1 if "(2)" in name or "(1)" in name else 0,
        1 if "con guia" in name or "con guía" in name else 0,
        pages,
        path.stat().st_size,
        path.name,
    )


def elegir_unicos() -> list[Path]:
    groups: dict[tuple, list[Path]] = defaultdict(list)
    for p in sorted(FOLDER.glob("*.pdf")):
        groups[key_of(p.name)].append(p)
    chosen = []
    for key, paths in sorted(groups.items()):
        best = sorted(paths, key=score)[0]
        chosen.append(best)
    return chosen


def primera_hoja(src: Path, dst: Path) -> int:
    doc = fitz.open(src)
    pages = doc.page_count
    out = fitz.open()
    out.insert_pdf(doc, from_page=0, to_page=0)
    out.save(dst)
    out.close()
    doc.close()
    return pages


def log(msg: str) -> None:
    line = msg
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        print(line.encode("ascii", "replace").decode("ascii"), flush=True)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> None:
    TMP.mkdir(parents=True, exist_ok=True)
    if LOG.exists():
        LOG.unlink()

    actual = win32print.GetDefaultPrinter()
    log(f"Impresora default actual: {actual}")
    if actual != PRINTER:
        win32print.SetDefaultPrinter(PRINTER)
        log(f"Default cambiada a: {PRINTER}")
    else:
        log(f"Default OK: {PRINTER}")

    if not ACROBAT.exists():
        raise SystemExit(f"No está Acrobat: {ACROBAT}")

    chosen = elegir_unicos()
    log(f"Comprobantes únicos a imprimir: {len(chosen)}")
    log("(Duplicados omitidos; multipagina -> solo hoja 1)")

    ok = fail = 0
    for i, src in enumerate(chosen, 1):
        safe = re.sub(r"[^\w.\-]+", "_", src.stem)[:80]
        dst = TMP / f"{i:03d}_{safe}.pdf"
        try:
            pages = primera_hoja(src, dst)
            log(f"[{i}/{len(chosen)}] {src.name}  (orig {pages}p → 1p)")
            # Acrobat /t: imprime y cierra
            cmd = [
                str(ACROBAT),
                "/t",
                str(dst),
                PRINTER,
                DRIVER,
                PORT,
            ]
            proc = subprocess.run(cmd, timeout=60)
            if proc.returncode not in (0, None):
                log(f"  WARN returncode={proc.returncode}")
            ok += 1
            # Dar tiempo al spooler / Acrobat
            time.sleep(2.5)
        except Exception as exc:
            fail += 1
            log(f"  ERROR {src.name}: {exc}")

    log(f"LISTO ok={ok} fail={fail}")
    log(f"Temp: {TMP}")


if __name__ == "__main__":
    main()
