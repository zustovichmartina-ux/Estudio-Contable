# -*- coding: utf-8 -*-
"""OCR rapido: solo 1 pagina de cada banco que necesita OCR."""
import sys, io
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import fitz
import easyocr
import numpy as np
from PIL import Image
from pathlib import Path
from collections import defaultdict

CARPETA = Path(r"C:\Users\recep\Desktop\Estudio Contable\extractos bancarios\Prestamos Financieros")

def ocr_page(pdf_path, pagina=0, dpi=150, rotar=0):
    lector = easyocr.Reader(["es", "en"], gpu=False, verbose=False)
    doc = fitz.open(str(pdf_path))
    page = doc[pagina]
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    if rotar:
        img = img.rotate(rotar, expand=True)
    resultados = lector.readtext(np.array(img))
    # Reconstruir lineas por Y
    filas = defaultdict(list)
    for bbox, txt, conf in resultados:
        y = int((bbox[0][1] + bbox[2][1]) / 2 / 15) * 15
        filas[y].append((bbox[0][0], txt, conf))
    lineas = []
    for y in sorted(filas):
        partes = sorted(filas[y], key=lambda x: x[0])
        linea = " ".join(t for _, t, _ in partes)
        confs = [c for _, _, c in partes]
        conf_avg = sum(confs)/len(confs)
        lineas.append((y, linea, conf_avg))
    return lineas

# ---- SANTANDER ----
pdfs_sant = sorted([f for f in CARPETA.glob("*.pdf") if "santander" in f.name.lower()])
if pdfs_sant:
    print(f"\n{'='*60}")
    print(f"SANTANDER: {pdfs_sant[0].name}")
    print(f"{'='*60}")
    lineas = ocr_page(pdfs_sant[0], pagina=0, dpi=150, rotar=0)
    for y, linea, conf in lineas:
        print(f"  [{conf:.2f}] {linea}")

# ---- PROVINCIA ----
pdfs_prov = sorted([f for f in CARPETA.glob("*.pdf") if "provincia" in f.name.lower()])
for rot in [0, 90, 270]:
    if pdfs_prov:
        print(f"\n{'='*60}")
        print(f"PROVINCIA (rot={rot}): {pdfs_prov[0].name}")
        print(f"{'='*60}")
        lineas = ocr_page(pdfs_prov[0], pagina=0, dpi=150, rotar=rot)
        for y, linea, conf in lineas:
            print(f"  [{conf:.2f}] {linea}")
        # Si hay texto util, no seguir probando rotaciones
        texto_all = " ".join(l for _, l, _ in lineas)
        if any(c.isdigit() for c in texto_all[:200]):
            break

# ---- NACION ----
pdfs_nac = sorted([f for f in CARPETA.glob("*.pdf") if "nacion" in f.name.lower()])
if pdfs_nac:
    print(f"\n{'='*60}")
    print(f"NACION: {pdfs_nac[0].name}")
    print(f"{'='*60}")
    lineas = ocr_page(pdfs_nac[0], pagina=0, dpi=150, rotar=0)
    for y, linea, conf in lineas:
        print(f"  [{conf:.2f}] {linea}")

print("\n\nDIAGNOSTICO COMPLETO")
