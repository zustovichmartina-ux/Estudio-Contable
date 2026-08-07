# -*- coding: utf-8 -*-
import sys, io
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import fitz, numpy as np
from PIL import Image
from collections import defaultdict

def ocr_pdf(pdf_path, rotar_grados=90):
    import easyocr
    lector = easyocr.Reader(["es", "en"], gpu=False, verbose=False)
    doc = fitz.open(str(pdf_path))
    textos = []
    for page in doc:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        if rotar_grados != 0:
            img = img.rotate(rotar_grados, expand=True)
        resultados = lector.readtext(np.array(img))
        filas = defaultdict(list)
        for bbox, txt, _ in resultados:
            y = int((bbox[0][1] + bbox[2][1]) / 2 / 15) * 15
            filas[y].append((bbox[0][0], txt))
        lineas = [" ".join(t for _, t in sorted(filas[y], key=lambda x: x[0])) for y in sorted(filas)]
        textos.append("\n".join(lineas))
    return textos

import glob
pdfs = sorted(glob.glob(r"C:\Users\recep\Desktop\Estudio Contable\extractos bancarios\Prestamos Financieros\*Provincia*.pdf"))
print(f"PDFs Provincia encontrados: {len(pdfs)}")
for pdf_path in pdfs:
    print(f"\n{'='*60}")
    print(f"Procesando: {pdf_path}")
    for grados in [90, 270, 0]:
        print(f"\n--- Rotacion {grados} grados ---")
        try:
            textos = ocr_pdf(pdf_path, rotar_grados=grados)
            for i, t in enumerate(textos):
                print(f"\n=== PAGINA {i+1} (rot={grados}) ===")
                lineas = t.splitlines()
                for j, linea in enumerate(lineas[:60]):
                    print(f"  L{j+1:03d}: {linea}")
            # Si hay texto razonable, parar con esta rotacion
            texto_total = "\n".join(textos)
            tiene_fechas = len([l for l in texto_total.splitlines() if any(c.isdigit() for c in l) and ('/' in l or '-' in l)]) > 3
            if tiene_fechas:
                print(f"\n[OK] Rotacion {grados} produce texto con fechas. Suficiente para analisis.")
                break
        except Exception as e:
            print(f"  ERROR: {e}")
    break  # solo primer PDF
