# -*- coding: utf-8 -*-
"""
Diagnostico rapido: extrae texto nativo de todos los PDFs para ver si tienen texto embebido.
Si tienen texto, muestra las primeras 80 lineas. No usa OCR (rapido).
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pdfplumber
from pathlib import Path

CARPETA = Path(r"C:\Users\recep\Desktop\Estudio Contable\extractos bancarios\Prestamos Financieros")

pdfs = sorted(CARPETA.glob("*.pdf"))
print(f"Total PDFs: {len(pdfs)}\n")

for pdf in pdfs:
    nombre = pdf.name
    banco = "?"
    n = nombre.lower()
    if "santander" in n: banco = "SANTANDER"
    elif "provincia" in n: banco = "PROVINCIA"
    elif "galicia" in n: banco = "GALICIA"
    elif "frances" in n or "francés" in n: banco = "FRANCES"
    elif "mercado" in n: banco = "MERCADOPAGO"
    elif "nacion" in n or "nació" in n: banco = "NACION"
    
    print(f"\n{'='*70}")
    print(f"[{banco}] {nombre[:60]}")
    
    try:
        with pdfplumber.open(str(pdf)) as doc:
            total_chars = 0
            for i, page in enumerate(doc.pages):
                t = page.extract_text() or ""
                total_chars += len(t.strip())
            
            print(f"  Paginas: {len(doc.pages)} | Chars totales: {total_chars}")
            tiene_texto = total_chars > 100
            print(f"  Tiene texto nativo: {tiene_texto}")
            
            if tiene_texto:
                # Mostrar primeras lineas utiles
                for i, page in enumerate(doc.pages):
                    t = page.extract_text() or ""
                    if t.strip():
                        print(f"  --- Pagina {i+1} (primeras 30 lineas) ---")
                        for j, linea in enumerate(t.splitlines()):
                            if linea.strip():
                                print(f"    {linea}")
                            if j >= 29:
                                print("    ...")
                                break
                        break  # Solo mostrar primera pagina con texto
            else:
                print("  [!] Sin texto nativo - necesitara OCR")
    except Exception as e:
        print(f"  ERROR: {e}")
