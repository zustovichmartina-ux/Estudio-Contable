# -*- coding: utf-8 -*-
"""Diagnostico detallado del parser de Provincia y Santander."""
import sys, io, re
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import fitz, numpy as np
from PIL import Image
from collections import defaultdict
from datetime import date as date_type

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

def _limpiar_monto(s):
    if not s:
        return 0.0
    s = str(s).strip().replace("$", "").replace(" ", "")
    s = s.replace("o", "0").replace("O", "0").replace("l", "1").replace("|", "1")
    if not re.match(r"^[\d.,]+$", s):
        s = re.sub(r"[^0-9.,]", "", s)
    if not s:
        return 0.0
    last_dot = s.rfind(".")
    last_comma = s.rfind(",")
    last_sep = max(last_dot, last_comma)
    if last_sep == -1:
        try:
            return float(s)
        except ValueError:
            return 0.0
    decimal_part = s[last_sep + 1:]
    integer_part = s[:last_sep]
    if re.match(r"^\d{1,2}$", decimal_part):
        int_clean = re.sub(r"[.,]", "", integer_part)
        s = (int_clean or "0") + "." + decimal_part
    else:
        s = re.sub(r"[.,]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0

import glob

# ===================== PROVINCIA =====================
print("=" * 60)
print("DIAGNOSTICO PROVINCIA")
print("=" * 60)

pdfs_prov = sorted(glob.glob(r"C:\Users\recep\Desktop\Estudio Contable\extractos bancarios\Prestamos Financieros\*Provincia*.pdf"))
print(f"PDFs Provincia: {len(pdfs_prov)}")

if pdfs_prov:
    pdf_path = pdfs_prov[0]
    print(f"PDF: {pdf_path}")
    textos = ocr_pdf(pdf_path, rotar_grados=90)

    # Normalizar
    textos_n = []
    for texto in textos:
        t = re.sub(r"(\d)\s+([,.])\s*(\d)", r"\1\2\3", texto)
        t = re.sub(r"([,.])\s+(\d)", r"\1\2", t)
        textos_n.append(t)

    pat_fecha_iso = re.compile(r"\b(\d{4})[\s\-](\d{2})[\s\-](\d{2})\b")
    pat_monto = re.compile(r"(?<!\d)[\d]+(?:[.,]\d{3})*[.,]\d{2}(?!\d)")

    print("\n--- Lineas con fechas detectadas ---")
    pagos_por_fecha = defaultdict(list)
    for pg_idx, texto in enumerate(textos_n):
        for linea_idx, linea in enumerate(texto.splitlines()):
            m = pat_fecha_iso.search(linea)
            if not m:
                continue
            try:
                year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if year < 2000 or year > 2040 or month < 1 or month > 12 or day < 1 or day > 31:
                    print(f"  P{pg_idx+1}:L{linea_idx+1} FECHA INVALIDA {year}-{month}-{day}: {linea[:80]}")
                    continue
                fecha = date_type(year, month, day)
            except ValueError as e:
                print(f"  P{pg_idx+1}:L{linea_idx+1} ERROR FECHA: {e}: {linea[:80]}")
                continue

            nums = pat_monto.findall(linea)
            importes = [_limpiar_monto(x) for x in nums if _limpiar_monto(x) >= 100]
            print(f"  P{pg_idx+1}:L{linea_idx+1} FECHA={fecha} | nums_raw={nums[:5]} | importes>100={importes[:5]}")
            if importes:
                pagos_por_fecha[fecha].append(importes)

    print(f"\n--- Fechas encontradas con importes: {len(pagos_por_fecha)} ---")
    cuotas_resultado = []
    for i, fecha in enumerate(sorted(pagos_por_fecha.keys()), 1):
        filas = pagos_por_fecha[fecha]
        totales_filas = [max(f) for f in filas]
        totales_filtrados = [v for v in totales_filas if v >= 100]
        if len(totales_filtrados) >= 2:
            tots = sorted(totales_filtrados, reverse=True)
            capital, intereses = tots[0], tots[1]
        elif len(totales_filtrados) == 1:
            capital, intereses = 0.0, totales_filtrados[0]
        else:
            continue
        total = round(capital + intereses, 2)
        cuotas_resultado.append({"cuota": i, "vencimiento": str(fecha), "capital": capital, "intereses": intereses, "total": total})
        print(f"  Cuota {i}: {fecha} | capital={capital:.2f} | intereses={intereses:.2f} | total={total:.2f}")

    print(f"\nTOTAL CUOTAS PROVINCIA: {len(cuotas_resultado)}")

# ===================== SANTANDER =====================
print("\n" + "=" * 60)
print("DIAGNOSTICO SANTANDER")
print("=" * 60)

pdfs_sant = sorted(glob.glob(r"C:\Users\recep\Desktop\Estudio Contable\extractos bancarios\Prestamos Financieros\*Santander*.pdf"))
print(f"PDFs Santander: {len(pdfs_sant)}")

if pdfs_sant:
    pdf_path = pdfs_sant[0]
    print(f"PDF: {pdf_path}")
    textos = ocr_pdf(pdf_path, rotar_grados=0)
    print(f"\n--- Texto OCR crudo (primeras 50 lineas pagina 1) ---")
    for j, linea in enumerate(textos[0].splitlines()[:50]):
        print(f"  L{j+1:03d}: {linea}")

    pat_fecha_flex = re.compile(r"\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}")
    pat_monto = re.compile(r"(?<!\d)[\d]+(?:[.,]\d{3})*[.,]\d{2}(?!\d)")

    print(f"\n--- Lineas con fechas detectadas ---")
    en_tabla = False
    for pg_idx, texto in enumerate(textos):
        for j, linea in enumerate(texto.splitlines()):
            low = linea.lower()
            if not en_tabla:
                if any(k in low for k in ("fecha", "vencim", "capital", "desarrollo", "cuota")):
                    en_tabla = True
                    print(f"  P{pg_idx+1}:L{j+1} [ENCABEZADO TABLA]: {linea[:80]}")
                continue
            m_fecha = pat_fecha_flex.search(linea)
            if not m_fecha:
                continue
            nums = pat_monto.findall(linea)
            importes = [_limpiar_monto(x) for x in nums if _limpiar_monto(x) >= 100]
            print(f"  P{pg_idx+1}:L{j+1} FECHA={m_fecha.group()} | nums={nums[:5]} | importes={importes[:5]}")
