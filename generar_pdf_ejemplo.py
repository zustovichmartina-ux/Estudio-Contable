#!/usr/bin/env python3
"""Genera PDFs de EJEMPLO (datos ficticios, CAE ficticio) a partir de las filas EJEMPLO del modelo."""
import datetime as dt
import os
import sys

import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from arca.comprobante import Comprobante, agrupar, leer_planilla  # noqa: E402
from arca.pdf import generar_pdf  # noqa: E402

xlsx = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "plantillas", "Facturas_a_emitir_MODELO.xlsx")
out = os.path.join(BASE, "salida", "pdf_ejemplo")
os.makedirs(out, exist_ok=True)
for i, g in enumerate(agrupar(leer_planilla(openpyxl.load_workbook(xlsx, data_only=True))), 1):
    cb = Comprobante(g)
    if not cb.es_ejemplo or not cb.validar():
        continue
    cae_ficticio = f"7{cb.cbte_tipo:03d}0000000{i:03d}"[:14]
    vto = (cb.fecha + dt.timedelta(days=10)).strftime("%Y%m%d")
    p = os.path.join(out, f"EJEMPLO_{cb.cuit}_{cb.cbte_tipo:03d}_{cb.pto_vta:05d}-{i:08d}.pdf")
    qr = generar_pdf(p, cb, i, cae_ficticio, vto, marca_agua="EJEMPLO - SIN VALIDEZ")
    print(p, "\n   QR:", qr)
