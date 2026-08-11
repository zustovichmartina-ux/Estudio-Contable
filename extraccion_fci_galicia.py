# -*- coding: utf-8 -*-
"""Compatibilidad: el extractor FCI ahora vive en extraccion_fci.py (cualquier banco)."""
from extraccion_fci import (  # noqa: F401
    extraer_texto_pdf,
    extraer_y_guardar,
    parsear_extracto_fci_galicia,
    procesar_pdfs_fci,
)
