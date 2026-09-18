"""Tema visual compartido del Estudio Contable (Streamlit).

Paleta y logo del estudio. La fluidez (radios, spacing, transiciones)
vive acá para que shell, botones, inputs, tablas y buzones se sientan
igual en toda la app.
"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

_CSS_PATH = Path(__file__).resolve().parent / "assets" / "estudio.css"

# Paleta canónica del estudio (no teal/gold).
COLORES_ESTUDIO = {
    "navy": "#0B1C33",
    "primary": "#2563EB",
    "primary_soft": "#3B82F6",
    "bg": "#F4F6FA",
    "ink": "#0F172A",
    "muted": "#64748B",
    "line": "#E2E8F0",
    "card": "#FFFFFF",
}

LOGO_ESTUDIO = Path(__file__).resolve().parent / "assets" / "estudio-zona-guemes-wordmark-oscuro.png"


def cargar_css() -> str:
    return _CSS_PATH.read_text(encoding="utf-8")


def inyectar_tema() -> None:
    """Inyecta el CSS compartido. Llamar después de st.set_page_config."""
    st.markdown(f"<style>\n{cargar_css()}\n</style>", unsafe_allow_html=True)
