"""UI Streamlit del módulo ARCA: facturación electrónica."""
from __future__ import annotations

import streamlit as st

from ui_arca_facturacion import render_facturacion_arca


def render_arca_module() -> None:
    """Facturación por Web Service y consulta (todavía sin armar)."""
    st.caption(
        "Facturación electrónica por Web Service. "
        "Las claves fiscales no se cargan en esta web."
    )
    tab_fe, tab_consulta = st.tabs(["Facturación", "Consulta"])
    with tab_fe:
        render_facturacion_arca()
    with tab_consulta:
        st.info("Próximamente")
