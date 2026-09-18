# -*- coding: utf-8 -*-
"""Herramienta: extracto PDF/Excel del banco → Excel descargable."""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd
import streamlit as st

from procesador import (
    PERFILES_BANCO,
    df_extracto_convertidor_sin_saldos,
    exportar_extracto_bancario_excel,
    exportar_zip_extractos_por_banco,
    procesar_extractos_bancarios_pdfs,
)


def render_herramienta_extractos() -> None:
    bancos_txt = ", ".join(
        sorted(
            {
                str(p.get("nombre_display") or k)
                for k, p in PERFILES_BANCO.items()
                if k != "desconocido"
            }
        )
    )
    st.markdown("#### Extractos de PDF → Excel")
    st.caption(
        "Convertidor: subís el extracto y bajás el Excel. OCR si el PDF es escaneo. "
        f"Bancos: {bancos_txt}. "
        "Excel: Fecha, Concepto, Débitos y Créditos — sin saldo inicial, saldo final ni columna de saldo. "
        "La conciliación (clasificar, retenciones, asiento) está en **Conciliación Bancaria**."
    )

    pdfs_ext = st.file_uploader(
        "Extractos (PDF o Excel del homebanking)",
        type=["pdf", "xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="uploader_herramientas_extractos_bancarios",
        help="PDF digital o escaneado. Si es imagen, corre OCR (la primera vez tarda; después usa cache).",
    )
    if st.button("Convertir a Excel", type="primary", key="btn_herramientas_extracto_excel"):
        if not pdfs_ext:
            st.warning("Subí al menos un PDF o Excel.")
        else:
            with st.spinner(
                "Leyendo archivos… OCR automático si el PDF es escaneo. No cierres la pestaña."
            ):
                df_ext, meta_ext, err_ext = procesar_extractos_bancarios_pdfs(pdfs_ext)
                st.session_state.extracto_santander_df = df_ext
                st.session_state.extracto_santander_meta = meta_ext
                st.session_state.extracto_santander_errores = err_ext
                paquetes_ui: list[dict] = []
                for p in meta_ext.get("por_banco") or []:
                    df_b = p.get("df")
                    meta_b = p.get("meta") or {}
                    if df_b is None or getattr(df_b, "empty", True):
                        continue
                    xlsx_b = exportar_extracto_bancario_excel(df_b, meta_b)
                    paquetes_ui.append(
                        {
                            "banco": p.get("banco") or meta_b.get("banco") or "Banco",
                            "banco_slug": p.get("banco_slug") or "",
                            "df": df_b,
                            "meta": meta_b,
                            "xlsx": xlsx_b,
                            "pdf_merged": p.get("pdf_merged"),
                        }
                    )
                st.session_state.extracto_paquetes = paquetes_ui
                if len(paquetes_ui) > 1:
                    st.session_state.extracto_zip = exportar_zip_extractos_por_banco(
                        paquetes_ui,
                        cuit=str(meta_ext.get("cuit") or st.session_state.get("cuit_activo") or ""),
                    )
                else:
                    st.session_state.extracto_zip = None
                if paquetes_ui:
                    st.session_state.extracto_santander_xlsx = paquetes_ui[0]["xlsx"]
                    st.session_state.extracto_santander_pdf_merged = (
                        None if len(paquetes_ui) > 1 else paquetes_ui[0].get("pdf_merged")
                    )
                else:
                    st.session_state.extracto_santander_xlsx = None
                    st.session_state.extracto_santander_pdf_merged = None
            st.rerun()

    err_ext = st.session_state.get("extracto_santander_errores") or []
    if err_ext:
        st.error("Algún archivo no salió. El detalle está abajo; el resto, si hubo, sí se armó.")
        with st.expander(f"Detalle ({len(err_ext)})", expanded=True):
            st.dataframe(pd.DataFrame(err_ext), use_container_width=True, hide_index=True)

    paquetes = st.session_state.get("extracto_paquetes") or []
    df_ext = st.session_state.get("extracto_santander_df")
    meta_ext = st.session_state.get("extracto_santander_meta") or {}
    if not paquetes or df_ext is None or getattr(df_ext, "empty", True):
        if not err_ext and st.session_state.get("extracto_santander_df") is not None:
            st.warning("No salieron movimientos. Probá otro PDF o el Excel del homebanking.")
        return

    n_mov = (
        int((df_ext["Tipo fila"] == "Movimiento").sum())
        if "Tipo fila" in df_ext.columns
        else len(df_ext)
    )
    bancos_txt_ok = " · ".join(p["banco"] for p in paquetes)
    st.success(
        f"Listo: **{n_mov}** movimientos · **{len(paquetes)}** banco(s) · "
        f"**{len(meta_ext.get('archivos') or [])}** archivo(s)"
        + (f" · {bancos_txt_ok}" if bancos_txt_ok else "")
    )

    cuit_limpio = re.sub(
        r"\D", "", str(meta_ext.get("cuit") or st.session_state.get("cuit_activo") or "")
    )
    zip_bytes = st.session_state.get("extracto_zip")
    if zip_bytes and len(paquetes) > 1:
        st.download_button(
            f"Descargar ZIP ({len(paquetes)} Excel, uno por banco)",
            data=zip_bytes,
            file_name=f"Extractos_por_banco_{cuit_limpio or 'cliente'}_{datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            type="primary",
            key="dl_herramientas_extracto_zip",
            use_container_width=True,
        )
    for i, p in enumerate(paquetes):
        banco = p.get("banco") or "Banco"
        slug_banco = re.sub(r"[^A-Za-z0-9]+", "_", str(banco).strip())[:24] or "banco"
        stamp = datetime.now().strftime("%Y-%m-%d_%H %M")
        st.download_button(
            f"Descargar Excel — {banco}",
            data=p["xlsx"],
            file_name=f"{stamp}_{slug_banco}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" if len(paquetes) == 1 else "secondary",
            key=f"dl_herramientas_extracto_xlsx_{i}_{slug_banco}",
            use_container_width=True,
        )

    vista = df_extracto_convertidor_sin_saldos(df_ext)
    if not vista.empty:
        st.dataframe(vista.head(80), use_container_width=True, hide_index=True)
        st.caption(f"Vista previa ({min(80, len(vista))} de {len(vista)} filas).")
