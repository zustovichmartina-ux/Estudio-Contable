# -*- coding: utf-8 -*-
"""Buzón OCR → grilla imputable → asiento Tango / papeles (estructura del demo)."""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd
import streamlit as st

from imputacion_bancaria import df_extracto_a_movs_imputacion, imputar_extracto
from procesador import (
    PERFILES_BANCO,
    enriquecer_df_extracto_formato_banco,
    exportar_extracto_bancario_excel,
    exportar_zip_extractos_por_banco,
    procesar_extractos_bancarios_pdfs,
)


def _padron_vacio() -> dict:
    return {"cliente": "", "cuit": "", "deudores": [], "proveedores": []}


def _padron_sesion() -> dict:
    p = st.session_state.get("extractos_padron")
    if not isinstance(p, dict):
        p = _padron_vacio()
        st.session_state.extractos_padron = p
    p.setdefault("deudores", [])
    p.setdefault("proveedores", [])
    return p


def _cargar_padron_excel(uploaded, tipo: str) -> list[dict]:
    df = pd.read_excel(uploaded) if not isinstance(uploaded, pd.DataFrame) else uploaded
    cols = {str(c).strip().lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            for k, orig in cols.items():
                if n in k:
                    return orig
        return None

    c_cuit = col("cuit", "cuil")
    c_nom = col("nombre", "razon", "razón", "proveedor", "cliente", "contraparte")
    c_cta = col("cuenta", "codigo", "código")
    out = []
    for _, row in df.iterrows():
        nombre = str(row[c_nom]).strip() if c_nom else ""
        cuit = re.sub(r"\D", "", str(row[c_cuit] if c_cuit else ""))
        if not nombre and not cuit:
            continue
        cuenta = str(row[c_cta]).strip() if c_cta and pd.notna(row[c_cta]) else ""
        pref = "Deudores por ventas" if tipo == "deudores" else "Proveedores"
        out.append(
            {
                "cuit": cuit,
                "nombre": nombre,
                "cuenta": cuenta,
                "imputacion": f"{pref} · {nombre}" if nombre else pref,
            }
        )
    return out


def _kpis_imputacion(resumen: dict, movs: list[dict]) -> None:
    cred = sum(float(m.get("credito") or 0) for m in movs)
    deb = sum(float(m.get("debito") or 0) for m in movs)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Movimientos", int(resumen.get("total") or len(movs)))
    c2.metric("Fijas", int(resumen.get("fijas") or 0))
    c3.metric("Padrón", int(resumen.get("sugeridas") or 0))
    c4.metric("A imputar", int(resumen.get("pendientes") or 0))
    c5.metric("Créditos", f"$ {cred:,.2f}")
    c6.metric("Débitos", f"$ {deb:,.2f}")


def render_herramienta_extractos() -> None:
    """Reemplaza el convertidor suelto por el flujo del demo de ayer."""
    bancos_txt = ", ".join(
        sorted(
            {
                str(p.get("nombre_display") or k)
                for k, p in PERFILES_BANCO.items()
                if k != "desconocido"
            }
        )
    )
    st.markdown("#### Extractos bancarios")
    st.caption(
        "Buzón con **OCR siempre** (digital o escaneado) → misma grilla del banco, "
        "con imputación. Los movimientos generales (IVA, IIBB, Ley 25.413, comisiones, VEP) "
        "salen con **cuenta fija**. Las transferencias se cruzan con deudores/proveedores. "
        f"Bancos: {bancos_txt}."
    )

    tab_buzon, tab_mov, tab_deu, tab_prov, tab_tango, tab_pap = st.tabs(
        [
            "Buzón OCR",
            "Movimientos",
            "Deudores",
            "Proveedores",
            "Asiento Tango",
            "Papeles de trabajo",
        ]
    )

    padron = _padron_sesion()

    with tab_buzon:
        pdfs_ext = st.file_uploader(
            "Extractos (PDF o Excel del homebanking)",
            type=["pdf", "xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="uploader_herramientas_extractos_bancarios",
            help="PDF digital o escaneado. Si es imagen, corre OCR (la primera vez tarda; después usa cache).",
        )
        if st.button("Leer extractos", type="primary", key="btn_herramientas_extracto_excel"):
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
                    movs = df_extracto_a_movs_imputacion(df_ext)
                    st.session_state.extractos_imputacion = imputar_extracto(movs, padron)
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
        else:
            n_mov = int((df_ext["Tipo fila"] == "Movimiento").sum()) if "Tipo fila" in df_ext.columns else len(df_ext)
            bancos_txt_ok = " · ".join(p["banco"] for p in paquetes)
            st.success(
                f"Listo: **{n_mov}** movimientos · **{len(paquetes)}** banco(s) · "
                f"**{len(meta_ext.get('archivos') or [])}** archivo(s)"
                + (f" · {bancos_txt_ok}" if bancos_txt_ok else "")
            )
            st.caption("La grilla imputable está en **Movimientos**. Acá descargás el Excel del extracto.")
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

    with tab_deu:
        st.caption("Padrón de deudores (CUIT / nombre / cuenta). Las transferencias a favor se cruzan acá.")
        up_d = st.file_uploader("Excel de deudores", type=["xlsx", "xls", "csv"], key="extractos_up_deudores")
        if up_d is not None and st.button("Cargar deudores", key="btn_extractos_deudores"):
            padron["deudores"] = _cargar_padron_excel(up_d, "deudores")
            st.session_state.extractos_padron = padron
            _reimputar_si_hay(padron)
            st.success(f"{len(padron['deudores'])} deudor(es).")
            st.rerun()
        if padron.get("deudores"):
            st.dataframe(pd.DataFrame(padron["deudores"]), use_container_width=True, hide_index=True)
        else:
            st.info("Sin padrón todavía. Podés seguir: las transferencias quedan en «a imputar».")

    with tab_prov:
        st.caption("Padrón de proveedores. Los débitos (TRF, CASH, ECHEQ) se cruzan acá.")
        up_p = st.file_uploader("Excel de proveedores", type=["xlsx", "xls", "csv"], key="extractos_up_proveedores")
        if up_p is not None and st.button("Cargar proveedores", key="btn_extractos_proveedores"):
            padron["proveedores"] = _cargar_padron_excel(up_p, "proveedores")
            st.session_state.extractos_padron = padron
            _reimputar_si_hay(padron)
            st.success(f"{len(padron['proveedores'])} proveedor(es).")
            st.rerun()
        if padron.get("proveedores"):
            st.dataframe(pd.DataFrame(padron["proveedores"]), use_container_width=True, hide_index=True)
        else:
            st.info("Sin padrón todavía.")

    imput = st.session_state.get("extractos_imputacion") or {}
    movs = imput.get("movimientos") or []

    with tab_mov:
        if not movs:
            st.info("Primero leé un extracto en **Buzón OCR**.")
        else:
            _kpis_imputacion(imput.get("resumen") or {}, movs)
            st.caption(
                "Generales = cuenta fija (no se editan). Transferencias = deudor/proveedor o pendiente."
            )
            filas_ui = []
            for m in movs:
                filas_ui.append(
                    {
                        "Fecha": m.get("fecha"),
                        "Descripción": m.get("descripcion"),
                        "Detalle": m.get("detalle"),
                        "Crédito": m.get("credito") or None,
                        "Débito": m.get("debito") or None,
                        "Saldo": m.get("saldo"),
                        "Imputación": m.get("imputacion"),
                        "Cuenta": m.get("cuenta"),
                        "Contraparte": m.get("contraparte"),
                        "Origen": m.get("origen_imputacion"),
                    }
                )
            st.dataframe(pd.DataFrame(filas_ui), use_container_width=True, hide_index=True)

    with tab_tango:
        st.caption(
            "El asiento sale de esta grilla. Hasta que pases la plantilla vacía de Tango, "
            "solo muestro el preview Debe/Haber (no invento el formato de importación)."
        )
        if not movs:
            st.info("No hay movimientos imputados.")
        else:
            renglones = []
            for m in movs:
                cuenta = str(m.get("cuenta") or "").strip() or "99999"
                desc = str(m.get("imputacion") or m.get("descripcion") or "")
                cred = float(m.get("credito") or 0)
                deb = float(m.get("debito") or 0)
                if deb:
                    renglones.append({"Fecha": m.get("fecha"), "Cuenta": cuenta, "Concepto": desc, "Debe": deb, "Haber": 0.0, "Origen": m.get("origen_imputacion")})
                if cred:
                    renglones.append({"Fecha": m.get("fecha"), "Cuenta": cuenta, "Concepto": desc, "Debe": 0.0, "Haber": cred, "Origen": m.get("origen_imputacion")})
            df_as = pd.DataFrame(renglones)
            debe = float(df_as["Debe"].sum()) if not df_as.empty else 0.0
            haber = float(df_as["Haber"].sum()) if not df_as.empty else 0.0
            k1, k2, k3 = st.columns(3)
            k1.metric("Debe", f"$ {debe:,.2f}")
            k2.metric("Haber", f"$ {haber:,.2f}")
            k3.metric("Diferencia", f"$ {debe - haber:,.2f}")
            st.dataframe(df_as, use_container_width=True, hide_index=True)
            st.button("Generar asiento Tango", disabled=True, help="Plantilla vacía pendiente", key="btn_extractos_tango_disabled")

    with tab_pap:
        st.caption(
            "Papeles de trabajo del balance (cuadros bancarios). "
            "El formato sale de la plantilla vacía que pases; hasta entonces no lo invento."
        )
        st.button(
            "Armar papeles de trabajo",
            disabled=True,
            help="Plantilla vacía pendiente",
            key="btn_extractos_papeles_disabled",
        )
        df_ext = st.session_state.get("extracto_santander_df")
        if df_ext is not None and not getattr(df_ext, "empty", True):
            prev = enriquecer_df_extracto_formato_banco(df_ext)
            cols = [c for c in ("Fecha", "Descripcion", "Detalle", "Importe", "Saldo", "Clasificacion") if c in prev.columns]
            st.dataframe(prev[cols].head(40), use_container_width=True, hide_index=True)


def _reimputar_si_hay(padron: dict) -> None:
    df_ext = st.session_state.get("extracto_santander_df")
    if df_ext is None or getattr(df_ext, "empty", True):
        return
    movs = df_extracto_a_movs_imputacion(df_ext)
    st.session_state.extractos_imputacion = imputar_extracto(movs, padron)
