# -*- coding: utf-8 -*-
"""Conciliación Bancaria con el formato AE-Studio, usando las reglas de la web."""
from __future__ import annotations

from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

import database as db
from capa_revision import resolver_codigo_plan
from conceptos_bancos import CACHE_PATH, cargar_instructivo
from motor_conciliacion import (
    CATEGORIA_A_CUENTA_HINT,
    bucket_ae,
    correr_motor,
    df_extracto_a_filas,
    money,
    movimientos_a_filas_grilla_tango,
    validar_saldos_corridos,
)
from procesador import clasificar_movimiento_extracto, procesar_extractos_bancarios_pdfs
from ui_motor_conciliacion import _cargar_proveedores_desde_upload, _cargar_veps_excel

_VISTAS = (
    ("ingreso", "Ingresos"),
    ("egreso", "Egresos"),
    ("retencion", "Retenciones"),
    ("deduccion", "Deducciones"),
    ("inter-cta", "Inter-cuentas"),
    ("sin-cat", "Sin clasificar"),
)


def _periodo_str(d: date | None) -> str:
    if not d:
        return date.today().strftime("%Y-%m-01")
    return d.replace(day=1).isoformat()


def _enriquecer(movs: list[dict]) -> list[dict]:
    out = []
    for m in movs:
        cred = float(money(m.get("credito")))
        deb = float(money(m.get("debito")))
        label = clasificar_movimiento_extracto(
            str(m.get("descripcion") or ""),
            importe=cred - deb,
        )
        fila = dict(m)
        fila["extracto_label"] = label
        fila["vista"] = bucket_ae(m, extracto_label=label)
        fila["monto"] = round(cred - deb, 2)
        out.append(fila)
    return out


def _kpis(movs: list[dict]) -> None:
    grupos = {k: [m for m in movs if m.get("vista") == k] for k, _ in _VISTAS}
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric(
        "Ingresos",
        f"$ {sum(m['monto'] for m in grupos['ingreso'] if m['monto'] > 0):,.2f}",
        f"{len(grupos['ingreso'])} mov.",
    )
    c2.metric(
        "Egresos",
        f"$ {sum(abs(m['monto']) for m in grupos['egreso']):,.2f}",
        f"{len(grupos['egreso'])} mov.",
    )
    c3.metric(
        "Retenciones",
        f"$ {sum(abs(m['monto']) for m in grupos['retencion']):,.2f}",
        f"{len(grupos['retencion'])} mov.",
    )
    c4.metric(
        "Deducciones",
        f"$ {sum(abs(m['monto']) for m in grupos['deduccion']):,.2f}",
        f"{len(grupos['deduccion'])} mov.",
    )
    c5.metric(
        "Inter-cuentas",
        f"$ {sum(abs(m['monto']) for m in grupos['inter-cta']):,.2f}",
        f"{len(grupos['inter-cta'])} mov.",
    )
    n_sin = len(grupos["sin-cat"])
    if n_sin:
        st.caption(f"{n_sin} movimiento(s) sin clasificar — hay que imputarlos a mano.")


def _tabla(movs: list[dict]) -> None:
    if not movs:
        st.info("Sin movimientos en esta vista.")
        return
    filas = []
    for m in movs:
        filas.append(
            {
                "Fecha": m.get("fecha"),
                "Descripción": m.get("descripcion"),
                "Crédito": float(money(m.get("credito"))) or None,
                "Débito": float(money(m.get("debito"))) or None,
                "Cuenta (Conceptos Bancos)": m.get("categoria"),
                "Código plan": m.get("cuenta_codigo") or "",
                "Vista": m.get("vista"),
                "Estado": m.get("estado"),
                "Confianza": m.get("confianza") or "",
                "Por qué": m.get("match_detalle") or m.get("citar") or "",
                "Banco": m.get("banco"),
            }
        )
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)


def _aplicar_plan(movs: list[dict], plan_df) -> list[dict]:
    out = []
    for m in movs:
        fila = dict(m)
        codigo, desc_plan, score = resolver_codigo_plan(
            str(m.get("categoria") or ""),
            plan_df,
            hints=CATEGORIA_A_CUENTA_HINT,
        )
        fila["cuenta_codigo"] = codigo
        fila["cuenta_plan"] = desc_plan
        fila["score_plan"] = score
        out.append(fila)
    return out


def render_conciliacion_ae(
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    st.caption(
        "Formato AE-Studio (cliente → extracto → Ingresos / Egresos / Retenciones / "
        "Deducciones / Inter-cuentas). La cuenta la arma la web: **Conceptos Bancos**, "
        "reglas locales, padrón de proveedores/VEPs y el **plan de cuentas de esta sociedad**. "
        "OCR siempre, también si el PDF es escaneo."
    )

    periodo_key = f"ae_periodo_{sociedad_id}"
    if periodo_key not in st.session_state:
        st.session_state[periodo_key] = date.today().replace(day=1)
    preview_key = f"ae_preview_{sociedad_id}"

    tab_imp, tab_an, tab_mov, tab_ing, tab_egr, tab_ret, tab_ded, tab_ic, tab_as, tab_reg = st.tabs(
        [
            "Importar extracto",
            "Análisis",
            "Movimientos",
            "Ingresos",
            "Egresos",
            "Retenciones",
            "Deducciones",
            "Inter-cuentas",
            "Asiento Tango",
            "Reglas",
        ]
    )

    with tab_imp:
        c1, c2, c3 = st.columns(3)
        with c1:
            archivos = st.file_uploader(
                "Extracto (PDF, Excel o CSV)",
                type=["pdf", "xlsx", "xls", "csv"],
                accept_multiple_files=True,
                key=f"ae_pdfs_{sociedad_id}",
            )
        with c2:
            excel_prov = st.file_uploader(
                "Proveedores pendientes (Excel Tango)",
                type=["xlsx", "xls", "csv"],
                key=f"ae_prov_{sociedad_id}",
            )
        with c3:
            excel_vep = st.file_uploader(
                "Padrón VEPs AFIP (Excel)",
                type=["xlsx", "xls", "csv"],
                key=f"ae_vep_{sociedad_id}",
            )
        periodo_ui = st.date_input(
            "Período (mes del extracto)",
            value=st.session_state[periodo_key],
            key=f"ae_periodo_ui_{sociedad_id}",
        )
        st.session_state[periodo_key] = periodo_ui.replace(day=1)
        periodo = _periodo_str(st.session_state[periodo_key])

        if not plan_vinculado:
            st.error("Vinculá el plan de cuentas de esta sociedad antes de clasificar.")

        if st.button("Leer y clasificar", type="primary", key=f"ae_run_{sociedad_id}"):
            if not archivos:
                st.error("Subí al menos un extracto.")
            elif not plan_vinculado:
                st.error("Falta el plan de cuentas de la sociedad.")
            else:
                with st.spinner("Leyendo extracto (OCR si es escaneo) con las reglas de la web…"):
                    df, meta, errores = procesar_extractos_bancarios_pdfs(archivos)
                    if errores:
                        st.warning(
                            "Algunos archivos tuvieron problemas: "
                            + "; ".join(
                                f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5]
                            )
                        )
                    if df is None or getattr(df, "empty", True):
                        st.error("No salieron movimientos del extracto.")
                    else:
                        if excel_prov is not None:
                            n_p = db.reemplazar_proveedores_pendientes(
                                sociedad_id, _cargar_proveedores_desde_upload(excel_prov)
                            )
                            st.info(f"Proveedores cargados: {n_p}")
                        if excel_vep is not None:
                            raw = excel_vep.read() if hasattr(excel_vep, "read") else excel_vep
                            n_v = db.reemplazar_veps_afip(
                                sociedad_id,
                                _cargar_veps_excel(BytesIO(raw) if isinstance(raw, bytes) else excel_vep),
                            )
                            st.info(f"VEPs cargados: {n_v}")

                        filas = df_extracto_a_filas(df)
                        ok_saldo, msg_saldo = validar_saldos_corridos(filas)
                        if not ok_saldo:
                            st.error(f"Extracto sospechoso: {msg_saldo}")
                        reglas = db.listar_reglas_clasificacion(solo_activas=True)
                        proveedores = db.listar_proveedores_pendientes(sociedad_id, solo_libres=False)
                        veps = db.listar_veps_afip(sociedad_id)
                        banco = (
                            str((meta or {}).get("banco") or "")
                            or banco_elegido
                            or ""
                        )
                        resultados = correr_motor(
                            filas,
                            reglas,
                            proveedores,
                            veps,
                            cliente_id=sociedad_id,
                            banco=banco,
                            periodo=st.session_state[periodo_key],
                            saldo_ok=ok_saldo,
                        )
                        plan_df = st.session_state.get("plan_cuentas_df")
                        movs = _aplicar_plan(_enriquecer(resultados), plan_df)
                        st.session_state[preview_key] = {
                            "movimientos": movs,
                            "proveedores": proveedores,
                            "banco": banco,
                            "periodo": periodo,
                            "meta": meta or {},
                        }
                        st.rerun()

        preview = st.session_state.get(preview_key)
        if preview:
            movs = preview.get("movimientos") or []
            st.success(
                f"Propuesta lista: **{len(movs)}** movimientos · banco **{preview.get('banco') or '—'}**. "
                "Todavía no se guardó. Revisá Análisis / Movimientos y confirmá."
            )
            reviso = st.checkbox(
                "Revisé las pendientes. El motor propone; yo confirmo.",
                key=f"ae_confirm_rev_{sociedad_id}",
            )
            if st.button("Confirmar y guardar", type="primary", key=f"ae_save_{sociedad_id}"):
                if not reviso:
                    st.error("Tildá que revisaste antes de guardar.")
                else:
                    db.borrar_movimientos_periodo(
                        sociedad_id, periodo=str(preview.get("periodo") or periodo), banco=None
                    )
                    db.insertar_movimientos_banco(preview.get("movimientos") or [])
                    st.session_state["motor_last_periodo"] = str(preview.get("periodo") or periodo)
                    st.session_state["motor_grilla_bridge"] = movimientos_a_filas_grilla_tango(
                        preview.get("movimientos") or [],
                        plan_cuentas=st.session_state.get("plan_cuentas_df"),
                    )
                    usuario = str(st.session_state.get("oficina_usuario") or "sistema")
                    db.registrar_auditoria_conciliacion(
                        cliente_id=sociedad_id,
                        movimiento_id=None,
                        usuario=usuario,
                        accion="correr_motor_ae",
                        detalle=f"{len(movs)} movimientos | banco={preview.get('banco')}",
                    )
                    st.session_state.pop(preview_key, None)
                    st.success("Guardado.")
                    st.rerun()
            if st.button("Descartar propuesta", key=f"ae_discard_{sociedad_id}"):
                st.session_state.pop(preview_key, None)
                st.rerun()
        else:
            st.info("Subí el extracto de esta sociedad y dale a **Leer y clasificar**.")

    preview = st.session_state.get(preview_key)
    if preview:
        movs_vista = preview.get("movimientos") or []
    else:
        periodo_activo = st.session_state.get("motor_last_periodo") or _periodo_str(
            st.session_state.get(periodo_key)
        )
        guardados = db.listar_movimientos_banco(sociedad_id, periodo=periodo_activo)
        plan_df = st.session_state.get("plan_cuentas_df")
        movs_vista = _aplicar_plan(_enriquecer(guardados), plan_df)

    banco_filtro = str(banco_elegido or "").strip()
    if banco_filtro:
        filtrados = [
            m
            for m in movs_vista
            if not m.get("banco")
            or banco_filtro.lower() in str(m.get("banco") or "").lower()
        ]
        if filtrados:
            movs_vista = filtrados

    with tab_an:
        if not movs_vista:
            st.info("Todavía no hay movimientos. Importá un extracto.")
        else:
            _kpis(movs_vista)
            st.caption(
                f"Sociedad: **{nombre_activo or '—'}** · CUIT `{cuit_activo or '—'}` · "
                f"Banco filtro: **{banco_elegido or 'todos'}**"
            )

    with tab_mov:
        _tabla(movs_vista)

    for key, tab in (
        ("ingreso", tab_ing),
        ("egreso", tab_egr),
        ("retencion", tab_ret),
        ("deduccion", tab_ded),
        ("inter-cta", tab_ic),
    ):
        with tab:
            _tabla([m for m in movs_vista if m.get("vista") == key])

    with tab_as:
        st.caption(
            "Renglones con el código del plan de esta sociedad (99999 = no hubo match, no se inventa). "
            "La plantilla vacía de importación Tango, cuando la pases, sale de acá."
        )
        bridge = movimientos_a_filas_grilla_tango(
            [m for m in movs_vista if m.get("estado") != "PENDIENTE"],
            plan_cuentas=st.session_state.get("plan_cuentas_df"),
        )
        if not bridge:
            st.info("No hay filas confirmadas para armar asiento.")
        else:
            df_as = pd.DataFrame(bridge)
            debe = float(df_as["debe"].sum()) if "debe" in df_as.columns else 0.0
            haber = float(df_as["haber"].sum()) if "haber" in df_as.columns else 0.0
            k1, k2, k3 = st.columns(3)
            k1.metric("Debe", f"$ {debe:,.2f}")
            k2.metric("Haber", f"$ {haber:,.2f}")
            k3.metric("Diferencia", f"$ {debe - haber:,.2f}")
            st.dataframe(df_as, use_container_width=True, hide_index=True)
            st.session_state["motor_grilla_bridge"] = bridge

    with tab_reg:
        st.caption("Las mismas reglas que ya usa la web: instructivo + reglas locales + VEPs.")
        try:
            info = cargar_instructivo()
            st.success(
                f"Conceptos Bancos: **{len(info.get('cuentas') or [])} cuentas** · "
                f"**{len(info.get('bancos') or {})} bancos**. "
                f"Origen: `{info.get('origen') or CACHE_PATH}`"
            )
        except Exception as exc:
            st.warning(f"Sin instructivo: {exc}")
        up_cb = st.file_uploader(
            "Actualizar Conceptos Bancos 2.xlsx",
            type=["xlsx"],
            key=f"ae_uploader_conceptos_{sociedad_id}",
        )
        if up_cb is not None:
            data = cargar_instructivo(up_cb.getvalue())
            st.success(
                f"Cache actualizado: {len(data.get('cuentas') or [])} cuentas, "
                f"{len(data.get('bancos') or {})} bancos."
            )
        reglas = db.listar_reglas_clasificacion(solo_activas=False)
        st.dataframe(
            pd.DataFrame(reglas)[["id", "orden", "patron", "categoria", "tipo", "activo"]]
            if reglas
            else pd.DataFrame(),
            use_container_width=True,
            hide_index=True,
        )
        with st.form(f"ae_alta_regla_{sociedad_id}"):
            p = st.text_input("Patrón")
            c = st.text_input("Categoría")
            t = st.selectbox(
                "Tipo",
                [
                    "INGRESO",
                    "DEBITO_IMPUESTO",
                    "DEBITO_FIJO",
                    "DEBITO_VEP",
                    "DEBITO_PROVEEDOR",
                    "DEBITO_REVISAR",
                    "INGRESO_O_DEBITO_PROPIO",
                ],
            )
            if st.form_submit_button("Agregar regla") and p and c:
                db.agregar_regla_clasificacion(p, c, t)
                st.success("Regla agregada.")
                st.rerun()
