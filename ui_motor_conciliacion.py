# -*- coding: utf-8 -*-
"""UI Streamlit del motor de conciliación bancaria."""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st

import database as db
from motor_conciliacion import (
    CATEGORIA_A_CUENTA_HINT,
    correr_motor,
    df_extracto_a_filas,
    money,
    movimientos_a_filas_grilla_tango,
    resumen_por_categoria,
    validar_saldos_corridos,
)


def _periodo_str(d: date | None) -> str:
    if not d:
        return date.today().strftime("%Y-%m-01")
    return d.replace(day=1).isoformat()


def _cargar_veps_excel(fuente) -> list[dict]:
    df = pd.read_excel(fuente) if not isinstance(fuente, pd.DataFrame) else fuente
    cols = {str(c).strip().lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            for k, orig in cols.items():
                if n in k:
                    return orig
        return None

    c_nro = col("vep", "numero")
    c_fecha = col("fecha")
    c_imp = col("importe", "monto", "total")
    c_imp_uesto = col("impuesto", "concepto", "tributo")
    c_per = col("periodo", "fiscal")
    out = []
    for _, row in df.iterrows():
        out.append(
            {
                "numero_vep": str(row[c_nro]).strip() if c_nro else "",
                "fecha": row[c_fecha] if c_fecha else None,
                "importe": str(money(row[c_imp] if c_imp else 0)),
                "impuesto": str(row[c_imp_uesto]).strip() if c_imp_uesto else "",
                "periodo_fiscal": str(row[c_per]).strip() if c_per else "",
            }
        )
    return out


def _cargar_proveedores_desde_upload(fuente) -> list[dict]:
    from procesador import cargar_facturas_proveedores_excel

    df = cargar_facturas_proveedores_excel(fuente)
    filas = []
    for _, row in df.iterrows():
        filas.append(
            {
                "fecha": row.get("fecha"),
                "tipo_comp": row.get("tipo") or "",
                "num_comp": row.get("comprobante") or "",
                "razon_social": row.get("proveedor") or "",
                "importe": str(money(row.get("importe"))),
                "usado": False,
            }
        )
    return filas


def render_motor_conciliacion(
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
) -> None:
    st.markdown("---")
    st.subheader("Motor de Conciliación (extracto → reglas → match)")
    st.caption(
        "Clasifica el extracto PDF con reglas editables, matchea proveedores y VEPs, "
        "y deja excepciones para revisión. No reemplaza el flujo Balance→Tango; lo complementa."
    )

    periodo_key = f"motor_periodo_{sociedad_id}"
    if periodo_key not in st.session_state:
        st.session_state[periodo_key] = date.today().replace(day=1)

    tab_imp, tab_res, tab_exc, tab_sum, tab_cfg = st.tabs(
        ["1. Importar y correr", "2. Conciliación", "3. Excepciones", "4. Resumen", "5. Reglas / VEPs"]
    )

    with tab_imp:
        c1, c2, c3 = st.columns(3)
        with c1:
            pdfs = st.file_uploader(
                "Extracto bancario (PDF)",
                type=["pdf"],
                accept_multiple_files=True,
                key=f"motor_pdfs_{sociedad_id}",
            )
        with c2:
            excel_prov = st.file_uploader(
                "Proveedores pendientes (Excel Tango 21101)",
                type=["xlsx", "xls", "csv"],
                key=f"motor_prov_{sociedad_id}",
            )
        with c3:
            excel_vep = st.file_uploader(
                "Padrón VEPs AFIP (Excel)",
                type=["xlsx", "xls", "csv"],
                key=f"motor_vep_{sociedad_id}",
            )

        periodo_ui = st.date_input(
            "Período (mes del extracto)",
            value=st.session_state[periodo_key],
            key=f"motor_periodo_ui_{sociedad_id}",
        )
        st.session_state[periodo_key] = periodo_ui.replace(day=1)
        periodo = _periodo_str(st.session_state[periodo_key])

        if st.button("Correr motor de conciliación", type="primary", key=f"motor_run_{sociedad_id}"):
            if not pdfs:
                st.error("Subí al menos un PDF de extracto.")
            else:
                with st.spinner("Parseando extracto y corriendo motor..."):
                    from procesador import procesar_extractos_bancarios_pdfs

                    df, meta, errores = procesar_extractos_bancarios_pdfs(pdfs)
                    if errores:
                        st.warning("Algunos PDF tuvieron problemas: " + "; ".join(
                            f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5]
                        ))
                    if df is None or df.empty:
                        st.error("No se extrajeron movimientos del PDF.")
                    else:
                        filas = df_extracto_a_filas(df)
                        ok_saldo, msg_saldo = validar_saldos_corridos(filas)
                        if not ok_saldo:
                            st.error(f"Extracto sospechoso: {msg_saldo}")

                        if excel_prov is not None:
                            n_p = db.reemplazar_proveedores_pendientes(
                                sociedad_id, _cargar_proveedores_desde_upload(excel_prov)
                            )
                            st.info(f"Proveedores cargados: {n_p}")
                        if excel_vep is not None:
                            raw = excel_vep.read() if hasattr(excel_vep, "read") else excel_vep
                            if isinstance(raw, bytes):
                                n_v = db.reemplazar_veps_afip(
                                    sociedad_id, _cargar_veps_excel(BytesIO(raw))
                                )
                            else:
                                n_v = db.reemplazar_veps_afip(
                                    sociedad_id, _cargar_veps_excel(excel_vep)
                                )
                            st.info(f"VEPs cargados: {n_v}")

                        reglas = db.listar_reglas_clasificacion(solo_activas=True)
                        proveedores = db.listar_proveedores_pendientes(sociedad_id, solo_libres=False)
                        veps = db.listar_veps_afip(sociedad_id)
                        banco = banco_elegido or (meta.get("banco") if isinstance(meta, dict) else "") or ""

                        resultados = correr_motor(
                            filas,
                            reglas,
                            proveedores,
                            veps,
                            cliente_id=sociedad_id,
                            banco=str(banco),
                            periodo=st.session_state[periodo_key],
                            saldo_ok=ok_saldo,
                        )
                        db.borrar_movimientos_periodo(sociedad_id, periodo=periodo, banco=None)
                        db.insertar_movimientos_banco(resultados)
                        # Persistir usados
                        for p in proveedores:
                            if p.get("usado") and p.get("id"):
                                db.marcar_proveedor_usado(int(p["id"]), True)

                        usuario = str(st.session_state.get("oficina_usuario") or "sistema")
                        db.registrar_auditoria_conciliacion(
                            cliente_id=sociedad_id,
                            movimiento_id=None,
                            usuario=usuario,
                            accion="correr_motor",
                            detalle=f"{len(resultados)} movimientos | banco={banco} | periodo={periodo}",
                        )
                        st.session_state["motor_last_periodo"] = periodo
                        st.session_state["motor_grilla_bridge"] = movimientos_a_filas_grilla_tango(
                            resultados
                        )
                        st.success(
                            f"Motor OK: {len(resultados)} movimientos. "
                            f"Conciliados: {sum(1 for r in resultados if r['estado']=='CONCILIADO')} · "
                            f"Pendientes: {sum(1 for r in resultados if r['estado']=='PENDIENTE')}"
                        )

    periodo_activo = st.session_state.get("motor_last_periodo") or _periodo_str(
        st.session_state.get(periodo_key)
    )
    movs = db.listar_movimientos_banco(sociedad_id, periodo=periodo_activo)

    with tab_res:
        if not movs:
            st.info("Todavía no hay movimientos para este período. Corré el motor en la pestaña 1.")
        else:
            f1, f2 = st.columns(2)
            with f1:
                est_filtro = st.multiselect(
                    "Estado",
                    ["OK", "CONCILIADO", "PENDIENTE"],
                    default=["OK", "CONCILIADO", "PENDIENTE"],
                    key=f"motor_filtro_est_{sociedad_id}",
                )
            with f2:
                cats = sorted({m.get("categoria") or "" for m in movs})
                cat_filtro = st.multiselect(
                    "Categoría",
                    cats,
                    default=cats,
                    key=f"motor_filtro_cat_{sociedad_id}",
                )
            data = [
                m
                for m in movs
                if m.get("estado") in est_filtro and (m.get("categoria") or "") in cat_filtro
            ]
            df_show = pd.DataFrame(
                [
                    {
                        "Fecha": m.get("fecha"),
                        "Descripción": m.get("descripcion"),
                        "Crédito": float(money(m.get("credito"))),
                        "Débito": float(money(m.get("debito"))),
                        "Categoría": m.get("categoria"),
                        "Estado": m.get("estado"),
                        "Match": m.get("match_detalle") or "",
                    }
                    for m in data
                ]
            )

            if not df_show.empty:
                st.dataframe(df_show, use_container_width=True, hide_index=True)
            else:
                st.warning("Sin filas con esos filtros.")
                # Colores: verde CONCILIADO / celeste OK / naranja PENDIENTE (leyenda)
            st.caption("Estados: CONCILIADO (verde en Excel export) · OK · PENDIENTE (excepciones)")

            if st.session_state.get("motor_grilla_bridge"):
                st.markdown("#### Puente a grilla Tango")
                st.caption(
                    "Filas no pendientes con cuenta sugerida (99999 = sin mapeo). "
                    "Usá el flujo Balance→grilla para el asiento final; esto es el sugerido del motor."
                )
                st.dataframe(
                    pd.DataFrame(st.session_state["motor_grilla_bridge"]),
                    use_container_width=True,
                    hide_index=True,
                )
                st.download_button(
                    "Descargar sugerido motor (CSV)",
                    data=pd.DataFrame(st.session_state["motor_grilla_bridge"]).to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"motor_conciliacion_{nombre_activo or sociedad_id}_{periodo_activo}.csv",
                    mime="text/csv",
                    key=f"dl_bridge_{sociedad_id}",
                )

    with tab_exc:
        pendientes = [m for m in movs if m.get("estado") == "PENDIENTE"]
        if not pendientes:
            st.success("No hay excepciones pendientes en este período.")
        else:
            st.write(f"**{len(pendientes)}** movimientos pendientes de revisión.")
            for m in pendientes:
                with st.expander(
                    f"{m.get('fecha')} · ${float(money(m.get('debito') or m.get('credito'))):,.2f} · "
                    f"{(m.get('descripcion') or '')[:80]}"
                ):
                    st.write(m.get("match_detalle") or "")
                    st.code(m.get("descripcion") or "", language=None)
                    nueva_cat = st.text_input(
                        "Categoría a asignar",
                        value=m.get("categoria") or "",
                        key=f"exc_cat_{m['id']}",
                    )
                    nuevo_tipo = st.selectbox(
                        "Tipo",
                        [
                            "INGRESO",
                            "DEBITO_IMPUESTO",
                            "DEBITO_FIJO",
                            "DEBITO_VEP",
                            "DEBITO_PROVEEDOR",
                            "DEBITO_REVISAR",
                        ],
                        index=5,
                        key=f"exc_tipo_{m['id']}",
                    )
                    patron_nuevo = st.text_input(
                        "Patrón a guardar en reglas (opcional)",
                        value="",
                        key=f"exc_pat_{m['id']}",
                        help="Si lo completás, se agrega a clasificacion_reglas para el próximo mes.",
                    )
                    if st.button("Resolver excepción", key=f"exc_ok_{m['id']}"):
                        ant = m.get("categoria")
                        db.actualizar_movimiento_banco(
                            int(m["id"]),
                            categoria=nueva_cat,
                            tipo=nuevo_tipo,
                            estado="OK",
                            match_detalle="Resuelto manualmente",
                        )
                        usuario = str(st.session_state.get("oficina_usuario") or "sistema")
                        db.registrar_auditoria_conciliacion(
                            cliente_id=sociedad_id,
                            movimiento_id=int(m["id"]),
                            usuario=usuario,
                            accion="resolver_excepcion",
                            categoria_anterior=ant,
                            categoria_nueva=nueva_cat,
                            detalle=m.get("descripcion"),
                        )
                        if patron_nuevo.strip():
                            db.agregar_regla_clasificacion(
                                patron_nuevo.strip(), nueva_cat, nuevo_tipo
                            )
                            st.success("Excepción resuelta y patrón agregado a reglas.")
                        else:
                            st.success("Excepción resuelta.")
                        st.rerun()

    with tab_sum:
        if not movs:
            st.info("Sin datos.")
        else:
            resumen = resumen_por_categoria(movs)
            st.dataframe(resumen, use_container_width=True, hide_index=True)
            tot_c = sum(float(money(m.get("credito"))) for m in movs)
            tot_d = sum(float(money(m.get("debito"))) for m in movs)
            st.metric("Total créditos extracto", f"$ {tot_c:,.2f}")
            st.metric("Total débitos extracto", f"$ {tot_d:,.2f}")

    with tab_cfg:
        st.markdown("##### Reglas de clasificación (editables)")
        reglas = db.listar_reglas_clasificacion(solo_activas=False)
        st.dataframe(
            pd.DataFrame(reglas)[
                ["id", "orden", "patron", "categoria", "tipo", "activo"]
            ]
            if reglas
            else pd.DataFrame(),
            use_container_width=True,
            hide_index=True,
        )
        with st.form(f"alta_regla_{sociedad_id}"):
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
                ],
            )
            if st.form_submit_button("Agregar regla") and p and c:
                db.agregar_regla_clasificacion(p, c, t)
                st.success("Regla agregada.")
                st.rerun()

        st.markdown("##### Padrón VEPs cargado")
        veps = db.listar_veps_afip(sociedad_id)
        st.caption(f"{len(veps)} VEPs en base para esta sociedad.")
        if veps:
            st.dataframe(pd.DataFrame(veps), use_container_width=True, hide_index=True)

        st.markdown("##### Mapa categoría → cuenta (puente Tango)")
        st.dataframe(
            pd.DataFrame(
                [{"categoria": k, "cuenta_sugerida": v} for k, v in CATEGORIA_A_CUENTA_HINT.items()]
            ),
            use_container_width=True,
            hide_index=True,
        )
