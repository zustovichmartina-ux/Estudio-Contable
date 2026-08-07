"""UI Streamlit — Liquidación de Sueldos."""

from __future__ import annotations

import importlib
from datetime import date

import pandas as pd
import streamlit as st

import database as db
from cursor_error_report import render_boton_enviar_error_cursor
from importar_legajos_sueldos import (
    generar_plantilla_legajos_bytes,
    leer_legajos_desde_excel,
)
from payroll_service import calcular_liquidacion, periodo_actual


def _recargar_database() -> None:
    """Evita AttributeError si Streamlit quedó con un database.py viejo en memoria."""
    global db
    db = importlib.reload(db)


def _fmt_money(v: float) -> str:
    try:
        return f"$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "$ 0,00"


def _etiqueta_cct(codigo: str | None, nombre: str | None = None) -> str:
    if not codigo:
        return "Sin CCT asignado"
    if nombre:
        return f"{codigo} — {nombre}"
    return str(codigo)


def _selector_sociedad(clientes: list[dict], key: str = "sueldos_sociedad_global") -> int | None:
    """Menú único de sociedades (buscable)."""
    if not clientes:
        st.warning("No hay sociedades cargadas. Cargalas en Gestión de Clientes.")
        return None

    labels = {f"{c['nombre']} ({c['cuit']})": int(c["id"]) for c in clientes}
    opciones = list(labels.keys())
    prev = st.session_state.get("sueldos_panel_cliente_id")
    index = 0
    if prev is not None:
        for i, lab in enumerate(opciones):
            if labels[lab] == int(prev):
                index = i
                break

    sel = st.selectbox(
        "Sociedad",
        opciones,
        index=min(index, len(opciones) - 1),
        key=key,
        help="Escribí para filtrar por nombre o CUIT.",
    )
    cliente_id = labels[sel]
    st.session_state["sueldos_panel_cliente_id"] = cliente_id
    return cliente_id


def _render_recibo(liq: dict) -> None:
    st.markdown(f"#### Recibo · {liq.get('empleado_nombre')} · {liq.get('periodo')}")
    st.caption(
        f"{liq.get('empresa_nombre')} · CUIT {liq.get('cuit')} · "
        f"CUIL {liq.get('cuil')} · {liq.get('categoria')} · CCT {liq.get('cct_asignado') or '—'}"
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total remunerativo", _fmt_money(liq.get("total_remunerativo", 0)))
    c2.metric("Total no remunerativo", _fmt_money(liq.get("total_no_remunerativo", 0)))
    c3.metric("Total descuentos", _fmt_money(liq.get("total_descuentos", 0)))
    c4.metric("Neto a percibir", _fmt_money(liq.get("neto_a_percibir", 0)))

    conceptos = liq.get("conceptos") or []
    if not conceptos:
        return
    df = pd.DataFrame(conceptos)
    if "importe" in df.columns:
        df["importe"] = df["importe"].map(_fmt_money)
    st.dataframe(df, use_container_width=True, hide_index=True)


def _vista_panel_estudio(periodo: str, cliente_id: int) -> None:
    st.caption(
        f"Período **{periodo}**. Estado: Pendiente → Novedades Recibidas → Liquidado."
    )
    filas = db.resumen_sueldos_empresas(periodo)
    f = next((x for x in filas if x["cliente_id"] == cliente_id), None)
    if not f:
        st.error("No se encontró la sociedad seleccionada.")
        return

    n_pend = sum(1 for x in filas if x["estado"] == "Pendiente")
    n_nov = sum(1 for x in filas if x["estado"] == "Novedades Recibidas")
    n_liq = sum(1 for x in filas if x["estado"] == "Liquidado")
    st.caption(
        f"Resumen global: {len(filas)} sociedades · 🟡 {n_pend} · 🟠 {n_nov} · 🟢 {n_liq}"
    )

    with st.container(border=True):
        col_a, col_b, col_c = st.columns([3, 2, 2])
        with col_a:
            st.markdown(f"**{f['nombre']}**")
            st.caption(f"CUIT {f['cuit']} · {f['empleados']} legajos")
        with col_b:
            if f["cct_asignado"]:
                st.markdown(f"`{f['cct_asignado']}`")
                st.caption(f["convenio_nombre"])
            else:
                st.warning("Sin CCT asignado")
                st.caption("Asignalo en la pestaña Legajos y CCT")
            badge = {
                "Pendiente": "🟡 Pendiente",
                "Novedades Recibidas": "🟠 Novedades Recibidas",
                "Liquidado": "🟢 Liquidado",
            }.get(f["estado"], f["estado"])
            st.write(badge)
            st.caption(f"{f['novedades']} nov. · {f['liquidaciones']} liq.")
        with col_c:
            sin_cct = not f["cct_asignado"]
            if st.button(
                "Ejecutar Liquidación",
                key=f"liq_run_{cliente_id}_{periodo}",
                disabled=(
                    f["estado"] == "Pendiente"
                    or f["empleados"] == 0
                    or sin_cct
                ),
                use_container_width=True,
            ):
                _ejecutar_liquidacion_empresa(cliente_id, periodo)
                st.rerun()
            if st.button(
                "Ver liquidaciones",
                key=f"liq_ver_{cliente_id}_{periodo}",
                use_container_width=True,
            ):
                st.session_state["sueldos_detalle_cliente"] = cliente_id

    if st.session_state.get("sueldos_detalle_cliente") == cliente_id:
        st.divider()
        st.subheader(f"Detalle · {f['nombre']}")
        if st.button("Cerrar detalle", key="sueldos_cerrar_detalle"):
            st.session_state.pop("sueldos_detalle_cliente", None)
            st.rerun()

        liqs = db.listar_liquidaciones_periodo(cliente_id, periodo)
        if not liqs:
            st.info(
                "Sin liquidaciones en este período. Ejecutá la liquidación o cargá novedades."
            )
            return

        opciones_recibo = {
            f"{l['empleado_nombre']} ({l['cuil']}) — neto {_fmt_money(l['neto_a_percibir'])}": l[
                "id"
            ]
            for l in liqs
        }
        sel = st.selectbox(
            "Previsualizar recibo",
            list(opciones_recibo.keys()),
            key="sueldos_sel_recibo",
        )
        liq = db.obtener_liquidacion(opciones_recibo[sel])
        if liq:
            _render_recibo(liq)


def _ejecutar_liquidacion_empresa(cliente_id: int, periodo: str) -> None:
    cliente = db.obtener_cliente(cliente_id)
    if not cliente:
        st.error("Cliente no encontrado.")
        return
    cct = (cliente.get("cct_asignado") or "").strip()
    if not cct:
        st.error("Asigná un CCT a esta sociedad antes de liquidar (pestaña Legajos y CCT).")
        return
    convenio = db.obtener_convenio(cct)
    if not convenio:
        st.error(f"No existe el convenio «{cct}» en el catálogo.")
        return
    reglas = convenio.get("reglas") or {}

    novedades = db.listar_novedades_periodo(cliente_id, periodo)
    if not novedades:
        st.warning(f"No hay novedades en el buzón para {periodo}.")
        return

    empleados = {e["id"]: e for e in db.listar_empleados_sueldos(cliente_id)}
    ok = 0
    sin_basico = []
    sin_categoria = []
    for nov in novedades:
        emp = empleados.get(nov["empleado_id"])
        if not emp:
            continue
        calc = calcular_liquidacion(emp, nov, reglas)
        if calc.get("fuente_basico") == "sin_dato":
            sin_categoria.append(emp.get("nombre") or emp.get("cuil"))
        if float(calc.get("sueldo_basico_usado") or 0) <= 0:
            sin_basico.append(emp.get("nombre") or emp.get("cuil"))
        db.upsert_liquidacion_resultado(
            cliente_id=cliente_id,
            empleado_id=emp["id"],
            periodo=periodo,
            total_remunerativo=calc["total_remunerativo"],
            total_no_remunerativo=calc["total_no_remunerativo"],
            total_descuentos=calc["total_descuentos"],
            neto_a_percibir=calc["neto_a_percibir"],
            conceptos=calc["conceptos"],
        )
        ok += 1
    st.success(f"Liquidación ejecutada: {ok} recibos para {periodo}.")
    if sin_categoria:
        st.warning(
            "Sin categoría CCT reconocida (básico en 0). Asigná categoría en Legajos y CCT: "
            + ", ".join(sin_categoria[:8])
            + ("…" if len(sin_categoria) > 8 else "")
        )
    elif sin_basico:
        st.warning(
            "Básico en 0: "
            + ", ".join(sin_basico[:8])
            + ("…" if len(sin_basico) > 8 else "")
        )


def _vista_novedades(periodo: str, cliente_id: int) -> None:
    empleados = db.listar_empleados_sueldos(cliente_id)
    if not empleados:
        st.warning("Esta empresa no tiene legajos. Importalos en la pestaña Legajos y CCT.")
        return

    st.caption(
        f"Período **{periodo}**. Editá ausencias, HE 50% y no remunerativo, luego enviá al estudio."
    )

    filas_edit = []
    for e in empleados:
        nov = db.obtener_novedad(cliente_id, e["id"], periodo)
        filas_edit.append(
            {
                "empleado_id": e["id"],
                "nombre": e["nombre"],
                "cuil": e["cuil"],
                "categoria": e["categoria"],
                "sueldo_basico": e["sueldo_basico"],
                "dias_ausencia": int(nov["dias_ausencia"]) if nov else 0,
                "horas_extras_50": float(nov["horas_extras_50"]) if nov else 0.0,
                "no_remunerativo_extra": float(nov["no_remunerativo_extra"]) if nov else 0.0,
            }
        )

    df = pd.DataFrame(filas_edit)
    editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        disabled=["empleado_id", "nombre", "cuil", "categoria", "sueldo_basico"],
        column_config={
            "empleado_id": None,
            "nombre": st.column_config.TextColumn("Empleado"),
            "cuil": st.column_config.TextColumn("CUIL"),
            "categoria": st.column_config.TextColumn("Categoría"),
            "sueldo_basico": st.column_config.NumberColumn("Básico", format="$ %.2f"),
            "dias_ausencia": st.column_config.NumberColumn("Días de Ausencia", min_value=0, step=1),
            "horas_extras_50": st.column_config.NumberColumn(
                "Horas Extras 50%", min_value=0.0, step=0.5, format="%.2f"
            ),
            "no_remunerativo_extra": st.column_config.NumberColumn(
                "Monto No Remunerativo", min_value=0.0, step=100.0, format="$ %.2f"
            ),
        },
        key=f"sueldos_editor_nov_{cliente_id}_{periodo}",
    )

    if st.button(
        "Guardar Novedades y Enviar al Estudio",
        type="primary",
        key=f"sueldos_enviar_nov_{cliente_id}",
    ):
        n = 0
        for _, row in editado.iterrows():
            db.upsert_novedad_buzon(
                cliente_id=cliente_id,
                empleado_id=int(row["empleado_id"]),
                periodo=periodo,
                dias_ausencia=int(row["dias_ausencia"] or 0),
                horas_extras_50=float(row["horas_extras_50"] or 0),
                no_remunerativo_extra=float(row["no_remunerativo_extra"] or 0),
            )
            n += 1
        st.success(f"{n} novedades guardadas y enviadas al estudio ({periodo}).")
        st.rerun()


def _vista_legajos(cliente_id: int) -> None:
    cliente = db.obtener_cliente(cliente_id)
    if not cliente:
        st.error("Sociedad no encontrada.")
        return

    convenios = db.listar_convenios()
    codigos = [c["codigo"] for c in convenios]
    nombres = {c["codigo"]: c["nombre"] for c in convenios}
    opciones_cct = ["— Sin asignar —"] + [
        f"{c} — {nombres.get(c, c)}" for c in codigos
    ]
    cct_actual = (cliente.get("cct_asignado") or "").strip()
    idx = 0
    if cct_actual:
        for i, c in enumerate(codigos):
            if c == cct_actual:
                idx = i + 1
                break

    st.subheader("Convenio colectivo (CCT)")
    st.caption(
        "El **sueldo básico** se toma de la escala del CCT según la categoría del legajo "
        "(FAECYS julio 2026 para Comercio). Antigüedad Comercio = 1% por año."
    )
    nuevo_label = st.selectbox(
        "CCT asignado",
        opciones_cct,
        index=idx,
        key=f"sueldos_cct_sel_{cliente_id}",
    )
    if st.button("Guardar CCT", key=f"sueldos_guardar_cct_{cliente_id}"):
        if nuevo_label.startswith("—"):
            db.actualizar_cct_cliente(cliente_id, "")
            st.success("CCT quitado (sin asignar).")
        else:
            codigo = nuevo_label.split(" — ", 1)[0].strip()
            db.actualizar_cct_cliente(cliente_id, codigo)
            st.success(f"CCT {codigo} asignado a {cliente['nombre']}.")
        st.rerun()

    # Escala del CCT asignado
    cct_ver = (cliente.get("cct_asignado") or "").strip()
    if cct_ver:
        from cct_escalas import obtener_escala_desde_reglas, listar_categorias_cct

        conv = db.obtener_convenio(cct_ver)
        reglas = (conv or {}).get("reglas") or {}
        escala = obtener_escala_desde_reglas(reglas)
        with st.expander(
            f"Escala salarial CCT ({reglas.get('escalaVigencia', 'vigente')}) — "
            f"{len(escala)} categorías",
            expanded=False,
        ):
            if reglas.get("escalaFuente"):
                st.caption(reglas["escalaFuente"])
            if escala:
                st.dataframe(
                    pd.DataFrame(
                        [{"Categoría": k, "Básico": v} for k, v in sorted(escala.items())]
                    ),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Básico": st.column_config.NumberColumn(format="$ %.2f"),
                    },
                )
            else:
                st.info("Este CCT aún no tiene escala de básicos cargada.")

            cats = listar_categorias_cct(reglas)
            if cats and st.button(
                "Aplicar básicos de la escala a legajos (según categoría)",
                key=f"sueldos_aplicar_escala_{cliente_id}",
            ):
                from cct_escalas import resolver_basico_cct

                n_ok, n_fail = 0, 0
                for e in db.listar_empleados_sueldos(cliente_id):
                    basico, canon = resolver_basico_cct(e.get("categoria") or "", reglas)
                    if basico is None:
                        n_fail += 1
                        continue
                    db.upsert_empleado_sueldo(
                        cliente_id=cliente_id,
                        cuil=e["cuil"],
                        nombre=e["nombre"],
                        categoria=canon or e["categoria"],
                        sueldo_basico=basico,
                        fecha_ingreso=e["fecha_ingreso"],
                        antiguedad_anios=e["antiguedad_anios"],
                        activo=bool(e["activo"]),
                        empleado_id=e["id"],
                    )
                    n_ok += 1
                st.success(
                    f"Básicos aplicados: {n_ok}. Sin categoría reconocida: {n_fail}."
                )
                st.rerun()

    st.divider()
    st.subheader("Importar legajos desde Excel / CSV (Tango)")
    st.caption(
        "1) Descargá la plantilla · 2) Pegá el export de Tango · 3) Subí el archivo · "
        "4) Confirmá la importación. Los CUIL existentes se actualizan."
    )

    c_dl, c_up = st.columns(2)
    with c_dl:
        st.download_button(
            "⬇️ Descargar plantilla Excel",
            data=generar_plantilla_legajos_bytes(),
            file_name="Plantilla_Legajos_Sueldos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"sueldos_dl_plantilla_{cliente_id}",
            use_container_width=True,
        )
    with c_up:
        uploaded = st.file_uploader(
            "Subir Excel o CSV de legajos",
            type=["xlsx", "xls", "csv"],
            key="sueldos_upload_legajos_fijo",
            help="Formatos: .xlsx, .xls o .csv",
        )

    if uploaded is not None:
        try:
            raw = uploaded.getvalue()
            # Pasar nombre para detectar CSV
            from io import BytesIO

            bio = BytesIO(raw)
            bio.name = uploaded.name
            filas, errores = leer_legajos_desde_excel(bio)
        except Exception as exc:
            st.error(f"Error al leer el archivo: {exc}")
            filas, errores = [], [str(exc)]

        if errores and not filas:
            for e in errores[:20]:
                st.error(e)
        else:
            st.write(f"**Vista previa:** {len(filas)} legajos listos para importar.")
            if errores:
                with st.expander(f"Avisos ({len(errores)})", expanded=False):
                    for e in errores[:40]:
                        st.warning(e)
            if filas:
                st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
                if st.button(
                    f"Importar {len(filas)} legajos a {cliente['nombre']}",
                    type="primary",
                    key=f"sueldos_btn_import_{cliente_id}",
                ):
                    n = 0
                    for row in filas:
                        db.upsert_empleado_sueldo(
                            cliente_id=cliente_id,
                            cuil=row["cuil"],
                            nombre=row["nombre"],
                            categoria=row["categoria"],
                            sueldo_basico=row["sueldo_basico"],
                            fecha_ingreso=row["fecha_ingreso"],
                            antiguedad_anios=row["antiguedad_anios"],
                            activo=True,
                        )
                        n += 1
                    st.success(f"Se importaron / actualizaron {n} legajos.")
                    st.rerun()

    st.divider()
    st.subheader("Alta / edición manual de un legajo")
    with st.form(f"form_empleado_sueldo_{cliente_id}", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            cuil = st.text_input("CUIL")
            nombre = st.text_input("Nombre")
            categoria = st.text_input("Categoría")
        with c2:
            sueldo = st.number_input("Sueldo básico", min_value=0.0, step=1000.0)
            fecha_ing = st.date_input("Fecha de ingreso", value=date(2020, 1, 1))
            antig = st.number_input("Antigüedad (años)", min_value=0, step=1)
        if st.form_submit_button("Guardar legajo"):
            if not cuil.strip() or not nombre.strip():
                st.error("CUIL y nombre son obligatorios.")
            else:
                db.upsert_empleado_sueldo(
                    cliente_id=cliente_id,
                    cuil=cuil.strip(),
                    nombre=nombre.strip(),
                    categoria=categoria.strip(),
                    sueldo_basico=float(sueldo),
                    fecha_ingreso=fecha_ing.isoformat(),
                    antiguedad_anios=int(antig),
                )
                st.success("Legajo guardado.")
                st.rerun()

    empleados = db.listar_empleados_sueldos(cliente_id, solo_activos=False)
    st.subheader(f"Legajos cargados ({len(empleados)})")
    if empleados:
        from cct_escalas import listar_categorias_cct, resolver_basico_cct

        conv = db.obtener_convenio((cliente.get("cct_asignado") or "").strip())
        reglas = (conv or {}).get("reglas") or {}
        cats_cct = listar_categorias_cct(reglas) if reglas.get("escalas") or conv else listar_categorias_cct(None)

        st.caption(
            "Asigná la **categoría del CCT** (ej. Administrativo A, Vendedor B). "
            "Al liquidar, el básico sale de la escala FAECYS, no hace falta tipearlo."
        )
        df_leg = pd.DataFrame(
            [
                {
                    "id": e["id"],
                    "nombre": e["nombre"],
                    "cuil": e["cuil"],
                    "categoria": e["categoria"] or "",
                    "basico_cct": (
                        resolver_basico_cct(e.get("categoria") or "", reglas)[0] or 0.0
                    ),
                    "fecha_ingreso": e["fecha_ingreso"],
                    "antiguedad_anios": int(e["antiguedad_anios"] or 0),
                    "activo": bool(e["activo"]),
                }
                for e in empleados
            ]
        )
        editado = st.data_editor(
            df_leg,
            use_container_width=True,
            hide_index=True,
            disabled=["id", "cuil", "basico_cct"],
            column_config={
                "id": None,
                "nombre": st.column_config.TextColumn("Nombre"),
                "cuil": st.column_config.TextColumn("CUIL"),
                "categoria": st.column_config.SelectboxColumn(
                    "Categoría CCT",
                    options=[""] + cats_cct,
                    required=False,
                ),
                "basico_cct": st.column_config.NumberColumn(
                    "Básico según CCT", format="$ %.2f",
                    help="Se calcula solo; no se edita a mano.",
                ),
                "fecha_ingreso": st.column_config.TextColumn("Ingreso (AAAA-MM-DD)"),
                "antiguedad_anios": st.column_config.NumberColumn(
                    "Antigüedad (años)", min_value=0, step=1
                ),
                "activo": st.column_config.CheckboxColumn("Activo"),
            },
            key=f"sueldos_editor_legajos_{cliente_id}",
        )
        if st.button(
            "Guardar cambios en legajos",
            type="primary",
            key=f"sueldos_guardar_legajos_grid_{cliente_id}",
        ):
            n = 0
            for _, row in editado.iterrows():
                cat = str(row.get("categoria") or "").strip()
                basico_res, canon = resolver_basico_cct(cat, reglas)
                db.upsert_empleado_sueldo(
                    cliente_id=cliente_id,
                    cuil=str(row["cuil"]).strip(),
                    nombre=str(row["nombre"]).strip(),
                    categoria=canon or cat,
                    sueldo_basico=float(basico_res or 0),
                    fecha_ingreso=str(row["fecha_ingreso"] or date.today().isoformat())[:10],
                    antiguedad_anios=int(row["antiguedad_anios"] or 0),
                    activo=bool(row["activo"]),
                    empleado_id=int(row["id"]),
                )
                n += 1
            st.success(
                f"{n} legajos actualizados. Volvé a Ejecutar Liquidación en el Panel."
            )
            st.rerun()

        baja_ids = {
            f"{e['nombre']} ({e['cuil']})": e["id"] for e in empleados if e["activo"]
        }
        if baja_ids:
            baja = st.selectbox(
                "Desactivar legajo", list(baja_ids.keys()), key=f"sueldos_baja_{cliente_id}"
            )
            if st.button("Desactivar", key=f"sueldos_btn_baja_{cliente_id}"):
                db.eliminar_empleado_sueldo(baja_ids[baja])
                st.success("Legajo desactivado.")
                st.rerun()
    else:
        st.info("Todavía no hay legajos. Subí el Excel/CSV de Tango arriba.")


def seccion_liquidacion_sueldos() -> None:
    """Pestaña principal de Liquidación de Sueldos."""
    try:
        _recargar_database()
        db.inicializar_bd()

        periodo = st.text_input(
            "Período (AAAA-MM)",
            value=st.session_state.get("sueldos_periodo") or periodo_actual(),
            key="sueldos_periodo_input",
        ).strip()
        if len(periodo) == 7 and periodo[4] == "-":
            st.session_state["sueldos_periodo"] = periodo
        else:
            st.warning("Usá formato AAAA-MM (ej. 2026-07).")
            periodo = periodo_actual()

        clientes = db.listar_clientes()
        st.markdown("##### Elegí la sociedad")
        cliente_id = _selector_sociedad(clientes)
        if cliente_id is None:
            return

        cli = db.obtener_cliente(cliente_id)
        cct = (cli or {}).get("cct_asignado") or ""
        st.caption(
            f"Seleccionada: **{(cli or {}).get('nombre', '')}** · "
            f"CCT: **{_etiqueta_cct(cct)}**"
        )

        tab_panel, tab_nov, tab_leg = st.tabs(
            ["Panel del estudio", "Carga de novedades", "Legajos y CCT"]
        )
        with tab_panel:
            _vista_panel_estudio(periodo, cliente_id)
        with tab_nov:
            _vista_novedades(periodo, cliente_id)
        with tab_leg:
            _vista_legajos(cliente_id)
    except Exception as exc:
        render_boton_enviar_error_cursor(
            exc,
            contexto="Liquidación de Sueldos",
            key="sueldos_enviar_error_cursor",
        )
