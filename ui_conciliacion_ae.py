# -*- coding: utf-8 -*-
"""Conciliación Bancaria — camino claro: importar → revisar → asiento Tango.

La clasificación es la de esta web (Conceptos Bancos + reglas de Configuración).
Los movimientos se guardan en la base del estudio.
"""
from __future__ import annotations

import calendar
import copy
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

import database as db
from capa_revision import gate_asiento, resolver_codigo_plan
from conceptos_bancos import CACHE_PATH, cargar_instructivo
from motor_conciliacion import (
    CATEGORIA_A_CUENTA_HINT,
    armar_papel_mes,
    bucket_ae,
    correr_motor,
    cuenta_banco_del_plan,
    df_extracto_a_filas,
    money,
    movimientos_a_filas_grilla_tango,
    papeles_por_mes,
    renglones_asiento_banco_mes,
    validar_saldos_corridos,
)
from excel_formato_estudio import exportar_informe_excel
from procesador import (
    AsientoDevengamiento,
    ExportacionTangoError,
    RenglonAsiento,
    clasificar_movimiento_extracto,
    generar_excel_tango_nativo,
    guardar_biblioteca_persistida,
    procesar_extractos_bancarios_pdfs,
)
from ui_motor_conciliacion import _cargar_proveedores_desde_upload, _cargar_veps_excel


class _UploadMemoria:
    """Bytes del PDF guardados entre reruns de Streamlit."""

    def __init__(self, name: str, data: bytes) -> None:
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data

_PASOS = (
    ("importar", "1 · Importar"),
    ("movimientos", "2 · Movimientos"),
    ("papeles", "3 · Papeles"),
    ("asiento", "4 · Asiento Tango"),
    ("config", "5 · Reglas"),
)
_VISTA_FILTRO = (
    ("", "Todas"),
    ("ingreso", "Ingresos"),
    ("egreso", "Egresos"),
    ("retencion", "Retenciones"),
    ("deduccion", "Deducciones"),
    ("inter-cta", "Inter-cuentas"),
    ("sin-cat", "Sin match"),
)


def _periodo_str(d: date | None) -> str:
    if not d:
        return date.today().strftime("%Y-%m-01")
    return d.replace(day=1).isoformat()


def _fmt(n: float | None) -> str:
    if n is None:
        return "—"
    return f"$ {n:,.2f}"


def _opciones_plan(plan_df: pd.DataFrame | None) -> list[str]:
    opts = ["99999 — A clasificar"]
    if plan_df is None or getattr(plan_df, "empty", True):
        return opts
    col_cod = "codigo" if "codigo" in plan_df.columns else plan_df.columns[0]
    col_desc = "descripcion" if "descripcion" in plan_df.columns else plan_df.columns[1]
    vistos: set[str] = set()
    for _, row in plan_df.iterrows():
        cod = str(row.get(col_cod) or "").strip()
        if not cod or cod in vistos or cod.lower() in {"nan", "codigo", "código"}:
            continue
        vistos.add(cod)
        desc = str(row.get(col_desc) or "").strip()
        opts.append(f"{cod} — {desc}" if desc else cod)
    return opts


def _codigo_desde_opcion(texto: str) -> str:
    s = str(texto or "").strip()
    if "—" in s:
        s = s.split("—", 1)[0].strip()
    elif " - " in s:
        s = s.split(" - ", 1)[0].strip()
    return s or "99999"


def _opcion_desde_codigo(codigo: str, opciones: list[str]) -> str:
    cod = str(codigo or "99999").strip() or "99999"
    for op in opciones:
        if _codigo_desde_opcion(op) == cod:
            return op
    return opciones[0] if opciones else "99999 — A clasificar"


def _asiento_desde_rows(
    rows: list[dict],
    *,
    banco: str,
    periodo: str,
    fecha_asiento: date,
) -> AsientoDevengamiento:
    renglones = [
        RenglonAsiento(
            codigo_cuenta=str(r.get("Código") or "99999"),
            descripcion_cuenta=str(r.get("Descripción") or ""),
            debe=float(r.get("Debe") or 0),
            haber=float(r.get("Haber") or 0),
        )
        for r in rows
        if abs(float(r.get("Debe") or 0)) > 0.005 or abs(float(r.get("Haber") or 0)) > 0.005
    ]
    asiento = AsientoDevengamiento(
        identificador=1,
        concepto=f"{banco} {periodo}".strip(),
        fecha=fecha_asiento,
        renglones=renglones,
    )
    asiento.tipo = banco  # type: ignore[attr-defined]
    asiento.periodo = periodo  # type: ignore[attr-defined]
    asiento.fecha_tango_str = fecha_asiento.strftime("%d/%m/%Y")  # type: ignore[attr-defined]
    return asiento


def _guardar_biblioteca_banco(
    sociedad_id: int,
    asientos: list,
    rows: list[dict],
    *,
    banco: str,
    periodo: str,
) -> str:
    etiqueta = str(periodo).replace("/", "-")
    existentes = st.session_state.get("biblioteca_bancos") or []
    if any(
        e.get("sociedad_id") == sociedad_id
        and e.get("banco") == banco
        and str(e.get("periodo") or "").replace("/", "-") == etiqueta
        for e in existentes
    ):
        raise ValueError(f"El período {etiqueta} ya está archivado para {banco}.")
    st.session_state.setdefault("biblioteca_bancos", []).append(
        {
            "sociedad_id": sociedad_id,
            "banco": banco,
            "periodo": periodo,
            "asientos": copy.deepcopy(asientos),
            "rows": copy.deepcopy(rows),
        }
    )
    st.session_state.setdefault("periodos_bancos_procesados", []).append(
        {"sociedad_id": sociedad_id, "banco": banco, "periodo": periodo}
    )
    guardar_biblioteca_persistida(
        biblioteca_asientos=st.session_state.get("biblioteca_asientos") or [],
        biblioteca_bancos=st.session_state.get("biblioteca_bancos") or [],
        periodos_procesados=st.session_state.get("periodos_procesados") or [],
        periodos_bancos_procesados=st.session_state.get("periodos_bancos_procesados") or [],
        usuario=str(st.session_state.get("oficina_usuario") or ""),
    )
    return etiqueta


def _norm(texto: str) -> str:
    return str(texto or "").lower()


def _filtra_banco(movs: list[dict], banco_elegido: str) -> list[dict]:
    banco_filtro = str(banco_elegido or "").strip().lower()
    if not banco_filtro:
        return movs
    tokens = [t for t in banco_filtro.replace("banco", " ").split() if len(t) > 2]
    filtrados = []
    for m in movs:
        b = str(m.get("banco") or "").lower()
        if not b:
            filtrados.append(m)
            continue
        if banco_filtro in b or b in banco_filtro or any(t in b for t in tokens):
            filtrados.append(m)
    return filtrados if filtrados else movs


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
        fila["sub_categoria"] = str(
            m.get("concepto_instructivo")
            or m.get("categoria")
            or (label if label and label != "Sin clasificar" else "-")
        )
        fila["monto"] = round(cred - deb, 2)
        out.append(fila)
    return out


def _aplicar_plan(movs: list[dict], plan_df) -> list[dict]:
    out = []
    for m in movs:
        fila = dict(m)
        codigo, desc_plan, score = resolver_codigo_plan(
            str(m.get("categoria") or ""),
            plan_df,
            hints=CATEGORIA_A_CUENTA_HINT,
        )
        if str(m.get("estado") or "") == "PENDIENTE" and (
            not fila.get("categoria")
            or "identificar" in str(fila.get("categoria") or "").lower()
            or codigo == "99999"
        ):
            codigo = "99999"
        fila["cuenta_codigo"] = codigo
        fila["cuenta_plan"] = desc_plan
        fila["score_plan"] = score
        out.append(fila)
    return out


def _grupo(movs: list[dict], vista: str) -> list[dict]:
    return [m for m in movs if m.get("vista") == vista]


def _suma(movs: list[dict], pred=None) -> float:
    tot = 0.0
    for m in movs:
        if pred is None or pred(m):
            tot += abs(float(m.get("monto") or 0))
    return tot


def _es_icdb(m: dict) -> bool:
    blob = _norm(f"{m.get('sub_categoria')} {m.get('descripcion')} {m.get('categoria')}")
    compacto = blob.replace(" ", "")
    return (
        "25413" in compacto
        or "icdb" in blob
        or "imp.ley" in compacto
        or ("debito" in blob and "credito" in blob and "ley" in blob)
    )


def _tabla(movs: list[dict], cols: list[str] | None = None) -> None:
    if not movs:
        st.info("Sin movimientos en esta vista. Andá a **Importar** y subí el extracto.")
        return
    filas = []
    for m in movs:
        filas.append(
            {
                "Fecha": m.get("fecha"),
                "Descripción": m.get("descripcion"),
                "Crédito": float(money(m.get("credito"))),
                "Débito": float(money(m.get("debito"))),
                "Monto": m.get("monto"),
                "Categoría": m.get("categoria") or m.get("sub_categoria"),
                "Vista": m.get("vista"),
                "Cuenta": m.get("cuenta_codigo") or "",
                "Plan": m.get("cuenta_plan") or "",
                "Estado": m.get("estado"),
                "Por qué": m.get("match_detalle") or "",
                "Banco": m.get("banco"),
            }
        )
    df = pd.DataFrame(filas)
    if cols:
        keep = [c for c in cols if c in df.columns]
        df = df[keep]
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"{len(movs)} movimiento(s). Categoría = reglas de esta web (Conceptos Bancos + Configuración).")


def _cargar_movs(sociedad_id: int, banco_elegido: str, periodo_key: str, preview_key: str) -> list[dict]:
    periodo_activo = _periodo_str(st.session_state.get(periodo_key))
    preview = st.session_state.get(preview_key)
    if preview and str(preview.get("periodo") or "") != periodo_activo:
        preview = None
    if preview:
        movs = list(preview.get("movimientos") or [])
    else:
        guardados = db.listar_movimientos_banco(sociedad_id, periodo=periodo_activo)
        plan_df = st.session_state.get("plan_cuentas_df")
        movs = _aplicar_plan(_enriquecer(guardados), plan_df)
    return _filtra_banco(movs, banco_elegido)


def _cargar_todos_banco(sociedad_id: int, banco_elegido: str, preview_key: str, periodo: str) -> list[dict]:
    todos = db.listar_movimientos_banco(sociedad_id)
    preview = st.session_state.get(preview_key)
    if preview:
        periodo_prev = str(preview.get("periodo") or periodo)
        todos = [m for m in todos if str(m.get("periodo") or "") != periodo_prev]
        todos.extend(preview.get("movimientos") or [])
    return _filtra_banco(_enriquecer(todos), banco_elegido)


def _kpis_resumen(movs: list[dict]) -> None:
    ing = _grupo(movs, "ingreso")
    egr = [m for m in movs if m.get("vista") in {"egreso", "impuesto"}]
    ret = _grupo(movs, "retencion")
    pend = [
        m
        for m in movs
        if str(m.get("estado") or "") == "PENDIENTE" or m.get("vista") == "sin-cat"
    ]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ingresos", _fmt(sum(float(m.get("monto") or 0) for m in ing if float(m.get("monto") or 0) > 0)), f"{len(ing)} mov.")
    c2.metric("Egresos", _fmt(_suma(egr)), f"{len(egr)} mov.")
    c3.metric("Retenciones", _fmt(_suma(ret)))
    c4.metric("A revisar", str(len(pend)))


def _bytes_informe(movs: list[dict], nombre: str | None, titulo: str) -> bytes:
    filas = []
    for m in movs:
        filas.append(
            {
                "Fecha": m.get("fecha"),
                "Descripción": m.get("descripcion"),
                "Crédito": float(money(m.get("credito"))),
                "Débito": float(money(m.get("debito"))),
                "Categoría": m.get("categoria"),
                "Cuenta": m.get("cuenta_codigo") or "",
                "Estado": m.get("estado"),
                "Banco": m.get("banco"),
            }
        )
    df = pd.DataFrame(filas)
    return exportar_informe_excel(
        titulo=titulo,
        subtitulo=nombre or "Cliente",
        periodo=_periodo_str(date.today()),
        kpis=[
            ("Movimientos", len(movs)),
            ("Ingresos", sum(float(m.get("monto") or 0) for m in movs if float(m.get("monto") or 0) > 0)),
            ("Egresos", sum(abs(float(m.get("monto") or 0)) for m in movs if float(m.get("monto") or 0) < 0)),
        ],
        detalle=df if not df.empty else pd.DataFrame(columns=["Fecha", "Descripción", "Crédito", "Débito"]),
        hoja_detalle="Movimientos",
        col_moneda=["Crédito", "Débito"],
        col_fecha=["Fecha"],
    )


def _view_importar(
    *,
    sociedad_id: int,
    banco_elegido: str,
    nombre_activo: str | None,
    plan_vinculado: bool,
    periodo_key: str,
    preview_key: str,
    view_key: str,
) -> None:
    st.markdown("#### Importar extracto")
    st.caption("PDF del banco o Excel/CSV de la plantilla. Después: revisar matches y bajar el asiento.")
    if nombre_activo:
        st.success(f"Cliente: **{nombre_activo}** · Banco: **{banco_elegido or '—'}**")
    else:
        st.error("Seleccioná el cliente antes de subir el archivo.")
        return
    if not plan_vinculado:
        st.warning("Vinculá el plan de cuentas de esta sociedad para imputar Asiento Tango. Se puede leer el extracto igual.")

    tab_pdf, tab_xl = st.tabs(["PDF del banco", "Excel / Plantilla"])
    with tab_pdf:
        archivos_pdf = st.file_uploader(
            "Arrastrá el PDF del extracto o hacé click",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"ae_pdfs_{sociedad_id}",
        )
        st.caption("Galicia · BNA · BBVA · Macro · Santander · Provincia · Mercado Pago")
    with tab_xl:
        archivos_xl = st.file_uploader(
            "Plantilla del estudio (.xlsx) o CSV",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key=f"ae_xls_{sociedad_id}",
        )
    archivos = list(archivos_pdf or []) or list(archivos_xl or [])
    cache_files = f"ae_files_bytes_{sociedad_id}"
    if archivos:
        st.session_state[cache_files] = [(a.name, a.getvalue()) for a in archivos]

    extras = st.expander("Padrón opcional (proveedores / VEP)", expanded=False)
    with extras:
        c2, c3 = st.columns(2)
        with c2:
            excel_prov = st.file_uploader(
                "Proveedores pendientes (opcional)",
                type=["xlsx", "xls", "csv"],
                key=f"ae_prov_{sociedad_id}",
            )
        with c3:
            excel_vep = st.file_uploader(
                "Padrón VEPs AFIP (opcional)",
                type=["xlsx", "xls", "csv"],
                key=f"ae_vep_{sociedad_id}",
            )

    periodo = _periodo_str(st.session_state[periodo_key])
    existentes = db.listar_movimientos_banco(sociedad_id, periodo=periodo)
    pisar_ok = True
    if existentes:
        pisar_ok = st.checkbox(
            f"Este período ya tiene {len(existentes)} movimiento(s). Pisarlos al guardar.",
            key=f"ae_overwrite_{sociedad_id}_{periodo}",
        )

    if st.button("Leer extracto", type="primary", key=f"ae_run_{sociedad_id}"):
        pares = [(a.name, a.getvalue()) for a in archivos] if archivos else list(
            st.session_state.get(cache_files) or []
        )
        if not pares:
            st.error("Subí un PDF o Excel.")
        else:
            fuentes = [_UploadMemoria(n, b) for n, b in pares]
            with st.spinner("Leyendo extracto…"):
                df, meta, errores = procesar_extractos_bancarios_pdfs(
                    fuentes, banco_hint=banco_elegido
                )
                if errores:
                    st.warning(
                        "Algunos archivos tuvieron problemas: "
                        + "; ".join(f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5])
                    )
                if df is None or getattr(df, "empty", True):
                    st.error("Este archivo no se pudo leer. Probá la otra solapa (PDF ↔ Excel).")
                else:
                    if excel_prov is not None:
                        db.reemplazar_proveedores_pendientes(
                            sociedad_id, _cargar_proveedores_desde_upload(excel_prov)
                        )
                    if excel_vep is not None:
                        raw = excel_vep.read() if hasattr(excel_vep, "read") else excel_vep
                        db.reemplazar_veps_afip(
                            sociedad_id,
                            _cargar_veps_excel(BytesIO(raw) if isinstance(raw, bytes) else excel_vep),
                        )
                    filas = df_extracto_a_filas(df)
                    ok_saldo, msg_saldo = validar_saldos_corridos(filas)
                    if not ok_saldo:
                        st.error(f"Extracto sospechoso: {msg_saldo}")
                    banco = str((meta or {}).get("banco") or "") or banco_elegido or ""
                    resultados = correr_motor(
                        filas,
                        db.listar_reglas_clasificacion(solo_activas=True),
                        db.listar_proveedores_pendientes(sociedad_id, solo_libres=False),
                        db.listar_veps_afip(sociedad_id),
                        cliente_id=sociedad_id,
                        banco=banco,
                        periodo=st.session_state[periodo_key],
                        saldo_ok=ok_saldo,
                    )
                    movs = _aplicar_plan(_enriquecer(resultados), st.session_state.get("plan_cuentas_df"))
                    st.session_state[preview_key] = {
                        "movimientos": movs,
                        "banco": banco,
                        "periodo": periodo,
                        "meta": meta or {},
                    }
                    st.rerun()

    preview = st.session_state.get(preview_key)
    if preview:
        movs = preview.get("movimientos") or []
        n_ok = sum(1 for m in movs if str(m.get("estado") or "") in {"OK", "CONCILIADO"})
        n_pend = sum(1 for m in movs if str(m.get("estado") or "") == "PENDIENTE")
        n_999 = sum(1 for m in movs if str(m.get("cuenta_codigo") or "") == "99999")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Banco detectado", str(preview.get("banco") or "—"))
        d2.metric("Movimientos", len(movs))
        d3.metric("Clasificados", n_ok)
        d4.metric("A revisar / 99999", f"{n_pend} / {n_999}")
        st.markdown("**Vista previa** — todavía no se guardó")
        _tabla(
            movs[:120],
            ["Fecha", "Descripción", "Crédito", "Débito", "Categoría", "Cuenta", "Estado", "Por qué"],
        )
        c_ok, c_no = st.columns(2)
        with c_ok:
            if st.button("Guardar y revisar", type="primary", key=f"ae_save_{sociedad_id}"):
                if existentes and not pisar_ok:
                    st.error("Este período ya tiene movimientos. Tildá que querés pisarlos.")
                else:
                    db.borrar_movimientos_periodo(
                        sociedad_id, periodo=str(preview.get("periodo") or periodo), banco=None
                    )
                    db.insertar_movimientos_banco(movs)
                    st.session_state["motor_last_periodo"] = str(preview.get("periodo") or periodo)
                    st.session_state["motor_grilla_bridge"] = movimientos_a_filas_grilla_tango(
                        movs, plan_cuentas=st.session_state.get("plan_cuentas_df")
                    )
                    db.registrar_auditoria_conciliacion(
                        cliente_id=sociedad_id,
                        movimiento_id=None,
                        usuario=str(st.session_state.get("oficina_usuario") or "sistema"),
                        accion="guardar_extracto_ae",
                        detalle=f"{len(movs)} movimientos | banco={preview.get('banco')}",
                    )
                    st.session_state.pop(preview_key, None)
                    st.session_state[view_key] = "movimientos"
                    st.success(f"{len(movs)} movimientos guardados. Revisá matches y confirmá pendientes.")
                    st.rerun()
        with c_no:
            if st.button("Cancelar", key=f"ae_discard_{sociedad_id}"):
                st.session_state.pop(preview_key, None)
                st.rerun()


def _view_movimientos(movs: list[dict], sociedad_id: int, nombre_activo: str | None, view_key: str) -> None:
    st.markdown("#### Movimientos")
    st.caption("La categoría sale de Conceptos Bancos y de las reglas de **Reglas**. 99999 = sin match.")
    if not movs:
        st.info("No hay movimientos de este período.")
        if st.button("Ir a Importar", type="primary", key=f"ae_go_imp_{sociedad_id}"):
            st.session_state[view_key] = "importar"
            st.rerun()
        return

    _kpis_resumen(movs)
    q = st.text_input("Buscar en la descripción", key="ae_q_mov")
    c1, c2 = st.columns(2)
    with c1:
        cat = st.selectbox(
            "Vista",
            [k for k, _ in _VISTA_FILTRO],
            format_func=dict(_VISTA_FILTRO).get,
            key="ae_f_cat",
        )
    with c2:
        est = st.selectbox(
            "Estado",
            ["", "OK", "CONCILIADO", "PENDIENTE"],
            format_func=lambda x: "Todos" if x == "" else x,
            key="ae_f_est",
        )
    res = movs
    if q:
        ql = q.lower()
        res = [m for m in res if ql in _norm(m.get("descripcion"))]
    if cat:
        if cat == "egreso":
            res = [m for m in res if m.get("vista") in {"egreso", "impuesto"}]
        else:
            res = [m for m in res if m.get("vista") == cat]
    if est:
        res = [m for m in res if str(m.get("estado") or "") == est]
    _tabla(
        res,
        ["Fecha", "Descripción", "Crédito", "Débito", "Categoría", "Cuenta", "Estado", "Por qué"],
    )

    pendientes = [
        m
        for m in movs
        if str(m.get("estado") or "") == "PENDIENTE"
        and m.get("id")
        and str(m.get("categoria") or "")
        and "identificar" not in str(m.get("categoria") or "").lower()
    ]
    if pendientes:
        st.markdown("**Pendientes con categoría propuesta**")
        st.caption("El motor pide confirmación en reglas locales y matches débiles. No imputa solo.")
        if st.button(
            f"Confirmar {len(pendientes)} categoría(s) propuesta(s)",
            key=f"ae_confirm_pend_{sociedad_id}",
        ):
            usuario = str(st.session_state.get("oficina_usuario") or "sistema")
            for m in pendientes:
                db.actualizar_movimiento_banco(int(m["id"]), estado="OK")
                db.registrar_auditoria_conciliacion(
                    cliente_id=sociedad_id,
                    movimiento_id=int(m["id"]),
                    usuario=usuario,
                    accion="confirmar_categoria",
                    categoria_nueva=str(m.get("categoria") or ""),
                    detalle=str(m.get("descripcion") or ""),
                )
            st.success("Confirmadas. El asiento ya no las marca como 99999 si el plan matchea.")
            st.rerun()

    st.download_button(
        "Descargar movimientos (.xlsx)",
        data=_bytes_informe(movs, nombre_activo, "Conciliación bancaria"),
        file_name="movimientos_banco.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ae_export_mov",
    )


def _view_papeles(
    movs: list[dict],
    todos: list[dict],
    nombre_activo: str | None,
    periodo: date | None,
) -> None:
    st.markdown("#### Papeles del mes")
    st.caption(
        "Primer mes: saldo inicial del extracto. Meses siguientes: arrastre del cierre. "
        "Cierre = apertura + créditos − débitos. Según resumen = saldo del extracto. "
        "**La diferencia no se fuerza a 0** — si no da, hay que revisar."
    )
    if not movs and not todos:
        st.info("Importá el extracto para armar el papel.")
        return

    cadena = papeles_por_mes(todos or movs)
    if not cadena and movs:
        cadena = [armar_papel_mes(movs)]
    if not cadena:
        st.info("No hay fechas para armar la cadena mensual.")
        return

    actual = cadena[-1]
    if periodo is not None:
        for p in cadena:
            if p.get("anio") == periodo.year and p.get("mes") == periodo.month:
                actual = p
                break

    origen = "extracto (primer mes)" if actual.get("origen_apertura") == "extracto" else "arrastre del mes anterior"
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Apertura", _fmt(actual.get("apertura")), origen)
    k2.metric("Créditos", _fmt(actual.get("creditos")))
    k3.metric("Débitos", _fmt(actual.get("debitos")))
    k4.metric("Cierre (fórmula)", _fmt(actual.get("cierre")))
    k5.metric("Según resumen", _fmt(actual.get("segun_resumen")))
    dif = actual.get("diferencia")
    st.metric("Diferencia (resumen − cierre)", _fmt(dif))
    if dif is not None and abs(float(dif)) > 0.05:
        st.warning(
            f"Hay diferencia de {_fmt(dif)} entre el saldo del extracto y el cierre. "
            "No se corrige sola: es un desvío a revisar."
        )
    elif dif is not None:
        st.success("Cierre y según resumen coinciden (dentro de $ 0,05).")

    df_cad = pd.DataFrame(
        [
            {
                "Período": p.get("periodo"),
                "Apertura": p.get("apertura"),
                "Créditos": p.get("creditos"),
                "Débitos": p.get("debitos"),
                "Cierre": p.get("cierre"),
                "Según resumen": p.get("segun_resumen"),
                "Diferencia": p.get("diferencia"),
                "Origen apertura": p.get("origen_apertura"),
                "Movs.": p.get("movimientos"),
            }
            for p in cadena
        ]
    )
    st.dataframe(df_cad, use_container_width=True, hide_index=True)

    ret = _grupo(movs, "retencion")
    ded = _grupo(movs, "deduccion")
    with st.expander("Proyección Ganancias (×12 del mes, no reemplaza DDJJ)", expanded=False):
        icdb = sum(abs(float(m.get("monto") or 0)) for m in ret if _es_icdb(m))
        ret_g = _suma(ret, lambda x: "ganancia" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"))
        c1, c2, c3 = st.columns(3)
        c1.metric("Ded. detectadas (anualiz.)", _fmt(_suma(ded) * 12))
        c2.metric("Ret. Ganancias (anualiz.)", _fmt(ret_g * 12))
        c3.metric("ICDB 34% (anualiz.)", _fmt(icdb * 0.34 * 12))

    st.download_button(
        "Descargar papeles (.xlsx)",
        data=exportar_informe_excel(
            titulo="Papeles de conciliación bancaria",
            subtitulo=nombre_activo or "Cliente",
            periodo=_periodo_str(date.today()),
            kpis=[
                ("Apertura", actual.get("apertura") or 0),
                ("Cierre", actual.get("cierre") or 0),
                ("Según resumen", actual.get("segun_resumen") or 0),
                ("Diferencia", actual.get("diferencia") or 0),
            ],
            detalle=df_cad,
            hoja_detalle="Papeles",
            col_moneda=["Apertura", "Créditos", "Débitos", "Cierre", "Según resumen", "Diferencia"],
        ),
        file_name="papeles_conciliacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ae_export_papeles",
    )


def _view_asiento(
    movs: list[dict],
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    periodo_key: str,
) -> None:
    st.markdown("#### Asiento Tango")
    st.caption(
        "Contrapartidas agrupadas y cuenta banco. **99999 = sin match**, no se inventa. "
        "El asiento tiene que cerrar (Debe = Haber) para exportar. "
        "La diferencia de papeles (según resumen) no se toca acá."
    )
    st.checkbox(
        "Mostrar asiento CN desde el Balance Excel (flujo anterior)",
        key="ae_mostrar_flujo_balance",
        help="Abre debajo el armado viejo a partir del Excel de bancos. No hace falta para el extracto.",
    )
    if not movs:
        st.info("Sin filas para el asiento. Importá el extracto y guardá.")
        return

    plan_df = st.session_state.get("plan_cuentas_df")
    opciones = _opciones_plan(plan_df)
    banco = str(banco_elegido or movs[0].get("banco") or "Banco")
    sug_cod, sug_desc = cuenta_banco_del_plan(plan_df, banco)
    cuenta_banco_opt = st.selectbox(
        "Cuenta banco (contrapartida del asiento)",
        opciones,
        index=max(0, opciones.index(_opcion_desde_codigo(sug_cod, opciones)))
        if _opcion_desde_codigo(sug_cod, opciones) in opciones
        else 0,
        key=f"ce_cta_banco_{sociedad_id}",
    )
    codigo_banco = _codigo_desde_opcion(cuenta_banco_opt)
    desc_banco = (
        cuenta_banco_opt.split("—", 1)[-1].strip() if "—" in cuenta_banco_opt else sug_desc
    )

    periodo_date = st.session_state.get(periodo_key)
    if isinstance(periodo_date, date):
        ultimo = calendar.monthrange(periodo_date.year, periodo_date.month)[1]
        fecha_asiento = date(periodo_date.year, periodo_date.month, ultimo)
        periodo = f"{periodo_date.month:02d}/{periodo_date.year}"
    else:
        fecha_asiento = date.today()
        periodo = fecha_asiento.strftime("%m/%Y")
    fecha_str = fecha_asiento.strftime("%d/%m/%Y")

    rows = renglones_asiento_banco_mes(
        movs,
        codigo_banco=codigo_banco,
        descripcion_banco=desc_banco or banco,
        periodo=periodo,
        fecha_str=fecha_str,
    )
    g = gate_asiento(rows)
    debe = sum(float(r.get("Debe") or 0) for r in rows)
    haber = sum(float(r.get("Haber") or 0) for r in rows)
    n_999 = sum(1 for r in rows if str(r.get("Código") or "") == "99999")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Debe", _fmt(debe))
    k2.metric("Haber", _fmt(haber))
    k3.metric("Diferencia", _fmt(debe - haber))
    k4.metric("Sin match (99999)", str(n_999))
    st.dataframe(
        pd.DataFrame(rows)[["Código", "Descripción", "Debe", "Haber"]] if rows else pd.DataFrame(),
        use_container_width=True,
        hide_index=True,
    )
    st.session_state["motor_grilla_bridge"] = movimientos_a_filas_grilla_tango(
        movs, plan_cuentas=plan_df
    )

    if g.get("bloqueantes"):
        st.error("Todavía no se puede guardar ni exportar:")
        for item in g["bloqueantes"]:
            st.markdown(f"- {item}")
        st.caption("Completá las líneas 99999 en Movimientos o elegí la cuenta banco del plan.")
        return

    asiento = _asiento_desde_rows(rows, banco=banco, periodo=periodo, fecha_asiento=fecha_asiento)
    c_bib, c_xls = st.columns(2)
    with c_bib:
        if st.button(
            "Guardar en biblioteca",
            type="primary",
            use_container_width=True,
            key=f"ce_bib_{sociedad_id}",
        ):
            try:
                etiqueta = _guardar_biblioteca_banco(
                    sociedad_id, [asiento], rows, banco=banco, periodo=periodo,
                )
                st.success(f"Asiento {banco} {etiqueta} guardado en la biblioteca.")
            except ValueError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"No se pudo guardar: {exc}")
    with c_xls:
        try:
            fp = (
                periodo,
                codigo_banco,
                tuple((r.get("Código"), r.get("Debe"), r.get("Haber")) for r in rows),
            )
            cache_fp = f"ce_xlsx_fp_{sociedad_id}"
            cache_bytes = f"ce_xlsx_bytes_{sociedad_id}"
            cache_name = f"ce_xlsx_name_{sociedad_id}"
            if st.session_state.get(cache_fp) != fp:
                ruta = generar_excel_tango_nativo(
                    [asiento],
                    nombre_activo or banco,
                    str(cuit_activo or ""),
                    fecha_asiento.month,
                    fecha_asiento.year,
                    plan_cuentas=plan_df,
                )
                st.session_state[cache_bytes] = ruta.read_bytes()
                st.session_state[cache_name] = ruta.name
                st.session_state[cache_fp] = fp
            st.download_button(
                "Excel para importar a Tango",
                data=st.session_state[cache_bytes],
                file_name=st.session_state.get(cache_name) or "asiento_banco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"ce_xlsx_{sociedad_id}",
            )
        except ExportacionTangoError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"No se pudo armar el Excel Tango: {exc}")


def _view_config(sociedad_id: int) -> None:
    st.markdown("#### Reglas de clasificación")
    st.caption(
        "Conceptos Bancos identifica si hay match. Si no, mandan estas reglas (patrón en la descripción). "
        "Lo que agregues acá se usa en Movimientos y en el asiento."
    )
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
            st.success("Regla agregada. El próximo **Leer extracto** la usa.")
            st.rerun()


def render_conciliacion_ae(
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    periodo_key = f"ae_periodo_{sociedad_id}"
    preview_key = f"ae_preview_{sociedad_id}"
    view_key = f"ae_view_{sociedad_id}"
    if periodo_key not in st.session_state:
        st.session_state[periodo_key] = date.today().replace(day=1)
    if view_key not in st.session_state:
        st.session_state[view_key] = "importar"

    # Migrar claves viejas del injerto MM Studio (dashboard, ingresos, …).
    legacy = {
        "dashboard": "movimientos",
        "ingresos": "movimientos",
        "egresos": "movimientos",
        "retenciones": "movimientos",
        "deducciones": "movimientos",
        "intercuentas": "papeles",
        "proyeccion": "papeles",
        "alertas": "movimientos",
    }
    if st.session_state.get(view_key) in legacy:
        st.session_state[view_key] = legacy[st.session_state[view_key]]

    periodo_activo = _periodo_str(st.session_state.get(periodo_key))
    preview = st.session_state.get(preview_key)
    if preview and str(preview.get("periodo") or "") != periodo_activo:
        st.session_state.pop(preview_key, None)

    st.markdown(
        """
        <style>
        div.stRadio > div[role="radiogroup"] {
            gap: 0.35rem 0.85rem;
            flex-wrap: wrap;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    top_a, top_b = st.columns([2, 3])
    with top_a:
        st.date_input(
            "Período del extracto",
            key=periodo_key,
            help="Mes que se guarda y se arrastra al mes siguiente en Papeles.",
        )
    with top_b:
        st.caption("Camino: elegir sociedad y banco → importar extracto → revisar matches → asiento Tango.")
        st.radio(
            "Paso",
            [k for k, _ in _PASOS],
            format_func=dict(_PASOS).get,
            horizontal=True,
            key=view_key,
            label_visibility="collapsed",
        )

    movs = _cargar_movs(sociedad_id, banco_elegido, periodo_key, preview_key)
    vista = st.session_state.get(view_key) or "importar"

    if vista == "importar":
        _view_importar(
            sociedad_id=sociedad_id,
            banco_elegido=banco_elegido,
            nombre_activo=nombre_activo,
            plan_vinculado=plan_vinculado,
            periodo_key=periodo_key,
            preview_key=preview_key,
            view_key=view_key,
        )
    elif vista == "movimientos":
        _view_movimientos(movs, sociedad_id, nombre_activo, view_key)
    elif vista == "papeles":
        todos = _cargar_todos_banco(
            sociedad_id, banco_elegido, preview_key, _periodo_str(st.session_state.get(periodo_key))
        )
        _view_papeles(movs, todos, nombre_activo, st.session_state.get(periodo_key))
    elif vista == "asiento":
        _view_asiento(
            movs,
            sociedad_id=sociedad_id,
            banco_elegido=banco_elegido,
            cuit_activo=cuit_activo,
            nombre_activo=nombre_activo,
            periodo_key=periodo_key,
        )
    else:
        _view_config(sociedad_id)
