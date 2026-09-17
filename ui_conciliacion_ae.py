# -*- coding: utf-8 -*-
"""Conciliación Bancaria con la estructura de AE_Studio.html.

La clasificación sigue siendo la de la web (Conceptos Bancos + plan del cliente).
Los movimientos se guardan en la base del estudio, como AE los guardaba en Supabase.
"""
from __future__ import annotations

from collections import defaultdict
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
from excel_formato_estudio import exportar_informe_excel
from procesador import clasificar_movimiento_extracto, procesar_extractos_bancarios_pdfs
from ui_motor_conciliacion import _cargar_proveedores_desde_upload, _cargar_veps_excel

_NAV = (
    (
        "Análisis",
        (
            ("dashboard", "Dashboard"),
            ("movimientos", "Movimientos"),
            ("ingresos", "Ingresos"),
            ("egresos", "Egresos"),
        ),
    ),
    (
        "Impositivo",
        (
            ("retenciones", "Retenciones"),
            ("deducciones", "Deducciones"),
            ("intercuentas", "Inter-cuentas"),
            ("proyeccion", "Proy. Ganancias"),
            ("alertas", "Alertas"),
        ),
    ),
    (
        "Sistema",
        (
            ("importar", "Importar"),
            ("asiento", "Asiento Tango"),
            ("config", "Configuración"),
        ),
    ),
)


def _periodo_str(d: date | None) -> str:
    if not d:
        return date.today().strftime("%Y-%m-01")
    return d.replace(day=1).isoformat()


def _fmt(n: float) -> str:
    return f"$ {n:,.2f}"


def _norm(texto: str) -> str:
    return str(texto or "").lower()


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
        fila["sub_categoria"] = label if label and label != "Sin clasificar" else str(
            m.get("concepto_instructivo") or m.get("categoria") or "-"
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


def _icdb(movs: list[dict]) -> float:
    return sum(abs(float(m.get("monto") or 0)) for m in movs if _es_icdb(m))


def _por_sub(movs: list[dict], n: int = 6) -> dict[str, float]:
    acc: dict[str, float] = defaultdict(float)
    for m in movs:
        acc[str(m.get("sub_categoria") or "-")] += abs(float(m.get("monto") or 0))
    return dict(sorted(acc.items(), key=lambda kv: kv[1], reverse=True)[:n])


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
                "Monto": m.get("monto"),
                "Tipo": "Cred" if float(m.get("monto") or 0) >= 0 else "Deb",
                "Categoria": m.get("vista"),
                "Sub-cat.": m.get("sub_categoria"),
                "Cuenta": m.get("cuenta_codigo") or "",
                "Plan": m.get("cuenta_plan") or m.get("categoria") or "",
                "Estado": m.get("estado"),
                "Banco": m.get("banco"),
            }
        )
    df = pd.DataFrame(filas)
    if cols:
        keep = [c for c in cols if c in df.columns]
        df = df[keep]
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"{len(movs)} movimiento(s)")


def _cargar_movs(sociedad_id: int, banco_elegido: str, periodo_key: str, preview_key: str) -> list[dict]:
    preview = st.session_state.get(preview_key)
    if preview:
        movs = list(preview.get("movimientos") or [])
    else:
        periodo_activo = st.session_state.get("motor_last_periodo") or _periodo_str(
            st.session_state.get(periodo_key)
        )
        guardados = db.listar_movimientos_banco(sociedad_id, periodo=periodo_activo)
        plan_df = st.session_state.get("plan_cuentas_df")
        movs = _aplicar_plan(_enriquecer(guardados), plan_df)
    banco_filtro = str(banco_elegido or "").strip()
    if banco_filtro:
        filtrados = [
            m
            for m in movs
            if not m.get("banco") or banco_filtro.lower() in str(m.get("banco") or "").lower()
        ]
        if filtrados:
            movs = filtrados
    return movs


def _bytes_asiento(movs: list[dict], nombre: str | None) -> bytes:
    filas = []
    for m in movs:
        filas.append(
            {
                "Fecha": m.get("fecha"),
                "Descripción": m.get("descripcion"),
                "Monto": float(m.get("monto") or 0),
                "Vista": m.get("vista"),
                "Subcategoría": m.get("sub_categoria"),
                "Cuenta": m.get("cuenta_codigo") or "",
                "Plan": m.get("cuenta_plan") or "",
                "Estado": m.get("estado"),
                "Banco": m.get("banco"),
            }
        )
    df = pd.DataFrame(filas)
    return exportar_informe_excel(
        titulo="Conciliación bancaria",
        subtitulo=nombre or "Cliente",
        periodo=_periodo_str(date.today()),
        kpis=[
            ("Movimientos", len(movs)),
            ("Ingresos", sum(m["monto"] for m in movs if float(m.get("monto") or 0) > 0)),
            ("Egresos", sum(abs(m["monto"]) for m in movs if float(m.get("monto") or 0) < 0)),
        ],
        detalle=df if not df.empty else pd.DataFrame(columns=["Fecha", "Descripción", "Monto"]),
        hoja_detalle="Movimientos",
        col_moneda=["Monto"],
        col_fecha=["Fecha"],
        total_col="Monto",
    )


def _view_dashboard(movs: list[dict], nombre: str | None) -> None:
    cab, acc = st.columns([3, 1])
    with cab:
        st.markdown("### Dashboard")
        st.caption(f"{nombre or 'Sin cliente'} — {len(movs)} mov." if movs else "Sin datos")
    with acc:
        if movs:
            st.download_button(
                "Exportar asiento",
                data=_bytes_asiento(movs, nombre),
                file_name="asiento_banco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="ae_export_dash",
            )
    if not movs:
        st.info("Importá extractos para comenzar. Menú **Importar**, a la izquierda.")
        return
    ing = _grupo(movs, "ingreso")
    egr = [m for m in movs if m.get("vista") in {"egreso", "impuesto"}]
    ret = _grupo(movs, "retencion")
    ded = _grupo(movs, "deduccion")
    ic = _grupo(movs, "inter-cta")
    icdb = _icdb(ret)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Ingresos", _fmt(sum(m["monto"] for m in ing if m["monto"] > 0)), f"{len(ing)} mov.")
    c2.metric("Egresos", _fmt(_suma(egr)), f"{len(egr)} mov.")
    c3.metric("Retenciones", _fmt(_suma(ret)))
    c4.metric("Deducciones", _fmt(_suma(ded)))
    c5.metric("Inter-cuentas", _fmt(_suma(ic)))
    c6.metric("ICDB comp 34%", _fmt(icdb * 0.34))
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Ingresos por tipo**")
        data = _por_sub(ing)
        st.bar_chart(pd.Series(data, name="ARS") if data else pd.Series(dtype=float), height=180)
    with col_b:
        st.markdown("**Egresos por tipo**")
        data = _por_sub(egr)
        st.bar_chart(pd.Series(data, name="ARS") if data else pd.Series(dtype=float), height=180)
    st.markdown("**Últimos movimientos**")
    _tabla(movs[:8], ["Fecha", "Descripción", "Monto", "Categoria", "Sub-cat."])


def _view_movimientos(movs: list[dict]) -> None:
    st.markdown("### Movimientos")
    st.caption("Grilla completa. La categoría sale de las reglas de la web.")
    q = st.text_input("Buscar", key="ae_q_mov")
    cat = st.selectbox(
        "Categoría",
        ["", "ingreso", "egreso", "retencion", "deduccion", "inter-cta", "sin-cat"],
        format_func=lambda x: "Todas" if x == "" else x,
        key="ae_f_cat",
    )
    res = movs
    if q:
        ql = q.lower()
        res = [m for m in res if ql in _norm(m.get("descripcion"))]
    if cat:
        res = [m for m in res if m.get("vista") == cat]
    _tabla(res)


def _view_ingresos(movs: list[dict]) -> None:
    st.markdown("### Ingresos")
    st.caption("Gravados · Exentos · A analizar")
    m = _grupo(movs, "ingreso")
    hab = _suma(m, lambda x: "habere" in _norm(x.get("sub_categoria")) or "sueldo" in _norm(x.get("descripcion")))
    hon = _suma(m, lambda x: "honorario" in _norm(x.get("sub_categoria")) or "factura" in _norm(x.get("descripcion")))
    ter = _suma(m, lambda x: "tercer" in _norm(x.get("sub_categoria")) or "recibid" in _norm(x.get("sub_categoria")))
    exe = _suma(m, lambda x: "interes" in _norm(x.get("sub_categoria")))
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Haberes/Sueldos", _fmt(hab))
    c2.metric("Honorarios", _fmt(hon))
    c3.metric("Terceros", _fmt(ter))
    c4.metric("Exentos", _fmt(exe))
    c5.metric("Otros", _fmt(max(0.0, _suma(m) - hab - hon - ter - exe)))
    q = st.text_input("Buscar", key="ae_q_ing")
    f = [x for x in m if not q or q.lower() in _norm(x.get("descripcion"))]
    _tabla(f, ["Fecha", "Descripción", "Monto", "Sub-cat.", "Banco"])


def _view_egresos(movs: list[dict]) -> None:
    st.markdown("### Egresos")
    st.caption("Gastos · Tarjetas · Débitos")
    m = [x for x in movs if x.get("vista") in {"egreso", "impuesto"}]
    tar = _suma(m, lambda x: "tarjeta" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"))
    afip = _suma(m, lambda x: "afip" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"))
    com = _suma(m, lambda x: "compra" in _norm(x.get("sub_categoria")))
    tra = _suma(m, lambda x: "emitid" in _norm(x.get("sub_categoria")) or "a tercero" in _norm(x.get("descripcion")))
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Pago tarjetas", _fmt(tar))
    c2.metric("Pagos AFIP", _fmt(afip))
    c3.metric("Compras débito", _fmt(com))
    c4.metric("Transf. salidas", _fmt(tra))
    c5.metric("Otros", _fmt(max(0.0, _suma(m) - tar - afip - com - tra)))
    q = st.text_input("Buscar", key="ae_q_egr")
    f = [x for x in m if not q or q.lower() in _norm(x.get("descripcion"))]
    _tabla(f, ["Fecha", "Descripción", "Monto", "Sub-cat.", "Banco"])


def _view_retenciones(movs: list[dict]) -> None:
    st.markdown("### Retenciones")
    st.caption("IIBB · Ganancias · IVA · ICDB")
    m = _grupo(movs, "retencion")
    icdb = _icdb(m)
    gc = _suma(m, lambda x: "ganancia" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"))
    iibb = _suma(m, lambda x: any(k in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}") for k in ("iibb", "arba", "sircreb", "bruto")))
    iva = _suma(m, lambda x: "iva" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"))
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("ICDB total (L.25.413)", _fmt(icdb))
    c2.metric("ICDB comp (34%)", _fmt(icdb * 0.34))
    c3.metric("Ret. Ganancias", _fmt(gc))
    c4.metric("Ret IIBB", _fmt(iibb))
    c5.metric("Percep IVA", _fmt(iva))
    c6.metric("Otras", _fmt(max(0.0, _suma(m) - icdb - gc - iibb - iva)))
    _tabla(m, ["Fecha", "Descripción", "Monto", "Sub-cat.", "Plan"])


def _view_deducciones(movs: list[dict]) -> None:
    st.markdown("### Deducciones")
    st.caption("Art. 85 LIG · Autónomos · Prepaga")
    m = _grupo(movs, "deduccion")
    alta = _suma(m, lambda x: any(k in _norm(x.get("sub_categoria")) for k in ("autonom", "monotributo")))
    media = _suma(m, lambda x: any(k in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}") for k in ("obra social", "prepaga", "seguro")))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total", _fmt(_suma(m)))
    c2.metric("Certeza alta", _fmt(alta))
    c3.metric("Certeza media", _fmt(media))
    c4.metric("A verificar", _fmt(max(0.0, _suma(m) - alta - media)))
    _tabla(m, ["Fecha", "Descripción", "Monto", "Sub-cat.", "Banco"])


def _view_intercuentas(movs: list[dict]) -> None:
    st.markdown("### Inter-cuentas")
    st.caption("Neutros · FCI · Dólar — excluidos del análisis impositivo")
    m = _grupo(movs, "inter-cta")
    sus = _suma(m, lambda x: "suscrip" in _norm(x.get("sub_categoria")))
    res = _suma(m, lambda x: "rescate" in _norm(x.get("sub_categoria")))
    usd = _suma(m, lambda x: "usd" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}") or "extranj" in _norm(x.get("descripcion")))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total", _fmt(_suma(m)))
    c2.metric("FCI suscripciones", _fmt(sus))
    c3.metric("FCI rescates", _fmt(res))
    c4.metric("Compra/venta USD", _fmt(usd))
    _tabla(m, ["Fecha", "Descripción", "Monto", "Sub-cat.", "Banco"])


def _view_proyeccion(movs: list[dict]) -> None:
    st.markdown("### Proy. Ganancias")
    st.caption("Estimación con el extracto del mes. No reemplaza la DDJJ.")
    if not movs:
        st.info("Sin movimientos. Importá el extracto.")
        return
    factor = 12.0
    ret = _grupo(movs, "retencion")
    ded = _grupo(movs, "deduccion")
    icdb34 = _icdb(ret) * 0.34 * factor
    ret_g = _suma(
        ret,
        lambda x: "ganancia" in _norm(f"{x.get('sub_categoria')} {x.get('descripcion')}"),
    ) * factor
    ded_a = _suma(ded) * factor
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ded. detectadas (anualiz.)", _fmt(ded_a))
    c2.metric("Ret. Ganancias (anualiz.)", _fmt(ret_g))
    c3.metric("ICDB 34% (anualiz.)", _fmt(icdb34))
    c4.metric("Pagos a cuenta (anualiz.)", _fmt(ret_g + icdb34))
    st.caption("Factor ×12 sobre el mes importado.")


def _view_alertas(movs: list[dict]) -> None:
    st.markdown("### Alertas")
    st.caption("Pendientes de confirmar y débitos AFIP a identificar.")
    pend = [
        m
        for m in movs
        if str(m.get("estado") or "") == "PENDIENTE" or m.get("vista") == "sin-cat"
    ]
    afip = [
        m
        for m in movs
        if m.get("vista") == "impuesto"
        or "afip" in _norm(f"{m.get('descripcion')} {m.get('sub_categoria')}")
    ]
    c1, c2 = st.columns(2)
    c1.metric("Pendientes", len(pend))
    c2.metric("Pagos AFIP", len(afip))
    st.markdown("**A revisar**")
    _tabla(pend[:80], ["Fecha", "Descripción", "Monto", "Categoria", "Estado"])


def _view_importar(
    *,
    sociedad_id: int,
    banco_elegido: str,
    nombre_activo: str | None,
    plan_vinculado: bool,
    periodo_key: str,
    preview_key: str,
) -> None:
    st.markdown("### Importar extracto")
    st.caption("PDF · Excel · igual que AE Studio. Paso 1: cliente ya elegido arriba.")
    if nombre_activo:
        st.success(f"Cliente: **{nombre_activo}**")
    else:
        st.error("Seleccioná el cliente antes de subir el archivo.")
        return
    if not plan_vinculado:
        st.error("Vinculá el plan de cuentas de esta sociedad antes de clasificar.")

    tab_pdf, tab_xl = st.tabs(["PDF del banco", "Excel / Plantilla"])
    with tab_pdf:
        archivos = st.file_uploader(
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
        archivos = archivos or archivos_xl

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
    periodo_ui = st.date_input(
        "Período (mes del extracto)",
        value=st.session_state[periodo_key],
        key=f"ae_periodo_ui_{sociedad_id}",
    )
    st.session_state[periodo_key] = periodo_ui.replace(day=1)
    periodo = _periodo_str(st.session_state[periodo_key])

    if st.button("Leer extracto", type="primary", key=f"ae_run_{sociedad_id}"):
        if not archivos:
            st.error("Subí un PDF o Excel.")
        elif not plan_vinculado:
            st.error("Falta el plan de cuentas.")
        else:
            with st.spinner("Leyendo extracto (OCR si es escaneo)…"):
                df, meta, errores = procesar_extractos_bancarios_pdfs(archivos)
                if errores:
                    st.warning(
                        "Algunos archivos tuvieron problemas: "
                        + "; ".join(f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5])
                    )
                if df is None or getattr(df, "empty", True):
                    st.error("Este PDF no se pudo leer. Probá la solapa Excel.")
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
                    n_ing = len(_grupo(movs, "ingreso"))
                    n_egr = len([x for x in movs if x.get("vista") in {"egreso", "impuesto"}])
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("Banco detectado", banco or "—")
                    d2.metric("Movimientos", len(movs))
                    d3.metric("Ingresos", n_ing)
                    d4.metric("Egresos", n_egr)
                    st.rerun()

    preview = st.session_state.get(preview_key)
    if preview:
        movs = preview.get("movimientos") or []
        st.markdown("**Vista previa**")
        _tabla(movs[:80], ["Fecha", "Descripción", "Monto", "Categoria", "Sub-cat."])
        c_ok, c_no = st.columns(2)
        with c_ok:
            if st.button("Guardar movimientos", type="primary", key=f"ae_save_{sociedad_id}"):
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
                st.session_state[f"ae_view_{sociedad_id}"] = "dashboard"
                st.success(f"{len(movs)} movimientos guardados.")
                st.rerun()
        with c_no:
            if st.button("Cancelar", key=f"ae_discard_{sociedad_id}"):
                st.session_state.pop(preview_key, None)
                st.rerun()


def _view_asiento(movs: list[dict]) -> None:
    st.markdown("### Asiento Tango")
    st.caption("Código del plan de esta sociedad. 99999 = sin match, no se inventa.")
    bridge = movimientos_a_filas_grilla_tango(
        [m for m in movs if m.get("estado") != "PENDIENTE"],
        plan_cuentas=st.session_state.get("plan_cuentas_df"),
    )
    if not bridge:
        st.info("Sin filas para el asiento. Importá el extracto y guardá.")
        return
    df_as = pd.DataFrame(bridge)
    debe = float(df_as["debe"].sum()) if "debe" in df_as.columns else 0.0
    haber = float(df_as["haber"].sum()) if "haber" in df_as.columns else 0.0
    k1, k2, k3 = st.columns(3)
    k1.metric("Debe", _fmt(debe))
    k2.metric("Haber", _fmt(haber))
    k3.metric("Diferencia", _fmt(debe - haber))
    st.dataframe(df_as, use_container_width=True, hide_index=True)
    st.session_state["motor_grilla_bridge"] = bridge


def _view_config(sociedad_id: int) -> None:
    st.markdown("### Configuración")
    st.caption("Instructivo Conceptos Bancos y reglas locales (respaldo).")
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
        st.session_state[view_key] = "dashboard"

    movs = _cargar_movs(sociedad_id, banco_elegido, periodo_key, preview_key)
    counts = {
        "dashboard": len(movs),
        "movimientos": len(movs),
        "ingresos": len(_grupo(movs, "ingreso")),
        "egresos": len([m for m in movs if m.get("vista") in {"egreso", "impuesto"}]),
        "retenciones": len(_grupo(movs, "retencion")),
        "deducciones": len(_grupo(movs, "deduccion")),
        "intercuentas": len(_grupo(movs, "inter-cta")),
        "proyeccion": 0,
        "alertas": len(
            [
                m
                for m in movs
                if str(m.get("estado") or "") == "PENDIENTE" or m.get("vista") == "sin-cat"
            ]
        ),
        "importar": 0,
        "asiento": 0,
        "config": 0,
    }

    col_nav, col_main = st.columns([1, 4], gap="large")
    with col_nav:
        st.caption("MM · Studio")
        actual = st.session_state.get(view_key) or "dashboard"
        for titulo, items in _NAV:
            st.markdown(f"**{titulo}**")
            for key, label in items:
                n = counts.get(key) or 0
                txt = f"{label}  ({n})" if n else label
                tipo = "primary" if actual == key else "secondary"
                if st.button(
                    txt,
                    key=f"ae_nav_{sociedad_id}_{key}",
                    use_container_width=True,
                    type=tipo,
                ):
                    st.session_state[view_key] = key
                    st.rerun()
        if st.button(
            "+ Importar",
            type="primary",
            key=f"ae_btn_imp_{sociedad_id}",
            use_container_width=True,
        ):
            st.session_state[view_key] = "importar"
            st.rerun()

    vista = st.session_state.get(view_key) or "dashboard"
    with col_main:
        if vista == "dashboard":
            _view_dashboard(movs, nombre_activo)
        elif vista == "movimientos":
            _view_movimientos(movs)
        elif vista == "ingresos":
            _view_ingresos(movs)
        elif vista == "egresos":
            _view_egresos(movs)
        elif vista == "retenciones":
            _view_retenciones(movs)
        elif vista == "deducciones":
            _view_deducciones(movs)
        elif vista == "intercuentas":
            _view_intercuentas(movs)
        elif vista == "proyeccion":
            _view_proyeccion(movs)
        elif vista == "alertas":
            _view_alertas(movs)
        elif vista == "importar":
            _view_importar(
                sociedad_id=sociedad_id,
                banco_elegido=banco_elegido,
                nombre_activo=nombre_activo,
                plan_vinculado=plan_vinculado,
                periodo_key=periodo_key,
                preview_key=preview_key,
            )
        elif vista == "asiento":
            _view_asiento(movs)
        else:
            _view_config(sociedad_id)
