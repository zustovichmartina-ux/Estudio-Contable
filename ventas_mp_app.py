"""App de ventas Mercado Pago de una sociedad (Streamlit). No mezclar con gastos personales."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

import database as db
from ventas_mp import (
    COLS,
    cargar_ledger,
    cargar_token,
    conectar,
    cuadro_por_orden,
    excel_ventas_bytes,
    exportar_excel_ventas,
    guardar_sociedad,
    guardar_token,
    kpis,
    listar_sociedades,
    parsear_reporte,
    sincronizar_ventas,
)

st.set_page_config(
    page_title="Ventas Mercado Pago · Sociedad",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ec-navy: #1F4E79; }
    h1, h2, h3 { color: var(--ec-navy) !important; }
    [data-testid="stMetricValue"] { color: var(--ec-navy); }
    </style>
    """,
    unsafe_allow_html=True,
)

DESKTOP = Path.home() / "Desktop"


def _clientes_estudio() -> list[dict]:
    try:
        return db.listar_clientes()
    except Exception:
        return []


if "slug" not in st.session_state:
    socs = listar_sociedades()
    st.session_state.slug = socs[0]["slug"] if socs else ""
    st.session_state.df = pd.DataFrame(columns=COLS)
    st.session_state.meta = {}
    st.session_state.cuenta = {}
    if st.session_state.slug:
        df_l, meta_l = cargar_ledger(st.session_state.slug)
        st.session_state.df = df_l
        st.session_state.meta = meta_l
        st.session_state.cuenta = {k: meta_l.get(k) for k in ("nombre", "email", "id", "cuit", "nickname") if meta_l.get(k)}

st.title("Ventas Mercado Pago · Sociedad")
st.caption(
    "Cuenta vendedora de la sociedad: ventas, MELI Envíos y comisiones. "
    "El libro personal de gastos sigue en el puerto 8505."
)

with st.sidebar:
    st.subheader("Sociedad")
    clientes = _clientes_estudio()
    opciones = ["(cargar a mano)"] + [f"{c.get('nombre')} · {c.get('cuit')}" for c in clientes]
    elegido = st.selectbox("Cliente del estudio", opciones)
    if elegido != "(cargar a mano)":
        idx = opciones.index(elegido) - 1
        nombre_def = str(clientes[idx].get("nombre") or "")
        cuit_def = str(clientes[idx].get("cuit") or "")
    else:
        actuales = {s["slug"]: s for s in listar_sociedades()}
        prev = actuales.get(st.session_state.slug) or {}
        nombre_def = str(prev.get("nombre") or "")
        cuit_def = str(prev.get("cuit") or "")
    nombre = st.text_input("Nombre de la sociedad", value=nombre_def)
    cuit = st.text_input("CUIT de la sociedad", value=cuit_def)
    if st.button("Usar esta sociedad", width="stretch"):
        if not nombre.strip():
            st.error("Poné el nombre de la sociedad.")
        else:
            fila = guardar_sociedad(nombre.strip(), cuit)
            st.session_state.slug = fila["slug"]
            df_l, meta_l = cargar_ledger(fila["slug"])
            st.session_state.df = df_l
            st.session_state.meta = meta_l
            st.rerun()

    slug = st.session_state.slug
    st.divider()
    st.subheader("API Mercado Pago")
    st.markdown(
        "Entrá a Mercado Pago **con el usuario de la sociedad** → "
        "[Tus integraciones](https://www.mercadopago.com.ar/developers/panel/app) → "
        "crear app → **credenciales de producción** (`APP_USR-…`). "
        "No uses el token de la cuenta personal."
    )
    hay_token = bool(slug and cargar_token(slug))
    st.caption(f"Token: {'guardado para esta sociedad' if hay_token else 'todavía no hay'}")
    token_in = st.text_input("Access Token de la sociedad", type="password")
    if st.button("Conectar", width="stretch", type="primary"):
        tok = token_in.strip() or (cargar_token(slug) if slug else "")
        if not tok:
            st.error("Pegá el APP_USR- de la cuenta vendedora.")
        else:
            try:
                cuenta = conectar(tok)
                nombre_soc = nombre.strip() or str(cuenta.get("nombre") or "Sociedad")
                cuit_soc = (cuit or "").strip() or str(cuenta.get("cuit") or "")
                fila = guardar_sociedad(nombre_soc, cuit_soc)
                slug = fila["slug"]
                st.session_state.slug = slug
                guardar_token(slug, tok)
                st.session_state.cuenta = cuenta
                st.success(f"Conectado: {cuenta['nombre']} ({cuenta.get('email') or cuenta.get('id')})")
            except Exception as exc:
                st.error(str(exc))

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        desde_api = st.date_input("Desde", date.today() - timedelta(days=30), key="v_desde")
    with col_d2:
        hasta_api = st.date_input("Hasta", date.today(), key="v_hasta")
    if st.button("Sincronizar ventas", width="stretch"):
        tok = token_in.strip() or (cargar_token(slug) if slug else "")
        if not tok:
            st.error("Pegá el APP_USR- y dale a Conectar.")
        else:
            try:
                if not slug:
                    cuenta0 = conectar(tok)
                    fila = guardar_sociedad(
                        nombre.strip() or str(cuenta0.get("nombre") or "Sociedad"),
                        (cuit or "").strip() or str(cuenta0.get("cuit") or ""),
                    )
                    slug = fila["slug"]
                    st.session_state.slug = slug
                if token_in.strip():
                    guardar_token(slug, token_in.strip())
                with st.spinner("Bajando ventas, envíos y comisiones…"):
                    df_sync, meta = sincronizar_ventas(slug, tok, desde_api, hasta_api)
                    st.session_state.df = df_sync
                    st.session_state.meta = meta
                    st.session_state.cuenta = {
                        k: meta.get(k) for k in ("nombre", "email", "id", "cuit", "nickname") if meta.get(k)
                    }
                    st.rerun()
            except Exception as exc:
                st.error(str(exc))

    st.divider()
    up = st.file_uploader("O CSV de dinero en cuenta", type=["csv"])
    if up is not None and st.button("Importar CSV", width="stretch"):
        st.session_state.df = parsear_reporte(up.getvalue())
        st.session_state.meta = {"avisos": ["CSV importado"]}
        st.rerun()

cuenta = st.session_state.get("cuenta") or {}
if cuenta.get("nombre"):
    extra = f" · {cuenta['email']}" if cuenta.get("email") else ""
    st.success(f"Sociedad conectada: **{cuenta['nombre']}**{extra}")
for aviso in (st.session_state.get("meta") or {}).get("avisos") or []:
    st.warning(aviso)

df = st.session_state.df if isinstance(st.session_state.df, pd.DataFrame) else pd.DataFrame(columns=COLS)
metricas = kpis(df)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Ventas brutas", f"$ {metricas['ventas']:,.0f}")
c2.metric("MELI Envíos", f"$ {metricas['envios']:,.0f}")
c3.metric("Comisiones", f"$ {metricas['comisiones']:,.0f}")
c4.metric("Neto acreditado", f"$ {metricas['neto']:,.0f}")

st.subheader("Por orden (venta vs envío vs comisión)")
por_orden = cuadro_por_orden(df)
if por_orden.empty:
    st.caption("Todavía no hay ventas. Conectá el token de la sociedad y sincronizá.")
else:
    st.dataframe(por_orden, width="stretch", hide_index=True)

st.subheader("Movimientos")
st.dataframe(df, width="stretch", hide_index=True)

st.divider()
periodo = f"{desde_api.strftime('%d/%m/%Y')} – {hasta_api.strftime('%d/%m/%Y')}"
nombre_soc = (nombre or "").strip() or slug or "Sociedad"
ruta_xlsx = DESKTOP / f"Ventas_MP_{slug or 'sociedad'}.xlsx"
b1, b2 = st.columns(2)
with b1:
    if st.button("Generar Excel en el Escritorio", type="primary", width="stretch"):
        exportar_excel_ventas(df, ruta_xlsx, sociedad=nombre_soc, periodo=periodo)
        st.success(f"Guardado: {ruta_xlsx}")
with b2:
    st.download_button(
        "Descargar Excel",
        data=excel_ventas_bytes(df, sociedad=nombre_soc, periodo=periodo),
        file_name=f"Ventas_MP_{slug or 'sociedad'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
