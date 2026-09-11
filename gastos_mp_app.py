"""App de gastos personales Mercado Pago (Streamlit)."""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from gastos_mp import (
    ARCHIVO_TOKEN,
    CATEGORIAS,
    COLS_MOV,
    aplicar_clasificacion,
    aplicar_saldo,
    aprender_desde_edicion,
    cargar_reglas,
    cargar_token,
    cargar_ultimo_sync,
    df_vacio,
    guardar_ledger,
    guardar_token,
    incorporar_al_libro,
    kpis,
    leer_csv_mp,
    movimientos_demo,
    resumen_por_categoria,
    sincronizar_mercadopago,
)
from gastos_mp_excel import excel_gastos_bytes, exportar_excel_gastos

st.set_page_config(
    page_title="Mi Mercado Pago",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ec-navy: #1F4E79; --ec-muted: #6B7280; --ec-bg: #F3F4F6; }
    .stApp { background: var(--ec-bg); }
    h1, h2, h3 { color: var(--ec-navy) !important; }
    [data-testid="stMetricValue"] { color: var(--ec-navy); }
    </style>
    """,
    unsafe_allow_html=True,
)

DESKTOP = Path.home() / "Desktop"
RUTA_EXCEL = DESKTOP / "Gastos_MercadoPago.xlsx"


def _filtrar(df: pd.DataFrame, desde: date, hasta: date, cats: list[str]) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["Fecha"] = pd.to_datetime(out["Fecha"]).dt.date
    out = out[(out["Fecha"] >= desde) & (out["Fecha"] <= hasta)]
    if cats:
        out = out[out["Categoria"].isin(cats)]
    return out.reset_index(drop=True)


def _chart_categorias(resumen: pd.DataFrame) -> alt.Chart:
    data = resumen[resumen["Gastado"] > 0][["Categoria", "Gastado"]]
    return (
        alt.Chart(data)
        .mark_bar(color="#1F4E79")
        .encode(
            x=alt.X("Gastado:Q", title="Gastado (ARS)"),
            y=alt.Y("Categoria:N", sort="-x", title=None),
            tooltip=[
                alt.Tooltip("Categoria:N", title="Categoría"),
                alt.Tooltip("Gastado:Q", title="Gastado (ARS)", format=",.0f"),
            ],
        )
        .properties(title="Gastos por categoría", height=max(240, 26 * max(len(data), 1)))
    )


def _chart_saldo(df: pd.DataFrame) -> alt.Chart:
    data = df.copy()
    data["Fecha"] = pd.to_datetime(data["Fecha"])
    data = data.sort_values(["Fecha", "Hora", "Id"], ascending=True)
    return (
        alt.Chart(data)
        .mark_line(color="#1F4E79")
        .encode(
            x=alt.X("Fecha:T", title="Fecha"),
            y=alt.Y("Saldo:Q", title="Saldo Mercado Pago (ARS)"),
            tooltip=[
                alt.Tooltip("Fecha:T", title="Fecha"),
                alt.Tooltip("Contraparte:N", title="Contraparte"),
                alt.Tooltip("Importe:Q", title="Importe", format=",.0f"),
                alt.Tooltip("Saldo:Q", title="Saldo", format=",.0f"),
            ],
        )
        .properties(title="Saldo después de cada movimiento", height=280)
    )


if "df" not in st.session_state or "Contraparte" not in getattr(st.session_state.df, "columns", []):
    ultimo = cargar_ultimo_sync()
    if ultimo is not None:
        st.session_state.df, st.session_state.meta = ultimo
        st.session_state.fuente = "api"
    else:
        st.session_state.df = movimientos_demo()
        st.session_state.fuente = "demo"
        st.session_state.meta = {}

if (
    st.session_state.get("fuente") == "api"
    and cargar_token()
    and not st.session_state.get("_auto_sync")
):
    st.session_state._auto_sync = True
    try:
        df_sync, meta_sync = sincronizar_mercadopago(
            cargar_token(),
            date.today() - timedelta(days=14),
            date.today(),
            rapido=True,
        )
        st.session_state.df = df_sync
        st.session_state.meta = meta_sync
        st.session_state.fuente = "api"
    except Exception:
        pass

st.title("Mi Mercado Pago")
st.caption(
    "Libro personal permanente. Cada sync suma lo nuevo (compras, transferencias, reservas, "
    "crédito, reintegros) y no borra el historial. Las sociedades van aparte."
)

with st.sidebar:
    st.subheader("Fuente")
    fuente = st.radio(
        "Datos",
        ["api", "csv", "demo"],
        format_func=lambda x: {
            "api": "API Mercado Pago",
            "csv": "CSV de Actividad",
            "demo": "Ejemplo",
        }[x],
        index={"api": 0, "csv": 1, "demo": 2}.get(st.session_state.fuente, 0),
    )

    if fuente == "demo":
        if st.button("Cargar ejemplo", width="stretch"):
            st.session_state.df = movimientos_demo()
            st.session_state.fuente = "demo"
            st.session_state.meta = {}
            st.rerun()
        st.caption("Datos de muestra para ver contraparte, detalle y saldo encadenado.")

    elif fuente == "csv":
        up = st.file_uploader("CSV de Actividad de Mercado Pago", type=["csv"])
        st.caption("En la app o web: Actividad → Descargar. Ahí viene origen/destino.")
        if up is not None and st.button("Importar CSV", width="stretch"):
            st.session_state.df = incorporar_al_libro(leer_csv_mp(up.getvalue(), up.name))
            st.session_state.fuente = "csv"
            st.session_state.meta = {"avisos": ["CSV sumado al libro permanente."]}
            st.rerun()

    else:
        token_guardado = bool(cargar_token())
        st.markdown(
            "[Tus integraciones](https://www.mercadopago.com.ar/developers/panel/app) → "
            "crear una app → **credenciales de producción** (APP_USR-…)."
        )
        st.caption(
            "Token guardado. El libro queda en tu PC y cada sync suma lo nuevo."
            if token_guardado
            else "Todavía no hay token. Pegá el APP_USR- de producción de tu cuenta."
        )
        token_in = st.text_input("Access Token", type="password")
        if token_in and st.button("Guardar token", width="stretch"):
            guardar_token(token_in)
            st.success(f"Guardado en {ARCHIVO_TOKEN}")
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            desde_api = st.date_input("Desde", date.today() - timedelta(days=90), key="api_desde")
        with col_d2:
            hasta_api = st.date_input("Hasta", date.today(), key="api_hasta")
        if st.button("Sincronizar Mercado Pago", width="stretch", type="primary"):
            tok = token_in.strip() or cargar_token()
            if not tok:
                st.error("Pegá el Access Token de producción de tu cuenta.")
            else:
                if token_in.strip():
                    guardar_token(token_in.strip())
                with st.spinner("Sumando movimientos al libro permanente…"):
                    try:
                        df_sync, meta = sincronizar_mercadopago(tok, desde_api, hasta_api)
                        st.session_state.df = df_sync
                        st.session_state.fuente = "api"
                        st.session_state.meta = meta
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
        st.caption("Suma al historial. No borra meses anteriores.")

    st.divider()
    df_all = st.session_state.df if isinstance(st.session_state.df, pd.DataFrame) else df_vacio()
    if df_all.empty:
        min_f = date.today() - timedelta(days=30)
        max_f = date.today()
    else:
        fechas = pd.to_datetime(df_all["Fecha"])
        min_f = fechas.min().date()
        max_f = fechas.max().date()
    rango = st.date_input("Período", (min_f, max_f), min_value=min_f, max_value=max_f)
    if isinstance(rango, tuple) and len(rango) == 2:
        desde, hasta = rango
    else:
        desde, hasta = min_f, max_f
    cats_sel = st.multiselect("Categorías", CATEGORIAS, default=[])
    saldo_manual = st.number_input(
        "Saldo actual en Mercado Pago (opcional)",
        min_value=0.0,
        value=0.0,
        step=1000.0,
        help="La API no deja leer el saldo. Si lo cargás, la columna Saldo se encadena hacia atrás desde este valor.",
    )
    if st.button("Aplicar saldo actual", width="stretch") and saldo_manual > 0:
        st.session_state.df = aplicar_saldo(df_all, saldo_final=float(saldo_manual))
        meta_now = dict(st.session_state.get("meta") or {})
        meta_now["saldo_actual"] = float(saldo_manual)
        st.session_state.meta = meta_now
        st.rerun()

meta = st.session_state.get("meta") or {}
if st.session_state.fuente == "demo":
    st.info("Ejemplo. Para tus datos: API (token APP_USR) o CSV de Actividad.")
elif meta.get("usuario"):
    extra = f" · saldo MP $ {meta['saldo_actual']:,.0f}" if meta.get("saldo_actual") is not None else ""
    st.success(f"Conectado como **{meta['usuario']}**{extra}")
    for aviso in meta.get("avisos") or []:
        st.warning(aviso)

df_view = _filtrar(st.session_state.df, desde, hasta, cats_sel)
metricas = kpis(df_view)
saldo_kpi = kpis(st.session_state.df)["saldo"]
c1, c2, c3, c4 = st.columns(4)
c1.metric("Saldo", f"$ {saldo_kpi:,.0f}")
c2.metric("Gastado", f"$ {metricas['gastado']:,.0f}")
c3.metric("Ingresos", f"$ {metricas['ingresos']:,.0f}")
c4.metric("Sin clasificar", f"{int(metricas['sin_clasificar'])} / {int(metricas['movimientos'])}")

resumen = resumen_por_categoria(df_view)
col_a, col_b = st.columns((1.15, 1))
with col_a:
    if not resumen.empty and (resumen["Gastado"] > 0).any():
        st.altair_chart(_chart_categorias(resumen), width="stretch")
    else:
        st.caption("No hay gastos en el filtro actual.")
with col_b:
    if not cats_sel and not df_view.empty and "Saldo" in df_view.columns:
        st.altair_chart(_chart_saldo(df_view), width="stretch")
    elif cats_sel:
        st.caption("El gráfico de saldo usa todos los movimientos. Sacá el filtro de categorías para verlo.")

st.subheader("Movimientos")
st.caption(
    "Contraparte es la persona o el comercio. Detalle es el resto (concepto, CVU, QR). "
    "Saldo es el de tu cuenta de MP después de esa fila. Si corregís la categoría y guardás, se recuerda el comercio."
)

df_editor = df_view.copy()
if not df_editor.empty:
    df_editor["Fecha"] = pd.to_datetime(df_editor["Fecha"]).dt.date
    df_editor["Id"] = df_editor["Id"].astype(str)
    keep = [c for c in COLS_MOV if c in df_editor.columns]
    df_editor = df_editor[keep]

editado = st.data_editor(
    df_editor,
    width="stretch",
    hide_index=True,
    column_config={
        "Id": st.column_config.TextColumn("Id", disabled=True, width="small"),
        "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
        "Hora": st.column_config.TextColumn("Hora", width="small"),
        "Contraparte": st.column_config.TextColumn("Contraparte", width="medium"),
        "Detalle": st.column_config.TextColumn("Detalle", width="large"),
        "Categoria": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS),
        "Importe": st.column_config.NumberColumn("Importe", format="$ %.2f"),
        "Saldo": st.column_config.NumberColumn("Saldo", format="$ %.2f", disabled=True),
        "Tipo": st.column_config.TextColumn("Tipo MP", disabled=True, width="small"),
        "Medio": st.column_config.TextColumn("Medio", disabled=True, width="small"),
    },
    key="editor_movs",
)

if st.button("Guardar clasificaciones"):
    n = aprender_desde_edicion(df_view, editado)
    base = st.session_state.df.copy()
    editado_idx = editado.set_index("Id")
    for mid, row in editado_idx.iterrows():
        mask = base["Id"].astype(str) == str(mid)
        if mask.any():
            base.loc[mask, "Categoria"] = row["Categoria"]
    saldo_final = None
    if "Saldo" in base.columns and base["Saldo"].notna().any():
        saldo_final = float(pd.to_numeric(base["Saldo"], errors="coerce").dropna().iloc[0])
    reglas = cargar_reglas()
    base = aplicar_clasificacion(base, reglas)
    for mid, row in editado_idx.iterrows():
        mask = base["Id"].astype(str) == str(mid)
        base.loc[mask, "Categoria"] = row["Categoria"]
    st.session_state.df = aplicar_saldo(base, saldo_final=saldo_final)
    guardar_ledger(st.session_state.df, dict(st.session_state.get("meta") or {}))
    st.success(f"Listo. {n} contraparte(s) quedaron en el libro.")
    st.rerun()

st.divider()
periodo = f"{desde.strftime('%d/%m/%Y')} – {hasta.strftime('%d/%m/%Y')}"
df_excel = df_view if not df_view.empty else st.session_state.df
b1, b2 = st.columns(2)
with b1:
    if st.button("Generar Excel en el Escritorio", type="primary", width="stretch"):
        exportar_excel_gastos(df_excel, RUTA_EXCEL, periodo=periodo)
        st.success(f"Guardado: {RUTA_EXCEL}")
with b2:
    st.download_button(
        "Descargar Excel",
        data=excel_gastos_bytes(df_excel, periodo=periodo),
        file_name="Gastos_MercadoPago.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
