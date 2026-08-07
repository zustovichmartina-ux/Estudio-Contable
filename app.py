"""Aplicación Streamlit — Estudio Contable."""

import copy
import json
import calendar
import re
import shutil
import socket
import tempfile
import time
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from generar_auditoria import (
    procesar_todos as _ga_procesar_todos,
    generar_excel as _ga_generar_excel,
    CARPETA as _GA_CARPETA_PRESTAMOS,
)

import database as db
from cruce_facturas_arca import procesar_cruce_facturas_arca
from procesador import (
    BANCOS_ARGENTINOS,
    COMPRAS_TANGO_PATH,
    MAX_PDFS_ANUALES,
    PERFILES_BANCO,
    PLANTILLA_CONCILIACION,
    PLAN_CUENTAS_DEFAULT,
    RUTA_RAIZ_CLIENTES,
    TAX_REGISTRY,
    BANK_REGISTRY,
    obtener_ficha_banco,
    AsientoDevengamiento,
    _agregar_renglon,
    aplicar_saldos_al_resultado,
    cargar_balance_desde_ruta_unc,
    sanitizar_ruta_unc,
    es_ruta_http_legacy,
    ruta_balance_local_por_sociedad,
    BALANCE_EXCEL_PROYECTO,
    extraer_filas_universales_balance_por_periodo_con_errores,
    extraer_filas_universales_balance_por_banco_con_errores,
    formatear_periodo_mm_yyyy,
    formatear_fecha_dd_mm_yyyy,
    monto_neto_fila_grilla,
    aplicar_monto_editable_fila,
    resolver_cuenta_banco_hibrida,
    listar_periodos_disponibles_balance,
    listar_bancos_conciliacion,
    listar_solapas_excel,
    resolver_solapa_balance,
    localizar_encabezado_meses_balance,
    mapear_columnas_periodo_balance,
    resolver_indice_columna_periodo,
    celda_coincide_periodo_seleccionado,
    MESES_MAP,
    _parsear_periodo_balance_celda as _parsear_periodo_balance_celda_proc,
    cargar_compras_tango,
    cargar_movimientos_contables,
    cargar_plan_cuentas,
    leer_dataframe_balance_solapa,
    leer_datos_balance_por_ficha,
    listar_solapas_excel,
    obtener_ficha_impuesto,
    conciliar_movimientos,
    escanear_carpeta_cliente,
    extraer_movimientos_anuales,
    generar_excel_tango_nativo,
    generar_planilla_conciliacion,
    generar_txt_tango,
    movimientos_a_dataframe,
    resultado_a_dataframe,
    guardar_biblioteca_persistida,
    cargar_biblioteca_persistida,
    guardar_borrador_grilla_persistido,
    cargar_borrador_grilla_persistido,
    limpiar_borrador_grilla_persistido,
    ruta_biblioteca_persistida_activa,
    preparar_asientos_export_tango,
    auditar_renglones_imputables_tango,
    auditar_exportacion_tango,
    resumir_informe_export_tango,
    ExportacionTangoError,
    extraer_datos_pdf_galicia,
    conciliar_banco_con_tango,
    cargar_subdiario_tango,
    extraer_movimientos_banco,
    movimientos_banco_a_dataframe_conciliacion,
    extraer_codigo_cuenta_tango_desde_concepto,
    parsear_fecha_export_tango,
    procesar_facturas_monotributo,
    exportar_monotributo_excel,
    procesar_extractos_santander_pdfs,
    procesar_extractos_bancarios_pdfs,
    exportar_extracto_santander_excel,
    exportar_extracto_bancario_excel,
    enriquecer_df_extracto_formato_banco,
    exportar_zip_extractos_por_banco,
    unir_pdfs_en_memoria,
    cargar_facturas_proveedores_excel,
    cargar_debitos_desde_extracto_df,
    matchear_debitos_con_facturas,
    ejecutar_match_debitos_proveedores,
    exportar_match_proveedores_excel,
    TOL_MATCH_DIAS,
    TOL_MATCH_DIAS_SOLO_MONTO,
    TOL_MATCH_DIAS_MONTO_UNICO,
    PROVEEDORES_DEFAULT_PATH,

    extraer_datos_liquidacion_pdf,
    procesar_liquidaciones_tarjeta_pdfs,
    exportar_liquidaciones_tarjeta_excel,
    PLANTILLAS_TARJETAS,
    detectar_entidad_por_texto,
    extraer_con_plantilla,
    extraer_texto_liquidacion_pdf,
    PERFILES_BANCO,
)

from inversiones import (
    GRUPOS_ESPECIE,
    aplicar_fifo,
    exportar_inversiones_excel,
    extraer_saldo_inicial_ddjj_pdf,
    leer_saldo_inicial_excel,
    plantilla_saldo_inicial_excel,
    procesar_archivos_inversiones,
    reclasificar_movimientos,
)
from caja_usd import (
    armar_caja_usd,
    exportar_caja_usd_excel,
    leer_cotizaciones_excel,
    movimientos_desde_excel_extracto,
    movimientos_desde_pdfs_caja_usd,
    plantilla_cotizaciones_excel,
)
from completar_cuadro_bancario import (
    completar_cuadro_bancario_existente,
    explorar_buzon_cuadros_bancarios,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_PLANES_DIR = BASE_DIR / "data" / "planes_cuentas"
# Canónico compartido en la red del estudio (mismo T: que biblioteca/borradores).
PLANES_RED_DIR = Path(r"T:\Estudio Contable") / "planes_cuentas"
LEGACY_PLANES_DIR = BASE_DIR / "planes de cuentas"
PLANILLA_IVA_DEFAULT = BASE_DIR / "PLANILLA PARA IMPORTAR IVA.xlsx"


def _es_entorno_cloud() -> bool:
    """True en Streamlit Community Cloud (sin T:/ ni Escritorio del estudio)."""
    import os

    flags = (
        os.environ.get("STREAMLIT_SHARING_MODE"),
        os.environ.get("STREAMLIT_CLOUD"),
        os.environ.get("IS_STREAMLIT_CLOUD"),
    )
    if any(str(f).strip().lower() in {"1", "true", "yes"} for f in flags if f):
        return True
    host = str(os.environ.get("HOSTNAME") or os.environ.get("COMPUTERNAME") or "").lower()
    if host.endswith(".streamlit.app") or "streamlit" in host:
        return True
    # Layout típico de Community Cloud
    if Path("/mount/src").is_dir() or Path("/home/appuser").is_dir():
        return True
    # Linux sin red del estudio: tratar como cloud/remoto para UI de rutas
    if os.name != "nt" and not PLANES_RED_DIR.exists():
        return True
    return False

st.set_page_config(
    page_title="Estudio Contable",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&display=swap');

    :root {
        --ec-navy: #1F4E79;
        --ec-navy-soft: #2E5F8A;
        --ec-slate: #334155;
        --ec-muted: #64748B;
        --ec-line: #E2E8F0;
        --ec-bg: #F7F9FC;
        --ec-card: #FFFFFF;
        --ec-accent-soft: rgba(31, 78, 121, 0.08);
        --ec-radius: 8px;
        --ec-radius-sm: 6px;
        --ec-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
        --ec-font: "Source Sans 3", "Segoe UI", "Calibri", system-ui, sans-serif;
    }

    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"],
    [data-testid="stSidebar"], .stMarkdown, .stButton, .stTextInput {
        font-family: var(--ec-font) !important;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
    }

    .stApp {
        background: var(--ec-bg) !important;
    }

    [data-testid="stHeader"] {
        background: rgba(247, 249, 252, 0.92) !important;
        border-bottom: 1px solid var(--ec-line);
    }

    [data-testid="stAppViewContainer"] > .main {
        opacity: 1 !important;
    }

    .main .block-container {
        padding-top: 1.1rem !important;
        padding-bottom: 2.5rem !important;
        max-width: 1400px;
        opacity: 1 !important;
    }

    /* Selectboxes: no CSS agresivo de BaseWeb (anti removeChild). */

    [data-testid="stException"][data-ec-hidden-noise="1"],
    .stException[data-ec-hidden-noise="1"],
    [data-testid="stAlert"][data-ec-hidden-noise="1"] {
        display: none !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
        border: none !important;
    }

    .stMarkdown p, [data-testid="stCaptionContainer"] {
        font-size: 15px !important;
        color: var(--ec-slate) !important;
        line-height: 1.5 !important;
    }
    h1 {
        font-size: 1.75rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
        color: var(--ec-navy) !important;
        margin-bottom: 0.35rem !important;
    }
    h2 {
        font-size: 1.25rem !important;
        font-weight: 600 !important;
        color: var(--ec-navy) !important;
    }
    h3, h4 {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: var(--ec-slate) !important;
    }
    small, [data-testid="stCaptionContainer"] {
        color: var(--ec-muted) !important;
    }

    [data-testid="stSidebar"] {
        background: #FFFFFF !important;
        border-right: 1px solid var(--ec-line) !important;
    }
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0.85rem;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label {
        font-size: 14px !important;
    }

    div[data-testid="stExpander"],
    div[data-testid="stForm"],
    div[data-testid="stAlert"],
    div[data-testid="stMetric"],
    .stDataFrame,
    [data-testid="stFileUploader"] {
        border-radius: var(--ec-radius) !important;
    }
    div[data-testid="stExpander"] {
        background: var(--ec-card) !important;
        border: 1px solid var(--ec-line) !important;
        box-shadow: var(--ec-shadow);
    }

    .stTextInput input, .stNumberInput input, .stTextArea textarea {
        border-radius: var(--ec-radius-sm) !important;
        border: 1px solid var(--ec-line) !important;
        background: #FFFFFF !important;
        color: var(--ec-slate) !important;
        font-size: 15px !important;
        box-shadow: none !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus {
        border-color: rgba(31, 78, 121, 0.45) !important;
        box-shadow: 0 0 0 3px var(--ec-accent-soft) !important;
    }

    .stButton > button,
    .stDownloadButton > button,
    .stFormSubmitButton > button {
        border-radius: var(--ec-radius-sm) !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        padding: 0.45rem 1rem !important;
        transition: background 0.15s ease, border-color 0.15s ease, opacity 0.15s ease !important;
        box-shadow: none !important;
    }
    .stButton > button[kind="primary"],
    .stFormSubmitButton > button[kind="primary"],
    .stDownloadButton > button {
        background: var(--ec-navy) !important;
        color: #FFFFFF !important;
        border: 1px solid var(--ec-navy) !important;
    }
    .stButton > button[kind="secondary"] {
        background: #FFFFFF !important;
        color: var(--ec-navy) !important;
        border: 1px solid var(--ec-line) !important;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover,
    .stFormSubmitButton > button:hover {
        opacity: 0.92;
    }

    div[data-testid="stAlert"] {
        border: 1px solid var(--ec-line) !important;
        box-shadow: var(--ec-shadow);
    }

    .stDataFrame, [data-testid="stDataFrame"] {
        border: 1px solid var(--ec-line) !important;
        box-shadow: var(--ec-shadow);
        overflow: hidden;
        background: var(--ec-card) !important;
    }
    th, td {
        font-size: 14px !important;
        color: var(--ec-slate) !important;
    }

    [data-testid="stFileUploader"] section {
        border-radius: var(--ec-radius) !important;
        border: 1px dashed rgba(31, 78, 121, 0.35) !important;
        background: var(--ec-accent-soft) !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: var(--ec-navy) !important;
    }

    [data-testid="stMetric"] {
        background: var(--ec-card);
        border: 1px solid var(--ec-line);
        border-radius: var(--ec-radius);
        padding: 0.75rem 0.9rem;
        box-shadow: var(--ec-shadow);
    }
    [data-testid="stMetricValue"] {
        font-weight: 700 !important;
        color: var(--ec-navy) !important;
    }

    .stProgress > div > div > div > div {
        background: var(--ec-navy) !important;
        border-radius: 4px !important;
    }

    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb {
        background: rgba(100, 116, 139, 0.35);
        border-radius: 8px;
        border: 2px solid transparent;
        background-clip: content-box;
    }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header[data-testid="stHeader"] { height: 2.4rem; }

    a { color: var(--ec-navy-soft) !important; text-decoration: none !important; }
    a:hover { text-decoration: underline !important; }

    code, pre {
        border-radius: 4px !important;
        font-size: 13px !important;
    }

    /* Ocultar SOLO Deploy/Desplegar (File change / Recargar quedan) */
    [data-testid="stAppDeployButton"],
    [data-testid="stDeployButton"],
    .stDeployButton,
    div[data-testid="stToolbar"] a[href*="share.streamlit"] {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
        pointer-events: none !important;
    }

    /* Chips Biblioteca/Menú: fijos arriba a la derecha.
       NO usar position:absolute + width/height 0 en wrappers: eso interceptaba
       clicks del nav (radio/botones) y dejaba el módulo desincronizado. */
    div[data-testid="stHorizontalBlock"]:has(.st-key-ec_toolbar_bib):has(.st-key-ec_toolbar_menu) {
        position: fixed !important;
        top: 0.35rem !important;
        right: 2.75rem !important;
        z-index: 999990 !important;
        width: max-content !important;
        max-width: none !important;
        height: auto !important;
        margin: 0 !important;
        padding: 0 !important;
        gap: 0.85rem !important;
        column-gap: 0.85rem !important;
        row-gap: 0 !important;
        align-items: center !important;
        justify-content: flex-end !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        pointer-events: none;
    }
    div[data-testid="stHorizontalBlock"]:has(.st-key-ec_toolbar_bib):has(.st-key-ec_toolbar_menu)
      > div[data-testid="stColumn"] {
        width: auto !important;
        flex: 0 0 auto !important;
        min-width: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        pointer-events: auto;
    }
    /* Separación extra entre chips (Streamlit a veces comprime el gap de columns). */
    div[data-testid="stHorizontalBlock"]:has(.st-key-ec_toolbar_bib):has(.st-key-ec_toolbar_menu)
      > div[data-testid="stColumn"]:has(.st-key-ec_toolbar_menu) {
        margin-left: 0.35rem !important;
    }
    /* Colapsar solo el alto reservado en el flujo (sin overlay absoluto).
       overflow:visible para que los popovers fijos sigan clickeables. */
    div[data-testid="stElementContainer"]:has(.st-key-ec_toolbar_bib),
    div[data-testid="stElementContainer"]:has(.st-key-ec_toolbar_menu),
    div[data-testid="element-container"]:has(.st-key-ec_toolbar_bib),
    div[data-testid="element-container"]:has(.st-key-ec_toolbar_menu) {
        margin-bottom: 0 !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }
    .st-key-ec_toolbar_bib,
    .st-key-ec_toolbar_menu {
        width: auto !important;
        margin: 0 !important;
        pointer-events: auto;
    }
    .st-key-ec_toolbar_bib button,
    .st-key-ec_toolbar_menu button,
    .st-key-ec_toolbar_bib [data-testid="stPopoverButton"] button,
    .st-key-ec_toolbar_menu [data-testid="stPopoverButton"] button {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-sizing: border-box !important;
        min-height: 2rem !important;
        height: 2rem !important;
        padding: 0 0.95rem !important;
        font-size: 0.8125rem !important;
        font-weight: 500 !important;
        letter-spacing: 0.01em !important;
        line-height: 1 !important;
        border-radius: 0.5rem !important;
        background: #FFFFFF !important;
        color: #1F4E79 !important;
        border: 1px solid #D0D7DE !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04) !important;
        white-space: nowrap !important;
        transition: background 0.12s ease, border-color 0.12s ease !important;
    }
    .st-key-ec_toolbar_bib button:hover,
    .st-key-ec_toolbar_menu button:hover,
    .st-key-ec_toolbar_bib [data-testid="stPopoverButton"] button:hover,
    .st-key-ec_toolbar_menu [data-testid="stPopoverButton"] button:hover {
        background: #F8FAFC !important;
        border-color: #94A3B8 !important;
    }
    /* Nav principal (ventanas): pills espaciadas, coherente con chips Biblioteca/Menú */
    .st-key-ventana_principal_v3,
    div[data-testid="stElementContainer"]:has(.st-key-ventana_principal_v3),
    div[data-testid="element-container"]:has(.st-key-ventana_principal_v3) {
        margin: 0.15rem 0 0.35rem 0 !important;
    }
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"],
    .st-key-ventana_principal_v3 [data-baseweb="button-group"],
    .st-key-ventana_principal_v3 [role="radiogroup"] {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 0.45rem !important;
        column-gap: 0.45rem !important;
        row-gap: 0.45rem !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
    }
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] label,
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] button,
    .st-key-ventana_principal_v3 [data-baseweb="button-group"] button,
    .st-key-ventana_principal_v3 [role="radiogroup"] label,
    .st-key-ventana_principal_v3 [role="radiogroup"] button {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-sizing: border-box !important;
        min-height: 2.15rem !important;
        height: auto !important;
        padding: 0.45rem 1.05rem !important;
        margin: 0 !important;
        font-size: 0.875rem !important;
        font-weight: 500 !important;
        letter-spacing: 0.01em !important;
        line-height: 1.15 !important;
        white-space: nowrap !important;
        border-radius: 0.55rem !important;
        background: #FFFFFF !important;
        color: #1F4E79 !important;
        border: 1px solid #D0D7DE !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04) !important;
        transition: background 0.12s ease, border-color 0.12s ease, color 0.12s ease !important;
    }
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] label:hover,
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] button:hover,
    .st-key-ventana_principal_v3 [data-baseweb="button-group"] button:hover,
    .st-key-ventana_principal_v3 [role="radiogroup"] label:hover,
    .st-key-ventana_principal_v3 [role="radiogroup"] button:hover {
        background: #F8FAFC !important;
        border-color: #94A3B8 !important;
    }
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] label[data-checked="true"],
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] label[aria-checked="true"],
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] button[aria-checked="true"],
    .st-key-ventana_principal_v3 [data-testid="stSegmentedControl"] button[aria-pressed="true"],
    .st-key-ventana_principal_v3 [data-baseweb="button-group"] button[aria-checked="true"],
    .st-key-ventana_principal_v3 [role="radiogroup"] label[aria-checked="true"],
    .st-key-ventana_principal_v3 [role="radiogroup"] button[aria-checked="true"],
    .st-key-ventana_principal_v3 [role="radiogroup"] [data-checked="true"] {
        background: #1F4E79 !important;
        color: #FFFFFF !important;
        border-color: #1F4E79 !important;
        box-shadow: 0 1px 3px rgba(31, 78, 121, 0.28) !important;
        font-weight: 600 !important;
    }
    .main .block-container {
        padding-top: 0.85rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def _inyectar_autoheal_frontend() -> None:
    """Oculta/elimina banners rojos de removeChild (bug Streamlit/React) sin ensuciar la UI.

    El script corre en un iframe de components.html → opera sobre parent.document.
    No recarga la página: eso dejaba varios errores rojos iguales en pantalla.
    """
    import streamlit.components.v1 as components

    components.html(
        """
<script>
(function () {
  function rootDoc() {
    try { return window.parent && window.parent.document ? window.parent.document : document; }
    catch (e) { return document; }
  }
  function rootWin() {
    try { return window.parent || window; }
    catch (e) { return window; }
  }

  const w = rootWin();
  if (w.__ecAutohealArmed) return;
  w.__ecAutohealArmed = true;

  function isNoise(msg) {
    return /removeChild|NotFoundError|Error al ejecutar ['"]removeChild['"]|nodo que se va a eliminar/i.test(
      String(msg || "")
    );
  }

  function hideDeploy() {
    try {
      const doc = rootDoc();
      const kill = function (el) {
        if (!el) return;
        el.style.setProperty("display", "none", "important");
        el.style.setProperty("visibility", "hidden", "important");
        el.style.setProperty("width", "0", "important");
        el.style.setProperty("height", "0", "important");
        el.style.setProperty("margin", "0", "important");
        el.style.setProperty("padding", "0", "important");
        el.style.setProperty("overflow", "hidden", "important");
        el.style.setProperty("pointer-events", "none", "important");
      };
      doc.querySelectorAll(
        '[data-testid="stAppDeployButton"], [data-testid="stDeployButton"], .stDeployButton'
      ).forEach(kill);
      doc.querySelectorAll(
        '[data-testid="stToolbar"] button, [data-testid="stToolbar"] a, [data-testid="stHeader"] button'
      ).forEach(function (el) {
        const t = (
          (el.innerText || "") +
          " " +
          (el.textContent || "") +
          " " +
          (el.getAttribute("title") || "") +
          " " +
          (el.getAttribute("aria-label") || "")
        );
        if (/deploy|desplegar/i.test(t)) kill(el);
      });
    } catch (e) {}
  }

  function mountEcChipsToHeader() {
    /* No mover nodos React (rompe el main). Solo CSS fixed + hideDeploy. */
    return;
  }

  function scrub() {
    try {
      hideDeploy();
      mountEcChipsToHeader();
      const doc = rootDoc();
      const sels = [
        '[data-testid="stException"]',
        '.stException',
        '[data-testid="stAlert"]',
        '[kind="error"]',
        'div[role="alert"]',
      ];
      doc.querySelectorAll(sels.join(",")).forEach(function (el) {
        const t = (el.innerText || el.textContent || "");
        if (isNoise(t)) {
          // Solo ocultar: NO el.remove() — romper nodos React desincroniza el nav/módulo.
          el.style.setProperty("display", "none", "important");
          el.style.setProperty("height", "0", "important");
          el.style.setProperty("margin", "0", "important");
          el.style.setProperty("padding", "0", "important");
          el.style.setProperty("overflow", "hidden", "important");
          el.setAttribute("data-ec-hidden-noise", "1");
        }
      });
    } catch (e) {}
  }

  w.addEventListener("error", function (e) {
    if (isNoise((e && (e.message || (e.error && e.error.message))) || "")) {
      try { e.preventDefault(); e.stopImmediatePropagation(); } catch (err) {}
      scrub();
      return true;
    }
  }, true);
  w.addEventListener("unhandledrejection", function (e) {
    const r = e && e.reason;
    if (isNoise((r && (r.message || String(r))) || "")) {
      try { e.preventDefault(); e.stopImmediatePropagation(); } catch (err) {}
      scrub();
    }
  }, true);

  try {
    const doc = rootDoc();
    const obs = new MutationObserver(function () { scrub(); });
    function arm() {
      if (!doc.body) return;
      obs.observe(doc.body, { childList: true, subtree: true });
      scrub();
      setInterval(scrub, 700);
    }
    if (doc.body) arm();
    else doc.addEventListener("DOMContentLoaded", arm);
  } catch (e) {}
})();
</script>
        """,
        height=0,
        width=0,
    )


_inyectar_autoheal_frontend()

db.inicializar_bd()

NUEVOS_MONOTRIBUTISTAS: list[dict[str, str]] = [
    {"nombre": "PERNAS ROSARIO", "cuit": "27274162282", "tipo": "Monotributista"},
    {"nombre": "SELVA, MAXIMILIANO", "cuit": "20253936755", "tipo": "Monotributista"},
    {"nombre": "SCARPELLO, PAOLA", "cuit": "27282992707", "tipo": "Monotributista"},
    {"nombre": "MONCHIETTI, TRISTANA", "cuit": "27219760677", "tipo": "Monotributista"},
    {"nombre": "MONCHIETTI, TITO", "cuit": "20244442278", "tipo": "Monotributista"},
    {"nombre": "PANASCI, ANGELES MARIA", "cuit": "27255699038", "tipo": "Monotributista"},
    {"nombre": "VARELA, GONZALO", "cuit": "23255699059", "tipo": "Monotributista"},
    {"nombre": "DIEGO LAWRIE", "cuit": "23283669009", "tipo": "Monotributista"},
    {"nombre": "GUTIERREZ CARLOS", "cuit": "20056216848", "tipo": "Monotributista"},
    {"nombre": "BAZAN CARLOS", "cuit": "20114552802", "tipo": "Monotributista"},
    {"nombre": "AGUSTINA RACITI", "cuit": "27363830523", "tipo": "Monotributista"},
    {"nombre": "ADRIANA DAVICO", "cuit": "23167407234", "tipo": "Monotributista"},
    {"nombre": "RODRIGUEZ ASTUNO PAMELA", "cuit": "27304498612", "tipo": "Monotributista"},
    {"nombre": "PARODI DOLORES", "cuit": "27294424917", "tipo": "Monotributista"},
    {"nombre": "PARODI MANUEL", "cuit": "20254298876", "tipo": "Monotributista"},
    {"nombre": "PARODI RICARDO", "cuit": "20053337458", "tipo": "Monotributista"},
    {"nombre": "DEPREZ PRUVOST SERGIO", "cuit": "20188117105", "tipo": "Monotributista"},
    {"nombre": "REBORA JORGELINA", "cuit": "27245392848", "tipo": "Monotributista"},
    {"nombre": "NEGROTTI PABLO", "cuit": "20242515901", "tipo": "Monotributista"},
    {"nombre": "QUILOGRAN LEILA", "cuit": "27434562967", "tipo": "Monotributista"},
    {"nombre": "LAZO LUCIANO SEBASTIAN", "cuit": "20453005829", "tipo": "Monotributista"},
    {"nombre": "MUÑOZ NAHUEL ANDRES", "cuit": "20445626512", "tipo": "Monotributista"},
    {"nombre": "CRISTIANO FERNANDO PASCUAL", "cuit": "20312465877", "tipo": "Monotributista"},
    {"nombre": "ACHE CLAUDIO DAVID", "cuit": "20230872083", "tipo": "Monotributista"},
    {"nombre": "COLOMBO GUILLERMO", "cuit": "23263466373", "tipo": "Monotributista"},
    {"nombre": "DIAZ SILVIA BEATRIZ", "cuit": "27177966245", "tipo": "Monotributista"},
    {"nombre": "ERRECABORDE ROCIO", "cuit": "27312641556", "tipo": "Monotributista"},
    {"nombre": "FARIAS JONATAN", "cuit": "20338985356", "tipo": "Monotributista"},
    {"nombre": "FARIAS MARCELINO", "cuit": "20170178093", "tipo": "Monotributista"},
    {"nombre": "FARIAS YESICA DAIANA", "cuit": "27326618115", "tipo": "Monotributista"},
    {"nombre": "CARE ANDREA", "cuit": "27238186043", "tipo": "Monotributista"},
    {"nombre": "FERNANDEZ M MAIRA", "cuit": "27329884037", "tipo": "Monotributista"},
    {"nombre": "GOMEZ SILVINA", "cuit": "27273795559", "tipo": "Monotributista"},
    {"nombre": "LUCCA MAURO", "cuit": "20318501492", "tipo": "Monotributista"},
    {"nombre": "RODRIGUEZ FLORENCIA", "cuit": "27321040484", "tipo": "Monotributista"},
    {"nombre": "VIVAS NATALIA SOLEDAD", "cuit": "27294955955", "tipo": "Monotributista"},
    {"nombre": "GIVONETTI CAROLINA", "cuit": "27222929941", "tipo": "Monotributista"},
    {"nombre": "ALBARELLO CAMILA", "cuit": "27420430340", "tipo": "Monotributista"},
    {"nombre": "MERODIO JIMENA", "cuit": "27398502529", "tipo": "Monotributista"},
    {"nombre": "NICOLAO LUCRECIA", "cuit": "27362171003", "tipo": "Monotributista"},
    {"nombre": "WERTHEIMER SOFIA", "cuit": "27447596054", "tipo": "Monotributista"},
    {"nombre": "GALLO ANA ESTHER", "cuit": "27223606852", "tipo": "Monotributista"},
    {"nombre": "GARCIA MARIELA ALEJANDRA", "cuit": "27237063991", "tipo": "Monotributista"},
    {"nombre": "LIGORE PEDRO", "cuit": "20345016326", "tipo": "Monotributista"},
    {"nombre": "LOPEZ JOAQUIN", "cuit": "20210277782", "tipo": "Monotributista"},
    {"nombre": "MIGNANI SOLANA", "cuit": "27270834227", "tipo": "Monotributista"},
    {"nombre": "MORTEO JORGELINA", "cuit": "27264194232", "tipo": "Monotributista"},
    {"nombre": "ARIAS CARLOS", "cuit": "20083720140", "tipo": "Monotributista"},
    {"nombre": "BUNGE AGUSTIN", "cuit": "20363835741", "tipo": "Monotributista"},
    {"nombre": "BUSTAMANTE AGUSTINA", "cuit": "27392831636", "tipo": "Monotributista"},
    {"nombre": "FRAYSSINET FERMIN", "cuit": "20290677301", "tipo": "Monotributista"},
    {"nombre": "PETERSEN MARIA JOSE", "cuit": "27232246419", "tipo": "Monotributista"},
    {"nombre": "KELLER LORENA", "cuit": "23221386124", "tipo": "Monotributista"},
    {"nombre": "PEREZ MARTA", "cuit": "27064978336", "tipo": "Monotributista"},
    {"nombre": "SIRI EMILIO", "cuit": "20291182659", "tipo": "Monotributista"},
    {"nombre": "BERAZA PATRICIA EDITH", "cuit": "27163370935", "tipo": "Monotributista"},
    {"nombre": "SANTIAGO MONACCHI", "cuit": "20429550719", "tipo": "Monotributista"},
    {"nombre": "ETTIEN LEROUX", "cuit": "20274162466", "tipo": "Monotributista"},
    {"nombre": "PAL YAMILA", "cuit": "27326686153", "tipo": "Monotributista"},
    {"nombre": "DOMINGUEZ MARCELO", "cuit": "20168259779", "tipo": "Monotributista"},
]
db.sincronizar_clientes_catalogo(NUEVOS_MONOTRIBUTISTAS)
# Cloud / repo limpio: siembra PJ + planes desde data/seed y data/planes_cuentas
_cargar_seed_pj = getattr(db, "cargar_seed_sociedades_pj", None)
if callable(_cargar_seed_pj):
    _cargar_seed_pj()
else:
    # Fallback si Cloud sirve database.py viejo sin el método (redeploy a medias)
    _seed_pj = Path(__file__).resolve().parent / "data" / "seed" / "sociedades_pj.json"
    if _seed_pj.is_file():
        try:
            _data_pj = json.loads(_seed_pj.read_text(encoding="utf-8"))
            if isinstance(_data_pj, list):
                db.sincronizar_clientes_catalogo(_data_pj)
        except (OSError, json.JSONDecodeError):
            pass

_SOCiedad_KEY = "sociedad_activa"
_IMPUESTO_KEY = "selector_impuesto"
_MODULO_TRABAJO_KEY = "selector_modulo_trabajo"
_BANCO_KEY = "selector_banco_conciliar_v2"
_RUTA_BALANCE_BANCOS_INPUT = "ruta_balance_bancos_input"
_BTN_CARGAR_SALDO_BANCOS = "btn_cargar_saldo_bancos"
_MODULOS_TRABAJO = ("Devengamiento de Impuestos", "Conciliación Bancaria")
_VENTANAS_PRINCIPALES = (
    "Devengamiento de Impuestos",
    "Conciliación Bancaria",
    "Préstamos Financieros",
    "Recategorización Monotributo",
    "Herramientas",
)
# v3: botones con keys fijas (el radio + CSS absolute del toolbar desincronizaba UI↔módulo)
_VENTANA_KEY = "ventana_principal_v3"
_VENTANA_KEY_LEGACY = "ventana_principal_activa"
_VENTANA_NAV_LABELS = {
    "Devengamiento de Impuestos": "Devengamiento",
    "Conciliación Bancaria": "Conciliación",
    "Préstamos Financieros": "Préstamos",
    "Recategorización Monotributo": "Monotributo",
    "Herramientas": "Herramientas",
}


def _es_ventana_herramientas(nombre: str | None) -> bool:
    return "Herramientas" in str(nombre or "")


def _migrar_ventana_principal_session() -> None:
    """Normaliza labels viejos / key legacy hacia ventana_principal_v3."""
    actual = st.session_state.get(_VENTANA_KEY)
    if not actual and st.session_state.get(_VENTANA_KEY_LEGACY):
        actual = st.session_state.get(_VENTANA_KEY_LEGACY)
    actual = str(actual or "")
    # Alias legacy (emoji / nombres cortos) → canónicos sin iconos
    if actual in ("🧰 Herramientas", "Herramientas", "💰 Liquidación de Sueldos"):
        actual = "Herramientas"
    elif actual in ("📊 Recategorización Monotributo", "Recategorización Monotributo"):
        actual = "Recategorización Monotributo"
    if actual not in _VENTANAS_PRINCIPALES:
        actual = _VENTANAS_PRINCIPALES[0]
    st.session_state[_VENTANA_KEY] = actual
    # Evitar que un radio legacy (si queda en session) pelee con los botones
    st.session_state.pop(_VENTANA_KEY_LEGACY, None)


def _render_nav_ventanas_principales() -> str:
    """Un solo segmented_control keyed al nombre canónico del módulo (sin mapeo intermedio)."""
    _migrar_ventana_principal_session()
    opciones = list(_VENTANAS_PRINCIPALES)
    actual = str(st.session_state.get(_VENTANA_KEY, opciones[0]))
    if actual not in opciones:
        actual = opciones[0]
        st.session_state[_VENTANA_KEY] = actual

    # Limpiar key intermedia de intentos previos (label corto) que desincronizaba.
    st.session_state.pop("nav_ventana_segmented_v1", None)

    elegido = st.segmented_control(
        "Módulo",
        options=opciones,
        format_func=lambda x: _VENTANA_NAV_LABELS.get(x, x),
        key=_VENTANA_KEY,
        label_visibility="collapsed",
        required=True,
    )
    if elegido not in opciones:
        elegido = actual
        st.session_state[_VENTANA_KEY] = elegido
    return str(elegido)

_DEBUG_LOG_PATH = BASE_DIR / "debug-46b61e.log"
STREAMLIT_PUERTO_DEFAULT = 8501
# Arranque con tolerancia de red: iniciar_estudio_contable.bat
# streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.disconnectedSessionTTL 3600 --server.websocketPingInterval 10


def obtener_ip_red() -> str:
    """IP privada de la PC en la red local (para compartir la app en la oficina)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 1))
        return s.getsockname()[0]
    except OSError:
        return "localhost"
    finally:
        s.close()


def _obtener_puerto_streamlit() -> int:
    try:
        from streamlit import config as st_config
        return int(st_config.get_option("server.port"))
    except Exception:
        return STREAMLIT_PUERTO_DEFAULT


def _render_panel_compartir_red_local(*, en_sidebar: bool = False) -> None:
    """Link Network URL copiable (sidebar legacy o barra superior)."""
    destino = st.sidebar if en_sidebar else st
    with destino.expander("Compartir en la red local", expanded=False):
        ip = obtener_ip_red()
        puerto = _obtener_puerto_streamlit()
        network_url = f"http://{ip}:{puerto}"
        st.code(network_url, language=None)
        st.caption(
            "Copiá este link y compartilo. "
            "Ambos deben estar conectados al mismo Wi-Fi de la oficina."
        )


def _listar_excels_exportaciones_recientes(limite: int = 12) -> list[Path]:
    """Últimos Excels generados en la carpeta exportaciones/ del proyecto."""
    carpeta = BASE_DIR / "exportaciones"
    if not carpeta.is_dir():
        return []
    archivos = [
        p for p in carpeta.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    archivos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return archivos[:limite]


def _resumen_biblioteca_sociedad_activa() -> tuple[int, dict[str, list], int, dict[str, list]]:
    """Totales de biblioteca (impuestos + bancos) para la sociedad activa."""
    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    por_imp: dict[str, list] = {}
    por_banco: dict[str, list] = {}
    if sociedad_id is None:
        return 0, por_imp, 0, por_banco
    for e in st.session_state.get("biblioteca_asientos") or []:
        if e.get("sociedad_id") != sociedad_id:
            continue
        imp = str(e.get("impuesto") or "IVA")
        por_imp.setdefault(imp, []).append(e)
    for e in st.session_state.get("biblioteca_bancos") or []:
        if e.get("sociedad_id") != sociedad_id:
            continue
        banco = str(e.get("banco") or "Banco")
        por_banco.setdefault(banco, []).append(e)
    total_asi = sum(len(v) for v in por_imp.values())
    total_bco = sum(len(v) for v in por_banco.values())
    return total_asi, por_imp, total_bco, por_banco


def _render_titulo_estudio() -> None:
    st.markdown(
        '<p style="margin:0.1rem 0 1.05rem 0;font-size:1.55rem;font-weight:700;color:#1F4E79;'
        'letter-spacing:-0.02em;line-height:1.2;">Estudio Contable</p>',
        unsafe_allow_html=True,
    )


def _render_barra_superior_cuenta() -> None:
    """Biblioteca+Menú anclados al header Streamlit (CSS fixed, sin hueco)."""
    nombre_u = (
        st.session_state.get("usuario_oficina_nombre")
        or st.session_state.get("usuario_oficina")
        or "—"
    )
    total_asi, por_imp, total_bco, por_banco = _resumen_biblioteca_sociedad_activa()
    total_meses = total_asi + total_bco
    label_bib = f"Biblioteca · {total_meses}"

    # Fila mínima: solo chips (CSS las sube al header y colapsa el hueco del main)
    col_bib, col_menu = st.columns([1.2, 1.0], gap="medium")
    with col_bib:
        with st.popover(label_bib, key="ec_toolbar_bib", width="content"):
            sociedad_id = st.session_state.get(_SOCiedad_KEY)
            if sociedad_id is None:
                st.caption("Elegí una sociedad en la barra lateral.")
            else:
                st.caption(
                    f"**Meses archivados:** {total_asi} (impuestos) · {total_bco} (bancos)"
                )
                if por_imp:
                    st.markdown("**Devengamientos**")
                    for imp, entradas in sorted(por_imp.items()):
                        periodos = sorted(
                            {
                                _periodo_a_etiqueta(str(e.get("periodo") or ""))
                                for e in entradas
                                if e.get("periodo")
                            }
                        )
                        st.caption(
                            f"{imp}: {len(entradas)}"
                            + (f" · [{', '.join(periodos)}]" if periodos else "")
                        )
                    with st.expander("Vaciar impuestos", expanded=False):
                        for imp, entradas in sorted(por_imp.items()):
                            if not entradas:
                                continue
                            if st.button(
                                f"Vaciar {imp}",
                                key=f"btn_vaciar_bib_top_{_slug_impuesto(imp)}",
                                type="secondary",
                                use_container_width=True,
                            ):
                                _vaciar_biblioteca_sociedad(sociedad_id, imp)
                                st.rerun()
                else:
                    st.caption("Sin asientos de impuestos archivados.")

                if por_banco:
                    st.markdown("**Bancos**")
                    for banco, entradas in sorted(por_banco.items()):
                        periodos = sorted(
                            {
                                _periodo_a_etiqueta(str(e.get("periodo") or ""))
                                for e in entradas
                                if e.get("periodo")
                            }
                        )
                        st.caption(
                            f"{banco}: {len(entradas)}"
                            + (f" · [{', '.join(periodos)}]" if periodos else "")
                        )
                    with st.expander("Vaciar bancos", expanded=False):
                        for banco in sorted(por_banco.keys()):
                            if st.button(
                                f"Vaciar {banco}",
                                key=f"btn_vaciar_bib_banco_top_{_slug_banco(banco)}",
                                type="secondary",
                                use_container_width=True,
                            ):
                                _vaciar_biblioteca_banco_sociedad(sociedad_id, banco)
                                st.rerun()
                st.caption(
                    f"`{ruta_biblioteca_persistida_activa(usuario=_usuario_oficina_actual())}`"
                )

    with col_menu:
        with st.popover("☰ Menú", key="ec_toolbar_menu", width="content"):
            st.caption(f"Usuario: **{nombre_u}**")
            if st.button("Cerrar sesión", use_container_width=True, key="btn_logout_oficina"):
                _cerrar_sesion_oficina()
                st.rerun()

            st.markdown("---")
            if not _es_entorno_cloud():
                st.markdown("**Compartir en la red local**")
                ip = obtener_ip_red()
                puerto = _obtener_puerto_streamlit()
                st.code(f"http://{ip}:{puerto}", language=None)
                st.caption("Mismo Wi-Fi de la oficina para compartirlo.")
                st.markdown("---")
            else:
                try:
                    from seguridad_datos import estado_cifrado_ui

                    est = estado_cifrado_ui()
                    if est.get("fernet_ok"):
                        st.caption("Cifrado at-rest: activo")
                    else:
                        st.caption("Cifrado at-rest: falta DATA_ENCRYPTION_KEY")
                except Exception:
                    pass
                st.markdown("---")

            st.markdown("**Administración**")
            if st.button(
                "Gestión de Clientes",
                use_container_width=True,
                key="btn_admin_clientes_top",
            ):
                st.session_state.vista_admin = "clientes"
                st.rerun()
            if st.session_state.get("usuario_oficina_admin") and st.button(
                "Usuarios de la oficina",
                use_container_width=True,
                key="btn_admin_usuarios_top",
            ):
                st.session_state.vista_admin = "usuarios_oficina"
                st.rerun()
            if st.button("Acerca de", use_container_width=True, key="btn_admin_acerca_top"):
                st.session_state.vista_admin = "acerca"
                st.rerun()
            if st.session_state.get("vista_admin") and st.button(
                "← Volver a procesos",
                use_container_width=True,
                key="btn_admin_volver_top",
            ):
                st.session_state.vista_admin = None
                st.rerun()

            st.markdown("---")
            st.markdown("**Opciones avanzadas**")
            modo_auto = st.checkbox(
                "Ingesta automática desde carpeta (por CUIT)",
                value=st.session_state.modo_ingesta_automatica,
                key="chk_ingesta_auto_top",
                help=(
                    "Si está activo, al procesar se escanea clientes/{CUIT}/ "
                    "en lugar de usar archivos subidos."
                ),
            )
            if modo_auto != st.session_state.modo_ingesta_automatica:
                st.session_state.modo_ingesta_automatica = modo_auto
                st.session_state.resultado = None
                st.session_state.mov_banco = None
            if st.session_state.modo_ingesta_automatica:
                nueva_ruta = st.text_input(
                    "Ruta raíz de clientes",
                    value=st.session_state.ruta_raiz_clientes,
                    key="ruta_raiz_clientes_top",
                )
                if nueva_ruta != st.session_state.ruta_raiz_clientes:
                    st.session_state.ruta_raiz_clientes = nueva_ruta


def _render_sidebar_sociedad_y_biblioteca(cliente: dict | None) -> None:
    """Barra lateral: solo sociedad activa (biblioteca va al toolbar superior)."""
    st.session_state["_sidebar_unificada"] = True
    st.sidebar.markdown("### Sociedad de trabajo")

    clientes = db.listar_clientes()
    if clientes:
        opciones = {f"{c['nombre']} ({c['cuit']})": c["id"] for c in clientes}
        ids = list(opciones.values())
        labels = list(opciones.keys())
        actual = st.session_state.get(_SOCiedad_KEY)
        idx = ids.index(actual) if actual in ids else 0
        # Mantener el selectbox del sidebar alineado con sociedad_activa
        # (si no, el key viejo pisa al selector del módulo y aparece "sin plan").
        label_deseado = labels[idx]
        if st.session_state.get("sidebar_sociedad_select_v2") != label_deseado:
            st.session_state["sidebar_sociedad_select_v2"] = label_deseado
        elegido_label = st.sidebar.selectbox(
            "Sociedad activa",
            labels,
            key="sidebar_sociedad_select_v2",
        )
        elegido_id = opciones[elegido_label]
        if elegido_id != st.session_state.get(_SOCiedad_KEY):
            st.session_state[_SOCiedad_KEY] = elegido_id
            actualizar_sociedad_activa()
            st.rerun()
        cli = db.obtener_cliente(elegido_id) or cliente
    else:
        cli = cliente
        st.sidebar.warning("No hay clientes cargados.")

    if cli:
        st.sidebar.caption(
            f"**{cli.get('nombre') or '—'}** · CUIT `{cli.get('cuit') or '—'}` · "
            f"{cli.get('tipo_persona') or '—'}"
        )
    else:
        st.sidebar.warning("Sin sociedad seleccionada")

    st.sidebar.caption("Usá **Biblioteca** y **Menú** en la barra superior.")



def _dbg_log(hypothesis_id: str, location: str, message: str, data: dict | None = None) -> None:
    # #region agent log
    try:
        payload = {
            "sessionId": "46b61e",
            "hypothesisId": hypothesis_id,
            "location": location,
            "message": message,
            "data": data or {},
            "timestamp": int(time.time() * 1000),
        }
        with open(_DEBUG_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except OSError:
        pass
    # #endregion


def _init_session_state() -> None:
    """Inicializa variables de sesión persistentes."""
    defaults = {
        "sociedad_activa": None,
        "selector_impuesto": "IVA",
        "cliente_id_seleccionado": None,
        "resultado": None,
        "plan_cuentas": None,
        "mov_banco": None,
        "cliente_activo": None,
        "bancos_detectados": [],
        "modo_ingesta_automatica": False,
        "ruta_raiz_clientes": str(RUTA_RAIZ_CLIENTES),
        "devengamiento_resultado": None,
        "devengamiento_datos": None,
        "saldos_iniciales_bancos": [{"banco": BANCOS_ARGENTINOS[0], "saldo_inicial": 0.0}],
        "auditoria_prestamos_resultado": None,
        "vista_admin": None,
        "iva_grilla_preview": None,
        "iva_asientos_generados": None,
        "plan_cuentas_df": None,
        "plan_cuentas_cliente_id": None,
        "plan_cuentas_path_resuelto": None,
        "plan_cuentas_es_default": False,
        "cuit_activo": None,
        "nombre_activo": None,
        "biblioteca_asientos": [],
        "biblioteca_bancos": [],
        "periodos_procesados": [],
        "periodos_bancos_procesados": [],
        "reset_counter": 0,
        "iva_skip_default_planilla": False,
        "saldos_tecnicos_list": [0.0],
        "saldos_libre_list": [0.0],
        "saldos_favor_iibb_list": [0.0],
        "saldos_favor_cm_list": [0.0],
        "saldos_favor_sueldos_list": [0.0],
        "saldos_favor_tish_list": [0.0],
        "balance_servidor_rutas_por_sociedad": {},
        "balance_servidor_buffer_por_sociedad": {},
        "balance_servidor_sync_at_por_sociedad": {},
        "iibb_grilla_preview": None,
        "sueldos_grilla_preview": None,
        "sueldos_asientos_generados": None,
        "sueldos_resumen_analitico": None,
        "sueldos_mensaje_biblioteca_ok": None,
        "sueldos_limpiar_formulario_pendiente": None,
        "sueldos_planilla_ref_fecha": None,
        "sueldos_skip_default_planilla": False,
        "sueldos_saldos_rc_active": None,
        "tish_grilla_preview": None,
        "tish_asientos_generados": None,
        "tish_resumen_analitico": None,
        "tish_mensaje_biblioteca_ok": None,
        "tish_limpiar_formulario_pendiente": None,
        "tish_planilla_ref_fecha": None,
        "tish_skip_default_planilla": False,
        "tish_saldos_rc_active": None,
        "iibb_asientos_generados": None,
        "iibb_resumen_analitico": None,
        "iibb_mensaje_biblioteca_ok": None,
        "iibb_limpiar_formulario_pendiente": None,
        "iibb_planilla_ref_fecha": None,
        "iibb_skip_default_planilla": False,
        "iibb_saldos_rc_active": None,
        "_impuesto_previo_devengamientos": None,
        "_modulo_trabajo_previo": None,
        "_banco_previo_devengamientos": None,
        "_persistencia_hidratada": False,
        "_borrador_recuperado_slug": None,
        "idx_debe": None,
        "idx_haber": None,
        "selector_modulo_trabajo": "Devengamiento de Impuestos",
        "ventana_principal_v3": "Devengamiento de Impuestos",
        "selector_banco_conciliar_v2": "Santander",
        "mono_facturas_df": None,
        "mono_errores_extraccion": [],
        "extracto_santander_df": None,
        "extracto_santander_meta": {},
        "extracto_santander_errores": [],
        "extracto_santander_pdf_merged": None,
        "extracto_santander_xlsx": None,
        "cuadro_bancario_resultado": None,
        "cuadro_bancario_resultado_signature": None,
        "cuadro_bancario_buzon": None,
        "cuadro_bancario_buzon_ruta": None,
        "match_prov_resultado": None,
        "match_prov_xlsx": None,
        "match_prov_meta": {},
        "match_prov_errores": [],
        "usuario_oficina": None,
        "usuario_oficina_nombre": None,
        "usuario_oficina_admin": False,
    }
    for clave, valor in defaults.items():
        if clave not in st.session_state:
            st.session_state[clave] = valor
    if (
        "selector_banco_conciliar_v2" not in st.session_state
        and "selector_banco_conciliar" in st.session_state
    ):
        st.session_state["selector_banco_conciliar_v2"] = st.session_state.pop(
            "selector_banco_conciliar"
        )
    if "selector_impuesto" not in st.session_state and "impuesto_elegido" in st.session_state:
        st.session_state["selector_impuesto"] = st.session_state.pop("impuesto_elegido")
    if st.session_state.sociedad_activa is None and st.session_state.cliente_id_seleccionado is not None:
        st.session_state.sociedad_activa = st.session_state.cliente_id_seleccionado
    elif st.session_state.cliente_id_seleccionado is None and st.session_state.sociedad_activa is not None:
        st.session_state.cliente_id_seleccionado = st.session_state.sociedad_activa
    _hidratar_persistencia_desde_disco()


def _usuario_oficina_actual() -> str | None:
    return st.session_state.get("usuario_oficina")


def _persistir_biblioteca_en_disco() -> None:
    """Escribe biblioteca de devengamientos y bancos en JSON físico (por usuario)."""
    try:
        guardar_biblioteca_persistida(
            biblioteca_asientos=st.session_state.get("biblioteca_asientos") or [],
            biblioteca_bancos=st.session_state.get("biblioteca_bancos") or [],
            periodos_procesados=st.session_state.get("periodos_procesados") or [],
            periodos_bancos_procesados=st.session_state.get("periodos_bancos_procesados") or [],
            usuario=_usuario_oficina_actual(),
        )
    except OSError as exc:
        st.session_state["_persistencia_ultimo_error"] = str(exc)


def _hidratar_persistencia_desde_disco() -> None:
    """Al iniciar sesión Streamlit, precarga biblioteca archivada desde disco."""
    if st.session_state.get("_persistencia_hidratada"):
        return
    if not _usuario_oficina_actual():
        return
    datos = cargar_biblioteca_persistida(usuario=_usuario_oficina_actual())
    if datos:
        st.session_state.biblioteca_asientos = datos.get("biblioteca_asientos") or []
        st.session_state.biblioteca_bancos = datos.get("biblioteca_bancos") or []
        st.session_state.periodos_procesados = datos.get("periodos_procesados") or []
        st.session_state.periodos_bancos_procesados = datos.get("periodos_bancos_procesados") or []
        st.session_state["_persistencia_biblioteca_ruta"] = datos.get("ruta")
    st.session_state["_persistencia_hidratada"] = True


def _contexto_modulo_por_slug(slug: str) -> tuple[str, str]:
    impuesto = st.session_state.get(_IMPUESTO_KEY)
    if impuesto and _slug_impuesto(str(impuesto)) == slug:
        return "devengamiento", str(impuesto)
    banco = st.session_state.get(_BANCO_KEY)
    if banco and _slug_banco(str(banco)) == slug:
        return "banco", str(banco)
    ventana = str(st.session_state.get(_VENTANA_KEY, ""))
    if "Bancaria" in ventana or "banco" in slug:
        return "banco", slug
    return "devengamiento", slug


def _autosalvar_borrador_grilla(slug: str) -> None:
    """Auto-guardado de emergencia cada vez que la grilla cambia."""
    rows = st.session_state.get(f"{slug}_grilla_preview")
    if not rows:
        return
    modulo, contexto = _contexto_modulo_por_slug(slug)
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    try:
        guardar_borrador_grilla_persistido({
            "modulo": modulo,
            "slug": slug,
            "sociedad_id": st.session_state.get(_SOCiedad_KEY),
            "contexto": contexto,
            "grilla_preview": copy.deepcopy(rows),
            "asientos": copy.deepcopy(asientos),
            "resumen_analitico": copy.deepcopy(st.session_state.get(f"{slug}_resumen_analitico")),
            "columna_auditoria": copy.deepcopy(st.session_state.get(f"{slug}_columna_auditoria")),
            "auto_fp": st.session_state.get(f"{slug}_auto_fp"),
        }, usuario=_usuario_oficina_actual())
    except OSError:
        pass


def _marcar_grilla_dirty(slug: str) -> None:
    """Marca la grilla como editada; invalida Excel Tango cacheado."""
    st.session_state[f"{slug}_grilla_dirty"] = True
    st.session_state.pop(f"{slug}_tango_xlsx_bytes", None)
    st.session_state.pop(f"{slug}_tango_xlsx_fp", None)
    st.session_state.pop(f"{slug}_tango_xlsx_name", None)


def _grilla_esta_dirty(slug: str) -> bool:
    return bool(st.session_state.get(f"{slug}_grilla_dirty"))


def _limpiar_grilla_dirty(slug: str) -> None:
    st.session_state[f"{slug}_grilla_dirty"] = False


def _autosalvar_borrador_grilla_throttled(slug: str, min_interval_s: float = 3.0) -> None:
    """Autosave a disco con throttle para no trabar la UI en cada tecla."""
    import time

    now = time.monotonic()
    key = f"{slug}_last_autosave_ts"
    last = float(st.session_state.get(key) or 0.0)
    forzar = bool(st.session_state.pop(f"{slug}_force_autosave", False))
    if not forzar and (now - last) < min_interval_s:
        return
    st.session_state[key] = now
    _autosalvar_borrador_grilla(slug)


def _fingerprint_asientos_export(asientos: list) -> str:
    """Huella liviana de renglones para cachear el Excel Tango."""
    partes: list[str] = []
    for asiento in asientos or []:
        for renglon in getattr(asiento, "renglones", []) or []:
            partes.append(
                f"{getattr(renglon, 'codigo_cuenta', '')}|"
                f"{round(float(getattr(renglon, 'debe', 0) or 0), 2)}|"
                f"{round(float(getattr(renglon, 'haber', 0) or 0), 2)}"
            )
    return str(hash(tuple(partes)))


def _render_descarga_excel_tango_diferida(
    *,
    slug: str,
    ficha: dict,
    asientos: list,
    puede_exportar: bool,
    nombre_activo: str | None,
    cuit_activo: str | None,
    mes_ref: int,
    anio_ref: int,
    key_suffix: str,
    label_generar: str = "Generar Excel Tango (mes actual)",
    label_descargar: str = "Descargar Excel Tango — mes actual",
) -> None:
    """
    No regenera openpyxl en cada rerun: el Excel se arma solo al click
    de «Generar», y la descarga usa bytes cacheados.
    """
    fp = _fingerprint_asientos_export(asientos)
    bytes_key = f"{slug}_tango_xlsx_bytes"
    fp_key = f"{slug}_tango_xlsx_fp"
    name_key = f"{slug}_tango_xlsx_name"
    gen_key = f"btn_gen_tango_{slug}_{key_suffix}"

    c1, c2 = st.columns([1.2, 1])
    with c1:
        if not puede_exportar:
            st.button(
                label_generar,
                key=gen_key,
                use_container_width=True,
                disabled=True,
                help="Completá las cuentas 99999 de la grilla para habilitar.",
            )
        elif st.button(label_generar, key=gen_key, use_container_width=True):
            try:
                _marcar_grilla_dirty(slug)
                _finalizar_balance_grilla_slug(slug, ficha, forzar=True)
                asientos_act = st.session_state.get(f"{slug}_asientos_generados") or asientos
                plan_df = st.session_state.get("plan_cuentas_df")
                with st.spinner("Armando Excel Tango…"):
                    ruta_xl = generar_excel_tango_nativo(
                        asientos=asientos_act,
                        nombre_cliente=nombre_activo or "",
                        cuit=cuit_activo or "000",
                        mes=mes_ref,
                        anio=anio_ref,
                        plan_cuentas=plan_df,
                    )
                    st.session_state[bytes_key] = Path(ruta_xl).read_bytes()
                    st.session_state[fp_key] = _fingerprint_asientos_export(asientos_act)
                    st.session_state[name_key] = ruta_xl.name
                st.rerun()
            except ExportacionTangoError as exc:
                _mostrar_errores_exportacion_tango(exc)
            except Exception as exc:
                st.error(f"No se pudo generar el Excel: {exc}")
    with c2:
        data = st.session_state.get(bytes_key)
        if puede_exportar and data and st.session_state.get(fp_key) == fp:
            st.download_button(
                label_descargar,
                data=data,
                file_name=st.session_state.get(name_key) or f"asiento_{slug}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_tango_cached_{slug}_{key_suffix}",
                use_container_width=True,
            )
        elif not puede_exportar:
            st.caption("Hay cuentas sin vincular (99999).")
        else:
            st.caption("Generá el Excel para habilitar la descarga.")


def _restaurar_borrador_grilla_si_aplica(slug: str) -> bool:
    """Restaura borrador de disco si la grilla en memoria está vacía."""
    if st.session_state.get(f"{slug}_grilla_preview"):
        return False
    borrador = cargar_borrador_grilla_persistido(usuario=_usuario_oficina_actual())
    if not borrador:
        return False
    if str(borrador.get("slug", "")) != slug:
        return False
    sociedad_actual = st.session_state.get(_SOCiedad_KEY)
    if borrador.get("sociedad_id") is not None and sociedad_actual is not None:
        if int(borrador["sociedad_id"]) != int(sociedad_actual):
            return False
    st.session_state[f"{slug}_grilla_preview"] = copy.deepcopy(borrador.get("grilla_preview") or [])
    st.session_state[f"{slug}_asientos_generados"] = copy.deepcopy(borrador.get("asientos") or [])
    if borrador.get("resumen_analitico") is not None:
        st.session_state[f"{slug}_resumen_analitico"] = copy.deepcopy(borrador["resumen_analitico"])
    if borrador.get("columna_auditoria") is not None:
        st.session_state[f"{slug}_columna_auditoria"] = copy.deepcopy(borrador["columna_auditoria"])
    if borrador.get("auto_fp") is not None:
        st.session_state[f"{slug}_auto_fp"] = borrador["auto_fp"]
    st.session_state["_borrador_recuperado_slug"] = slug
    return True


def _limpiar_borrador_si_corresponde(slug: str) -> None:
    borrador = cargar_borrador_grilla_persistido(usuario=_usuario_oficina_actual())
    if borrador and str(borrador.get("slug", "")) == slug:
        limpiar_borrador_grilla_persistido(usuario=_usuario_oficina_actual())


def _slug_impuesto(nombre: str) -> str:
    try:
        return str(obtener_ficha_impuesto(nombre).get("slug", "iva"))
    except ValueError:
        try:
            return str(obtener_ficha_banco(nombre).get("slug", "banco"))
        except ValueError:
            return nombre.lower().replace(" ", "_")


def _slug_banco(nombre: str) -> str:
    try:
        return str(obtener_ficha_banco(nombre).get("slug", "banco"))
    except ValueError:
        return nombre.lower().replace(" ", "_")


def _impuestos_devengamientos() -> tuple[str, ...]:
    return tuple(TAX_REGISTRY.keys())


def _inicializar_estado_coordenadas_debe_haber(slug: str | None = None) -> None:
    """Inicializador seguro: evita KeyError en idx_debe / idx_haber (Loop Review)."""
    if "idx_debe" not in st.session_state:
        st.session_state["idx_debe"] = None
    if "idx_haber" not in st.session_state:
        st.session_state["idx_haber"] = None
    if slug:
        if f"{slug}_idx_debe" not in st.session_state:
            st.session_state[f"{slug}_idx_debe"] = None
        if f"{slug}_idx_haber" not in st.session_state:
            st.session_state[f"{slug}_idx_haber"] = None


def _sincronizar_coordenadas_session(
    slug: str,
    idx_debe: int | None,
    idx_haber: int | None,
) -> None:
    _inicializar_estado_coordenadas_debe_haber(slug)
    st.session_state["idx_debe"] = idx_debe
    st.session_state["idx_haber"] = idx_haber
    st.session_state[f"{slug}_idx_debe"] = idx_debe
    st.session_state[f"{slug}_idx_haber"] = idx_haber


def _limpiar_coordenadas_session(slug: str) -> None:
    _sincronizar_coordenadas_session(slug, None, None)


def _obtener_coordenadas_session(slug: str) -> tuple[int | None, int | None]:
    _inicializar_estado_coordenadas_debe_haber(slug)
    idx_debe = st.session_state.get(f"{slug}_idx_debe")
    if idx_debe is None:
        idx_debe = st.session_state.get("idx_debe")
    idx_haber = st.session_state.get(f"{slug}_idx_haber")
    if idx_haber is None:
        idx_haber = st.session_state.get("idx_haber")
    return idx_debe, idx_haber


def _etiqueta_mes_periodo(periodo_mensual: str) -> str:
    meses = (
        "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
    )
    parsed = _parsear_periodo_texto(str(periodo_mensual or "").replace("/", "-"))
    if parsed and 1 <= parsed[0] <= 12:
        return meses[parsed[0]]
    return "el mes seleccionado"


def _salvaguarda_render_grilla_coordenadas(
    slug: str,
    periodo_mensual: str = "",
) -> bool:
    """
    Control condicional antes de renderizar grilla.
    Si ya hay filas generadas, siempre permite editar (Debe/Haber / ⇄).
    Solo bloquea cuando aún no hay grilla y faltan coordenadas del Excel.
    """
    _inicializar_estado_coordenadas_debe_haber(slug)
    idx_debe, idx_haber = _obtener_coordenadas_session(slug)
    if idx_debe is not None and idx_haber is not None:
        return True
    rows = st.session_state.get(f"{slug}_grilla_preview") or []
    if rows:
        return True
    mes_label = _etiqueta_mes_periodo(periodo_mensual)
    st.warning(
        f"⚠️ No se pudo localizar de forma automática la columna de Debe/Haber "
        f"para el mes de {mes_label} en la planilla. "
        "Por favor, verifique el nombre de la columna en el Excel."
    )
    return False


def _flush_estado_modulo_por_slug(slug: str) -> None:
    """Flush módulo de un impuesto (sin tocar biblioteca archivada)."""
    _limpiar_coordenadas_session(slug)
    st.session_state.pop(f"{slug}_grilla_preview", None)
    st.session_state.pop(f"{slug}_asientos_generados", None)
    st.session_state.pop(f"{slug}_planilla_ref_fecha", None)
    st.session_state.pop(f"{slug}_resumen_analitico", None)
    st.session_state.pop(f"{slug}_mensaje_biblioteca_ok", None)
    st.session_state.pop(f"{slug}_limpiar_formulario_pendiente", None)
    st.session_state.pop(f"{slug}_saldos_rc_active", None)
    st.session_state[f"{slug}_skip_default_planilla"] = False
    ficha = None
    for imp in _impuestos_devengamientos():
        if _slug_impuesto(imp) == slug:
            ficha = obtener_ficha_impuesto(imp)
            break
    if ficha:
        for inp in ficha.get("inputs_contingencia") or []:
            clave = inp.get("clave")
            if clave:
                st.session_state[clave] = [0.0]
    st.session_state["reset_counter"] = int(st.session_state.get("reset_counter", 0)) + 1


def _flush_estado_iva_al_cambiar_sociedad() -> None:
    _flush_estado_modulo_por_slug("iva")


def _flush_estado_iibb_al_cambiar_sociedad() -> None:
    _flush_estado_modulo_por_slug("iibb")


def _flush_estado_sueldos_al_cambiar_sociedad() -> None:
    _flush_estado_modulo_por_slug("sueldos")


def _flush_estado_cm_al_cambiar_sociedad() -> None:
    _flush_estado_modulo_por_slug("cm")


def _flush_estado_tish_al_cambiar_sociedad() -> None:
    _flush_estado_modulo_por_slug("tish")


def _flush_estado_iva_modulo() -> None:
    _flush_estado_modulo_por_slug("iva")


def _flush_estado_iibb_modulo() -> None:
    _flush_estado_modulo_por_slug("iibb")


def _flush_estado_sueldos_modulo() -> None:
    _flush_estado_modulo_por_slug("sueldos")


def _flush_estado_tish_modulo() -> None:
    _flush_estado_modulo_por_slug("tish")


def _detectar_cambio_impuesto_y_flush() -> None:
    """Al cambiar impuesto o sociedad, vacía módulos activos (biblioteca intacta)."""
    actual = str(st.session_state.get(_IMPUESTO_KEY, "IVA")).strip() or "IVA"
    previo = st.session_state.get("_impuesto_previo_devengamientos")
    if previo is not None and previo != actual:
        for imp in _impuestos_devengamientos():
            if imp != actual:
                _flush_estado_modulo_por_slug(_slug_impuesto(imp))
    st.session_state["_impuesto_previo_devengamientos"] = actual


def _detectar_cambio_banco_y_flush() -> None:
    """Al cambiar banco activo, vacía grilla del banco anterior (biblioteca bancaria intacta)."""
    actual = str(st.session_state.get(_BANCO_KEY, "") or "").strip()
    previo = st.session_state.get("_banco_previo_conciliacion")
    if previo is not None and previo != actual and previo:
        try:
            slug_prev = str(obtener_ficha_banco(previo).get("slug", _slug_banco(previo)))
            _flush_estado_modulo_por_slug(slug_prev)
        except ValueError:
            pass
    if actual:
        st.session_state["_banco_previo_conciliacion"] = actual


def _clave_periodo_mensual_banco(slug: str, sociedad_id: int) -> str:
    return f"banco_periodo_v2_{slug}_{sociedad_id}"


def _clave_fecha_tango_banco(slug: str, rc: int, periodo_mensual: str | None = None) -> str:
    """Key del date_input bancario. Incluye el período para remount seguro al cambiar mes."""
    base = f"banco_fecha_tango_v2_{slug}_{rc}"
    p = _periodo_clave_fecha(periodo_mensual)
    return f"{base}_{p}" if p else base


def _clave_fecha_tango_slug(slug: str, rc: int, periodo_mensual: str | None = None) -> str:
    """Key del date_input de devengamientos. Incluye el período (anti removeChild)."""
    base = f"{slug}_fecha_tango_{rc}"
    p = _periodo_clave_fecha(periodo_mensual)
    return f"{base}_{p}" if p else base


def _periodo_clave_fecha(periodo_mensual: str | None) -> str | None:
    if not periodo_mensual:
        return None
    parsed = _parsear_periodo_texto(str(periodo_mensual).replace("/", "-"))
    if not parsed:
        return None
    return f"{int(parsed[0]):02d}-{int(parsed[1])}"


def _safe_wkey_uid(uid: str) -> str:
    return re.sub(r"[^A-Za-z0-9_\-]", "_", str(uid or "x"))[:96]


def _wkey_banco_cuenta_row(slug: str, uid: str) -> str:
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    return f"banco_cuenta_v3_{slug}_{gen}_{_safe_wkey_uid(uid)}"


def _wkey_banco_tipo_row(slug: str, uid: str) -> str:
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    return f"banco_tipo_v3_{slug}_{gen}_{_safe_wkey_uid(uid)}"


def _wkey_banco_del_row(slug: str, uid: str) -> str:
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    return f"banco_del_v3_{slug}_{gen}_{_safe_wkey_uid(uid)}"


def _wkey_banco_swap_row(slug: str, uid: str) -> str:
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    return f"banco_swap_v3_{slug}_{gen}_{_safe_wkey_uid(uid)}"


def _wkey_banco_monto_row(slug: str, uid: str, tipo: str) -> str:
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    return f"banco_monto_v3_{slug}_{gen}_{_safe_wkey_uid(uid)}_{tipo}"


def _bump_ui_grilla_banco(slug: str) -> None:
    """Nueva generación de widgets: evita removeChild al regenerar el asiento."""
    st.session_state[f"banco_ui_gen_{slug}"] = int(
        st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0
    ) + 1
    st.session_state[f"banco_grilla_page_{slug}"] = 0
    _limpiar_widgets_grilla_banco(slug)


def _wkey_monto_editable(slug: str, rc: int, idx: int, tipo: str) -> str:
    return f"monto_editable_{slug}_{rc}_{idx}_{tipo}"


def _btn_quitar_bancos_v2(slug: str, sociedad_id: int | None) -> str:
    return f"btn_quitar_bancos_v2_{slug}_{sociedad_id}"


def _fecha_asiento_seleccionada_banco(slug: str, periodo_mensual: str | None = None) -> date | None:
    rc = _iva_reset_counter()
    if periodo_mensual:
        val = st.session_state.get(_clave_fecha_tango_banco(slug, rc, periodo_mensual))
        if isinstance(val, date):
            return val
    val = st.session_state.get(_clave_fecha_tango_banco(slug, rc))
    return val if isinstance(val, date) else None


def _purge_banco_widget_keys_v2() -> None:
    """Elimina keys de widgets bancarios v2/v3 al salir de Conciliación Bancaria."""
    prefixes = (
        "banco_cuenta_row_",
        "banco_cuenta_v3_",
        "banco_tipo_row_",
        "banco_tipo_v3_",
        "banco_monto_row_",
        "banco_monto_v3_",
        "banco_del_row_",
        "banco_del_v3_",
        "banco_swap_row_",
        "banco_swap_v3_",
        "btn_quitar_bancos_v2_",
        "banco_periodo_v2_",
        "banco_fecha_tango_v2_",
        "dl_tango_banco_v2_",
        "btn_guardar_biblioteca_banco_v2_",
        "btn_vaciar_biblioteca_banco_v2_",
        "vent_bco_",
        "monto_editable_",
        "banco_grilla_page_",
        "banco_grilla_editor_v1_",
    )
    for key in list(st.session_state.keys()):
        ks = str(key)
        if any(ks.startswith(p) for p in prefixes):
            st.session_state.pop(key, None)


def _detectar_cambio_ventana_y_flush() -> None:
    """Al cambiar ventana principal, reinicia coordenadas globales sin tocar bibliotecas."""
    actual = str(st.session_state.get(_VENTANA_KEY, _VENTANAS_PRINCIPALES[0]))
    previo = st.session_state.get("_ventana_principal_previa")
    if previo is not None and previo != actual:
        _inicializar_estado_coordenadas_debe_haber()
        if previo == "Conciliación Bancaria":
            _purge_banco_widget_keys_v2()
    st.session_state["_ventana_principal_previa"] = actual


def _detectar_cambio_modulo_trabajo_y_flush() -> None:
    """Separa estado entre Devengamientos e Conciliación Bancaria."""
    actual = str(st.session_state.get(_MODULO_TRABAJO_KEY, _MODULOS_TRABAJO[0]))
    previo = st.session_state.get("_modulo_trabajo_previo")
    if previo is not None and previo != actual:
        _inicializar_estado_coordenadas_debe_haber()
    st.session_state["_modulo_trabajo_previo"] = actual


def _limpiar_resultados_cambio_sociedad() -> None:
    """State flush al cambiar de sociedad: evita contaminación cruzada de datos."""
    st.session_state.resultado = None
    st.session_state.mov_banco = None
    _flush_estado_iva_al_cambiar_sociedad()
    _flush_estado_iibb_al_cambiar_sociedad()
    _flush_estado_cm_al_cambiar_sociedad()
    _flush_estado_sueldos_al_cambiar_sociedad()
    _flush_estado_tish_al_cambiar_sociedad()


def _limpiar_estado_devengamientos_desincronizado() -> None:
    """Resetea estado de Devengamientos ante falla de sincronización."""
    st.session_state.cuit_activo = None
    st.session_state.nombre_activo = None
    st.session_state.plan_cuentas_df = None
    st.session_state.plan_cuentas_cliente_id = None
    st.session_state.plan_cuentas_path_resuelto = None
    st.session_state.plan_cuentas_es_default = False
    st.session_state.pop("iva_grilla_preview", None)
    st.session_state.pop("iva_asientos_generados", None)
    st.session_state.pop("iva_planilla_ref_fecha", None)
    st.session_state.devengamiento_resultado = None
    st.session_state.devengamiento_datos = None


def _cliente_sociedad_activa(clientes: list[dict]) -> dict | None:
    """Lee sociedad_activa sin escribirla; otras pestañas no pisan el selector de Devengamientos."""
    if not clientes:
        return None
    ids = [c["id"] for c in clientes]
    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    if sociedad_id not in ids:
        # Fallback solo lectura: no mutar st.session_state.sociedad_activa
        sociedad_id = ids[0]
        # #region agent log
        _dbg_log("I", "_cliente_sociedad_activa", "fallback_readonly", {
            "sociedad_activa": st.session_state.get(_SOCiedad_KEY),
            "fallback_id": sociedad_id,
        })
        # #endregion
    else:
        st.session_state.cliente_id_seleccionado = sociedad_id
    return next(c for c in clientes if c["id"] == sociedad_id)


def _guardar_upload(archivo) -> Path:
    """Persiste un archivo subido en un temporal para procesarlo.

    En Cloud, si hay DATA_ENCRYPTION_KEY, también guarda copia cifrada
    namespaced por usuario en data/secure/<user>/uploads/.
    """
    sufijo = Path(archivo.name).suffix
    contenido = bytes(archivo.getbuffer())
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=sufijo)
    tmp.write(contenido)
    tmp.close()
    if _es_entorno_cloud():
        try:
            from seguridad_datos import guardar_upload_cifrado, tiene_clave_cifrado

            if tiene_clave_cifrado():
                guardar_upload_cifrado(
                    _usuario_oficina_actual(),
                    "uploads",
                    archivo.name,
                    contenido,
                )
        except Exception:
            pass
    return Path(tmp.name)


def _mostrar_cliente_activo_sidebar(cliente: dict | None) -> None:
    """Muestra el cliente congelado en la barra lateral."""
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Cliente activo")
    if cliente:
        st.sidebar.info(
            f"**{cliente['nombre']}**\n\n"
            f"CUIT: `{cliente['cuit']}`\n\n"
            f"Tipo: {cliente['tipo_persona']}"
        )
    else:
        st.sidebar.warning("Sin cliente seleccionado")


def _seccion_clientes() -> None:
    st.subheader("Gestión de Clientes")

    if st.button("Actualizar CUITs desde auxiliares", type="secondary"):
        with st.spinner("Actualizando CUITs..."):
            stats = db.actualizar_cuits_desde_auxiliares()
            st.success(
                f"Actualizados: {stats['actualizados']} | "
                f"Insertados: {stats['insertados']} | "
                f"Temporales: {stats['temporales']}"
            )

    tab_lista, tab_alta = st.tabs(["Clientes registrados", "Alta / Edición"])
    clientes = db.listar_clientes()

    with tab_alta:
        opciones_edicion = {"— Nuevo cliente —": None}
        opciones_edicion.update({f"{c['nombre']} ({c['cuit']})": c["id"] for c in clientes})
        seleccion = st.selectbox("Editar cliente existente", list(opciones_edicion.keys()))
        cliente_edit = db.obtener_cliente(opciones_edicion[seleccion]) if opciones_edicion[seleccion] else None

        with st.form("form_cliente", clear_on_submit=not cliente_edit):
            nombre = st.text_input("Nombre / Razón Social", value=cliente_edit["nombre"] if cliente_edit else "")
            cuit = st.text_input("CUIT", value=cliente_edit["cuit"] if cliente_edit else "")
            tipo_persona = st.selectbox(
                "Tipo de Persona",
                db.TIPOS_PERSONA,
                index=db.TIPOS_PERSONA.index(cliente_edit["tipo_persona"]) if cliente_edit else 0,
            )
            mes_cierre = st.number_input(
                "Mes cierre de balance",
                min_value=1,
                max_value=12,
                value=int(cliente_edit.get("mes_cierre_balance") or 12) if cliente_edit else 12,
            )
            col1, col2 = st.columns(2)
            guardar = col1.form_submit_button("Guardar", type="primary")
            eliminar = col2.form_submit_button("Eliminar") if cliente_edit else False

            if guardar:
                if not nombre or not cuit:
                    st.error("Nombre y CUIT son obligatorios.")
                else:
                    try:
                        if cliente_edit:
                            db.actualizar_cliente(
                                cliente_edit["id"],
                                nombre,
                                cuit,
                                tipo_persona,
                                cliente_edit.get("plan_cuentas_path"),
                                int(mes_cierre),
                            )
                            st.success(f"Cliente '{nombre}' actualizado.")
                        else:
                            db.crear_cliente(nombre, cuit, tipo_persona, None, int(mes_cierre))
                            st.success(f"Cliente '{nombre}' registrado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Error al guardar: {exc}")

            if eliminar and cliente_edit:
                db.eliminar_cliente(cliente_edit["id"])
                if st.session_state.cliente_id_seleccionado == cliente_edit["id"]:
                    st.session_state.cliente_id_seleccionado = None
                    st.session_state.sociedad_activa = None
                st.success("Cliente eliminado.")
                st.rerun()

        st.markdown("---")
        st.markdown("**Plan de Cuentas del Cliente**")
        pc_actual = cliente_edit.get("plan_cuentas_path") if cliente_edit else None
        if pc_actual and Path(pc_actual).exists():
            st.success(f"Plan vinculado: {Path(pc_actual).name}")
        else:
            st.info("Sin plan asociado. Se usará el plan por defecto del proyecto.")

        archivo_pc = st.file_uploader(
            "Subir Plan de Cuentas (.xlsx)",
            type=["xlsx"],
            key=f"uploader_plan_{opciones_edicion.get(seleccion, 'nuevo')}",
        )
        if archivo_pc and cliente_edit:
            try:
                xlsx_path = _guardar_plan_cliente_en_disco(cliente_edit["cuit"], archivo_pc.getvalue())
                db.actualizar_cliente(
                    cliente_edit["id"],
                    cliente_edit["nombre"],
                    cliente_edit["cuit"],
                    cliente_edit["tipo_persona"],
                    str(xlsx_path),
                    cliente_edit.get("mes_cierre_balance", 12),
                )
                _invalidar_cache_plan(cliente_edit["id"])
                _sincronizar_plan_cuentas_session(cliente_edit["id"], forzar=True)
                if st.session_state.get(_SOCiedad_KEY) == cliente_edit["id"]:
                    actualizar_sociedad_activa()
                st.success(f"Plan Excel guardado: {xlsx_path.name}")
                st.rerun()
            except Exception as exc:
                st.error(f"Error al guardar el plan de cuentas: {exc}")

    with tab_lista:
        if not clientes:
            st.info("No hay clientes registrados.")
        else:
            df = pd.DataFrame(clientes)[["id", "nombre", "cuit", "tipo_persona", "creado_en"]]
            df["cuit_temporal"] = df["cuit"].astype(str).str.startswith("99").map({True: "⚠️ Temporal", False: ""})
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.caption(f"Total: {len(clientes)} clientes registrados")


def _seccion_conciliacion() -> None:
    st.subheader("Conciliación Bancaria")

    clientes = db.listar_clientes()
    if not clientes:
        st.warning("Debe registrar al menos un cliente antes de conciliar.")
        return

    cliente = _cliente_sociedad_activa(clientes)
    if not cliente:
        return

    st.caption(f"Sociedad activa: **{cliente['nombre']}** (elegila en Devengamientos si necesitás cambiarla).")

    st.session_state.cliente_activo = cliente
    es_juridica = cliente["tipo_persona"] == "Persona Jurídica"

    col_info1, col_info2, col_info3 = st.columns(3)
    col_info1.metric("CUIT", cliente["cuit"])
    col_info2.metric("Tipo", cliente["tipo_persona"])
    col_info3.metric("ID Cliente", cliente["id"])

    modo_auto = st.session_state.modo_ingesta_automatica
    ruta_raiz = Path(st.session_state.ruta_raiz_clientes)

    extractos = None
    contables = None
    compras_file = None
    plan_file = None

    if modo_auto:
        st.markdown("#### Ingesta automática de archivos")
        st.info(f"Modo automático activo: se escaneará `{ruta_raiz / cliente['cuit']}` al procesar.")
    else:
        st.markdown("#### Carga de archivos")
        extractos = st.file_uploader(
            f"Extractos bancarios PDF (hasta {MAX_PDFS_ANUALES} archivos — año completo)",
            type=["pdf"],
            accept_multiple_files=True,
            help="Seleccione uno o más extractos mensuales. El sistema los consolidará cronológicamente.",
        )
        col_a, col_b = st.columns(2)
        with col_a:
            contables = st.file_uploader(
                "Movimientos contables (Excel/CSV) — opcional",
                type=["xlsx", "xls", "csv"],
            )
        with col_b:
            compras_file = st.file_uploader(
                "Compras Tango (Excel) — opcional",
                type=["xlsx", "xls"],
                help=f"Si no sube archivo, se usa `{COMPRAS_TANGO_PATH.name}` del proyecto.",
            )
        plan_file = st.file_uploader(
            "Plan de cuentas (Excel) — opcional",
            type=["xlsx", "xls"],
            help="Si no sube archivo, se usa el plan del cliente o el plan por defecto.",
        )

    tolerancia = st.slider("Tolerancia de clearing bancario (días)", 1, 7, 3)

    if st.button("Procesar conciliación anual", type="primary"):
        mensaje_spinner = (
            "Escaneando carpeta y procesando archivos..."
            if modo_auto
            else "Extrayendo y clasificando movimientos (OCR + detección de banco)..."
        )
        with st.spinner(mensaje_spinner):
            try:
                ruta_compras = None
                ruta_cuentas = None

                if modo_auto:
                    rutas_pdf, ruta_compras, ruta_cuentas = escanear_carpeta_cliente(
                        cliente["cuit"], ruta_raiz
                    )
                    if not rutas_pdf:
                        st.error(
                            f"No se encontraron extractos PDF en la carpeta del cliente: "
                            f"{ruta_raiz / cliente['cuit']}"
                        )
                        return
                else:
                    if not extractos:
                        st.error("Debe subir al menos un extracto PDF.")
                        return
                    if len(extractos) > MAX_PDFS_ANUALES:
                        st.warning(f"Se procesarán solo los primeros {MAX_PDFS_ANUALES} archivos.")
                        extractos = extractos[:MAX_PDFS_ANUALES]
                    rutas_pdf = [_guardar_upload(e) for e in extractos]
                    if compras_file:
                        ruta_compras = _guardar_upload(compras_file)
                    if plan_file:
                        ruta_cuentas = _guardar_upload(plan_file)

                mov_banco, bancos, saldos_mes = extraer_movimientos_anuales(rutas_pdf)
                if not mov_banco:
                    st.warning("No se detectaron movimientos en los extractos.")
                    return

                if cliente.get("tipo_persona") == "Persona Jurídica" and not cliente.get("mes_cierre_balance"):
                    meses_ordenados = sorted(saldos_mes.keys())
                    if meses_ordenados:
                        primer_mes = meses_ordenados[0][1]
                        mes_cierre = primer_mes - 1
                        if mes_cierre == 0:
                            mes_cierre = 12
                        db.actualizar_cliente(
                            cliente["id"],
                            cliente["nombre"],
                            cliente["cuit"],
                            cliente["tipo_persona"],
                            cliente.get("plan_cuentas_path"),
                            mes_cierre_balance=mes_cierre,
                        )
                        cliente["mes_cierre_balance"] = mes_cierre
                        st.info(f"Ciclo fiscal deducido: cierre en mes {mes_cierre}")

                if ruta_cuentas:
                    plan_path = str(ruta_cuentas)
                else:
                    plan_path = cliente.get("plan_cuentas_path") or str(PLAN_CUENTAS_DEFAULT)
                plan_cuentas = cargar_plan_cuentas(plan_path)
                compras = cargar_compras_tango(ruta_compras)

                mov_contables = (
                    []
                    if modo_auto
                    else (cargar_movimientos_contables(contables) if contables else [])
                )
                resultado = conciliar_movimientos(
                    mov_banco,
                    mov_contables,
                    tolerancia_dias=tolerancia,
                    compras=compras,
                )
                resultado = aplicar_saldos_al_resultado(resultado, saldos_mes)

                st.session_state.resultado = resultado
                st.session_state.plan_cuentas = plan_cuentas
                st.session_state.mov_banco = mov_banco
                st.session_state.bancos_detectados = bancos
                st.session_state.cliente_activo = cliente

                bancos_nombres = [PERFILES_BANCO.get(b, {}).get("nombre_display", b) for b in bancos]
                st.success(
                    f"✅ {len(mov_banco)} movimientos | "
                    f"{len(resultado.conciliados)} conciliados | "
                    f"{len(resultado.solo_banco)} solo banco | "
                    f"{len(resultado.anomalias)} anomalías | "
                    f"Bancos: {', '.join(bancos_nombres)}"
                )
            except Exception as exc:
                st.error(f"Error en el procesamiento: {exc}")

    if (
        st.session_state.resultado
        and st.session_state.get("cliente_activo", {}).get("id") == cliente["id"]
    ):
        resultado = st.session_state.resultado
        plan_cuentas = st.session_state.plan_cuentas
        mov_banco = st.session_state.mov_banco

        tab_ext, tab_conc, tab_clas, tab_exp = st.tabs(
            ["Extracto consolidado", "Conciliación", "Clasificación contable", "Exportar"]
        )

        with tab_ext:
            st.dataframe(movimientos_a_dataframe(mov_banco), use_container_width=True, hide_index=True)

        with tab_conc:
            df_res = resultado_a_dataframe(resultado)
            st.dataframe(df_res, use_container_width=True, hide_index=True)
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Conciliados", len(resultado.conciliados))
            m2.metric("Solo banco", len(resultado.solo_banco))
            m3.metric("Solo contabilidad", len(resultado.solo_contabilidad))
            m4.metric("Anomalías", len(resultado.anomalias))
            if resultado.saldo_extracto:
                m5.metric("Saldo final", f"${resultado.saldo_extracto:,.2f}")

        with tab_clas:
            if resultado.resumen_por_categoria:
                df_cat = pd.DataFrame(
                    [{"Categoría planilla": k, "Importe acumulado": v} for k, v in resultado.resumen_por_categoria.items()]
                )
                st.dataframe(df_cat, use_container_width=True, hide_index=True)
            if resultado.anomalias:
                st.markdown("#### ⚠️ Anomalías detectadas")
                st.dataframe(pd.DataFrame(resultado.anomalias), use_container_width=True, hide_index=True)
            if resultado.resumen_anual_por_mes:
                st.markdown("#### Resumen anual por mes")
                filas_mes = []
                for (anio, mes), cats in sorted(resultado.resumen_anual_por_mes.items()):
                    for cat, imp in cats.items():
                        filas_mes.append({"Año": anio, "Mes": mes, "Categoría": cat, "Importe": imp})
                st.dataframe(pd.DataFrame(filas_mes), use_container_width=True, hide_index=True)

        with tab_exp:
            try:
                excel_bytes = generar_planilla_conciliacion(
                    resultado,
                    nombre_cliente=cliente["nombre"],
                    ruta_plantilla=PLANTILLA_CONCILIACION,
                )
                st.download_button(
                    label="Descargar Planilla de Conciliación (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Conciliacion_{cliente['cuit']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as exc:
                st.error(f"Error al generar planilla: {exc}")

            if es_juridica:
                txt_content = generar_txt_tango(resultado, plan_cuentas)
                st.download_button(
                    label="Descargar Asientos Tango (.txt)",
                    data=txt_content.encode("utf-8"),
                    file_name=f"Asientos_Tango_{cliente['cuit']}.txt",
                    mime="text/plain",
                )
            else:
                st.info("Exportación TXT deshabilitada para Persona Física.")



_MARC_DEBITO_FISCAL = ("Debito Fiscal Actividades", "Débito Fiscal Actividades")
_MARC_CREDITO_FISCAL = ("Crédito Fiscal Actividades", "Credito Fiscal Actividades")
_MARC_LIQUIDACION = ("Liquidación", "Liquidacion")
_PAT_ALIC_IVA = re.compile(r"10[,.]?5|10[,.]?50|(?<!\d)21(?!\d)|(?<!\d)27(?!\d)")


def _directorio_planes_canonico() -> Path:
    """
    Carpeta canónica del plan de cuentas.
    En oficina (Windows): prioriza T:\\…\\planes_cuentas.
    En Cloud/Linux: solo data/planes_cuentas del repo (sin T:).
    """
    import os

    candidatos: list[Path] = []
    if os.name == "nt" and not _es_entorno_cloud():
        candidatos.append(PLANES_RED_DIR)
    candidatos.append(DATA_PLANES_DIR)
    for carpeta in candidatos:
        try:
            carpeta.mkdir(parents=True, exist_ok=True)
            probe = carpeta / ".planes_probe"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return carpeta
        except OSError:
            continue
    DATA_PLANES_DIR.mkdir(parents=True, exist_ok=True)
    return DATA_PLANES_DIR


def _mtime_archivo(ruta: Path | None) -> float | None:
    if ruta is None:
        return None
    try:
        p = Path(ruta)
        if p.is_file():
            return float(p.stat().st_mtime)
    except OSError:
        return None
    return None


def _ruta_plan_xlsx(cuit: str) -> Path:
    """Ruta de escritura/lectura preferida: red compartida si está disponible."""
    return _directorio_planes_canonico() / f"plan_{cuit}.xlsx"


def _ruta_plan_xlsx_red(cuit: str) -> Path:
    return PLANES_RED_DIR / f"plan_{cuit}.xlsx"


def _ruta_plan_xlsx_local(cuit: str) -> Path:
    return DATA_PLANES_DIR / f"plan_{cuit}.xlsx"


def _indice_sociedades_pj(clientes_pj: list[dict]) -> dict[int, dict]:
    """Mapa id → {id, nombre, cuit, plan_path} para sincronización atómica."""
    indice: dict[int, dict] = {}
    for c in clientes_pj:
        cuit = str(c.get("cuit", "")).strip()
        indice[c["id"]] = {
            "id": c["id"],
            "nombre": c["nombre"],
            "cuit": cuit,
            "plan_path": _ruta_plan_xlsx(cuit),
        }
    return indice


def _ruta_plan_csv(cuit: str) -> Path:
    return _directorio_planes_canonico() / f"plan_{cuit}.csv"


def _rutas_plan_candidatas(cliente: dict) -> list[Path]:
    cuit = str(cliente.get("cuit", "")).strip()
    # Primero la red compartida, después local — así todos ven el último upload.
    candidatas: list[Path] = [
        _ruta_plan_xlsx_red(cuit),
        _ruta_plan_xlsx_local(cuit),
        LEGACY_PLANES_DIR / f"{cuit}_plan.xlsx",
    ]
    # Cloud: planes cifrados por usuario (y fallback sin usuario)
    try:
        from seguridad_datos import ruta_plan_cifrado

        candidatas.insert(0, ruta_plan_cifrado(_usuario_oficina_actual(), cuit))
        candidatas.insert(1, ruta_plan_cifrado(None, cuit))
    except Exception:
        pass
    carpeta_cliente = RUTA_RAIZ_CLIENTES / cuit
    if carpeta_cliente.is_dir():
        for archivo in sorted(carpeta_cliente.iterdir()):
            if not archivo.is_file() or archivo.name.startswith("~$"):
                continue
            nombre = archivo.name.lower()
            if archivo.suffix.lower() in (".xlsx", ".xls") and (
                "cuentas" in nombre or "plan" in nombre
            ):
                candidatas.append(archivo)
    if cliente.get("plan_cuentas_path"):
        bd_path = Path(cliente["plan_cuentas_path"])
        if bd_path.suffix.lower() in (".xlsx", ".enc") and bd_path not in candidatas:
            candidatas.append(bd_path)
    candidatas.append(_directorio_planes_canonico() / f"plan_{cuit}.csv")
    candidatas.append(DATA_PLANES_DIR / f"plan_{cuit}.csv")
    vistos: set[str] = set()
    unicas: list[Path] = []
    for p in candidatas:
        key = str(p).lower()
        if key not in vistos:
            vistos.add(key)
            unicas.append(p)
    return unicas


def _resolver_ruta_plan_para_lectura(ruta: Path) -> Path:
    """Si el plan está cifrado (.enc), materializa un temporal descifrado."""
    ruta = Path(ruta)
    if str(ruta).endswith(".enc") or ruta.suffix.lower() == ".enc":
        from seguridad_datos import materializar_descifrado

        return materializar_descifrado(ruta, suffix=".xlsx")
    return ruta


def _es_plan_generico_default(ruta: Path) -> bool:
    """True si la ruta es el plan genérico del estudio (nunca cuenta como vinculado a una SA)."""
    ruta = Path(ruta)
    nombre = ruta.name.lower()
    if nombre in ("plan_default.xlsx", "plan_default.xls") or "cuentas contables (4)" in nombre:
        return True
    defaults = [
        Path(PLAN_CUENTAS_DEFAULT),
        Path(getattr(db, "PLAN_CUENTAS_DEFAULT_REPO", "")),
        Path(getattr(db, "PLAN_CUENTAS_DEFAULT_LEGACY", "")),
    ]
    try:
        defaults.append(Path(db._plan_cuentas_default_path()))
    except Exception:
        pass
    for d in defaults:
        if not d or str(d) in ("", "."):
            continue
        try:
            if ruta.resolve() == d.resolve():
                return True
        except OSError:
            if str(ruta).lower() == str(d).lower():
                return True
    return False


def _es_plan_propio_cliente(cuit: str, ruta: Path) -> bool:
    """True solo si el archivo es el Excel específico del cliente (no el genérico del proyecto)."""
    ruta = Path(ruta)
    if not ruta.exists():
        return False
    if _es_plan_generico_default(ruta):
        return False
    # Plan cifrado del usuario = propio
    if "data" in ruta.parts and "secure" in ruta.parts and f"plan_{cuit}" in ruta.name:
        return True
    for cand in (_ruta_plan_xlsx_red(cuit), _ruta_plan_xlsx_local(cuit), _ruta_plan_xlsx(cuit)):
        try:
            if cand.exists() and ruta.resolve() == cand.resolve():
                return True
        except OSError:
            if str(ruta).lower() == str(cand).lower():
                return True
    try:
        if ruta.resolve() == (LEGACY_PLANES_DIR / f"{cuit}_plan.xlsx").resolve():
            return True
    except OSError:
        pass
    carpeta_cliente = RUTA_RAIZ_CLIENTES / cuit
    try:
        if carpeta_cliente.is_dir() and carpeta_cliente.resolve() == ruta.parent.resolve():
            # En carpeta del cliente: Excel con 'plan'/'cuentas' en el nombre (no default)
            nombre = ruta.name.lower()
            if ruta.suffix.lower() in (".xlsx", ".xls") and (
                "cuentas" in nombre or "plan" in nombre
            ):
                return True
    except OSError:
        pass
    # plan_{cuit}.xlsx en carpeta planes_cuentas (red o local)
    if ruta.name.lower() == f"plan_{cuit}.xlsx".lower() and "planes_cuentas" in str(ruta).lower():
        return True
    # Path registrado en BD (nombre libre) solo si el nombre contiene el CUIT
    nombre = ruta.name.lower()
    cuit_norm = re.sub(r"\D", "", str(cuit))
    if cuit_norm and cuit_norm in re.sub(r"\D", "", nombre):
        return True
    return False


def _promover_plan_a_canonico(cuit: str, ruta: Path) -> Path:
    """Copia el Excel encontrado al canónico compartido (red o local) si aún no existe."""
    canon = _ruta_plan_xlsx(cuit)
    if canon.exists():
        return canon
    if not _es_plan_propio_cliente(cuit, ruta):
        return ruta
    canon.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta, canon)
    df = cargar_plan_cuentas(canon)
    df[["codigo", "descripcion", "imputable"]].to_csv(_ruta_plan_csv(cuit), index=False)
    # Espejo local para cuando la red no esté
    try:
        local = _ruta_plan_xlsx_local(cuit)
        if local.resolve() != canon.resolve():
            local.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(canon, local)
            df[["codigo", "descripcion", "imputable"]].to_csv(
                DATA_PLANES_DIR / f"plan_{cuit}.csv", index=False,
            )
    except OSError:
        pass
    _dbg_log("C", "_promover_plan_a_canonico", "promoted", {"cuit": cuit, "from": str(ruta), "to": str(canon)})
    return canon


def _guardar_plan_cliente_en_disco(cuit: str, archivo_bytes: bytes) -> Path:
    """
    Persiste el plan en el servidor compartido (T:\\…\\planes_cuentas) y espejo local.
    En Cloud: cifrado Fernet namespaced por usuario (data/secure/<user>/planes/).
    """
    # Cloud: no dejar plaintext en disco del contenedor
    if _es_entorno_cloud():
        from seguridad_datos import (
            guardar_plan_cifrado,
            materializar_descifrado,
            tiene_clave_cifrado,
        )

        if not tiene_clave_cifrado():
            raise RuntimeError(
                "En Cloud hace falta DATA_ENCRYPTION_KEY en Secrets para guardar planes."
            )
        enc = guardar_plan_cifrado(_usuario_oficina_actual(), cuit, archivo_bytes)
        tmp = materializar_descifrado(enc, suffix=".xlsx")
        df = cargar_plan_cuentas(tmp)
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        _dbg_log(
            "C",
            "_guardar_plan_cliente_en_disco",
            "saved_xlsx_encrypted",
            {"cuit": cuit, "path": str(enc), "rows": len(df)},
        )
        return enc

    ruta_xlsx = _ruta_plan_xlsx(cuit)
    ruta_xlsx.parent.mkdir(parents=True, exist_ok=True)
    ruta_xlsx.write_bytes(archivo_bytes)
    df = cargar_plan_cuentas(ruta_xlsx)
    df[["codigo", "descripcion", "imputable"]].to_csv(_ruta_plan_csv(cuit), index=False)

    # Espejo en la otra ubicación (red↔local) para resiliencia offline
    try:
        espejo = (
            _ruta_plan_xlsx_local(cuit)
            if ruta_xlsx.parent.resolve() == PLANES_RED_DIR.resolve()
            else _ruta_plan_xlsx_red(cuit)
        )
        if espejo.resolve() != ruta_xlsx.resolve():
            espejo.parent.mkdir(parents=True, exist_ok=True)
            espejo.write_bytes(archivo_bytes)
            csv_espejo = espejo.with_suffix(".csv")
            df[["codigo", "descripcion", "imputable"]].to_csv(csv_espejo, index=False)
    except OSError:
        pass

    _dbg_log(
        "C",
        "_guardar_plan_cliente_en_disco",
        "saved_xlsx_shared",
        {
            "cuit": cuit,
            "path": str(ruta_xlsx),
            "shared": str(PLANES_RED_DIR),
            "rows": len(df),
        },
    )
    return ruta_xlsx


def _invalidar_cache_plan(cliente_id: int) -> None:
    st.session_state.pop(f"plan_cuentas_df_{cliente_id}", None)
    st.session_state.pop(f"plan_cuentas_meta_{cliente_id}", None)
    if st.session_state.get("plan_cuentas_cliente_id") == cliente_id:
        st.session_state.pop("plan_cuentas_df", None)
        st.session_state.pop("plan_cuentas_cliente_id", None)
        st.session_state.pop("plan_cuentas_path_resuelto", None)
        st.session_state.pop("plan_cuentas_es_default", None)
        st.session_state.pop("plan_cuentas_mtime", None)


def _cargar_plan_cuentas_cliente(
    cliente_id: int,
    usar_cache: bool = True,
) -> tuple[pd.DataFrame, Path, bool]:
    """
    Carga plan del cliente desde disco (BD → CSV canónico → legacy → default).
    Retorna (dataframe, ruta_resuelta, es_plan_default).
    Invalida cache si el archivo en red/local cambió de mtime (otro usuario lo actualizó).
    """
    cliente = db.obtener_cliente(cliente_id)
    if not cliente:
        raise ValueError(f"Cliente id={cliente_id} no encontrado.")
    cliente = _limpiar_plan_bd_si_archivo_ausente(cliente)

    cache_key = f"plan_cuentas_df_{cliente_id}"
    meta_key = f"plan_cuentas_meta_{cliente_id}"

    ruta_resuelta: Path | None = None
    for candidata in _rutas_plan_candidatas(cliente):
        if candidata.exists() and not _es_plan_generico_default(candidata):
            # Preferir primera candidata no-genérica; propio se prioriza abajo
            if ruta_resuelta is None:
                ruta_resuelta = candidata
            if _es_plan_propio_cliente(str(cliente.get("cuit", "")).strip(), candidata):
                ruta_resuelta = candidata
                break

    cuit = str(cliente.get("cuit", "")).strip()
    bd_raw = (cliente.get("plan_cuentas_path") or "").strip()
    bd_path = Path(bd_raw) if bd_raw else None
    bd_vinculado = bool(
        bd_path and bd_path.exists() and not _es_plan_generico_default(bd_path)
    )

    # Verde/vinculado: plan_{cuit} o path BD real (no plan_default).
    if ruta_resuelta is not None and _es_plan_propio_cliente(cuit, ruta_resuelta):
        using_default = False
        if not str(ruta_resuelta).endswith(".enc"):
            ruta_resuelta = _promover_plan_a_canonico(cuit, ruta_resuelta)
    elif bd_vinculado:
        ruta_resuelta = bd_path  # type: ignore[assignment]
        using_default = False
    else:
        propio_hallado: Path | None = None
        for candidata in _rutas_plan_candidatas(cliente):
            if candidata.exists() and _es_plan_propio_cliente(cuit, candidata):
                propio_hallado = candidata
                break
        if propio_hallado is not None:
            ruta_resuelta = propio_hallado
            using_default = False
            if not str(ruta_resuelta).endswith(".enc"):
                ruta_resuelta = _promover_plan_a_canonico(cuit, ruta_resuelta)
        else:
            ruta_resuelta = Path(db._plan_cuentas_default_path())
            if not ruta_resuelta.is_file():
                ruta_resuelta = Path(PLAN_CUENTAS_DEFAULT)
            using_default = True

    mtime = _mtime_archivo(ruta_resuelta)

    if usar_cache and cache_key in st.session_state:
        meta = st.session_state.get(meta_key) or {}
        if (
            meta.get("cliente_id") == cliente_id
            and meta.get("path") == str(ruta_resuelta)
            and bool(meta.get("using_default")) == using_default
            and meta.get("mtime") == mtime
        ):
            return (
                st.session_state[cache_key],
                Path(meta["path"]),
                bool(meta.get("using_default", using_default)),
            )
        # #region agent log
        _dbg_log("H", "_cargar_plan_cuentas_cliente", "per_client_cache_stale", {
            "cliente_id": cliente_id,
            "meta_path": meta.get("path"),
            "resolved_path": str(ruta_resuelta),
            "meta_mtime": meta.get("mtime"),
            "resolved_mtime": mtime,
            "meta_default": meta.get("using_default"),
            "resolved_default": using_default,
        })
        # #endregion
        st.session_state.pop(cache_key, None)
        st.session_state.pop(meta_key, None)

    ruta_lectura = _resolver_ruta_plan_para_lectura(ruta_resuelta)
    df = cargar_plan_cuentas(ruta_lectura)
    if ruta_lectura != ruta_resuelta:
        try:
            Path(ruta_lectura).unlink(missing_ok=True)
        except OSError:
            pass
    if usar_cache:
        st.session_state[cache_key] = df
        st.session_state[meta_key] = {
            "cliente_id": cliente_id,
            "path": str(ruta_resuelta),
            "using_default": using_default,
            "cuit": cliente.get("cuit"),
            "mtime": mtime,
        }
    _dbg_log("A", "_cargar_plan_cuentas_cliente", "loaded_from_disk", {
        "cliente_id": cliente_id,
        "path": str(ruta_resuelta),
        "using_default": using_default,
        "rows": len(df),
        "suffix": ruta_resuelta.suffix,
        "mtime": mtime,
    })
    return df, ruta_resuelta, using_default


def _vincular_plan_bd_si_corresponde(cliente_id: int, ruta: Path) -> None:
    """Persiste en SQLite la ruta del Excel del cliente cuando se resolvió desde disco."""
    cliente = db.obtener_cliente(cliente_id)
    if not cliente:
        return
    cuit = str(cliente.get("cuit", "")).strip()
    canon = _ruta_plan_xlsx(cuit)
    ruta_str = str(canon if canon.exists() else ruta)
    if (cliente.get("plan_cuentas_path") or "") != ruta_str:
        db.actualizar_cliente(
            cliente_id,
            cliente["nombre"],
            cliente["cuit"],
            cliente["tipo_persona"],
            ruta_str,
            cliente.get("mes_cierre_balance", 12),
        )


def _sincronizar_plan_cuentas_session(cliente_id: int, forzar: bool = False) -> bool:
    """
    Carga el Excel del cliente desde disco a st.session_state.plan_cuentas_df.
    Retorna True si se encontró plan propio (no default).
    """
    cliente = db.obtener_cliente(cliente_id)
    if not cliente:
        return False
    cuit = str(cliente.get("cuit", "")).strip()

    if not forzar:
        for candidata in _rutas_plan_candidatas(cliente):
            if candidata.exists() and _es_plan_propio_cliente(cuit, candidata):
                if bool(st.session_state.get("plan_cuentas_es_default", False)):
                    forzar = True
                    # #region agent log
                    _dbg_log("H", "_sincronizar_plan_cuentas_session", "stale_default_invalidated", {
                        "cliente_id": cliente_id,
                        "cuit": cuit,
                        "found": str(candidata),
                    })
                    # #endregion
                break

    cached_id = st.session_state.get("plan_cuentas_cliente_id")
    plan_ss = st.session_state.get("plan_cuentas_df")
    tiene_df = (
        plan_ss is not None
        and hasattr(plan_ss, "empty")
        and not plan_ss.empty
    )
    if not forzar and cached_id == cliente_id and tiene_df:
        # Si otro usuario actualizó el plan en T:\, el mtime cambia → recargar.
        path_ss = st.session_state.get("plan_cuentas_path_resuelto")
        mtime_ss = st.session_state.get("plan_cuentas_mtime")
        mtime_disk = _mtime_archivo(Path(path_ss)) if path_ss else None
        if path_ss and mtime_ss is not None and mtime_disk is not None and mtime_ss != mtime_disk:
            forzar = True
            _dbg_log("H", "_sincronizar_plan_cuentas_session", "mtime_changed_reload", {
                "cliente_id": cliente_id,
                "path": path_ss,
                "mtime_ss": mtime_ss,
                "mtime_disk": mtime_disk,
            })
        else:
            es_default = bool(st.session_state.get("plan_cuentas_es_default", False))
            _dbg_log("D", "_sincronizar_plan_cuentas_session", "session_cache_hit", {
                "cliente_id": cliente_id,
                "rows": len(st.session_state.plan_cuentas_df),
                "using_default": es_default,
            })
            return not es_default

    if forzar:
        _invalidar_cache_plan(cliente_id)

    df, ruta, es_default = _cargar_plan_cuentas_cliente(cliente_id, usar_cache=True)
    if not es_default:
        _vincular_plan_bd_si_corresponde(cliente_id, ruta)
    st.session_state.plan_cuentas_df = df
    st.session_state.plan_cuentas_cliente_id = cliente_id
    st.session_state.plan_cuentas_path_resuelto = str(ruta)
    st.session_state.plan_cuentas_es_default = es_default
    st.session_state.plan_cuentas_mtime = _mtime_archivo(ruta)
    _dbg_log("B", "_sincronizar_plan_cuentas_session", "session_updated", {
        "cliente_id": cliente_id,
        "path": str(ruta),
        "rows": len(df),
        "using_default": es_default,
        "mtime": st.session_state.plan_cuentas_mtime,
    })
    return not es_default


def actualizar_sociedad_activa() -> None:
    """Callback on_change: actualiza cuit, nombre y plan Excel en un solo viaje."""
    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    if sociedad_id is None:
        return
    cliente = db.obtener_cliente(sociedad_id)
    if not cliente:
        return
    st.session_state.cuit_activo = str(cliente.get("cuit", "")).strip()
    st.session_state.nombre_activo = cliente["nombre"]
    st.session_state.cliente_id_seleccionado = sociedad_id
    _limpiar_resultados_cambio_sociedad()
    _invalidar_cache_plan(sociedad_id)
    _sincronizar_plan_cuentas_session(sociedad_id, forzar=True)
    _inicializar_balance_servidor_por_sociedad()
    _aplicar_ruta_balance_default_sociedad(sociedad_id)
    _auto_cargar_balance_local_si_existe(sociedad_id)
    # #region agent log
    _dbg_log("F", "actualizar_sociedad_activa", "sociedad_synced", {
        "sociedad_id": sociedad_id,
        "cuit_activo": st.session_state.cuit_activo,
        "nombre_activo": st.session_state.nombre_activo,
        "plan_rows": len(st.session_state.plan_cuentas_df) if st.session_state.get("plan_cuentas_df") is not None else 0,
        "plan_cliente_id": st.session_state.get("plan_cuentas_cliente_id"),
    })
    # #endregion


def _verificar_sincronizacion_devengamientos(indice: dict[int, dict]) -> None:
    """Si selector y memoria no coinciden, re-sincroniza (no congela la UI)."""
    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    if sociedad_id is None or sociedad_id not in indice:
        return
    esperado = indice[sociedad_id]
    nombre_mem = st.session_state.get("nombre_activo")
    cuit_mem = st.session_state.get("cuit_activo")
    plan_id = st.session_state.get("plan_cuentas_cliente_id")
    if (
        nombre_mem != esperado["nombre"]
        or cuit_mem != esperado["cuit"]
        or plan_id != sociedad_id
    ):
        _dbg_log("J", "_verificar_sincronizacion_devengamientos", "desync_resync", {
            "sociedad_id": sociedad_id,
            "esperado_nombre": esperado["nombre"],
            "esperado_cuit": esperado["cuit"],
            "nombre_mem": nombre_mem,
            "cuit_mem": cuit_mem,
            "plan_cliente_id": plan_id,
        })
        actualizar_sociedad_activa()


def _selector_sociedad_devengamientos(clientes: list[dict]) -> None:
    """Selectbox único con key permanente sociedad_activa; persiste entre reruns."""
    if not clientes:
        return

    opciones = {c["id"]: f"{c['nombre']} — {c['tipo_persona']}" for c in clientes}
    ids = list(opciones.keys())

    if st.session_state.get(_SOCiedad_KEY) not in ids:
        st.session_state[_SOCiedad_KEY] = ids[0]
        st.session_state.cliente_id_seleccionado = ids[0]

    st.selectbox(
        "Sociedad",
        options=ids,
        format_func=lambda x: opciones[x],
        key=_SOCiedad_KEY,
        on_change=actualizar_sociedad_activa,
    )

    sociedad_id = st.session_state[_SOCiedad_KEY]
    st.session_state.cliente_id_seleccionado = sociedad_id
    # #region agent log
    _dbg_log("I", "_selector_sociedad_devengamientos", "after_selectbox", {
        "sociedad_id": sociedad_id,
    })
    # #endregion


def _resolver_plan_cuentas_cliente(cliente: dict) -> pd.DataFrame:
    plan_df, _, _ = _cargar_plan_cuentas_cliente(cliente["id"])
    return plan_df


def _plan_col_descripcion(plan_df: pd.DataFrame) -> str:
    for col in ("descripcion", "Descripción", "Descripcion"):
        if col in plan_df.columns:
            return col
    return "descripcion"


def _plan_col_codigo(plan_df: pd.DataFrame) -> str:
    for col in ("codigo", "Código", "Codigo"):
        if col in plan_df.columns:
            return col
    return "codigo"


def _norm_desc_iva(texto: str) -> str:
    t = str(texto).lower()
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u")):
        t = t.replace(a, b)
    return t


def _matchea_alicuota_desc(descripcion: str, alicuota: str) -> bool:
    d = _norm_desc_iva(descripcion)
    if alicuota == "10.5":
        return bool(re.search(r"10[,.]?5|10[,.]?50", d))
    if alicuota == "21":
        return bool(re.search(r"(?<!\d)21(?!\d)", d))
    if alicuota == "27":
        return bool(re.search(r"(?<!\d)27(?!\d)", d))
    return alicuota.lower() in d


def _es_linea_ventas_iva(desc: str) -> bool:
    d = _norm_desc_iva(desc)
    return ("debito" in d or "ventas" in d) and ("fiscal" in d or "ventas" in d or "iva" in d)


def _es_alias_debito_iva(d: str) -> bool:
    return "debito" in d or "ventas" in d or bool(re.search(r"\bdf\b", d))


def _es_alias_credito_iva(d: str) -> bool:
    return "credito" in d or "compras" in d or bool(re.search(r"\bcf\b", d))


def _es_cuenta_iva_no_mapeada(codigo: str, descripcion: str) -> bool:
    return str(codigo).strip() == "99999" or str(descripcion).startswith("CUENTA_NO_MAPEADA")


def _plan_base_imputable(plan_cuentas_df: pd.DataFrame) -> pd.DataFrame:
    if plan_cuentas_df is None or plan_cuentas_df.empty:
        return plan_cuentas_df
    base = plan_cuentas_df
    if "imputable" in base.columns:
        filtrada = base[base["imputable"] == True].copy()
        if not filtrada.empty:
            base = filtrada
    if "usa_auxiliares" in base.columns:
        sin_aux = base[base["usa_auxiliares"] != True].copy()
        if not sin_aux.empty:
            base = sin_aux
    return base


def _match_grupo_inclusion(d: str, grupo: tuple[str, ...]) -> bool:
    """Grupo OR con contexto IVA: 'ventas'/'compras' solos no alcanzan."""
    if grupo == ("debito", "ventas", "df") or grupo == ("debito", "ventas", "df", "restitu"):
        if "debito" in d or bool(re.search(r"\bdf\b", d)):
            return True
        if "ventas" in d and ("iva" in d or "fiscal" in d):
            return True
        if grupo[-1] == "restitu" and "restitu" in d and ("iva" in d or "fiscal" in d):
            return True
        return False
    if grupo == ("credito", "compras", "cf") or grupo == ("credito", "compras", "cf", "restitu"):
        if "credito" in d or bool(re.search(r"\bcf\b", d)):
            return True
        if "compras" in d and ("iva" in d or "fiscal" in d):
            return True
        if grupo[-1] == "restitu" and "restitu" in d and ("iva" in d or "fiscal" in d):
            return True
        return False
    if grupo == ("saldo a favor", "libre disponibilidad", "saldo activo"):
        return (
            "saldo a favor" in d
            or "saldo del impuesto a favor" in d
            or "libre disponibilidad" in d
            or ("libre" in d and "disp" in d)
            or "saldo activo" in d
        )
    return any(t in d for t in grupo)


def _alicuota_contraria(alicuota: str) -> str | None:
    if alicuota == "21":
        return "10.5"
    if alicuota == "10.5":
        return "21"
    return None


def _es_cuenta_iva_generica_sin_alicuota(d: str, desc: str) -> bool:
    """Plan Tango habitual: 'IVA Débito Fiscal' / 'IVA Crédito Fiscal' sin % en el nombre."""
    if "iva" not in d:
        return False
    if not (
        "fiscal" in d
        or "debito" in d
        or "credito" in d
        or bool(re.search(r"\bdf\b", d))
        or bool(re.search(r"\bcf\b", d))
    ):
        return False
    if _matchea_alicuota_desc(desc, "21") or _matchea_alicuota_desc(desc, "10.5"):
        return False
    return True


def _excluida_por_regla(d: str, exclusiones: tuple[str, ...]) -> bool:
    return any(ex in d for ex in exclusiones)


REGLAS_CUENTA_IVA: dict[str, dict] = {
    "debito_21": {
        "incluir_grupo": ("debito", "ventas", "df"),
        "alicuota": "21",
        "excluir": (
            "caja", "banco", "credito", "compras", "proveedor",
            "retencion", "percepcion", "moratoria",
            "deudor", "servicios", "productos",
        ),
    },
    "debito_105": {
        "incluir_grupo": ("debito", "ventas", "df"),
        "alicuota": "10.5",
        "excluir": (
            "caja", "banco", "credito", "compras", "proveedor",
            "retencion", "percepcion",
            "deudor", "servicios", "productos",
        ),
    },
    "credito_21": {
        "incluir_grupo": ("credito", "compras", "cf"),
        "alicuota": "21",
        "excluir": (
            "caja", "banco", "debito", "ventas", "cliente",
            "retencion", "percepcion", "pagar",
            "deudor", "servicios", "productos",
        ),
    },
    "credito_105": {
        "incluir_grupo": ("credito", "compras", "cf"),
        "alicuota": "10.5",
        "excluir": (
            "caja", "banco", "debito", "ventas", "cliente",
            "retencion", "percepcion", "pagar",
            "deudor", "servicios", "productos",
        ),
    },
    "retenciones": {
        "incluir_fn": lambda d: ("retencion" in d or bool(re.search(r"\bret\b", d))) and "iva" in d,
        "excluir": ("iibb", "ganancias", "suss", "bancaria", "pagar"),
    },
    "percepciones": {
        "incluir_fn": lambda d: ("percepcion" in d or "percep" in d) and "iva" in d,
        "excluir": ("iibb", "ganancias", "aduanera", "arba", "agip"),
    },
    "iva_pagar": {
        "incluir_todos": ("iva",),
        "incluir_alguno": ("pagar", "saldo pasivo", "subcuenta oblig"),
        "excluir": ("credito", "debito", "gasto", "activo"),
    },
    "saldo_favor": {
        "incluir_grupo": ("saldo a favor", "libre disponibilidad", "saldo activo"),
        "incluir_todos": ("iva",),
        "excluir": ("debito", "pagar", "iibb"),
    },
    "nc_compras": {
        "incluir_grupo": ("credito", "compras", "cf", "restitu"),
        "excluir": (
            "caja", "banco", "debito", "ventas", "cliente",
            "pagar", "retencion", "percepcion",
        ),
    },
    "nc_ventas": {
        "incluir_grupo": ("debito", "ventas", "df", "restitu"),
        "excluir": (
            "caja", "banco", "credito", "compras", "proveedor",
            "pagar", "retencion", "percepcion",
        ),
    },
    "debito_27": {
        "incluir_grupo": ("debito", "ventas", "df"),
        "alicuota": "27",
        "excluir": (
            "caja", "banco", "credito", "compras", "proveedor",
            "retencion", "percepcion", "moratoria",
            "deudor", "servicios", "productos", "21", "10.5", "10,5",
        ),
    },
    "credito_27": {
        "incluir_grupo": ("credito", "compras", "cf"),
        "alicuota": "27",
        "excluir": (
            "caja", "banco", "debito", "ventas", "cliente",
            "retencion", "percepcion", "pagar",
            "deudor", "servicios", "productos", "21", "10.5", "10,5",
        ),
    },
    "nc_compras_27": {
        "incluir_grupo": ("credito", "compras", "cf", "restitu"),
        "alicuota": "27",
        "excluir": (
            "caja", "banco", "debito", "ventas", "cliente",
            "pagar", "retencion", "percepcion", "21", "10.5", "10,5",
        ),
    },
    "nc_ventas_27": {
        "incluir_grupo": ("debito", "ventas", "df", "restitu"),
        "alicuota": "27",
        "excluir": (
            "caja", "banco", "credito", "compras", "proveedor",
            "pagar", "retencion", "percepcion", "21", "10.5", "10,5",
        ),
    },
}


REGLAS_CUENTA_IIBB: dict[str, dict] = {
    "impuesto_determinado": {
        "incluir_todos": ("iibb",),
        "incluir_alguno": ("impuesto", "gasto", "devengado", "determinado"),
        "excluir": (
            "caja", "banco", "pagar", "retencion", "percepcion", "favor", "iva",
        ),
    },
    "retenciones_iibb": {
        "incluir_fn": lambda d: (
            ("retencion" in d or bool(re.search(r"\bret\b", d)))
            and "iibb" in d
            and "bancaria" not in d
            and "sircreb" not in d
            and "creb" not in d
        ),
        "excluir": ("iva", "ganancias", "suss", "pagar", "percepcion"),
    },
    "retenciones_bancarias": {
        "incluir_fn": lambda d: (
            "bancaria" in d or "sircreb" in d or "creb" in d
        ) and ("iibb" in d or "retencion" in d or "ingresos brutos" in d),
        "excluir": ("iva", "ganancias", "pagar"),
    },
    "percepciones_iibb": {
        "incluir_fn": lambda d: (
            ("percepcion" in d or "percep" in d) and "iibb" in d
        ),
        "excluir": ("iva", "ganancias", "aduanera", "pagar", "retencion"),
    },
    "saldo_favor_anterior": {
        "incluir_fn": lambda d: (
            "saldo" in d and "favor" in d and "iibb" in d and "nuevo" not in d
        ),
        "excluir": ("iva", "pagar", "retencion", "percepcion"),
    },
    "iibb_pagar": {
        "incluir_todos": ("iibb",),
        "incluir_alguno": ("pagar",),
        "excluir": ("retencion", "percepcion", "favor", "iva", "devengado", "determinado"),
    },
    "saldo_favor_nuevo": {
        "incluir_fn": lambda d: (
            "saldo" in d and "favor" in d and "iibb" in d
        ),
        "excluir": ("iva", "anterior", "pagar", "retencion", "percepcion"),
    },
}


REGLAS_CUENTA_SUELDOS: dict[str, dict] = {
    "sueldos_jornales": {
        "incluir_fn": lambda d: (
            ("sueldo" in d or "jornal" in d or "remuneracion" in d)
            and "pagar" not in d
            and "carga" not in d
        ),
        "excluir": ("iva", "iibb", "tish", "sindicato", "f931"),
    },
    "sueldos_pagar": {
        "incluir_fn": lambda d: (
            ("sueldo" in d or "jornal" in d) and "pagar" in d
        ),
        "excluir": ("iva", "iibb", "carga"),
    },
    "cargas_sociales_pagar": {
        "incluir_fn": lambda d: (
            "carga" in d and "social" in d and "pagar" in d
        ),
        "excluir": ("iva", "iibb", "deveng"),
    },
    "aportes": {
        "incluir_fn": lambda d: "aporte" in d or "f931" in d,
        "excluir": ("iva", "iibb", "pagar"),
    },
    "contribuciones": {
        "incluir_fn": lambda d: "contribucion" in d,
        "excluir": ("iva", "iibb", "pagar"),
    },
    "sueldos_saldo_favor": {
        "incluir_fn": lambda d: "saldo" in d and "favor" in d and ("sueldo" in d or "asignacion" in d),
        "excluir": ("iva", "iibb", "tish"),
    },
    "sueldos_pagar_cierre": {
        "incluir_fn": lambda d: ("sueldo" in d or "jornal" in d) and "pagar" in d,
        "excluir": ("iva", "iibb", "carga"),
    },
}


REGLAS_CUENTA_TISH: dict[str, dict] = {
    "gasto_tasa": {
        "incluir_fn": lambda d: (
            ("tasa" in d or "tish" in d or "seguridad" in d or "higiene" in d)
            and ("determinada" in d or "deveng" in d or "gasto" in d)
            and "pagar" not in d
            and "retencion" not in d
        ),
        "excluir": ("iva", "iibb", "sueldo"),
    },
    "tasa_pagar": {
        "incluir_fn": lambda d: (
            ("tasa" in d or "tish" in d) and "pagar" in d
        ),
        "excluir": ("iva", "iibb", "retencion"),
    },
    "retenciones_tish": {
        "incluir_fn": lambda d: (
            ("retencion" in d or "retenciones" in d)
            and ("tish" in d or "seguridad" in d or "higiene" in d or "municipal" in d)
        ),
        "excluir": ("iva", "iibb", "sueldo"),
    },
    "derecho_oficina": {
        "incluir_fn": lambda d: "derecho" in d and ("oficina" in d or "municipal" in d),
        "excluir": ("iva", "iibb"),
    },
    "tish_saldo_favor": {
        "incluir_fn": lambda d: "saldo" in d and "favor" in d and ("tish" in d or "tasa" in d),
        "excluir": ("iva", "iibb", "anterior"),
    },
    "tish_saldo_favor_nuevo": {
        "incluir_fn": lambda d: "saldo" in d and "favor" in d and ("tish" in d or "tasa" in d),
        "excluir": ("iva", "iibb", "anterior", "pagar"),
    },
}


def _cumple_regla_inclusion_base(d: str, regla: dict) -> bool:
    if regla.get("incluir_fn"):
        return bool(regla["incluir_fn"](d))
    grupo = regla.get("incluir_grupo")
    if grupo and not _match_grupo_inclusion(d, grupo):
        return False
    todos = regla.get("incluir_todos")
    if todos and not all(t in d for t in todos):
        return False
    alguno = regla.get("incluir_alguno")
    if alguno and not any(t in d for t in alguno):
        return False
    return True


def _puntaje_candidato_iva(
    d: str,
    desc: str,
    regla: dict,
    subtipo: str | None = None,
) -> int:
    """Puntaje del candidato; -1 = descartado."""
    if not _cumple_regla_inclusion_base(d, regla):
        return -1

    score = 1
    ali = regla.get("alicuota")
    if ali:
        if _matchea_alicuota_desc(desc, ali):
            score += 6
        else:
            contra = _alicuota_contraria(ali)
            if contra and _matchea_alicuota_desc(desc, contra):
                return -1
            if ali == "21" and _es_cuenta_iva_generica_sin_alicuota(d, desc):
                score += 3
            else:
                return -1

    if "%" in desc:
        score += 1
    if "fiscal" in d:
        score += 2
    if "iva" in d:
        score += 1
    if regla.get("incluir_alguno"):
        for tok in regla["incluir_alguno"]:
            if tok in d:
                score += 2
    if subtipo == "tecnico" and "tecnico" in d:
        score += 3
    elif subtipo == "libre" and "libre" in d:
        score += 3
    elif subtipo == "saldo_favor":
        if "saldo a favor" in d or "saldo del impuesto a favor" in d:
            score += 2
    return score


def _cumple_regla_inclusion(
    d: str,
    desc: str,
    regla: dict,
    *,
    exigir_alicuota: bool = True,
) -> bool:
    if not _cumple_regla_inclusion_base(d, regla):
        return False
    ali = regla.get("alicuota")
    if exigir_alicuota and ali:
        if _matchea_alicuota_desc(desc, ali):
            return True
        if ali == "21" and _es_cuenta_iva_generica_sin_alicuota(d, desc):
            return True
        return False
    if ali:
        contra = _alicuota_contraria(ali)
        if contra and _matchea_alicuota_desc(desc, contra):
            return False
    return True


def _rol_a_tipo_concepto(rol: str, alicuota: str | None = None) -> str:
    if rol in REGLAS_CUENTA_IVA:
        return rol
    if rol in ("ventas_21", "ventas"):
        return "debito_21" if (alicuota or "21") != "10.5" else "debito_105"
    if rol in ("ventas_105",):
        return "debito_105"
    if rol in ("ventas_27",):
        return "debito_27"
    if rol in ("compras_21", "compras"):
        return "credito_21" if (alicuota or "21") != "10.5" else "credito_105"
    if rol in ("compras_105",):
        return "credito_105"
    if rol in ("compras_27",):
        return "credito_27"
    mapping = {
        "nc_compras": "nc_compras",
        "nc_compras_27": "nc_compras_27",
        "nc_credito": "nc_compras",
        "nc_ventas": "nc_ventas",
        "nc_ventas_27": "nc_ventas_27",
        "retenciones": "retenciones",
        "percepciones": "percepciones",
        "pagar": "iva_pagar",
        "tecnico": "saldo_favor",
        "libre": "saldo_favor",
        "saldo_favor": "saldo_favor",
    }
    return mapping.get(rol, rol)


def _nombre_concepto_rescate(tipo_concepto: str) -> str:
    nombres = {
        "debito_21": "IVA Débito Fiscal 21%",
        "debito_105": "IVA Débito Fiscal 10,5%",
        "debito_27": "IVA Débito Fiscal 27%",
        "credito_21": "IVA Crédito Fiscal 21%",
        "credito_105": "IVA Crédito Fiscal 10,5%",
        "credito_27": "IVA Crédito Fiscal 27%",
        "retenciones": "Retenciones IVA",
        "percepciones": "Percepciones IVA",
        "iva_pagar": "IVA a Pagar",
        "saldo_favor": "Saldo a Favor IVA Nuevo Período",
        "nc_compras": "NC Compras IVA",
        "nc_compras_27": "NC Compras IVA 27%",
        "nc_ventas": "NC Ventas IVA",
        "nc_ventas_27": "NC Ventas IVA 27%",
    }
    return nombres.get(tipo_concepto, tipo_concepto.replace("_", " ").title())


def obtener_cuenta_tango(
    plan_cuentas_df: pd.DataFrame,
    tipo_concepto: str,
    alicuota: str | None = None,
    subtipo: str | None = None,
) -> tuple[str, str]:
    """Algoritmo universal de búsqueda contable con inclusiones y exclusiones rigurosas."""
    regla = REGLAS_CUENTA_IVA.get(tipo_concepto)
    if regla is None:
        return "99999", f"CUENTA_NO_MAPEADA_{tipo_concepto}"

    plan_base = _plan_base_imputable(plan_cuentas_df)
    col_desc = _plan_col_descripcion(plan_cuentas_df)
    col_cod = _plan_col_codigo(plan_cuentas_df)
    exclusiones = regla.get("excluir", ())
    hint = subtipo or (tipo_concepto if tipo_concepto in ("tecnico", "libre") else None)

    mejor_cod, mejor_desc, mejor_score = "", "", -1
    for _, row in plan_base.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if not cod:
            continue
        d = _norm_desc_iva(desc)
        if _excluida_por_regla(d, exclusiones):
            continue
        score = _puntaje_candidato_iva(d, desc, regla, hint)
        if score > mejor_score:
            mejor_score, mejor_cod, mejor_desc = score, cod, desc

    if mejor_score > 0:
        return mejor_cod, mejor_desc
    return "99999", _nombre_concepto_rescate(tipo_concepto)


def _puntaje_candidato_iibb(d: str, desc: str, regla: dict) -> int:
    """Puntaje del candidato IIBB; -1 = descartado."""
    if not _cumple_regla_inclusion_base(d, regla):
        return -1
    score = 1
    if "iibb" in d or "ingresos brutos" in d or "ingreso bruto" in d:
        score += 3
    if "fiscal" in d or "impuesto" in d:
        score += 1
    if regla.get("incluir_alguno"):
        for tok in regla["incluir_alguno"]:
            if tok in d:
                score += 2
    if "%" in desc:
        score += 1
    return score


def _nombre_concepto_rescate_iibb(tipo_concepto: str) -> str:
    nombres = {
        "impuesto_determinado": "IIBB Devengado / Impuesto Determinado",
        "retenciones_iibb": "Retenciones IIBB",
        "retenciones_bancarias": "Retenciones Bancarias IIBB (Sircreb)",
        "percepciones_iibb": "Percepciones IIBB",
        "saldo_favor_anterior": "Saldo a Favor IIBB Período Anterior",
        "iibb_pagar": "IIBB a Pagar",
        "saldo_favor_nuevo": "Saldo a Favor IIBB Nuevo Período",
    }
    return nombres.get(tipo_concepto, tipo_concepto.replace("_", " ").title())


def obtener_cuenta_tango_iibb(
    plan_cuentas_df: pd.DataFrame,
    tipo_concepto: str,
) -> tuple[str, str]:
    """Algoritmo de búsqueda contable para cuentas IIBB."""
    regla = REGLAS_CUENTA_IIBB.get(tipo_concepto)
    if regla is None:
        return "99999", f"CUENTA_NO_MAPEADA_{tipo_concepto}"

    plan_base = _plan_base_imputable(plan_cuentas_df)
    col_desc = _plan_col_descripcion(plan_cuentas_df)
    col_cod = _plan_col_codigo(plan_cuentas_df)
    exclusiones = regla.get("excluir", ())

    mejor_cod, mejor_desc, mejor_score = "", "", -1
    for _, row in plan_base.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if not cod:
            continue
        d = _norm_desc_iva(desc)
        if _excluida_por_regla(d, exclusiones):
            continue
        score = _puntaje_candidato_iibb(d, desc, regla)
        if score > mejor_score:
            mejor_score, mejor_cod, mejor_desc = score, cod, desc

    if mejor_score > 0:
        return mejor_cod, mejor_desc
    return "99999", _nombre_concepto_rescate_iibb(tipo_concepto)


def _puntaje_candidato_generico(d: str, desc: str, regla: dict, bonus_tokens: tuple[str, ...] = ()) -> int:
    if not _cumple_regla_inclusion_base(d, regla):
        return -1
    score = 1
    for tok in bonus_tokens:
        if tok in d:
            score += 2
    if regla.get("incluir_alguno"):
        for tok in regla["incluir_alguno"]:
            if tok in d:
                score += 2
    if "%" in desc:
        score += 1
    return score


def _obtener_cuenta_tango_por_reglas(
    plan_cuentas_df: pd.DataFrame,
    tipo_concepto: str,
    reglas: dict[str, dict],
    nombre_rescate_fn,
    bonus_tokens: tuple[str, ...] = (),
) -> tuple[str, str]:
    regla = reglas.get(tipo_concepto)
    if regla is None:
        return "99999", f"CUENTA_NO_MAPEADA_{tipo_concepto}"
    plan_base = _plan_base_imputable(plan_cuentas_df)
    col_desc = _plan_col_descripcion(plan_cuentas_df)
    col_cod = _plan_col_codigo(plan_cuentas_df)
    exclusiones = regla.get("excluir", ())
    mejor_cod, mejor_desc, mejor_score = "", "", -1
    for _, row in plan_base.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if not cod:
            continue
        d = _norm_desc_iva(desc)
        if _excluida_por_regla(d, exclusiones):
            continue
        score = _puntaje_candidato_generico(d, desc, regla, bonus_tokens)
        if score > mejor_score:
            mejor_score, mejor_cod, mejor_desc = score, cod, desc
    if mejor_score > 0:
        return mejor_cod, mejor_desc
    return "99999", nombre_rescate_fn(tipo_concepto)


def _nombre_concepto_rescate_sueldos(tipo_concepto: str) -> str:
    nombres = {
        "sueldos_jornales": "Sueldos y Jornales",
        "sueldos_pagar": "Sueldos y Jornales a Pagar",
        "cargas_sociales_pagar": "Cargas Sociales a Pagar",
        "aportes": "Aportes / F931",
        "contribuciones": "Contribuciones Patronales",
        "sueldos_saldo_favor": "Saldo a Favor Sueldos",
        "sueldos_pagar_cierre": "Sueldos a Pagar (cierre)",
    }
    return nombres.get(tipo_concepto, tipo_concepto.replace("_", " ").title())


def obtener_cuenta_tango_sueldos(plan_cuentas_df: pd.DataFrame, tipo_concepto: str) -> tuple[str, str]:
    return _obtener_cuenta_tango_por_reglas(
        plan_cuentas_df, tipo_concepto, REGLAS_CUENTA_SUELDOS,
        _nombre_concepto_rescate_sueldos, ("sueldo", "jornal", "remuneracion"),
    )


def _nombre_concepto_rescate_tish(tipo_concepto: str) -> str:
    nombres = {
        "gasto_tasa": "Gasto Tasa / Tasa Determinada TISH",
        "tasa_pagar": "Tasa TISH a Pagar",
        "retenciones_tish": "Retenciones TISH",
        "derecho_oficina": "Derecho de Oficina",
        "tish_saldo_favor": "Saldo a Favor TISH Período Anterior",
        "tish_saldo_favor_nuevo": "Saldo a Favor TISH Nuevo Período",
    }
    return nombres.get(tipo_concepto, tipo_concepto.replace("_", " ").title())


def obtener_cuenta_tango_tish(plan_cuentas_df: pd.DataFrame, tipo_concepto: str) -> tuple[str, str]:
    return _obtener_cuenta_tango_por_reglas(
        plan_cuentas_df, tipo_concepto, REGLAS_CUENTA_TISH,
        _nombre_concepto_rescate_tish, ("tish", "tasa", "seguridad", "higiene"),
    )


def _opciones_cuentas_rescate_iibb(
    plan_df: pd.DataFrame,
    tipo_concepto: str,
) -> list[tuple[str, str]]:
    """Lista filtrada para selectbox de rescate IIBB."""
    regla = REGLAS_CUENTA_IIBB.get(tipo_concepto)
    if regla is None or plan_df is None or plan_df.empty:
        return []

    plan_base = _plan_base_imputable(plan_df)
    col_desc = _plan_col_descripcion(plan_df)
    col_cod = _plan_col_codigo(plan_df)
    exclusiones = regla.get("excluir", ())
    candidatos: list[tuple[int, str, str]] = []

    for _, row in plan_base.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if not cod:
            continue
        d = _norm_desc_iva(desc)
        if _excluida_por_regla(d, exclusiones):
            continue
        score = _puntaje_candidato_iibb(d, desc, regla)
        if score < 0:
            continue
        candidatos.append((score, cod, desc))

    candidatos.sort(key=lambda x: (-x[0], x[1]))
    return [(cod, f"{cod} - {desc}") for _, cod, desc in candidatos]


def _rol_a_tipo_concepto_iibb(rol: str) -> str:
    mapping = {
        "impuesto_determinado": "impuesto_determinado",
        "retenciones_iibb": "retenciones_iibb",
        "retenciones_bancarias": "retenciones_bancarias",
        "percepciones_iibb": "percepciones_iibb",
        "saldo_favor_anterior": "saldo_favor_anterior",
        "iibb_pagar": "iibb_pagar",
        "saldo_favor_nuevo": "saldo_favor_nuevo",
    }
    return mapping.get(rol, rol)


def _mapear_cuenta_tango_iibb(plan_df: pd.DataFrame, rol: str) -> tuple[str, str]:
    tipo = _rol_a_tipo_concepto_iibb(rol)
    return obtener_cuenta_tango_iibb(plan_df, tipo)


def _opciones_cuentas_rescate(
    plan_df: pd.DataFrame,
    tipo_concepto: str,
    subtipo: str | None = None,
) -> list[tuple[str, str]]:
    """Lista filtrada para selectbox de rescate: solo cuentas del concepto IVA activo."""
    regla = REGLAS_CUENTA_IVA.get(tipo_concepto)
    if regla is None or plan_df is None or plan_df.empty:
        return []

    plan_base = _plan_base_imputable(plan_df)
    col_desc = _plan_col_descripcion(plan_df)
    col_cod = _plan_col_codigo(plan_df)
    exclusiones = regla.get("excluir", ())
    hint = subtipo
    candidatos: list[tuple[int, str, str]] = []

    for _, row in plan_base.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if not cod:
            continue
        d = _norm_desc_iva(desc)
        if _excluida_por_regla(d, exclusiones):
            continue
        score = _puntaje_candidato_iva(d, desc, regla, hint)
        if score < 0:
            continue
        candidatos.append((score, cod, desc))

    candidatos.sort(key=lambda x: (-x[0], x[1]))
    return [(cod, f"{cod} - {desc}") for _, cod, desc in candidatos]


def _buscar_cuenta_iva_alias(
    plan_base: pd.DataFrame,
    col_desc: str,
    col_cod: str,
    rol: str,
    alicuota: str | None = None,
) -> tuple[str, str]:
    """Compatibilidad con tests: delega al motor universal obtener_cuenta_tango."""
    tipo = _rol_a_tipo_concepto(rol, alicuota)
    subtipo = rol if rol in ("tecnico", "libre") else None
    return obtener_cuenta_tango(plan_base, tipo, alicuota, subtipo=subtipo)


def _mapear_cuenta_tango(plan_df: pd.DataFrame, rol: str) -> tuple[str, str]:
    """Mapeo de cuentas IVA al plan Tango de la sociedad activa."""
    tipo = _rol_a_tipo_concepto(rol)
    subtipo = rol if rol in ("tecnico", "libre", "saldo_favor") else None
    return obtener_cuenta_tango(plan_df, tipo, subtipo=subtipo)


def _extraer_monto_de_texto(linea: str) -> float:
    matches = re.findall(r"[\d\.]+,[\d]{2}", linea)
    if matches:
        return float(matches[-1].replace(".", "").replace(",", "."))
    return 0.0


def _normalizar_lista_saldos_iva(valor: float | list[float] | None) -> list[float]:
    """Convierte float legacy o lista dinámica en montos positivos individuales."""
    if valor is None:
        return []
    if isinstance(valor, list):
        return [round(float(x), 2) for x in valor if round(float(x), 2) > 0]
    v = round(float(valor), 2)
    return [v] if v > 0 else []


def _calcular_desglose_posicion_iva(
    *,
    df_debito_21: float,
    df_debito_105: float,
    df_debito_27: float = 0.0,
    nc_compras: float,
    nc_compras_27: float = 0.0,
    cf_credito_21: float,
    cf_credito_105: float,
    cf_credito_27: float = 0.0,
    nc_ventas: float,
    nc_ventas_27: float = 0.0,
    retenciones: float,
    percepciones: float,
    saldos_tecnicos: list[float],
    saldos_libre: list[float],
) -> dict:
    """Ecuación de balance DDJJ: posición previa antes de la fila de cierre."""
    total_debito_mes = round(
        df_debito_21 + df_debito_105 + df_debito_27 + nc_compras + nc_compras_27, 2
    )
    total_credito_mes = round(
        cf_credito_21 + cf_credito_105 + cf_credito_27
        + nc_ventas + nc_ventas_27 + retenciones + percepciones, 2
    )
    tecnicos = [round(float(x), 2) for x in saldos_tecnicos if round(float(x), 2) > 0]
    libres = [round(float(x), 2) for x in saldos_libre if round(float(x), 2) > 0]
    total_saldos_ant = round(sum(tecnicos) + sum(libres), 2)
    total_haber_previo = round(total_credito_mes + total_saldos_ant, 2)
    diferencia_previa = round(total_debito_mes - total_haber_previo, 2)
    movimiento_puro = round(total_debito_mes - total_credito_mes, 2)

    if diferencia_previa > 0:
        resultado_tipo = "IVA a Pagar"
        resultado_lado = "Haber"
        resultado_monto = diferencia_previa
    elif diferencia_previa < 0:
        resultado_tipo = "Saldo a Favor IVA Nuevo Período"
        resultado_lado = "Debe"
        resultado_monto = abs(diferencia_previa)
    else:
        resultado_tipo = "Equilibrado (sin saldo de cierre)"
        resultado_lado = "—"
        resultado_monto = 0.0

    return {
        "total_debito_mes": total_debito_mes,
        "total_credito_mes": total_credito_mes,
        "movimiento_puro_mes": movimiento_puro,
        "saldos_tecnicos": tecnicos,
        "saldos_libre": libres,
        "total_saldos_anteriores": total_saldos_ant,
        "total_haber_previo": total_haber_previo,
        "diferencia_previa": diferencia_previa,
        "resultado_tipo": resultado_tipo,
        "resultado_lado": resultado_lado,
        "resultado_monto": resultado_monto,
    }


def _armar_lineas_movimiento_iva(
    plan_cuentas_df: pd.DataFrame,
    resumen: dict,
    roles_debe: tuple[tuple[str, float], ...],
    roles_haber: tuple[tuple[str, float], ...],
) -> list[dict]:
    """Arma líneas del mes + saldos arrastre + única fila de posición DDJJ."""
    lineas: list[dict] = []

    for rol, monto in roles_debe:
        if monto > 0:
            cod, desc = _mapear_cuenta_tango(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iva(cod, desc, monto, 0.0, rol))

    for rol, monto in roles_haber:
        if monto > 0:
            cod, desc = _mapear_cuenta_tango(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iva(cod, desc, 0.0, monto, rol))

    for monto in resumen.get("saldos_tecnicos", []):
        cod, desc = _mapear_cuenta_tango(plan_cuentas_df, "tecnico")
        lineas.append(_linea_asiento_iva(cod, desc, 0.0, monto, "tecnico"))

    for monto in resumen.get("saldos_libre", []):
        cod, desc = _mapear_cuenta_tango(plan_cuentas_df, "libre")
        lineas.append(_linea_asiento_iva(cod, desc, 0.0, monto, "libre"))

    dif = resumen["diferencia_previa"]
    if dif > 0:
        cod, desc = _mapear_cuenta_tango(plan_cuentas_df, "pagar")
        lineas.append(_linea_asiento_iva(cod, desc, 0.0, dif, "pagar"))
    elif dif < 0:
        cod, desc = _mapear_cuenta_tango(plan_cuentas_df, "saldo_favor")
        lineas.append(_linea_asiento_iva(
            cod, "Saldo a Favor IVA Nuevo Período", abs(dif), 0.0, "saldo_favor",
        ))

    return lineas


def _normalizar_lista_saldos_iibb(valor: float | list[float] | None) -> list[float]:
    if valor is None:
        return []
    if isinstance(valor, list):
        return [round(float(x), 2) for x in valor if round(float(x), 2) > 0]
    v = round(float(valor), 2)
    return [v] if v > 0 else []


def _calcular_desglose_posicion_iibb(
    *,
    impuesto_determinado: float,
    retenciones: float,
    percepciones: float,
    retenciones_bancarias: float,
    saldos_favor: list[float],
) -> dict:
    """Posición IIBB = impuesto_determinado − (ret + perc + bancarias + saldos favor)."""
    saldos = [round(float(x), 2) for x in saldos_favor if round(float(x), 2) > 0]
    total_descuentos = round(
        retenciones + percepciones + retenciones_bancarias + sum(saldos), 2
    )
    imp = round(float(impuesto_determinado), 2)
    diferencia_previa = round(imp - total_descuentos, 2)

    if diferencia_previa > 0:
        resultado_tipo = "IIBB a Pagar"
        resultado_lado = "Haber"
        resultado_monto = diferencia_previa
    elif diferencia_previa < 0:
        resultado_tipo = "Saldo a Favor IIBB Nuevo Período"
        resultado_lado = "Debe"
        resultado_monto = abs(diferencia_previa)
    else:
        resultado_tipo = "Equilibrado (sin saldo de cierre)"
        resultado_lado = "—"
        resultado_monto = 0.0

    return {
        "impuesto_determinado": imp,
        "retenciones": round(retenciones, 2),
        "percepciones": round(percepciones, 2),
        "retenciones_bancarias": round(retenciones_bancarias, 2),
        "saldos_favor": saldos,
        "total_saldos_anteriores": round(sum(saldos), 2),
        "total_descuentos": total_descuentos,
        "diferencia_previa": diferencia_previa,
        "resultado_tipo": resultado_tipo,
        "resultado_lado": resultado_lado,
        "resultado_monto": resultado_monto,
    }


def _linea_asiento_iibb(
    cod: str,
    desc: str,
    debe: float,
    haber: float,
    rol: str = "",
) -> dict:
    return {
        "Cuenta": cod,
        "Detalle": desc,
        "Debe": round(debe, 2),
        "Haber": round(haber, 2),
        "Estado": "Ingresado",
        "_rol": rol,
    }


def _armar_lineas_movimiento_iibb(
    plan_cuentas_df: pd.DataFrame,
    resumen: dict,
) -> list[dict]:
    """Arma líneas IIBB: impuesto determinado, ret/perc/bancarias, saldos favor y cierre."""
    lineas: list[dict] = []

    imp = resumen.get("impuesto_determinado", 0.0)
    if imp > 0:
        cod, desc = _mapear_cuenta_tango_iibb(plan_cuentas_df, "impuesto_determinado")
        lineas.append(_linea_asiento_iibb(cod, desc, imp, 0.0, "impuesto_determinado"))

    for rol, monto_key in (
        ("retenciones_iibb", "retenciones"),
        ("percepciones_iibb", "percepciones"),
        ("retenciones_bancarias", "retenciones_bancarias"),
    ):
        monto = resumen.get(monto_key, 0.0)
        if monto > 0:
            cod, desc = _mapear_cuenta_tango_iibb(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, rol))

    for monto in resumen.get("saldos_favor", []):
        cod, desc = _mapear_cuenta_tango_iibb(plan_cuentas_df, "saldo_favor_anterior")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, "saldo_favor_anterior"))

    dif = resumen["diferencia_previa"]
    if dif > 0:
        cod, desc = _mapear_cuenta_tango_iibb(plan_cuentas_df, "iibb_pagar")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, dif, "iibb_pagar"))
    elif dif < 0:
        cod, desc = _mapear_cuenta_tango_iibb(plan_cuentas_df, "saldo_favor_nuevo")
        lineas.append(_linea_asiento_iibb(
            cod, "Saldo a Favor IIBB Nuevo Período", abs(dif), 0.0, "saldo_favor_nuevo",
        ))

    return lineas


def _resolver_monto_manual_o_planilla(manual: float, planilla: float) -> float:
    """Manual override; si es cero, usa valor leído de la planilla."""
    m = round(float(manual or 0), 2)
    p = round(float(planilla or 0), 2)
    return m if m > 0 else p


def _generar_asiento_iibb_planilla_posicional(
    plan_cuentas_df: pd.DataFrame,
    datos: dict,
    saldos_favor: float | list[float],
    retenciones: float,
    percepciones: float,
    retenciones_bancarias: float,
) -> tuple[list[dict], dict, float, bool]:
    saldos = _normalizar_lista_saldos_iibb(saldos_favor)
    ret = _resolver_monto_manual_o_planilla(
        retenciones, datos.get("retenciones_planilla", 0.0),
    )
    perc = _resolver_monto_manual_o_planilla(
        percepciones, datos.get("percepciones_planilla", 0.0),
    )
    banc = _resolver_monto_manual_o_planilla(
        retenciones_bancarias, datos.get("retenciones_bancarias_planilla", 0.0),
    )
    resumen = _calcular_desglose_posicion_iibb(
        impuesto_determinado=datos.get("impuesto_determinado", 0.0),
        retenciones=ret,
        percepciones=perc,
        retenciones_bancarias=banc,
        saldos_favor=saldos,
    )
    lineas = _armar_lineas_movimiento_iibb(plan_cuentas_df, resumen)
    lineas, dif_ajuste, ajustado = _aplicar_loop_review_impuesto_determinado(lineas)
    return lineas, resumen, dif_ajuste, ajustado


def _fmt_pesos_ar(monto: float) -> str:
    """Importe en pesos argentinos (evita que Streamlit muestre 'dólares')."""
    neg = monto < 0
    s = f"{abs(monto):,.2f}"
    entero, dec = s.split(".")
    entero = entero.replace(",", ".")
    pref = "-" if neg else ""
    return f"{pref}$ {entero},{dec}"


def _obtener_valores_saldos_iva_desde_session(lista_key: str, prefix: str) -> list[float]:
    """Lee los montos reales de los widgets dinámicos ➕ en session_state."""
    rc = _iva_reset_counter()
    lista = st.session_state.get(lista_key) or [0.0]
    return [
        round(float(st.session_state.get(f"{prefix}_{rc}_{i}", 0.0)), 2)
        for i in range(len(lista))
    ]


def _render_cuadro_control_analitico_iva(
    resumen: dict | None,
    rows: list[dict],
) -> None:
    """Cuadro de control con desglose de arrastre y balance final en pesos."""
    if resumen is None:
        resumen = st.session_state.get("iva_resumen_analitico") or {}

    total_debe = round(sum(float(r.get("Debe") or 0) for r in rows), 2)
    total_haber = round(sum(float(r.get("Haber") or 0) for r in rows), 2)
    diferencia = round(total_debe - total_haber, 2)

    with st.container(border=True):
        st.markdown("### 📊 Resumen Analítico de la Liquidación")

        st.markdown("#### 🟢 Movimiento Puro del Mes (lo que genera el período)")
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.markdown(f"**Total Débitos Fiscales**  \n{_fmt_pesos_ar(resumen.get('total_debito_mes', 0))}")
        c_m2.markdown(f"**Total Créditos / Ret. / Perc.**  \n{_fmt_pesos_ar(resumen.get('total_credito_mes', 0))}")
        mov = resumen.get("movimiento_puro_mes", 0)
        c_m3.markdown(f"**Neto del Mes (D − C)**  \n{_fmt_pesos_ar(mov)}")

        st.markdown("#### 🔵 Detalle de Saldos que Arrastran (períodos anteriores)")
        detalle_saldos: list[str] = []
        for i, m in enumerate(resumen.get("saldos_tecnicos") or [], start=1):
            detalle_saldos.append(f"Saldo Técnico #{i}: {_fmt_pesos_ar(m)}")
        for i, m in enumerate(resumen.get("saldos_libre") or [], start=1):
            detalle_saldos.append(f"Saldo Libre Disponibilidad #{i}: {_fmt_pesos_ar(m)}")
        if detalle_saldos:
            for linea in detalle_saldos:
                st.markdown(f"- {linea}")
            st.markdown(
                f"**Subtotal arrastre:** {_fmt_pesos_ar(resumen.get('total_saldos_anteriores', 0))}"
            )
        else:
            st.caption("Sin saldos de períodos anteriores cargados.")

        st.markdown("#### 🏁 Resultado Final de la DDJJ")
        r1, r2, r3 = st.columns(3)
        r1.markdown(f"**Concepto imputado**  \n{resumen.get('resultado_tipo', '—')}")
        r2.markdown(f"**Lado contable**  \n{resumen.get('resultado_lado', '—')}")
        r3.markdown(f"**Importe de cierre**  \n{_fmt_pesos_ar(resumen.get('resultado_monto', 0))}")

        st.markdown("---")
        st.markdown("**Control de partida doble (asiento generado)**")
        t1, t2, t3 = st.columns(3)
        t1.markdown(f"**Total Debe**  \n{_fmt_pesos_ar(total_debe)}")
        t2.markdown(f"**Total Haber**  \n{_fmt_pesos_ar(total_haber)}")
        color = "green" if diferencia == 0 else "red"
        t3.markdown(
            f"**Diferencia**  \n"
            f":{color}[{_fmt_pesos_ar(diferencia)}]"
        )
        if diferencia != 0:
            st.warning(
                f"El asiento aún no cierra a cero ({_fmt_pesos_ar(diferencia)}). "
                "Revisá la grilla o esperá el ajuste automático de redondeo."
            )


def _linea_asiento_iva(
    cod: str,
    desc: str,
    debe: float,
    haber: float,
    rol: str = "",
) -> dict:
    return {
        "Cuenta": cod,
        "Detalle": desc,
        "Debe": round(debe, 2),
        "Haber": round(haber, 2),
        "Estado": "Ingresado",
        "_rol": rol,
    }


def _generar_asiento_iva_excel_robusto(
    file_buffer,
    plan_cuentas_df: pd.DataFrame,
    sal_tec: float,
    sal_lib: float,
    ret_mes: float,
    perc_mes: float,
    *,
    es_csv: bool = False,
) -> list[dict]:
    """Procesa Excel/CSV ARCA por recorrido fila a fila (sin slices frágiles)."""
    if es_csv:
        df = pd.read_csv(file_buffer, header=None)
    else:
        df = pd.read_excel(file_buffer, header=None)
    montos = _parsear_montos_fiscales_arca_desde_df(df)
    lineas, _, _, _ = _generar_asiento_iva_hibrido_excel(
        plan_cuentas_df, montos, sal_tec, sal_lib, ret_mes, perc_mes,
    )
    return lineas


def _aplicar_loop_review_por_rol(
    lineas_asiento: list[dict],
    rol_ajuste: str,
) -> tuple[list[dict], float, bool]:
    """Ajuste autónomo ≤ $5 sobre la línea del rol indicado en la ficha activa."""
    total_debe = sum(f["Debe"] for f in lineas_asiento)
    total_haber = sum(f["Haber"] for f in lineas_asiento)
    diferencia = round(total_debe - total_haber, 2)
    dif_antes_ajuste = diferencia
    ajuste_aplicado = False
    if diferencia != 0.0 and abs(diferencia) <= 5.0:
        candidatos = [
            f for f in lineas_asiento
            if f.get("_rol") == rol_ajuste and f["Debe"] > 0
        ]
        if candidatos:
            objetivo = max(candidatos, key=lambda f: f["Debe"])
            objetivo["Debe"] = round(objetivo["Debe"] - diferencia, 2)
            ajuste_aplicado = True
    return lineas_asiento, dif_antes_ajuste, ajuste_aplicado


def _aplicar_loop_review_ventas_21(
    lineas_asiento: list[dict],
) -> tuple[list[dict], float, bool]:
    """Ajuste autónomo ≤ $5 sobre la línea de IVA Débito Fiscal 21%."""
    total_debe = sum(f["Debe"] for f in lineas_asiento)
    total_haber = sum(f["Haber"] for f in lineas_asiento)
    diferencia = round(total_debe - total_haber, 2)
    dif_antes_ajuste = diferencia
    ajuste_aplicado = False
    if diferencia != 0.0 and abs(diferencia) <= 5.0:
        candidatos = [
            f for f in lineas_asiento
            if f.get("_rol") == "ventas_21" and f["Debe"] > 0
        ]
        if not candidatos:
            candidatos = [
                f for f in lineas_asiento
                if f["Debe"] > 0
                and _matchea_alicuota_desc(f["Detalle"], "21")
                and _es_linea_ventas_iva(f["Detalle"])
            ]
        if candidatos:
            objetivo = max(candidatos, key=lambda f: f["Debe"])
            objetivo["Debe"] = round(objetivo["Debe"] - diferencia, 2)
            ajuste_aplicado = True
    return lineas_asiento, dif_antes_ajuste, ajuste_aplicado


def _aplicar_loop_review_impuesto_determinado(
    lineas_asiento: list[dict],
) -> tuple[list[dict], float, bool]:
    """Ajuste autónomo ≤ $5 sobre la línea de IIBB impuesto determinado (Debe)."""
    total_debe = sum(f["Debe"] for f in lineas_asiento)
    total_haber = sum(f["Haber"] for f in lineas_asiento)
    diferencia = round(total_debe - total_haber, 2)
    dif_antes_ajuste = diferencia
    ajuste_aplicado = False
    if diferencia != 0.0 and abs(diferencia) <= 5.0:
        candidatos = [
            f for f in lineas_asiento
            if f.get("_rol") == "impuesto_determinado" and f["Debe"] > 0
        ]
        if candidatos:
            objetivo = max(candidatos, key=lambda f: f["Debe"])
            objetivo["Debe"] = round(objetivo["Debe"] - diferencia, 2)
            ajuste_aplicado = True
    return lineas_asiento, dif_antes_ajuste, ajuste_aplicado


def _generar_asiento_iva_hibrido_excel(
    plan_cuentas_df: pd.DataFrame,
    montos_excel: dict[str, float],
    saldo_tecnico: float | list[float],
    saldo_libre: float | list[float],
    retenciones: float,
    percepciones: float,
) -> tuple[list[dict], dict, float, bool]:
    """Consolida asiento mensual IVA: Excel alícuotas + cargas manuales + balance DDJJ."""
    saldos_tec = _normalizar_lista_saldos_iva(saldo_tecnico)
    saldos_lib = _normalizar_lista_saldos_iva(saldo_libre)
    resumen = _calcular_desglose_posicion_iva(
        df_debito_21=montos_excel.get("ventas_21", 0.0),
        df_debito_105=montos_excel.get("ventas_105", 0.0),
        df_debito_27=montos_excel.get("ventas_27", 0.0),
        nc_compras=montos_excel.get("nc_compras", 0.0),
        nc_compras_27=montos_excel.get("nc_compras_27", 0.0),
        cf_credito_21=montos_excel.get("compras_21", 0.0),
        cf_credito_105=montos_excel.get("compras_105", 0.0),
        cf_credito_27=montos_excel.get("compras_27", 0.0),
        nc_ventas=montos_excel.get("nc_ventas", 0.0),
        nc_ventas_27=montos_excel.get("nc_ventas_27", 0.0),
        retenciones=retenciones,
        percepciones=percepciones,
        saldos_tecnicos=saldos_tec,
        saldos_libre=saldos_lib,
    )
    roles_debe = (
        ("ventas_21", montos_excel.get("ventas_21", 0.0)),
        ("ventas_105", montos_excel.get("ventas_105", 0.0)),
        ("ventas_27", montos_excel.get("ventas_27", 0.0)),
        ("nc_compras", montos_excel.get("nc_compras", 0.0)),
        ("nc_compras_27", montos_excel.get("nc_compras_27", 0.0)),
    )
    roles_haber = (
        ("compras_21", montos_excel.get("compras_21", 0.0)),
        ("compras_105", montos_excel.get("compras_105", 0.0)),
        ("compras_27", montos_excel.get("compras_27", 0.0)),
        ("nc_ventas", montos_excel.get("nc_ventas", 0.0)),
        ("nc_ventas_27", montos_excel.get("nc_ventas_27", 0.0)),
        ("retenciones", retenciones),
        ("percepciones", percepciones),
    )
    lineas = _armar_lineas_movimiento_iva(plan_cuentas_df, resumen, roles_debe, roles_haber)
    lineas, dif_ajuste, ajustado = _aplicar_loop_review_ventas_21(lineas)
    return lineas, resumen, dif_ajuste, ajustado


def _celda_a_float(val) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return 0.0
    try:
        return float(s)
    except ValueError:
        pass
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return 0.0


def _normalizar_concepto_planilla(val) -> str:
    """Columna A: minúsculas, sin acentos, espacios colapsados."""
    if pd.isna(val):
        return ""
    t = _norm_desc_iva(str(val))
    return re.sub(r"\s+", " ", t).strip()


def _plantilla_datos_planilla_iva() -> dict:
    return {
        "df_debito_21": 0.0,
        "df_debito_105": 0.0,
        "df_debito_27": 0.0,
        "cf_credito_21": 0.0,
        "cf_credito_105": 0.0,
        "cf_credito_27": 0.0,
        "nc_ventas": 0.0,
        "nc_compras": 0.0,
        "nc_ventas_27": 0.0,
        "nc_compras_27": 0.0,
        "retenciones_planilla": 0.0,
        "percepciones_planilla": 0.0,
        "saldo_tecnico_planilla": 0.0,
        "saldo_libre_planilla": 0.0,
        "periodo_texto": "",
        "periodo_mes": None,
        "periodo_anio": None,
    }


def _matcher_etiqueta_planilla_iva(concepto: str) -> str | None:
    """Resuelve la clave de datos según palabras clave en columna A (orden = especificidad)."""
    t = concepto
    if not t:
        return None

    tiene = lambda *tokens: all(tok in t for tok in tokens)
    es_105 = "10.5" in t or "10,5" in t or bool(re.search(r"10\s*[,.]\s*5", t))
    es_21 = bool(re.search(r"(?<!\d)21(?!\d)", t))
    es_27 = bool(re.search(r"(?<!\d)27(?!\d)", t))

    if es_27 and (("nc" in t and "compr" in t) or (tiene("nota", "credito") and "compr" in t)):
        return "nc_compras_27"
    if es_27 and (("nc" in t and "vent" in t) or (tiene("nota", "credito") and "vent" in t)):
        return "nc_ventas_27"
    if ("nc" in t or tiene("nota", "credito")) and "compr" in t:
        return "nc_compras"
    if ("nc" in t or tiene("nota", "credito")) and "vent" in t:
        return "nc_ventas"

    if es_27 and ("debito" in t or "df" in t.split()):
        return "df_debito_27"
    if es_105 and ("debito" in t or "df" in t.split()):
        return "df_debito_105"
    if es_21 and ("debito" in t or "df" in t.split()):
        return "df_debito_21"
    if "debito fiscal" in t or t.startswith("debito"):
        if es_27:
            return "df_debito_27"
        if es_105:
            return "df_debito_105"
        if es_21:
            return "df_debito_21"

    if es_27 and ("credito" in t or "cf" in t.split()):
        return "cf_credito_27"
    if es_105 and ("credito" in t or "cf" in t.split()):
        return "cf_credito_105"
    if es_21 and ("credito" in t or "cf" in t.split()):
        return "cf_credito_21"
    if "credito fiscal" in t or t.startswith("credito"):
        if es_27:
            return "cf_credito_27"
        if es_105:
            return "cf_credito_105"
        if es_21:
            return "cf_credito_21"

    if "retencion" in t:
        return "retenciones_planilla"
    if "percepcion" in t or "percep" in t:
        return "percepciones_planilla"
    if "saldo" in t and "tecnico" in t:
        return "saldo_tecnico_planilla"
    if "libre" in t and ("disponibilidad" in t or "disp" in t):
        return "saldo_libre_planilla"

    return None


def _columna_a_planilla(df: pd.DataFrame) -> pd.Series:
    """Columna A normalizada: vacíos seguros, minúsculas, sin espacios extra."""
    if df.shape[1] == 0:
        return pd.Series([""] * len(df), dtype=str)
    return df[0].fillna("").astype(str).str.lower().str.strip()


def _extraer_periodo_planilla_df(
    df: pd.DataFrame,
    periodo_mensual: str | None = None,
) -> tuple[str, int | None, int | None]:
    """Busca fila de período por etiqueta, columna seleccionada o patrón MM-AAAA."""
    if periodo_mensual:
        periodo = _parsear_periodo_texto(periodo_mensual.replace("/", "-"))
        if periodo:
            mes, anio = periodo
            return periodo_mensual.replace("-", "/"), mes, anio
    if df.shape[1] == 0:
        return "", None, None

    col_a = _columna_a_planilla(df)
    col_b = df[1] if df.shape[1] > 1 else df[0]

    for concepto_raw, val in zip(col_a, col_b):
        concepto = _normalizar_concepto_planilla(concepto_raw)
        if "periodo" in concepto or concepto in ("mes", "mes/año", "mes/anio"):
            periodo = _parsear_celda_periodo_planilla(val)
            return _texto_celda_periodo(val), (
                periodo[0] if periodo else None
            ), (periodo[1] if periodo else None)

    if df.shape[1] > 1:
        for val in df[1]:
            periodo = _parsear_celda_periodo_planilla(val)
            if periodo:
                return _texto_celda_periodo(val), periodo[0], periodo[1]
    return "", None, None


def _columna_montos_planilla(df: pd.DataFrame, periodo_mensual: str | None = None) -> pd.Series:
    """Columna de montos: índice 1 por defecto o columna del período mensual."""
    if df.shape[1] == 0:
        return pd.Series(dtype=float)
    if periodo_mensual:
        idx = resolver_indice_columna_periodo(df, periodo_mensual)
        if idx is not None:
            return df[idx]
    return df[1] if df.shape[1] > 1 else df[0]


def leer_planilla_iva_por_etiquetas(
    file_buffer,
    *,
    es_csv: bool = False,
    nombre_solapa_impuesto: str = "IVA",
    periodo_mensual: str | None = None,
) -> dict:
    """Lee solapa del balance por coincidencia de etiquetas en columna A (estructura flexible)."""
    if es_csv:
        df_mes = pd.read_csv(file_buffer, header=None)
    else:
        if hasattr(file_buffer, "seek"):
            file_buffer.seek(0)
        df_mes = leer_dataframe_balance_solapa(
            file_buffer,
            nombre_solapa_impuesto,
            es_csv=False,
            header=None,
        )

    datos = _plantilla_datos_planilla_iva()
    asignados: set[str] = set()

    if df_mes.shape[1] > 0:
        col_a = _columna_a_planilla(df_mes)
        col_montos = _columna_montos_planilla(df_mes, periodo_mensual)
        for concepto_raw, monto_raw in zip(col_a, col_montos):
            concepto = _normalizar_concepto_planilla(concepto_raw)
            if not concepto or "periodo" in concepto:
                continue
            clave = _matcher_etiqueta_planilla_iva(concepto)
            if clave is None or clave in asignados:
                continue
            monto = _celda_a_float(monto_raw)
            if monto != 0.0:
                datos[clave] = round(float(datos.get(clave, 0.0)) + monto, 2)
                asignados.add(clave)

    texto_per, mes, anio = _extraer_periodo_planilla_df(df_mes, periodo_mensual)
    datos["periodo_texto"] = texto_per
    datos["periodo_mes"] = mes
    datos["periodo_anio"] = anio
    return datos


def leer_planilla_iva_posicional(
    file_buffer,
    *,
    es_csv: bool = False,
    nombre_solapa_impuesto: str = "IVA",
    periodo_mensual: str | None = None,
) -> dict:
    """Alias retrocompatible: lectura flexible por etiquetas en solapa del balance."""
    return leer_planilla_iva_por_etiquetas(
        file_buffer,
        es_csv=es_csv,
        nombre_solapa_impuesto=nombre_solapa_impuesto,
        periodo_mensual=periodo_mensual,
    )


def _plantilla_datos_planilla_iibb() -> dict:
    return {
        "impuesto_determinado": 0.0,
        "retenciones_planilla": 0.0,
        "percepciones_planilla": 0.0,
        "retenciones_bancarias_planilla": 0.0,
        "saldo_favor_planilla": 0.0,
        "periodo_texto": "",
        "periodo_mes": None,
        "periodo_anio": None,
    }


def _matcher_etiqueta_planilla_iibb(concepto: str) -> str | None:
    """Resuelve clave IIBB según palabras clave en columna A (orden = especificidad)."""
    t = concepto
    if not t:
        return None

    if any(x in t for x in ("retenciones bancarias", "sircreb", "bancarias iibb")):
        return "retenciones_bancarias_planilla"
    if "retencion" in t and ("bancaria" in t or "sircreb" in t):
        return "retenciones_bancarias_planilla"

    if "retenciones iibb" in t or (
        "retencion" in t and "iibb" in t and "bancaria" not in t and "sircreb" not in t
    ):
        return "retenciones_planilla"
    if t.strip() == "retenciones" or (
        "retencion" in t and "bancaria" not in t and "sircreb" not in t and "iva" not in t
    ):
        return "retenciones_planilla"

    if "percepciones iibb" in t or (
        ("percepcion" in t or "percep" in t) and "iibb" in t
    ):
        return "percepciones_planilla"
    if t.strip() == "percepciones" or (
        ("percepcion" in t or "percep" in t) and "iva" not in t and "bancaria" not in t
    ):
        return "percepciones_planilla"

    if "saldo a favor anterior" in t or "saldo anterior iibb" in t:
        return "saldo_favor_planilla"

    if (
        "impuesto determinado" in t
        or "iibb devengado" in t
        or "monto de impuesto" in t
    ):
        return "impuesto_determinado"

    return None


def leer_planilla_iibb_por_etiquetas(
    file_buffer,
    *,
    es_csv: bool = False,
    nombre_solapa_impuesto: str = "Ingresos Brutos",
    periodo_mensual: str | None = None,
) -> dict:
    """Lee solapa IIBB del balance por coincidencia de etiquetas en columna A."""
    if es_csv:
        df_mes = pd.read_csv(file_buffer, header=None)
    else:
        if hasattr(file_buffer, "seek"):
            file_buffer.seek(0)
        df_mes = leer_dataframe_balance_solapa(
            file_buffer,
            nombre_solapa_impuesto,
            es_csv=False,
            header=None,
        )

    datos = _plantilla_datos_planilla_iibb()
    asignados: set[str] = set()

    if df_mes.shape[1] > 0:
        col_a = _columna_a_planilla(df_mes)
        col_montos = _columna_montos_planilla(df_mes, periodo_mensual)
        for concepto_raw, monto_raw in zip(col_a, col_montos):
            concepto = _normalizar_concepto_planilla(concepto_raw)
            if not concepto or "periodo" in concepto:
                continue
            clave = _matcher_etiqueta_planilla_iibb(concepto)
            if clave is None or clave in asignados:
                continue
            monto = _celda_a_float(monto_raw)
            if monto != 0.0:
                datos[clave] = round(float(datos.get(clave, 0.0)) + monto, 2)
                asignados.add(clave)

    texto_per, mes, anio = _extraer_periodo_planilla_df(df_mes, periodo_mensual)
    datos["periodo_texto"] = texto_per
    datos["periodo_mes"] = mes
    datos["periodo_anio"] = anio
    return datos


def _texto_celda_periodo(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime, date)):
        return f"{val.month:02d}-{val.year}"
    return str(val).strip()


def _parsear_celda_periodo_planilla(val) -> tuple[int, int] | None:
    """Extrae mes y año desde celda de planilla (fecha nativa o mes en texto español)."""
    parsed = _parsear_periodo_balance_celda_proc(val)
    if parsed:
        return parsed
    if pd.isna(val):
        return None
    return _parsear_periodo_texto(str(val).strip())


def _parsear_periodo_texto(periodo_texto: str) -> tuple[int, int] | None:
    """Parsea MM-YYYY, MM/YYYY, MM-YY o fechas ISO exportadas por Excel."""
    s = str(periodo_texto).strip()
    if not s or s.lower() == "nan":
        return None

    m = re.match(r"^(0?[1-9]|1[0-2])[/\-](\d{4})$", s)
    if m:
        return int(m.group(1)), int(m.group(2))

    m = re.match(r"^(0?[1-9]|1[0-2])[/\-](\d{2})$", s)
    if m:
        yy = int(m.group(2))
        anio = 2000 + yy if yy < 50 else 1900 + yy
        return int(m.group(1)), anio

    m = re.match(r"^(\d{4})[/\-](0?[1-9]|1[0-2])(?:[/\-]\d{1,2})?", s)
    if m:
        return int(m.group(2)), int(m.group(1))

    return _extraer_periodo(s)


def _fecha_asiento_iva_tango(mes: int, anio: int) -> date:
    """Último día calendario del mes/año indicados en la planilla (formato Tango)."""
    ultimo_dia = calendar.monthrange(anio, mes)[1]
    return date(anio, mes, ultimo_dia)


def _asegurar_fecha_asiento_ultimo_dia(
    fecha_key: str,
    periodo_mensual: str | None,
) -> None:
    """Inicializa el date_input con el último día del período MM/YYYY.

    Cada período usa su propia widget key (ver _clave_fecha_tango_*), así que
    solo se asigna si la key aún no existe: no se muta un date_input ya montado
    (eso disparaba NotFoundError removeChild en Streamlit).
    """
    if not periodo_mensual or fecha_key in st.session_state:
        return
    parsed = _parsear_periodo_texto(str(periodo_mensual).replace("/", "-"))
    if not parsed:
        return
    st.session_state[fecha_key] = _fecha_asiento_iva_tango(int(parsed[0]), int(parsed[1]))


def _formatear_fecha_tango(fecha: date) -> str:
    return formatear_fecha_dd_mm_yyyy(fecha)


def _normalizar_celdas_fecha_grilla(row: dict) -> dict:
    """Período MM/YYYY y Fecha DD/MM/YYYY como strings independientes."""
    fila = dict(row)
    fila["Período"] = formatear_periodo_mm_yyyy(fila.get("Período", ""))
    fila["Fecha"] = formatear_fecha_dd_mm_yyyy(fila.get("Fecha", ""))
    return fila


def _normalizar_filas_grilla(rows: list[dict]) -> list[dict]:
    return [_normalizar_celdas_fecha_grilla(r) for r in rows]


def _texto_periodo_grilla(val) -> str:
    return formatear_periodo_mm_yyyy(val)


def _texto_fecha_grilla(val) -> str:
    return formatear_fecha_dd_mm_yyyy(val)


def _impuesto_activo_devengamientos() -> str:
    """Impuesto seleccionado en pantalla → nombre de solapa del balance."""
    try:
        if _IMPUESTO_KEY in st.session_state:
            val = str(st.session_state[_IMPUESTO_KEY]).strip()
            return val or "IVA"
    except Exception:
        pass
    return "IVA"


def _fecha_default_desde_planilla_iibb(
    archivo,
    nombre_solapa_impuesto: str | None = None,
) -> date | None:
    if archivo is None:
        return None
    solapa = nombre_solapa_impuesto or _impuesto_activo_devengamientos()
    try:
        buf, es_csv = _abrir_planilla_iva(archivo)
        datos = leer_planilla_iibb_por_etiquetas(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        )
        mes, anio = datos.get("periodo_mes"), datos.get("periodo_anio")
        if mes is not None and anio is not None:
            return _fecha_asiento_iva_tango(int(mes), int(anio))
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if periodo:
            return _fecha_asiento_iva_tango(periodo[0], periodo[1])
    except Exception:
        return None
    return None


def _fecha_default_desde_planilla(
    archivo,
    nombre_solapa_impuesto: str | None = None,
) -> date | None:
    """Último día del mes/año leído en la solapa del balance del impuesto activo."""
    if archivo is None:
        return None
    solapa = nombre_solapa_impuesto or _impuesto_activo_devengamientos()
    try:
        buf, es_csv = _abrir_planilla_iva(archivo)
        datos = leer_planilla_iva_posicional(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        )
        mes, anio = datos.get("periodo_mes"), datos.get("periodo_anio")
        if mes is not None and anio is not None:
            return _fecha_asiento_iva_tango(int(mes), int(anio))
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if periodo:
            return _fecha_asiento_iva_tango(periodo[0], periodo[1])
    except Exception:
        return None
    return None


def _clave_referencia_planilla(archivo, sociedad_id: int | None = None) -> str:
    if archivo is None:
        return ""
    servidor = _buffer_balance_servidor_sociedad(sociedad_id)
    if servidor is not None:
        if servidor is archivo or (
            hasattr(archivo, "getvalue")
            and hasattr(servidor, "getvalue")
            and archivo.getvalue() == servidor.getvalue()
        ):
            return f"servidor:{_ruta_balance_servidor_sociedad(sociedad_id)}"
    if hasattr(archivo, "name"):
        return str(archivo.name)
    return "planilla_memoria"


def _prefijar_fecha_input_desde_planilla_iibb(archivo, sociedad_id: int | None = None) -> None:
    imp = _impuesto_activo_devengamientos()
    slug = _slug_impuesto(imp)
    ref = _clave_referencia_planilla(archivo, sociedad_id)
    rc = _iva_reset_counter()
    fecha_key = f"{slug}_fecha_tango_{rc}"
    ref_key = f"{slug}_planilla_ref_fecha"
    if ref and st.session_state.get(ref_key) != ref:
        sugerida = _fecha_default_desde_planilla_iibb(archivo, nombre_solapa_impuesto=imp)
        if sugerida:
            st.session_state[fecha_key] = sugerida
        st.session_state[ref_key] = ref
    elif fecha_key not in st.session_state:
        sugerida = _fecha_default_desde_planilla_iibb(archivo, nombre_solapa_impuesto=imp)
        st.session_state[fecha_key] = sugerida or date.today()


def _prefijar_fecha_input_desde_planilla_banco(
    archivo,
    slug: str,
    sociedad_id: int | None = None,
) -> None:
    """Pre-puebla el date_input bancario v2 al cambiar de planilla."""
    ref = _clave_referencia_planilla(archivo, sociedad_id)
    rc = _iva_reset_counter()
    fecha_key = _clave_fecha_tango_banco(slug, rc)
    ref_key = f"{slug}_planilla_ref_fecha"
    if ref and st.session_state.get(ref_key) != ref:
        sugerida = _fecha_default_desde_planilla(archivo)
        if sugerida:
            st.session_state[fecha_key] = sugerida
        st.session_state[ref_key] = ref
    elif fecha_key not in st.session_state:
        sugerida = _fecha_default_desde_planilla(archivo)
        st.session_state[fecha_key] = sugerida or date.today()


def _prefijar_fecha_input_desde_planilla(archivo, sociedad_id: int | None = None) -> None:
    """Pre-puebla el date_input al cambiar de planilla; no pisa la elección manual."""
    ref = _clave_referencia_planilla(archivo, sociedad_id)
    rc = _iva_reset_counter()
    fecha_key = f"iva_fecha_tango_{rc}"
    if ref and st.session_state.get("iva_planilla_ref_fecha") != ref:
        sugerida = _fecha_default_desde_planilla(archivo)
        if sugerida:
            st.session_state[fecha_key] = sugerida
        st.session_state["iva_planilla_ref_fecha"] = ref
    elif fecha_key not in st.session_state:
        sugerida = _fecha_default_desde_planilla(archivo)
        st.session_state[fecha_key] = sugerida or date.today()


def _aplicar_fecha_tango_asientos(
    asientos: list[AsientoDevengamiento],
    rows: list[dict],
    fecha: date,
) -> tuple[list[AsientoDevengamiento], list[dict]]:
    """Fuente de verdad: fecha manual → string DD/MM/YYYY en grilla y export Tango."""
    fecha_final_str = formatear_fecha_dd_mm_yyyy(fecha)
    for asiento in asientos:
        asiento.fecha = fecha
        asiento.fecha_tango_str = fecha_final_str  # type: ignore[attr-defined]
    for row in rows:
        row["Fecha"] = fecha_final_str
    return asientos, rows


def _inicializar_sesion_iva() -> None:
    """Estado unificado del módulo Devengamientos IVA."""
    if "biblioteca_asientos" not in st.session_state:
        st.session_state.biblioteca_asientos = []
    if "periodos_procesados" not in st.session_state:
        st.session_state.periodos_procesados = []
    if "reset_counter" not in st.session_state:
        st.session_state.reset_counter = 0
    if "saldos_tecnicos_list" not in st.session_state:
        st.session_state.saldos_tecnicos_list = [0.0]
    if "saldos_libre_list" not in st.session_state:
        st.session_state.saldos_libre_list = [0.0]
    _inicializar_balance_servidor_por_sociedad()


def _iva_reset_counter() -> int:
    return int(st.session_state.get("reset_counter", 0))


def _fecha_asiento_seleccionada_iibb() -> date | None:
    rc = _iva_reset_counter()
    val = st.session_state.get(f"iibb_fecha_tango_{rc}")
    return val if isinstance(val, date) else None


def _fecha_asiento_seleccionada() -> date | None:
    rc = _iva_reset_counter()
    val = st.session_state.get(f"iva_fecha_tango_{rc}")
    return val if isinstance(val, date) else None


def _periodos_procesados_sociedad(
    sociedad_id: int | None,
    impuesto: str | None = None,
) -> list[str]:
    if sociedad_id is None:
        return []
    resultado: list[str] = []
    for p in (st.session_state.get("periodos_procesados") or []):
        if p.get("sociedad_id") != sociedad_id:
            continue
        if impuesto is not None and p.get("impuesto", "IVA") != impuesto:
            continue
        resultado.append(_periodo_a_etiqueta(p.get("periodo", "")))
    return resultado


def _periodo_ya_en_biblioteca(
    sociedad_id: int,
    periodo: str,
    impuesto: str = "IVA",
) -> bool:
    etiqueta = _periodo_a_etiqueta(periodo)
    return etiqueta in _periodos_procesados_sociedad(sociedad_id, impuesto)


def _biblioteca_por_sociedad(
    sociedad_id: int | None,
    impuesto: str | None = None,
) -> list[dict]:
    if sociedad_id is None:
        return []
    entradas = [
        b for b in (st.session_state.get("biblioteca_asientos") or [])
        if b.get("sociedad_id") == sociedad_id
    ]
    if impuesto is not None:
        entradas = [b for b in entradas if b.get("impuesto", "IVA") == impuesto]
    return entradas


def _vaciar_biblioteca_sociedad(
    sociedad_id: int | None,
    impuesto: str | None = None,
) -> None:
    if sociedad_id is None:
        return
    if impuesto is None:
        st.session_state.biblioteca_asientos = [
            b for b in (st.session_state.get("biblioteca_asientos") or [])
            if b.get("sociedad_id") != sociedad_id
        ]
        st.session_state.periodos_procesados = [
            p for p in (st.session_state.get("periodos_procesados") or [])
            if p.get("sociedad_id") != sociedad_id
        ]
        _persistir_biblioteca_en_disco()
        return
    st.session_state.biblioteca_asientos = [
        b for b in (st.session_state.get("biblioteca_asientos") or [])
        if not (b.get("sociedad_id") == sociedad_id and b.get("impuesto", "IVA") == impuesto)
    ]
    st.session_state.periodos_procesados = [
        p for p in (st.session_state.get("periodos_procesados") or [])
        if not (p.get("sociedad_id") == sociedad_id and p.get("impuesto", "IVA") == impuesto)
    ]
    _persistir_biblioteca_en_disco()


def _render_panel_biblioteca_asientos(
    sociedad_id: int | None,
    impuesto: str = "IVA",
) -> None:
    """Panel lateral: biblioteca acumulativa y períodos archivados por impuesto."""
    if st.session_state.get("_sidebar_unificada"):
        return
    _inicializar_sesion_iva()
    entradas = _biblioteca_por_sociedad(sociedad_id, impuesto)
    periodos = _periodos_procesados_sociedad(sociedad_id, impuesto)
    titulo = f"### Biblioteca de Asientos {impuesto}"
    btn_key = f"btn_vaciar_biblioteca_{_slug_impuesto(impuesto)}"

    with st.sidebar:
        st.markdown("---")
        st.markdown(titulo)
        st.metric("Meses archivados", len(entradas))
        if periodos:
            st.caption(f"Asientos listos: **[{', '.join(periodos)}]**")
        else:
            st.caption("Sin meses guardados aún.")
        st.caption(f"Archivo físico: `{ruta_biblioteca_persistida_activa(usuario=_usuario_oficina_actual())}`")
        if entradas and st.button(
            "Vaciar Biblioteca",
            key=btn_key,
            type="secondary",
            use_container_width=True,
        ):
            _vaciar_biblioteca_sociedad(sociedad_id, impuesto)
            st.rerun()


def _inicializar_sesion_iibb() -> None:
    if "saldos_favor_iibb_list" not in st.session_state:
        st.session_state.saldos_favor_iibb_list = [0.0]
    _inicializar_balance_servidor_por_sociedad()


def _inicializar_sesion_cm() -> None:
    if "saldos_favor_cm_list" not in st.session_state:
        st.session_state.saldos_favor_cm_list = [0.0]
    _inicializar_balance_servidor_por_sociedad()


def _inicializar_sesion_sueldos() -> None:
    if "saldos_favor_sueldos_list" not in st.session_state:
        st.session_state.saldos_favor_sueldos_list = [0.0]
    _inicializar_balance_servidor_por_sociedad()


def _inicializar_sesion_tish() -> None:
    if "saldos_favor_tish_list" not in st.session_state:
        st.session_state.saldos_favor_tish_list = [0.0]
    _inicializar_balance_servidor_por_sociedad()


def _aplicar_limpieza_formulario_mes_iibb_si_pendiente() -> None:
    if not st.session_state.pop("iibb_limpiar_formulario_pendiente", False):
        return
    st.session_state.pop("iibb_grilla_preview", None)
    st.session_state.pop("iibb_asientos_generados", None)
    st.session_state.pop("iibb_planilla_ref_fecha", None)
    st.session_state.pop("iibb_resumen_analitico", None)
    st.session_state.pop("iibb_auto_fp", None)
    st.session_state["iibb_skip_default_planilla"] = True
    st.session_state["reset_counter"] = _iva_reset_counter() + 1


def _marcar_limpieza_formulario_mes_iibb() -> None:
    st.session_state["iibb_limpiar_formulario_pendiente"] = True


def _guardar_asiento_en_biblioteca(
    sociedad_id: int,
    asientos: list[AsientoDevengamiento],
    rows: list[dict],
    *,
    impuesto: str = "IVA",
) -> str:
    """Archiva el mes actual en la biblioteca (bloqueo anti-duplicado por impuesto)."""
    _inicializar_sesion_iva()
    if not asientos or not rows:
        raise ValueError("No hay asiento generado para guardar.")

    periodo = getattr(asientos[0], "periodo", "") or rows[0].get("Período", "")
    if not periodo:
        raise ValueError("No se pudo determinar el período del asiento.")

    etiqueta = _periodo_a_etiqueta(periodo)
    if _periodo_ya_en_biblioteca(sociedad_id, periodo, impuesto):
        raise ValueError(
            f"El período {etiqueta} ya está archivado en la biblioteca de {impuesto}. "
            "Vacíe la biblioteca o cargue otro mes."
        )

    for row in rows:
        row["Estado"] = "Ingresado"

    resumen_key = f"{_slug_impuesto(impuesto)}_resumen_analitico"
    entrada = {
        "sociedad_id": sociedad_id,
        "impuesto": impuesto,
        "periodo": periodo,
        "periodo_orden": _periodo_a_orden(periodo),
        "asientos": copy.deepcopy(asientos),
        "rows": copy.deepcopy(rows),
        "resumen_analitico": copy.deepcopy(st.session_state.get(resumen_key)),
    }
    st.session_state.biblioteca_asientos.append(entrada)
    st.session_state.periodos_procesados.append({
        "sociedad_id": sociedad_id,
        "impuesto": impuesto,
        "periodo": periodo,
    })
    _persistir_biblioteca_en_disco()
    _limpiar_borrador_si_corresponde(_slug_impuesto(impuesto))
    return etiqueta


def _periodo_a_orden(periodo: str) -> tuple[int, int]:
    """Convierte '04/2025' o '04-2025' en tupla (año, mes) para orden cronológico."""
    s = str(periodo).strip().replace("-", "/")
    partes = s.split("/")
    if len(partes) != 2:
        return (9999, 99)
    a, b = partes[0].strip(), partes[1].strip()
    if len(b) == 4:
        return int(b), int(a)
    if len(a) == 4:
        return int(a), int(b)
    return (9999, 99)


def _periodo_a_etiqueta(periodo: str) -> str:
    return str(periodo).strip().replace("/", "-")


def _inicializar_biblioteca_asientos() -> None:
    _inicializar_sesion_iva()


def _aplicar_limpieza_formulario_mes_iva_si_pendiente() -> None:
    """Reset seguro: solo incrementa reset_counter (anti StreamlitAPIException)."""
    if not st.session_state.pop("iva_limpiar_formulario_pendiente", False):
        return
    st.session_state.pop("iva_grilla_preview", None)
    st.session_state.pop("iva_asientos_generados", None)
    st.session_state.pop("iva_planilla_ref_fecha", None)
    st.session_state.pop("iva_resumen_analitico", None)
    st.session_state.pop("iva_auto_fp", None)
    st.session_state["iva_skip_default_planilla"] = True
    st.session_state["reset_counter"] = _iva_reset_counter() + 1


def _marcar_limpieza_formulario_mes_iva() -> None:
    """Programa el reset del formulario para el próximo rerun (post-guardar biblioteca)."""
    st.session_state["iva_limpiar_formulario_pendiente"] = True


def _asientos_consolidados_biblioteca(
    sociedad_id: int | None,
    impuesto: str | None = None,
) -> list[AsientoDevengamiento]:
    """Asigna Asiento 1, 2, 3… correlativo por mes archivado (anti-duplicación Tango)."""
    entradas = sorted(
        _biblioteca_por_sociedad(sociedad_id, impuesto),
        key=lambda e: e.get("periodo_orden", (9999, 99)),
    )
    todos: list[AsientoDevengamiento] = []
    for num_asiento, entrada in enumerate(entradas, start=1):
        for asiento in entrada.get("asientos", []):
            asiento.identificador = num_asiento
            todos.append(asiento)
    return todos


def _dataframe_consolidado_biblioteca(
    sociedad_id: int | None,
    impuesto: str | None = None,
) -> pd.DataFrame:
    entradas = sorted(
        _biblioteca_por_sociedad(sociedad_id, impuesto),
        key=lambda e: e.get("periodo_orden", (9999, 99)),
    )
    partes: list[pd.DataFrame] = []
    for entrada in entradas:
        filas = entrada.get("rows") or []
        if filas:
            partes.append(pd.DataFrame(filas))
    if not partes:
        return pd.DataFrame()
    df = pd.concat(partes, ignore_index=True)
    df["Estado"] = "Ingresado"
    subset = [c for c in ("Período", "Fecha", "Código", "Debe", "Haber") if c in df.columns]
    if subset:
        df = df.drop_duplicates(subset=subset, keep="last")
    cols = [c for c in ("Período", "Fecha", "Código", "Descripción", "Debe", "Haber", "Estado") if c in df.columns]
    return df[cols] if cols else df


def _generar_excel_biblioteca_consolidada(
    sociedad_id: int | None,
    cuit: str,
    nombre_cliente: str,
    impuesto: str | None = None,
    plan_cuentas: pd.DataFrame | None = None,
) -> Path:
    entradas = sorted(
        _biblioteca_por_sociedad(sociedad_id, impuesto),
        key=lambda e: e.get("periodo_orden", (9999, 99)),
    )
    if not entradas:
        raise ValueError("La biblioteca está vacía.")
    asientos = _asientos_consolidados_biblioteca(sociedad_id, impuesto)
    primero = entradas[0]["periodo_orden"]
    ultimo = entradas[-1]["periodo_orden"]
    mes_ini, anio_ini = primero[1], primero[0]
    mes_fin, anio_fin = ultimo[1], ultimo[0]
    etiqueta = (
        f"consolidado_{mes_ini:02d}_{anio_ini}_{mes_fin:02d}_{anio_fin}"
        if len(entradas) > 1
        else f"{mes_fin:02d}_{anio_fin}"
    )
    nombre_archivo = f"Devengamientos_Tango_{cuit}_{etiqueta}.xlsx"
    return generar_excel_tango_nativo(
        asientos=asientos,
        nombre_cliente=nombre_cliente,
        cuit=cuit,
        mes=mes_fin,
        anio=anio_fin,
        id_base=1,
        nombre_archivo=nombre_archivo,
        plan_cuentas=plan_cuentas,
    )


def _mostrar_errores_exportacion_tango(exc: ExportacionTangoError) -> None:
    st.error(str(exc))
    for err in exc.errores or []:
        periodo = err.get("periodo") or "—"
        st.warning(
            f"Asiento **#{err.get('identificador', '?')}** ({err.get('concepto', '')}) — "
            f"período {periodo}, renglón **{err.get('nro', '?')}**: "
            f"cuenta **{err.get('codigo', '')}** ({err.get('descripcion', '')}). "
            f"{err.get('motivo', '')}"
        )


def _mostrar_informe_exportacion_tango(informe: dict[str, list[dict]]) -> None:
    """Muestra resumen agrupado por cuenta (no 158 líneas repetidas)."""
    bloqueantes = informe.get("bloqueantes") or []
    advertencias = informe.get("advertencias") or []
    if bloqueantes:
        st.error(
            f"Exportación bloqueada: {len(bloqueantes)} renglón/es con cuentas "
            "sin asignar o Rubro/Madre no imputable."
        )
    if advertencias and not bloqueantes:
        st.warning(
            f"Atención Tango: {len(advertencias)} renglón/es usan cuentas con "
            "**auxiliares contables**. El Excel se generará, pero Tango solo los aceptará si "
            "cada cuenta tiene **tipo de auxiliar** y **regla de apropiación al 100%** configurados."
        )
    if not bloqueantes and not advertencias:
        return
    resumen = resumir_informe_export_tango(informe)
    filas = []
    for item in resumen:
        etiqueta = "Bloqueante" if item["tipo"] == "bloqueantes" else "Advertencia auxiliar"
        filas.append({
            "Tipo": etiqueta,
            "Código": item["codigo"],
            "Descripción": item["descripcion"],
            "Veces en biblioteca": item["ocurrencias"],
            "Asientos": ", ".join(item["asientos"][:12]) + (
                "…" if len(item["asientos"]) > 12 else ""
            ),
            "Detalle": item["motivo"],
        })
    with st.expander(
        f"Detalle por cuenta ({len(resumen)} cuenta/s únicas)",
        expanded=bool(bloqueantes),
    ):
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
        if advertencias and not bloqueantes:
            st.info(
                "**Configuración obligatoria en Tango (una sola vez por cuenta):**\n\n"
                "1. **Plan de cuentas** → abrir cada cuenta listada (ej. 42201, 42203, 21312, 21302).\n"
                "2. En **Auxiliares contables**, asignar un **tipo de auxiliar** (ej. Legajos para sueldos).\n"
                "3. Crear **regla de apropiación al 100%** para ese tipo.\n"
                "4. Alternativa: desactivar «Usa auxiliares contables» si el devengamiento es agregado.\n\n"
                "El Excel de importación no incluye apropiaciones analíticas; sin ese paso Tango rechaza "
                "con *«La cuenta no tiene asignado ningún tipo de auxiliar»*.\n\n"
                "Los asientos de sueldos se exportan con tipo **SUELDOS** (no VARIOS)."
            )


def _asegurar_plan_cuentas_export(
    sociedad_id: int,
    cuit_activo: str | None,
    *,
    slug: str,
) -> pd.DataFrame | None:
    """Carga el plan del cliente (o default) y ofrece subida inline si falta."""
    try:
        _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=True)
    except Exception:
        pass

    plan_df = st.session_state.get("plan_cuentas_df")
    if plan_df is not None and not getattr(plan_df, "empty", True):
        if bool(st.session_state.get("plan_cuentas_es_default", False)):
            st.info(
                "Usando el **plan de cuentas genérico** del estudio. "
                "Para este cliente conviene vincular su Excel propio "
                "(Gestión de Clientes o el cargador de abajo)."
            )
        else:
            st.caption(
                f"Plan cargado: {Path(st.session_state.get('plan_cuentas_path_resuelto') or '').name} "
                f"({len(plan_df)} cuentas)."
            )
        return plan_df

    st.error(
        "Exportación bloqueada: no hay plan de cuentas cargado. "
        "Subí el Excel de cuentas de este cliente para continuar."
    )
    archivo = st.file_uploader(
        "Subir Plan de Cuentas del cliente (.xlsx)",
        type=["xlsx"],
        key=f"uploader_plan_export_{slug}_{sociedad_id}",
        help="Se guarda para este cliente y habilita el export a Tango.",
    )
    if archivo and cuit_activo:
        try:
            xlsx_path = _guardar_plan_cliente_en_disco(str(cuit_activo), archivo.getvalue())
            cli = db.obtener_cliente(int(sociedad_id))
            if cli:
                db.actualizar_cliente(
                    cli["id"],
                    cli["nombre"],
                    cli["cuit"],
                    cli["tipo_persona"],
                    str(xlsx_path),
                    cli.get("mes_cierre_balance", 12),
                )
            _invalidar_cache_plan(int(sociedad_id))
            _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=True)
            st.success(f"Plan vinculado: {xlsx_path.name}")
            st.rerun()
        except Exception as exc:
            st.error(f"No se pudo guardar el plan: {exc}")
    return None


def _auditar_asientos_antes_export_ui(
    asientos: list,
    plan_cuentas: pd.DataFrame | None,
) -> bool:
    """Validación intermedia visible en pantalla antes de compilar el Excel Tango."""
    if plan_cuentas is None or plan_cuentas.empty:
        st.error(
            "Exportación bloqueada: no hay plan de cuentas cargado. "
            "Vinculá el Excel de cuentas del cliente para validar imputabilidad y auxiliares."
        )
        return False
    informe = auditar_exportacion_tango(asientos, plan_cuentas)
    _mostrar_informe_exportacion_tango(informe)
    if informe.get("bloqueantes"):
        return False
    try:
        preparar_asientos_export_tango(asientos, plan_cuentas)
    except ExportacionTangoError as exc:
        _mostrar_errores_exportacion_tango(exc)
        return False
    return True


def _periodos_bancos_procesados_sociedad(
    sociedad_id: int | None,
    banco: str | None = None,
) -> list[str]:
    if sociedad_id is None:
        return []
    resultado: list[str] = []
    for p in (st.session_state.get("periodos_bancos_procesados") or []):
        if p.get("sociedad_id") != sociedad_id:
            continue
        if banco is not None and p.get("banco") != banco:
            continue
        resultado.append(_periodo_a_etiqueta(p.get("periodo", "")))
    return resultado


def _biblioteca_bancos_por_sociedad(
    sociedad_id: int | None,
    banco: str | None = None,
) -> list[dict]:
    if sociedad_id is None:
        return []
    entradas = [
        b for b in (st.session_state.get("biblioteca_bancos") or [])
        if b.get("sociedad_id") == sociedad_id
    ]
    if banco is not None:
        entradas = [b for b in entradas if b.get("banco") == banco]
    return entradas


def _periodo_ya_en_biblioteca_banco(
    sociedad_id: int,
    periodo: str,
    banco: str,
) -> bool:
    etiqueta = _periodo_a_etiqueta(periodo)
    return etiqueta in _periodos_bancos_procesados_sociedad(sociedad_id, banco)


def _vaciar_biblioteca_banco_sociedad(
    sociedad_id: int | None,
    banco: str | None = None,
) -> None:
    if sociedad_id is None:
        return
    if banco is None:
        st.session_state.biblioteca_bancos = [
            b for b in (st.session_state.get("biblioteca_bancos") or [])
            if b.get("sociedad_id") != sociedad_id
        ]
        st.session_state.periodos_bancos_procesados = [
            p for p in (st.session_state.get("periodos_bancos_procesados") or [])
            if p.get("sociedad_id") != sociedad_id
        ]
        _persistir_biblioteca_en_disco()
        return
    st.session_state.biblioteca_bancos = [
        b for b in (st.session_state.get("biblioteca_bancos") or [])
        if not (b.get("sociedad_id") == sociedad_id and b.get("banco") == banco)
    ]
    st.session_state.periodos_bancos_procesados = [
        p for p in (st.session_state.get("periodos_bancos_procesados") or [])
        if not (p.get("sociedad_id") == sociedad_id and p.get("banco") == banco)
    ]
    _persistir_biblioteca_en_disco()


def _render_bloque_excel_consolidado_banco(
    sociedad_id: int | None,
    banco: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    *,
    key_suffix: str,
    asegurar_plan: bool = True,
    auditoria_ui: bool = True,
) -> None:
    """Descarga del Excel consolidado de la biblioteca bancaria (sidebar o panel)."""
    if sociedad_id is None:
        return
    entradas_bib = _biblioteca_bancos_por_sociedad(sociedad_id, banco)
    if not entradas_bib:
        return

    slug = _slug_banco(banco)
    st.markdown(f"#### 📦 Excel consolidado ({banco})")
    st.caption(
        f"{len(entradas_bib)} mes/es archivados listos para importar a Tango."
    )

    if asegurar_plan:
        plan_df = _asegurar_plan_cuentas_export(
            sociedad_id, cuit_activo, slug=f"banco_{slug}_{key_suffix}",
        )
    else:
        try:
            _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=False)
        except Exception:
            pass
        plan_df = st.session_state.get("plan_cuentas_df")
        if plan_df is None or getattr(plan_df, "empty", True):
            st.caption(
                "Vinculá el plan de cuentas del cliente para habilitar la descarga."
            )
            return

    if plan_df is None:
        return

    asientos_cons = _asientos_consolidados_biblioteca_banco(sociedad_id, banco)
    if auditoria_ui:
        if not _auditar_asientos_antes_export_ui(asientos_cons, plan_df):
            return
    else:
        # Sidebar: validar en silencio (evita expanders duplicados → removeChild).
        try:
            if plan_df.empty:
                st.caption("Falta el plan de cuentas del cliente.")
                return
            preparar_asientos_export_tango(asientos_cons, plan_df)
        except ExportacionTangoError:
            st.caption(
                "Hay errores de exportación. Revisá el panel principal para el detalle."
            )
            return
        except Exception as exc:
            st.caption(f"No se puede exportar aún: {exc}")
            return

    try:
        ruta_consolidada = _generar_excel_biblioteca_banco_consolidada(
            sociedad_id, cuit_activo or "000", nombre_activo or "", banco,
            plan_cuentas=plan_df,
        )
        with open(ruta_consolidada, "rb") as f:
            st.download_button(
                "📦 Generar Archivo de Importación Consolidado (Tango)",
                f,
                file_name=ruta_consolidada.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key=f"dl_tango_banco_v2_{slug}_consolidado_{key_suffix}",
                use_container_width=True,
            )
    except ExportacionTangoError as exc:
        if auditoria_ui:
            _mostrar_errores_exportacion_tango(exc)
        else:
            st.caption("Error al armar el Excel; mirá el panel principal.")
    except Exception as exc:
        st.error(f"No se pudo generar el archivo consolidado: {exc}")


def _render_panel_biblioteca_bancos(
    sociedad_id: int | None,
    banco: str,
    cuit_activo: str | None = None,
    nombre_activo: str | None = None,
) -> None:
    """Panel lateral: biblioteca acumulativa de asientos bancarios."""
    if st.session_state.get("_sidebar_unificada"):
        return
    entradas = _biblioteca_bancos_por_sociedad(sociedad_id, banco)
    periodos = _periodos_bancos_procesados_sociedad(sociedad_id, banco)
    slug = _slug_banco(banco)
    with st.sidebar:
        st.markdown("---")
        st.markdown("### Biblioteca de Asientos Bancarios")
        st.caption(f"**{banco}**")
        st.metric("Meses archivados", len(entradas))
        if periodos:
            st.caption(f"Asientos listos: **[{', '.join(periodos)}]**")
        else:
            st.caption("Sin meses guardados aún.")
        if entradas:
            if not cuit_activo and sociedad_id is not None:
                try:
                    cli = db.obtener_cliente(int(sociedad_id))
                    if cli:
                        cuit_activo = str(cli.get("cuit") or "")
                        nombre_activo = str(cli.get("nombre") or nombre_activo or "")
                except Exception:
                    pass
            _render_bloque_excel_consolidado_banco(
                sociedad_id,
                banco,
                cuit_activo,
                nombre_activo,
                key_suffix="sidebar",
                asegurar_plan=False,
                auditoria_ui=False,
            )
        if entradas and st.button(
            "Vaciar Biblioteca Bancaria",
            key=f"btn_vaciar_biblioteca_banco_v2_{slug}",
            type="secondary",
            use_container_width=True,
        ):
            _vaciar_biblioteca_banco_sociedad(sociedad_id, banco)
            st.rerun()


def _guardar_asiento_en_biblioteca_banco(
    sociedad_id: int,
    asientos: list[AsientoDevengamiento],
    rows: list[dict],
    *,
    banco: str,
) -> str:
    if not asientos or not rows:
        raise ValueError("No hay asiento bancario generado para guardar.")
    periodo = getattr(asientos[0], "periodo", "") or rows[0].get("Período", "")
    if not periodo:
        raise ValueError("No se pudo determinar el período del asiento bancario.")
    if _periodo_ya_en_biblioteca_banco(sociedad_id, periodo, banco):
        raise ValueError(
            f"El período {_periodo_a_etiqueta(periodo)} ya está archivado para {banco}."
        )
    for row in rows:
        row["Estado"] = "Ingresado"
    slug = _slug_banco(banco)
    entrada = {
        "sociedad_id": sociedad_id,
        "banco": banco,
        "periodo": periodo,
        "periodo_orden": _periodo_a_orden(periodo),
        "asientos": copy.deepcopy(asientos),
        "rows": copy.deepcopy(rows),
        "resumen_analitico": copy.deepcopy(st.session_state.get(f"{slug}_resumen_analitico")),
    }
    st.session_state.biblioteca_bancos.append(entrada)
    st.session_state.periodos_bancos_procesados.append({
        "sociedad_id": sociedad_id,
        "banco": banco,
        "periodo": periodo,
    })
    _persistir_biblioteca_en_disco()
    _limpiar_borrador_si_corresponde(_slug_banco(banco))
    return _periodo_a_etiqueta(periodo)


def _asientos_consolidados_biblioteca_banco(
    sociedad_id: int | None,
    banco: str | None = None,
) -> list[AsientoDevengamiento]:
    entradas = sorted(
        _biblioteca_bancos_por_sociedad(sociedad_id, banco),
        key=lambda e: e.get("periodo_orden", (9999, 99)),
    )
    todos: list[AsientoDevengamiento] = []
    for num_asiento, entrada in enumerate(entradas, start=1):
        for asiento in entrada.get("asientos", []):
            asiento.identificador = num_asiento
            todos.append(asiento)
    return todos


def _generar_excel_biblioteca_banco_consolidada(
    sociedad_id: int | None,
    cuit: str,
    nombre_cliente: str,
    banco: str,
    plan_cuentas: pd.DataFrame | None = None,
) -> Path:
    entradas = sorted(
        _biblioteca_bancos_por_sociedad(sociedad_id, banco),
        key=lambda e: e.get("periodo_orden", (9999, 99)),
    )
    if not entradas:
        raise ValueError("La biblioteca bancaria está vacía.")
    asientos = _asientos_consolidados_biblioteca_banco(sociedad_id, banco)
    primero = entradas[0]["periodo_orden"]
    ultimo = entradas[-1]["periodo_orden"]
    mes_ini, anio_ini = primero[1], primero[0]
    mes_fin, anio_fin = ultimo[1], ultimo[0]
    slug = _slug_banco(banco)
    etiqueta = (
        f"consolidado_{slug}_{mes_ini:02d}_{anio_ini}_{mes_fin:02d}_{anio_fin}"
        if len(entradas) > 1
        else f"{slug}_{mes_fin:02d}_{anio_fin}"
    )
    nombre_archivo = f"Conciliacion_Bancaria_Tango_{cuit}_{etiqueta}.xlsx"
    return generar_excel_tango_nativo(
        asientos=asientos,
        nombre_cliente=nombre_cliente,
        cuit=cuit,
        mes=mes_fin,
        anio=anio_fin,
        id_base=1,
        nombre_archivo=nombre_archivo,
        plan_cuentas=plan_cuentas,
    )


def _inicializar_sesion_banco(slug: str) -> None:
    _inicializar_balance_servidor_por_sociedad()
    _inicializar_estado_coordenadas_debe_haber(slug)


def _marcar_limpieza_formulario_mes_banco(slug: str) -> None:
    st.session_state[f"{slug}_limpiar_formulario_pendiente"] = True


def _aplicar_limpieza_formulario_mes_banco(slug: str) -> None:
    avanzar = st.session_state.pop(f"{slug}_periodo_mensual_avanzar", None)
    if avanzar:
        sid = st.session_state.get(_SOCiedad_KEY)
        if sid is not None:
            st.session_state[_clave_periodo_mensual_banco(slug, sid)] = avanzar
    if not st.session_state.pop(f"{slug}_limpiar_formulario_pendiente", False):
        return
    _limpiar_grilla_periodo_slug(slug)
    st.session_state["reset_counter"] = _iva_reset_counter() + 1


def _extension_planilla_iva(file_or_path) -> str:
    if isinstance(file_or_path, (str, Path)):
        return Path(file_or_path).suffix.lower()
    name = getattr(file_or_path, "name", "") or ""
    return Path(name).suffix.lower()


def _planilla_a_bytesio(archivo: Any) -> BytesIO:
    """Normaliza cualquier fuente de planilla a BytesIO en memoria (path-independent)."""
    if hasattr(archivo, "getvalue"):
        buf = BytesIO(archivo.getvalue())
        buf.name = getattr(archivo, "name", "planilla.xlsx")
        return buf
    if isinstance(archivo, (str, Path)):
        p = Path(archivo)
        if not p.is_file():
            raise FileNotFoundError(f"No se encontró la planilla: {p.name}")
        buf = BytesIO(p.read_bytes())
        buf.name = p.name
        return buf
    raise TypeError(f"Tipo de planilla no soportado: {type(archivo)!r}")


def _abrir_planilla_iva(file_or_path) -> tuple[BytesIO, bool]:
    """Abre planilla IVA siempre en memoria (compatible con red local y cloud)."""
    buf = _planilla_a_bytesio(file_or_path)
    es_csv = _extension_planilla_iva(buf.name) == ".csv"
    buf.seek(0)
    return buf, es_csv


def _cargar_planilla_default_servidor() -> BytesIO | None:
    """Planilla base opcional del despliegue (ruta relativa al directorio de la app)."""
    if not PLANILLA_IVA_DEFAULT.is_file():
        return None
    return _planilla_a_bytesio(PLANILLA_IVA_DEFAULT)


def _ruta_balance_default_sociedad(
    sociedad_id: int | None,
    *,
    nombre: str = "",
    cuit: str = "",
) -> str:
    """Ruta relativa preconfigurada del balance local para la sociedad activa."""
    return ruta_balance_local_por_sociedad(
        nombre=nombre or st.session_state.get("nombre_activo", ""),
        cuit=cuit or st.session_state.get("cuit_activo", ""),
        sociedad_id=sociedad_id,
    )


def _aplicar_ruta_balance_default_sociedad(sociedad_id: int | None) -> None:
    """Asigna ruta local preconfigurada y descarta URLs http legacy."""
    if sociedad_id is None:
        return
    rutas = st.session_state.balance_servidor_rutas_por_sociedad
    actual = sanitizar_ruta_unc(rutas.get(sociedad_id, ""))
    default = _ruta_balance_default_sociedad(sociedad_id)
    if actual and es_ruta_http_legacy(actual):
        rutas[sociedad_id] = default
    elif not actual and default:
        rutas.setdefault(sociedad_id, default)


def _auto_cargar_balance_local_si_existe(sociedad_id: int | None) -> bool:
    """Carga en memoria el balance local si existe en disco y no hay buffer activo."""
    if sociedad_id is None or _es_entorno_cloud():
        return False
    _inicializar_balance_servidor_por_sociedad()
    if st.session_state.balance_servidor_buffer_por_sociedad.get(sociedad_id):
        return False
    ruta = _ruta_balance_servidor_sociedad(sociedad_id)
    if not ruta or es_ruta_http_legacy(ruta):
        return False
    try:
        buf = cargar_balance_desde_ruta_unc(ruta)
    except (FileNotFoundError, ValueError, OSError):
        return False
    st.session_state.balance_servidor_buffer_por_sociedad[sociedad_id] = buf
    st.session_state.balance_servidor_sync_at_por_sociedad[sociedad_id] = (
        datetime.now().strftime("%d/%m/%Y %H:%M")
    )
    return True


def _inicializar_balance_servidor_por_sociedad() -> None:
    """Rutas UNC y buffers de balance por sociedad (reemplaza Excel Cloud)."""
    if "balance_servidor_rutas_por_sociedad" not in st.session_state:
        st.session_state.balance_servidor_rutas_por_sociedad = {}
    if "balance_servidor_buffer_por_sociedad" not in st.session_state:
        st.session_state.balance_servidor_buffer_por_sociedad = {}
    if "balance_servidor_sync_at_por_sociedad" not in st.session_state:
        st.session_state.balance_servidor_sync_at_por_sociedad = {}
    # Migración legacy Excel Cloud → servidor local
    legacy_urls = st.session_state.pop("iva_excel_cloud_urls", None)
    legacy_bufs = st.session_state.pop("iva_planilla_cloud_por_sociedad", None)
    legacy_sync = st.session_state.pop("iva_planilla_cloud_sync_at_por_sociedad", None)
    sid = st.session_state.get("sociedad_activa")
    if sid is not None and legacy_urls and sid in legacy_urls:
        st.session_state.balance_servidor_rutas_por_sociedad.setdefault(sid, legacy_urls[sid])
    if sid is not None and legacy_bufs and sid in legacy_bufs:
        st.session_state.balance_servidor_buffer_por_sociedad[sid] = legacy_bufs[sid]
    if sid is not None and legacy_sync and sid in legacy_sync:
        st.session_state.balance_servidor_sync_at_por_sociedad[sid] = legacy_sync[sid]
    legacy_url = st.session_state.pop("iva_excel_cloud_url", None)
    legacy_buf = st.session_state.pop("iva_planilla_cloud_buffer", None)
    legacy_at = st.session_state.pop("iva_planilla_cloud_sync_at", None)
    if sid is not None and legacy_url:
        st.session_state.balance_servidor_rutas_por_sociedad.setdefault(sid, legacy_url)
    if sid is not None and legacy_buf is not None:
        st.session_state.balance_servidor_buffer_por_sociedad[sid] = legacy_buf
    if sid is not None and legacy_at:
        st.session_state.balance_servidor_sync_at_por_sociedad[sid] = legacy_at
    if sid is not None:
        _aplicar_ruta_balance_default_sociedad(sid)


def _ruta_balance_servidor_sociedad(sociedad_id: int | None) -> str:
    _inicializar_balance_servidor_por_sociedad()
    if sociedad_id is None:
        return ""
    raw = st.session_state.balance_servidor_rutas_por_sociedad.get(sociedad_id, "")
    return sanitizar_ruta_unc(str(raw))


def _cargar_balance_cifrado_si_existe(sociedad_id: int) -> BytesIO | None:
    """En Cloud, rehidrata el buffer desde el último balance cifrado del usuario."""
    if not _es_entorno_cloud():
        return None
    try:
        from seguridad_datos import directorio_seguro, leer_cifrado, tiene_clave_cifrado

        if not tiene_clave_cifrado():
            return None
        carpeta = directorio_seguro(_usuario_oficina_actual(), "balances")
        prefijo = f"soc_{sociedad_id}_"
        candidatos = sorted(
            [p for p in carpeta.glob("*.enc") if p.name.startswith(prefijo)],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidatos:
            return None
        return BytesIO(leer_cifrado(candidatos[0]))
    except Exception:
        return None


def _buffer_balance_servidor_sociedad(sociedad_id: int | None) -> BytesIO | None:
    _inicializar_balance_servidor_por_sociedad()
    if sociedad_id is None:
        return None
    buf = st.session_state.balance_servidor_buffer_por_sociedad.get(sociedad_id)
    if buf is not None:
        return buf
    hidratado = _cargar_balance_cifrado_si_existe(int(sociedad_id))
    if hidratado is not None:
        st.session_state.balance_servidor_buffer_por_sociedad[sociedad_id] = hidratado
        st.session_state.balance_servidor_sync_at_por_sociedad[sociedad_id] = (
            datetime.now().strftime("%d/%m/%Y %H:%M")
        )
        return hidratado
    return None


def _resolver_fuente_balance(
    archivo_uploader,
    sociedad_id: int | None,
) -> tuple[BytesIO | None, str, str]:
    """
    Prioridad: buffer servidor local (UNC) > uploader de contingencia.
    Retorna (buffer, origen, etiqueta).
    """
    servidor = _buffer_balance_servidor_sociedad(sociedad_id)
    if servidor is not None:
        buf = _planilla_a_bytesio(servidor)
        ruta = _ruta_balance_servidor_sociedad(sociedad_id)
        etiqueta = ruta[:80] + ("…" if len(ruta) > 80 else "")
        return buf, "servidor", etiqueta or "Servidor local (carpeta compartida)"

    if archivo_uploader is not None:
        buf = _planilla_a_bytesio(archivo_uploader)
        return buf, "upload", getattr(archivo_uploader, "name", "balance.xlsx")

    return None, "", ""


def _solapas_disponibles_archivo(archivo) -> list[str]:
    """Lista solapas del workbook en memoria (diagnóstico UI bancos)."""
    try:
        buf, _ = _abrir_planilla_iva(archivo)
        return listar_solapas_excel(buf)
    except Exception:
        return []


def _ruta_balance_bancos_efectiva(sociedad_id: int | None) -> str:
    """Ruta pegada por el usuario en bancos; si está vacía, usa el Excel local del proyecto."""
    raw = str(st.session_state.get(_RUTA_BALANCE_BANCOS_INPUT, "") or "").strip()
    if raw:
        return sanitizar_ruta_unc(raw)
    _aplicar_ruta_balance_default_sociedad(sociedad_id)
    default = _ruta_balance_default_sociedad(sociedad_id) or BALANCE_EXCEL_PROYECTO
    return sanitizar_ruta_unc(default)


def _cargar_balance_servidor_bancos(
    sociedad_id: int | None,
    *,
    force: bool = False,
    ruta: str | None = None,
) -> tuple[bool, str]:
    """Carga el Excel del balance para Conciliación Bancaria desde la ruta indicada."""
    if sociedad_id is None:
        return False, ""
    _inicializar_balance_servidor_por_sociedad()
    if not force and st.session_state.balance_servidor_buffer_por_sociedad.get(sociedad_id):
        ruta_activa = st.session_state.balance_servidor_rutas_por_sociedad.get(sociedad_id, "")
        return True, ruta_activa
    ruta_limpia = sanitizar_ruta_unc(ruta or _ruta_balance_bancos_efectiva(sociedad_id))
    if not ruta_limpia:
        return False, ""
    st.session_state.balance_servidor_rutas_por_sociedad[sociedad_id] = ruta_limpia
    try:
        buf = cargar_balance_desde_ruta_unc(ruta_limpia)
    except (FileNotFoundError, ValueError, OSError):
        return False, ruta_limpia
    st.session_state.balance_servidor_buffer_por_sociedad[sociedad_id] = buf
    st.session_state.balance_servidor_sync_at_por_sociedad[sociedad_id] = (
        datetime.now().strftime("%d/%m/%Y %H:%M")
    )
    return True, ruta_limpia


def _cargar_balance_proyecto_obligatorio_bancos(
    sociedad_id: int | None,
    *,
    force: bool = False,
) -> bool:
    """Compatibilidad: auto-carga usando la ruta del buzón bancario o el default del proyecto."""
    ok, _ = _cargar_balance_servidor_bancos(sociedad_id, force=force)
    return ok


def _render_conexion_servidor_banco(sociedad_id: int | None, banco: str) -> None:
    """Panel UNC / carpeta compartida para Conciliación Bancaria (espejo de Devengamientos)."""
    _inicializar_balance_servidor_por_sociedad()
    slug = _slug_banco(banco)
    ok_key = f"banco_sync_ok_v2_{slug}_{sociedad_id}"
    err_key = f"banco_sync_err_v2_{slug}_{sociedad_id}"
    err_ruta_key = f"banco_sync_err_ruta_v2_{slug}_{sociedad_id}"
    if st.session_state.pop(ok_key, False):
        ruta_ok = st.session_state.pop(f"banco_sync_ruta_v2_{slug}_{sociedad_id}", "")
        st.success(
            f"✓ Balance cargado en memoria desde `{ruta_ok or _ruta_balance_bancos_efectiva(sociedad_id)}`."
        )
    if st.session_state.pop(err_key, False):
        ruta_err = st.session_state.pop(err_ruta_key, "")
        st.error(
            f"❌ No se encontró el archivo. Ruta intentada: `{ruta_err}`. "
            "Verificá que el Excel esté en la carpeta del proyecto o en la ruta UNC."
        )

    with st.expander("🏦 Conexión Servidor — Conciliación Bancaria", expanded=False):
        if sociedad_id is None:
            st.warning("Seleccioná una sociedad para configurar la ruta del balance.")
            return

        st.caption(
            f"Ingresá la ruta UNC o relativa al proyecto del **Balance completo** para **{banco}**. "
            "El sistema leerá la solapa del banco activo con escaneo matricial. "
            "Ejemplo UNC: `\\\\TANGOSRV\\Balances\\OFTALMOLOGIA RELE MAR DEL PLATA SRL.xlsx` — "
            f"Ejemplo local: `{BALANCE_EXCEL_PROYECTO}`"
        )

        if _es_entorno_cloud():
            st.info(
                "Estás en **Streamlit Cloud**: no hay Escritorio ni `T:\\`. "
                "Subí el Excel del balance con el cargador de abajo."
            )
        up_bal = st.file_uploader(
            "Subir Balance (.xlsx / .xls / .csv)",
            type=["xlsx", "xls", "csv"],
            key=f"uploader_balance_bancos_{slug}_{sociedad_id}",
            help="Alternativa a la ruta local/UNC. Obligatorio en Cloud.",
        )
        if up_bal is not None:
            fp = f"{up_bal.name}:{getattr(up_bal, 'size', 0)}"
            fp_key = f"_bal_up_fp_bancos_{sociedad_id}"
            if st.session_state.get(fp_key) != fp:
                raw = up_bal.getvalue()
                buffers = st.session_state.balance_servidor_buffer_por_sociedad
                sync_at_dict = st.session_state.balance_servidor_sync_at_por_sociedad
                buffers[sociedad_id] = BytesIO(raw)
                sync_at_dict[sociedad_id] = datetime.now().strftime("%d/%m/%Y %H:%M")
                st.session_state[f"{slug}_skip_default_planilla"] = True
                st.session_state[fp_key] = fp
                if _es_entorno_cloud():
                    try:
                        from seguridad_datos import guardar_balance_cifrado, tiene_clave_cifrado

                        if tiene_clave_cifrado():
                            enc = guardar_balance_cifrado(
                                _usuario_oficina_actual(),
                                sociedad_id,
                                up_bal.name,
                                raw,
                            )
                            st.caption(f"Balance cifrado en disco: `{enc.name}`")
                        else:
                            st.warning(
                                "Balance solo en memoria de esta sesión. "
                                "Configurá DATA_ENCRYPTION_KEY en Secrets para guardarlo cifrado."
                            )
                    except Exception as exc:
                        st.warning(f"No se pudo cifrar el balance en disco: {exc}")
                st.success(f"✓ Balance en memoria: `{up_bal.name}`.")
                st.rerun()

        rutas = st.session_state.balance_servidor_rutas_por_sociedad
        _aplicar_ruta_balance_default_sociedad(sociedad_id)
        if _RUTA_BALANCE_BANCOS_INPUT not in st.session_state:
            st.session_state[_RUTA_BALANCE_BANCOS_INPUT] = (
                rutas.get(sociedad_id) or BALANCE_EXCEL_PROYECTO
            )

        if not _es_entorno_cloud():
            ruta_input = st.text_input(
                "Ruta del Balance (fija para bancos)",
                key=_RUTA_BALANCE_BANCOS_INPUT,
                placeholder=r".\Copia de OFTALMOLOGIA RELE Balance 2026.xlsx",
                help="Ruta de red Windows (UNC), ruta absoluta o relativa (./) al Excel/CSV del balance mensual.",
            )
            rutas[sociedad_id] = sanitizar_ruta_unc(ruta_input)
            if es_ruta_http_legacy(rutas.get(sociedad_id, "")):
                st.warning(
                    "Las URLs web (Excel Cloud) ya no se usan. "
                    "Reemplazá la ruta por un archivo local relativo (./) o una carpeta compartida UNC."
                )
                default = _ruta_balance_default_sociedad(sociedad_id) or BALANCE_EXCEL_PROYECTO
                rutas[sociedad_id] = default
                st.session_state[_RUTA_BALANCE_BANCOS_INPUT] = default
        else:
            rutas[sociedad_id] = rutas.get(sociedad_id) or ""

        buffers = st.session_state.balance_servidor_buffer_por_sociedad
        sync_at_dict = st.session_state.balance_servidor_sync_at_por_sociedad

        col_sync, col_clear = st.columns(2)
        with col_sync:
            if not _es_entorno_cloud() and st.button(
                "🔄 Cargar Saldo desde Servidor",
                type="primary",
                key=_BTN_CARGAR_SALDO_BANCOS,
                use_container_width=True,
            ):
                ruta_limpia = sanitizar_ruta_unc(rutas.get(sociedad_id, ""))
                rutas[sociedad_id] = ruta_limpia
                if not ruta_limpia:
                    st.session_state[err_key] = True
                    st.session_state[err_ruta_key] = "(vacía)"
                    st.rerun()
                elif es_ruta_http_legacy(ruta_limpia):
                    st.session_state[err_key] = True
                    st.session_state[err_ruta_key] = ruta_limpia
                    st.rerun()
                else:
                    ok, ruta_usada = _cargar_balance_servidor_bancos(
                        sociedad_id, force=True, ruta=ruta_limpia,
                    )
                    if ok:
                        st.session_state[ok_key] = True
                        st.session_state[f"banco_sync_ruta_v2_{slug}_{sociedad_id}"] = ruta_usada
                        st.session_state[f"{slug}_skip_default_planilla"] = True
                        st.session_state.pop(f"{slug}_auto_fp", None)
                    else:
                        st.session_state[err_key] = True
                        st.session_state[err_ruta_key] = ruta_usada
                    st.rerun()

        with col_clear:
            if buffers.get(sociedad_id) and st.button(
                "✖ Quitar balance en memoria",
                key=_btn_quitar_bancos_v2(slug, sociedad_id),
                use_container_width=True,
            ):
                buffers.pop(sociedad_id, None)
                sync_at_dict.pop(sociedad_id, None)
                st.session_state.pop(f"{slug}_auto_fp", None)
                st.rerun()

        if buffers.get(sociedad_id):
            sync_at = sync_at_dict.get(sociedad_id, "")
            hojas = _solapas_disponibles_archivo(buffers[sociedad_id])
            st.info(
                f"📁 **Balance activo** — última carga: {sync_at or 'reciente'}. "
                f"Solapas detectadas: **{', '.join(hojas) if hojas else '—'}**"
            )


def _aviso_diagnostico_extraccion_banco(
    *,
    impuesto: str,
    periodo_mensual: str,
    resultado,
    archivo,
    aviso_periodo: str,
) -> str:
    """Mensaje informativo cuando falla solapa, período o filas vacías (módulo bancos)."""
    hojas = _solapas_disponibles_archivo(archivo)
    hojas_txt = ", ".join(hojas) if hojas else "(ninguna)"
    if resultado.error_tipo == "sheet_not_found":
        return (
            f"⚠️ No se encontró una solapa que coincida con el banco seleccionado (**{impuesto}**). "
            f"Las solapas disponibles en este Excel son: **{hojas_txt}**"
        )
    if resultado.error_tipo == "month_column_not_found":
        periodos = resultado.periodos_disponibles or []
        periodos_txt = ", ".join(periodos[:24]) if periodos else "(ninguno detectado)"
        solapa = resultado.solapa_resuelta or impuesto
        return (
            f"⚠️ No se encontró la columna del período **{periodo_mensual}** en la solapa "
            f"«{solapa}». Períodos detectados en cabecera: **{periodos_txt}**"
        )
    if not (resultado.filas or []):
        solapa = resultado.solapa_resuelta or impuesto
        return (
            f"⚠️ No se encontraron filas con importes en la solapa «{solapa}» "
            f"para el período **{periodo_mensual}**. "
            f"Solapas del archivo: **{hojas_txt}**. "
            "Revisá el banco seleccionado o probá otro mes."
        )
    return aviso_periodo


def _render_conexion_servidor_local(sociedad_id: int | None, impuesto: str) -> None:
    """Panel UNC / carpeta compartida por sociedad."""
    _inicializar_balance_servidor_por_sociedad()
    slug = _slug_impuesto(impuesto)
    with st.expander("🏢 Conexión Servidor Local (Carpeta Compartida)", expanded=False):
        if sociedad_id is None:
            st.warning("Seleccioná una sociedad para configurar la ruta del balance.")
            return

        st.caption(
            "Ingresá la ruta UNC o relativa al proyecto del **Balance completo**. "
            f"El sistema leerá la solapa del impuesto activo (**{impuesto}**) con escaneo matricial. "
            "Ejemplo UNC: `\\\\TANGOSRV\\Balances\\OFTALMOLOGIA RELE MAR DEL PLATA SRL.xlsx` — "
            "Ejemplo local: `./Copia de OFTALMOLOGIA RELE Balance 2026.xlsx`"
        )

        if _es_entorno_cloud():
            st.info(
                "Estás en **Streamlit Cloud**: no hay Escritorio ni `T:\\`. "
                "Subí el Excel del balance con el cargador."
            )
        up_bal = st.file_uploader(
            "Subir Balance (.xlsx / .xls / .csv)",
            type=["xlsx", "xls", "csv"],
            key=f"uploader_balance_dev_{slug}_{sociedad_id}",
            help="Alternativa a la ruta local/UNC. En Cloud usá siempre el uploader.",
        )
        if up_bal is not None:
            fp = f"{up_bal.name}:{getattr(up_bal, 'size', 0)}"
            fp_key = f"_bal_up_fp_dev_{slug}_{sociedad_id}"
            if st.session_state.get(fp_key) != fp:
                raw = up_bal.getvalue()
                buffers = st.session_state.balance_servidor_buffer_por_sociedad
                sync_at_dict = st.session_state.balance_servidor_sync_at_por_sociedad
                buffers[sociedad_id] = BytesIO(raw)
                sync_at_dict[sociedad_id] = datetime.now().strftime("%d/%m/%Y %H:%M")
                st.session_state[f"{slug}_skip_default_planilla"] = True
                st.session_state[fp_key] = fp
                if _es_entorno_cloud():
                    try:
                        from seguridad_datos import guardar_balance_cifrado, tiene_clave_cifrado

                        if tiene_clave_cifrado():
                            guardar_balance_cifrado(
                                _usuario_oficina_actual(),
                                sociedad_id,
                                up_bal.name,
                                raw,
                            )
                        else:
                            st.warning(
                                "Balance solo en memoria. Configurá DATA_ENCRYPTION_KEY en Secrets."
                            )
                    except Exception as exc:
                        st.warning(f"No se pudo cifrar el balance: {exc}")
                st.success(f"✓ Balance en memoria: `{up_bal.name}`.")
                st.rerun()

        rutas = st.session_state.balance_servidor_rutas_por_sociedad
        _aplicar_ruta_balance_default_sociedad(sociedad_id)
        if not _es_entorno_cloud():
            _auto_cargar_balance_local_si_existe(sociedad_id)
        input_key = f"balance_servidor_ruta_{slug}_{sociedad_id}"
        if input_key not in st.session_state:
            st.session_state[input_key] = rutas.get(sociedad_id, "")

        if not _es_entorno_cloud():
            ruta_input = st.text_input(
                "Ruta del Balance (UNC o relativa al proyecto)",
                key=input_key,
                placeholder=r".\Copia de OFTALMOLOGIA RELE Balance 2026.xlsx",
                help="Ruta de red Windows (UNC), ruta absoluta o relativa (./) al Excel/CSV del balance mensual.",
            )
            rutas[sociedad_id] = sanitizar_ruta_unc(ruta_input)
            if es_ruta_http_legacy(rutas.get(sociedad_id, "")):
                st.warning(
                    "Las URLs web (Excel Cloud) ya no se usan. "
                    "Reemplazá la ruta por un archivo local relativo (./) o una carpeta compartida UNC."
                )
                default = _ruta_balance_default_sociedad(sociedad_id)
                if default:
                    rutas[sociedad_id] = default
                    st.session_state[input_key] = default
        else:
            rutas[sociedad_id] = rutas.get(sociedad_id) or ""

        buffers = st.session_state.balance_servidor_buffer_por_sociedad
        sync_at_dict = st.session_state.balance_servidor_sync_at_por_sociedad

        col_sync, col_clear = st.columns(2)
        with col_sync:
            if not _es_entorno_cloud() and st.button(
                "🔄 Cargar Balance desde Servidor",
                type="primary",
                key=f"btn_sync_balance_servidor_{slug}_{sociedad_id}",
                use_container_width=True,
            ):
                ruta_limpia = sanitizar_ruta_unc(rutas.get(sociedad_id, ""))
                rutas[sociedad_id] = ruta_limpia
                if not ruta_limpia:
                    st.error("Ingresá la ruta del balance antes de cargar.")
                elif es_ruta_http_legacy(ruta_limpia):
                    st.error(
                        "No se pueden cargar balances desde URLs web. "
                        "Usá una ruta local relativa (./archivo.xlsx) o UNC de red."
                    )
                else:
                    try:
                        with st.spinner("Leyendo balance (escaneo matricial)…"):
                            buf = cargar_balance_desde_ruta_unc(ruta_limpia)
                            buffers[sociedad_id] = buf
                            sync_at_dict[sociedad_id] = datetime.now().strftime("%d/%m/%Y %H:%M")
                            st.session_state[f"{slug}_skip_default_planilla"] = True
                        st.success(
                            f"✓ Balance cargado en memoria ({len(buf.getvalue()):,} bytes)."
                        )
                        st.rerun()
                    except FileNotFoundError:
                        st.error(
                            f"❌ No se encontró el archivo. Ruta intentada: `{ruta_limpia}`. "
                            "Verificá que el Excel esté en la carpeta del proyecto o en la ruta UNC."
                        )
                    except ValueError as exc:
                        st.error(f"❌ {exc}")
                    except Exception as exc:
                        st.error(
                            f"❌ Error al cargar balance desde `{ruta_limpia}`: {exc}"
                        )

        with col_clear:
            if buffers.get(sociedad_id) and st.button(
                "✖ Quitar balance en memoria",
                key=f"btn_clear_balance_servidor_{slug}_{sociedad_id}",
                use_container_width=True,
            ):
                buffers.pop(sociedad_id, None)
                sync_at_dict.pop(sociedad_id, None)
                st.rerun()

        if buffers.get(sociedad_id):
            sync_at = sync_at_dict.get(sociedad_id, "")
            st.info(
                f"📁 **Balance del servidor activo para esta sociedad** — última carga: "
                f"{sync_at or 'reciente'}. Tiene prioridad sobre el uploader local."
            )


def _resolver_fuente_planilla_iva(
    archivo_uploader,
    sociedad_id: int | None,
) -> tuple[BytesIO | None, str, str]:
    """Compatibilidad: delega a _resolver_fuente_balance."""
    return _resolver_fuente_balance(archivo_uploader, sociedad_id)


def _generar_asiento_iva_planilla_posicional(
    plan_cuentas_df: pd.DataFrame,
    datos: dict,
    saldo_tecnico: float | list[float],
    saldo_libre: float | list[float],
    retenciones: float,
    percepciones: float,
) -> tuple[list[dict], dict, float, bool]:
    """Asiento mensual IVA desde planilla posicional + cargas manuales multi-saldo."""
    saldos_tec = _normalizar_lista_saldos_iva(saldo_tecnico)
    saldos_lib = _normalizar_lista_saldos_iva(saldo_libre)
    resumen = _calcular_desglose_posicion_iva(
        df_debito_21=datos.get("df_debito_21", 0.0),
        df_debito_105=datos.get("df_debito_105", 0.0),
        df_debito_27=datos.get("df_debito_27", 0.0),
        nc_compras=datos.get("nc_compras", 0.0),
        nc_compras_27=datos.get("nc_compras_27", 0.0),
        cf_credito_21=datos.get("cf_credito_21", 0.0),
        cf_credito_105=datos.get("cf_credito_105", 0.0),
        cf_credito_27=datos.get("cf_credito_27", 0.0),
        nc_ventas=datos.get("nc_ventas", 0.0),
        nc_ventas_27=datos.get("nc_ventas_27", 0.0),
        retenciones=retenciones,
        percepciones=percepciones,
        saldos_tecnicos=saldos_tec,
        saldos_libre=saldos_lib,
    )
    roles_debe = (
        ("ventas_21", datos.get("df_debito_21", 0.0)),
        ("ventas_105", datos.get("df_debito_105", 0.0)),
        ("ventas_27", datos.get("df_debito_27", 0.0)),
        ("nc_compras", datos.get("nc_compras", 0.0)),
        ("nc_compras_27", datos.get("nc_compras_27", 0.0)),
    )
    roles_haber = (
        ("compras_21", datos.get("cf_credito_21", 0.0)),
        ("compras_105", datos.get("cf_credito_105", 0.0)),
        ("compras_27", datos.get("cf_credito_27", 0.0)),
        ("nc_ventas", datos.get("nc_ventas", 0.0)),
        ("nc_ventas_27", datos.get("nc_ventas_27", 0.0)),
        ("retenciones", retenciones),
        ("percepciones", percepciones),
    )
    lineas = _armar_lineas_movimiento_iva(plan_cuentas_df, resumen, roles_debe, roles_haber)
    lineas, dif_ajuste, ajustado = _aplicar_loop_review_ventas_21(lineas)
    return lineas, resumen, dif_ajuste, ajustado


def procesar_planilla_iva(
    file_or_path,
    plan_df: pd.DataFrame,
    sal_tec: float | list[float],
    sal_lib: float | list[float],
    ret: float,
    perc: float,
    nombre_solapa_impuesto: str | None = None,
    periodo_mensual: str | None = None,
) -> tuple[list[AsientoDevengamiento], dict]:
    """Motor unificado: balance Excel (solapa del impuesto) + inputs manuales → asiento Tango."""
    solapa = nombre_solapa_impuesto or _impuesto_activo_devengamientos()
    buf, es_csv = _abrir_planilla_iva(file_or_path)
    datos = leer_planilla_iva_posicional(
        buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        periodo_mensual=periodo_mensual,
    )
    mes = datos.get("periodo_mes")
    anio = datos.get("periodo_anio")
    if mes is None or anio is None:
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if not periodo:
            raise ValueError(
                f"No se pudo leer el período desde la solapa '{solapa}' del balance: "
                f"'{datos.get('periodo_texto', '')}'. Use formato MM-AAAA (ej: 04-2025)."
            )
        mes, anio = periodo
    mes, anio = int(mes), int(anio)
    lineas, resumen, _, _ = _generar_asiento_iva_planilla_posicional(
        plan_df, datos, sal_tec, sal_lib, ret, perc,
    )
    if not lineas:
        return [], resumen
    asiento = _lineas_a_asiento(lineas, plan_df, mes, anio, identificador=1)
    asiento._roles_renglones = [str(l.get("_rol", "")) for l in lineas]  # type: ignore[attr-defined]
    asiento._resumen_analitico = resumen  # type: ignore[attr-defined]
    return [asiento], resumen


def procesar_planilla_iibb(
    file_or_path,
    plan_df: pd.DataFrame,
    saldos_favor: float | list[float],
    ret: float,
    perc: float,
    ret_banc: float,
    nombre_solapa_impuesto: str | None = None,
    periodo_mensual: str | None = None,
    tipo_impuesto: str | None = None,
) -> tuple[list[AsientoDevengamiento], dict]:
    """Motor unificado IIBB/CM: balance Excel + inputs manuales → asiento Tango."""
    solapa = nombre_solapa_impuesto or _impuesto_activo_devengamientos()
    tipo = tipo_impuesto or "IIBB"
    buf, es_csv = _abrir_planilla_iva(file_or_path)
    datos = leer_planilla_iibb_por_etiquetas(
        buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        periodo_mensual=periodo_mensual,
    )
    mes = datos.get("periodo_mes")
    anio = datos.get("periodo_anio")
    if mes is None or anio is None:
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if not periodo:
            raise ValueError(
                f"No se pudo leer el período desde la solapa '{solapa}' del balance: "
                f"'{datos.get('periodo_texto', '')}'. Use formato MM-AAAA (ej: 04-2025)."
            )
        mes, anio = periodo
    mes, anio = int(mes), int(anio)
    lineas, resumen, _, _ = _generar_asiento_iibb_planilla_posicional(
        plan_df, datos, saldos_favor, ret, perc, ret_banc,
    )
    if not lineas:
        return [], resumen
    asiento = _lineas_a_asiento(
        lineas, plan_df, mes, anio, identificador=1, tipo_impuesto=tipo,
    )
    asiento._roles_renglones = [str(l.get("_rol", "")) for l in lineas]  # type: ignore[attr-defined]
    asiento._resumen_analitico = resumen  # type: ignore[attr-defined]
    return [asiento], resumen


def leer_planilla_sueldos_por_etiquetas(
    file_buffer,
    *,
    es_csv: bool = False,
    nombre_solapa_impuesto: str = "Sueldos",
    periodo_mensual: str | None = None,
) -> dict:
    buf = file_buffer
    if hasattr(buf, "seek"):
        buf.seek(0)
    if es_csv:
        return leer_datos_balance_por_ficha(
            buf, nombre_solapa_impuesto, es_csv=True, periodo_mensual=periodo_mensual,
        )
    return leer_datos_balance_por_ficha(
        buf, nombre_solapa_impuesto, es_csv=False, periodo_mensual=periodo_mensual,
    )


def _calcular_desglose_posicion_sueldos(
    *,
    sueldos_brutos: float,
    sueldos_netos: float,
    aportes: float,
    contribuciones: float,
    sueldos_pagar: float,
    cargas_sociales_pagar: float,
    saldos_favor: list[float],
) -> dict:
    bruto_o_neto = round(sueldos_brutos or sueldos_netos, 2)
    total_debe = round(bruto_o_neto + aportes + contribuciones, 2)
    saldos = [round(float(x), 2) for x in saldos_favor if round(float(x), 2) > 0]
    total_haber = round(sueldos_pagar + cargas_sociales_pagar + sum(saldos), 2)
    diferencia_previa = round(total_debe - total_haber, 2)
    if diferencia_previa > 0:
        resultado_tipo, resultado_lado, resultado_monto = "Sueldos a Pagar (cierre)", "Haber", diferencia_previa
    elif diferencia_previa < 0:
        resultado_tipo, resultado_lado, resultado_monto = "Saldo a Favor Sueldos", "Debe", abs(diferencia_previa)
    else:
        resultado_tipo, resultado_lado, resultado_monto = "Equilibrado (sin saldo de cierre)", "—", 0.0
    return {
        "sueldos_brutos": round(sueldos_brutos, 2),
        "sueldos_netos": round(sueldos_netos, 2),
        "aportes": round(aportes, 2),
        "contribuciones": round(contribuciones, 2),
        "sueldos_pagar": round(sueldos_pagar, 2),
        "cargas_sociales_pagar": round(cargas_sociales_pagar, 2),
        "saldos_favor": saldos,
        "total_debe_mes": total_debe,
        "total_haber_previo": total_haber,
        "diferencia_previa": diferencia_previa,
        "resultado_tipo": resultado_tipo,
        "resultado_lado": resultado_lado,
        "resultado_monto": resultado_monto,
    }


def _armar_lineas_movimiento_sueldos(plan_cuentas_df: pd.DataFrame, resumen: dict) -> list[dict]:
    lineas: list[dict] = []
    bruto = resumen.get("sueldos_brutos", 0.0)
    neto = resumen.get("sueldos_netos", 0.0)
    sueldo_mes = bruto if bruto > 0 else neto
    if sueldo_mes > 0:
        cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, "sueldos_jornales")
        lineas.append(_linea_asiento_iibb(cod, desc, sueldo_mes, 0.0, "sueldos_jornales"))
    for rol, key in (("aportes", "aportes"), ("contribuciones", "contribuciones")):
        monto = resumen.get(key, 0.0)
        if monto > 0:
            cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iibb(cod, desc, monto, 0.0, rol))
    for rol, key in (("sueldos_pagar", "sueldos_pagar"), ("cargas_sociales_pagar", "cargas_sociales_pagar")):
        monto = resumen.get(key, 0.0)
        if monto > 0:
            cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, rol))
    for monto in resumen.get("saldos_favor", []):
        cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, "sueldos_saldo_favor")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, "sueldos_saldo_favor"))
    dif = resumen["diferencia_previa"]
    if dif > 0:
        cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, "sueldos_pagar_cierre")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, dif, "sueldos_pagar_cierre"))
    elif dif < 0:
        cod, desc = obtener_cuenta_tango_sueldos(plan_cuentas_df, "sueldos_saldo_favor")
        lineas.append(_linea_asiento_iibb(cod, "Saldo a Favor Sueldos", abs(dif), 0.0, "sueldos_saldo_favor_nuevo"))
    return lineas


def _generar_asiento_sueldos_planilla_posicional(
    plan_cuentas_df: pd.DataFrame,
    datos: dict,
    saldos_favor: float | list[float],
) -> tuple[list[dict], dict, float, bool]:
    saldos = _normalizar_lista_saldos_iibb(saldos_favor)
    resumen = _calcular_desglose_posicion_sueldos(
        sueldos_brutos=datos.get("sueldos_brutos", 0.0),
        sueldos_netos=datos.get("sueldos_netos", 0.0),
        aportes=datos.get("aportes", 0.0),
        contribuciones=datos.get("contribuciones", 0.0),
        sueldos_pagar=datos.get("sueldos_pagar", 0.0),
        cargas_sociales_pagar=datos.get("cargas_sociales_pagar", 0.0),
        saldos_favor=saldos,
    )
    lineas = _armar_lineas_movimiento_sueldos(plan_cuentas_df, resumen)
    ficha = obtener_ficha_impuesto("Sueldos")
    lineas, dif_ajuste, ajustado = _aplicar_loop_review_por_rol(
        lineas, ficha.get("cuenta_ajuste_centavos_rol", "sueldos_jornales"),
    )
    return lineas, resumen, dif_ajuste, ajustado


def procesar_planilla_sueldos(
    file_or_path,
    plan_df: pd.DataFrame,
    saldos_favor: float | list[float],
    nombre_solapa_impuesto: str | None = None,
    periodo_mensual: str | None = None,
) -> tuple[list[AsientoDevengamiento], dict]:
    solapa = nombre_solapa_impuesto or "Sueldos"
    buf, es_csv = _abrir_planilla_iva(file_or_path)
    datos = leer_planilla_sueldos_por_etiquetas(
        buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        periodo_mensual=periodo_mensual,
    )
    mes = datos.get("periodo_mes")
    anio = datos.get("periodo_anio")
    if mes is None or anio is None:
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if not periodo:
            raise ValueError(
                f"No se pudo leer el período desde la solapa '{solapa}' del balance: "
                f"'{datos.get('periodo_texto', '')}'. Use formato MM-AAAA (ej: 04-2025)."
            )
        mes, anio = periodo
    mes, anio = int(mes), int(anio)
    lineas, resumen, _, _ = _generar_asiento_sueldos_planilla_posicional(plan_df, datos, saldos_favor)
    if not lineas:
        return [], resumen
    asiento = _lineas_a_asiento(lineas, plan_df, mes, anio, identificador=1, tipo_impuesto="SUELDOS")
    asiento._roles_renglones = [str(l.get("_rol", "")) for l in lineas]  # type: ignore[attr-defined]
    asiento._resumen_analitico = resumen  # type: ignore[attr-defined]
    return [asiento], resumen


def leer_planilla_tish_por_etiquetas(
    file_buffer,
    *,
    es_csv: bool = False,
    nombre_solapa_impuesto: str = "TISH",
    periodo_mensual: str | None = None,
) -> dict:
    buf = file_buffer
    if hasattr(buf, "seek"):
        buf.seek(0)
    if es_csv:
        return leer_datos_balance_por_ficha(
            buf, nombre_solapa_impuesto, es_csv=True, periodo_mensual=periodo_mensual,
        )
    return leer_datos_balance_por_ficha(
        buf, nombre_solapa_impuesto, es_csv=False, periodo_mensual=periodo_mensual,
    )


def _calcular_desglose_posicion_tish(
    *,
    tasa_determinada: float,
    retenciones_tish: float,
    derecho_oficina: float,
    saldos_favor: list[float],
) -> dict:
    saldos = [round(float(x), 2) for x in saldos_favor if round(float(x), 2) > 0]
    total_descuentos = round(retenciones_tish + derecho_oficina + sum(saldos), 2)
    tasa = round(float(tasa_determinada), 2)
    diferencia_previa = round(tasa - total_descuentos, 2)
    if diferencia_previa > 0:
        resultado_tipo, resultado_lado, resultado_monto = "Tasa TISH a Pagar", "Haber", diferencia_previa
    elif diferencia_previa < 0:
        resultado_tipo, resultado_lado, resultado_monto = "Saldo a Favor TISH Nuevo Período", "Debe", abs(diferencia_previa)
    else:
        resultado_tipo, resultado_lado, resultado_monto = "Equilibrado (sin saldo de cierre)", "—", 0.0
    return {
        "tasa_determinada": tasa,
        "retenciones_tish": round(retenciones_tish, 2),
        "derecho_oficina": round(derecho_oficina, 2),
        "saldos_favor": saldos,
        "total_descuentos": total_descuentos,
        "diferencia_previa": diferencia_previa,
        "resultado_tipo": resultado_tipo,
        "resultado_lado": resultado_lado,
        "resultado_monto": resultado_monto,
    }


def _armar_lineas_movimiento_tish(plan_cuentas_df: pd.DataFrame, resumen: dict) -> list[dict]:
    lineas: list[dict] = []
    tasa = resumen.get("tasa_determinada", 0.0)
    if tasa > 0:
        cod, desc = obtener_cuenta_tango_tish(plan_cuentas_df, "gasto_tasa")
        lineas.append(_linea_asiento_iibb(cod, desc, tasa, 0.0, "gasto_tasa"))
    for rol, key in (("retenciones_tish", "retenciones_tish"), ("derecho_oficina", "derecho_oficina")):
        monto = resumen.get(key, 0.0)
        if monto > 0:
            cod, desc = obtener_cuenta_tango_tish(plan_cuentas_df, rol)
            lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, rol))
    for monto in resumen.get("saldos_favor", []):
        cod, desc = obtener_cuenta_tango_tish(plan_cuentas_df, "tish_saldo_favor")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, monto, "tish_saldo_favor"))
    dif = resumen["diferencia_previa"]
    if dif > 0:
        cod, desc = obtener_cuenta_tango_tish(plan_cuentas_df, "tasa_pagar")
        lineas.append(_linea_asiento_iibb(cod, desc, 0.0, dif, "tasa_pagar"))
    elif dif < 0:
        cod, desc = obtener_cuenta_tango_tish(plan_cuentas_df, "tish_saldo_favor_nuevo")
        lineas.append(_linea_asiento_iibb(cod, "Saldo a Favor TISH Nuevo Período", abs(dif), 0.0, "tish_saldo_favor_nuevo"))
    return lineas


def _generar_asiento_tish_planilla_posicional(
    plan_cuentas_df: pd.DataFrame,
    datos: dict,
    saldos_favor: float | list[float],
) -> tuple[list[dict], dict, float, bool]:
    saldos = _normalizar_lista_saldos_iibb(saldos_favor)
    resumen = _calcular_desglose_posicion_tish(
        tasa_determinada=datos.get("tasa_determinada", 0.0),
        retenciones_tish=datos.get("retenciones_tish", 0.0),
        derecho_oficina=datos.get("derecho_oficina", 0.0),
        saldos_favor=saldos,
    )
    lineas = _armar_lineas_movimiento_tish(plan_cuentas_df, resumen)
    ficha = obtener_ficha_impuesto("TISH")
    lineas, dif_ajuste, ajustado = _aplicar_loop_review_por_rol(
        lineas, ficha.get("cuenta_ajuste_centavos_rol", "gasto_tasa"),
    )
    return lineas, resumen, dif_ajuste, ajustado


def procesar_planilla_tish(
    file_or_path,
    plan_df: pd.DataFrame,
    saldos_favor: float | list[float],
    nombre_solapa_impuesto: str | None = None,
    periodo_mensual: str | None = None,
) -> tuple[list[AsientoDevengamiento], dict]:
    solapa = nombre_solapa_impuesto or "TISH"
    buf, es_csv = _abrir_planilla_iva(file_or_path)
    datos = leer_planilla_tish_por_etiquetas(
        buf, es_csv=es_csv, nombre_solapa_impuesto=solapa,
        periodo_mensual=periodo_mensual,
    )
    mes = datos.get("periodo_mes")
    anio = datos.get("periodo_anio")
    if mes is None or anio is None:
        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
        if not periodo:
            raise ValueError(
                f"No se pudo leer el período desde la solapa '{solapa}' del balance: "
                f"'{datos.get('periodo_texto', '')}'. Use formato MM-AAAA (ej: 04-2025)."
            )
        mes, anio = periodo
    mes, anio = int(mes), int(anio)
    lineas, resumen, _, _ = _generar_asiento_tish_planilla_posicional(plan_df, datos, saldos_favor)
    if not lineas:
        return [], resumen
    asiento = _lineas_a_asiento(lineas, plan_df, mes, anio, identificador=1, tipo_impuesto="TISH")
    asiento._roles_renglones = [str(l.get("_rol", "")) for l in lineas]  # type: ignore[attr-defined]
    asiento._resumen_analitico = resumen  # type: ignore[attr-defined]
    return [asiento], resumen


def _extraer_texto_pdf_upload(uploaded_file) -> str:
    import pdfplumber

    partes: list[str] = []
    with pdfplumber.open(BytesIO(uploaded_file.getvalue())) as pdf:
        for pagina in pdf.pages:
            partes.append(pagina.extract_text() or "")
    return "\n".join(partes)


def _norm_fila(texto: str) -> str:
    return _norm_desc_iva(str(texto))


def _texto_fila(df: pd.DataFrame, idx) -> str:
    partes: list[str] = []
    for v in df.loc[idx]:
        if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan":
            partes.append(str(v).strip())
    return " ".join(partes)


def _fila_contiene_marcador(texto_fila: str, marcadores: tuple[str, ...]) -> bool:
    norm = _norm_fila(texto_fila)
    return any(_norm_fila(m) in norm for m in marcadores)


def _buscar_fila_marcador(df: pd.DataFrame, marcadores: tuple[str, ...]) -> int | None:
    for idx in df.index:
        if _fila_contiene_marcador(_texto_fila(df, idx), marcadores):
            return int(idx)
    return None


def _es_fila_total(texto_fila: str) -> bool:
    t = _norm_fila(texto_fila)
    return any(
        x in t
        for x in (
            "total del debito",
            "total de debito",
            "total del credito",
            "total de credito",
            "total general",
        )
    )


def _extraer_monto_fila(row: pd.Series) -> float:
    montos: list[float] = []
    for v in row.tolist():
        if pd.isna(v) or not str(v).strip() or str(v).strip().lower() == "nan":
            continue
        texto = str(v).strip()
        try:
            num = float(texto.replace(".", "").replace(",", "."))
            if num > 0:
                montos.append(num)
                continue
        except ValueError:
            pass
        for match in re.findall(r"[\d\.]+,[\d]{2}", texto):
            montos.append(float(match.replace(".", "").replace(",", ".")))
    return montos[-1] if montos else 0.0


def _detectar_alicuota_en_fila(texto_fila: str) -> str | None:
    if any(x in texto_fila for x in ("10,50", "10.50", "10,5 %", "10,5%")):
        return "10.5"
    if any(x in texto_fila for x in ("21,00", "21.00", "21,0 %", "21,0%")):
        return "21"
    if any(x in texto_fila for x in ("27,00", "27.00")):
        return "27"
    return None


def _parsear_montos_fiscales_arca_desde_df(df: pd.DataFrame) -> dict[str, float]:
    """Recorre fila a fila la liquidación ARCA y extrae débito/crédito por alícuota."""
    montos = {
        "ventas_21": 0.0,
        "ventas_105": 0.0,
        "compras_21": 0.0,
        "compras_105": 0.0,
    }
    seccion_actual: str | None = None
    df = df.fillna("")

    for _, row in df.iterrows():
        texto = " ".join(
            str(v) for v in row
            if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"
        )
        if not texto.strip():
            continue
        norm = _norm_fila(texto)

        if "debito fiscal actividades" in norm:
            seccion_actual = "VENTAS"
            continue
        if "credito fiscal actividades" in norm or "crédito fiscal actividades" in norm:
            seccion_actual = "COMPRAS"
            continue
        if "liquidacion" in norm or "liquidación" in norm:
            seccion_actual = "LIQUIDACION"
            continue

        if seccion_actual == "VENTAS":
            if _es_fila_total(texto):
                continue
            ali = _detectar_alicuota_en_fila(texto)
            if ali in ("10.5", "21"):
                monto = _extraer_monto_fila(row)
                if monto > 0:
                    key = "ventas_105" if ali == "10.5" else "ventas_21"
                    montos[key] += monto
        elif seccion_actual == "COMPRAS":
            if _es_fila_total(texto):
                continue
            ali = _detectar_alicuota_en_fila(texto)
            if ali in ("10.5", "21"):
                monto = _extraer_monto_fila(row)
                if monto > 0:
                    key = "compras_105" if ali == "10.5" else "compras_21"
                    montos[key] += monto

    return montos


def _aplicar_loop_review_partida_doble(
    lineas_asiento: list[dict],
) -> tuple[list[dict], float, bool]:
    total_debe = sum(f["Debe"] for f in lineas_asiento)
    total_haber = sum(f["Haber"] for f in lineas_asiento)
    diferencia = round(total_debe - total_haber, 2)
    dif_antes_ajuste = diferencia
    ajuste_aplicado = False
    if diferencia != 0.0 and abs(diferencia) <= 5.0:
        candidatos = [f for f in lineas_asiento if f["Debe"] > 0 and _es_linea_ventas_iva(f["Detalle"])]
        if candidatos:
            objetivo = max(candidatos, key=lambda f: f["Debe"])
            objetivo["Debe"] = round(objetivo["Debe"] - diferencia, 2)
            ajuste_aplicado = True
    return lineas_asiento, dif_antes_ajuste, ajuste_aplicado


def _texto_a_df_liquidacion(texto: str) -> pd.DataFrame:
    lineas = [l.strip() for l in texto.replace("\r", "").split("\n") if l.strip()]
    return pd.DataFrame({0: lineas}).fillna("")


def _generar_asiento_cierre_iva_desde_df(
    df_liquidacion: pd.DataFrame,
    plan_cuentas_df: pd.DataFrame,
) -> tuple[list[dict], float, bool]:
    """
    Asiento de refundición mensual segregado por alícuota desde Excel/CSV ARCA.
    Ventas → DEBE | Compras → HABER | NC compras → DEBE | Liquidación → HABER.
    Recorrido fila a fila (sin slices) para evitar TypeError con índices no enteros.
    """
    df = df_liquidacion.fillna("").copy()
    lineas_asiento: list[dict] = []
    col_desc = _plan_col_descripcion(plan_cuentas_df)
    col_cod = _plan_col_codigo(plan_cuentas_df)
    plan_base = plan_cuentas_df
    if "imputable" in plan_cuentas_df.columns:
        plan_base = plan_cuentas_df[plan_cuentas_df["imputable"] == True].copy()
        if plan_base.empty:
            plan_base = plan_cuentas_df

    seccion_actual: str | None = None
    debito_por_ali: dict[str, float] = {"10.5": 0.0, "21": 0.0, "27": 0.0}

    for _, row in df.iterrows():
        texto = " ".join(str(v) for v in row if pd.notna(v) and str(v).strip())
        norm = _norm_fila(texto)

        if "debito fiscal actividades" in norm:
            seccion_actual = "VENTAS"
            continue
        if "credito fiscal actividades" in norm or "crédito fiscal actividades" in norm:
            seccion_actual = "COMPRAS"
            continue
        if "liquidacion" in norm or "liquidación" in norm:
            seccion_actual = "LIQUIDACION"
            continue

        if seccion_actual == "VENTAS":
            if _es_fila_total(texto):
                continue
            ali = _detectar_alicuota_en_fila(texto)
            if ali:
                monto = _extraer_monto_fila(row)
                if monto > 0:
                    debito_por_ali[ali] = debito_por_ali.get(ali, 0.0) + monto

        elif seccion_actual == "COMPRAS":
            if _es_fila_total(texto):
                continue
            if "credito fiscal a restituir" in norm:
                monto_nc = _extraer_monto_fila(row)
                if monto_nc > 0:
                    cod, desc = _buscar_cuenta_iva_alias(plan_base, col_desc, col_cod, "nc_credito")
                    lineas_asiento.append({
                        "Cuenta": cod, "Detalle": desc, "Debe": monto_nc, "Haber": 0.0, "Estado": "Ingresado",
                    })
                continue
            ali = _detectar_alicuota_en_fila(texto)
            if ali:
                monto_c = _extraer_monto_fila(row)
                if monto_c > 0:
                    cod, desc = _buscar_cuenta_iva_alias(plan_base, col_desc, col_cod, "compras", ali)
                    lineas_asiento.append({
                        "Cuenta": cod, "Detalle": desc, "Debe": 0.0, "Haber": monto_c, "Estado": "Ingresado",
                    })

        elif seccion_actual == "LIQUIDACION":
            if "percepciones impositivas sufridas" in norm:
                monto_percep = _extraer_monto_fila(row)
                if monto_percep > 0:
                    cod, desc = _buscar_cuenta_iva_alias(plan_base, col_desc, col_cod, "percepciones")
                    lineas_asiento.append({
                        "Cuenta": cod, "Detalle": desc, "Debe": 0.0, "Haber": monto_percep, "Estado": "Ingresado",
                    })
            if any(
                x in texto
                for x in ("Saldo del Impuesto a Favor de ARCA", "Saldo de impuesto a favor de AFIP")
            ):
                monto_pagar = _extraer_monto_fila(row)
                if monto_pagar > 0:
                    cod, desc = _buscar_cuenta_iva_alias(plan_base, col_desc, col_cod, "pagar")
                    lineas_asiento.append({
                        "Cuenta": cod, "Detalle": desc, "Debe": 0.0, "Haber": monto_pagar, "Estado": "Ingresado",
                    })

    for ali, monto in debito_por_ali.items():
        if monto > 0:
            cod, desc = _buscar_cuenta_iva_alias(plan_base, col_desc, col_cod, "ventas", ali)
            lineas_asiento.append({
                "Cuenta": cod, "Detalle": desc, "Debe": monto, "Haber": 0.0, "Estado": "Ingresado",
            })

    return _aplicar_loop_review_partida_doble(lineas_asiento)


def _generar_asiento_cierre_iva_segregado(
    texto_pdf: str,
    plan_cuentas_df: pd.DataFrame,
) -> tuple[list[dict], float, bool]:
    """Wrapper de compatibilidad: convierte texto plano a DataFrame y delega al motor Pandas."""
    return _generar_asiento_cierre_iva_desde_df(_texto_a_df_liquidacion(texto_pdf), plan_cuentas_df)


def _extraer_periodo_desde_df(df: pd.DataFrame) -> tuple[int, int] | None:
    textos = [_texto_fila(df, idx) for idx in list(df.index[:30])]
    return _extraer_periodo("\n".join(textos))


def _leer_liquidacion_iva(uploaded_file) -> pd.DataFrame:
    ext = Path(uploaded_file.name).suffix.lower()
    buf = BytesIO(uploaded_file.getvalue())
    if ext == ".csv":
        df = pd.read_csv(buf, header=None, dtype=str)
    else:
        df = pd.read_excel(buf, header=None, dtype=str)
    return df.fillna("")


def _procesar_liquidacion_iva(uploaded_file, plan_df: pd.DataFrame) -> list[AsientoDevengamiento]:
    df = _leer_liquidacion_iva(uploaded_file)
    periodo = _extraer_periodo_desde_df(df)
    if periodo:
        mes, anio = periodo
    else:
        hoy = pd.Timestamp.today()
        mes, anio = hoy.month, hoy.year
    lineas, _, _ = _generar_asiento_cierre_iva_desde_df(df, plan_df)
    if not lineas:
        return []
    return [_lineas_a_asiento(lineas, plan_df, mes, anio, identificador=1)]


def procesar_liquidacion_arca(
    uploaded_file,
    plan_df: pd.DataFrame,
    sal_tec: float,
    sal_lib: float,
    ret_mes: float,
    perc_mes: float,
) -> list[AsientoDevengamiento]:
    """Motor unificado: PDF/XLSX/CSV ARCA + inputs manuales → asiento Tango."""
    ext = Path(uploaded_file.name).suffix.lower()

    if ext == ".pdf":
        texto = _extraer_texto_pdf_upload(uploaded_file)
        df = _texto_a_df_liquidacion(texto)
        periodo = _extraer_periodo(texto)
    else:
        buf = BytesIO(uploaded_file.getvalue())
        if ext == ".csv":
            df = pd.read_csv(buf, header=None)
        else:
            df = pd.read_excel(buf, header=None)
        periodo = _extraer_periodo_desde_df(df)

    montos = _parsear_montos_fiscales_arca_desde_df(df)
    lineas, _, _, _ = _generar_asiento_iva_hibrido_excel(
        plan_df, montos, sal_tec, sal_lib, ret_mes, perc_mes,
    )
    if not lineas:
        return []

    if periodo:
        mes, anio = periodo
    else:
        hoy = pd.Timestamp.today()
        mes, anio = hoy.month, hoy.year
    return [_lineas_a_asiento(lineas, plan_df, mes, anio, identificador=1)]


def _extraer_periodo(texto: str) -> tuple[int, int] | None:
    meses_abr = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
        "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    }
    meses_nombre = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
        "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10,
        "noviembre": 11, "diciembre": 12,
    }
    mp = re.search(r"\b(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)-(\d{2})\b", texto, re.I)
    if mp:
        yy = int(mp.group(2))
        return meses_abr[mp.group(1).lower()], (2000 + yy if yy < 50 else 1900 + yy)
    for fuente in (texto[:300].lower(), texto.lower()):
        m = re.search(r"\b(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[.\-/\s](\d{2})\b", fuente)
        if m:
            yy = int(m.group(2))
            return meses_abr[m.group(1)], (2000 + yy if yy < 50 else 1900 + yy)
        m = re.search(r"\b(0[1-9]|1[0-2])[/\-](20\d{2})\b", fuente)
        if m:
            return int(m.group(1)), int(m.group(2))
        for nombre, num in meses_nombre.items():
            m = re.search(rf"\b{nombre}\b[^\d]{{0,30}}\b(20\d{{2}})\b", fuente)
            if m:
                return num, int(m.group(1))
    return None


def _nombre_banco_display_desde_ficha(ficha: dict | None) -> str:
    """Nombre legible del banco (ej. 'Banco Galicia') desde la ficha BANK_REGISTRY."""
    if not ficha:
        return "Banco"
    slug = str(ficha.get("slug") or "").strip().lower()
    codigo = str(ficha.get("codigo_tango") or "").strip().upper()
    for nombre, reg in BANK_REGISTRY.items():
        if slug and str(reg.get("slug") or "").lower() == slug:
            return str(nombre)
        if codigo and str(reg.get("codigo_tango") or "").upper() == codigo:
            return str(nombre)
    return "Banco"


def _es_tipo_asiento_banco(tipo_impuesto: str) -> bool:
    tipo = str(tipo_impuesto or "").strip().upper()
    if tipo in {"BANCO", "VARIOS"}:
        return tipo == "BANCO"
    for ficha in BANK_REGISTRY.values():
        if str(ficha.get("codigo_tango") or "").upper() == tipo:
            return True
        if str(ficha.get("slug") or "").upper() == tipo:
            return True
    return False


def _lineas_a_asiento(
    lineas: list[dict],
    plan_df: pd.DataFrame,
    mes: int,
    anio: int,
    identificador: int = 1,
    tipo_impuesto: str = "IVA",
    *,
    concepto_override: str | None = None,
    nombre_banco: str | None = None,
) -> AsientoDevengamiento:
    periodo_str = f"{mes:02d}/{anio}"
    # Leyenda siempre vacía: Tango no necesita texto repetido en cada renglón.
    leyenda = ""
    if concepto_override:
        concepto = str(concepto_override).strip()
    elif _es_tipo_asiento_banco(tipo_impuesto) or nombre_banco:
        banco = (nombre_banco or "Banco").strip() or "Banco"
        concepto = f"{banco} {periodo_str}"
    elif tipo_impuesto == "CM":
        concepto = f"DEVENGAMIENTO CM {periodo_str}"
    elif tipo_impuesto == "IIBB":
        concepto = f"Devengamiento IIBB {periodo_str}"
    elif tipo_impuesto == "SUELDOS":
        concepto = f"Devengamiento Sueldos {periodo_str}"
    elif tipo_impuesto == "TISH":
        concepto = f"Devengamiento TISH {periodo_str}"
    else:
        concepto = f"Devengamiento IVA {periodo_str}"
    renglones = []
    advertencias: list[str] = []

    for fila in lineas:
        cod = fila["Cuenta"]
        desc = fila["Detalle"]
        if cod == "99999" or str(desc).startswith("CUENTA_NO_MAPEADA"):
            advertencias.append(f"Sin cuenta en plan para: {desc}")
        _agregar_renglon(
            renglones,
            cod,
            desc,
            debe=fila["Debe"],
            haber=fila["Haber"],
            leyenda=leyenda,
        )

    asiento = AsientoDevengamiento(
        identificador=identificador,
        concepto=concepto,
        fecha=_fecha_asiento_iva_tango(mes, anio),
        renglones=renglones,
    )
    asiento.tipo = tipo_impuesto  # type: ignore[attr-defined]
    asiento.periodo = periodo_str  # type: ignore[attr-defined]
    asiento.advertencias = advertencias  # type: ignore[attr-defined]
    return asiento


def _opciones_cuentas_plan(plan_df: pd.DataFrame) -> list[tuple[str, str]]:
    """Plan completo de la sociedad activa: (codigo, 'codigo - descripcion'). Cacheado por mtime."""
    if plan_df is None or plan_df.empty:
        return []
    sid = st.session_state.get("plan_cuentas_cliente_id")
    mtime = st.session_state.get("plan_cuentas_mtime")
    path = st.session_state.get("plan_cuentas_path_resuelto")
    cache_key = f"plan_opciones_cache_{sid}_{mtime}_{path}"
    cached = st.session_state.get(cache_key)
    if isinstance(cached, list) and cached:
        return cached

    col_cod = _plan_col_codigo(plan_df)
    col_desc = _plan_col_descripcion(plan_df)
    vistos: set[str] = set()
    opciones: list[tuple[str, str]] = []
    for _, row in plan_df.iterrows():
        cod = str(row[col_cod]).strip()
        if not cod or cod.lower() in {"nan", "none", "codigo", "código"}:
            continue
        if cod in vistos:
            continue
        vistos.add(cod)
        desc = str(row[col_desc]).strip()
        if desc.lower() in {"nan", "none"}:
            desc = ""
        opciones.append((cod, f"{cod} - {desc}" if desc else cod))
    opciones.sort(key=lambda x: (len(x[0]), x[0], x[1].lower()))
    st.session_state[cache_key] = opciones
    return opciones


def _opciones_con_cuenta_actual(
    opciones: list[tuple[str, str]],
    codigo_actual: str = "",
    desc_actual: str = "",
) -> list[tuple[str, str]]:
    """Si la cuenta de la fila no está en el plan, la agrega al menú para no perderla."""
    cod = str(codigo_actual or "").strip()
    desc = str(desc_actual or "").strip()
    if not cod or cod == "99999":
        return list(opciones)
    if _indice_opcion_cuenta(opciones, cod) is not None:
        return list(opciones)
    label = f"{cod} - {desc}" if desc else cod
    return [(cod, f"{label} (fuera del plan)")] + list(opciones)


def _opciones_desde_filas_grilla(rows: list[dict]) -> list[tuple[str, str]]:
    """Fallback: armar menú con las cuentas ya presentes en la grilla."""
    vistos: set[str] = set()
    opciones: list[tuple[str, str]] = []
    for r in rows or []:
        cod = str(r.get("Código") or "").strip()
        if not cod or cod in vistos or cod == "99999":
            continue
        vistos.add(cod)
        desc = str(r.get("Descripción") or "").strip()
        opciones.append((cod, f"{cod} - {desc}" if desc else cod))
    return opciones


def _resolver_plan_y_opciones_cliente(
    plan_df: pd.DataFrame | None = None,
    *,
    sociedad_id: int | None = None,
    codigo_actual: str = "",
    desc_actual: str = "",
    rows_grilla: list[dict] | None = None,
) -> tuple[pd.DataFrame | None, list[tuple[str, str]], str]:
    """
    Asegura el plan del cliente activo en session y arma el menú desplegable.

    Retorna (plan_df, opciones[(codigo, label)], mensaje_estado).
    """
    sid = sociedad_id or st.session_state.get("sociedad_activa")
    nombre_cliente = ""
    if sid is not None:
        try:
            sid_int = int(sid)
            plan_ss = st.session_state.get("plan_cuentas_df")
            forzar = plan_ss is None or getattr(plan_ss, "empty", True)
            _sincronizar_plan_cuentas_session(sid_int, forzar=forzar)
            cli = db.obtener_cliente(sid_int)
            if cli:
                nombre_cliente = str(cli.get("nombre") or "").strip()
        except Exception:
            pass

    plan = st.session_state.get("plan_cuentas_df")
    if plan is None or getattr(plan, "empty", True):
        plan = plan_df
    opciones = _opciones_cuentas_plan(plan)
    if not opciones and rows_grilla:
        opciones = _opciones_desde_filas_grilla(rows_grilla)
    opciones = _opciones_con_cuenta_actual(opciones, codigo_actual, desc_actual)

    if opciones and plan is not None and not getattr(plan, "empty", True):
        quien = f" de **{nombre_cliente}**" if nombre_cliente else " del cliente activo"
        msg = f"Menú de cuentas{quien}: **{len(opciones)}** cuentas del plan cargado."
    elif opciones:
        msg = (
            f"Menú parcial (**{len(opciones)}** cuentas de la grilla). "
            "Vinculá el plan del cliente para ver todas las cuentas Tango."
        )
    else:
        msg = (
            "No hay plan de cuentas cargado para este cliente. "
            "Vinculá el Excel del plan en Gestión de Clientes para poder elegir cuenta."
        )
    return plan, opciones, msg


def _indice_opcion_cuenta(opciones: list[tuple[str, str]], codigo: str) -> int | None:
    codigo = str(codigo).strip()
    if not codigo or codigo == "99999":
        return None
    for i, (cod, _) in enumerate(opciones):
        if cod == codigo:
            return i
    return None


def _inicializar_listas_saldos_iva() -> None:
    rc = _iva_reset_counter()
    if st.session_state.get("iva_saldos_rc_active") != rc:
        st.session_state.saldos_tecnicos_list = [0.0]
        st.session_state.saldos_libre_list = [0.0]
        st.session_state.iva_saldos_rc_active = rc


def _render_lista_saldos_iva(titulo: str, lista_key: str, prefix: str) -> float:
    """Inputs dinámicos con botón ➕; retorna la suma total."""
    _inicializar_listas_saldos_iva()
    lista: list[float] = st.session_state[lista_key]
    rc = _iva_reset_counter()
    total = 0.0
    for i in range(len(lista)):
        ca, cb = st.columns([4, 1])
        with ca:
            etiqueta = titulo if i == 0 else f"{titulo} #{i + 1}"
            total += st.number_input(
                etiqueta,
                min_value=0.0,
                step=0.01,
                format="%.2f",
                key=f"{prefix}_{rc}_{i}",
            )
        with cb:
            st.write("")
            if i == len(lista) - 1 and st.button(
                "➕",
                key=f"{prefix}_add_{rc}_{i}",
                help=f"Agregar otro {titulo.lower()}",
            ):
                lista.append(0.0)
                st.session_state[lista_key] = lista
                st.rerun()
    return total


def _tipo_default_por_rol(rol: str) -> str:
    """Debe: débitos y NC compras; Haber: créditos, NC ventas, ret/perc/saldos/cierre."""
    if rol in ("ventas_21", "ventas_105", "ventas_27", "nc_compras", "nc_compras_27", "saldo_favor"):
        return "Debe"
    return "Haber"


def _monto_neto_fila(row: dict) -> float:
    return monto_neto_fila_grilla(row)


def _aplicar_tipo_a_fila_dict(row: dict, tipo: str) -> None:
    monto = _monto_neto_fila(row)
    aplicar_monto_editable_fila(row, monto, tipo)


def _invertir_tipo_debe_haber(tipo: str | None) -> str:
    """Debe ↔ Haber (botón ⇄ de la grilla)."""
    return "Haber" if str(tipo or "Debe").strip() == "Debe" else "Debe"


def _render_selector_tipo_con_swap(
    col,
    *,
    tipo_actual: str,
    key_tipo: str,
    key_swap: str,
) -> tuple[str, bool]:
    """Botón ⇄ que muestra el lado actual y lo invierte al hacer clic."""
    del key_tipo  # compat: ya no usamos selectbox anidado (rompía el layout)
    tipo_opciones = ["Debe", "Haber"]
    tipo_norm = tipo_actual if tipo_actual in tipo_opciones else "Debe"
    pidio_swap = col.button(
        f"⇄ {tipo_norm}",
        key=key_swap,
        help="Cambiar Debe ↔ Haber",
        use_container_width=True,
    )
    if pidio_swap:
        return _invertir_tipo_debe_haber(tipo_norm), True
    return tipo_norm, False


def _sincronizar_widget_monto(row: dict, monto_key: str) -> None:
    """Inicializa el number_input; no pisa el valor que acaba de tipear el usuario.

    Sobrescribir siempre (widget ≠ fila) impedía corregir un 0: Streamlit guarda el
    nuevo importe en session_state y este sync lo volvía a 0 antes del number_input.
    Regeneración/cierre invalidan las keys vía reset_counter / _sync_monto.
    """
    neto = _monto_neto_fila(row)
    if monto_key not in st.session_state:
        st.session_state[monto_key] = neto
        return
    if row.pop("_sync_monto", False):
        st.session_state[monto_key] = neto


def _aplicar_monto_fila_slug(
    slug: str,
    ficha: dict,
    idx: int,
    monto: float,
    *,
    finalizar: bool = True,
) -> None:
    grilla = st.session_state.get(f"{slug}_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    row = grilla[idx]
    aplicar_monto_editable_fila(row, monto, str(row.get("_tipo") or "Debe"))
    row["_monto_manual"] = True
    st.session_state[f"{slug}_grilla_preview"] = grilla
    _marcar_grilla_dirty(slug)
    if finalizar:
        _finalizar_balance_grilla_slug(slug, ficha)


def _render_importe_editable_en_grilla(
    cols,
    *,
    slug: str,
    idx: int,
    row: dict,
    rc: int,
    tipo: str,
    ficha: dict,
    monto_key: str | None = None,
    finalizar: bool = True,
) -> bool:
    """Pinta number_input en Debe o Haber según tipo. Retorna True si el monto cambió."""
    key = monto_key or _wkey_monto_editable(slug, rc, idx, tipo)
    _sincronizar_widget_monto(row, key)
    neto_previo = _monto_neto_fila(row)

    if tipo == "Debe":
        nuevo = cols[4].number_input(
            "Debe",
            min_value=0.0,
            step=0.01,
            format="%.2f",
            key=key,
            label_visibility="collapsed",
        )
        cols[5].write("—")
    else:
        cols[4].write("—")
        nuevo = cols[5].number_input(
            "Haber",
            min_value=0.0,
            step=0.01,
            format="%.2f",
            key=key,
            label_visibility="collapsed",
        )

    if round(float(nuevo), 2) != round(neto_previo, 2):
        _aplicar_monto_fila_slug(slug, ficha, idx, float(nuevo), finalizar=finalizar)
        return True
    return False


def _aplicar_monto_fila_iva(idx: int, monto: float) -> None:
    ficha = {"motor": "iva", "cuenta_ajuste_centavos_rol": "ventas_21"}
    _aplicar_monto_fila_slug("iva", ficha, idx, monto)


def _aplicar_monto_fila_iibb(idx: int, monto: float) -> None:
    ficha = {"motor": "iibb", "cuenta_ajuste_centavos_rol": "impuesto_determinado"}
    _aplicar_monto_fila_slug("iibb", ficha, idx, monto)


def _sincronizar_asientos_desde_grilla(rows: list[dict]) -> None:
    """Propaga grilla (cuentas, debe/haber) a asientos para export Tango."""
    asientos = st.session_state.get("iva_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx >= len(rows):
                return
            r = rows[flat_idx]
            renglon.codigo_cuenta = str(r.get("Código", ""))
            renglon.descripcion_cuenta = str(r.get("Descripción", ""))
            renglon.debe = round(float(r.get("Debe") or 0), 2)
            renglon.haber = round(float(r.get("Haber") or 0), 2)
            flat_idx += 1


def _aplicar_loop_review_filas_grilla(rows: list[dict]) -> list[dict]:
    """Loop review ≤ $5 sobre IVA Débito 21% en filas de grilla."""
    if not rows:
        return rows
    lineas = [
        {
            "Debe": float(r.get("Debe") or 0),
            "Haber": float(r.get("Haber") or 0),
            "_rol": r.get("_rol", ""),
            "Detalle": r.get("Descripción", ""),
        }
        for r in rows
    ]
    lineas, _, _ = _aplicar_loop_review_ventas_21(lineas)
    for i, lin in enumerate(lineas):
        rows[i]["Debe"] = lin["Debe"]
        rows[i]["Haber"] = lin["Haber"]
        rows[i]["_monto"] = _monto_neto_fila(rows[i])
    return rows


def _finalizar_balance_grilla_iva() -> None:
    rows = st.session_state.get("iva_grilla_preview")
    if not rows:
        return
    rows = _aplicar_loop_review_filas_grilla(list(rows))
    st.session_state["iva_grilla_preview"] = rows
    _sincronizar_asientos_desde_grilla(rows)


def _aplicar_tipo_fila_iva(idx: int, tipo: str) -> None:
    grilla = st.session_state.get("iva_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    _aplicar_tipo_a_fila_dict(grilla[idx], tipo)
    st.session_state["iva_grilla_preview"] = grilla
    _finalizar_balance_grilla_iva()


def _fila_grilla_desde_renglon(
    asiento_periodo: str,
    fecha_str: str,
    renglon,
    rol: str,
) -> dict:
    debe = round(float(renglon.debe or 0), 2)
    haber = round(float(renglon.haber or 0), 2)
    tipo = _tipo_default_por_rol(rol)
    if debe > 0 and haber == 0:
        tipo = "Debe"
    elif haber > 0 and debe == 0:
        tipo = "Haber"
    return {
        "Período": formatear_periodo_mm_yyyy(asiento_periodo),
        "Fecha": formatear_fecha_dd_mm_yyyy(fecha_str),
        "Código": renglon.codigo_cuenta,
        "Descripción": renglon.descripcion_cuenta,
        "Debe": debe,
        "Haber": haber,
        "Estado": "Ingresado",
        "_rol": rol,
        "_tipo": tipo,
        "_monto": max(debe, haber),
    }


def _aplicar_seleccion_manual_cuenta_iva(idx: int, codigo: str, descripcion: str) -> None:
    """Sincroniza cuenta seleccionada manualmente en grilla y asientos para export Tango."""
    grilla = st.session_state.get("iva_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    if grilla[idx].get("Código") == codigo and grilla[idx].get("Descripción") == descripcion:
        return

    grilla[idx]["Código"] = codigo
    grilla[idx]["Descripción"] = descripcion
    grilla[idx]["_cuenta_manual"] = True
    st.session_state["iva_grilla_preview"] = grilla

    asientos = st.session_state.get("iva_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx == idx:
                renglon.codigo_cuenta = codigo
                renglon.descripcion_cuenta = descripcion
                advs = getattr(asiento, "advertencias", None) or []
                asiento.advertencias = [  # type: ignore[attr-defined]
                    a for a in advs if not str(a).startswith("Sin cuenta en plan para:")
                ]
                return
            flat_idx += 1


def _etiqueta_concepto_fila_iva(row: dict) -> str:
    """Nombre legible del concepto contable para alertas de auditoría."""
    rol = str(row.get("_rol") or "").strip()
    etiquetas_rol = {
        "ventas_21": "IVA Débito Fiscal 21%",
        "ventas_105": "IVA Débito Fiscal 10,5%",
        "ventas_27": "IVA Débito Fiscal 27%",
        "compras_21": "IVA Crédito Fiscal 21%",
        "compras_105": "IVA Crédito Fiscal 10,5%",
        "compras_27": "IVA Crédito Fiscal 27%",
        "nc_compras": "NC Compras IVA",
        "nc_compras_27": "NC Compras IVA 27%",
        "nc_ventas": "NC Ventas IVA",
        "nc_ventas_27": "NC Ventas IVA 27%",
        "retenciones": "Retenciones IVA",
        "percepciones": "Percepciones IVA",
        "tecnico": "Saldo Técnico IVA Período Anterior",
        "libre": "Saldo Libre Disponibilidad IVA",
        "pagar": "IVA a Pagar",
        "saldo_favor": "Saldo a Favor IVA Nuevo Período",
    }
    if rol in etiquetas_rol:
        return etiquetas_rol[rol]
    if rol:
        return _nombre_concepto_rescate(_rol_a_tipo_concepto(rol))
    desc = str(row.get("Descripción", "")).strip()
    if desc and not desc.startswith("CUENTA_NO_MAPEADA"):
        return desc
    return "Concepto IVA sin identificar"


def _filas_iva_pendientes_vinculacion(rows: list[dict]) -> list[dict]:
    return [
        r for r in rows
        if _es_cuenta_iva_no_mapeada(r.get("Código", ""), r.get("Descripción", ""))
    ]


def _conceptos_iva_sin_vincular(rows: list[dict]) -> list[str]:
    vistos: list[str] = []
    for fila in _filas_iva_pendientes_vinculacion(rows):
        etiqueta = _etiqueta_concepto_fila_iva(fila)
        if etiqueta not in vistos:
            vistos.append(etiqueta)
    return vistos


def _mostrar_alerta_auditoria_cuentas_iva(rows: list[dict]) -> bool:
    """Alerta temprana antes de la grilla. Retorna True si hay cuentas pendientes."""
    conceptos = _conceptos_iva_sin_vincular(rows)
    if not conceptos:
        return False
    lista = "\n".join(f"- {c}" for c in conceptos)
    st.warning(
        "⚠️ Atención: Se detectaron conceptos que no pudieron vincularse automáticamente "
        "con el Plan de Cuentas de Tango. Por favor, asígnales una cuenta manualmente "
        f"en la grilla inferior antes de exportar.\n\n{lista}"
    )
    return True


def _puede_exportar_asiento_iva_tango(rows: list[dict]) -> bool:
    return len(_filas_iva_pendientes_vinculacion(rows)) == 0


def _eliminar_fila_grilla_iva(idx: int) -> None:
    """Elimina una línea del asiento en curso y recalcula el balance."""
    grilla = st.session_state.get("iva_grilla_preview") or []
    if idx < 0 or idx >= len(grilla):
        return
    grilla.pop(idx)
    st.session_state["iva_grilla_preview"] = grilla

    asientos = st.session_state.get("iva_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        nuevos_renglones = []
        nuevos_roles: list[str] = []
        roles = getattr(asiento, "_roles_renglones", []) or []
        for ri, renglon in enumerate(asiento.renglones):
            if flat_idx == idx:
                flat_idx += 1
                continue
            nuevos_renglones.append(renglon)
            if ri < len(roles):
                nuevos_roles.append(roles[ri])
            flat_idx += 1
        asiento.renglones = nuevos_renglones
        if hasattr(asiento, "_roles_renglones"):
            asiento._roles_renglones = nuevos_roles  # type: ignore[attr-defined]

    st.session_state["iva_asientos_generados"] = asientos
    _finalizar_balance_grilla_iva()
    st.rerun()


def _render_grilla_iva_con_selectores(plan_df: pd.DataFrame) -> None:
    """Grilla IVA interactiva: selectbox con plan completo en cada fila."""
    rows = st.session_state.get("iva_grilla_preview") or []
    if not rows:
        return

    plan_session = st.session_state.get("plan_cuentas_df")
    plan_completo = plan_session if plan_session is not None else plan_df
    opciones_plan = _opciones_cuentas_plan(plan_completo)
    sin_plan = not opciones_plan
    if sin_plan:
        st.warning(
            "No hay plan de cuentas disponible para elegir cuenta Tango. "
            "Igual podés cambiar Debe ↔ Haber con el botón ⇄."
        )

    labels = [label for _, label in opciones_plan]
    cod_por_label = {label: cod for cod, label in opciones_plan}
    desc_por_cod = {
        cod: (label.split(" - ", 1)[1] if " - " in label else label)
        for cod, label in opciones_plan
    }

    hdr = st.columns([1, 1.2, 2.9, 0.8, 1, 1, 0.75, 0.55])
    for col, titulo in zip(
        hdr,
        ["Período", "Fecha", "Cuenta (plan completo)", "Tipo ⇄", "Debe", "Haber", "Estado", ""],
    ):
        col.markdown(f"**{titulo}**")

    rc = _iva_reset_counter()
    ficha_iva = {"motor": "iva", "cuenta_ajuste_centavos_rol": "ventas_21"}
    for idx, row in enumerate(rows):
        c = st.columns([1, 1.2, 2.9, 0.8, 1, 1, 0.75, 0.55])
        c[0].write(_texto_periodo_grilla(row.get("Período", "")))
        c[1].write(_texto_fecha_grilla(row.get("Fecha", "")))

        cod_actual = str(row.get("Código", ""))
        desc_actual = str(row.get("Descripción", ""))
        if sin_plan:
            c[2].write(f"{cod_actual} — {desc_actual}" if cod_actual else desc_actual)
        else:
            sel_index = _indice_opcion_cuenta(opciones_plan, cod_actual)
            select_kwargs: dict = {
                "label": "Cuenta",
                "options": labels,
                "key": f"iva_cuenta_{rc}_{idx}",
                "label_visibility": "collapsed",
            }
            if sel_index is not None:
                select_kwargs["index"] = sel_index
            else:
                select_kwargs["index"] = None
                select_kwargs["placeholder"] = "❌ SELECCIONAR CUENTA TANGO..."

            sel_label = c[2].selectbox(**select_kwargs)
            if sel_label:
                cod_sel = cod_por_label[sel_label]
                desc_sel = desc_por_cod.get(cod_sel, sel_label)
                if cod_sel != cod_actual or desc_sel != desc_actual:
                    _aplicar_seleccion_manual_cuenta_iva(idx, cod_sel, desc_sel)
                    row = (st.session_state.get("iva_grilla_preview") or rows)[idx]

        rol = str(row.get("_rol", ""))
        tipo_actual = row.get("_tipo") or _tipo_default_por_rol(rol)
        tipo_sel, pidio_swap = _render_selector_tipo_con_swap(
            c[3],
            tipo_actual=str(tipo_actual),
            key_tipo=f"iva_tipo_{rc}_{idx}",
            key_swap=f"iva_swap_{rc}_{idx}",
        )
        if pidio_swap or tipo_sel != row.get("_tipo"):
            _aplicar_tipo_fila_iva(idx, tipo_sel)
            row = (st.session_state.get("iva_grilla_preview") or rows)[idx]

        if _render_importe_editable_en_grilla(
            c, slug="iva", idx=idx, row=row, rc=rc, tipo=str(tipo_sel), ficha=ficha_iva,
        ):
            row = (st.session_state.get("iva_grilla_preview") or rows)[idx]

        c[6].write(row.get("Estado", ""))
        if c[7].button("🗑️", key=f"iva_del_{rc}_{idx}", help="Eliminar fila"):
            _eliminar_fila_grilla_iva(idx)

    # No st.rerun() acá: el widget ya disparó un rerun; otro rompe el DOM (removeChild).
    _finalizar_balance_grilla_iva()


def _tipo_default_por_rol_iibb(rol: str) -> str:
    if rol in ("impuesto_determinado", "saldo_favor_nuevo"):
        return "Debe"
    return "Haber"


def _obtener_valores_saldos_iibb_desde_session(lista_key: str, prefix: str) -> list[float]:
    rc = _iva_reset_counter()
    lista = st.session_state.get(lista_key) or [0.0]
    return [
        round(float(st.session_state.get(f"{prefix}_{rc}_{i}", 0.0)), 2)
        for i in range(len(lista))
    ]


def _inicializar_listas_saldos_iibb() -> None:
    rc = _iva_reset_counter()
    if st.session_state.get("iibb_saldos_rc_active") != rc:
        st.session_state.saldos_favor_iibb_list = [0.0]
        st.session_state.iibb_saldos_rc_active = rc


def _render_lista_saldos_iibb(titulo: str, lista_key: str, prefix: str) -> float:
    _inicializar_listas_saldos_iibb()
    lista: list[float] = st.session_state[lista_key]
    rc = _iva_reset_counter()
    total = 0.0
    for i in range(len(lista)):
        ca, cb = st.columns([4, 1])
        with ca:
            etiqueta = titulo if i == 0 else f"{titulo} #{i + 1}"
            total += st.number_input(
                etiqueta,
                min_value=0.0,
                step=0.01,
                format="%.2f",
                key=f"{prefix}_{rc}_{i}",
            )
        with cb:
            st.write("")
            if i == len(lista) - 1 and st.button(
                "➕",
                key=f"{prefix}_add_{rc}_{i}",
                help=f"Agregar otro {titulo.lower()}",
            ):
                lista.append(0.0)
                st.session_state[lista_key] = lista
                st.rerun()
    return total


def _render_cuadro_control_analitico_iibb(
    resumen: dict | None,
    rows: list[dict],
) -> None:
    if resumen is None:
        resumen = st.session_state.get("iibb_resumen_analitico") or {}

    total_debe = round(sum(float(r.get("Debe") or 0) for r in rows), 2)
    total_haber = round(sum(float(r.get("Haber") or 0) for r in rows), 2)
    diferencia = round(total_debe - total_haber, 2)

    with st.container(border=True):
        st.markdown("### 📊 Resumen Analítico de la Liquidación IIBB")
        c1, c2, c3 = st.columns(3)
        c1.markdown(
            f"**Impuesto Determinado**  \n{_fmt_pesos_ar(resumen.get('impuesto_determinado', 0))}"
        )
        c2.markdown(
            f"**Total Descuentos (Ret/Perc/Banc/Saldos)**  \n"
            f"{_fmt_pesos_ar(resumen.get('total_descuentos', 0))}"
        )
        c3.markdown(
            f"**Posición Neta**  \n{_fmt_pesos_ar(resumen.get('diferencia_previa', 0))}"
        )

        detalle_saldos: list[str] = []
        for i, m in enumerate(resumen.get("saldos_favor") or [], start=1):
            detalle_saldos.append(f"Saldo a Favor #{i}: {_fmt_pesos_ar(m)}")
        if detalle_saldos:
            for linea in detalle_saldos:
                st.markdown(f"- {linea}")
            st.markdown(
                f"**Subtotal saldos anteriores:** "
                f"{_fmt_pesos_ar(resumen.get('total_saldos_anteriores', 0))}"
            )

        st.markdown("#### 🏁 Resultado Final de la DDJJ")
        r1, r2, r3 = st.columns(3)
        r1.markdown(f"**Concepto imputado**  \n{resumen.get('resultado_tipo', '—')}")
        r2.markdown(f"**Lado contable**  \n{resumen.get('resultado_lado', '—')}")
        r3.markdown(f"**Importe de cierre**  \n{_fmt_pesos_ar(resumen.get('resultado_monto', 0))}")

        st.markdown("---")
        t1, t2, t3 = st.columns(3)
        t1.markdown(f"**Total Debe**  \n{_fmt_pesos_ar(total_debe)}")
        t2.markdown(f"**Total Haber**  \n{_fmt_pesos_ar(total_haber)}")
        color = "green" if diferencia == 0 else "red"
        t3.markdown(f"**Diferencia**  \n:{color}[{_fmt_pesos_ar(diferencia)}]")


def _sincronizar_asientos_desde_grilla_iibb(rows: list[dict]) -> None:
    asientos = st.session_state.get("iibb_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx >= len(rows):
                return
            r = rows[flat_idx]
            renglon.codigo_cuenta = str(r.get("Código", ""))
            renglon.descripcion_cuenta = str(r.get("Descripción", ""))
            renglon.debe = round(float(r.get("Debe") or 0), 2)
            renglon.haber = round(float(r.get("Haber") or 0), 2)
            flat_idx += 1


def _aplicar_loop_review_filas_grilla_iibb(rows: list[dict]) -> list[dict]:
    if not rows:
        return rows
    lineas = [
        {
            "Debe": float(r.get("Debe") or 0),
            "Haber": float(r.get("Haber") or 0),
            "_rol": r.get("_rol", ""),
            "Detalle": r.get("Descripción", ""),
        }
        for r in rows
    ]
    lineas, _, _ = _aplicar_loop_review_impuesto_determinado(lineas)
    for i, lin in enumerate(lineas):
        rows[i]["Debe"] = lin["Debe"]
        rows[i]["Haber"] = lin["Haber"]
        rows[i]["_monto"] = _monto_neto_fila(rows[i])
    return rows


def _finalizar_balance_grilla_iibb() -> None:
    rows = st.session_state.get("iibb_grilla_preview")
    if not rows:
        return
    rows = _aplicar_loop_review_filas_grilla_iibb(list(rows))
    st.session_state["iibb_grilla_preview"] = rows
    _sincronizar_asientos_desde_grilla_iibb(rows)


def _aplicar_tipo_fila_iibb(idx: int, tipo: str) -> None:
    grilla = st.session_state.get("iibb_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    _aplicar_tipo_a_fila_dict(grilla[idx], tipo)
    st.session_state["iibb_grilla_preview"] = grilla
    _finalizar_balance_grilla_iibb()


def _aplicar_seleccion_manual_cuenta_iibb(idx: int, codigo: str, descripcion: str) -> None:
    grilla = st.session_state.get("iibb_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    if grilla[idx].get("Código") == codigo and grilla[idx].get("Descripción") == descripcion:
        return
    grilla[idx]["Código"] = codigo
    grilla[idx]["Descripción"] = descripcion
    grilla[idx]["_cuenta_manual"] = True
    st.session_state["iibb_grilla_preview"] = grilla
    _marcar_grilla_dirty("iibb")
    asientos = st.session_state.get("iibb_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx == idx:
                renglon.codigo_cuenta = codigo
                renglon.descripcion_cuenta = descripcion
                advs = getattr(asiento, "advertencias", None) or []
                asiento.advertencias = [  # type: ignore[attr-defined]
                    a for a in advs if not str(a).startswith("Sin cuenta en plan para:")
                ]
                return
            flat_idx += 1


def _etiqueta_concepto_fila_iibb(row: dict) -> str:
    rol = str(row.get("_rol") or "").strip()
    etiquetas = {
        "impuesto_determinado": "IIBB Devengado / Impuesto Determinado",
        "retenciones_iibb": "Retenciones IIBB",
        "retenciones_bancarias": "Retenciones Bancarias IIBB (Sircreb)",
        "percepciones_iibb": "Percepciones IIBB",
        "saldo_favor_anterior": "Saldo a Favor IIBB Período Anterior",
        "iibb_pagar": "IIBB a Pagar",
        "saldo_favor_nuevo": "Saldo a Favor IIBB Nuevo Período",
    }
    if rol in etiquetas:
        return etiquetas[rol]
    if rol:
        return _nombre_concepto_rescate_iibb(_rol_a_tipo_concepto_iibb(rol))
    desc = str(row.get("Descripción", "")).strip()
    if desc and not desc.startswith("CUENTA_NO_MAPEADA"):
        return desc
    return "Concepto IIBB sin identificar"


def _mostrar_alerta_auditoria_cuentas_iibb(rows: list[dict]) -> bool:
    conceptos: list[str] = []
    for fila in rows:
        if not _es_cuenta_iva_no_mapeada(fila.get("Código", ""), fila.get("Descripción", "")):
            continue
        etiqueta = _etiqueta_concepto_fila_iibb(fila)
        if etiqueta not in conceptos:
            conceptos.append(etiqueta)
    if not conceptos:
        return False
    lista = "\n".join(f"- {c}" for c in conceptos)
    st.warning(
        "⚠️ Atención: Se detectaron conceptos IIBB sin vincular al Plan de Cuentas. "
        f"Asigná cuentas manualmente en la grilla.\n\n{lista}"
    )
    return True


def _puede_exportar_asiento_iibb_tango(rows: list[dict]) -> bool:
    return not any(
        _es_cuenta_iva_no_mapeada(r.get("Código", ""), r.get("Descripción", ""))
        for r in rows
    )


def _eliminar_fila_grilla_iibb(idx: int) -> None:
    grilla = st.session_state.get("iibb_grilla_preview") or []
    if idx < 0 or idx >= len(grilla):
        return
    grilla.pop(idx)
    st.session_state["iibb_grilla_preview"] = grilla
    asientos = st.session_state.get("iibb_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        nuevos_renglones = []
        nuevos_roles: list[str] = []
        roles = getattr(asiento, "_roles_renglones", []) or []
        for ri, renglon in enumerate(asiento.renglones):
            if flat_idx == idx:
                flat_idx += 1
                continue
            nuevos_renglones.append(renglon)
            if ri < len(roles):
                nuevos_roles.append(roles[ri])
            flat_idx += 1
        asiento.renglones = nuevos_renglones
        if hasattr(asiento, "_roles_renglones"):
            asiento._roles_renglones = nuevos_roles  # type: ignore[attr-defined]
    st.session_state["iibb_asientos_generados"] = asientos
    _finalizar_balance_grilla_iibb()
    st.rerun()


def _render_grilla_iibb_con_selectores(plan_df: pd.DataFrame) -> None:
    rows = st.session_state.get("iibb_grilla_preview") or []
    if not rows:
        return

    plan_session = st.session_state.get("plan_cuentas_df")
    plan_completo = plan_session if plan_session is not None else plan_df
    opciones_plan = _opciones_cuentas_plan(plan_completo)
    sin_plan = not opciones_plan
    if sin_plan:
        st.warning(
            "No hay plan de cuentas disponible para elegir cuenta Tango. "
            "Igual podés cambiar Debe ↔ Haber con el botón ⇄."
        )

    labels = [label for _, label in opciones_plan]
    cod_por_label = {label: cod for cod, label in opciones_plan}
    desc_por_cod = {
        cod: (label.split(" - ", 1)[1] if " - " in label else label)
        for cod, label in opciones_plan
    }

    hdr = st.columns([1, 1.2, 2.9, 0.8, 1, 1, 0.75, 0.55])
    for col, titulo in zip(
        hdr,
        ["Período", "Fecha", "Cuenta (plan completo)", "Tipo ⇄", "Debe", "Haber", "Estado", ""],
    ):
        col.markdown(f"**{titulo}**")

    rc = _iva_reset_counter()
    ficha_iibb = {"motor": "iibb", "cuenta_ajuste_centavos_rol": "impuesto_determinado"}
    for idx, row in enumerate(rows):
        c = st.columns([1, 1.2, 2.9, 0.8, 1, 1, 0.75, 0.55])
        c[0].write(_texto_periodo_grilla(row.get("Período", "")))
        c[1].write(_texto_fecha_grilla(row.get("Fecha", "")))

        cod_actual = str(row.get("Código", ""))
        desc_actual = str(row.get("Descripción", ""))
        if sin_plan:
            c[2].write(f"{cod_actual} — {desc_actual}" if cod_actual else desc_actual)
        else:
            sel_index = _indice_opcion_cuenta(opciones_plan, cod_actual)
            select_kwargs: dict = {
                "label": "Cuenta",
                "options": labels,
                "key": f"iibb_cuenta_{rc}_{idx}",
                "label_visibility": "collapsed",
            }
            if sel_index is not None:
                select_kwargs["index"] = sel_index
            else:
                select_kwargs["index"] = None
                select_kwargs["placeholder"] = "❌ SELECCIONAR CUENTA TANGO..."

            sel_label = c[2].selectbox(**select_kwargs)
            if sel_label:
                cod_sel = cod_por_label[sel_label]
                desc_sel = desc_por_cod.get(cod_sel, sel_label)
                if cod_sel != cod_actual or desc_sel != desc_actual:
                    _aplicar_seleccion_manual_cuenta_iibb(idx, cod_sel, desc_sel)
                    row = (st.session_state.get("iibb_grilla_preview") or rows)[idx]

        rol = str(row.get("_rol", ""))
        tipo_actual = row.get("_tipo") or _tipo_default_por_rol_iibb(rol)
        tipo_sel, pidio_swap = _render_selector_tipo_con_swap(
            c[3],
            tipo_actual=str(tipo_actual),
            key_tipo=f"iibb_tipo_{rc}_{idx}",
            key_swap=f"iibb_swap_{rc}_{idx}",
        )
        if pidio_swap or tipo_sel != row.get("_tipo"):
            _aplicar_tipo_fila_iibb(idx, tipo_sel)
            row = (st.session_state.get("iibb_grilla_preview") or rows)[idx]

        if _render_importe_editable_en_grilla(
            c, slug="iibb", idx=idx, row=row, rc=rc, tipo=str(tipo_sel), ficha=ficha_iibb,
        ):
            row = (st.session_state.get("iibb_grilla_preview") or rows)[idx]

        c[6].write(row.get("Estado", ""))
        if c[7].button("🗑️", key=f"iibb_del_{rc}_{idx}", help="Eliminar fila"):
            _eliminar_fila_grilla_iibb(idx)

    # No st.rerun() acá: el widget ya disparó un rerun; otro rompe el DOM (removeChild).
    _finalizar_balance_grilla_iibb()


def _procesar_pdf_iva(ruta_pdf: Path, plan_df: pd.DataFrame) -> list[AsientoDevengamiento]:
    import pdfplumber

    asientos: list[AsientoDevengamiento] = []
    with pdfplumber.open(str(ruta_pdf)) as pdf:
        for idx, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            if not texto.strip():
                continue
            periodo = _extraer_periodo(texto)
            if not periodo:
                continue
            mes, anio = periodo
            lineas, _, _ = _generar_asiento_cierre_iva_segregado(texto, plan_df)
            if lineas:
                asientos.append(_lineas_a_asiento(lineas, plan_df, mes, anio, identificador=idx))
    return asientos


def _widget_subir_plan_inline(
    sociedad_id: int | None,
    cuit_activo: str | None,
    *,
    key_suffix: str,
) -> None:
    """Uploader compacto para vincular el plan cuando falta en la sociedad activa."""
    if sociedad_id is None or not cuit_activo:
        return
    archivo = st.file_uploader(
        "Subir Plan de Cuentas (.xlsx) para esta sociedad",
        type=["xlsx"],
        key=f"uploader_plan_inline_{key_suffix}_{sociedad_id}",
        help="Se guarda en planes_cuentas (red/local) y queda vinculado a este CUIT.",
    )
    if not archivo:
        return
    try:
        xlsx_path = _guardar_plan_cliente_en_disco(str(cuit_activo), archivo.getvalue())
        cli = db.obtener_cliente(int(sociedad_id))
        if cli:
            db.actualizar_cliente(
                cli["id"],
                cli["nombre"],
                cli["cuit"],
                cli["tipo_persona"],
                str(xlsx_path),
                cli.get("mes_cierre_balance", 12),
            )
        _invalidar_cache_plan(int(sociedad_id))
        _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=True)
        st.success(f"Plan vinculado: {xlsx_path.name}")
        st.rerun()
    except Exception as exc:
        st.error(f"No se pudo guardar el plan: {exc}")


def _mensaje_plan_no_vinculado(cliente: dict | None = None) -> str:
    """Texto de alerta: distingue archivo faltante vs nunca vinculado."""
    sociedad_id = (cliente or {}).get("id") or st.session_state.get(_SOCiedad_KEY)
    cli = cliente or (db.obtener_cliente(sociedad_id) if sociedad_id else None)
    if not cli:
        return (
            "⚠️ Esta sociedad no tiene un Plan de Cuentas asociado. "
            "Subí el Excel de Tango en Gestión de Clientes."
        )
    cuit = str(cli.get("cuit") or "").strip()
    ruta_bd = cli.get("plan_cuentas_path")
    if ruta_bd and not Path(ruta_bd).exists():
        # Ruta fantasma en BD: limpiar para no confundir
        try:
            db.actualizar_cliente(
                cli["id"],
                cli["nombre"],
                cli["cuit"],
                cli["tipo_persona"],
                None,
                cli.get("mes_cierre_balance", 12),
            )
        except Exception:
            pass
        return (
            f"⚠️ El plan de **{cli.get('nombre')}** figura en la base pero el archivo "
            f"no está en disco (`{Path(ruta_bd).name}`). "
            "Volvé a subirlo en Gestión de Clientes (CUIT "
            f"`{cuit}`)."
        )
    return (
        f"⚠️ **{cli.get('nombre')}** no tiene Plan de Cuentas propio. "
        "Subí el Excel de Tango en Gestión de Clientes o con el cargador del módulo."
    )


def _limpiar_plan_bd_si_archivo_ausente(cliente: dict) -> dict:
    """Limpia path fantasma o plan_default en BD (evita falso 'vinculado')."""
    raw = (cliente.get("plan_cuentas_path") or "").strip()
    if not raw:
        return cliente
    ruta = Path(raw)
    cuit = str(cliente.get("cuit", "")).strip()
    # Ausente en disco, o genérico/default: no es vínculo real de la sociedad
    debe_limpiar = (not ruta.exists()) or _es_plan_generico_default(ruta)
    if not debe_limpiar:
        return cliente
    try:
        db.actualizar_cliente(
            int(cliente["id"]),
            cliente["nombre"],
            cliente["cuit"],
            cliente["tipo_persona"],
            None,
            cliente.get("mes_cierre_balance", 12),
        )
        _dbg_log("P", "_limpiar_plan_bd_si_archivo_ausente", "cleared_stale_path", {
            "cliente_id": cliente.get("id"),
            "cuit": cuit,
            "old": raw,
            "reason": "missing" if not ruta.exists() else "generic_default",
        })
    except Exception:
        pass
    frescos = db.obtener_cliente(int(cliente["id"]))
    return frescos or {**cliente, "plan_cuentas_path": None}


def _plan_existe_para_cliente(cliente: dict) -> bool:
    """True solo si hay plan propio de ESA sociedad (no plan_default ni genérico)."""
    cliente = _limpiar_plan_bd_si_archivo_ausente(cliente)
    cuit = str(cliente.get("cuit", "")).strip()
    if not cuit:
        return False
    bd_raw = (cliente.get("plan_cuentas_path") or "").strip()
    if bd_raw:
        bd_path = Path(bd_raw)
        # Path BD real existente (nombre libre) cuenta; plan_default ya fue limpiado arriba
        if bd_path.exists() and not _es_plan_generico_default(bd_path):
            return True
    return any(
        p.exists() and _es_plan_propio_cliente(cuit, p)
        for p in _rutas_plan_candidatas(cliente)
    )


def _sociedad_tiene_plan_vinculado_por_session() -> bool:
    """True si el plan Excel de la sociedad activa está en session_state o en disco."""
    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    cuit = st.session_state.get("cuit_activo")
    if sociedad_id is None or not cuit:
        return False
    plan_df = st.session_state.get("plan_cuentas_df")
    if (
        plan_df is not None
        and len(plan_df) > 0
        and st.session_state.get("plan_cuentas_cliente_id") == sociedad_id
        and not bool(st.session_state.get("plan_cuentas_es_default", False))
    ):
        return True
    cliente = db.obtener_cliente(sociedad_id)
    if not cliente:
        return False
    return _plan_existe_para_cliente(cliente)


def _sociedad_tiene_plan_vinculado(cliente: dict) -> bool:
    """True si el Excel del plan existe en disco o está cargado en session_state."""
    plan_df = st.session_state.get("plan_cuentas_df")
    if (
        plan_df is not None
        and len(plan_df) > 0
        and st.session_state.get("plan_cuentas_cliente_id") == cliente["id"]
        and not bool(st.session_state.get("plan_cuentas_es_default", False))
    ):
        return True
    return _plan_existe_para_cliente(cliente)


def _inicializar_sesion_por_impuesto(impuesto: str) -> None:
    slug = _slug_impuesto(impuesto)
    if slug == "iva":
        _inicializar_sesion_iva()
    elif slug == "iibb":
        _inicializar_sesion_iibb()
    elif slug == "cm":
        _inicializar_sesion_cm()
    elif slug == "sueldos":
        _inicializar_sesion_sueldos()
    elif slug == "tish":
        _inicializar_sesion_tish()


def _aplicar_limpieza_formulario_mes_por_impuesto(impuesto: str) -> None:
    slug = _slug_impuesto(impuesto)
    avanzar = st.session_state.pop(f"{slug}_periodo_mensual_avanzar", None)
    if avanzar:
        sid = st.session_state.get(_SOCiedad_KEY)
        if sid is not None:
            st.session_state[_clave_periodo_mensual(slug, sid)] = avanzar
    if slug == "iva":
        _aplicar_limpieza_formulario_mes_iva_si_pendiente()
    elif slug == "iibb":
        _aplicar_limpieza_formulario_mes_iibb_si_pendiente()
    elif not st.session_state.pop(f"{slug}_limpiar_formulario_pendiente", False):
        return
    else:
        _limpiar_grilla_periodo_slug(slug)
        st.session_state["reset_counter"] = _iva_reset_counter() + 1


def _marcar_limpieza_formulario_mes_por_impuesto(impuesto: str) -> None:
    slug = _slug_impuesto(impuesto)
    if slug == "iva":
        _marcar_limpieza_formulario_mes_iva()
    elif slug == "iibb":
        _marcar_limpieza_formulario_mes_iibb()
    else:
        st.session_state[f"{slug}_limpiar_formulario_pendiente"] = True


def _fecha_asiento_seleccionada_slug(slug: str, periodo_mensual: str | None = None) -> date | None:
    rc = _iva_reset_counter()
    if periodo_mensual:
        val = st.session_state.get(_clave_fecha_tango_slug(slug, rc, periodo_mensual))
        if isinstance(val, date):
            return val
    val = st.session_state.get(_clave_fecha_tango_slug(slug, rc))
    return val if isinstance(val, date) else None


def _render_lista_saldos_por_ficha(inp: dict, slug: str) -> None:
    titulo = inp["titulo"]
    lista_key = inp["clave"]
    prefix = inp["prefix"]
    rc = _iva_reset_counter()
    active_key = f"{slug}_saldos_rc_active"
    if st.session_state.get(active_key) != rc:
        st.session_state[lista_key] = [0.0]
        st.session_state[active_key] = rc
    lista: list[float] = st.session_state.get(lista_key) or [0.0]
    for i in range(len(lista)):
        ca, cb = st.columns([4, 1])
        with ca:
            etiqueta = titulo if i == 0 else f"{titulo} #{i + 1}"
            st.number_input(
                etiqueta,
                min_value=0.0,
                step=0.01,
                format="%.2f",
                key=f"{prefix}_{rc}_{i}",
            )
        with cb:
            st.write("")
            if i == len(lista) - 1 and st.button(
                "➕",
                key=f"{prefix}_add_{rc}_{i}",
                help=f"Agregar otro {titulo.lower()}",
            ):
                lista.append(0.0)
                st.session_state[lista_key] = lista
                st.rerun()


def _obtener_valores_saldos_desde_session(lista_key: str, prefix: str) -> list[float]:
    rc = _iva_reset_counter()
    lista = st.session_state.get(lista_key) or [0.0]
    return [
        round(float(st.session_state.get(f"{prefix}_{rc}_{i}", 0.0)), 2)
        for i in range(len(lista))
    ]


def _prefijar_fecha_input_desde_planilla_por_motor(
    archivo,
    sociedad_id: int | None,
    motor: str,
) -> None:
    if motor == "iibb":
        _prefijar_fecha_input_desde_planilla_iibb(archivo, sociedad_id)
    elif motor == "iva":
        _prefijar_fecha_input_desde_planilla(archivo, sociedad_id)
    else:
        slug = motor
        ref = _clave_referencia_planilla(archivo, sociedad_id)
        rc = _iva_reset_counter()
        fecha_key = f"{slug}_fecha_tango_{rc}"
        if ref and st.session_state.get(f"{slug}_planilla_ref_fecha") != ref:
            sugerida = None
            if archivo is not None:
                try:
                    buf, es_csv = _abrir_planilla_iva(archivo)
                    impuesto = "Sueldos" if motor == "sueldos" else "TISH"
                    leer = leer_planilla_sueldos_por_etiquetas if motor == "sueldos" else leer_planilla_tish_por_etiquetas
                    datos = leer(buf, es_csv=es_csv, nombre_solapa_impuesto=impuesto)
                    mes, anio = datos.get("periodo_mes"), datos.get("periodo_anio")
                    if mes is not None and anio is not None:
                        sugerida = _fecha_asiento_iva_tango(int(mes), int(anio))
                    else:
                        periodo = _parsear_periodo_texto(datos.get("periodo_texto", ""))
                        if periodo:
                            sugerida = _fecha_asiento_iva_tango(periodo[0], periodo[1])
                except Exception:
                    sugerida = None
            if sugerida:
                st.session_state[fecha_key] = sugerida
            st.session_state[f"{slug}_planilla_ref_fecha"] = ref
        elif fecha_key not in st.session_state:
            st.session_state[fecha_key] = date.today()


def _procesar_balance_por_motor(
    motor: str,
    archivo,
    plan_df: pd.DataFrame,
    solapa: str,
    ficha: dict,
    inputs_manuales: dict[str, float],
    saldos_contingencia: dict[str, list[float]],
    periodo_mensual: str | None = None,
) -> tuple[list[AsientoDevengamiento], dict]:
    if motor == "iva":
        return procesar_planilla_iva(
            archivo, plan_df,
            saldos_contingencia.get("saldos_tecnicos_list", [0.0]),
            saldos_contingencia.get("saldos_libre_list", [0.0]),
            inputs_manuales.get("retenciones", 0.0),
            inputs_manuales.get("percepciones", 0.0),
            nombre_solapa_impuesto=solapa,
            periodo_mensual=periodo_mensual,
        )
    if motor == "iibb":
        lista_saldos = "saldos_favor_iibb_list"
        for inp in ficha.get("inputs_contingencia") or []:
            clave = inp.get("clave")
            if clave and "favor" in str(clave).lower():
                lista_saldos = clave
                break
        return procesar_planilla_iibb(
            archivo, plan_df,
            saldos_contingencia.get(lista_saldos, [0.0]),
            inputs_manuales.get("retenciones", 0.0),
            inputs_manuales.get("percepciones", 0.0),
            inputs_manuales.get("retenciones_bancarias", 0.0),
            nombre_solapa_impuesto=solapa,
            periodo_mensual=periodo_mensual,
            tipo_impuesto=str(ficha.get("codigo_tango", "IIBB")),
        )
    if motor == "sueldos":
        return procesar_planilla_sueldos(
            archivo, plan_df,
            saldos_contingencia.get("saldos_favor_sueldos_list", [0.0]),
            nombre_solapa_impuesto=solapa,
            periodo_mensual=periodo_mensual,
        )
    if motor == "tish":
        return procesar_planilla_tish(
            archivo, plan_df,
            saldos_contingencia.get("saldos_favor_tish_list", [0.0]),
            nombre_solapa_impuesto=solapa,
            periodo_mensual=periodo_mensual,
        )
    raise ValueError(f"Motor de impuesto no soportado: {motor}")


def _periodo_siguiente(periodo: str) -> str | None:
    """Avanza un mes en formato MM/YYYY."""
    parsed = _parsear_periodo_texto(str(periodo).replace("/", "-"))
    if not parsed:
        return None
    mes, anio = parsed
    if mes >= 12:
        return f"01/{anio + 1}"
    return f"{mes + 1:02d}/{anio}"


def _clave_periodo_mensual(slug: str, sociedad_id: int) -> str:
    return f"{slug}_periodo_mensual_{sociedad_id}"


def _clave_periodo_widget(slug: str, sociedad_id: int) -> str:
    return _clave_periodo_mensual(slug, sociedad_id)


def _leer_datos_balance_periodo(
    archivo,
    motor: str,
    solapa: str,
    periodo: str,
) -> dict:
    buf, es_csv = _abrir_planilla_iva(archivo)
    if motor == "iva":
        return leer_planilla_iva_por_etiquetas(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa, periodo_mensual=periodo,
        )
    if motor == "iibb":
        return leer_planilla_iibb_por_etiquetas(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa, periodo_mensual=periodo,
        )
    if motor == "sueldos":
        return leer_planilla_sueldos_por_etiquetas(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa, periodo_mensual=periodo,
        )
    if motor == "tish":
        return leer_planilla_tish_por_etiquetas(
            buf, es_csv=es_csv, nombre_solapa_impuesto=solapa, periodo_mensual=periodo,
        )
    raise ValueError(f"Motor no soportado: {motor}")


def _aplicar_datos_periodo_a_inputs(
    slug: str,
    ficha: dict,
    motor: str,
    datos: dict,
    rc: int,
) -> None:
    """Auto-puebla inputs manuales y saldos de contingencia desde el balance del período."""
    manuales_map: dict[str, str] = {}
    contingencia_map: dict[str, tuple[str, str]] = {}
    if motor == "iva":
        manuales_map = {
            "retenciones": "retenciones_planilla",
            "percepciones": "percepciones_planilla",
        }
        contingencia_map = {
            "saldos_tecnicos_list": ("saldo_tecnico_planilla", "iva_stec"),
            "saldos_libre_list": ("saldo_libre_planilla", "iva_slib"),
        }
    elif motor == "iibb":
        manuales_map = {
            "retenciones": "retenciones_planilla",
            "percepciones": "percepciones_planilla",
            "retenciones_bancarias": "retenciones_bancarias_planilla",
        }
        contingencia_map = {}
        for inp in ficha.get("inputs_contingencia") or []:
            clave = inp.get("clave")
            prefix = inp.get("prefix")
            if clave and prefix:
                contingencia_map[clave] = ("saldo_favor_planilla", prefix)
    elif motor == "sueldos":
        contingencia_map = {
            "saldos_favor_sueldos_list": ("sueldos_saldo_favor", "suel_sfav"),
        }
    elif motor == "tish":
        contingencia_map = {
            "saldos_favor_tish_list": ("tish_saldo_favor", "tish_sfav"),
        }

    for clave_input, clave_datos in manuales_map.items():
        monto = round(float(datos.get(clave_datos, 0.0) or 0.0), 2)
        st.session_state[f"{slug}_{clave_input}_{rc}"] = monto

    for lista_key, (clave_datos, prefix) in contingencia_map.items():
        monto = round(float(datos.get(clave_datos, 0.0) or 0.0), 2)
        if monto > 0:
            st.session_state[lista_key] = [monto]
            st.session_state[f"{slug}_saldos_rc_active"] = rc
            st.session_state[f"{prefix}_{rc}_0"] = monto

    # Fecha asiento: la fija _asegurar_fecha_asiento_ultimo_dia (key por período).


def _descripcion_cuenta_en_plan(plan_df: pd.DataFrame | None, codigo: str) -> str | None:
    if plan_df is None or plan_df.empty or not codigo or codigo == "99999":
        return None
    col_cod = _plan_col_codigo(plan_df)
    col_desc = _plan_col_descripcion(plan_df)
    for _, row in plan_df.iterrows():
        if str(row[col_cod]).strip() == str(codigo).strip():
            return str(row[col_desc]).strip()
    return None


def _marcar_fila_ajuste_loop_review_universal(rows: list[dict], ficha: dict) -> None:
    """Marca la fila principal del impuesto/banco para el ajuste de centavos (Loop Review)."""
    rol = str(ficha.get("cuenta_ajuste_centavos_rol", "") or "")
    if not rol or not rows:
        return
    if ficha.get("motor") == "banco":
        cod_ajuste = str(ficha.get("cuenta_ajuste_codigo", "") or "").strip()
        if cod_ajuste:
            for row in rows:
                if str(row.get("Código", "")).strip() == cod_ajuste:
                    row["_rol"] = rol
                    return
        debe_rows = [r for r in rows if float(r.get("Debe") or 0) > 0]
        if debe_rows:
            max(debe_rows, key=lambda r: float(r.get("Debe") or 0))["_rol"] = rol
        return
    keywords = (
        "impuesto determinado", "impuesto sobre", "debito fiscal", "iibb devengado",
        "tasa determinada", "sueldos y jornales", "sueldo bruto",
    )
    for row in rows:
        desc = _normalizar_concepto_planilla(str(row.get("Descripción", "")))
        if any(kw in desc for kw in keywords) and float(row.get("Debe") or 0) > 0:
            row["_rol"] = rol
            return
    debe_rows = [r for r in rows if float(r.get("Debe") or 0) > 0]
    if debe_rows:
        max(debe_rows, key=lambda r: float(r.get("Debe") or 0))["_rol"] = rol


def _plan_cuentas_lista_banco(plan_df: pd.DataFrame | None) -> list[tuple[str, str]]:
    """Lista (codigo, descripcion) del plan activo para el intérprete híbrido bancario."""
    if plan_df is None or plan_df.empty:
        return []
    col_cod = _plan_col_codigo(plan_df)
    col_desc = _plan_col_descripcion(plan_df)
    resultado: list[tuple[str, str]] = []
    for _, row in plan_df.iterrows():
        cod = str(row[col_cod]).strip()
        desc = str(row[col_desc]).strip()
        if cod:
            resultado.append((cod, desc))
    return resultado


def _resolver_cuenta_grilla_banco(
    concepto: str,
    plan_df: pd.DataFrame | None,
) -> tuple[str, str]:
    """Delega al intérprete híbrido (código + NLP) sin bloquear la fila."""
    return resolver_cuenta_banco_hibrida(concepto, _plan_cuentas_lista_banco(plan_df))


_ROLES_CIERRE_PARTIDA_DOBLE = frozenset({
    "pagar",
    "saldo_favor",
    "iibb_pagar",
    "saldo_favor_nuevo",
    "tasa_pagar",
    "tish_saldo_favor_nuevo",
})

_ROLES_INPUTS_MANUALES_INYECTABLES = frozenset({
    "retenciones",
    "percepciones",
    "tecnico",
    "libre",
    "retenciones_iibb",
    "percepciones_iibb",
    "retenciones_bancarias",
    "saldo_favor_anterior",
    "tish_saldo_favor",
})


def _motor_usa_cierre_partida_doble(motor: str) -> bool:
    return motor in {"iva", "iibb", "tish"}


def _fila_grilla_cuenta(
    *,
    codigo: str,
    descripcion: str,
    debe: float,
    haber: float,
    rol: str,
    periodo_mensual: str,
    fecha_str: str,
) -> dict:
    debe = round(float(debe or 0), 2)
    haber = round(float(haber or 0), 2)
    tipo = "Debe" if debe > 0 else "Haber"
    return {
        "Período": formatear_periodo_mm_yyyy(periodo_mensual),
        "Fecha": formatear_fecha_dd_mm_yyyy(fecha_str),
        "Código": codigo,
        "Descripción": descripcion,
        "Debe": debe,
        "Haber": haber,
        "Estado": "Ingresado",
        "_rol": rol,
        "_tipo": tipo,
        "_monto": max(debe, haber),
    }


def _mapear_cuenta_cierre_por_motor(plan_df: pd.DataFrame | None, motor: str, rol: str) -> tuple[str, str]:
    plan = plan_df if plan_df is not None else pd.DataFrame()
    if motor == "iibb":
        return _mapear_cuenta_tango_iibb(plan, rol)
    if motor == "tish":
        return obtener_cuenta_tango_tish(plan, rol)
    return _mapear_cuenta_tango(plan, rol)


def _quitar_roles_grilla(rows: list[dict], roles: frozenset[str]) -> list[dict]:
    return [r for r in rows if str(r.get("_rol") or "") not in roles]


def _inyectar_inputs_manuales_en_grilla(
    rows: list[dict],
    *,
    plan_df: pd.DataFrame | None,
    motor: str,
    inputs_manuales: dict[str, float] | None,
    saldos_contingencia: dict[str, list[float]] | None,
    periodo_mensual: str,
    fecha_str: str,
) -> list[dict]:
    """Agrega retenciones/percepciones/saldos cargados a mano (no vienen del Excel)."""
    if not _motor_usa_cierre_partida_doble(motor):
        return rows
    out = _quitar_roles_grilla(list(rows), _ROLES_INPUTS_MANUALES_INYECTABLES)
    inputs_manuales = inputs_manuales or {}
    saldos_contingencia = saldos_contingencia or {}

    def _append(rol: str, debe: float, haber: float) -> None:
        monto_d = round(float(debe or 0), 2)
        monto_h = round(float(haber or 0), 2)
        if monto_d <= 0 and monto_h <= 0:
            return
        cod, desc = _mapear_cuenta_cierre_por_motor(plan_df, motor, rol)
        out.append(
            _fila_grilla_cuenta(
                codigo=cod,
                descripcion=desc,
                debe=monto_d,
                haber=monto_h,
                rol=rol,
                periodo_mensual=periodo_mensual,
                fecha_str=fecha_str,
            )
        )

    if motor == "iva":
        _append("retenciones", 0.0, float(inputs_manuales.get("retenciones") or 0))
        _append("percepciones", 0.0, float(inputs_manuales.get("percepciones") or 0))
        for monto in saldos_contingencia.get("saldos_tecnicos_list") or []:
            _append("tecnico", 0.0, float(monto or 0))
        for monto in saldos_contingencia.get("saldos_libre_list") or []:
            _append("libre", 0.0, float(monto or 0))
    elif motor == "iibb":
        _append("retenciones_iibb", 0.0, float(inputs_manuales.get("retenciones") or 0))
        _append("percepciones_iibb", 0.0, float(inputs_manuales.get("percepciones") or 0))
        _append("retenciones_bancarias", 0.0, float(inputs_manuales.get("retenciones_bancarias") or 0))
        for clave in ("saldos_favor_iibb_list", "saldos_favor_cm_list"):
            for monto in saldos_contingencia.get(clave) or []:
                _append("saldo_favor_anterior", 0.0, float(monto or 0))
    elif motor == "tish":
        for monto in saldos_contingencia.get("saldos_favor_tish_list") or []:
            _append("tish_saldo_favor", 0.0, float(monto or 0))
    return out


def _clave_preservar_cuenta_fila(row: dict) -> tuple:
    """Clave estable para reaplicar una cuenta elegida a mano tras regenerar."""
    fila_idx = row.get("_fila_idx")
    if fila_idx is not None and str(fila_idx).strip() != "":
        return ("fila", int(fila_idx))
    rol = str(row.get("_rol") or "").strip()
    if rol:
        return ("rol", rol)
    desc = str(row.get("Descripción") or row.get("concepto_raw") or "").strip().lower()
    tipo = str(row.get("_tipo") or "").strip()
    return ("desc", desc, tipo)


def _preservar_cuentas_manuales_en_grilla(
    rows_anteriores: list[dict] | None,
    rows_nuevas: list[dict],
) -> list[dict]:
    """Reaplica Código/Descripción que el usuario eligió a mano."""
    if not rows_anteriores or not rows_nuevas:
        return rows_nuevas
    manuales: dict[tuple, tuple[str, str]] = {}
    for r in rows_anteriores:
        if not r.get("_cuenta_manual"):
            continue
        cod = str(r.get("Código") or "").strip()
        if not cod or cod == "99999":
            continue
        manuales[_clave_preservar_cuenta_fila(r)] = (
            cod,
            str(r.get("Descripción") or "").strip(),
        )
    if not manuales:
        return rows_nuevas
    for r in rows_nuevas:
        hit = manuales.get(_clave_preservar_cuenta_fila(r))
        if not hit:
            continue
        r["Código"], r["Descripción"] = hit
        r["_cuenta_manual"] = True
    return rows_nuevas


def _marcar_cuenta_manual_en_fila(row: dict, codigo: str, descripcion: str) -> None:
    row["Código"] = codigo
    row["Descripción"] = descripcion
    row["_cuenta_manual"] = True


def _inyectar_linea_cierre_partida_doble(
    rows: list[dict],
    *,
    plan_df: pd.DataFrame | None,
    motor: str,
    periodo_mensual: str = "",
    fecha_str: str = "",
) -> list[dict]:
    """
    Cierra el asiento con IVA/IIBB/TISH a pagar o saldo a favor.
    Sin esta línea el resumen muestra 'a pagar' pero el asiento no balancea.
    Respeta cuentas marcadas `_cuenta_manual` en la línea de cierre.
    """
    if not rows or not _motor_usa_cierre_partida_doble(motor):
        return rows

    # Guardar cuenta elegida a mano en líneas de cierre antes de quitarlas.
    manual_cierre: dict[str, tuple[str, str]] = {}
    for r in rows:
        rol = str(r.get("_rol") or "")
        if rol in _ROLES_CIERRE_PARTIDA_DOBLE and r.get("_cuenta_manual"):
            cod = str(r.get("Código") or "").strip()
            if cod and cod != "99999":
                manual_cierre[rol] = (cod, str(r.get("Descripción") or "").strip())

    out = _quitar_roles_grilla(list(rows), _ROLES_CIERRE_PARTIDA_DOBLE)
    if not periodo_mensual:
        periodo_mensual = str(out[0].get("Período") or "") if out else ""
    if not fecha_str:
        fecha_str = str(out[0].get("Fecha") or "") if out else ""

    total_debe = round(sum(float(r.get("Debe") or 0) for r in out), 2)
    total_haber = round(sum(float(r.get("Haber") or 0) for r in out), 2)
    dif = round(total_debe - total_haber, 2)
    if abs(dif) < 0.005:
        return out

    if dif > 0:
        if motor == "iibb":
            rol = "iibb_pagar"
        elif motor == "tish":
            rol = "tasa_pagar"
        else:
            rol = "pagar"
        debe, haber = 0.0, dif
    else:
        if motor == "iibb":
            rol = "saldo_favor_nuevo"
        elif motor == "tish":
            rol = "tish_saldo_favor_nuevo"
        else:
            rol = "saldo_favor"
        debe, haber = abs(dif), 0.0

    if rol in manual_cierre:
        cod, desc = manual_cierre[rol]
    else:
        cod, desc = _mapear_cuenta_cierre_por_motor(plan_df, motor, rol)
        if motor == "iva" and rol == "saldo_favor":
            desc = "Saldo a Favor IVA Nuevo Período"
        elif motor == "iibb" and rol == "saldo_favor_nuevo":
            desc = "Saldo a Favor IIBB Nuevo Período"
        elif motor == "tish" and rol == "tish_saldo_favor_nuevo":
            desc = "Saldo a Favor TISH Nuevo Período"

    fila_cierre = _fila_grilla_cuenta(
        codigo=cod,
        descripcion=desc,
        debe=debe,
        haber=haber,
        rol=rol,
        periodo_mensual=periodo_mensual,
        fecha_str=fecha_str,
    )
    if rol in manual_cierre:
        fila_cierre["_cuenta_manual"] = True
    fila_cierre["_sync_monto"] = True
    out.append(fila_cierre)
    return out


def _filas_grilla_desde_extractor_universal(
    filas_extraidas: list[dict],
    *,
    periodo_mensual: str,
    fecha_str: str,
    ficha: dict,
    plan_df: pd.DataFrame | None,
) -> list[dict]:
    """Convierte filas del extractor universal al formato de grilla interactiva."""
    rows: list[dict] = []
    es_banco = ficha.get("motor") == "banco"
    es_iva = ficha.get("motor") == "iva"
    for item in filas_extraidas:
        concepto_raw = str(item.get("concepto_raw") or item.get("descripcion") or "").strip()
        if es_banco:
            codigo, desc = _resolver_cuenta_grilla_banco(concepto_raw, plan_df)
        else:
            codigo_raw = str(item.get("codigo") or "").strip()
            codigo = codigo_raw or "99999"
            # Preferir el concepto del Excel (no pisar con nombre genérico del plan).
            desc = str(item.get("descripcion") or "").strip() or concepto_raw
            if not desc or re.fullmatch(r"\d{5}(?:\.0+)?", desc):
                desc_plan = _descripcion_cuenta_en_plan(plan_df, codigo)
                if desc_plan:
                    desc = desc_plan
        tipo = str(item.get("tipo") or "Debe")
        debe = round(float(item.get("debe", 0) or 0), 2)
        haber = round(float(item.get("haber", 0) or 0), 2)
        # Permitir filas a $0 (conceptos proyectados del Excel sin movimiento en el mes).
        if debe <= 0 and haber <= 0:
            monto = round(abs(float(item.get("monto") or 0)), 2)
            if monto > 0:
                debe = monto if tipo == "Debe" else 0.0
                haber = monto if tipo == "Haber" else 0.0
            else:
                debe, haber = 0.0, 0.0
        monto = max(debe, haber)
        rol = ""
        if es_iva and concepto_raw:
            rol = _matcher_etiqueta_planilla_iva(_normalizar_concepto_planilla(concepto_raw)) or ""
            # NC: proyectar en asiento con lado correcto (aunque resten en el neto / a pagar).
            if rol in ("nc_compras", "nc_compras_27", "nc_ventas", "nc_ventas_27") and monto > 0:
                tipo_rol = _tipo_default_por_rol(rol)
                if tipo_rol != tipo:
                    tipo = tipo_rol
                    debe = monto if tipo == "Debe" else 0.0
                    haber = monto if tipo == "Haber" else 0.0
        row_dict = {
            "Período": formatear_periodo_mm_yyyy(periodo_mensual),
            "Fecha": formatear_fecha_dd_mm_yyyy(fecha_str),
            "Código": codigo,
            "Descripción": desc,
            "Debe": debe,
            "Haber": haber,
            "Estado": "Ingresado",
            "_rol": rol,
            "_tipo": tipo,
            "_monto": monto,
        }
        if ficha.get("motor") == "banco":
            fila_idx = item.get("fila_idx", len(rows))
            tipo_uid = str(item.get("tipo") or "Debe")
            orden = item.get("orden_secuencial", len(rows))
            row_dict["_fila_idx"] = fila_idx
            row_dict["_orden_secuencial"] = orden
            row_dict["_uid"] = f"b{orden}_{tipo_uid.lower()}"
        rows.append(row_dict)
    _marcar_fila_ajuste_loop_review_universal(rows, ficha)
    return rows


def _resumen_analitico_universal_desde_grilla(rows: list[dict]) -> dict:
    total_debe = round(sum(float(r.get("Debe") or 0) for r in rows), 2)
    total_haber = round(sum(float(r.get("Haber") or 0) for r in rows), 2)
    diferencia = round(total_debe - total_haber, 2)
    if diferencia > 0:
        resultado_tipo, resultado_lado = "Saldo Deudor / a Pagar", "Haber"
    elif diferencia < 0:
        resultado_tipo, resultado_lado = "Saldo Acreedor / a Favor", "Debe"
    else:
        resultado_tipo, resultado_lado = "Equilibrado (partida doble)", "—"
    return {
        "total_debe_mes": total_debe,
        "total_haber_mes": total_haber,
        "total_debito_mes": total_debe,
        "total_credito_mes": total_haber,
        "movimiento_puro_mes": diferencia,
        "diferencia_previa": diferencia,
        "resultado_tipo": resultado_tipo,
        "resultado_lado": resultado_lado,
        "resultado_monto": abs(diferencia),
        "impuesto_determinado": round(
            sum(float(r.get("Debe") or 0) for r in rows if float(r.get("Debe") or 0) > 0), 2,
        ),
        "total_descuentos": total_haber,
        "total_saldos_anteriores": 0.0,
        "saldos_tecnicos": [],
        "saldos_libre": [],
        "saldos_favor": [],
    }


def _tipo_impuesto_asiento_desde_motor(motor: str, ficha: dict | None = None) -> str:
    if ficha and ficha.get("codigo_tango"):
        return str(ficha["codigo_tango"])
    if motor == "banco" and ficha:
        return str(ficha.get("codigo_tango", "BANCO"))
    return {
        "iva": "IVA",
        "iibb": "IIBB",
        "sueldos": "SUELDOS",
        "tish": "TISH",
        "banco": "BANCO",
    }.get(motor, motor.upper())


def _asiento_desde_grilla_universal(
    rows: list[dict],
    plan_df: pd.DataFrame,
    mes: int,
    anio: int,
    motor: str,
    ficha: dict | None = None,
) -> AsientoDevengamiento:
    lineas = [
        {
            "Cuenta": str(r.get("Código", "")),
            "Detalle": str(r.get("Descripción", "")),
            "Debe": float(r.get("Debe") or 0),
            "Haber": float(r.get("Haber") or 0),
            "_rol": str(r.get("_rol", "")),
        }
        for r in rows
    ]
    nombre_banco = _nombre_banco_display_desde_ficha(ficha) if motor == "banco" else None
    concepto_banco = (
        f"{nombre_banco} {mes:02d}/{anio}" if motor == "banco" and nombre_banco else None
    )
    asiento = _lineas_a_asiento(
        lineas, plan_df, mes, anio, identificador=1,
        tipo_impuesto=_tipo_impuesto_asiento_desde_motor(motor, ficha),
        concepto_override=concepto_banco,
        nombre_banco=nombre_banco,
    )
    asiento._roles_renglones = [str(r.get("_rol", "")) for r in rows]  # type: ignore[attr-defined]
    asiento._resumen_analitico = _resumen_analitico_universal_desde_grilla(rows)  # type: ignore[attr-defined]
    return asiento


def _reconstruir_asientos_desde_grilla_slug(slug: str, rows: list[dict], ficha: dict) -> None:
    """Reconstruye el asiento Tango completo desde la grilla (alta/baja de renglones)."""
    if not rows:
        st.session_state[f"{slug}_asientos_generados"] = []
        st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla([])
        return
    periodo = str(rows[0].get("Período") or "")
    parsed = _parsear_periodo_texto(periodo.replace("/", "-"))
    if parsed:
        mes, anio = parsed
    else:
        hoy = date.today()
        mes, anio = hoy.month, hoy.year
    plan_df = st.session_state.get("plan_cuentas_df")
    motor = str(ficha.get("motor") or "iva")
    asiento = _asiento_desde_grilla_universal(rows, plan_df, mes, anio, motor, ficha=ficha)
    st.session_state[f"{slug}_asientos_generados"] = [asiento]
    st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla(rows)


def _sincronizar_asientos_desde_grilla_slug(slug: str, rows: list[dict]) -> None:
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    n_renglones = sum(len(a.renglones) for a in asientos)
    if n_renglones != len(rows):
        # Si cambió la cantidad de filas, hace falta reconstruir (cierre / + / borrar).
        return
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx >= len(rows):
                return
            r = rows[flat_idx]
            renglon.codigo_cuenta = str(r.get("Código", ""))
            renglon.descripcion_cuenta = str(r.get("Descripción", ""))
            renglon.debe = round(float(r.get("Debe") or 0), 2)
            renglon.haber = round(float(r.get("Haber") or 0), 2)
            flat_idx += 1


def _finalizar_balance_grilla_slug(
    slug: str,
    ficha: dict,
    *,
    forzar: bool = False,
) -> None:
    """Recalcula asiento solo si la grilla está dirty (o forzar=True)."""
    if not forzar and not _grilla_esta_dirty(slug):
        return
    rows = st.session_state.get(f"{slug}_grilla_preview")
    if not rows:
        _limpiar_grilla_dirty(slug)
        return
    rows = list(rows)
    motor = str(ficha.get("motor") or "")
    plan_df = st.session_state.get("plan_cuentas_df")
    rows = _aplicar_loop_review_filas_grilla_por_ficha(rows, ficha)
    if _motor_usa_cierre_partida_doble(motor):
        rows = _inyectar_linea_cierre_partida_doble(
            rows,
            plan_df=plan_df,
            motor=motor,
        )
    rows = _normalizar_filas_grilla(rows)
    for r in rows:
        # Empujar importes tras cierre/loop review sin pisar tipeos del usuario en otros runs.
        if not r.get("_monto_manual"):
            r["_sync_monto"] = True
    st.session_state[f"{slug}_grilla_preview"] = rows
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    n_renglones = sum(len(a.renglones) for a in asientos)
    if n_renglones != len(rows):
        _reconstruir_asientos_desde_grilla_slug(slug, rows, ficha)
    else:
        _sincronizar_asientos_desde_grilla_slug(slug, rows)
        st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla(rows)
    _limpiar_grilla_dirty(slug)
    _autosalvar_borrador_grilla_throttled(slug)


def _agregar_fila_manual_grilla_slug(slug: str, ficha: dict, plan_df: pd.DataFrame | None) -> None:
    """Agrega una fila vacía editable al asiento (botón ➕)."""
    rows = list(st.session_state.get(f"{slug}_grilla_preview") or [])
    periodo = ""
    fecha_str = ""
    if rows:
        periodo = str(rows[0].get("Período") or "")
        fecha_str = str(rows[0].get("Fecha") or "")
    codigo, descripcion = "99999", "Cuenta a asignar"
    if plan_df is not None and not plan_df.empty:
        try:
            _, opciones, _ = _resolver_plan_y_opciones_cliente(plan_df, rows_grilla=rows)
            if opciones:
                codigo, label = opciones[0]
                descripcion = label.split(" - ", 1)[1] if " - " in label else label
        except Exception:
            pass
    rows.append(
        _fila_grilla_cuenta(
            codigo=str(codigo),
            descripcion=str(descripcion),
            debe=0.0,
            haber=0.0,
            rol="manual",
            periodo_mensual=periodo,
            fecha_str=fecha_str,
        )
    )
    # Dejar Debe/Haber en 0 y tipo Debe por defecto
    rows[-1]["_tipo"] = "Debe"
    st.session_state[f"{slug}_grilla_preview"] = rows
    _reconstruir_asientos_desde_grilla_slug(slug, rows, ficha)
    _marcar_grilla_dirty(slug)
    _limpiar_grilla_dirty(slug)
    st.session_state[f"{slug}_force_autosave"] = True
    _autosalvar_borrador_grilla_throttled(slug)
    st.rerun()


def _aplicar_seleccion_manual_cuenta_slug(slug: str, idx: int, codigo: str, descripcion: str) -> None:
    grilla = st.session_state.get(f"{slug}_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    if grilla[idx].get("Código") == codigo and grilla[idx].get("Descripción") == descripcion:
        return
    grilla[idx]["Código"] = codigo
    grilla[idx]["Descripción"] = descripcion
    grilla[idx]["_cuenta_manual"] = True
    st.session_state[f"{slug}_grilla_preview"] = grilla
    _marcar_grilla_dirty(slug)
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        for renglon in asiento.renglones:
            if flat_idx == idx:
                renglon.codigo_cuenta = codigo
                renglon.descripcion_cuenta = descripcion
                _autosalvar_borrador_grilla_throttled(slug)
                return
            flat_idx += 1


def _aplicar_tipo_fila_slug(
    slug: str,
    ficha: dict,
    idx: int,
    tipo: str,
    *,
    finalizar: bool = True,
) -> None:
    grilla = st.session_state.get(f"{slug}_grilla_preview")
    if grilla is None or idx < 0 or idx >= len(grilla):
        return
    _aplicar_tipo_a_fila_dict(grilla[idx], tipo)
    st.session_state[f"{slug}_grilla_preview"] = grilla
    _marcar_grilla_dirty(slug)
    if finalizar:
        _finalizar_balance_grilla_slug(slug, ficha)


def _eliminar_fila_grilla_slug(slug: str, ficha: dict, idx: int) -> None:
    grilla = st.session_state.get(f"{slug}_grilla_preview") or []
    if idx < 0 or idx >= len(grilla):
        return
    grilla.pop(idx)
    st.session_state[f"{slug}_grilla_preview"] = grilla
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        nuevos_renglones = []
        nuevos_roles: list[str] = []
        roles = getattr(asiento, "_roles_renglones", []) or []
        for ri, renglon in enumerate(asiento.renglones):
            if flat_idx == idx:
                flat_idx += 1
                continue
            nuevos_renglones.append(renglon)
            if ri < len(roles):
                nuevos_roles.append(roles[ri])
            flat_idx += 1
        asiento.renglones = nuevos_renglones
        if hasattr(asiento, "_roles_renglones"):
            asiento._roles_renglones = nuevos_roles  # type: ignore[attr-defined]
    st.session_state[f"{slug}_asientos_generados"] = asientos
    _marcar_grilla_dirty(slug)
    _finalizar_balance_grilla_slug(slug, ficha, forzar=True)
    st.rerun()


def _uid_fila_banco(row: dict, idx: int) -> str:
    uid = str(row.get("_uid") or "").strip()
    if uid:
        return uid
    fila_idx = row.get("_fila_idx", idx)
    codigo = str(row.get("Código", "x"))
    return f"b{fila_idx}_{codigo}"


def _ensure_uids_filas_banco(rows: list[dict]) -> None:
    for idx, row in enumerate(rows):
        if not row.get("_uid"):
            row["_uid"] = _uid_fila_banco(row, idx)


def _limpiar_widgets_grilla_banco(slug: str) -> None:
    prefixes = (
        f"banco_cuenta_row_{slug}_",
        f"banco_cuenta_v3_{slug}_",
        f"banco_tipo_row_{slug}_",
        f"banco_tipo_v3_{slug}_",
        f"banco_monto_row_{slug}_",
        f"banco_monto_v3_{slug}_",
        f"banco_del_row_{slug}_",
        f"banco_del_v3_{slug}_",
        f"banco_swap_row_{slug}_",
        f"banco_swap_v3_{slug}_",
        f"monto_editable_{slug}_",
        f"vent_bco_{slug}_",
    )
    for key in list(st.session_state.keys()):
        key_str = str(key)
        if any(key_str.startswith(p) for p in prefixes):
            st.session_state.pop(key, None)


def _limpiar_widgets_fila_banco(slug: str, uid: str) -> None:
    """Borra solo las keys de UNA fila (por uid). No vaciar toda la grilla: provoca removeChild."""
    uid_s = _safe_wkey_uid(uid)
    gen = int(st.session_state.get(f"banco_ui_gen_{slug}", 0) or 0)
    prefixes = (
        f"banco_cuenta_v3_{slug}_{gen}_{uid_s}",
        f"banco_tipo_v3_{slug}_{gen}_{uid_s}",
        f"banco_monto_v3_{slug}_{gen}_{uid_s}",
        f"banco_del_v3_{slug}_{gen}_{uid_s}",
        f"banco_swap_v3_{slug}_{gen}_{uid_s}",
    )
    for key in list(st.session_state.keys()):
        ks = str(key)
        if any(ks == p or ks.startswith(p) for p in prefixes):
            st.session_state.pop(key, None)


def _eliminar_fila_grilla_banco(slug: str, ficha: dict, idx: int) -> None:
    grilla = st.session_state.get(f"{slug}_grilla_preview") or []
    if idx < 0 or idx >= len(grilla):
        return
    uid = str(grilla[idx].get("_uid") or _uid_fila_banco(grilla[idx], idx))
    _limpiar_widgets_fila_banco(slug, uid)
    grilla.pop(idx)
    st.session_state[f"{slug}_grilla_preview"] = grilla
    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    flat_idx = 0
    for asiento in asientos:
        nuevos_renglones = []
        nuevos_roles: list[str] = []
        roles = getattr(asiento, "_roles_renglones", []) or []
        for ri, renglon in enumerate(asiento.renglones):
            if flat_idx == idx:
                flat_idx += 1
                continue
            nuevos_renglones.append(renglon)
            if ri < len(roles):
                nuevos_roles.append(roles[ri])
            flat_idx += 1
        asiento.renglones = nuevos_renglones
        if hasattr(asiento, "_roles_renglones"):
            asiento._roles_renglones = nuevos_roles  # type: ignore[attr-defined]
    st.session_state[f"{slug}_asientos_generados"] = asientos
    if grilla:
        _marcar_grilla_dirty(slug)
        _finalizar_balance_grilla_slug(slug, ficha, forzar=True)
    else:
        st.session_state[f"{slug}_asientos_generados"] = []
        _limpiar_grilla_dirty(slug)
    st.session_state[f"{slug}_force_autosave"] = True
    _autosalvar_borrador_grilla_throttled(slug)
    st.rerun()


def _render_grilla_interactiva_slug(
    slug: str,
    ficha: dict,
    plan_df: pd.DataFrame,
    *,
    periodo_mensual: str = "",
    sociedad_id: int | None = None,
) -> None:
    """Grilla: cuenta desplegable (plan del cliente) + ⇄ Debe/Haber + importes."""
    if not _salvaguarda_render_grilla_coordenadas(slug, periodo_mensual):
        return
    rows = list(st.session_state.get(f"{slug}_grilla_preview") or [])
    if not rows:
        return

    motor = str(ficha.get("motor") or "")
    # No reinyectar cierre en cada pintado: pisaba cuentas elegidas a mano.
    # El cierre se aplica al generar el asiento y en _finalizar_balance_grilla_slug.

    _, opciones_base, msg_plan = _resolver_plan_y_opciones_cliente(
        plan_df, sociedad_id=sociedad_id, rows_grilla=rows,
    )
    st.caption(msg_plan)
    if not opciones_base:
        st.error(
            "No hay cuentas para el desplegable. "
            "Vinculá el plan de cuentas del cliente en Gestión de Clientes."
        )
        return

    # Período | Fecha | Cuenta | Tipo | Debe | Haber | Estado | Del
    ratios = [0.7, 0.9, 3.45, 0.95, 1.15, 1.15, 0.8, 0.55]
    hdr = st.columns(ratios)
    for col, titulo in zip(
        hdr,
        ["Período", "Fecha", "Cuenta (plan del cliente)", "Tipo", "Debe", "Haber", "Estado", ""],
    ):
        col.markdown(f"**{titulo}**")

    rc = _iva_reset_counter()
    for idx, row in enumerate(rows):
        c = st.columns(ratios)
        c[0].write(_texto_periodo_grilla(row.get("Período", "")))
        c[1].write(_texto_fecha_grilla(row.get("Fecha", "")))

        cod_actual = str(row.get("Código", ""))
        desc_actual = str(row.get("Descripción", ""))
        opciones_fila = _opciones_con_cuenta_actual(opciones_base, cod_actual, desc_actual)
        labels = [label for _, label in opciones_fila]
        cod_por_label = {label: cod for cod, label in opciones_fila}
        desc_por_cod = {}
        for cod, label in opciones_fila:
            if " - " in label:
                desc_por_cod[cod] = label.split(" - ", 1)[1].replace(" (fuera del plan)", "")
            else:
                desc_por_cod[cod] = label.replace(" (fuera del plan)", "")

        sel_index = _indice_opcion_cuenta(opciones_fila, cod_actual)
        select_kwargs: dict = {
            "label": "Cuenta",
            "options": labels,
            "key": f"{slug}_cuenta_{rc}_{idx}",
            "label_visibility": "collapsed",
            "help": "Elegí cualquier cuenta del plan de este cliente. No se pisa sola.",
        }
        if sel_index is not None:
            select_kwargs["index"] = sel_index
        else:
            # Nunca forzar index=0: eso pisaba la cuenta con la primera del plan.
            select_kwargs["index"] = None
            select_kwargs["placeholder"] = "Seleccioná cuenta Tango..."

        with c[2]:
            sel_label = st.selectbox(**select_kwargs)
        if sel_label:
            cod_sel = cod_por_label[sel_label]
            desc_sel = desc_por_cod.get(cod_sel, sel_label)
            if cod_sel != cod_actual or desc_sel != desc_actual:
                _aplicar_seleccion_manual_cuenta_slug(slug, idx, cod_sel, desc_sel)

        tipo_actual = row.get("_tipo") or "Debe"
        tipo_sel, pidio_swap = _render_selector_tipo_con_swap(
            c[3],
            tipo_actual=str(tipo_actual),
            key_tipo=f"{slug}_tipo_{rc}_{idx}",
            key_swap=f"{slug}_swap_{rc}_{idx}",
        )
        if pidio_swap or tipo_sel != row.get("_tipo"):
            _aplicar_tipo_fila_slug(slug, ficha, idx, tipo_sel)

        row = (st.session_state.get(f"{slug}_grilla_preview") or rows)[idx]
        _render_importe_editable_en_grilla(
            c, slug=slug, idx=idx, row=row, rc=rc, tipo=str(tipo_sel), ficha=ficha,
        )

        row = (st.session_state.get(f"{slug}_grilla_preview") or rows)[idx]
        c[6].write(row.get("Estado", ""))
        if c[7].button("🗑️", key=f"{slug}_del_{rc}_{idx}", help="Eliminar fila"):
            _eliminar_fila_grilla_slug(slug, ficha, idx)
            return

    if st.button(
        "➕ Agregar cuenta",
        key=f"{slug}_add_row_{rc}",
        help="Sumá una fila manual al asiento (elegí cuenta e importe)",
    ):
        _agregar_fila_manual_grilla_slug(slug, ficha, plan_df)
        return

    # Solo si hubo edición; evita recálculo+autosave en cada pintado.
    if _grilla_esta_dirty(slug):
        _finalizar_balance_grilla_slug(slug, ficha)


def _render_grilla_interactiva_banco(
    slug: str,
    ficha: dict,
    plan_df: pd.DataFrame,
    banco: str,
    *,
    periodo_mensual: str = "",
    sociedad_id: int | None = None,
) -> None:
    """Grilla bancaria en formulario: se edita sin recargar; aplica al Guardar."""
    del banco
    st.session_state.pop(f"banco_grilla_editor_v1_{slug}", None)
    st.session_state.pop(f"banco_edit_cuenta_idx_{slug}", None)
    st.session_state.pop(f"banco_cuenta_unica_{slug}", None)

    if not _salvaguarda_render_grilla_coordenadas(slug, periodo_mensual):
        return
    rows = list(st.session_state.get(f"{slug}_grilla_preview") or [])
    if not rows:
        return

    _ensure_uids_filas_banco(rows)
    st.session_state[f"{slug}_grilla_preview"] = rows

    _, opciones_base, msg_plan = _resolver_plan_y_opciones_cliente(
        plan_df, sociedad_id=sociedad_id, rows_grilla=rows,
    )
    st.caption(msg_plan)
    if not opciones_base:
        st.error(
            "No hay cuentas para el desplegable. "
            "Vinculá el plan de cuentas del cliente en Gestión de Clientes."
        )
        return

    st.markdown(f"**Grilla de asiento bancario** · {len(rows)} línea(s)")
    st.caption(
        "Editá cuenta, tipo e importe sin que se recargue la página. "
        "Cuando termines, tocá **Aplicar cambios en la grilla**."
    )

    ratios = [0.7, 0.9, 3.2, 0.85, 1.1, 1.1, 0.7, 0.55]
    form_key = f"form_grilla_banco_{slug}"
    with st.form(form_key, clear_on_submit=False):
        hdr = st.columns(ratios)
        for col, titulo in zip(
            hdr,
            ["Período", "Fecha", "Cuenta", "Tipo", "Debe", "Haber", "Estado", "Borrar"],
        ):
            col.markdown(f"**{titulo}**")

        for i, row in enumerate(rows):
            uid = str(row.get("_uid") or _uid_fila_banco(row, i))
            c = st.columns(ratios)
            c[0].write(_texto_periodo_grilla(row.get("Período", "")))
            c[1].write(_texto_fecha_grilla(row.get("Fecha", "")))

            cod_actual = str(row.get("Código", "") or "")
            desc_actual = str(row.get("Descripción", "") or "")
            opciones_fila = _opciones_con_cuenta_actual(opciones_base, cod_actual, desc_actual)
            labels = [label for _, label in opciones_fila]
            cod_por_label = {label: cod for cod, label in opciones_fila}
            desc_por_cod: dict[str, str] = {}
            for cod, label in opciones_fila:
                if " - " in label:
                    desc_por_cod[cod] = label.split(" - ", 1)[1].replace(" (fuera del plan)", "")
                else:
                    desc_por_cod[cod] = label.replace(" (fuera del plan)", "")

            sel_index = _indice_opcion_cuenta(opciones_fila, cod_actual)
            wkey_cuenta = _wkey_banco_cuenta_row(slug, uid)
            if wkey_cuenta not in st.session_state and sel_index is not None and labels:
                st.session_state[wkey_cuenta] = labels[sel_index]

            cuenta_kwargs: dict = {
                "label": "Cuenta",
                "options": labels,
                "key": wkey_cuenta,
                "label_visibility": "collapsed",
            }
            if sel_index is None and labels:
                cuenta_kwargs["placeholder"] = "Seleccione Cuenta..."
            with c[2]:
                st.selectbox(**cuenta_kwargs)

            tipo_actual = row.get("_tipo") or (
                "Debe" if float(row.get("Debe") or 0) >= float(row.get("Haber") or 0) else "Haber"
            )
            wkey_tipo = _wkey_banco_tipo_row(slug, uid)
            if wkey_tipo not in st.session_state:
                st.session_state[wkey_tipo] = tipo_actual if tipo_actual in ("Debe", "Haber") else "Debe"
            with c[3]:
                st.selectbox(
                    "Tipo",
                    options=["Debe", "Haber"],
                    key=wkey_tipo,
                    label_visibility="collapsed",
                )

            # Ambos importes siempre visibles: el form no recarga al cambiar Tipo.
            wkey_debe = _wkey_banco_monto_row(slug, uid, "Debe")
            wkey_haber = _wkey_banco_monto_row(slug, uid, "Haber")
            if wkey_debe not in st.session_state:
                st.session_state[wkey_debe] = float(row.get("Debe") or 0) or (
                    _monto_neto_fila(row) if (row.get("_tipo") or tipo_actual) == "Debe" else 0.0
                )
            if wkey_haber not in st.session_state:
                st.session_state[wkey_haber] = float(row.get("Haber") or 0) or (
                    _monto_neto_fila(row) if (row.get("_tipo") or tipo_actual) == "Haber" else 0.0
                )
            c[4].number_input(
                "Debe",
                min_value=0.0,
                step=0.01,
                format="%.2f",
                key=wkey_debe,
                label_visibility="collapsed",
            )
            c[5].number_input(
                "Haber",
                min_value=0.0,
                step=0.01,
                format="%.2f",
                key=wkey_haber,
                label_visibility="collapsed",
            )

            c[6].write(row.get("Estado", ""))
            c[7].checkbox(
                "Borrar",
                key=_wkey_banco_del_row(slug, uid),
                label_visibility="collapsed",
                help="Marcar para borrar al aplicar",
            )

        aplicado = st.form_submit_button(
            "💾 Aplicar cambios en la grilla",
            type="primary",
            use_container_width=True,
        )

    if not aplicado:
        return

    # Una sola pasada: aplicar widgets → dirty → recalcular asiento
    grilla = list(st.session_state.get(f"{slug}_grilla_preview") or [])
    a_borrar: list[int] = []
    for i, row in enumerate(grilla):
        uid = str(row.get("_uid") or _uid_fila_banco(row, i))
        if bool(st.session_state.get(_wkey_banco_del_row(slug, uid))):
            a_borrar.append(i)
            continue

        opciones_fila = _opciones_con_cuenta_actual(
            opciones_base,
            str(row.get("Código", "") or ""),
            str(row.get("Descripción", "") or ""),
        )
        labels = [label for _, label in opciones_fila]
        cod_por_label = {label: cod for cod, label in opciones_fila}
        desc_por_cod: dict[str, str] = {}
        for cod, label in opciones_fila:
            if " - " in label:
                desc_por_cod[cod] = label.split(" - ", 1)[1].replace(" (fuera del plan)", "")
            else:
                desc_por_cod[cod] = label.replace(" (fuera del plan)", "")

        sel_label = st.session_state.get(_wkey_banco_cuenta_row(slug, uid))
        if sel_label and sel_label in cod_por_label:
            cod_sel = cod_por_label[sel_label]
            desc_sel = desc_por_cod.get(cod_sel, sel_label)
            if cod_sel != row.get("Código") or desc_sel != row.get("Descripción"):
                row["Código"] = cod_sel
                row["Descripción"] = desc_sel
                row["_cuenta_manual"] = True

        tipo_sel = str(st.session_state.get(_wkey_banco_tipo_row(slug, uid)) or row.get("_tipo") or "Debe")
        if tipo_sel not in ("Debe", "Haber"):
            tipo_sel = "Debe"
        debe_v = float(st.session_state.get(_wkey_banco_monto_row(slug, uid, "Debe")) or 0)
        haber_v = float(st.session_state.get(_wkey_banco_monto_row(slug, uid, "Haber")) or 0)
        # Si el usuario cargó un solo lado, respeta ese monto; si ambos, usa el del Tipo.
        if debe_v > 0 and haber_v == 0:
            tipo_sel = "Debe"
            monto = debe_v
        elif haber_v > 0 and debe_v == 0:
            tipo_sel = "Haber"
            monto = haber_v
        else:
            monto = debe_v if tipo_sel == "Debe" else haber_v
        _aplicar_tipo_a_fila_dict(row, tipo_sel)
        aplicar_monto_editable_fila(row, monto, tipo_sel)
        row["_monto_manual"] = True
        row["_tipo"] = tipo_sel

    for i in reversed(a_borrar):
        if 0 <= i < len(grilla):
            grilla.pop(i)

    st.session_state[f"{slug}_grilla_preview"] = grilla
    _marcar_grilla_dirty(slug)
    _finalizar_balance_grilla_slug(slug, ficha, forzar=True)
    # Sincronizar widgets Debe/Haber con la fila real (un solo lado con importe)
    for i, row in enumerate(list(st.session_state.get(f"{slug}_grilla_preview") or [])):
        uid = str(row.get("_uid") or _uid_fila_banco(row, i))
        st.session_state[_wkey_banco_monto_row(slug, uid, "Debe")] = float(row.get("Debe") or 0)
        st.session_state[_wkey_banco_monto_row(slug, uid, "Haber")] = float(row.get("Haber") or 0)
        st.session_state[_wkey_banco_tipo_row(slug, uid)] = str(row.get("_tipo") or "Debe")
        st.session_state[_wkey_banco_del_row(slug, uid)] = False
    st.session_state[f"{slug}_force_autosave"] = True
    _autosalvar_borrador_grilla_throttled(slug, min_interval_s=0.0)
    st.success("Cambios aplicados al asiento.")
    st.rerun()


def _totales_debe_haber_grilla(rows: list[dict]) -> tuple[float, float, float]:
    """Suma Debe/Haber de todas las líneas (asiento completo)."""
    total_debe = round(sum(float(r.get("Debe") or 0) for r in rows), 2)
    total_haber = round(sum(float(r.get("Haber") or 0) for r in rows), 2)
    diferencia = round(total_debe - total_haber, 2)
    return total_debe, total_haber, diferencia


def _asiento_partida_doble_ok(diferencia: float, *, tol: float = 0.05) -> bool:
    """Tolerancia de centavos / redondeo para no marcar falso desbalance."""
    return abs(float(diferencia or 0)) <= tol


def _render_cuadro_control_analitico_universal(
    resumen: dict | None,
    rows: list[dict],
    titulo: str = "Liquidación",
) -> None:
    if resumen is None:
        resumen = {}
    total_debe, total_haber, diferencia = _totales_debe_haber_grilla(rows)
    ok = _asiento_partida_doble_ok(diferencia)
    with st.container(border=True):
        st.markdown(f"### 📊 Resumen Analítico — {titulo}")
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**Total Debe**  \n{_fmt_pesos_ar(total_debe)}")
        c2.markdown(f"**Total Haber**  \n{_fmt_pesos_ar(total_haber)}")
        color = "green" if ok else "red"
        c3.markdown(f"**Diferencia (Debe − Haber)**  \n:{color}[{_fmt_pesos_ar(diferencia)}]")
        st.caption(
            f"{len(rows)} línea(s) · "
            + ("✓ Partida doble OK (balancea)." if ok else "⚠ No balancea: revisá la grilla o aplicá cambios.")
        )
        if not ok:
            st.warning(
                f"El asiento no cierra a cero ({_fmt_pesos_ar(diferencia)}). "
                "Si ya corregiste importes en la grilla, tocá **Aplicar cambios en la grilla** "
                "para recalcular estos totales."
            )


def _render_barra_fija_export_tango(
    *,
    slug: str,
    ficha: dict,
    asientos: list,
    rows: list[dict],
    puede_exportar: bool,
    nombre_activo: str | None,
    cuit_activo: str | None,
    mes_ref: int,
    anio_ref: int,
    key_suffix: str,
) -> None:
    """Barra fija inferior: totales Debe/Haber + Generar Excel Tango (siempre visible)."""
    total_debe, total_haber, diferencia = _totales_debe_haber_grilla(rows)
    ok = _asiento_partida_doble_ok(diferencia)
    color = "#15803d" if ok else "#b91c1c"
    btn_key_frag = f"btn_gen_tango_{slug}_{key_suffix}"
    st.markdown(
        f"""
        <style>
        .ec-tango-sticky-spacer {{ height: 5.75rem; }}
        div[data-testid="stHorizontalBlock"]:has(.st-key-{btn_key_frag}),
        div[data-testid="stHorizontalBlock"]:has([class*="st-key-{btn_key_frag}"]) {{
            position: fixed !important;
            left: max(0px, calc((100vw - 100%) / 2));
            right: 0;
            bottom: 0;
            z-index: 999980 !important;
            width: 100% !important;
            max-width: 100% !important;
            background: #F8FAFC !important;
            border-top: 1px solid #CBD5E1 !important;
            padding: 0.65rem 1.25rem 0.75rem 1.25rem !important;
            margin: 0 !important;
            box-shadow: 0 -4px 16px rgba(15, 23, 42, 0.08) !important;
            align-items: center !important;
        }}
        </style>
        <div class="ec-tango-sticky-spacer"></div>
        """,
        unsafe_allow_html=True,
    )
    col_tot, col_btn = st.columns([1.55, 1.0], gap="medium")
    with col_tot:
        st.markdown(
            f"**Total Debe** {_fmt_pesos_ar(total_debe)}&nbsp;&nbsp;·&nbsp;&nbsp;"
            f"**Total Haber** {_fmt_pesos_ar(total_haber)}&nbsp;&nbsp;·&nbsp;&nbsp;"
            f"<span style='color:{color};font-weight:700'>Diferencia {_fmt_pesos_ar(diferencia)}</span>"
            + ("&nbsp;&nbsp;✓ Balancea" if ok else "&nbsp;&nbsp;⚠ No balancea"),
            unsafe_allow_html=True,
        )
        st.caption(f"{len(rows)} líneas del asiento (todas, no solo la pantalla).")
    with col_btn:
        _render_descarga_excel_tango_diferida(
            slug=slug,
            ficha=ficha,
            asientos=asientos,
            puede_exportar=puede_exportar,
            nombre_activo=nombre_activo,
            cuit_activo=cuit_activo,
            mes_ref=mes_ref,
            anio_ref=anio_ref,
            key_suffix=key_suffix,
            label_generar="Generar archivo Tango",
            label_descargar="Descargar Excel Tango",
        )


def _establecer_grilla_vacia_operativa(
    slug: str,
    *,
    periodo_mensual: str = "",
    aviso: str = "",
    limpiar_coords: bool = True,
) -> None:
    """Grilla vacía pero operativa: preserva biblioteca, plan y sociedad."""
    if limpiar_coords:
        _limpiar_coordenadas_session(slug)
    st.session_state[f"{slug}_grilla_preview"] = []
    st.session_state[f"{slug}_asientos_generados"] = []
    st.session_state[f"{slug}_resumen_analitico"] = {}
    if aviso:
        st.session_state[f"{slug}_extraccion_aviso"] = aviso
    if periodo_mensual:
        st.session_state[f"{slug}_ultimo_periodo_intentado"] = periodo_mensual


def _limpiar_grilla_periodo_slug(slug: str) -> None:
    """Limpieza segura al avanzar de mes: solo datos del período actual del impuesto."""
    st.session_state.pop(f"{slug}_grilla_preview", None)
    st.session_state.pop(f"{slug}_asientos_generados", None)
    st.session_state.pop(f"{slug}_planilla_ref_fecha", None)
    st.session_state.pop(f"{slug}_resumen_analitico", None)
    st.session_state.pop(f"{slug}_auto_fp", None)
    st.session_state.pop(f"{slug}_columna_auditoria", None)
    st.session_state.pop(f"{slug}_extraccion_aviso", None)
    st.session_state[f"{slug}_skip_default_planilla"] = True


def _fingerprint_asiento_auto(
    slug: str,
    periodo: str,
    rc: int,
    inputs_manuales: dict[str, float],
    saldos_contingencia: dict[str, list[float]],
    origen: str,
) -> str:
    # v4: forzar regen tras fix Bahía (11419/11404) + no bloquear por _cuenta_manual
    payload = (
        periodo,
        rc,
        origen,
        "extract_asiento_v4",
    )
    return repr(payload)


def _generar_asiento_automatico(
    *,
    slug: str,
    motor: str,
    impuesto: str,
    ficha: dict,
    sociedad_id: int,
    archivo,
    solapa: str,
    periodo_mensual: str,
    inputs_manuales: dict[str, float],
    saldos_contingencia: dict[str, list[float]],
    origen_planilla: str,
) -> None:
    """Extractor universal fila por fila → grilla interactiva + asiento Tango."""
    periodo_mensual = formatear_periodo_mm_yyyy(periodo_mensual)
    fp = _fingerprint_asiento_auto(
        slug, periodo_mensual, _iva_reset_counter(),
        inputs_manuales, saldos_contingencia, origen_planilla,
    )
    fp_key = f"{slug}_auto_fp"
    if st.session_state.get(fp_key) == fp:
        return

    aviso_periodo = (
        "⚠️ El período seleccionado no contiene movimientos o columnas válidas en este balance. "
        "Seleccione otro mes."
    )
    es_banco = ficha.get("motor") == "banco" or motor == "banco"

    try:
        _sincronizar_plan_cuentas_session(sociedad_id, forzar=True)
        plan_df = st.session_state.get("plan_cuentas_df")
        plan_cliente_id = st.session_state.get("plan_cuentas_cliente_id")
        plan_ok_proc = not bool(st.session_state.get("plan_cuentas_es_default", False))
        if plan_df is None or plan_cliente_id != sociedad_id or not plan_ok_proc:
            return

        buf, es_csv = _abrir_planilla_iva(archivo)
        if es_banco:
            resultado = extraer_filas_universales_balance_por_banco_con_errores(
                buf, solapa, periodo_mensual, es_csv=es_csv,
            )
        else:
            resultado = extraer_filas_universales_balance_por_periodo_con_errores(
                buf, solapa, periodo_mensual, es_csv=es_csv,
            )
        if resultado.error:
            if es_banco:
                aviso = _aviso_diagnostico_extraccion_banco(
                    impuesto=impuesto,
                    periodo_mensual=periodo_mensual,
                    resultado=resultado,
                    archivo=archivo,
                    aviso_periodo=aviso_periodo,
                )
            elif resultado.error_tipo == "sheet_not_found":
                aviso = (
                    f"No se encontró solapa para **{impuesto}** en el balance cargado. "
                    f"{resultado.error}"
                )
            elif resultado.error_tipo == "month_column_not_found":
                aviso = aviso_periodo
            else:
                aviso = f"{aviso_periodo} ({resultado.error})"
            _establecer_grilla_vacia_operativa(
                slug,
                periodo_mensual=periodo_mensual,
                aviso=aviso,
            )
            st.session_state[fp_key] = fp
            return

        filas_extraidas = resultado.filas
        _sincronizar_coordenadas_session(
            slug, resultado.idx_debe, resultado.idx_haber,
        )
        if resultado.columna_cabecera_texto is not None:
            st.session_state[f"{slug}_columna_auditoria"] = {
                "periodo": periodo_mensual,
                "cabecera": resultado.columna_cabecera_texto,
                "col_idx": resultado.columna_indice,
                "idx_debe": resultado.idx_debe,
                "idx_haber": resultado.idx_haber,
                "solapa": resultado.solapa_resuelta or solapa,
            }
        if not filas_extraidas:
            aviso_vacio = (
                _aviso_diagnostico_extraccion_banco(
                    impuesto=impuesto,
                    periodo_mensual=periodo_mensual,
                    resultado=resultado,
                    archivo=archivo,
                    aviso_periodo=aviso_periodo,
                )
                if es_banco
                else aviso_periodo
            )
            _establecer_grilla_vacia_operativa(
                slug,
                periodo_mensual=periodo_mensual,
                aviso=aviso_vacio,
                limpiar_coords=resultado.idx_debe is None,
            )
            st.session_state[fp_key] = fp
            return

        parsed = _parsear_periodo_texto(periodo_mensual.replace("/", "-"))
        if not parsed:
            _establecer_grilla_vacia_operativa(
                slug,
                periodo_mensual=formatear_periodo_mm_yyyy(periodo_mensual),
                aviso=f"Período inválido: {periodo_mensual}",
            )
            st.session_state[fp_key] = fp
            return
        mes, anio = parsed
        periodo_grilla = formatear_periodo_mm_yyyy(periodo_mensual)

        if motor == "banco":
            fecha_asiento = _fecha_asiento_seleccionada_banco(slug, periodo_grilla)
        else:
            fecha_asiento = _fecha_asiento_seleccionada_slug(slug, periodo_grilla)
        if not isinstance(fecha_asiento, date):
            fecha_asiento = _fecha_asiento_iva_tango(mes, anio)
        fecha_str = _formatear_fecha_tango(fecha_asiento)

        rows = _filas_grilla_desde_extractor_universal(
            filas_extraidas,
            periodo_mensual=periodo_grilla,
            fecha_str=fecha_str,
            ficha=ficha,
            plan_df=plan_df,
        )
        # No pisar cuentas que el usuario ya corrigió a mano en esta grilla.
        rows = _preservar_cuentas_manuales_en_grilla(
            st.session_state.get(f"{slug}_grilla_preview"),
            rows,
        )
        if not es_banco:
            rows = _inyectar_inputs_manuales_en_grilla(
                rows,
                plan_df=plan_df,
                motor=motor,
                inputs_manuales=inputs_manuales,
                saldos_contingencia=saldos_contingencia,
                periodo_mensual=periodo_grilla,
                fecha_str=fecha_str,
            )
        rows = _aplicar_loop_review_filas_grilla_por_ficha(rows, ficha)
        if not es_banco:
            rows = _inyectar_linea_cierre_partida_doble(
                rows,
                plan_df=plan_df,
                motor=motor,
                periodo_mensual=periodo_grilla,
                fecha_str=fecha_str,
            )

        asiento = _asiento_desde_grilla_universal(rows, plan_df, mes, anio, motor, ficha=ficha)
        asientos = [asiento]
        resumen = _resumen_analitico_universal_desde_grilla(rows)
        st.session_state[f"{slug}_resumen_analitico"] = resumen

        if isinstance(fecha_asiento, date):
            asientos, rows = _aplicar_fecha_tango_asientos(asientos, rows, fecha_asiento)

        rows = _normalizar_filas_grilla(rows)
        for r in rows:
            r["_sync_monto"] = True
        st.session_state.pop(f"{slug}_extraccion_aviso", None)

    except Exception as exc:
        _establecer_grilla_vacia_operativa(
            slug,
            periodo_mensual=periodo_mensual,
            aviso=f"{aviso_periodo} ({exc})",
        )
        st.session_state[fp_key] = fp
        return

    st.session_state[f"{slug}_grilla_preview"] = rows
    st.session_state[f"{slug}_asientos_generados"] = asientos
    st.session_state[fp_key] = fp
    _limpiar_grilla_dirty(slug)
    if es_banco:
        _bump_ui_grilla_banco(slug)
    st.session_state.pop(f"{slug}_tango_xlsx_bytes", None)
    st.session_state.pop(f"{slug}_tango_xlsx_fp", None)
    st.session_state.pop(f"{slug}_tango_xlsx_name", None)
    st.session_state[f"{slug}_force_autosave"] = True
    _autosalvar_borrador_grilla_throttled(slug)


def _aplicar_loop_review_filas_grilla_por_ficha(rows: list[dict], ficha: dict) -> list[dict]:
    if not rows:
        return rows
    rol_ajuste = ficha.get("cuenta_ajuste_centavos_rol", "")
    lineas = [
        {
            "Debe": float(r.get("Debe") or 0),
            "Haber": float(r.get("Haber") or 0),
            "_rol": r.get("_rol", ""),
            "Detalle": r.get("Descripción", ""),
        }
        for r in rows
    ]
    if rol_ajuste and any(lin.get("_rol") == rol_ajuste and lin["Debe"] > 0 for lin in lineas):
        lineas, _, _ = _aplicar_loop_review_por_rol(lineas, rol_ajuste)
    elif ficha.get("motor") == "banco":
        rol_banco = ficha.get("cuenta_ajuste_centavos_rol", "cuenta_banco_principal")
        lineas, _, _ = _aplicar_loop_review_por_rol(lineas, rol_banco)
    elif ficha.get("motor") == "iva":
        lineas, _, _ = _aplicar_loop_review_ventas_21(lineas)
    elif ficha.get("motor") == "iibb":
        lineas, _, _ = _aplicar_loop_review_impuesto_determinado(lineas)
    elif rol_ajuste:
        lineas, _, _ = _aplicar_loop_review_por_rol(lineas, rol_ajuste)
    for i, lin in enumerate(lineas):
        rows[i]["Debe"] = lin["Debe"]
        rows[i]["Haber"] = lin["Haber"]
        rows[i]["_monto"] = _monto_neto_fila(rows[i])
    return rows


def _puede_exportar_asiento_generico(rows: list[dict]) -> bool:
    return all(str(r.get("Código", "")).strip() not in ("", "99999") for r in rows)


def _render_grilla_vacia_operativa(slug: str, impuesto: str) -> None:
    """Placeholder operativo cuando no hay filas extraídas para el período."""
    st.markdown("**Grilla de asiento**")
    st.caption(
        f"Sin movimientos cargados para el período actual de {impuesto}. "
        "Seleccioná otro mes en el selector o revisá la solapa del balance."
    )


def _render_resultados_conciliacion_bancaria(
    banco: str,
    slug: str,
    ficha: dict,
    sociedad_id: int,
    cuit_activo: str | None,
    nombre_activo: str | None,
) -> None:
    """Bloque inferior bancario — espejo de Devengamientos con keys y biblioteca propias."""
    aviso_extraccion = st.session_state.get(f"{slug}_extraccion_aviso")
    if aviso_extraccion:
        st.warning(aviso_extraccion)

    try:
        _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=False)
    except Exception:
        pass

    if _restaurar_borrador_grilla_si_aplica(slug):
        st.info(
            "♻️ Se recuperó el borrador de la grilla desde disco "
            f"(`borrador_actual.json`). Revisá los importes antes de exportar."
        )

    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    rows = st.session_state.get(f"{slug}_grilla_preview")
    if rows is None:
        rows = []
    fecha_vigente = _fecha_asiento_seleccionada_banco(
        slug, str(rows[0].get("Período", "") or "") if rows else None,
    )
    if rows and asientos and isinstance(fecha_vigente, date):
        asientos, rows = _aplicar_fecha_tango_asientos(asientos, rows, fecha_vigente)
        st.session_state[f"{slug}_asientos_generados"] = asientos
        st.session_state[f"{slug}_grilla_preview"] = rows

    if not rows:
        _render_grilla_vacia_operativa(slug, banco)
        # Tras guardar un mes la grilla se limpia; el Excel consolidado
        # debe seguir disponible mientras haya biblioteca.
        st.markdown("---")
        _render_bloque_excel_consolidado_banco(
            sociedad_id, banco, cuit_activo, nombre_activo,
            key_suffix="main_vacio",
            asegurar_plan=True,
        )
        return
    if not asientos:
        st.warning(f"No se detectaron líneas de asiento en la solapa {banco}.")
        st.markdown("---")
        _render_bloque_excel_consolidado_banco(
            sociedad_id, banco, cuit_activo, nombre_activo,
            key_suffix="main_sin_asiento",
            asegurar_plan=True,
        )
        return

    periodo_ref = str(rows[0].get("Período", "") or "")
    if not _salvaguarda_render_grilla_coordenadas(slug, periodo_ref):
        return

    st.success(f"{len(asientos)} asiento(s) bancario(s) {banco} generado(s).")
    audit_col = st.session_state.get(f"{slug}_columna_auditoria")
    if audit_col and audit_col.get("cabecera"):
        st.info(
            f"🔍 Sistema leyendo la columna del Excel titulada: "
            f"[{audit_col['cabecera']}]"
        )

    _ensure_uids_filas_banco(rows)
    st.session_state[f"{slug}_grilla_preview"] = rows
    # Solo recalcular si hubo edición aplicada (no en cada tecla).
    if _grilla_esta_dirty(slug):
        _finalizar_balance_grilla_slug(slug, ficha)
    rows = st.session_state.get(f"{slug}_grilla_preview") or rows
    st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla(rows)

    # Totales del asiento completo ANTES de la grilla (siempre visibles)
    _render_cuadro_control_analitico_universal(
        st.session_state.get(f"{slug}_resumen_analitico"), rows, banco,
    )

    _render_grilla_interactiva_banco(
        slug, ficha, st.session_state.get("plan_cuentas_df"), banco,
        periodo_mensual=periodo_ref,
        sociedad_id=sociedad_id,
    )
    rows = st.session_state.get(f"{slug}_grilla_preview") or rows
    st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla(rows)

    rows = st.session_state.get(f"{slug}_grilla_preview") or rows
    asientos = st.session_state.get(f"{slug}_asientos_generados") or asientos
    puede_exportar = _puede_exportar_asiento_generico(rows)
    finalizar = lambda: _finalizar_balance_grilla_slug(slug, ficha, forzar=True)

    if not puede_exportar:
        st.error(
            "Exportación bloqueada: hay conceptos con código 99999 sin cuenta Tango "
            "asignada. Completá todos los selectores de la grilla antes de descargar."
        )

    try:
        periodo_ref = getattr(asientos[0], "periodo", "") or ""
        mes_ref, anio_ref = (int(x) for x in periodo_ref.split("/"))
    except Exception:
        hoy = pd.Timestamp.today()
        mes_ref, anio_ref = hoy.month, hoy.year

    _render_barra_fija_export_tango(
        slug=slug,
        ficha=ficha,
        asientos=asientos,
        rows=rows,
        puede_exportar=puede_exportar,
        nombre_activo=nombre_activo,
        cuit_activo=cuit_activo,
        mes_ref=mes_ref,
        anio_ref=anio_ref,
        key_suffix="banco_mes",
    )

    if puede_exportar and st.button(
        "💾 Guardar Asiento en Biblioteca",
        type="primary",
        key=f"btn_guardar_biblioteca_banco_v2_{slug}",
        use_container_width=True,
    ):
        try:
            finalizar()
            rows = st.session_state.get(f"{slug}_grilla_preview") or rows
            asientos = st.session_state.get(f"{slug}_asientos_generados") or asientos
            periodo_guardado = _guardar_asiento_en_biblioteca_banco(
                sociedad_id, asientos, rows, banco=banco,
            )
            siguiente = _periodo_siguiente(_periodo_a_etiqueta(periodo_guardado).replace("-", "/"))
            if siguiente:
                st.session_state[f"{slug}_periodo_mensual_avanzar"] = siguiente
            st.session_state.pop(f"{slug}_auto_fp", None)
            st.session_state.pop(f"{slug}_periodo_prefill_{sociedad_id}", None)
            _marcar_limpieza_formulario_mes_banco(slug)
            st.session_state[f"{slug}_mensaje_biblioteca_ok"] = periodo_guardado
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"No se pudo archivar el asiento: {exc}")

    entradas_bib = _biblioteca_bancos_por_sociedad(sociedad_id, banco)
    if entradas_bib:
        st.markdown("---")
        _render_bloque_excel_consolidado_banco(
            sociedad_id, banco, cuit_activo, nombre_activo,
            key_suffix="main",
            asegurar_plan=True,
        )


def _render_resultados_devengamiento(
    impuesto: str,
    slug: str,
    ficha: dict,
    sociedad_id: int,
    cuit_activo: str | None,
    nombre_activo: str | None,
) -> None:
    aviso_extraccion = st.session_state.get(f"{slug}_extraccion_aviso")
    if aviso_extraccion:
        st.warning(aviso_extraccion)

    # Asegurar plan del cliente activo antes de armar los desplegables.
    try:
        _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=False)
    except Exception:
        pass

    if _restaurar_borrador_grilla_si_aplica(slug):
        st.info(
            "♻️ Se recuperó el borrador de la grilla desde disco "
            f"(`borrador_actual.json`). Revisá los importes antes de exportar."
        )

    asientos = st.session_state.get(f"{slug}_asientos_generados") or []
    rows = st.session_state.get(f"{slug}_grilla_preview")
    if rows is None:
        rows = []
    fecha_vigente = _fecha_asiento_seleccionada_slug(
        slug, str(rows[0].get("Período", "") or "") if rows else None,
    )
    if rows and asientos and isinstance(fecha_vigente, date):
        asientos, rows = _aplicar_fecha_tango_asientos(asientos, rows, fecha_vigente)
        st.session_state[f"{slug}_asientos_generados"] = asientos
        st.session_state[f"{slug}_grilla_preview"] = rows

    motor = ficha.get("motor", slug)
    es_motor_grilla = motor in ("iva", "iibb", "sueldos", "tish")
    if not rows:
        if es_motor_grilla:
            _render_grilla_vacia_operativa(slug, impuesto)
        return
    if not asientos:
        st.warning(f"No se detectaron líneas de asiento en la planilla {impuesto}.")
        return

    periodo_ref = str(rows[0].get("Período", "") or "")
    if not _salvaguarda_render_grilla_coordenadas(slug, periodo_ref):
        return

    st.success(f"{len(asientos)} asiento(s) {impuesto} generado(s).")
    audit_col = st.session_state.get(f"{slug}_columna_auditoria")
    if audit_col and audit_col.get("cabecera"):
        st.info(
            f"🔍 Sistema leyendo la columna del Excel titulada: "
            f"[{audit_col['cabecera']}]"
        )
    if es_motor_grilla:
        if motor == "iva":
            _mostrar_alerta_auditoria_cuentas_iva(rows)
        elif motor == "iibb":
            _mostrar_alerta_auditoria_cuentas_iibb(rows)
        _render_grilla_interactiva_slug(
            slug, ficha, st.session_state.get("plan_cuentas_df"),
            periodo_mensual=periodo_ref,
            sociedad_id=sociedad_id,
        )
        rows = st.session_state.get(f"{slug}_grilla_preview") or rows
        st.session_state[f"{slug}_resumen_analitico"] = _resumen_analitico_universal_desde_grilla(rows)
        if motor == "iva":
            _render_cuadro_control_analitico_iva(st.session_state.get(f"{slug}_resumen_analitico"), rows)
        elif motor == "iibb":
            _render_cuadro_control_analitico_iibb(st.session_state.get(f"{slug}_resumen_analitico"), rows)
        else:
            _render_cuadro_control_analitico_universal(
                st.session_state.get(f"{slug}_resumen_analitico"), rows, impuesto,
            )
        puede_exportar = _puede_exportar_asiento_generico(rows)
        finalizar = lambda: _finalizar_balance_grilla_slug(slug, ficha, forzar=True)
    else:
        # Motores sin grilla dedicada: igual permitir ⇄ Debe/Haber.
        _render_grilla_interactiva_slug(
            slug, ficha, st.session_state.get("plan_cuentas_df"),
            periodo_mensual=periodo_ref,
            sociedad_id=sociedad_id,
        )
        rows = st.session_state.get(f"{slug}_grilla_preview") or rows
        resumen = st.session_state.get(f"{slug}_resumen_analitico") or {}
        if resumen:
            st.markdown(
                f"**Posición:** {_fmt_pesos_ar(resumen.get('diferencia_previa', 0))} — "
                f"{resumen.get('resultado_tipo', '—')}"
            )
        puede_exportar = _puede_exportar_asiento_generico(rows)
        finalizar = lambda: _finalizar_balance_grilla_slug(slug, ficha, forzar=True)

    rows = st.session_state.get(f"{slug}_grilla_preview") or rows
    asientos = st.session_state.get(f"{slug}_asientos_generados") or asientos

    if not puede_exportar:
        st.error(
            "Exportación bloqueada: hay conceptos con código 99999 sin cuenta Tango "
            "asignada. Completá todos los selectores de la grilla antes de descargar."
        )

    try:
        periodo_ref = getattr(asientos[0], "periodo", "") or ""
        mes_ref, anio_ref = (int(x) for x in periodo_ref.split("/"))
    except Exception:
        hoy = pd.Timestamp.today()
        mes_ref, anio_ref = hoy.month, hoy.year

    if puede_exportar:
        plan_df = st.session_state.get("plan_cuentas_df")
        informe_mes = (
            auditar_exportacion_tango(asientos, plan_df)
            if plan_df is not None and not plan_df.empty and asientos
            else {"bloqueantes": [], "advertencias": []}
        )
        if informe_mes.get("bloqueantes") or informe_mes.get("advertencias"):
            _mostrar_informe_exportacion_tango(informe_mes)
        if not informe_mes.get("bloqueantes"):
            _render_descarga_excel_tango_diferida(
                slug=slug,
                ficha=ficha,
                asientos=asientos,
                puede_exportar=True,
                nombre_activo=nombre_activo,
                cuit_activo=cuit_activo,
                mes_ref=mes_ref,
                anio_ref=anio_ref,
                key_suffix="dev_mes",
            )

    if puede_exportar and st.button(
        "💾 Guardar Asiento en Biblioteca",
        type="primary",
        key=f"btn_guardar_biblioteca_{slug}",
        use_container_width=True,
    ):
        try:
            finalizar()
            rows = st.session_state.get(f"{slug}_grilla_preview") or rows
            asientos = st.session_state.get(f"{slug}_asientos_generados") or asientos
            periodo_guardado = _guardar_asiento_en_biblioteca(
                sociedad_id, asientos, rows, impuesto=impuesto,
            )
            siguiente = _periodo_siguiente(_periodo_a_etiqueta(periodo_guardado).replace("-", "/"))
            if siguiente:
                st.session_state[f"{slug}_periodo_mensual_avanzar"] = siguiente
            st.session_state.pop(f"{slug}_auto_fp", None)
            st.session_state.pop(f"{slug}_periodo_prefill_{sociedad_id}", None)
            _marcar_limpieza_formulario_mes_por_impuesto(impuesto)
            st.session_state[f"{slug}_mensaje_biblioteca_ok"] = periodo_guardado
            st.rerun()
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"No se pudo archivar el asiento: {exc}")

    entradas_bib = _biblioteca_por_sociedad(sociedad_id, impuesto)
    if entradas_bib:
        st.markdown("---")
        st.markdown(f"### 📦 Exportación consolidada a Tango ({impuesto})")
        df_consolidado = _dataframe_consolidado_biblioteca(sociedad_id, impuesto)
        if not df_consolidado.empty:
            with st.expander(
                f"Vista previa consolidada ({len(entradas_bib)} mes/es, "
                f"{len(df_consolidado)} renglones)",
                expanded=False,
            ):
                st.dataframe(df_consolidado, use_container_width=True, hide_index=True)
        plan_df = _asegurar_plan_cuentas_export(
            sociedad_id, cuit_activo, slug=slug,
        )
        asientos_cons = _asientos_consolidados_biblioteca(sociedad_id, impuesto)
        if plan_df is not None and _auditar_asientos_antes_export_ui(asientos_cons, plan_df):
            try:
                ruta_consolidada = _generar_excel_biblioteca_consolidada(
                    sociedad_id, cuit_activo or "000", nombre_activo or "", impuesto=impuesto,
                    plan_cuentas=plan_df,
                )
                with open(ruta_consolidada, "rb") as f:
                    st.download_button(
                        "📦 Generar Archivo de Importación Consolidado (Tango)",
                        f,
                        file_name=ruta_consolidada.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        key=f"dl_tango_{slug}_consolidado",
                        use_container_width=True,
                    )
            except ExportacionTangoError as exc:
                _mostrar_errores_exportacion_tango(exc)
            except Exception as exc:
                st.error(f"No se pudo generar el archivo consolidado: {exc}")


def _seccion_devengamientos_impuesto(
    impuesto: str,
    sociedad_id: int,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    """Flujo unificado de devengamiento guiado por TAX_REGISTRY."""
    ficha = obtener_ficha_impuesto(impuesto)
    slug = ficha.get("slug", _slug_impuesto(impuesto))
    motor = ficha.get("motor", slug)

    _inicializar_sesion_por_impuesto(impuesto)
    _inicializar_estado_coordenadas_debe_haber(_slug_impuesto(impuesto))
    _aplicar_limpieza_formulario_mes_por_impuesto(impuesto)
    _render_panel_biblioteca_asientos(sociedad_id, impuesto)

    rc = _iva_reset_counter()
    periodo_ok = st.session_state.pop(f"{slug}_mensaje_biblioteca_ok", None)
    if periodo_ok:
        st.success(
            f"✓ Asiento {impuesto} de {periodo_ok} guardado con éxito. "
            "Ya podés cargar el siguiente mes."
        )

    st.markdown("---")
    _render_conexion_servidor_local(sociedad_id, impuesto)

    uploader_key = f"uploader_balance_{slug}_{rc}"
    solapa_activa = impuesto
    archivo_planilla = st.file_uploader(
        "Subir Balance completo (respaldo / contingencia)",
        type=["xlsx", "xls", "csv"],
        key=uploader_key,
        help=(
            f"Arrastrá el archivo de Balance del estudio. Se procesará la solapa «{solapa_activa}» "
            "buscando conceptos por palabras clave en columna A."
        ),
    )

    archivo_a_procesar, origen_planilla, etiqueta_planilla = _resolver_fuente_balance(
        archivo_planilla, sociedad_id,
    )
    if archivo_a_procesar is not None:
        if origen_planilla == "servidor":
            st.success(
                f"📁 Usando Balance cargado desde **servidor local** "
                f"(solapa **{solapa_activa}**)."
            )
        elif origen_planilla == "upload":
            st.caption(
                f"📁 Balance cargado en memoria: `{etiqueta_planilla}` — solapa **{solapa_activa}**."
            )

    periodo_mensual: str | None = None
    periodo_key = _clave_periodo_mensual(slug, sociedad_id)
    if archivo_a_procesar is not None:
        buf_periodos, es_csv_periodos = _abrir_planilla_iva(archivo_a_procesar)
        periodos_disponibles = listar_periodos_disponibles_balance(
            buf_periodos, solapa_activa, es_csv=es_csv_periodos,
        )
        idx_default = 0
        periodo_previo = st.session_state.get(periodo_key)
        if periodo_previo and periodo_previo in periodos_disponibles:
            idx_default = periodos_disponibles.index(periodo_previo)
        periodo_mensual = st.selectbox(
            "📅 Período Mensual a Procesar",
            periodos_disponibles,
            index=idx_default,
            key=periodo_key,
        )
        prefill_key = f"{slug}_periodo_prefill_{sociedad_id}"
        if st.session_state.get(prefill_key) != f"{periodo_mensual}_{rc}":
            try:
                datos_periodo = _leer_datos_balance_periodo(
                    archivo_a_procesar, motor, solapa_activa, periodo_mensual,
                )
                _aplicar_datos_periodo_a_inputs(slug, ficha, motor, datos_periodo, rc)
                st.session_state[prefill_key] = f"{periodo_mensual}_{rc}"
                st.session_state.pop(f"{slug}_auto_fp", None)
            except Exception:
                pass

    inputs_manuales: dict[str, float] = {}
    saldos_contingencia: dict[str, list[float]] = {}
    contingencia = ficha.get("inputs_contingencia") or []
    manuales = ficha.get("inputs_manuales") or []
    n_cols = 2 if contingencia or manuales else 1
    cols = st.columns(n_cols)
    mitad = (len(contingencia) + len(manuales) + 1) // 2
    items = [(c, "cont") for c in contingencia] + [(m, "man") for m in manuales]
    for idx, (item, tipo) in enumerate(items):
        col = cols[0 if idx < mitad else 1]
        with col:
            if tipo == "cont":
                _render_lista_saldos_por_ficha(item, slug)
            else:
                clave = item["clave"]
                inputs_manuales[clave] = st.number_input(
                    item["titulo"],
                    min_value=0.0,
                    step=0.01,
                    format="%.2f",
                    key=f"{slug}_{clave}_{rc}",
                )

    for inp in contingencia:
        clave = inp["clave"]
        saldos_contingencia[clave] = _obtener_valores_saldos_desde_session(clave, inp["prefix"])

    fecha_key = _clave_fecha_tango_slug(slug, rc, periodo_mensual)
    if periodo_mensual:
        _asegurar_fecha_asiento_ultimo_dia(fecha_key, periodo_mensual)
    else:
        _prefijar_fecha_input_desde_planilla_por_motor(archivo_a_procesar, sociedad_id, motor)
    st.date_input(
        "Fecha del Asiento para Tango",
        format="DD/MM/YYYY",
        key=fecha_key,
        help=(
            "Fuente de verdad para Tango. Se fija al último día del período "
            "mensual seleccionado (podés ajustarla a mano si hace falta)."
        ),
    )

    if archivo_a_procesar is not None and not plan_vinculado:
        st.error(
            "No podés procesar el devengamiento hasta vincular el Plan de Cuentas de esta sociedad "
            "en el Editor de Clientes."
        )

    puede_procesar = archivo_a_procesar is not None and plan_vinculado and bool(periodo_mensual)
    if puede_procesar:
        _generar_asiento_automatico(
            slug=slug,
            motor=motor,
            impuesto=impuesto,
            ficha=ficha,
            sociedad_id=sociedad_id,
            archivo=archivo_a_procesar,
            solapa=solapa_activa,
            periodo_mensual=periodo_mensual,
            inputs_manuales=inputs_manuales,
            saldos_contingencia=saldos_contingencia,
            origen_planilla=origen_planilla or "",
        )

    _render_resultados_devengamiento(
        impuesto, slug, ficha, sociedad_id, cuit_activo, nombre_activo,
    )


def _parsear_cuenta_sugerida_pdf_banco(texto: str) -> tuple[str, str]:
    """Convierte «11301 - Deudores» o texto suelto en (código, descripción)."""
    s = str(texto or "").strip()
    if not s or s.lower().startswith("seleccione"):
        return "", ""
    if " - " in s:
        cod, desc = s.split(" - ", 1)
        return cod.strip(), desc.strip()
    codigo, resto = extraer_codigo_cuenta_tango_desde_concepto(s)
    if codigo and codigo != "99999":
        return codigo, resto or s
    return "", s


def _aplicar_matcheo_pdf_a_grilla_banco(
    slug: str,
    lineas: list[dict],
    *,
    ficha: dict,
    plan_df: pd.DataFrame | None,
    periodo_mensual: str | None,
    banco: str,
) -> None:
    """Carga en session_state la grilla editable a partir del matcheo PDF + Tango."""
    if not lineas:
        return

    periodo_grilla = formatear_periodo_mm_yyyy(periodo_mensual) if periodo_mensual else ""
    if not periodo_grilla:
        for lin in lineas:
            fecha_raw = lin.get("fecha")
            fecha_obj = parsear_fecha_export_tango(fecha_raw) if fecha_raw else None
            if fecha_obj:
                periodo_grilla = f"{fecha_obj.month:02d}/{fecha_obj.year}"
                break
    if not periodo_grilla:
        hoy = date.today()
        periodo_grilla = f"{hoy.month:02d}/{hoy.year}"

    parsed = _parsear_periodo_texto(periodo_grilla.replace("/", "-"))
    if not parsed:
        st.error(f"Período inválido para la grilla: {periodo_grilla}")
        return
    mes, anio = parsed

    fecha_asiento = _fecha_asiento_seleccionada_banco(slug, periodo_grilla)
    if not isinstance(fecha_asiento, date):
        fecha_asiento = _fecha_asiento_iva_tango(mes, anio)
    fecha_default_str = _formatear_fecha_tango(fecha_asiento)

    rows: list[dict] = []
    for orden, lin in enumerate(lineas):
        codigo, desc = _parsear_cuenta_sugerida_pdf_banco(str(lin.get("cuenta", "")))
        if not codigo:
            codigo = "99999"
        if not desc:
            desc = str(lin.get("concepto", "") or "").strip()
        desc_plan = _descripcion_cuenta_en_plan(plan_df, codigo)
        if desc_plan:
            desc = desc_plan

        monto = round(float(lin.get("monto") or 0), 2)
        tipo = str(lin.get("tipo") or "Debe")
        debe = monto if tipo == "Debe" else 0.0
        haber = monto if tipo == "Haber" else 0.0

        fecha_lin = parsear_fecha_export_tango(lin.get("fecha")) if lin.get("fecha") else None
        fecha_str = (
            formatear_fecha_dd_mm_yyyy(fecha_lin.strftime("%d/%m/%Y"))
            if isinstance(fecha_lin, date)
            else fecha_default_str
        )

        rows.append({
            "Período": periodo_grilla,
            "Fecha": fecha_str,
            "Código": codigo,
            "Descripción": desc,
            "Debe": debe,
            "Haber": haber,
            "Estado": "Ingresado",
            "_rol": "",
            "_tipo": tipo,
            "_monto": monto,
            "_fila_idx": orden,
            "_orden_secuencial": orden,
            "_uid": f"pdf{orden}_{tipo.lower()}",
        })

    _marcar_fila_ajuste_loop_review_universal(rows, ficha)
    rows = _aplicar_loop_review_filas_grilla_por_ficha(rows, ficha)
    asiento = _asiento_desde_grilla_universal(
        rows, plan_df or pd.DataFrame(), mes, anio, "banco", ficha=ficha,
    )
    asiento.tipo = "BANCO"  # type: ignore[attr-defined]
    asiento.periodo = periodo_grilla  # type: ignore[attr-defined]
    asiento.concepto = f"CONCILIACION {banco} {periodo_grilla}"  # type: ignore[attr-defined]

    if isinstance(fecha_asiento, date):
        asientos, rows = _aplicar_fecha_tango_asientos([asiento], rows, fecha_asiento)
    else:
        asientos = [asiento]

    rows = _normalizar_filas_grilla(rows)
    st.session_state[f"{slug}_grilla_preview"] = rows
    st.session_state[f"{slug}_asientos_generados"] = asientos
    st.session_state.pop(f"{slug}_extraccion_aviso", None)
    st.session_state.pop(f"{slug}_auto_fp", None)
    _autosalvar_borrador_grilla(slug)


def _render_conexion_inteligente_extractos_pdf(
    banco: str,
    slug: str,
    ficha: dict,
    sociedad_id: int,
    plan_vinculado: bool,
    periodo_mensual: str | None,
    *,
    como_herramienta: bool = False,
) -> None:
    """Buzón PDF + subdiario Tango → matcheo inteligente → grilla editable."""
    if como_herramienta:
        st.markdown("### Conexión Inteligente — Procesador de Extractos PDF")
        st.caption(
            "Subí el extracto digital y el listado por imputación de Tango. "
            "Al terminar, abrí **Conciliación Bancaria** con el mismo banco para ver "
            "la grilla completa y exportar el asiento."
        )
    else:
        st.markdown("### Configuración de la Conciliación Automática")
        st.caption(
            f"Banco activo: **{banco}**. Subí el extracto digital y el listado por imputación de Tango "
            "para sugerir cuentas contables antes de editar la grilla."
        )

    st.info(f"Banco: **{banco}** · Período: **{periodo_mensual or '(según movimientos del PDF)'}**")

    col_pdf, col_tango = st.columns(2)
    with col_pdf:
        archivo_pdf = st.file_uploader(
            "1. Subir Extracto Bancario (PDF digital)",
            type=["pdf"],
            key=f"{slug}_pdf_extracto_inteligente",
        )
    with col_tango:
        archivo_tango = st.file_uploader(
            "2. Subir Listado por Imputación de Tango (CSV o Excel)",
            type=["csv", "xlsx", "xls"],
            key=f"{slug}_tango_subdiario_inteligente",
        )

    if st.button(
        "⚡ Ejecutar Matcheo Inteligente y Generar Asiento",
        key=f"{slug}_btn_matcheo_pdf",
        type="primary",
    ):
        if not plan_vinculado:
            st.error("Vinculá el Plan de Cuentas de la sociedad antes de ejecutar el matcheo.")
            return
        if archivo_pdf is None or archivo_tango is None:
            st.warning("Cargá el PDF del banco y el reporte de Tango para continuar.")
            return

        with st.spinner("Procesando archivos y cruzando datos con Tango..."):
            try:
                df_tango = cargar_subdiario_tango(archivo_tango)
                plan_df = st.session_state.get("plan_cuentas_df")
                plan_lista = _plan_cuentas_lista_banco(plan_df)

                banco_norm = _normalizar_concepto_planilla(banco)
                if "galicia" in banco_norm:
                    df_banco = extraer_datos_pdf_galicia(archivo_pdf)
                else:
                    import tempfile
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(archivo_pdf.getvalue())
                        tmp_path = tmp.name
                    try:
                        banco_clave = "galicia" if "galicia" in banco_norm else (
                            "santander" if "santander" in banco_norm else banco_norm.replace(" ", "_")
                        )
                        movs = extraer_movimientos_banco(tmp_path, banco=banco_clave)
                        df_banco = movimientos_banco_a_dataframe_conciliacion(movs)
                    finally:
                        Path(tmp_path).unlink(missing_ok=True)

                if df_banco is None or df_banco.empty:
                    st.error("No se detectaron movimientos en el PDF. Verificá que sea un extracto digital legible.")
                    return

                lineas_calculadas = conciliar_banco_con_tango(df_banco, df_tango, plan_lista)
                st.session_state.lineas_asiento_bancos = lineas_calculadas
                _aplicar_matcheo_pdf_a_grilla_banco(
                    slug,
                    lineas_calculadas,
                    ficha=ficha,
                    plan_df=plan_df,
                    periodo_mensual=periodo_mensual,
                    banco=banco,
                )
                st.success(
                    f"¡Cruce completado! Se detectaron {len(lineas_calculadas)} movimiento(s) reales. "
                    f"Abrí Conciliación Bancaria → **{banco}** para ver la grilla."
                )
                if not como_herramienta:
                    st.rerun()
            except Exception as exc:
                st.error(f"Error en el procesamiento de archivos: {exc}")


def _herramienta_matcheo_inteligente_pdf() -> None:
    """Herramienta: PDF extracto + subdiario Tango → grilla del banco (fuera de Conciliación)."""
    st.caption(
        "Matcheo inteligente de extracto PDF con el listado por imputación de Tango. "
        "Quedó acá para dejar más espacio a la grilla en Conciliación Bancaria."
    )

    clientes = db.listar_clientes()
    if not clientes:
        st.warning("Registrá al menos un cliente antes de usar esta herramienta.")
        return

    _selector_sociedad_devengamientos(clientes)
    if (
        st.session_state.cuit_activo is None
        or st.session_state.nombre_activo is None
        or st.session_state.get("plan_cuentas_cliente_id") != st.session_state.get(_SOCiedad_KEY)
    ):
        actualizar_sociedad_activa()

    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    if not sociedad_id:
        st.info("Seleccioná la sociedad activa.")
        return

    try:
        _sincronizar_plan_cuentas_session(int(sociedad_id), forzar=False)
    except Exception:
        pass

    plan_vinculado = (
        st.session_state.get("plan_cuentas_df") is not None
        and st.session_state.get("plan_cuentas_cliente_id") == sociedad_id
        and not bool(st.session_state.get("plan_cuentas_es_default", False))
    )
    if not plan_vinculado:
        st.warning("Vinculá el Plan de Cuentas de esta sociedad en Gestión de Clientes.")

    bancos = list(BANK_REGISTRY.keys())
    banco = st.selectbox(
        "Banco del extracto",
        bancos,
        index=bancos.index("Banco Galicia") if "Banco Galicia" in bancos else 0,
        key="herramienta_matcheo_banco_v1",
    )
    periodo_mensual = st.text_input(
        "Período (MM/AAAA)",
        value="",
        placeholder="12/2025",
        key="herramienta_matcheo_periodo_v1",
        help="Opcional: si lo dejás vacío se toma del PDF. Formato MM/AAAA.",
    ).strip() or None
    if periodo_mensual:
        periodo_mensual = formatear_periodo_mm_yyyy(periodo_mensual)

    ficha = obtener_ficha_banco(banco)
    slug = ficha.get("slug", _slug_banco(banco))
    _inicializar_sesion_banco(slug)

    _render_conexion_inteligente_extractos_pdf(
        banco,
        slug,
        ficha,
        int(sociedad_id),
        plan_vinculado,
        periodo_mensual,
        como_herramienta=True,
    )

    n_lineas = len(st.session_state.get("lineas_asiento_bancos") or [])
    n_grilla = len(st.session_state.get(f"{slug}_grilla_preview") or [])
    if n_grilla or n_lineas:
        st.caption(
            f"Último matcheo en memoria: {n_lineas or n_grilla} línea(s) para **{banco}**. "
            "Revisalas en Conciliación Bancaria."
        )


def _seccion_conciliacion_bancaria_banco(
    banco: str,
    sociedad_id: int,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    """Flujo de conciliación bancaria guiado por BANK_REGISTRY (espejo de devengamientos)."""
    ficha = obtener_ficha_banco(banco)
    slug = ficha.get("slug", _slug_banco(banco))
    motor = ficha.get("motor", "banco")

    _inicializar_sesion_banco(slug)
    _aplicar_limpieza_formulario_mes_banco(slug)
    _render_panel_biblioteca_bancos(
        sociedad_id, banco, cuit_activo=cuit_activo, nombre_activo=nombre_activo,
    )
    if not st.session_state.get("balance_servidor_buffer_por_sociedad", {}).get(sociedad_id):
        _cargar_balance_proyecto_obligatorio_bancos(sociedad_id)

    rc = _iva_reset_counter()
    periodo_ok = st.session_state.pop(f"{slug}_mensaje_biblioteca_ok", None)
    if periodo_ok:
        st.success(
            f"✓ Asiento bancario de {periodo_ok} guardado con éxito. "
            "Ya podés cargar el siguiente mes."
        )

    st.markdown("---")
    _render_conexion_servidor_banco(sociedad_id, banco)
    st.caption(
        "El matcheo inteligente PDF + Tango está en **Herramientas → Matcheo inteligente PDF + Tango**, "
        "para dejar más espacio a la grilla del asiento."
    )

    solapa_activa = banco
    archivo_a_procesar, origen_planilla, etiqueta_planilla = _resolver_fuente_balance(
        None, sociedad_id,
    )
    if archivo_a_procesar is not None:
        try:
            buf_diag, _ = _abrir_planilla_iva(archivo_a_procesar)
            solapa_resuelta_ui = resolver_solapa_balance(buf_diag, banco)
            st.caption(f"📑 Solapa resuelta para **{banco}**: `{solapa_resuelta_ui}`")
        except ValueError:
            hojas = _solapas_disponibles_archivo(archivo_a_procesar)
            st.warning(
                f"⚠️ No se encontró una solapa que coincida con el banco seleccionado. "
                f"Las solapas disponibles en este Excel son: "
                f"**{', '.join(hojas) if hojas else '(ninguna)'}**"
            )
        if origen_planilla == "servidor":
            st.success(
                f"📁 Usando Balance cargado desde **servidor local** "
                f"(`{etiqueta_planilla}`) — solapa **{solapa_activa}**."
            )

    periodo_mensual: str | None = None
    periodo_key = _clave_periodo_mensual_banco(slug, sociedad_id)
    if archivo_a_procesar is not None:
        buf_periodos, es_csv_periodos = _abrir_planilla_iva(archivo_a_procesar)
        periodos_disponibles = listar_periodos_disponibles_balance(
            buf_periodos, solapa_activa, es_csv=es_csv_periodos,
        )
        idx_default = 0
        periodo_previo = st.session_state.get(periodo_key)
        if periodo_previo and periodo_previo in periodos_disponibles:
            idx_default = periodos_disponibles.index(periodo_previo)
        periodo_mensual = st.selectbox(
            "📅 Período Mensual a Procesar",
            periodos_disponibles,
            index=idx_default,
            key=periodo_key,
        )

    fecha_key = _clave_fecha_tango_banco(slug, rc, periodo_mensual)
    if periodo_mensual:
        _asegurar_fecha_asiento_ultimo_dia(fecha_key, periodo_mensual)
    elif archivo_a_procesar is not None:
        _prefijar_fecha_input_desde_planilla_banco(archivo_a_procesar, slug, sociedad_id)
    st.date_input(
        "Fecha del Asiento para Tango",
        format="DD/MM/YYYY",
        key=fecha_key,
        help=(
            "Fuente de verdad para Tango. Se fija al último día del período "
            "mensual seleccionado (podés ajustarla a mano si hace falta)."
        ),
    )

    if archivo_a_procesar is not None and not plan_vinculado:
        st.error(
            "No podés procesar la conciliación hasta vincular el Plan de Cuentas de esta sociedad "
            "en el Editor de Clientes."
        )

    puede_procesar = archivo_a_procesar is not None and plan_vinculado and bool(periodo_mensual)
    if puede_procesar:
        _generar_asiento_automatico(
            slug=slug,
            motor=motor,
            impuesto=banco,
            ficha=ficha,
            sociedad_id=sociedad_id,
            archivo=archivo_a_procesar,
            solapa=solapa_activa,
            periodo_mensual=periodo_mensual,
            inputs_manuales={},
            saldos_contingencia={},
            origen_planilla=origen_planilla or "",
        )

    _render_resultados_conciliacion_bancaria(
        banco, slug, ficha, sociedad_id, cuit_activo, nombre_activo,
    )


def _seccion_conciliacion_bancaria_balance() -> None:
    with st.container():
        st.caption(
            "Extracción mensual desde el Balance por solapa de banco. El motor de coordenadas "
            "Debe/Haber congela columnas mellizas y genera asientos listos para Tango (PES, Ingresado)."
        )

        clientes = db.listar_clientes()
        clientes_pj = [c for c in clientes if c.get("tipo_persona") == "Persona Jurídica"]
        if not clientes_pj:
            st.warning("Debe registrar al menos una Persona Jurídica para conciliar bancos.")
            return

        indice = _indice_sociedades_pj(clientes_pj)

        col_sociedad, col_banco = st.columns(2)
        with col_sociedad:
            _selector_sociedad_devengamientos(clientes_pj)
        with col_banco:
            banco_elegido = st.selectbox(
                "🏦 Banco a Conciliar",
                listar_bancos_conciliacion(),
                key=_BANCO_KEY,
            )

        sociedad_id = st.session_state.get(_SOCiedad_KEY)
        if sociedad_id is None:
            return

        _detectar_cambio_banco_y_flush()

        if (
            st.session_state.cuit_activo is None
            or st.session_state.nombre_activo is None
            or st.session_state.get("plan_cuentas_cliente_id") != sociedad_id
        ):
            actualizar_sociedad_activa()

        _verificar_sincronizacion_devengamientos(indice)

        plan_vinculado = _sociedad_tiene_plan_vinculado_por_session()
        with col_sociedad:
            if plan_vinculado:
                st.success(
                    "✓ Sociedad vinculada: El Plan de Cuentas está asociado correctamente en el backend."
                )
            else:
                st.error(_mensaje_plan_no_vinculado())
                _widget_subir_plan_inline(sociedad_id, st.session_state.get("cuit_activo"), key_suffix="conc")

        cuit_activo = st.session_state.cuit_activo
        nombre_activo = st.session_state.nombre_activo

        _seccion_conciliacion_bancaria_banco(
            banco_elegido, sociedad_id, cuit_activo, nombre_activo, plan_vinculado,
        )


def _seccion_devengamientos_iibb(
    sociedad_id: int,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    """Wrapper legacy → motor unificado TAX_REGISTRY."""
    _seccion_devengamientos_impuesto(
        "Ingresos Brutos", sociedad_id, cuit_activo, nombre_activo, plan_vinculado,
    )


def _seccion_devengamientos() -> None:
    st.caption(
        "Refundición mensual por impuesto. Seleccioná sociedad e impuesto, "
        "cargá el Balance desde el servidor local (UNC) o subilo como respaldo; el sistema lee "
        "la solapa del impuesto activo y completa los saldos manuales con el plan de cuentas vinculado."
    )

    clientes = db.listar_clientes()
    clientes_pj = [c for c in clientes if c.get("tipo_persona") == "Persona Jurídica"]
    if not clientes_pj:
        st.warning("Debe registrar al menos una Persona Jurídica para generar devengamientos.")
        return

    indice = _indice_sociedades_pj(clientes_pj)

    col_sociedad, col_impuesto = st.columns(2)
    with col_sociedad:
        _selector_sociedad_devengamientos(clientes_pj)
    with col_impuesto:
        impuesto_elegido = st.selectbox(
            "Impuesto",
            list(TAX_REGISTRY.keys()),
            key=_IMPUESTO_KEY,
        )

    sociedad_id = st.session_state.get(_SOCiedad_KEY)
    if sociedad_id is None:
        return

    _detectar_cambio_impuesto_y_flush()

    if (
        st.session_state.cuit_activo is None
        or st.session_state.nombre_activo is None
        or st.session_state.get("plan_cuentas_cliente_id") != sociedad_id
    ):
        actualizar_sociedad_activa()

    _verificar_sincronizacion_devengamientos(indice)

    plan_vinculado = _sociedad_tiene_plan_vinculado_por_session()
    with col_sociedad:
        if plan_vinculado:
            st.success(
                "✓ Sociedad vinculada: El Plan de Cuentas está asociado correctamente en el backend."
            )
        else:
            st.error(_mensaje_plan_no_vinculado())
            _widget_subir_plan_inline(sociedad_id, st.session_state.get("cuit_activo"), key_suffix="dev")

    cuit_activo = st.session_state.cuit_activo
    nombre_activo = st.session_state.nombre_activo

    # #region agent log
    _dbg_log("D", "_seccion_devengamientos", "render", {
        "sociedad_id": sociedad_id,
        "impuesto": st.session_state.get(_IMPUESTO_KEY),
        "cuit_activo": cuit_activo,
        "nombre_activo": nombre_activo,
        "plan_vinculado": plan_vinculado,
        "plan_rows": len(st.session_state.plan_cuentas_df) if st.session_state.get("plan_cuentas_df") is not None else 0,
        "plan_cliente_id": st.session_state.get("plan_cuentas_cliente_id"),
    })
    # #endregion

    _seccion_devengamientos_impuesto(
        impuesto_elegido, sociedad_id, cuit_activo, nombre_activo, plan_vinculado,
    )

    historial = db.listar_devengamientos(sociedad_id)
    if historial:
        with st.expander("Historial de devengamientos guardados"):
            st.dataframe(pd.DataFrame(historial), use_container_width=True, hide_index=True)



def _formulario_saldos_iniciales_bancos() -> list[dict]:
    """Formulario dinámico de saldos iniciales pasivos por banco (al 01/01)."""
    st.markdown("#### Saldos iniciales pasivos por banco (al 01/01)")
    st.caption("Indique el saldo inicial de préstamos bancarios en pesos para cada entidad.")

    saldos = st.session_state.saldos_iniciales_bancos
    filas_actualizadas: list[dict] = []

    for idx, fila in enumerate(saldos):
        col_banco, col_saldo, col_del = st.columns([3, 2, 1])
        with col_banco:
            banco_sel = st.selectbox(
                "Banco",
                BANCOS_ARGENTINOS,
                index=BANCOS_ARGENTINOS.index(fila["banco"]) if fila["banco"] in BANCOS_ARGENTINOS else 0,
                key=f"saldo_banco_{idx}",
                label_visibility="collapsed",
            )
        with col_saldo:
            saldo_val = st.number_input(
                "Saldo Inicial en Pesos",
                min_value=0.0,
                value=float(fila.get("saldo_inicial", 0)),
                step=1000.0,
                format="%.2f",
                key=f"saldo_valor_{idx}",
                label_visibility="collapsed",
            )
        with col_del:
            if len(saldos) > 1:
                if st.button("🗑️", key=f"del_saldo_{idx}", help="Eliminar fila"):
                    saldos.pop(idx)
                    st.session_state.saldos_iniciales_bancos = saldos
                    st.rerun()
            else:
                st.write("")

        filas_actualizadas.append({"banco": banco_sel, "saldo_inicial": saldo_val})

    st.session_state.saldos_iniciales_bancos = filas_actualizadas

    if st.button("➕ Agregar banco", key="add_saldo_banco"):
        st.session_state.saldos_iniciales_bancos.append(
            {"banco": BANCOS_ARGENTINOS[0], "saldo_inicial": 0.0}
        )
        st.rerun()

    return filas_actualizadas


def _herramienta_pdf_extractos_a_excel() -> None:
    """Convierte extractos multi-banco: un Excel por banco (ZIP si hay varios)."""
    bancos_txt = ", ".join(
        sorted(
            {
                str(p.get("nombre_display") or k)
                for k, p in PERFILES_BANCO.items()
                if k != "desconocido"
            }
        )
    )
    st.markdown("#### PDF extractos → Excel")
    st.caption(
        "Subí uno o varios extractos PDF (digital o escaneado). "
        "Se detecta el banco automáticamente y **todos** salen con el **mismo formato universal**: "
        "`Fecha | Descripcion | Detalle | Importe | Saldo | Clasificacion | Nueva_Clasificacion | Tipo Movimiento` "
        "+ hoja `Resumen_Clasificacion`. "
        "Importe con signo: **(−) resta** · **(+) suma**. Un Excel por banco. "
        f"Soportados: {bancos_txt}."
    )
    pdfs_ext = st.file_uploader(
        "Extractos bancarios (PDF)",
        type=["pdf"],
        accept_multiple_files=True,
        key="uploader_herramientas_extractos_bancarios",
        help="Podés mezclar bancos: cada uno sale en su propio Excel. Varios meses del mismo banco sí se unifican.",
    )
    if st.button("Convertir a Excel", type="primary", key="btn_herramientas_extracto_excel"):
        if not pdfs_ext:
            st.warning("Subí al menos un PDF para convertir.")
        else:
            with st.spinner(
                "Detectando banco(s) y leyendo PDFs… Si está escaneado, OCR página por página "
                "(puede tardar). No cierres la ventana."
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
                    paquetes_ui.append({
                        "banco": p.get("banco") or meta_b.get("banco") or "Banco",
                        "banco_slug": p.get("banco_slug") or "",
                        "df": df_b,
                        "meta": meta_b,
                        "xlsx": xlsx_b,
                        "pdf_merged": p.get("pdf_merged"),
                    })
                st.session_state.extracto_paquetes = paquetes_ui
                if len(paquetes_ui) > 1:
                    st.session_state.extracto_zip = exportar_zip_extractos_por_banco(
                        paquetes_ui,
                        cuit=str(meta_ext.get("cuit") or st.session_state.get("cuit_activo") or ""),
                    )
                else:
                    st.session_state.extracto_zip = None

                if paquetes_ui:
                    # Compatibilidad: primer banco como xlsx "principal"
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
        with st.expander(f"Advertencias ({len(err_ext)})", expanded=True):
            st.dataframe(pd.DataFrame(err_ext), use_container_width=True, hide_index=True)

    paquetes = st.session_state.get("extracto_paquetes") or []
    df_ext = st.session_state.get("extracto_santander_df")
    meta_ext = st.session_state.get("extracto_santander_meta") or {}
    if not paquetes or df_ext is None or df_ext.empty:
        return

    n_mov = int((df_ext["Tipo fila"] == "Movimiento").sum()) if "Tipo fila" in df_ext.columns else len(df_ext)
    n_bancos = len(paquetes)
    bancos_txt_ok = " · ".join(p["banco"] for p in paquetes)
    st.success(
        f"Listo: **{n_mov}** movimientos · **{n_bancos}** banco(s) · "
        f"**{len(meta_ext.get('archivos') or [])}** PDF(s)"
        + (f" · {bancos_txt_ok}" if bancos_txt_ok else "")
    )
    if n_bancos > 1:
        st.info(
            "Se generó **un Excel por banco**. Los meses se unifican solo dentro del mismo banco."
        )

    cuit_limpio = re.sub(
        r"\D",
        "",
        str(meta_ext.get("cuit") or st.session_state.get("cuit_activo") or ""),
    )
    zip_bytes = st.session_state.get("extracto_zip")
    if zip_bytes and n_bancos > 1:
        st.download_button(
            f"Descargar ZIP ({n_bancos} Excel, uno por banco)",
            data=zip_bytes,
            file_name=f"Extractos_por_banco_{cuit_limpio or 'cliente'}_{date.today().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            type="primary",
            key="dl_herramientas_extracto_zip",
            use_container_width=True,
        )

    for i, p in enumerate(paquetes):
        df_b = p["df"]
        meta_b = p.get("meta") or {}
        n_mov_b = int((df_b["Tipo fila"] == "Movimiento").sum()) if "Tipo fila" in df_b.columns else len(df_b)
        n_meses_b = int(df_b["Mes"].nunique()) if "Mes" in df_b.columns else 0
        banco = p.get("banco") or "Banco"
        with st.expander(
            f"{banco}: {n_mov_b} mov. · {n_meses_b} mes(es) · {len(meta_b.get('archivos') or [])} PDF(s)",
            expanded=(n_bancos == 1),
        ):
            if meta_b.get("cliente") or meta_b.get("cuit") or meta_b.get("formato"):
                partes_cap = []
                if meta_b.get("cliente") or meta_b.get("cuit"):
                    partes_cap.append(
                        f"Cliente: **{meta_b.get('cliente') or '—'}** · "
                        f"CUIT: `{meta_b.get('cuit') or '—'}` · "
                        f"Cuenta: `{meta_b.get('cuenta') or '—'}`"
                    )
                if meta_b.get("formato"):
                    partes_cap.append(f"Formato detectado: **{meta_b.get('formato')}**")
                st.caption(" · ".join(partes_cap))
            cols_preview = [
                c for c in (
                    "Fecha", "Descripcion", "Detalle", "Importe", "Saldo",
                    "Clasificacion", "Tipo Movimiento",
                ) if c in df_b.columns
            ]
            if "Clasificacion" not in df_b.columns or "Tipo Movimiento" not in df_b.columns:
                df_prev = enriquecer_df_extracto_formato_banco(df_b)
                cols_preview = [
                    c for c in (
                        "Fecha", "Descripcion", "Detalle", "Importe", "Saldo",
                        "Clasificacion", "Tipo Movimiento",
                    ) if c in df_prev.columns
                ]
            else:
                df_prev = df_b
            st.info(
                "Formato banco: **Importe** con signo (− resta / + suma). "
                "**Tipo Movimiento** = Debito/Credito según el signo. "
                "Podés completar **Nueva_Clasificacion** en el Excel."
            )
            st.dataframe(df_prev[cols_preview].head(40), use_container_width=True, hide_index=True)
            slug_banco = re.sub(r"[^A-Za-z0-9]+", "_", str(banco).strip())[:24] or "banco"
            stamp = datetime.now().strftime("%Y-%m-%d_%H %M")
            nombre_xlsx = f"{stamp}_{slug_banco}.xlsx"
            st.download_button(
                f"Descargar Excel — {banco}",
                data=p["xlsx"],
                file_name=nombre_xlsx,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" if n_bancos == 1 else "secondary",
                key=f"dl_herramientas_extracto_xlsx_{i}_{slug_banco}",
                use_container_width=True,
            )
            pdf_merged = p.get("pdf_merged")
            archivos_b = meta_b.get("archivos") or []
            if pdf_merged and len(archivos_b) > 1:
                st.download_button(
                    f"Descargar PDF unificado — {banco}",
                    data=pdf_merged,
                    file_name=f"Extractos_{slug_banco}_{cuit_limpio or 'cliente'}.pdf",
                    mime="application/pdf",
                    key=f"dl_herramientas_extracto_pdf_{i}_{slug_banco}",
                    use_container_width=True,
                )


def _herramienta_completar_cuadro_bancario() -> None:
    """Completa cuadros mensuales: buzón por carpeta UNC o carga manual."""
    st.markdown("#### Completar cuadro bancario existente")
    st.caption(
        "Motor con **cadena de saldos** (como en Trujillo): el saldo del extracto "
        "corrige signos/montos del OCR. Formato estándar único "
        "(meses / créditos / débitos / fórmulas DIF). "
        "Si el cliente no tiene el cuadro, se genera `CUADROS_BANCARIOS_AAAA.xlsx` limpio."
    )
    st.info(
        "Si el cliente ya tiene el cuadro estándar, se completa ahí. "
        "Si no, se genera un Excel nuevo `CUADROS_BANCARIOS_AAAA.xlsx` "
        "(solo hojas de cuadro + control), sin mezclar la planilla de Ganancias. "
        "Los `.xlsm` conservan macros solo cuando se edita el libro original."
    )
    ejercicio = st.number_input(
        "Ejercicio a completar",
        min_value=2000,
        max_value=2100,
        value=max(2000, date.today().year - 1),
        step=1,
        key="cuadro_bancario_ejercicio_v1",
        help="Se ignoran movimientos de otros años (p. ej. extractos trimestrales).",
    )

    st.markdown("##### Buzón de carpeta")
    ruta_buzon = st.text_input(
        "Ruta de la carpeta del cliente",
        key="cuadro_bancario_ruta_buzon_v1",
        placeholder=r"\\TANGOSRV\Compartido\CLIENTES\TRUJILLO HERNAN\Ganancias Personas Fisicas\GANANCIAS 2025",
        help="Pegá la carpeta (Copiar como ruta). Se buscan Excel y PDF en subcarpetas.",
    )
    col_explorar, _ = st.columns([1, 2])
    with col_explorar:
        explorar = st.button(
            "Explorar carpeta",
            key="cuadro_bancario_explorar_v1",
            use_container_width=True,
        )

    if explorar:
        st.session_state.cuadro_bancario_resultado = None
        st.session_state.cuadro_bancario_resultado_signature = None
        ruta_limpia = sanitizar_ruta_unc(ruta_buzon)
        if not ruta_limpia:
            st.warning("Pegá la ruta de la carpeta.")
        else:
            try:
                buzon = explorar_buzon_cuadros_bancarios(
                    ruta_limpia,
                    ejercicio=int(ejercicio),
                )
                st.session_state.cuadro_bancario_buzon = buzon
                st.session_state.cuadro_bancario_buzon_ruta = buzon["carpeta"]
            except Exception as exc:
                st.session_state.cuadro_bancario_buzon = None
                st.session_state.cuadro_bancario_buzon_ruta = None
                st.error(f"No pude explorar la carpeta: {exc}")
                return

    buzon = st.session_state.get("cuadro_bancario_buzon")
    excel_destino = None
    pdfs_upload = None
    excel_ruta = None
    pdfs_rutas: list[str] = []

    if buzon:
        st.success(
            f"Carpeta: `{buzon.get('carpeta')}` · "
            f"**{len(buzon.get('excels') or [])}** Excel · "
            f"**{len(buzon.get('pdfs') or [])}** PDF"
        )
        for adv in buzon.get("advertencias") or []:
            st.warning(adv)

        excels = buzon.get("excels") or []
        pdfs_hallados = buzon.get("pdfs") or []
        if not excels:
            st.error(
                "En esa carpeta hay extractos PDF, pero **no hay un Excel de Ganancias** "
                "para completar. Subí el cuadro Excel abajo (carga manual) o pegá una "
                "carpeta que sí lo tenga (por ejemplo la de Banco Galicia / Ganancia 2025)."
            )
        if excels:
            etiquetas = {
                item["ruta"]: (
                    f"{'★ ' if item.get('sugerido') else ''}"
                    f"{item['relativa']}  ({item['size'] // 1024} KB)"
                )
                for item in excels
            }
            default_excel = buzon.get("excel_sugerido") or excels[0]["ruta"]
            excel_ruta = st.selectbox(
                "Excel destino (detectado)",
                options=list(etiquetas.keys()),
                index=list(etiquetas.keys()).index(default_excel)
                if default_excel in etiquetas
                else 0,
                format_func=lambda r: etiquetas.get(r, r),
                key="cuadro_bancario_excel_sel_v1",
            )
        if pdfs_hallados:
            etiquetas_pdf = {
                item["ruta"]: f"{item['relativa']}  ({item['size'] // 1024} KB)"
                for item in pdfs_hallados
            }
            default_pdfs = [
                item["ruta"] for item in pdfs_hallados if item.get("sugerido", True)
            ] or list(etiquetas_pdf.keys())
            n_omitidos = len(pdfs_hallados) - len(default_pdfs)
            if n_omitidos > 0:
                st.caption(
                    f"Dejé afuera {n_omitidos} PDF que no parecen extractos bancarios "
                    "(tarjetas, compras/ventas, facturas). Podés sumarlos igual abajo."
                )
            pdfs_rutas = st.multiselect(
                "Extractos PDF a usar",
                options=list(etiquetas_pdf.keys()),
                default=default_pdfs,
                format_func=lambda r: etiquetas_pdf.get(r, r),
                key="cuadro_bancario_pdfs_sel_v1",
            )

    # Si exploramos carpeta y hay PDF pero no Excel, abrir carga manual del Excel.
    abrir_manual = (not bool(buzon)) or (bool(buzon) and not (buzon.get("excels") or []))
    with st.expander("Carga manual (opcional, si no usás carpeta)", expanded=abrir_manual):
        col_excel, col_pdfs = st.columns(2)
        with col_excel:
            excel_destino = st.file_uploader(
                "Excel destino (.xlsx o .xlsm)",
                type=["xlsx", "xlsm"],
                accept_multiple_files=False,
                key="cuadro_bancario_excel_destino_v1",
                help="Obligatorio si la carpeta no tiene el cuadro de Ganancias.",
            )
        with col_pdfs:
            pdfs_upload = st.file_uploader(
                "Extractos bancarios (PDF)",
                type=["pdf"],
                accept_multiple_files=True,
                key="cuadro_bancario_pdfs_v1",
            )

    # Carpeta con Excel+PDF, o híbrido: Excel subido + PDF de la carpeta.
    if excel_ruta and pdfs_rutas:
        origen = "carpeta"
    elif excel_destino is not None and pdfs_rutas:
        origen = "hibrido"
    elif excel_destino is not None and pdfs_upload:
        origen = "upload"
    else:
        origen = "incompleto"

    if origen == "carpeta":
        input_signature = (
            "carpeta",
            int(ejercicio),
            excel_ruta,
            tuple(sorted(pdfs_rutas)),
        )
    elif origen == "hibrido":
        input_signature = (
            "hibrido",
            int(ejercicio),
            (excel_destino.name, getattr(excel_destino, "size", 0)),
            tuple(sorted(pdfs_rutas)),
        )
    else:
        input_signature = (
            "upload",
            int(ejercicio),
            (excel_destino.name, getattr(excel_destino, "size", 0)) if excel_destino else None,
            tuple((a.name, getattr(a, "size", 0)) for a in (pdfs_upload or [])),
        )

    if st.button(
        "Completar y controlar Excel",
        type="primary",
        key="cuadro_bancario_procesar_v1",
        use_container_width=True,
    ):
        try:
            if origen == "incompleto":
                if buzon and pdfs_rutas and not excel_ruta and excel_destino is None:
                    st.error(
                        "Los PDF de la carpeta están listos, pero **falta el Excel destino**. "
                        "En esa carpeta no hay un archivo Ganancias (.xlsx/.xlsm). "
                        "Subilo en «Carga manual» o explorá otra carpeta que lo tenga."
                    )
                elif buzon and not pdfs_rutas and not pdfs_upload:
                    st.warning("Seleccioná al menos un PDF de la carpeta o subí extractos.")
                elif excel_destino is None and not excel_ruta:
                    st.warning("Falta el Excel destino (cuadro de Ganancias).")
                else:
                    st.warning("Faltan extractos PDF.")
                return

            if origen == "carpeta":
                excel_path = Path(excel_ruta)
                excel_bytes = excel_path.read_bytes()
                excel_nombre = excel_path.name
                # Usar ruta relativa (BBVA\Enero.pdf) para detectar cuenta/carpeta.
                root_buzon = Path(
                    st.session_state.get("cuadro_bancario_buzon_ruta")
                    or Path(excel_ruta).parent
                )
                pares_pdf = []
                for r in pdfs_rutas:
                    p = Path(r)
                    try:
                        etiqueta = str(p.relative_to(root_buzon))
                    except Exception:
                        etiqueta = p.name
                    pares_pdf.append((etiqueta, p))
            elif origen == "hibrido":
                excel_bytes = excel_destino.getvalue()
                excel_nombre = excel_destino.name
                root_buzon = Path(
                    st.session_state.get("cuadro_bancario_buzon_ruta")
                    or "."
                )
                pares_pdf = []
                for r in pdfs_rutas:
                    p = Path(r)
                    try:
                        etiqueta = str(p.relative_to(root_buzon))
                    except Exception:
                        etiqueta = p.name
                    pares_pdf.append((etiqueta, p))
            else:
                excel_bytes = excel_destino.getvalue()
                excel_nombre = excel_destino.name
                pares_pdf = [(a.name, a.getvalue()) for a in pdfs_upload]

            with st.spinner(
                f"Leyendo {len(pares_pdf)} extracto(s) de a uno (más estable)… "
                "Los PDF escaneados pueden tardar por el OCR. No cierres la pestaña."
            ):
                resultado = completar_cuadro_bancario_existente(
                    excel_bytes,
                    excel_nombre,
                    pares_pdf,
                    ejercicio=int(ejercicio),
                )
                # Guardar en la carpeta del buzón (o junto al Excel origen).
                carpeta_destino = None
                if st.session_state.get("cuadro_bancario_buzon_ruta"):
                    carpeta_destino = Path(st.session_state.cuadro_bancario_buzon_ruta)
                elif origen == "carpeta" and excel_ruta:
                    carpeta_destino = Path(excel_ruta).parent
                if resultado.get("excel") and carpeta_destino and carpeta_destino.is_dir():
                    try:
                        nombre_out = str(
                            resultado.get("nombre") or "Cuadros_bancarios_completados.xlsx"
                        )
                        ruta_out = carpeta_destino / nombre_out
                        if ruta_out.exists():
                            stem = ruta_out.stem
                            suf = ruta_out.suffix
                            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            ruta_out = carpeta_destino / f"{stem}_{stamp}{suf}"
                            resultado["nombre"] = ruta_out.name
                        ruta_out.write_bytes(resultado["excel"])
                        resultado["ruta_guardada"] = str(ruta_out)
                    except Exception as exc_save:
                        resultado["ruta_guardada"] = None
                        resultado.setdefault("errores", []).append(
                            {
                                "archivo": resultado.get("nombre"),
                                "motivo": f"No pude guardar en la carpeta: {exc_save}",
                            }
                        )
                st.session_state.cuadro_bancario_resultado = resultado
                st.session_state.cuadro_bancario_resultado_signature = input_signature
        except Exception as exc:
            st.session_state.cuadro_bancario_resultado = {
                "excel": None,
                "resultados": [],
                "errores": [{"motivo": str(exc)}],
                "nombre": "error.xlsx",
                "modo": "error",
            }
            st.session_state.cuadro_bancario_resultado_signature = input_signature
        st.rerun()

    resultado = st.session_state.get("cuadro_bancario_resultado")
    if (
        resultado
        and st.session_state.get("cuadro_bancario_resultado_signature") != input_signature
    ):
        st.info("Cambiaste los archivos, la carpeta o el ejercicio. Volvé a procesar.")
        return
    if not resultado:
        return

    errores = resultado.get("errores") or []
    modo = resultado.get("modo") or "cuadros"
    if modo == "movimientos":
        st.warning(
            "No se pudo completar el cuadro estándar (revisá el control de saldos). "
            "Guardé un archivo con los **movimientos detectados** para no perder el trabajo."
        )
    elif any(
        "plantilla del estudio" in str(e.get("motivo") or "").lower()
        for e in errores
    ):
        st.info(
            "Se aplicó el **formato estándar del estudio** (meses / créditos / débitos / fórmulas) "
            "porque el Excel del cliente no lo tenía."
        )
    if errores:
        st.warning(
            "Hay archivos o cuentas que requieren revisión. "
            "Revisá el detalle abajo."
        )
        st.dataframe(pd.DataFrame(errores), use_container_width=True, hide_index=True)

    grupos = resultado.get("resultados") or []
    excel_salida = resultado.get("excel")
    if not grupos and not excel_salida:
        st.error("No se generó un Excel: no quedaron movimientos válidos.")
        return

    controles = []
    for grupo in grupos:
        for control in grupo.get("controles") or []:
            controles.append(
                {
                    "Banco": grupo.get("banco"),
                    "Cuenta": grupo.get("cuenta"),
                    "Hoja": grupo.get("hoja"),
                    **{k: v for k, v in control.items() if k != "Mes_num"},
                }
            )
    df_control = pd.DataFrame(controles)
    n_ok = int((df_control["Estado"] == "OK").sum()) if not df_control.empty else 0
    if modo == "cuadros":
        st.success(
            f"Excel preparado: **{len(grupos)} cuenta(s)** · "
            f"**{n_ok} mes(es) conciliados** · formato estándar · hoja `_CONTROL_BANCOS`."
        )
    else:
        st.success(
            f"Excel de movimientos: **{len(grupos)} cuenta(s)** · "
            f"**{n_ok} mes(es) OK** en el control."
        )
    if not df_control.empty:
        st.dataframe(df_control, use_container_width=True, hide_index=True)

    if excel_salida:
        nombre = str(resultado.get("nombre") or "Cuadros_bancarios_completados.xlsx")
        ruta_guardada = resultado.get("ruta_guardada")
        if ruta_guardada:
            st.success(f"Guardado en la carpeta del cliente:\n`{ruta_guardada}`")
        else:
            st.info(
                "No pude guardar en la carpeta del buzón (faltó la ruta o no es accesible). "
                "Usá la descarga de abajo."
            )
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if nombre.lower().endswith(".xlsm")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "Descargar Excel generado",
            data=excel_salida,
            file_name=nombre,
            mime=mime,
            type="primary",
            key="cuadro_bancario_descargar_v1",
            use_container_width=True,
        )


def _herramienta_match_debitos_proveedores() -> None:
    """Cruza débitos del extracto con facturas de proveedores (Excel del período)."""
    st.markdown("#### Match débitos - proveedores")
    st.caption(
        "Mismo motor que el asistente (chat): beneficiario en **Detalle**, "
        f"nombre truncado (ej. Trujillo), ventanas **±{TOL_MATCH_DIAS}d** con nombre / "
        f"**±{TOL_MATCH_DIAS_SOLO_MONTO}d** solo monto (único ±{TOL_MATCH_DIAS_MONTO_UNICO}d), "
        "sumas **2 transf ↔ 1 fct**, excluye FCC del propio banco. "
        "Solo transferencias/pagos (no comisiones, impuestos, haberes)."
    )

    extracto_default = Path(r"C:\Users\recep\Downloads\Extracto_Santander_unificado_3071802274_20260710.xlsx")
    extracto_marzo = Path(r"C:\Users\recep\Downloads\Extracto_Santander_2026-03_OCR.xlsx")
    if not extracto_marzo.exists():
        extracto_marzo = Path(r"C:\Users\recep\Desktop\Extracto_Santander_Oftalmologia_RELE_2026-03.xlsx")

    for err in st.session_state.get("match_prov_errores") or []:
        st.error(err)

    c1, c2 = st.columns(2)
    with c1:
        pdfs_match = st.file_uploader(
            "Extractos bancarios (PDF) — Galicia, Macro, Santander, etc.",
            type=["pdf"],
            accept_multiple_files=True,
            key="uploader_match_extractos_pdf",
            help="Detecta el banco de cada PDF (multi-banco). Opcional si usás Excel o sesión.",
        )
        xlsx_extracto = st.file_uploader(
            "O Excel de extracto ya convertido",
            type=["xlsx"],
            key="uploader_match_extracto_xlsx",
        )
        tiene_sesion = bool(
            st.session_state.get("extracto_santander_df") is not None
            and not getattr(st.session_state.get("extracto_santander_df"), "empty", True)
        )
        usar_sesion = st.checkbox(
            "Usar extracto de esta sesión (PDF→Excel)",
            value=tiene_sesion,
            key="chk_match_usar_extracto_sesion",
        )
        usar_extracto_analizado = st.checkbox(
            "Usar extractos ya analizados (unificado + marzo OCR)",
            value=(not tiene_sesion) and (extracto_default.exists() or extracto_marzo.exists()),
            key="chk_match_usar_extractos_analizados",
            help="Toma los Excel generados previamente sin volver a subir archivos.",
        )
    with c2:
        xlsx_prov = st.file_uploader(
            "Excel proveedores / facturas del período",
            type=["xlsx"],
            key="uploader_match_proveedores",
            help=(
                "Exportá de Tango el listado por imputación contable (21101) "
                f"del período a cruzar. Opcional por defecto: {PROVEEDORES_DEFAULT_PATH}"
            ),
        )
        usar_default_prov = st.checkbox(
            f"Usar archivo por defecto ({PROVEEDORES_DEFAULT_PATH.name})",
            value=xlsx_prov is None and PROVEEDORES_DEFAULT_PATH.exists(),
            key="chk_match_prov_default",
        )
        if xlsx_prov is not None:
            st.caption(f"Usando archivo subido: `{getattr(xlsx_prov, 'name', 'proveedores.xlsx')}`")
        elif usar_default_prov and PROVEEDORES_DEFAULT_PATH.exists():
            st.caption(f"Proveedores: `{PROVEEDORES_DEFAULT_PATH}`")
        elif not PROVEEDORES_DEFAULT_PATH.exists():
            st.warning("Subí el Excel de proveedores del período (no hay archivo por defecto en T:).")
        else:
            st.info("Subí el Excel del período o marcá el archivo por defecto.")

    if st.button("Matchear débitos con facturas", type="primary", key="btn_match_debitos_prov"):
        df_ext = None
        origen_ext = ""
        errores: list[str] = []
        st.session_state.match_prov_errores = []

        with st.spinner("Leyendo extracto y facturas, matcheando débitos..."):
            if usar_sesion and st.session_state.get("extracto_santander_df") is not None:
                df_ext = st.session_state.extracto_santander_df.copy()
                origen_ext = "sesión PDF→Excel"
            elif pdfs_match:
                # Multi-banco (Galicia + Macro + Santander…), mismo parser que el chat
                df_ext, meta_ext, err_ext = procesar_extractos_bancarios_pdfs(pdfs_match)
                origen_ext = ", ".join(meta_ext.get("archivos") or [f.name for f in pdfs_match])
                bancos = meta_ext.get("bancos") or (
                    [meta_ext.get("banco")] if meta_ext.get("banco") else []
                )
                if bancos:
                    origen_ext = f"{origen_ext} [{', '.join(str(b) for b in bancos if b)}]"
                if err_ext:
                    errores.extend(f"{e.get('archivo')}: {e.get('motivo')}" for e in err_ext)
            elif xlsx_extracto is not None:
                try:
                    df_ext = pd.read_excel(xlsx_extracto, sheet_name="Movimientos")
                except Exception:
                    df_ext = pd.read_excel(xlsx_extracto)
                origen_ext = getattr(xlsx_extracto, "name", "extracto.xlsx")
            elif usar_extracto_analizado:
                frames = []
                nombres = []
                for path in (extracto_default, extracto_marzo):
                    if not path.exists():
                        continue
                    try:
                        parte = pd.read_excel(path, sheet_name="Movimientos")
                    except Exception:
                        parte = pd.read_excel(path)
                    frames.append(parte)
                    nombres.append(path.name)
                if frames:
                    df_ext = pd.concat(frames, ignore_index=True)
                    origen_ext = " + ".join(nombres)
                else:
                    errores.append("No se encontraron los Excel de extractos ya analizados.")
            else:
                errores.append("Falta extracto: marcá 'extractos ya analizados', subí PDF/Excel o usá el de la sesión.")

            origen_prov = ""
            fuente_prov = None
            if xlsx_prov is not None:
                fuente_prov = xlsx_prov
                origen_prov = getattr(xlsx_prov, "name", "proveedores.xlsx")
            elif usar_default_prov and PROVEEDORES_DEFAULT_PATH.exists():
                fuente_prov = PROVEEDORES_DEFAULT_PATH
                origen_prov = str(PROVEEDORES_DEFAULT_PATH)
            else:
                errores.append(
                    "Falta Excel de proveedores del período: subilo "
                    "(listado por imputación 21101 exportado de Tango)."
                )
            if df_ext is None or (hasattr(df_ext, "empty") and df_ext.empty):
                if "No hay movimientos" not in " ".join(errores):
                    errores.append("No hay movimientos en el extracto.")

            if errores or fuente_prov is None:
                st.session_state.match_prov_errores = errores or ["Falta Excel de proveedores."]
                st.session_state.match_prov_resultado = None
                st.session_state.match_prov_xlsx = None
            else:
                try:
                    resultado, meta_pipe = ejecutar_match_debitos_proveedores(
                        df_ext,
                        fuente_prov,
                        priorizar_fecha_monto=False,
                        excluir_cargos_banco=True,
                    )
                    meta = {
                        "cliente": st.session_state.get("nombre_activo")
                        or (st.session_state.get("extracto_santander_meta") or {}).get("cliente")
                        or "",
                        "cuit": st.session_state.get("cuit_activo") or "",
                        "origen_extracto": origen_ext,
                        "origen_facturas": origen_prov,
                        "pipeline": meta_pipe.get("pipeline"),
                        "tol_dias_nombre": meta_pipe.get("tol_dias_nombre"),
                        "tol_dias_monto": meta_pipe.get("tol_dias_monto"),
                    }
                    st.session_state.match_prov_resultado = resultado
                    st.session_state.match_prov_meta = meta
                    st.session_state.match_prov_xlsx = exportar_match_proveedores_excel(resultado, meta)
                    st.session_state.match_prov_errores = []
                except Exception as exc:
                    st.session_state.match_prov_errores = [f"Error al matchear: {exc}"]
                    st.session_state.match_prov_resultado = None
                    st.session_state.match_prov_xlsx = None
        st.rerun()

    res = st.session_state.get("match_prov_resultado")
    xlsx_out = st.session_state.get("match_prov_xlsx")
    if not res or not xlsx_out:
        return

    meta = st.session_state.get("match_prov_meta") or {}
    if meta.get("origen_extracto") or meta.get("origen_facturas"):
        st.success(
            f"Match OK · extracto: {meta.get('origen_extracto') or '—'} · "
            f"facturas: {meta.get('origen_facturas') or '—'} · "
            f"ventanas ±{meta.get('tol_dias_nombre', TOL_MATCH_DIAS)}d / "
            f"±{meta.get('tol_dias_monto', TOL_MATCH_DIAS_SOLO_MONTO)}d"
        )

    resumen = res.get("resumen")
    if resumen is not None and not resumen.empty:
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    calz = res.get("calzados")
    if calz is not None and not calz.empty:
        st.markdown("**Calzados** (muestra)")
        st.dataframe(calz.head(30), use_container_width=True, hide_index=True)

    cuit_dl = re.sub(
        r"\D",
        "",
        str(
            (st.session_state.get("match_prov_meta") or {}).get("cuit")
            or st.session_state.get("cuit_activo")
            or ""
        ),
    ) or "cliente"
    st.download_button(
        "Descargar Excel del match",
        data=xlsx_out,
        file_name=f"Match_Debitos_Proveedores_{cuit_dl}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_match_debitos_prov",
        type="primary",
        use_container_width=True,
    )


def _herramienta_liquidaciones_tarjeta() -> None:
    """Convertidor de liquidaciones: flujo simple para evitar crashes de Streamlit."""
    st.caption(
        "Subí PDFs de liquidaciones (CABAL, First Data, Naranja, Prisma, Mercado Pago, etc.). "
        "El sistema detecta la entidad; después podés corregirla en la tabla."
    )

    archivos = st.file_uploader(
        "Liquidaciones de tarjeta (PDF)",
        type=["pdf"],
        accept_multiple_files=True,
        key="liq_uploader_v3",
    )
    procesar = st.button(
        "Procesar liquidaciones",
        type="primary",
        key="liq_btn_procesar_v3",
        disabled=not bool(archivos),
    )

    if procesar and archivos:
        filas = []
        textos = {}
        with st.spinner(f"Procesando {len(archivos)} PDF(s)…"):
            for i, archivo in enumerate(archivos):
                nombre = str(getattr(archivo, "name", f"archivo_{i + 1}.pdf"))
                texto = extraer_texto_liquidacion_pdf(archivo)
                sugerida = detectar_entidad_por_texto(f"{texto}\n{nombre}")
                if sugerida not in PLANTILLAS_TARJETAS:
                    sugerida = "Otra / No detectada"
                valores = extraer_con_plantilla(texto, sugerida)
                valores["Archivo"] = nombre
                valores["Entidad"] = (
                    sugerida if sugerida != "Otra / No detectada" else "Desconocida"
                )
                valores["Entidad_detectada"] = sugerida
                filas.append(valores)
                textos[nombre] = texto
        st.session_state["liq_v3_filas"] = filas
        st.session_state["liq_v3_textos"] = textos
        st.session_state["liq_v3_xlsx"] = None

    filas = st.session_state.get("liq_v3_filas")
    if not filas:
        return

    opciones = list(PLANTILLAS_TARJETAS.keys()) + ["Otra / No detectada"]
    df = pd.DataFrame(filas)
    cols_orden = [
        c for c in (
            "Archivo", "Entidad_detectada", "Entidad", "Fecha", "Nro_Liquidacion",
            "Neto_Gravado", "IVA_21", "Percepcion_IVA", "Retencion_IVA",
            "Retencion_IIBB", "Percepcion_IIBB", "Total_Descontado",
        ) if c in df.columns
    ]
    df = df[cols_orden].copy()

    editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        key="liq_v3_editor",
        disabled=[c for c in df.columns if c != "Entidad"],
        column_config={
            "Entidad": st.column_config.SelectboxColumn(
                "Entidad",
                options=opciones,
                required=True,
            ),
        },
    )

    c1, c2 = st.columns(2)
    with c1:
        recalc = st.button("Recalcular con entidades confirmadas", key="liq_v3_recalc")
    with c2:
        gen = st.button("Generar Excel consolidado", type="primary", key="liq_v3_excel")

    if recalc or gen:
        textos = st.session_state.get("liq_v3_textos") or {}
        nuevas = []
        for _, row in editado.iterrows():
            nombre = str(row.get("Archivo") or "")
            entidad = str(row.get("Entidad") or "Otra / No detectada")
            texto = textos.get(nombre, "")
            valores = extraer_con_plantilla(texto, entidad)
            valores["Archivo"] = nombre
            valores["Entidad"] = (
                entidad if entidad != "Otra / No detectada" else "Desconocida"
            )
            valores["Entidad_detectada"] = row.get("Entidad_detectada") or entidad
            nuevas.append(valores)
        st.session_state["liq_v3_filas"] = nuevas
        if gen:
            st.session_state["liq_v3_xlsx"] = exportar_liquidaciones_tarjeta_excel(pd.DataFrame(nuevas))
        st.rerun()

    xlsx = st.session_state.get("liq_v3_xlsx")
    if xlsx:
        cuit_limpio = re.sub(r"\D", "", str(st.session_state.get("cuit_activo") or ""))
        st.download_button(
            "Descargar Excel consolidado",
            data=xlsx,
            file_name=(
                f"Liquidaciones_Tarjetas_Consolidado_{cuit_limpio or 'cliente'}_"
                f"{date.today().strftime('%Y%m%d')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="liq_v3_dl",
            use_container_width=True,
        )

def _herramienta_analizador_inversiones() -> None:
    """PDF/Excel AlyC → clasificación por especie → FIFO/PEPS + saldo inicial DDJJ."""
    st.markdown("#### Analizador de inversiones (FIFO)")
    st.caption(
        "1) Convertí movimientos de AlyC/broker (PDF o Excel) · "
        "2) Clasificá por especie (bonos, FCI, dólar/MEP, acciones…) · "
        "3) Sembrá saldo inicial desde DDJJ o Excel · "
        "4) Aplicá **FIFO / PEPS** y descargá el Excel de trabajo."
    )

    c1, c2 = st.columns(2)
    with c1:
        archivos_mov = st.file_uploader(
            "Movimientos del año (PDF / Excel AlyC)",
            type=["pdf", "xlsx", "xls", "xlsm", "csv"],
            accept_multiple_files=True,
            key="inv_uploader_movimientos",
            help="Preferí Excel del broker si el PDF no es tabular.",
        )
    with c2:
        archivo_ddjj = st.file_uploader(
            "Saldo inicial — DDJJ / BIENES año anterior (PDF)",
            type=["pdf"],
            accept_multiple_files=False,
            key="inv_uploader_ddjj",
            help="Preferí el PDF de papeles de trabajo BIENES (F.711) si el comprobante DDJJ no trae tenencias.",
        )
        archivo_saldo_xlsx = st.file_uploader(
            "Saldo inicial — Excel manual (opcional)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="inv_uploader_saldo_xlsx",
            help="Si la DDJJ no se lee bien, usá la plantilla.",
        )

    st.download_button(
        "Descargar plantilla saldo inicial",
        data=plantilla_saldo_inicial_excel(),
        file_name="Plantilla_saldo_inicial_inversiones.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="inv_dl_plantilla",
    )

    if st.button("Procesar inversiones (FIFO)", type="primary", key="inv_btn_procesar"):
        if not archivos_mov and not archivo_ddjj and not archivo_saldo_xlsx:
            st.warning("Subí al menos movimientos del año o un saldo inicial.")
        else:
            with st.spinner("Leyendo archivos, clasificando especies y aplicando FIFO…"):
                df_mov = pd.DataFrame()
                errs: list[dict] = []
                if archivos_mov:
                    df_mov, errs = procesar_archivos_inversiones(archivos_mov)
                    df_mov = reclasificar_movimientos(df_mov)

                df_ini = pd.DataFrame()
                avisos_ini: list[str] = []
                if archivo_saldo_xlsx is not None:
                    df_ini = leer_saldo_inicial_excel(
                        archivo_saldo_xlsx.getvalue(),
                        archivo_saldo_xlsx.name,
                    )
                    if df_ini.empty:
                        avisos_ini.append("El Excel de saldo inicial no trajo filas válidas.")
                elif archivo_ddjj is not None:
                    df_ini, avisos_ini = extraer_saldo_inicial_ddjj_pdf(
                        archivo_ddjj.getvalue(),
                        archivo_ddjj.name,
                    )

                resultado = aplicar_fifo(df_mov, df_ini if not df_ini.empty else None)
                xlsx = exportar_inversiones_excel(
                    df_mov,
                    resultado,
                    df_inicial=df_ini if not df_ini.empty else None,
                    meta={"nota": st.session_state.get("nombre_activo") or ""},
                )
                st.session_state.inv_df_mov = df_mov
                st.session_state.inv_df_ini = df_ini
                st.session_state.inv_resultado = resultado
                st.session_state.inv_errores = errs
                st.session_state.inv_avisos_ini = avisos_ini
                st.session_state.inv_xlsx = xlsx
            st.rerun()

    errs = st.session_state.get("inv_errores") or []
    if errs:
        with st.expander(f"Advertencias de lectura ({len(errs)})", expanded=True):
            st.dataframe(pd.DataFrame(errs), use_container_width=True, hide_index=True)

    for msg in st.session_state.get("inv_avisos_ini") or []:
        st.info(msg)

    df_mov = st.session_state.get("inv_df_mov")
    df_ini = st.session_state.get("inv_df_ini")
    resultado = st.session_state.get("inv_resultado")
    xlsx = st.session_state.get("inv_xlsx")

    if df_mov is None and resultado is None:
        st.caption(
            "Grupos de clasificación: " + " · ".join(GRUPOS_ESPECIE)
        )
        return

    if df_mov is not None and not df_mov.empty:
        st.success(f"Movimientos normalizados: **{len(df_mov)}**")
        # Editor liviano de reclasificación
        cols_prev = [
            c for c in (
                "Fecha", "Especie", "Grupo", "Tipo_Operacion",
                "Cantidad", "Precio", "Monto_Total", "Moneda", "Descripcion",
            ) if c in df_mov.columns
        ]
        editado = st.data_editor(
            df_mov[cols_prev + (["Nueva_Clasificacion"] if "Nueva_Clasificacion" in df_mov.columns else [])],
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="inv_editor_mov",
            column_config={
                "Nueva_Clasificacion": st.column_config.SelectboxColumn(
                    "Nueva_Clasificacion",
                    options=[""] + list(GRUPOS_ESPECIE),
                    required=False,
                ),
            } if "Nueva_Clasificacion" in df_mov.columns else None,
        )
        if st.button("Reaplicar FIFO con clasificación editada", key="inv_btn_reaplicar"):
            work = df_mov.copy()
            if "Nueva_Clasificacion" in editado.columns:
                work["Nueva_Clasificacion"] = editado["Nueva_Clasificacion"].values
            if "Grupo" in editado.columns:
                # Si editaron Grupo directo, respetarlo vía Nueva_Clasificacion
                for i, g in enumerate(editado["Grupo"].tolist()):
                    if g in GRUPOS_ESPECIE:
                        work.at[work.index[i], "Nueva_Clasificacion"] = g
            work = reclasificar_movimientos(work)
            res2 = aplicar_fifo(work, df_ini if df_ini is not None and not df_ini.empty else None)
            st.session_state.inv_df_mov = work
            st.session_state.inv_resultado = res2
            st.session_state.inv_xlsx = exportar_inversiones_excel(
                work, res2, df_inicial=df_ini, meta={"nota": st.session_state.get("nombre_activo") or ""},
            )
            st.rerun()
    elif df_mov is not None:
        st.warning("No se extrajeron movimientos. Probá con Excel del AlyC.")

    if df_ini is not None and not df_ini.empty:
        st.markdown("##### Saldo inicial")
        st.dataframe(df_ini, use_container_width=True, hide_index=True)

    if resultado is not None:
        saldos = pd.DataFrame(resultado.saldos)
        if not saldos.empty:
            st.markdown("##### Saldos al cierre (FIFO)")
            st.dataframe(saldos, use_container_width=True, hide_index=True)
        if resultado.avisos:
            with st.expander(f"Avisos FIFO ({len(resultado.avisos)})", expanded=False):
                for a in resultado.avisos:
                    st.write(f"- {a}")

    if xlsx:
        stamp = datetime.now().strftime("%Y-%m-%d_%H %M")
        st.download_button(
            "Descargar Excel — Analizador inversiones FIFO",
            data=xlsx,
            file_name=f"{stamp}_Inversiones_FIFO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="inv_dl_xlsx",
            use_container_width=True,
        )


@st.fragment
def _herramienta_caja_usd() -> None:
    """Caja en USD: físico/bancario + TC posición vs TC egreso → dif. de cambio.

    UI aislada en fragment + form (sin columns ni download pegado a uploaders)
    para evitar el bug de Streamlit ``insertBefore`` / NotFoundError en el DOM.
    """
    st.markdown("#### Caja USD (cta. cte. + diferencia de cotización)")
    st.caption(
        "Separá **dólares físicos** y **bancarios**. "
        "Cada **ingreso** actualiza el **TC de posición** (promedio ponderado). "
        "Cada **egreso** se valúa al **TC del movimiento** → "
        "**Dif = monto × (TC_mov − TC_posición)**."
    )

    if "caja_usd_plantilla_tc_bytes" not in st.session_state:
        st.session_state.caja_usd_plantilla_tc_bytes = plantilla_cotizaciones_excel()

    with st.form("caja_usd_form_v3", clear_on_submit=False):
        pdfs_caja = st.file_uploader(
            "Extractos CA en USD (PDF) → bolsillo bancario",
            type=["pdf"],
            accept_multiple_files=True,
            key="caja_usd_uploader_pdf_v3",
        )
        xlsx_mov = st.file_uploader(
            "O Excel de movimientos (opc. col. Tipo_Dolar: fisico/bancario)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="caja_usd_uploader_xlsx_mov_v3",
        )
        xlsx_tc = st.file_uploader(
            "Cotizaciones diarias (Excel Fecha | TC)",
            type=["xlsx", "xls"],
            accept_multiple_files=False,
            key="caja_usd_uploader_tc_v3",
            help="BNA comprador u otra serie. Si falta un día, se usa el TC hábil anterior.",
        )
        metodo = st.selectbox(
            "Método de costo",
            options=["cta_cte", "fifo"],
            format_func=lambda x: (
                "Cta. cte. (físico/bancario + TC posición)"
                if x == "cta_cte"
                else "FIFO por lotes"
            ),
            key="caja_usd_metodo_v3",
        )
        saldo_banc = st.number_input(
            "Saldo inicial bancario USD",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.2f",
            key="caja_usd_saldo_banc_v3",
        )
        saldo_fis = st.number_input(
            "Saldo inicial físico USD",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.2f",
            key="caja_usd_saldo_fis_v3",
        )
        tc_ini = st.number_input(
            "TC saldo inicial",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.4f",
            key="caja_usd_tc_ini_v3",
            help="Si queda en 0, se busca en la tabla de cotizaciones.",
        )
        fecha_ini = st.text_input(
            "Fecha saldo inicial (dd/mm/aaaa)",
            value="01/01/2025",
            key="caja_usd_fecha_ini_v3",
        )
        armar = st.form_submit_button("Armar caja USD", type="primary")

    if armar:
        if not pdfs_caja and xlsx_mov is None:
            st.warning("Subí extractos PDF o un Excel de movimientos.")
        elif xlsx_tc is None:
            st.warning(
                "Subí el Excel de cotizaciones (Fecha | TC). "
                "Sin eso no hay diferencia de cotización."
            )
        else:
            from datetime import datetime as _dt

            df_mov = pd.DataFrame()
            errs: list[dict] = []
            if pdfs_caja:
                df_mov, errs = movimientos_desde_pdfs_caja_usd(pdfs_caja)
            if xlsx_mov is not None:
                df_x = movimientos_desde_excel_extracto(
                    xlsx_mov.getvalue(), xlsx_mov.name
                )
                if not df_x.empty:
                    df_mov = (
                        pd.concat([df_mov, df_x], ignore_index=True)
                        if not df_mov.empty
                        else df_x
                    )

            df_tc, avisos_tc = leer_cotizaciones_excel(
                xlsx_tc.getvalue(), xlsx_tc.name
            )

            f_ini = None
            try:
                f_ini = _dt.strptime(fecha_ini.strip(), "%d/%m/%Y").date()
            except Exception:
                f_ini = None

            tc_val = float(tc_ini) if tc_ini and tc_ini > 0 else None
            res = armar_caja_usd(
                df_mov,
                df_tc,
                saldo_inicial_bancario=float(saldo_banc or 0),
                saldo_inicial_fisico=float(saldo_fis or 0),
                tc_inicial=tc_val,
                fecha_inicial=f_ini,
                metodo=metodo,
            )
            xlsx_out = exportar_caja_usd_excel(
                res,
                df_cotiz=df_tc,
                meta={"nota": st.session_state.get("nombre_activo") or ""},
            )
            st.session_state.caja_usd_resultado = res
            st.session_state.caja_usd_xlsx = xlsx_out
            st.session_state.caja_usd_errs = errs
            st.session_state.caja_usd_avisos_tc = avisos_tc

    errs = st.session_state.get("caja_usd_errs") or []
    for e in errs:
        st.warning(f"{e.get('archivo')}: {e.get('motivo')}")
    for a in st.session_state.get("caja_usd_avisos_tc") or []:
        st.info(a)

    res = st.session_state.get("caja_usd_resultado")
    if res is not None:
        r = res.resumen or {}
        st.markdown(
            f"- **Saldo bancario:** {r.get('Saldo_Bancario', r.get('Saldo_USD', 0)):,.2f}  \n"
            f"- **Saldo físico:** {r.get('Saldo_Fisico', 0):,.2f}  \n"
            f"- **Dif. cotización:** ${r.get('Total_Dif_Cotizacion', 0):,.2f}  \n"
            f"- **Saldo ARS (costo):** ${r.get('Saldo_ARS_Costo', 0):,.2f}"
        )
        for a in (res.avisos or [])[:15]:
            st.caption(f"⚠ {a}")
        if res.caja:
            st.dataframe(
                pd.DataFrame(res.caja),
                use_container_width=True,
                hide_index=True,
            )
        xlsx_bytes = st.session_state.get("caja_usd_xlsx")
        if xlsx_bytes:
            st.download_button(
                "Descargar Excel — Caja USD",
                data=xlsx_bytes,
                file_name="Caja_USD_diferencia_cotizacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="caja_usd_dl_xlsx_v3",
                use_container_width=True,
            )

    # Plantilla al final: nunca junto a file_uploader (dispara insertBefore)
    with st.expander("Plantilla de cotizaciones (Fecha | TC)", expanded=False):
        st.download_button(
            "Descargar plantilla",
            data=st.session_state.caja_usd_plantilla_tc_bytes,
            file_name="Plantilla_cotizaciones_USD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="caja_usd_dl_plantilla_tc_v3",
        )


def _herramienta_cruce_facturas_arca() -> None:
    """Cruza facturas PDF/fotos vs Mis Comprobantes ARCA (portal IVA)."""
    st.markdown("#### Cruce Facturas vs ARCA")
    st.caption(
        "Subí las facturas del período (PDF o fotos) y el Excel/CSV de "
        "**Mis Comprobantes Recibidos** del portal IVA de ARCA. "
        "El cruce arma Matcheadas / A revisar / Faltantes / Diferencias de importe."
    )

    c1, c2 = st.columns(2)
    with c1:
        facturas = st.file_uploader(
            "Facturas (PDF / imágenes / ZIP)",
            type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp", "zip"],
            accept_multiple_files=True,
            key="uploader_cruce_facturas_arca",
            help="Varios PDF o fotos; también un ZIP con comprobantes.",
        )
        usar_ocr = st.checkbox(
            "OCR si el PDF viene escaneado o es foto",
            value=True,
            key="chk_cruce_arca_ocr",
            help="Más lento. Desactivá si todos los PDF tienen texto nativo.",
        )
    with c2:
        arca = st.file_uploader(
            "Listado ARCA — Mis Comprobantes (xlsx / xls / csv)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=False,
            key="uploader_cruce_arca_listado",
            help="Export del portal Mis Comprobantes Recibidos.",
        )
        cuit_manual = st.text_input(
            "CUIT contribuyente (opcional)",
            value=str(st.session_state.get("cuit_activo") or ""),
            key="txt_cruce_arca_cuit",
            help="Si está vacío se toma del nombre/archivo ARCA o del CUIT receptor más frecuente.",
        )

    puede = bool(facturas) and arca is not None
    if st.button(
        "Procesar cruce",
        type="primary",
        key="btn_cruce_facturas_arca",
        disabled=not puede,
    ):
        with st.spinner("Leyendo ARCA, extrayendo facturas y cruzando…"):
            try:
                resultado, errores, cuit_det, xlsx = procesar_cruce_facturas_arca(
                    facturas,
                    arca,
                    usar_ocr=usar_ocr,
                    cuit_contribuyente=cuit_manual,
                    nombre_arca=str(getattr(arca, "name", "arca.xlsx")),
                )
            except Exception as exc:
                st.error(f"No se pudo procesar el cruce: {exc}")
                return
        st.session_state["cruce_arca_resultado"] = resultado
        st.session_state["cruce_arca_errores"] = errores
        st.session_state["cruce_arca_cuit"] = cuit_det
        st.session_state["cruce_arca_xlsx"] = xlsx
        st.success("Cruce listo.")
        st.rerun()

    errores = st.session_state.get("cruce_arca_errores") or []
    if errores:
        with st.expander(f"Archivos con advertencias ({len(errores)})", expanded=False):
            st.dataframe(pd.DataFrame(errores), use_container_width=True, hide_index=True)

    resultado = st.session_state.get("cruce_arca_resultado")
    xlsx = st.session_state.get("cruce_arca_xlsx")
    if not resultado or not xlsx:
        return

    m = resultado.get("matcheadas")
    r = resultado.get("a_revisar")
    f = resultado.get("faltantes")
    d = resultado.get("diferencias")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Matcheadas", 0 if m is None else len(m))
    k2.metric("A revisar", 0 if r is None else len(r))
    k3.metric("Faltantes", 0 if f is None else len(f))
    k4.metric("Diferencias", 0 if d is None else len(d))

    cuit_det = st.session_state.get("cruce_arca_cuit") or ""
    if cuit_det:
        st.caption(f"CUIT detectado: {cuit_det}")

    tabs = st.tabs(["Matcheadas", "A revisar", "Faltantes", "Diferencias"])
    with tabs[0]:
        st.dataframe(m if m is not None else pd.DataFrame(), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.dataframe(r if r is not None else pd.DataFrame(), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(f if f is not None else pd.DataFrame(), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(d if d is not None else pd.DataFrame(), use_container_width=True, hide_index=True)

    cuit_limpio = re.sub(r"\D", "", str(cuit_det or st.session_state.get("cuit_activo") or ""))
    st.download_button(
        "Descargar Excel del cruce",
        data=xlsx,
        file_name=(
            f"Cruce_Facturas_vs_ARCA_{cuit_limpio or 'cliente'}_"
            f"{date.today().strftime('%Y%m%d')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="dl_cruce_facturas_arca",
        use_container_width=True,
    )


def _seccion_herramientas() -> None:
    """Solapa de utilidades de oficina (selectbox → herramienta activa)."""
    st.write(
        "Utilidades rápidas para el día a día del estudio, "
        "sin mezclar con los procesos contables anuales."
    )

    herramienta_activa = st.selectbox(
        "Seleccioná la herramienta que vas a usar:",
        options=[
            "Completar cuadro bancario existente",
            "Matcheo inteligente PDF + Tango",
            "Extractos de PDF ➔ Excel",
            "Analizador de inversiones (FIFO)",
            "Caja USD (dif. cotización)",
            "Convertidor de Liquidaciones de Tarjeta",
            "Match débitos - proveedores",
            "Cruce Facturas vs ARCA",
        ],
        index=0,
        key="herramientas_selectbox_v9",
    )
    st.divider()

    if herramienta_activa == "Completar cuadro bancario existente":
        _herramienta_completar_cuadro_bancario()
    elif herramienta_activa == "Matcheo inteligente PDF + Tango":
        _herramienta_matcheo_inteligente_pdf()
    elif herramienta_activa == "Extractos de PDF ➔ Excel":
        _herramienta_pdf_extractos_a_excel()
    elif herramienta_activa == "Analizador de inversiones (FIFO)":
        _herramienta_analizador_inversiones()
    elif herramienta_activa == "Caja USD (dif. cotización)":
        _herramienta_caja_usd()
    elif herramienta_activa == "Convertidor de Liquidaciones de Tarjeta":
        _herramienta_liquidaciones_tarjeta()
    elif herramienta_activa == "Match débitos - proveedores":
        _herramienta_match_debitos_proveedores()
    elif herramienta_activa == "Cruce Facturas vs ARCA":
        _herramienta_cruce_facturas_arca()

def _seccion_recategorizacion_monotributo() -> None:
    """Análisis de períodos devengados en facturas electrónicas AFIP (PDF / ZIP)."""
    st.caption(
        "Cargá facturas electrónicas AFIP (PDF sueltos o ZIP) para consolidar "
        "períodos facturados y preparar el papel de trabajo de recategorización."
    )

    clientes = db.listar_clientes()
    if not clientes:
        st.warning("Debe registrar al menos un cliente antes de analizar facturas.")
        return

    _selector_sociedad_devengamientos(clientes)
    if (
        st.session_state.cuit_activo is None
        or st.session_state.nombre_activo is None
        or st.session_state.get("plan_cuentas_cliente_id") != st.session_state.get(_SOCiedad_KEY)
    ):
        actualizar_sociedad_activa()

    nombre = st.session_state.get("nombre_activo") or "—"
    cuit = st.session_state.get("cuit_activo") or "—"
    cliente = db.obtener_cliente(st.session_state.get(_SOCiedad_KEY)) if st.session_state.get(_SOCiedad_KEY) else None
    tipo_persona = cliente.get("tipo_persona", "—") if cliente else "—"

    col_n, col_c, col_t = st.columns(3)
    col_n.metric("Sociedad activa", nombre)
    col_c.metric("CUIT", cuit)
    col_t.metric("Tipo", tipo_persona)

    archivos = st.file_uploader(
        "Facturas electrónicas AFIP (PDF o ZIP)",
        type=["pdf", "zip"],
        accept_multiple_files=True,
        key="mono_uploader_facturas",
        help="Podés subir varios PDFs o un ZIP con múltiples comprobantes.",
    )

    if st.button("🚀 Analizar Períodos Devengados", type="primary", key="mono_btn_analizar"):
        if not archivos:
            st.warning("Subí al menos un archivo PDF o ZIP para analizar.")
        else:
            with st.spinner("Procesando facturas y extrayendo períodos devengados..."):
                df, errores = procesar_facturas_monotributo(archivos)
                st.session_state.mono_facturas_df = df
                st.session_state.mono_errores_extraccion = errores
            if df.empty:
                st.warning("No se extrajeron comprobantes válidos de los archivos cargados.")
            else:
                st.success(f"Se analizaron {len(df)} comprobante(s) correctamente.")
            st.rerun()

    df_mono = st.session_state.get("mono_facturas_df")
    errores_mono = st.session_state.get("mono_errores_extraccion") or []

    if errores_mono:
        with st.expander(f"Archivos con advertencias ({len(errores_mono)})", expanded=False):
            st.dataframe(pd.DataFrame(errores_mono), use_container_width=True, hide_index=True)

    if df_mono is not None and not df_mono.empty:
        total_importe = round(float(df_mono["Importe Total"].sum()), 2)
        total_fc = round(float(df_mono.loc[df_mono["Importe Total"] > 0, "Importe Total"].sum()), 2)
        total_nc = round(float(df_mono.loc[df_mono["Importe Total"] < 0, "Importe Total"].sum()), 2)
        n_nc = int((df_mono["Importe Total"] < 0).sum())
        periodos = df_mono["Período Desde"].astype(str).tolist()
        st.markdown(
            f"**Resumen:** {len(df_mono)} comprobante(s) "
            f"({n_nc} nota(s) de crédito) · "
            f"Facturado **${total_fc:,.2f}** · "
            f"NC **${total_nc:,.2f}** · "
            f"Neto **${total_importe:,.2f}** · "
            f"Período **{periodos[0]}** a **{periodos[-1]}**"
        )
        st.caption("Las notas de crédito figuran con importe negativo y ya están descontadas del neto.")
        st.dataframe(df_mono, use_container_width=True, hide_index=True)

        cuit_limpio = re.sub(r"\D", "", str(cuit)) or "00000000000"
        nombre_xlsx = f"Recategorizacion_Monotributo_{cuit_limpio}_{date.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            "Descargar papel de trabajo (Excel)",
            data=exportar_monotributo_excel(df_mono),
            file_name=nombre_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="mono_dl_excel",
            use_container_width=True,
        )


def _seccion_auditoria_prestamos() -> None:
    # Usa generar_auditoria.py: parsers por banco + OCR 200DPI para Provincia + reintentos automáticos
    st.caption(
        "Auditoría de cuotas desde PDFs bancarios (proyecciones de préstamos)."
    )

    import os as _os

    _base_dir = _os.path.dirname(__file__)

    # Botón para el Excel completo con datos reales (si ya fue generado)
    _completo_path = _os.path.join(_base_dir, "Auditoria_Prestamos_Completa.xlsx")
    if _os.path.exists(_completo_path):
        with open(_completo_path, "rb") as _f:
            st.download_button(
                "⬇️ Descargar Auditoría Completa (datos reales)",
                _f.read(),
                file_name="Auditoria_Prestamos_Completa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_auditoria_completa",
                type="primary",
                help="Excel con todos los préstamos procesados desde los PDFs reales",
            )

    # Botón para el Excel demo (datos ficticios de ejemplo)
    _demo_path = _os.path.join(_base_dir, "demo_prestamos_auditoria.xlsx")
    if _os.path.exists(_demo_path):
        with open(_demo_path, "rb") as _f:
            st.download_button(
                "⬇️ Descargar Excel Demo (Préstamos)",
                _f.read(),
                file_name="demo_prestamos_auditoria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_demo_prestamos",
                help="Ejemplo de planilla de auditoría con datos ficticios (Banco Galicia + Banco Santander)",
            )

    clientes = db.listar_clientes()
    if not clientes:
        st.warning("Debe registrar al menos un cliente antes de ejecutar la auditoría.")
        return

    cliente = _cliente_sociedad_activa(clientes)
    if not cliente:
        return

    st.caption(f"Sociedad activa: **{cliente['nombre']}** (elegila en Devengamientos si necesitás cambiarla).")

    st.session_state.cliente_activo = cliente

    col_info1, col_info2 = st.columns(2)
    col_info1.metric("CUIT", cliente["cuit"])
    col_info2.metric("Tipo", cliente["tipo_persona"])

    st.markdown("#### Carga de PDFs de préstamos")
    st.caption(
        "Suba las proyecciones / tablas de amortización de cada préstamo "
        "(Santander, Provincia, Galicia, Francés, Nación, Mercado Pago). "
        f"Si no sube archivos, se procesará la carpeta `extractos bancarios/Prestamos Financieros`."
    )
    pdfs_prestamo = st.file_uploader(
        "PDFs de Proyección de Préstamos",
        type=["pdf"],
        accept_multiple_files=True,
        help="Contratos y tablas de amortización con la grilla de cuotas por banco.",
        key="upload_pdfs_prestamo_ga",
    )

    saldos_iniciales = _formulario_saldos_iniciales_bancos()

    if st.button("Generar Excel de Auditoría", type="primary", key="btn_generar_auditoria_ga"):
        try:
            with st.spinner("Procesando PDFs con parsers especializados por banco..."):
                if pdfs_prestamo:
                    tmp_dir = Path(tempfile.mkdtemp(prefix="ec_prestamos_"))
                    try:
                        for archivo in pdfs_prestamo:
                            dest = tmp_dir / archivo.name
                            dest.write_bytes(archivo.getbuffer())
                        bancos_data = _ga_procesar_todos(tmp_dir)
                    finally:
                        shutil.rmtree(str(tmp_dir), ignore_errors=True)
                else:
                    if not _GA_CARPETA_PRESTAMOS.exists():
                        st.error(
                            f"No subió PDFs y la carpeta predeterminada no existe: "
                            f"`{_GA_CARPETA_PRESTAMOS}`"
                        )
                        return
                    bancos_data = _ga_procesar_todos(_GA_CARPETA_PRESTAMOS)

            if not bancos_data:
                st.warning(
                    "No se encontraron cuotas en los PDFs. "
                    "Verifique que los archivos sean tablas de amortización de préstamos."
                )
                return

            saldos_dict = {
                s["banco"]: float(s.get("saldo_inicial", 0.0))
                for s in saldos_iniciales
                if s.get("banco")
            }
            ruta_salida = BASE_DIR / "Auditoria_Prestamos_Completa.xlsx"

            with st.spinner("Generando Excel con formato de auditoría..."):
                _ga_generar_excel(bancos_data, saldos_dict, ruta_salida)

            total_cuotas = sum(
                sum(len(p["cuotas"]) for p in prestamos)
                for prestamos in bancos_data.values()
            )
            bancos_ok = list(bancos_data.keys())
            st.success(
                f"✅ Excel generado: {total_cuotas} cuotas | "
                f"{len(bancos_ok)} banco(s): {', '.join(bancos_ok)}"
            )

            st.markdown("#### Resumen por banco")
            filas_res = []
            for banco, prestamos in bancos_data.items():
                filas_res.append({
                    "Banco": banco,
                    "Préstamos": len(prestamos),
                    "Cuotas": sum(len(p["cuotas"]) for p in prestamos),
                    "Capital Total ($)": sum(p.get("capital_original", 0) for p in prestamos),
                })
            st.dataframe(pd.DataFrame(filas_res), use_container_width=True, hide_index=True)

            with open(str(ruta_salida), "rb") as _f:
                st.download_button(
                    label="⬇️ Descargar Auditoria_Prestamos_Completa.xlsx",
                    data=_f.read(),
                    file_name="Auditoria_Prestamos_Completa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="dl_auditoria_generada",
                )

        except Exception as exc:
            st.error(f"Error en la generación: {exc}")

def _cerrar_sesion_oficina() -> None:
    """Sale del usuario de oficina y limpia estado de trabajo en memoria."""
    claves_limpiar = [
        "usuario_oficina", "usuario_oficina_nombre", "usuario_oficina_admin",
        "_persistencia_hidratada", "biblioteca_asientos", "biblioteca_bancos",
        "periodos_procesados", "periodos_bancos_procesados",
        "extracto_santander_df", "extracto_santander_xlsx", "match_prov_resultado",
        "match_prov_xlsx", "vista_admin", "cuadro_bancario_resultado",
        "cuadro_bancario_resultado_signature", "cuadro_bancario_excel_destino_v1",
        "cuadro_bancario_pdfs_v1", "cuadro_bancario_buzon", "cuadro_bancario_buzon_ruta",
        "cuadro_bancario_ruta_buzon_v1", "cuadro_bancario_excel_sel_v1",
        "cuadro_bancario_pdfs_sel_v1",
    ]
    for k in claves_limpiar:
        if k in st.session_state:
            del st.session_state[k]
    st.session_state.usuario_oficina = None
    st.session_state.usuario_oficina_nombre = None
    st.session_state.usuario_oficina_admin = False
    st.session_state._persistencia_hidratada = False


def _pantalla_login_oficina() -> None:
    """Pantalla de ingreso: cada persona de la oficina elige su usuario."""
    st.title("Estudio Contable")
    st.subheader("Ingreso a la oficina")
    st.caption(
        "Cada persona entra con su usuario. Las sesiones son independientes: "
        "varios pueden trabajar a la vez sin pisarse."
    )
    if _es_entorno_cloud():
        # Re-sincronizar usuarios desde Secrets en cada visita al login (Cloud)
        try:
            db._aplicar_usuarios_desde_secrets()
        except Exception:
            pass

    usuarios = db.listar_usuarios_oficina(solo_activos=True)
    if not usuarios:
        st.error(
            "No hay usuarios cargados. En Cloud, pegá el bloque `oficina_usuarios` en Secrets "
            "y reiniciá la app."
        )
        return

    opciones = {u["usuario"]: f"{u['nombre']} ({u['usuario']})" for u in usuarios}
    elegido = st.selectbox(
        "Usuario",
        options=list(opciones.keys()),
        format_func=lambda x: opciones[x],
        key="login_usuario_select",
    )
    pin = st.text_input(
        "PIN",
        type="password",
        key="login_pin_input",
        help="Obligatorio en Cloud. En local, si el usuario no tiene PIN, dejalo vacío.",
    )
    if st.button("Entrar", type="primary", key="btn_login_oficina"):
        ok = db.verificar_login_oficina(elegido, pin)
        if not ok:
            st.error("Usuario o PIN incorrecto.")
            return
        st.session_state.usuario_oficina = ok["usuario"]
        st.session_state.usuario_oficina_nombre = ok["nombre"]
        st.session_state.usuario_oficina_admin = bool(ok.get("es_admin"))
        st.session_state._persistencia_hidratada = False
        st.rerun()

    with st.expander(
        "¿Olvidaste tu PIN? Recuperar acceso",
        expanded=False,
        key="expander_recup_pin_login",
    ):
        st.caption(
            "Un administrador puede autorizar el reset del PIN acá. "
            "Después entrá con el PIN nuevo."
        )
        admins = [u for u in usuarios if u.get("es_admin")]
        if not admins:
            st.warning("No hay administradores activos para autorizar la recuperación.")
            return

        objetivo = st.selectbox(
            "Usuario a recuperar",
            options=list(opciones.keys()),
            format_func=lambda x: opciones[x],
            key="recup_usuario_objetivo",
        )
        mapa_admin = {u["usuario"]: f"{u['nombre']} ({u['usuario']})" for u in admins}
        admin_elegido = st.selectbox(
            "Administrador que autoriza",
            options=list(mapa_admin.keys()),
            format_func=lambda x: mapa_admin[x],
            key="recup_usuario_admin",
        )
        pin_admin = st.text_input(
            "PIN del administrador",
            type="password",
            key="recup_pin_admin",
        )
        c_pin1, c_pin2 = st.columns(2)
        with c_pin1:
            pin_nuevo = st.text_input(
                "Nuevo PIN",
                type="password",
                key="recup_pin_nuevo",
                help="Mínimo 4 caracteres.",
            )
        with c_pin2:
            pin_nuevo_ok = st.text_input(
                "Repetir nuevo PIN",
                type="password",
                key="recup_pin_nuevo_ok",
            )
        if st.button("Restablecer PIN", type="secondary", key="btn_recuperar_pin"):
            if str(pin_nuevo or "").strip() != str(pin_nuevo_ok or "").strip():
                st.error("Los PIN nuevos no coinciden.")
            else:
                try:
                    res = db.resetear_pin_usuario_oficina(
                        objetivo,
                        pin_nuevo,
                        usuario_admin=admin_elegido,
                        pin_admin=pin_admin,
                    )
                    st.success(
                        f"PIN de **{res['nombre']}** ({res['usuario']}) restablecido. "
                        "Ya podés ingresar arriba con el PIN nuevo."
                    )
                except Exception as exc:
                    st.error(str(exc))


def _seccion_usuarios_oficina() -> None:
    """Alta/edición de usuarios de la oficina (admin)."""
    st.subheader("Usuarios de la oficina")
    st.caption("Cada persona tiene su propio espacio de trabajo (borradores y biblioteca).")

    usuarios = db.listar_usuarios_oficina(solo_activos=False)
    if usuarios:
        st.dataframe(
            pd.DataFrame([
                {
                    "Usuario": u["usuario"],
                    "Nombre": u["nombre"],
                    "Admin": "Sí" if u.get("es_admin") else "No",
                    "Activo": "Sí" if u.get("activo") else "No",
                }
                for u in usuarios
            ]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("#### Nuevo usuario")
    c1, c2, c3 = st.columns(3)
    with c1:
        nuevo_user = st.text_input("Usuario (login)", key="nuevo_user_login")
    with c2:
        nuevo_nombre = st.text_input("Nombre para mostrar", key="nuevo_user_nombre")
    with c3:
        nuevo_pin = st.text_input("PIN (opcional)", type="password", key="nuevo_user_pin")
    es_admin = st.checkbox("Es administrador", key="nuevo_user_admin")
    if st.button("Crear usuario", key="btn_crear_usuario_oficina"):
        try:
            db.crear_usuario_oficina(
                nuevo_user, nuevo_nombre, pin=nuevo_pin or "", es_admin=es_admin
            )
            st.success(f"Usuario `{nuevo_user}` creado.")
            st.rerun()
        except Exception as exc:
            st.error(str(exc))

    st.markdown("#### Editar usuario")
    if not usuarios:
        return
    mapa = {u["id"]: f"{u['nombre']} ({u['usuario']})" for u in usuarios}
    uid = st.selectbox(
        "Usuario a editar",
        options=list(mapa.keys()),
        format_func=lambda x: mapa[x],
        key="edit_user_id",
    )
    actual = next(u for u in usuarios if u["id"] == uid)
    nom_edit = st.text_input("Nombre", value=actual["nombre"], key="edit_user_nombre")
    pin_edit = st.text_input(
        "Nuevo PIN (dejar vacío para no cambiar; escribir espacio+borrar para quitar PIN)",
        type="password",
        key="edit_user_pin",
    )
    admin_edit = st.checkbox("Admin", value=bool(actual.get("es_admin")), key="edit_user_admin")
    activo_edit = st.checkbox("Activo", value=bool(actual.get("activo")), key="edit_user_activo")
    if st.button("Guardar cambios", key="btn_guardar_usuario_oficina"):
        kwargs = {
            "nombre": nom_edit,
            "es_admin": admin_edit,
            "activo": activo_edit,
        }
        # Solo actualizar PIN si el campo tiene contenido explícito
        if pin_edit is not None and str(pin_edit) != "":
            kwargs["pin"] = pin_edit
        db.actualizar_usuario_oficina(int(uid), **kwargs)
        st.success("Usuario actualizado.")
        st.rerun()


def main() -> None:
    _init_session_state()

    # Gate de login: sin usuario de oficina no se entra a la app
    if not st.session_state.get("usuario_oficina"):
        _pantalla_login_oficina()
        return

    _hidratar_persistencia_desde_disco()

    # #region agent log
    _dbg_log("G", "main", "run", {
        "usuario": st.session_state.get("usuario_oficina"),
        "sociedad_activa": st.session_state.get(_SOCiedad_KEY),
        "cliente_id": st.session_state.get("cliente_id_seleccionado"),
        "plan_df_set": st.session_state.get("plan_cuentas_df") is not None,
        "plan_cliente_id": st.session_state.get("plan_cuentas_cliente_id"),
    })
    # #endregion

    clientes = db.listar_clientes()
    cliente_sidebar = None
    if st.session_state.get("cuit_activo") and st.session_state.get("nombre_activo"):
        cliente_sidebar = {
            "nombre": st.session_state.nombre_activo,
            "cuit": st.session_state.cuit_activo,
            "tipo_persona": "Persona Jurídica",
        }
    elif st.session_state.get(_SOCiedad_KEY):
        cliente_sidebar = db.obtener_cliente(st.session_state[_SOCiedad_KEY])
    elif st.session_state.cliente_id_seleccionado:
        cliente_sidebar = db.obtener_cliente(st.session_state.cliente_id_seleccionado)
    elif clientes:
        cliente_sidebar = clientes[0]

    _render_sidebar_sociedad_y_biblioteca(cliente_sidebar)

    # --- Vista administración (ocupa el área principal) ---
    vista_admin = st.session_state.get("vista_admin")
    if vista_admin == "clientes":
        _render_titulo_estudio()
        _render_barra_superior_cuenta()
        _seccion_clientes()
        return
    if vista_admin == "usuarios_oficina":
        _render_titulo_estudio()
        _render_barra_superior_cuenta()
        if not st.session_state.get("usuario_oficina_admin"):
            st.warning("Solo administradores pueden gestionar usuarios.")
            return
        _seccion_usuarios_oficina()
        return
    if vista_admin == "acerca":
        _render_titulo_estudio()
        _render_barra_superior_cuenta()
        st.subheader("Acerca del sistema")
        st.markdown(
            f"""
            - **Devengamientos de Fin de Mes**: solo Personas Jurídicas → Excel asientos Tango.
            - **Conciliación Bancaria**: extractos PDF + lista Tango → planilla Excel clonada.
            - **Préstamos Financieros**: auditoría de cuotas desde PDFs bancarios.
            - **Recategorización Monotributo**: facturas AFIP PDF/ZIP → períodos devengados + Excel.
            - **Herramientas**: **matcheo inteligente PDF + Tango**; completar cuadro bancario; PDF extractos → Excel; **analizador de inversiones (FIFO)**; **caja USD (dif. cotización)**; match débitos ↔ proveedores; liquidaciones de tarjeta; **cruce facturas vs ARCA**.
            - **Usuarios de oficina**: cada persona entra con su usuario; sesiones independientes.
            - **Cloud**: link público + muro de login (PIN). Planes/balances subidos se cifran con `DATA_ENCRYPTION_KEY`.
            - **Multi-PDF anual**: hasta {MAX_PDFS_ANUALES} extractos consolidados cronológicamente.
            - **Detección automática de banco**: Santander, Galicia, Francés, Credicoop, Provincia, Macro.
            - **Cruce con Compras Tango**: `{COMPRAS_TANGO_PATH.name}`.
            - **CUITs auxiliares**: `{db.CUITS_AUXILIARES_PATH.name}`.
            """
        )
        return

    # Título → Nav → Toolbar (orden crítico: el CSS fixed del toolbar no debe
    # anteponerse al nav en el árbol de widgets / hit-testing).
    _render_titulo_estudio()
    ventana_activa = _render_nav_ventanas_principales()
    _detectar_cambio_ventana_y_flush()
    _render_barra_superior_cuenta()
    st.divider()

    # Título dinámico = exactamente la opción seleccionada en la barra
    st.markdown(f"### **{ventana_activa}**")

    try:
        if ventana_activa == "Devengamiento de Impuestos":
            _seccion_devengamientos()
        elif ventana_activa == "Conciliación Bancaria":
            with st.container():
                _seccion_conciliacion_bancaria_balance()
        elif ventana_activa == "Préstamos Financieros":
            _seccion_auditoria_prestamos()
        elif ventana_activa == "Recategorización Monotributo":
            _seccion_recategorizacion_monotributo()
        elif _es_ventana_herramientas(ventana_activa):
            _seccion_herramientas()
    except Exception as _exc_modulo:
        from cursor_error_report import render_boton_enviar_error_cursor

        render_boton_enviar_error_cursor(
            _exc_modulo,
            contexto=f"Módulo: {ventana_activa}",
            key=f"enviar_error_cursor_{ventana_activa}",
        )


if __name__ == "__main__":
    main()
