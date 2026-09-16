"""Inversiones — Análisis fiscal de inversiones (Ganancias / Bienes Personales).

Etapa 1 de la integración: base de datos compartida + Carga de datos.

A diferencia de la versión original (HTML/JS autocontenido, localStorage por
navegador), esta pestaña usa los MISMOS clientes que el resto de Estudio
Contable (``database.listar_clientes()``) y guarda las operaciones en la
base SQLite del estudio (ver ``inversiones_db.py``): los datos se comparten
entre usuarios y dispositivos.

Etapa 2 agrega Depuración (pool de dólares con TC de origen y asignación
PEPS a las compras en USD). Etapa 3 agrega Patrimonio (posición PEPS al
cierre, valuada "al origen" — solo Persona Humana por ahora, sin
reexpresión por inflación) con control de cierre opcional contra el
extracto del broker. Etapa 3.5 agrega Movimientos: el detalle de todas las
operaciones agrupadas por instrumento, en orden cronológico, con el
resultado realizado de cada venta/rescate calculado contra el costo PEPS
(misma lógica que Patrimonio). Bienes Personales todavía queda pendiente.
La pestaña de la herramienta original (legado, localStorage) se retiró:
ya no aporta nada que no esté cubierto acá.
"""

from __future__ import annotations

import datetime as _dt
import io
import re
import unicodedata
from typing import Optional

import pandas as pd
import streamlit as st

import database as db
import inversiones_db as idb

_MOV_SALDO_INICIAL_RAW = {"saldo_inicial", "saldo inicial", "saldoinicial", "inicial"}


# ────────────────────────────────────────────────────────────────────────
# Helpers de parsing (reimplementación en Python de la lógica de
# importarPlantilla() / parseDate() / pv() de InversionARG_Final.html)
# ────────────────────────────────────────────────────────────────────────

def _normalizar(texto: str) -> str:
    t = unicodedata.normalize("NFD", str(texto or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return t.lower().strip()


def _vacio(v) -> bool:
    """True si la celda está vacía (None, NaN de pandas, o string en blanco)."""
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip().lower()
    return s in ("", "nan", "nat", "none")


def _parse_fecha(v) -> str:
    if _vacio(v):
        return ""
    if isinstance(v, (pd.Timestamp, _dt.date, _dt.datetime)):
        try:
            return pd.Timestamp(v).date().isoformat()
        except Exception:
            return ""
    s = str(v).strip()
    if not s or s.lower() == "nat":
        return ""
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        return s[:10]
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        d, mth, y = m.groups()
        return f"{y}-{mth.zfill(2)}-{d.zfill(2)}"
    return ""


def _pv(v) -> float:
    """Número robusto: acepta formato EU "1.234,56" o US "1,234.56"."""
    if _vacio(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if not s or s == "-":
        return 0.0
    ld, lc = s.rfind("."), s.rfind(",")
    try:
        if ld > 0 and lc > 0:
            return float(s.replace(",", "")) if ld > lc else float(s.replace(".", "").replace(",", "."))
        if lc > 0 and ld < 0:
            return float(s.replace(",", "."))
        return float(s)
    except ValueError:
        return 0.0


def _detectar_tipo(instrumento: str) -> str:
    tk = instrumento.upper().strip()
    if re.match(r"^(GD|AL|AE|TX|TV)\d", tk):
        return "bono"
    if re.match(r"^(BPOB|BPOC|BPOD|BPY)\d*", tk):
        return "on"
    if re.match(r"^(GGAL|YPF|PAMP|COME|BMA|BBAR|CRES|EDN|TECO|SUPV|TXAR|IRSA|VALO)", tk):
        return "accion"
    if "CEDEAR" in instrumento.upper():
        return "cedear"
    if re.match(r"^(AAPL|MSFT|GOOGL|AMZN|META|TSLA|NVDA|SPY|QQQ|ARKK)", tk):
        return "cedear"
    if re.match(r"^(BTC|ETH|USDT|SOL|BNB)", tk):
        return "crypto"
    if re.match(r"^S\d{2}[A-Z]\d", tk) or re.search(r"lecap|letes", instrumento, re.I):
        return "letra"
    if re.search(r"bal multiac|adcap|schrod|premium|superfondo|fci", instrumento, re.I):
        return "fci"
    if re.search(r"caucion", instrumento, re.I):
        return "caucion"
    return "on"


# ────────────────────────────────────────────────────────────────────────
# Helpers de formato (PESOS / USD / CANTIDAD) para las tablas de Inversiones
# ────────────────────────────────────────────────────────────────────────

def _fmt_monto(v, prefijo: str = "", decimales: int = 2) -> str:
    """Formatea un número con separador de miles '.' y decimal ',' (convención
    argentina), con un prefijo opcional ('$' para pesos, 'USD' para dólares).
    Sin prefijo queda como formato de CANTIDAD simple. El signo (para
    resultados negativos) va siempre adelante de todo, ej. "-$ 1.234,50"."""
    try:
        v = float(v)
    except (TypeError, ValueError):
        v = 0.0
    signo = "-" if v < 0 else ""
    v = abs(v)
    txt = f"{v:,.{decimales}f}" if decimales else f"{v:,.0f}"
    entero, _, dec = txt.partition(".")
    entero = entero.replace(",", ".")
    numero = f"{entero},{dec}" if decimales else entero
    cuerpo = f"{prefijo} {numero}".strip() if prefijo else numero
    return f"{signo}{cuerpo}"


def _fmt_pesos(v, decimales: int = 2) -> str:
    return _fmt_monto(v, "$", decimales)


def _fmt_usd(v, decimales: int = 2) -> str:
    return _fmt_monto(v, "USD", decimales)


def _fmt_cantidad(v, decimales: int = 2) -> str:
    return _fmt_monto(v, "", decimales)


def _fmt_moneda(v, moneda: str, decimales: int = 2) -> str:
    """Formatea según la moneda de la operación (para columnas como Comisión
    que pueden estar en ARS o USD según cada fila)."""
    return _fmt_usd(v, decimales) if moneda == "USD" else _fmt_pesos(v, decimales)


# ────────────────────────────────────────────────────────────────────────
# Sección principal
# ────────────────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def _asegurar_tablas_inversiones() -> bool:
    """Red de seguridad: crea (si hace falta) las tablas de Inversiones y
    siembra el TC BNA histórico, por si ``database.inicializar_bd()`` no
    llegó a esta parte (p. ej. si algo previo en esa función global tira una
    excepción y corta la secuencia antes de llegar a Inversiones — así se
    vio el error "no such table: inversiones_tc_bna" en un despliegue nuevo).
    Todas las sentencias que corre son CREATE TABLE IF NOT EXISTS / ALTER con
    manejo de "ya existe", así que es seguro llamarlo siempre — y
    ``@st.cache_resource`` además asegura que solo se ejecute una vez por
    proceso corriendo, no en cada rerun."""
    with db.obtener_conexion() as conn:
        idb.inicializar_tablas_inversiones(conn)
        conn.commit()
    idb.sembrar_tc_bna_default()
    return True


def _selector_cliente_periodo() -> tuple[Optional[dict], str]:
    """Selector único de Cliente y Período, compartido por todas las
    solapas de Inversiones: elegirlos acá alcanza para que el resto de las
    solapas (Carga de datos, Tenencia USD, Depuración, Movimientos,
    Patrimonio) ya los tengan seleccionados, sin tener que volver a elegir
    en cada una."""
    clientes = db.listar_clientes()
    if not clientes:
        st.warning("Todavía no hay clientes cargados en el estudio.")
        return None, ""

    opciones = {f"{c['nombre']} ({c['cuit']})": c for c in clientes}
    anio_actual = _dt.date.today().year

    c1, c2 = st.columns([2, 1])
    nombre_sel = c1.selectbox("Cliente", list(opciones.keys()), key="inv_cliente_global")
    cliente = opciones[nombre_sel]
    periodo = c2.selectbox(
        "Período (año fiscal)",
        [str(a) for a in range(anio_actual + 1, anio_actual - 6, -1)],
        key="inv_periodo_global",
    )
    st.caption(
        "El cliente y el período elegidos acá se usan en todas las solapas "
        "de Inversiones (Carga de datos, Tenencia USD, Depuración, "
        "Movimientos, Patrimonio)."
    )
    return cliente, periodo


def seccion_inversiones_arg() -> None:
    """Renderiza la pestaña Inversiones."""
    _asegurar_tablas_inversiones()
    st.subheader("📈 Inversiones")

    cliente, periodo = _selector_cliente_periodo()
    if cliente is None:
        return

    tab_carga, tab_usd, tab_dep, tab_mov, tab_pat = st.tabs(
        ["Carga de datos", "Tenencia USD", "Depuración", "Movimientos", "Patrimonio"]
    )

    with tab_carga:
        _tab_carga_datos(cliente, periodo)

    with tab_usd:
        _tab_tenencia_usd(cliente)

    with tab_dep:
        _tab_depuracion(cliente, periodo)

    with tab_mov:
        _tab_movimientos(cliente, periodo)

    with tab_pat:
        _tab_patrimonio(cliente, periodo)


def _tab_carga_datos(cliente: dict, periodo: str) -> None:
    st.caption(
        "Los datos se guardan en la base del estudio: se comparten entre "
        "todos los usuarios y no dependen de este navegador/equipo."
    )

    _form_nueva_operacion(cliente, periodo)
    st.divider()
    _form_saldo_inicial(cliente, periodo)
    st.divider()
    _importar_plantilla(cliente)
    st.divider()
    _tabla_operaciones(cliente, periodo)


_OPCION_MANUAL = "✍️ Escribir manualmente…"


def _fecha_default_periodo(periodo: str) -> _dt.date:
    """Fecha por defecto para el campo Fecha de Nueva operación: hoy si cae
    dentro del período fiscal seleccionado, o el 31/12 de ese período si no
    (evita el caso típico de cargar una operación de un período pasado y que
    quede con la fecha de hoy — no aparecería en la tabla de ese período)."""
    hoy = _dt.date.today()
    try:
        anio = int(periodo)
    except (TypeError, ValueError):
        return hoy
    if hoy.year == anio:
        return hoy
    return _dt.date(anio, 12, 31)


@st.cache_data(show_spinner=False)
def _catalogo_opciones() -> tuple[list[str], dict[str, dict]]:
    """Lista de opciones para el selector de instrumento + mapa label -> entrada
    del catálogo (ticker/nombre/tipo), armado a partir de las planillas de
    valuaciones 2025 (ver inversiones_db.listar_catalogo_instrumentos)."""
    catalogo = idb.listar_catalogo_instrumentos()
    opciones = [_OPCION_MANUAL]
    mapa: dict[str, dict] = {}
    for c in catalogo:
        tipo_label = idb.TIPO_LABEL.get(c["tipo"], c["tipo"])
        label = f"{c['ticker']} — {c['nombre']} ({tipo_label})" if c.get("ticker") else f"{c['nombre']} ({tipo_label})"
        if label in mapa:
            continue
        opciones.append(label)
        mapa[label] = c
    return opciones, mapa


def _form_nueva_operacion(cliente: dict, periodo: str) -> None:
    st.markdown("**Nueva operación**")
    st.caption(
        "Buscá el instrumento en la lista (tickers/nombres de las planillas de "
        "valuaciones 2025) para que el tipo y el tratamiento impositivo se "
        "sugieran solos — siempre editables. Si no está en la lista, elegí "
        f"«{_OPCION_MANUAL}»."
    )

    version = st.session_state.get("inv_op_version", 0)

    opciones_instr, mapa_instr = _catalogo_opciones()

    c1, c2 = st.columns([1, 2])
    fecha = c1.date_input(
        "Fecha", value=_fecha_default_periodo(periodo), key=f"inv_op_fecha_{version}",
    )
    instr_sel = c2.selectbox(
        "Instrumento (buscar por ticker o nombre)", opciones_instr,
        key=f"inv_op_instrsel_{version}",
    )
    entry = mapa_instr.get(instr_sel)
    instr_idx = opciones_instr.index(instr_sel)

    if str(fecha.year) != str(periodo):
        st.warning(
            f"⚠️ La fecha elegida ({fecha.year}) no coincide con el período "
            f"seleccionado arriba ({periodo}). La operación se guarda igual, "
            "pero no vas a verla en la tabla de abajo ni en Depuración/Patrimonio/"
            "Movimientos hasta que cambies el período a ese año — por eso a "
            "veces parece que 'no se cargó'."
        )

    if entry is None:
        instrumento = st.text_input("Instrumento (nombre libre)", key=f"inv_op_instman_{version}")
        tipo_sugerido = "otro"
    else:
        instrumento = f"{entry['ticker']} - {entry['nombre']}" if entry.get("ticker") else entry["nombre"]
        st.text_input("Instrumento", value=instrumento, disabled=True, key=f"inv_op_instro_{version}_{instr_idx}")
        tipo_sugerido = entry["tipo"]

    c3, c4, c5 = st.columns(3)
    tipo = c3.selectbox(
        "Tipo", idb.TIPOS_INSTRUMENTO,
        index=idb.TIPOS_INSTRUMENTO.index(tipo_sugerido),
        format_func=lambda t: idb.TIPO_LABEL.get(t, t), key=f"inv_op_tipo_{version}_{instr_idx}",
    )
    movimiento = c4.selectbox(
        "Movimiento", idb.MOVIMIENTOS,
        format_func=lambda m: idb.MOV_LABEL.get(m, m), key=f"inv_op_mov_{version}",
    )
    moneda = c5.radio("Moneda", ["ARS", "USD"], horizontal=True, key=f"inv_op_moneda_{version}")
    es_usd = moneda == "USD"

    sugerencia = idb.sugerir_fiscal(tipo, movimiento, moneda)
    combo_key = f"{version}_{tipo}_{movimiento}_{moneda}"

    c6, c7, c8 = st.columns(3)
    vn = c6.number_input("VN / Cantidad", min_value=0.0, step=1.0, key=f"inv_op_vn_{version}")
    precio = c7.number_input("Precio (en la moneda elegida)", min_value=0.0, step=0.01, key=f"inv_op_precio_{version}")
    if es_usd and movimiento in idb.MOVIMIENTOS_COSTO_BASE:
        tc_origen = c8.number_input(
            "TC origen", min_value=0.0, step=0.01, key=f"inv_op_tcorigen_{version}",
            help="TC al que se compraron los dólares usados en esta operación "
                 "(para calcular el costo fiscal en pesos).",
        )
    elif es_usd:
        tc_dia = idb.tc_para_movimiento(fecha.isoformat(), movimiento)
        tipo_tc = "comprador" if movimiento in ("dividendo", "renta") else "vendedor"
        c8.number_input(
            f"TC del día ({idb.MOV_LABEL.get(movimiento, movimiento)})",
            value=float(tc_dia or 0.0), step=0.01, disabled=True,
            # La key incluye la fecha: al ser un campo deshabilitado, Streamlit
            # solo usa `value` la primera vez que crea el widget — si la key no
            # cambia junto con la fecha, el campo queda "pegado" mostrando el
            # TC de la fecha anterior aunque el usuario elija otra fecha.
            key=f"inv_op_tcdia_{combo_key}_{fecha.isoformat()}",
            help=f"TC BNA {tipo_tc} del día de la operación (o el hábil más "
                 "cercano) — se toma automático de la base cargada en el "
                 "estudio, no hace falta cargarlo a mano.",
        )
        if tc_dia is None:
            c8.caption("⚠️ No hay TC BNA cargado cerca de esta fecha.")
        tc_origen = 0.0
    else:
        tc_origen = 0.0
        c8.caption("TC origen: no aplica en operaciones en pesos.")

    c9, c10, c11 = st.columns(3)
    fis = c9.selectbox(
        "Tratamiento fiscal (Ganancias)", idb.TRATAMIENTOS_FISCALES,
        index=idb.TRATAMIENTOS_FISCALES.index(sugerencia["tratamiento_fiscal"]),
        format_func=lambda f: idb.FIS_LABEL.get(f, f), key=f"inv_op_fis_{combo_key}",
    )
    bp = c10.selectbox(
        "Bienes Personales", idb.ALCANZA_BP,
        index=idb.ALCANZA_BP.index(sugerencia["alcanza_bienes_personales"]),
        format_func=lambda b: idb.ALCANZA_BP_LABEL.get(b, b), key=f"inv_op_bp_{combo_key}",
    )
    comision = c11.number_input("Comisión", min_value=0.0, step=0.01, key=f"inv_op_comision_{version}")

    notas = st.text_input("Notas", key=f"inv_op_notas_{version}")

    st.caption(
        "La sugerencia de tratamiento fiscal / Bienes Personales es un punto de "
        "partida (según el cuadro general del estudio) — revisala en casos "
        "particulares (ONs/FCI sin oferta pública, acciones sin cotización, "
        "ADRs, instrumentos en el exterior, etc.)."
    )

    if st.button("Guardar operación", key=f"inv_op_guardar_{version}"):
        if entry is None and not instrumento.strip():
            st.error("Ingresá el instrumento.")
            return

        fecha_str = fecha.isoformat()
        tc = idb.tc_para_movimiento(fecha_str, movimiento) or 0

        if es_usd:
            precio_usd, precio_ars = precio, precio * tc
            total_usd = vn * precio if vn > 0 else precio
            total_ars = total_usd * tc
        else:
            precio_ars, precio_usd = precio, (precio / tc if tc else 0)
            total_ars = vn * precio if vn > 0 else precio
            total_usd = total_ars / tc if tc else 0

        div_usd = div_ars = 0.0
        if movimiento in ("dividendo", "renta"):
            if es_usd:
                div_usd, div_ars = total_usd, total_ars
            else:
                div_ars = total_ars
            total_ars = total_usd = 0.0
        elif movimiento in ("amort", "caucion"):
            div_ars = total_ars
            total_ars = total_usd = 0.0

        costo_fiscal = (total_usd * tc_origen) if (tc_origen > 0 and es_usd) else total_ars
        asignaciones = (
            [{
                "lote_id": None, "cant_usd": total_usd, "tc_origen": tc_origen,
                "ars": total_usd * tc_origen, "origen": "manual", "fecha_lote": fecha_str,
            }]
            if (tc_origen > 0 and es_usd and total_usd > 0) else []
        )

        idb.crear_operacion(
            cliente_id=cliente["id"], fecha=fecha_str, instrumento=instrumento, tipo=tipo,
            movimiento=movimiento, tratamiento_fiscal=fis, vn=vn, precio_ars=precio_ars,
            precio_usd=precio_usd, tc=tc, tc_origen=tc_origen, total_ars=total_ars,
            total_usd=total_usd, div_usd=div_usd, div_ars=div_ars, comision=comision,
            moneda=moneda, notas=notas, costo_fiscal=costo_fiscal, asignaciones_usd=asignaciones,
            alcanza_bienes_personales=bp,
        )
        st.session_state["inv_op_version"] = version + 1
        st.success("Operación guardada ✓")
        st.rerun()


def _form_saldo_inicial(cliente: dict, periodo: str) -> None:
    st.markdown("**Saldo inicial (tenencia al 31/12 del año anterior)**")

    version = st.session_state.get("inv_si_version", 0)
    opciones_instr, mapa_instr = _catalogo_opciones()

    instr_sel = st.selectbox(
        "Instrumento (buscar por ticker o nombre)", opciones_instr,
        key=f"inv_si_instrsel_{version}",
    )
    entry = mapa_instr.get(instr_sel)
    instr_idx = opciones_instr.index(instr_sel)

    if entry is None:
        instrumento = st.text_input("Instrumento (nombre libre)", key=f"inv_si_instman_{version}")
        tipo_sugerido = "otro"
    else:
        instrumento = f"{entry['ticker']} - {entry['nombre']}" if entry.get("ticker") else entry["nombre"]
        st.text_input("Instrumento", value=instrumento, disabled=True, key=f"inv_si_instro_{version}_{instr_idx}")
        tipo_sugerido = entry["tipo"]

    c1, c2, c3 = st.columns(3)
    tipo = c1.selectbox(
        "Tipo", idb.TIPOS_INSTRUMENTO,
        index=idb.TIPOS_INSTRUMENTO.index(tipo_sugerido),
        format_func=lambda t: idb.TIPO_LABEL.get(t, t), key=f"inv_si_tipo_{version}_{instr_idx}",
    )
    vn = c2.number_input("VN / Cantidad", min_value=0.0, step=1.0, key=f"inv_si_vn_{version}")
    total_ars = c3.number_input("Valor total en pesos", min_value=0.0, step=100.0, key=f"inv_si_ars_{version}")

    # El valor en pesos suele salir de la última DDJJ (Bienes Personales /
    # Ganancias) del cliente. Si esa tenencia venía de una compra en dólares,
    # hace falta un TC de origen para que, cuando más adelante se venda parte
    # de esta posición, calcular_movimientos pueda separar Rendimiento vs.
    # Diferencia de cambio — igual que ya hace con una Compra asignada en
    # Depuración. Sin esto, un lote de Alta queda siempre en pesos puros (sin
    # TC) y su venta futura carga el 100% del resultado a Rendimiento, sin
    # posibilidad de diferencia de cambio, aunque la tenencia real haya sido
    # en dólares.
    fecha_str = f"{periodo}-01-01"
    tc_bna_aprox = idb.obtener_tc(fecha_str, "venta") or 0.0

    c4, c5 = st.columns(2)
    usa_tc_origen = c4.checkbox(
        "¿Costo original en dólares?", value=(tc_bna_aprox > 0),
        key=f"inv_si_usatc_{version}",
        help="Marcalo si esta tenencia viene de una compra en dólares (aunque "
             "el valor de arriba esté en pesos, como en la DDJJ). Así, cuando "
             "más adelante vendas parte de esta posición, el resultado se "
             "separa en Rendimiento vs. Diferencia de cambio en vez de "
             "cargarse todo a Rendimiento.",
    )
    # Sugerencia de TC de origen: Valor total en pesos ÷ VN/Cantidad. Asume
    # que el VN cargado equivale 1 a 1 al monto en dólares de origen (así es
    # para bonos/ONs/Letras, donde el VN ya está denominado en dólares por
    # convención de mercado; para acciones/CEDEARs/FCI, donde el VN es
    # cantidad de unidades, el resultado deja de ser un TC real salvo que la
    # unidad valga ≈ USD 1 — el usuario puede corregirlo a mano en ese caso).
    # Si todavía no se cargaron VN o pesos, se usa el TC BNA histórico como
    # referencia de partida.
    tc_sugerido = (total_ars / vn) if (vn > 0 and total_ars > 0) else (tc_bna_aprox or 0.0)

    tc_origen = 0.0
    if usa_tc_origen:
        tc_origen = c5.number_input(
            "TC de origen (aprox.)", min_value=0.0, step=0.01,
            value=float(tc_sugerido), key=f"inv_si_tcorigen_{version}",
            help="Precargado como Valor total en pesos ÷ VN/Cantidad (el TC "
                 "implícito en lo que se pagó por esta tenencia). Si todavía "
                 "no cargaste VN o el valor en pesos, se precarga con el TC "
                 "BNA vendedor al 01/01 del período como referencia. Es una "
                 "aproximación — corregilo si tenés un dato más preciso.",
        )
        if vn > 0 and total_ars > 0 and abs(tc_origen - tc_sugerido) > 0.01:
            c5.caption(f"↻ Sugerido con los valores actuales: {_fmt_cantidad(tc_sugerido)}")
        if tc_origen > 0 and total_ars > 0:
            c5.caption(f"≈ {_fmt_usd(total_ars / tc_origen)}")
        if tc_bna_aprox <= 0:
            c5.caption("⚠️ Sin TC BNA cargado cerca de esta fecha — ingresalo a mano.")

    sugerencia = idb.sugerir_fiscal(tipo, "apertura", "ARS")

    if st.button("Guardar saldo inicial", key=f"inv_si_guardar_{version}"):
        if entry is None and not instrumento.strip():
            st.error("Ingresá el instrumento.")
            return

        es_usd_origen = usa_tc_origen and tc_origen > 0
        total_usd = (total_ars / tc_origen) if es_usd_origen else 0.0

        idb.crear_operacion(
            cliente_id=cliente["id"], fecha=fecha_str, instrumento=instrumento, tipo=tipo,
            movimiento="apertura", tratamiento_fiscal=idb.default_fis(tipo),
            vn=vn, precio_ars=(total_ars / vn if vn else total_ars),
            precio_usd=(total_usd / vn if (vn and es_usd_origen) else 0),
            tc=1, tc_origen=(tc_origen if es_usd_origen else 0),
            total_ars=total_ars, total_usd=total_usd, div_usd=0, div_ars=0,
            comision=0, moneda=("USD" if es_usd_origen else "ARS"),
            notas="Saldo inicial DDJJ anterior",
            costo_fiscal=total_ars, es_saldo_inicial=True,
            alcanza_bienes_personales=sugerencia["alcanza_bienes_personales"],
        )
        st.session_state["inv_si_version"] = version + 1
        st.success("Saldo inicial guardado ✓")
        st.rerun()


def _importar_plantilla(cliente: dict) -> None:
    st.markdown("**Importar planilla (Excel/CSV)**")
    st.caption(
        "Columnas reconocidas por nombre (sin importar mayúsculas ni orden): "
        "Fecha, Instrumento, Tipo, Movimiento, VN/Cantidad, Precio, Moneda, "
        "TC Origen, Importe neto, Comisión. Se importan todas las filas para "
        "el cliente seleccionado arriba."
    )
    archivo = st.file_uploader("Archivo", type=["xlsx", "xls", "csv"], key="inv_import_file")
    if archivo is None:
        return

    try:
        if archivo.name.lower().endswith(".csv"):
            df = pd.read_csv(archivo, header=None, dtype=object)
        else:
            df = pd.read_excel(archivo, header=None, dtype=object)
    except Exception as e:
        st.error(f"No pude leer el archivo: {e}")
        return

    header_row = None
    for i in range(min(10, len(df))):
        fila = [_normalizar(c) for c in df.iloc[i].tolist()]
        if any(c in ("fecha", "instrumento", "movimiento") for c in fila):
            header_row = i
            break
    if header_row is None:
        st.error("No encontré una fila de encabezados (Fecha / Instrumento / Movimiento).")
        return

    headers = [_normalizar(c) for c in df.iloc[header_row].tolist()]

    def col(*kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return -1

    i_f = col("fecha")
    i_inst = col("instrumento", "ticker", "especie")
    i_tipo = col("tipo")
    i_mov = col("movimiento", "mov")
    i_vn = col("vn", "cantidad", "cuotapart", "nominal")
    i_prec = col("precio")
    i_mon = col("moneda", "currency")
    i_tco = col("tc origen", "tc_origen", "origen")
    i_neto = col("importe neto", "neto", "imp neto", "net")
    i_comis = col("comision", "comision", "com ")

    if i_f < 0 or i_inst < 0:
        st.error("La planilla debe tener al menos columnas Fecha e Instrumento.")
        return

    mov_map = {
        "compra": "compra", "venta": "venta", "dividendo": "dividendo", "div": "dividendo",
        "renta": "renta", "amort": "amort", "amortizacion": "amort", "apertura": "apertura",
        "suscripcion": "apertura", "rescate": "rescate", "caucion": "caucion",
    }

    n, sk = 0, 0
    for _, row in df.iloc[header_row + 1:].iterrows():
        vals = row.tolist()
        if not any(not _vacio(v) for v in vals):
            continue

        fecha_str = _parse_fecha(vals[i_f] if i_f >= 0 and i_f < len(vals) else None)
        if not fecha_str:
            sk += 1
            continue
        instrumento = str(vals[i_inst]).strip() if i_inst >= 0 and not _vacio(vals[i_inst]) else ""
        if not instrumento or instrumento == "-" or re.search(r"ejemplo|borr[aá]|muestra", instrumento, re.I):
            sk += 1
            continue

        tipo_raw = str(vals[i_tipo]).strip().lower() if i_tipo >= 0 and not _vacio(vals[i_tipo]) else ""
        tipo = tipo_raw if tipo_raw in idb.TIPOS_INSTRUMENTO else _detectar_tipo(instrumento)
        mov_raw = str(vals[i_mov]).strip().lower() if i_mov >= 0 and not _vacio(vals[i_mov]) else "compra"
        es_saldo_inicial = mov_raw in _MOV_SALDO_INICIAL_RAW
        movimiento = "apertura" if es_saldo_inicial else mov_map.get(mov_raw, mov_raw if mov_raw in idb.MOVIMIENTOS else "compra")

        mon_raw = str(vals[i_mon]).strip().upper() if i_mon >= 0 and not _vacio(vals[i_mon]) else ""
        vn = abs(_pv(vals[i_vn])) if i_vn >= 0 else 0.0
        tc_origen = _pv(vals[i_tco]) if (i_tco >= 0 and movimiento in idb.MOVIMIENTOS_COSTO_BASE) else 0.0
        precio = abs(_pv(vals[i_prec])) if i_prec >= 0 else 0.0
        es_usd = mon_raw == "USD"

        tc = idb.tc_para_movimiento(fecha_str, movimiento) or 1

        if es_usd:
            precio_usd, precio_ars = precio, precio * tc
            total_usd = vn * precio if vn > 0 else precio
            total_ars = total_usd * tc
        else:
            precio_ars, precio_usd = precio, (precio / tc if tc else 0)
            total_ars = vn * precio if vn > 0 else precio
            total_usd = total_ars / tc if tc else 0

        div_usd = div_ars = 0.0
        if movimiento in ("dividendo", "renta"):
            if es_usd:
                div_usd, div_ars = total_usd, total_ars
            else:
                div_ars = total_ars
            total_ars = total_usd = 0.0
        elif movimiento in ("amort", "caucion"):
            div_ars = total_ars
            total_ars = total_usd = 0.0

        bruto = total_usd if es_usd else total_ars
        neto_v = _pv(vals[i_neto]) if i_neto >= 0 else 0.0
        comis_directa = _pv(vals[i_comis]) if i_comis >= 0 else 0.0
        if neto_v > 0 and bruto > 0:
            comision = abs(abs(bruto) - abs(neto_v))
        elif comis_directa > 0:
            comision = comis_directa
        else:
            comision = 0.0

        costo_fiscal = (total_usd * tc_origen) if (tc_origen > 0 and es_usd) else total_ars
        sugerencia = idb.sugerir_fiscal(tipo, movimiento, "USD" if es_usd else "ARS")

        idb.crear_operacion(
            cliente_id=cliente["id"], fecha=fecha_str, instrumento=instrumento, tipo=tipo,
            movimiento=movimiento, tratamiento_fiscal=sugerencia["tratamiento_fiscal"],
            vn=vn, precio_ars=precio_ars, precio_usd=precio_usd, tc=tc, tc_origen=tc_origen,
            total_ars=total_ars, total_usd=total_usd, div_usd=div_usd, div_ars=div_ars,
            comision=0 if es_saldo_inicial else comision, moneda="USD" if es_usd else "ARS",
            notas="Saldo inicial DDJJ anterior" if es_saldo_inicial else "",
            costo_fiscal=costo_fiscal, es_saldo_inicial=es_saldo_inicial,
            alcanza_bienes_personales=sugerencia["alcanza_bienes_personales"],
        )
        n += 1

    msg = f"✓ {n} operaciones importadas."
    if sk:
        msg += f" ({sk} omitidas)"
    st.success(msg)


def _tabla_operaciones(cliente: dict, periodo: str) -> None:
    ops = idb.listar_operaciones(cliente_id=cliente["id"], periodo=periodo)
    st.markdown(f"**Operaciones cargadas — {periodo}** ({len(ops)})")
    if not ops:
        st.info("Sin operaciones para este cliente y período.")
        return

    filas = [{
        "ID": o["id"],
        "Fecha": o["fecha"],
        "Instrumento": o["instrumento"],
        "Tipo": idb.TIPO_LABEL.get(o["tipo"], o["tipo"]),
        "Movimiento": idb.MOV_LABEL.get(o["movimiento"], o["movimiento"]) + (" (saldo inicial)" if o["es_saldo_inicial"] else ""),
        "Trat. fiscal": idb.FIS_LABEL.get(o["tratamiento_fiscal"], o["tratamiento_fiscal"]),
        "Bienes Personales": idb.ALCANZA_BP_LABEL.get(
            o.get("alcanza_bienes_personales", "gravado"), o.get("alcanza_bienes_personales", "")
        ),
        "VN": _fmt_cantidad(o["vn"]),
        "Total ARS": _fmt_pesos(o["total_ars"]),
        "Total USD": _fmt_usd(o["total_usd"]),
        "Div. ARS": _fmt_pesos(o["div_ars"]),
        "Comisión": _fmt_moneda(o["comision"], o["moneda"]),
        "Moneda": o["moneda"],
    } for o in ops]
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    _editar_operacion(ops)

    with st.expander("Eliminar una operación"):
        id_borrar = st.number_input("ID a eliminar", min_value=0, step=1, key="inv_op_id_borrar")
        if st.button("Eliminar", key="inv_op_btn_borrar") and id_borrar:
            idb.eliminar_operacion(int(id_borrar))
            st.success("Operación eliminada.")
            st.rerun()


def _editar_operacion(ops: list[dict]) -> None:
    with st.expander("Editar una operación"):
        opciones_op = {f"#{o['id']} — {o['fecha']} — {o['instrumento']}": o for o in ops}
        clave_sel = st.selectbox(
            "Operación a editar", list(opciones_op.keys()), key="inv_editar_op_sel",
        )
        op = opciones_op[clave_sel]
        op_id = op["id"]

        st.caption(
            "Modificá los datos y guardá — se recalculan los totales en ARS/USD "
            "igual que al cargar una operación nueva. Si cambiás mucho el total "
            "en USD de una compra ya asignada en Depuración, puede que haga "
            "falta reasignar los lotes del pool ahí."
        )

        c1, c2 = st.columns([1, 2])
        fecha = c1.date_input(
            "Fecha", value=_dt.date.fromisoformat(op["fecha"][:10]), key=f"inv_edit_fecha_{op_id}",
        )
        instrumento = c2.text_input(
            "Instrumento", value=op["instrumento"], key=f"inv_edit_instr_{op_id}",
        )

        c3, c4, c5 = st.columns(3)
        tipo = c3.selectbox(
            "Tipo", idb.TIPOS_INSTRUMENTO,
            index=(idb.TIPOS_INSTRUMENTO.index(op["tipo"]) if op["tipo"] in idb.TIPOS_INSTRUMENTO else 0),
            format_func=lambda t: idb.TIPO_LABEL.get(t, t), key=f"inv_edit_tipo_{op_id}",
        )
        movimiento = c4.selectbox(
            "Movimiento", idb.MOVIMIENTOS,
            index=(idb.MOVIMIENTOS.index(op["movimiento"]) if op["movimiento"] in idb.MOVIMIENTOS else 0),
            format_func=lambda m: idb.MOV_LABEL.get(m, m), key=f"inv_edit_mov_{op_id}",
        )
        moneda = c5.radio(
            "Moneda", ["ARS", "USD"], horizontal=True,
            index=(0 if op["moneda"] == "ARS" else 1), key=f"inv_edit_moneda_{op_id}",
        )
        es_usd = moneda == "USD"

        c6, c7, c8 = st.columns(3)
        vn = c6.number_input(
            "VN / Cantidad", min_value=0.0, step=1.0, value=float(op["vn"] or 0),
            key=f"inv_edit_vn_{op_id}",
        )
        precio_actual = op["precio_usd"] if op["moneda"] == "USD" else op["precio_ars"]
        precio = c7.number_input(
            "Precio (en la moneda elegida)", min_value=0.0, step=0.01,
            value=float(precio_actual or 0), key=f"inv_edit_precio_{op_id}",
        )
        if es_usd and movimiento in idb.MOVIMIENTOS_COSTO_BASE:
            tc_origen = c8.number_input(
                "TC origen", min_value=0.0, step=0.01, value=float(op["tc_origen"] or 0),
                key=f"inv_edit_tcorigen_{op_id}",
            )
        elif es_usd:
            tc_dia = idb.tc_para_movimiento(fecha.isoformat(), movimiento)
            c8.number_input(
                f"TC del día ({idb.MOV_LABEL.get(movimiento, movimiento)})",
                value=float(tc_dia if tc_dia is not None else (op["tc"] or 0)), step=0.01, disabled=True,
                key=f"inv_edit_tcdia_{op_id}_{movimiento}_{fecha.isoformat()}",
            )
            tc_origen = 0.0
        else:
            tc_origen = 0.0
            c8.caption("TC origen: no aplica en operaciones en pesos.")

        c9, c10, c11 = st.columns(3)
        fis = c9.selectbox(
            "Tratamiento fiscal (Ganancias)", idb.TRATAMIENTOS_FISCALES,
            index=(idb.TRATAMIENTOS_FISCALES.index(op["tratamiento_fiscal"])
                   if op["tratamiento_fiscal"] in idb.TRATAMIENTOS_FISCALES else 0),
            format_func=lambda f: idb.FIS_LABEL.get(f, f), key=f"inv_edit_fis_{op_id}",
        )
        bp_actual = op.get("alcanza_bienes_personales", "gravado")
        bp = c10.selectbox(
            "Bienes Personales", idb.ALCANZA_BP,
            index=(idb.ALCANZA_BP.index(bp_actual) if bp_actual in idb.ALCANZA_BP else 0),
            format_func=lambda b: idb.ALCANZA_BP_LABEL.get(b, b), key=f"inv_edit_bp_{op_id}",
        )
        comision = c11.number_input(
            "Comisión", min_value=0.0, step=0.01, value=float(op["comision"] or 0),
            key=f"inv_edit_comision_{op_id}",
        )

        notas = st.text_input("Notas", value=op["notas"] or "", key=f"inv_edit_notas_{op_id}")

        if st.button("Guardar cambios", key=f"inv_edit_guardar_{op_id}"):
            if not instrumento.strip():
                st.error("Ingresá el instrumento.")
                return

            fecha_str = fecha.isoformat()
            tc = idb.tc_para_movimiento(fecha_str, movimiento) or 0

            if es_usd:
                precio_usd, precio_ars = precio, precio * tc
                total_usd = vn * precio if vn > 0 else precio
                total_ars = total_usd * tc
            else:
                precio_ars, precio_usd = precio, (precio / tc if tc else 0)
                total_ars = vn * precio if vn > 0 else precio
                total_usd = total_ars / tc if tc else 0

            div_usd = div_ars = 0.0
            if movimiento in ("dividendo", "renta"):
                if es_usd:
                    div_usd, div_ars = total_usd, total_ars
                else:
                    div_ars = total_ars
                total_ars = total_usd = 0.0
            elif movimiento in ("amort", "caucion"):
                div_ars = total_ars
                total_ars = total_usd = 0.0

            costo_fiscal = (total_usd * tc_origen) if (tc_origen > 0 and es_usd) else total_ars

            idb.actualizar_operacion(
                operacion_id=op_id, fecha=fecha_str, instrumento=instrumento, tipo=tipo,
                movimiento=movimiento, tratamiento_fiscal=fis, vn=vn, precio_ars=precio_ars,
                precio_usd=precio_usd, tc=tc, tc_origen=tc_origen, total_ars=total_ars,
                total_usd=total_usd, div_usd=div_usd, div_ars=div_ars, comision=comision,
                moneda=moneda, notas=notas, costo_fiscal=costo_fiscal,
                alcanza_bienes_personales=bp,
            )
            st.success("Operación actualizada ✓")
            st.rerun()


def _tab_tenencia_usd(cliente: dict) -> None:
    st.caption(
        "Acá se arma la tenencia de dólares del cliente: cada compra o "
        "recepción de USD (MEP, oficial, informal, remesa) se carga como un "
        "lote con su TC de origen. Esos lotes se van consumiendo por orden "
        "PEPS (el más antiguo primero) a medida que se usan para pagar "
        "compras en USD en Depuración — la tabla de abajo muestra cuánto de "
        "cada lote ya se usó y cuánto queda disponible."
    )

    lotes_actuales = idb.listar_pool_usd(cliente["id"])
    total_cargado = sum(l["cant_usd"] for l in lotes_actuales)
    total_disponible = sum(l["disponible_usd"] for l in lotes_actuales)
    total_usado = total_cargado - total_disponible

    m1, m2, m3 = st.columns(3)
    m1.metric("USD cargados (histórico)", _fmt_usd(total_cargado))
    m2.metric("USD ya utilizados", _fmt_usd(total_usado))
    m3.metric("USD disponibles (tenencia actual)", _fmt_usd(total_disponible))
    st.divider()

    _pool_usd(cliente)


def _tab_depuracion(cliente: dict, periodo: str) -> None:
    st.caption(
        "Revisá cada instrumento: confirmá si la operación fue en ARS o USD. "
        "Para compras en USD, asignale lotes del pool con su TC de origen "
        "(cargados en la solapa «Tenencia USD») — las operaciones en ARS no "
        "requieren nada acá."
    )

    _depuracion_por_instrumento(cliente, periodo)


def _pool_usd(cliente: dict) -> None:
    st.markdown("**Pool de dólares (lotes con TC de origen)**")
    st.caption(
        "Cargá los dólares que el cliente compró o recibió (MEP, oficial, "
        "informal, remesa) con su TC de origen. Se usan para asignar el "
        "costo fiscal de las compras en USD por orden PEPS (más antiguo primero)."
    )

    version = st.session_state.get("inv_pool_version", 0)
    c1, c2, c3, c4 = st.columns(4)
    fecha = c1.date_input("Fecha", value=_dt.date.today(), key=f"inv_pool_fecha_{version}")
    cant = c2.number_input("Cantidad USD", min_value=0.0, step=1.0, key=f"inv_pool_cant_{version}")
    tc_origen = c3.number_input("TC de origen", min_value=0.0, step=0.01, key=f"inv_pool_tc_{version}")
    origen = c4.selectbox(
        "Origen", idb.ORIGENES_POOL,
        format_func=lambda o: idb.ORIGEN_POOL_LABEL.get(o, o), key=f"inv_pool_origen_{version}",
    )
    desc = st.text_input("Descripción (opcional)", key=f"inv_pool_desc_{version}")

    if st.button("+ Agregar lote", key=f"inv_pool_guardar_{version}"):
        if cant <= 0 or tc_origen <= 0:
            st.error("Ingresá cantidad de USD y TC de origen mayores a cero.")
        else:
            idb.crear_lote_pool(
                cliente_id=cliente["id"], fecha=fecha.isoformat(), cant_usd=cant,
                tc_origen=tc_origen, origen=origen, descripcion=desc,
            )
            st.session_state["inv_pool_version"] = version + 1
            st.success("Lote agregado ✓")
            st.rerun()

    lotes = idb.listar_pool_usd(cliente["id"])
    if not lotes:
        st.info("Sin lotes cargados todavía.")
        return

    filas = [{
        "ID": l["id"], "Fecha": l["fecha"],
        "Origen": idb.ORIGEN_POOL_LABEL.get(l["origen"], l["origen"]),
        "Cant. USD": _fmt_usd(l["cant_usd"], 4),
        "TC origen": _fmt_cantidad(l["tc_origen"]),
        "Costo ARS": _fmt_pesos(l["cant_usd"] * l["tc_origen"], 0),
        "Disponible USD": _fmt_usd(l["disponible_usd"], 4),
        "Descripción": l["descripcion"],
    } for l in lotes]
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    with st.expander("Eliminar un lote"):
        id_borrar = st.number_input("ID a eliminar", min_value=0, step=1, key="inv_pool_id_borrar")
        if st.button("Eliminar lote", key="inv_pool_btn_borrar") and id_borrar:
            idb.eliminar_lote_pool(int(id_borrar))
            st.success("Lote eliminado.")
            st.rerun()


def _depuracion_por_instrumento(cliente: dict, periodo: str) -> None:
    ops = idb.listar_operaciones(cliente_id=cliente["id"], periodo=periodo)
    if not ops:
        st.info("Sin operaciones cargadas para este cliente y período.")
        return

    grupos: dict[str, list[dict]] = {}
    for o in ops:
        grupos.setdefault(o["instrumento"], []).append(o)

    pendientes = 0
    resumen = {}
    for instrumento, ops_inst in grupos.items():
        compras_usd = [
            o for o in ops_inst
            if o["moneda"] == "USD" and o["movimiento"] in ("compra", "apertura")
        ]
        sin_completar = [
            o for o in compras_usd
            if sum(a["cant_usd"] for a in o["asignaciones_usd"]) < o["total_usd"] * 0.95
        ]
        resumen[instrumento] = (compras_usd, sin_completar)
        if sin_completar:
            pendientes += 1

    st.markdown(f"**Instrumentos — {periodo}** ({len(grupos)})")
    if pendientes:
        st.warning(f"⚠️ {pendientes} instrumento(s) con compras en USD sin TC de origen completo.")

    for instrumento, ops_inst in grupos.items():
        compras_usd, sin_completar = resumen[instrumento]
        todo_ars = all(o["moneda"] == "ARS" for o in ops_inst)
        if todo_ars:
            estado = "✅ Todo ARS"
        elif sin_completar:
            estado = f"⚠️ {len(sin_completar)} compra(s) USD sin TC origen"
        else:
            estado = "✅ USD asignado"

        tipo_inst = idb.TIPO_LABEL.get(ops_inst[0]["tipo"], ops_inst[0]["tipo"])
        with st.expander(f"{instrumento} — {tipo_inst} — {estado}"):
            filas = [{
                "ID": o["id"], "Fecha": o["fecha"],
                "Movimiento": idb.MOV_LABEL.get(o["movimiento"], o["movimiento"]),
                "VN": _fmt_cantidad(o["vn"]), "Moneda": o["moneda"],
                "Total ARS": _fmt_pesos(o["total_ars"]), "Total USD": _fmt_usd(o["total_usd"], 4),
                "TC BNA": _fmt_cantidad(o["tc"]),
                "USD asignado": _fmt_usd(sum(a["cant_usd"] for a in o["asignaciones_usd"]), 4),
                "Costo fiscal $": _fmt_pesos(o["costo_fiscal"]),
            } for o in ops_inst]
            st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

            if compras_usd:
                st.caption("Lotes del pool USD asignados a cada compra:")
                for o in compras_usd:
                    _fila_asignacion_usd(cliente, o)

            st.markdown("_Corregir moneda o tratamiento fiscal de este instrumento_")
            ids_grupo = [o["id"] for o in ops_inst]
            cm1, cm2, cm3 = st.columns([1, 1, 1])
            id_corregir = cm1.selectbox(
                "Operación (ID)", ids_grupo, key=f"dep_monid_{cliente['id']}_{periodo}_{instrumento}",
            )
            nueva_moneda = cm2.radio(
                "Nueva moneda", ["ARS", "USD"], horizontal=True,
                key=f"dep_monval_{cliente['id']}_{periodo}_{instrumento}",
            )
            if cm3.button("Aplicar moneda", key=f"dep_monbtn_{cliente['id']}_{periodo}_{instrumento}"):
                idb.cambiar_moneda_operacion(id_corregir, nueva_moneda)
                st.success("Moneda actualizada.")
                st.rerun()

            fis_actual = ops_inst[0]["tratamiento_fiscal"]
            cf1, cf2 = st.columns([2, 1])
            nuevo_fis = cf1.selectbox(
                "Tratamiento fiscal (aplica a todas las operaciones de este instrumento en el período)",
                idb.TRATAMIENTOS_FISCALES, index=idb.TRATAMIENTOS_FISCALES.index(fis_actual),
                format_func=lambda f: idb.FIS_LABEL.get(f, f),
                key=f"dep_fis_{cliente['id']}_{periodo}_{instrumento}",
            )
            if cf2.button("Actualizar", key=f"dep_fis_btn_{cliente['id']}_{periodo}_{instrumento}"):
                idb.cambiar_tratamiento_fiscal_instrumento(cliente["id"], instrumento, periodo, nuevo_fis)
                st.success("Tratamiento fiscal actualizado.")
                st.rerun()


def _fila_asignacion_usd(cliente: dict, op: dict) -> None:
    asignado = sum(a["cant_usd"] for a in op["asignaciones_usd"])
    restante = max(0.0, op["total_usd"] - asignado)

    st.markdown(f"— **{op['fecha']}** · VN {op['vn']:g} · USD {op['total_usd']:.4f} total")

    for i, a in enumerate(op["asignaciones_usd"]):
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1])
        c1.caption(a.get("fecha_lote", ""))
        c2.caption(idb.ORIGEN_POOL_LABEL.get(a.get("origen", ""), a.get("origen", "")))
        c3.caption(f"USD {a['cant_usd']:.4f}")
        c4.caption(f"TC {a['tc_origen']:.2f}")
        if c5.button("Quitar", key=f"dep_quitar_{op['id']}_{i}"):
            idb.quitar_asignacion(op["id"], i)
            st.rerun()

    if restante > 0.0001:
        st.caption(f"Resta asignar: USD {restante:.4f}")
    else:
        st.caption("✓ Completo")

    if st.button("Asignar PEPS automático", key=f"dep_peps_{op['id']}"):
        idb.asignar_peps_automatico(op["id"])
        st.rerun()

    lotes_disp = [l for l in idb.listar_pool_usd(cliente["id"]) if l["disponible_usd"] > 0.0001]
    if lotes_disp:
        opciones_lote = {
            f"{l['fecha']} · {idb.ORIGEN_POOL_LABEL.get(l['origen'], l['origen'])} · "
            f"disp. USD {l['disponible_usd']:.4f} · TC {l['tc_origen']:.2f}": l
            for l in lotes_disp
        }
        cl1, cl2, cl3 = st.columns([2, 1, 1])
        lote_sel = cl1.selectbox(
            "Lote a asignar manualmente", list(opciones_lote.keys()), key=f"dep_lotesel_{op['id']}",
        )
        lote = opciones_lote[lote_sel]
        valor_sugerido = float(min(restante, lote["disponible_usd"])) if restante > 0 else 0.0
        cant_usar = cl2.number_input(
            "USD a usar", min_value=0.0, value=valor_sugerido, step=0.01, key=f"dep_lotecant_{op['id']}",
        )
        if cl3.button("Asignar", key=f"dep_loteasign_{op['id']}"):
            if cant_usar > 0:
                idb.asignar_lote_a_operacion(op["id"], lote["id"], cant_usar)
                st.rerun()
    st.divider()


def _tab_patrimonio(cliente: dict, periodo: str) -> None:
    st.caption(
        "Tenencia al cierre del ejercicio (31/12 para Persona Humana), valuada al "
        "costo histórico PEPS — se consumen las compras en el orden en que se "
        "cargaron a medida que aparecen ventas o rescates. Por ahora la valuación "
        "es «al origen» (sin reexpresión por inflación): pensado para clientes "
        "Persona Humana."
    )

    posicion = idb.calcular_posicion(cliente["id"], periodo)
    if not posicion:
        st.info("Sin operaciones cargadas para este cliente y período.")
        return

    _metricas_patrimonio(posicion)
    st.divider()
    _grafico_composicion(posicion)
    st.divider()
    _tabla_posicion(posicion)
    st.divider()
    _control_cierre(cliente, periodo, posicion)
    st.divider()
    _resumen_patrimonial(cliente, periodo)


def _metricas_patrimonio(posicion: list[dict]) -> None:
    mantenidas = [p for p in posicion if p["estado"] == "mantenida"]
    costo_hist = sum(p["costo_ars"] for p in mantenidas)
    div_cobrados = sum(p["div_ars"] for p in posicion)
    rentas_amorts = sum(p["amorts_ars"] + p["rentas_ars"] for p in posicion)
    activos = len(mantenidas)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Costo histórico invertido", _fmt_pesos(costo_hist, 0))
    c2.metric("Dividendos/rentas cobrados", _fmt_pesos(div_cobrados, 0))
    c3.metric("Amort. + rentas fija/caución", _fmt_pesos(rentas_amorts, 0))
    c4.metric("Activos en cartera", _fmt_cantidad(activos, 0))


def _grafico_composicion(posicion: list[dict]) -> None:
    mantenidas = [p for p in posicion if p["estado"] == "mantenida" and p["costo_ars"] > 0]
    if not mantenidas:
        st.info("Sin tenencias abiertas al cierre del período.")
        return

    st.markdown("**Composición por tipo de instrumento (costo histórico $)**")
    df_tipo = (
        pd.DataFrame([{"Tipo": idb.TIPO_LABEL.get(p["tipo"], p["tipo"]), "Costo $": p["costo_ars"]} for p in mantenidas])
        .groupby("Tipo", as_index=True)["Costo $"].sum().sort_values(ascending=False)
    )
    st.bar_chart(df_tipo)

    st.markdown("**Tratamiento fiscal de la cartera (costo histórico $)**")
    df_fis = (
        pd.DataFrame([{
            "Trat. fiscal": idb.FIS_LABEL.get(p["tratamiento_fiscal"], p["tratamiento_fiscal"]),
            "Costo $": p["costo_ars"],
        } for p in mantenidas])
        .groupby("Trat. fiscal", as_index=True)["Costo $"].sum().sort_values(ascending=False)
    )
    st.bar_chart(df_fis)


def _tabla_posicion(posicion: list[dict]) -> None:
    st.markdown(f"**Tenencias al cierre — costo PEPS histórico** ({len(posicion)} instrumentos)")
    filas = [{
        "Instrumento": p["instrumento"],
        "Tipo": idb.TIPO_LABEL.get(p["tipo"], p["tipo"]),
        "Trat. fiscal": idb.FIS_LABEL.get(p["tratamiento_fiscal"], p["tratamiento_fiscal"]),
        "Bienes Personales": idb.ALCANZA_BP_LABEL.get(p["alcanza_bienes_personales"], p["alcanza_bienes_personales"]),
        "Cant. / VN": _fmt_cantidad(p["cantidad"], 4),
        "Val. hist. unit. $": _fmt_pesos(p["precio_unit_ars"]),
        "Costo total $ (PEPS)": _fmt_pesos(p["costo_ars"]),
        "Costo total USD": _fmt_usd(p["costo_usd"]),
        "Costo $ fiscal (TC origen)": _fmt_pesos(p["costo_fiscal"]),
        "Div./Renta cobrados $": _fmt_pesos(p["div_ars"]),
        "Amort. $": _fmt_pesos(p["amorts_ars"]),
        "Estado": "Mantenida" if p["estado"] == "mantenida" else "Cerrada",
    } for p in posicion]
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)


def _control_cierre(cliente: dict, periodo: str, posicion: list[dict]) -> None:
    st.markdown("**Control de cierre (opcional)**")
    st.caption(
        "Cargá acá la cantidad real según el extracto/comprobante del broker al "
        "cierre (31/12) para detectar diferencias contra la posición PEPS "
        "calculada arriba — por ejemplo, operaciones del período que todavía no "
        "se cargaron."
    )

    version = st.session_state.get("pat_control_version", 0)
    instrumentos_pos = [p["instrumento"] for p in posicion] or [_OPCION_MANUAL]
    c1, c2, c3 = st.columns([2, 1, 1])
    instrumento = c1.selectbox("Instrumento", instrumentos_pos, key=f"pat_ctrl_inst_{version}")
    vn_control = c2.number_input("Cantidad real (extracto)", min_value=0.0, step=1.0, key=f"pat_ctrl_vn_{version}")
    valor_control = c3.number_input("Valor $ real (opcional)", min_value=0.0, step=100.0, key=f"pat_ctrl_val_{version}")
    notas = st.text_input("Notas (opcional)", key=f"pat_ctrl_notas_{version}")

    if st.button("+ Cargar control de cierre", key=f"pat_ctrl_guardar_{version}"):
        idb.crear_tenencia_control(
            cliente_id=cliente["id"], periodo=periodo, instrumento=instrumento,
            vn_control=vn_control, valor_ars_control=valor_control, notas=notas,
        )
        st.session_state["pat_control_version"] = version + 1
        st.success("Control de cierre guardado ✓")
        st.rerun()

    controles = idb.listar_tenencias_control(cliente["id"], periodo)
    if not controles:
        return

    posicion_por_inst = {p["instrumento"]: p for p in posicion}
    filas = []
    for c in controles:
        calc = posicion_por_inst.get(c["instrumento"])
        cant_calc = calc["cantidad"] if calc else 0.0
        diff = c["vn_control"] - cant_calc
        filas.append({
            "ID": c["id"], "Instrumento": c["instrumento"],
            "Cant. calculada (PEPS)": _fmt_cantidad(cant_calc, 4),
            "Cant. real (extracto)": _fmt_cantidad(c["vn_control"], 4),
            "Diferencia": _fmt_cantidad(diff, 4),
            "¿Coincide?": "✓" if abs(diff) < 0.0001 else "⚠️ revisar",
            "Notas": c["notas"],
        })
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    con_diferencia = [f for f in filas if f["¿Coincide?"] != "✓"]
    if con_diferencia:
        st.warning(
            f"⚠️ {len(con_diferencia)} instrumento(s) con diferencia entre la posición "
            "calculada y el extracto — revisá si falta cargar alguna operación del período."
        )

    with st.expander("Eliminar un control de cierre"):
        id_borrar = st.number_input("ID a eliminar", min_value=0, step=1, key="pat_ctrl_id_borrar")
        if st.button("Eliminar", key="pat_ctrl_btn_borrar") and id_borrar:
            idb.eliminar_tenencia_control(int(id_borrar))
            st.success("Control eliminado.")
            st.rerun()


def _resumen_patrimonial(cliente: dict, periodo: str) -> None:
    st.markdown("### 📊 Resumen patrimonial (Ganancias / Bienes Personales)")
    st.caption(
        "Composición de la cartera a costo histórico (Ganancias) y a valor de "
        "cierre (Bienes Personales, según las planillas oficiales de "
        "valuaciones 2025 o la carga manual de abajo), variación patrimonial "
        "del período e impuesto aproximado por operación/instrumento. Todos "
        "los totales salen de la posición y los movimientos calculados arriba."
    )

    alic_key = f"pat_alicuota_escala_{cliente['id']}_{periodo}"
    alicuota_pct = st.number_input(
        "Alícuota de referencia para 'Gravado a escala' (%)",
        min_value=0.0, max_value=100.0, step=1.0,
        value=float(st.session_state.get(alic_key, idb.ALICUOTA_ESCALA_DEFAULT * 100)),
        key=alic_key,
        help="El tratamiento 'Gravado a escala' (Impuesto a las Ganancias, "
             "persona humana) no tiene una tasa fija: depende de TODOS los "
             "ingresos anuales del cliente, no solo de Inversiones. Se usa "
             "esta alícuota de referencia (por defecto, la escala máxima "
             "vigente, 35%) para estimar el impuesto — ajustala al caso real "
             "de cada cliente.",
    )
    alicuota = alicuota_pct / 100.0

    resumen = idb.calcular_resumen_patrimonial(cliente["id"], periodo, alicuota)

    if not resumen["tenencias"]:
        st.info("Sin tenencias abiertas al cierre para armar el resumen.")
        return

    sin_dato = [t for t in resumen["tenencias"] if t["fuente_valor_cierre"] == "sin_dato"]
    if sin_dato:
        st.warning(
            f"⚠️ {len(sin_dato)} instrumento(s) sin cotización de cierre encontrada "
            "(no está en las planillas oficiales 2025 cargadas, o el período no es "
            "2025) — completá el valor manual más abajo para que entren correctamente "
            "en Bienes Personales."
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Costo histórico (Ganancias)", _fmt_pesos(resumen["costo_historico_total_ars"], 0))
    c2.metric("Valor de cierre (Bienes Personales)", _fmt_pesos(resumen["valor_cierre_total_ars"], 0))
    c3.metric("TC cierre (31/12) usado", _fmt_cantidad(resumen["tc_cierre"]))
    c4.metric("Alícuota escala usada", f"{alicuota_pct:.0f}%")

    st.markdown("**Variación patrimonial del período**")
    vp = resumen["variacion_patrimonial"]
    v1, v2, v3, v4 = st.columns(4)
    v1.markdown(_html_metric("Rendimientos / intereses / rentas", vp["rendimientos_ars"]), unsafe_allow_html=True)
    v2.markdown(_html_metric("Compra-venta", vp["compraventa_ars"]), unsafe_allow_html=True)
    v3.markdown(_html_metric("Diferencia de cambio", vp["diferencia_cambio_ars"]), unsafe_allow_html=True)
    v4.markdown(_html_metric("Total variación patrimonial", vp["total_ars"]), unsafe_allow_html=True)

    st.markdown("**Composición de tenencias al cierre**")
    fuente_label = {"pdf": "Planilla oficial 2025", "manual": "Carga manual", "sin_dato": "⚠️ Sin dato"}
    filas_t = [{
        "Instrumento": t["instrumento"], "Tipo": idb.TIPO_LABEL.get(t["tipo"], t["tipo"]),
        "Cantidad": _fmt_cantidad(t["cantidad"], 4),
        "Costo histórico $ (Ganancias)": _fmt_pesos(t["costo_historico_ars"]),
        "Cotización cierre": f"{_fmt_cantidad(t['cotizacion_cierre'], 4)} {t['moneda_cotizacion']}",
        "Valor de cierre $ (Bienes Pers.)": _fmt_pesos(t["valor_cierre_total_ars"]),
        "Bienes Personales": idb.ALCANZA_BP_LABEL.get(t["alcanza_bienes_personales"], t["alcanza_bienes_personales"]),
        "Fuente cotización": fuente_label.get(t["fuente_valor_cierre"], t["fuente_valor_cierre"]),
    } for t in resumen["tenencias"]]
    st.dataframe(pd.DataFrame(filas_t), use_container_width=True, hide_index=True)

    with st.expander("Cargar/corregir valor de cierre manual de un instrumento"):
        st.caption(
            "Usalo para instrumentos sin dato (ON y FCI 2025 no se pudieron "
            "extraer automáticamente por un problema en esas dos planillas — "
            "ver aviso en el chat) o para cualquier período distinto de 2025."
        )
        instrumentos = [t["instrumento"] for t in resumen["tenencias"]]
        vversion = st.session_state.get("pat_vcierre_version", 0)
        i1, i2, i3, i4 = st.columns([2, 1, 1, 1])
        instr_sel = i1.selectbox("Instrumento", instrumentos, key=f"pat_vcierre_inst_{vversion}")
        cotiz = i2.number_input("Cotización", min_value=0.0, step=0.01, key=f"pat_vcierre_cot_{vversion}")
        moneda_v = i3.radio("Moneda", ["ARS", "USD"], horizontal=True, key=f"pat_vcierre_mon_{vversion}")
        if i4.button("Guardar", key=f"pat_vcierre_btn_{vversion}"):
            idb.guardar_valor_cierre_manual(cliente["id"], periodo, instr_sel, cotiz, moneda_v)
            st.session_state["pat_vcierre_version"] = vversion + 1
            st.success("Valor de cierre guardado ✓")
            st.rerun()

    st.markdown("**Desglose de impuesto aproximado por operación/instrumento**")
    st.caption(
        "Cedular 15%/5% y Retención 7% son el impuesto exacto (tasa fija sobre "
        "la base). 'Gravado a escala' usa la alícuota de referencia de arriba "
        "— es una aproximación a revisar, no el impuesto real del cliente."
    )
    filas_imp = [{
        "Fecha": d["fecha"], "Instrumento": d["instrumento"], "Movimiento": d["movimiento"],
        "Tratamiento": idb.FIS_LABEL.get(d["tratamiento_fiscal"], d["tratamiento_fiscal"]),
        "Base imponible $": _fmt_pesos(d["base_ars"]),
        "Alícuota": f"{d['alicuota'] * 100:.0f}%",
        "Impuesto aprox. $": _fmt_pesos(d["impuesto_ars"]),
    } for d in resumen["detalle_impuesto"]]
    st.dataframe(pd.DataFrame(filas_imp), use_container_width=True, hide_index=True)

    ti1, ti2 = st.columns(2)
    ti1.metric("Impuesto exacto (cedular/retención)", _fmt_pesos(resumen["impuesto_exacto_total_ars"]))
    ti2.metric("Impuesto aprox. (gravado a escala)", _fmt_pesos(resumen["impuesto_aproximado_escala_ars"]))

    st.divider()
    try:
        excel_bytes = _generar_excel_resumen(cliente, periodo, resumen)
    except Exception as e:
        st.error(f"No pude generar el Excel: {e}")
        return
    nombre_archivo = f"Resumen_Patrimonial_{cliente['nombre'].replace(' ', '_')}_{periodo}.xlsx"
    st.download_button(
        "📥 Exportar resumen a Excel",
        data=excel_bytes,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"pat_export_excel_{cliente['id']}_{periodo}",
    )


# ────────────────────────────────────────────────────────────────────────
# Exportación a Excel del resumen patrimonial (una sola solapa, con
# fórmulas que enlazan los totales a los datos puros — no hay ningún total
# hardcodeado como número fijo)
# ────────────────────────────────────────────────────────────────────────

_XLSX_AZUL = "1F4E79"
_XLSX_AZUL_CLARO = "DCE6F1"
_XLSX_GRIS = "F2F2F2"
_XLSX_VERDE = "1A7F37"
_XLSX_ROJO = "C62828"
_XLSX_BLANCO = "FFFFFF"


def _generar_excel_resumen(cliente: dict, periodo: str, resumen: dict) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Patrimonial"

    fill_header = PatternFill("solid", fgColor=_XLSX_AZUL)
    fill_subheader = PatternFill("solid", fgColor=_XLSX_AZUL_CLARO)
    fill_gris = PatternFill("solid", fgColor=_XLSX_GRIS)
    font_header = Font(color=_XLSX_BLANCO, bold=True, size=14)
    font_subheader = Font(color=_XLSX_AZUL, bold=True, size=11)
    font_col = Font(color=_XLSX_BLANCO, bold=True, size=10)
    font_normal = Font(size=10)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 20

    fila = 1

    def _titulo(texto, filas_alto=1):
        nonlocal fila
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila + filas_alto - 1, end_column=8)
        c = ws.cell(row=fila, column=1, value=texto)
        c.fill = fill_header
        c.font = font_header
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for r in range(fila, fila + filas_alto):
            for col in range(1, 9):
                ws.cell(row=r, column=col).fill = fill_header
        fila += filas_alto

    def _subtitulo(texto):
        nonlocal fila
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=8)
        c = ws.cell(row=fila, column=1, value=texto)
        c.font = font_subheader
        c.fill = fill_subheader
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        fila += 1

    def _fila_cabecera(cols):
        nonlocal fila
        for i, texto in enumerate(cols, start=1):
            c = ws.cell(row=fila, column=i, value=texto)
            c.font = font_col
            c.fill = fill_header
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        fila += 1

    _titulo(f"Resumen patrimonial — {cliente['nombre']} ({cliente['cuit']}) — Período {periodo}", filas_alto=2)
    fila += 1

    # ── Sección 1: Composición de tenencias al cierre ────────────────────
    _subtitulo("Composición de tenencias al cierre (a costo histórico y a valor de cierre)")
    fila_header_tenencias = fila
    _fila_cabecera([
        "Instrumento", "Tipo", "Cantidad", "Costo histórico $ (Ganancias)",
        "Cotización cierre", "Moneda cotiz.", "TC cierre (31/12)", "Valor de cierre $ (Bienes Pers.)",
    ])
    primera_fila_tenencias = fila
    tc_cierre_celda = None
    for t in resumen["tenencias"]:
        ws.cell(row=fila, column=1, value=t["instrumento"]).border = border
        ws.cell(row=fila, column=2, value=idb.TIPO_LABEL.get(t["tipo"], t["tipo"])).border = border
        ws.cell(row=fila, column=3, value=t["cantidad"]).border = border
        ws.cell(row=fila, column=3).number_format = "#,##0.0000"
        ws.cell(row=fila, column=4, value=t["costo_historico_ars"]).border = border
        ws.cell(row=fila, column=4).number_format = '"$" #,##0.00'
        ws.cell(row=fila, column=5, value=t["cotizacion_cierre"]).border = border
        ws.cell(row=fila, column=5).number_format = "#,##0.0000"
        ws.cell(row=fila, column=6, value=t["moneda_cotizacion"]).border = border
        ws.cell(row=fila, column=7, value=resumen["tc_cierre"]).border = border
        ws.cell(row=fila, column=7).number_format = "#,##0.00"
        if tc_cierre_celda is None:
            tc_cierre_celda = f"G{fila}"
        col_letra_cant, col_letra_cotiz, col_letra_mon, col_letra_tc = "C", "E", "F", "G"
        formula = (
            f"={col_letra_cant}{fila}*{col_letra_cotiz}{fila}"
            f"*IF({col_letra_mon}{fila}=\"USD\",{col_letra_tc}{fila},1)"
        )
        cvc = ws.cell(row=fila, column=8, value=formula)
        cvc.border = border
        cvc.number_format = '"$" #,##0.00'
        fila += 1
    ultima_fila_tenencias = fila - 1

    # Totales de la sección 1 (fórmulas SUM sobre el rango de arriba)
    ws.cell(row=fila, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=fila, column=4, value=f"=SUM(D{primera_fila_tenencias}:D{ultima_fila_tenencias})")
    ws.cell(row=fila, column=4).font = Font(bold=True)
    ws.cell(row=fila, column=4).number_format = '"$" #,##0.00'
    ws.cell(row=fila, column=8, value=f"=SUM(H{primera_fila_tenencias}:H{ultima_fila_tenencias})")
    ws.cell(row=fila, column=8).font = Font(bold=True)
    ws.cell(row=fila, column=8).number_format = '"$" #,##0.00'
    for col in range(1, 9):
        ws.cell(row=fila, column=col).fill = fill_gris
        ws.cell(row=fila, column=col).border = border
    fila_total_costo_hist = fila
    fila_total_valor_cierre = fila
    fila += 2

    # ── Sección 2: Detalle de variación patrimonial (por operación) ──────
    _subtitulo("Detalle de variación patrimonial (rendimientos / compra-venta / diferencia de cambio)")
    _fila_cabecera([
        "Fecha", "Instrumento", "Movimiento", "Rendimientos/intereses/rentas $",
        "Compra-venta $", "Diferencia de cambio $", "", "",
    ])
    primera_fila_var = fila
    for d in resumen["detalle_variacion"]:
        ws.cell(row=fila, column=1, value=str(d["fecha"])).border = border
        ws.cell(row=fila, column=2, value=d["instrumento"]).border = border
        ws.cell(row=fila, column=3, value=d["movimiento"]).border = border
        c4 = ws.cell(row=fila, column=4, value=d["rendimiento_intereses_rentas_ars"])
        c4.border = border
        c4.number_format = '"$" #,##0.00'
        c5 = ws.cell(row=fila, column=5, value=d["compraventa_ars"])
        c5.border = border
        c5.number_format = '"$" #,##0.00'
        c6 = ws.cell(row=fila, column=6, value=d["diferencia_cambio_ars"])
        c6.border = border
        c6.number_format = '"$" #,##0.00'
        fila += 1
    ultima_fila_var = fila - 1 if resumen["detalle_variacion"] else primera_fila_var

    fila_var_rend, fila_var_cv, fila_var_dif, fila_var_total = fila, fila + 1, fila + 2, fila + 3
    if resumen["detalle_variacion"]:
        rango_d = f"D{primera_fila_var}:D{ultima_fila_var}"
        rango_e = f"E{primera_fila_var}:E{ultima_fila_var}"
        rango_f = f"F{primera_fila_var}:F{ultima_fila_var}"
    else:
        rango_d = rango_e = rango_f = "D1:D1"  # sin movimientos: no hay rango, queda en 0

    def _fila_total_variacion(etiqueta, formula):
        nonlocal fila
        ws.cell(row=fila, column=1, value=etiqueta).font = Font(bold=True)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=3)
        c = ws.cell(row=fila, column=4, value=formula if resumen["detalle_variacion"] else 0)
        c.font = Font(bold=True)
        c.number_format = '"$" #,##0.00'
        for col in range(1, 5):
            ws.cell(row=fila, column=col).fill = fill_gris
            ws.cell(row=fila, column=col).border = border
        fila += 1
        return f"D{fila - 1}"

    ref_rend = _fila_total_variacion("Total Rendimientos/intereses/rentas", f"=SUM({rango_d})")
    ref_cv = _fila_total_variacion("Total Compra-venta", f"=SUM({rango_e})")
    ref_dif = _fila_total_variacion("Total Diferencia de cambio", f"=SUM({rango_f})")
    ws.cell(row=fila, column=1, value="TOTAL variación patrimonial").font = Font(bold=True)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=3)
    ctv = ws.cell(row=fila, column=4, value=f"={ref_rend}+{ref_cv}+{ref_dif}")
    ctv.font = Font(bold=True)
    ctv.number_format = '"$" #,##0.00'
    for col in range(1, 5):
        ws.cell(row=fila, column=col).fill = fill_subheader
        ws.cell(row=fila, column=col).border = border
    fila += 2

    # ── Sección 3: Desglose de impuesto por operación/instrumento ────────
    _subtitulo("Desglose de impuesto aproximado por operación/instrumento")
    fila_alicuota_ref = fila
    ws.cell(row=fila, column=1, value="Alícuota de referencia 'Gravado a escala' (editable):").font = Font(bold=True)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=3)
    calic = ws.cell(row=fila, column=4, value=resumen["alicuota_escala"])
    calic.number_format = "0%"
    calic.fill = PatternFill("solid", fgColor="FFF3CD")
    calic.font = Font(bold=True)
    celda_alicuota = f"D{fila}"
    fila += 2

    _fila_cabecera(["Fecha", "Instrumento", "Movimiento", "Tratamiento", "Base imponible $", "Alícuota", "Impuesto aprox. $", ""])
    primera_fila_imp = fila
    for d in resumen["detalle_impuesto"]:
        ws.cell(row=fila, column=1, value=str(d["fecha"])).border = border
        ws.cell(row=fila, column=2, value=d["instrumento"]).border = border
        ws.cell(row=fila, column=3, value=d["movimiento"]).border = border
        ws.cell(row=fila, column=4, value=idb.FIS_LABEL.get(d["tratamiento_fiscal"], d["tratamiento_fiscal"])).border = border
        cb = ws.cell(row=fila, column=5, value=d["base_ars"])
        cb.border = border
        cb.number_format = '"$" #,##0.00'
        es_escala = d["tratamiento_fiscal"] in ("gravado", "ordinario")
        ca = ws.cell(row=fila, column=6, value=(f"={celda_alicuota}" if es_escala else d["alicuota"]))
        ca.border = border
        ca.number_format = "0%"
        ci = ws.cell(row=fila, column=7, value=f"=E{fila}*F{fila}")
        ci.border = border
        ci.number_format = '"$" #,##0.00'
        fila += 1
    ultima_fila_imp = fila - 1 if resumen["detalle_impuesto"] else primera_fila_imp

    if resumen["detalle_impuesto"]:
        ws.cell(row=fila, column=1, value="TOTAL impuesto aproximado").font = Font(bold=True)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=6)
        ctot = ws.cell(row=fila, column=7, value=f"=SUM(G{primera_fila_imp}:G{ultima_fila_imp})")
        ctot.font = Font(bold=True)
        ctot.number_format = '"$" #,##0.00'
        for col in range(1, 8):
            ws.cell(row=fila, column=col).fill = fill_gris
            ws.cell(row=fila, column=col).border = border
        fila += 1

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _color_monto(valor: float) -> str:
    """Verde si es positivo, rojo si es negativo, gris si es cero — para
    remarcar de un vistazo si un rendimiento o diferencia de cambio favorece
    o no al cliente."""
    if valor > 1e-9:
        return "#1a7f37"
    if valor < -1e-9:
        return "#c62828"
    return "#57606a"


def _html_metric(label: str, valor: float, moneda: str = "$") -> str:
    """Igual que st.metric pero coloreando el número según el signo (verde/
    rojo) — st.metric no permite colorear el valor principal, solo el delta."""
    color = _color_monto(valor)
    monto = f"{moneda} {valor:,.0f}".replace(",", ".")
    return (
        '<div style="line-height:1.2; margin-bottom:0.5rem">'
        f'<div style="font-size:0.875rem;color:#57606a">{label}</div>'
        f'<div style="font-size:1.75rem;font-weight:600;color:{color}">{monto}</div>'
        '</div>'
    )


def _tab_movimientos(cliente: dict, periodo: str) -> None:
    st.caption(
        "Todas las operaciones del período, agrupadas por instrumento y en "
        "orden cronológico. Cada venta/rescate muestra el resultado realizado "
        "contra el costo PEPS consumido — misma lógica que la posición de "
        "Patrimonio (lo que no se vendió sigue como tenencia abierta ahí), para "
        "que ambas vistas coincidan siempre. Cuando la venta es en USD, ese "
        "resultado se abre además en Rendimiento (la ganancia/pérdida del "
        "instrumento — al TC del día de venta para la porción cuyo costo se "
        "originó en dólares, o directamente en pesos para la porción cuyo "
        "costo se originó en pesos) y Diferencia de cambio (solo para la "
        "porción con costo en dólares: lo que aporta nada más que el "
        "movimiento del tipo de cambio, no la actividad del cliente) — así se "
        "distingue lo que es variación patrimonial propia de lo que es "
        "variación cambiaria."
    )

    grupos = idb.calcular_movimientos(cliente["id"], periodo)
    if not grupos:
        st.info("Sin operaciones cargadas para este cliente y período.")
        return

    total_realizado = sum(
        m["resultado_ars"] for g in grupos for m in g["movimientos"]
        if m["resultado_ars"] is not None
    )
    total_rendimiento = sum(
        m["rendimiento_ars"] for g in grupos for m in g["movimientos"]
        if m["rendimiento_ars"] is not None
    )
    total_dif_cambio = sum(
        m["diferencia_cambio_ars"] for g in grupos for m in g["movimientos"]
        if m["diferencia_cambio_ars"] is not None
    )
    c1, c2, c3 = st.columns(3)
    c1.markdown(
        _html_metric("Resultado realizado del período (ventas/rescates)", total_realizado),
        unsafe_allow_html=True,
    )
    c2.markdown(_html_metric("de las cuales Rendimiento", total_rendimiento), unsafe_allow_html=True)
    c3.markdown(
        _html_metric("de las cuales Diferencia de cambio", total_dif_cambio),
        unsafe_allow_html=True,
    )
    st.divider()

    for g in grupos:
        tipo_label = idb.TIPO_LABEL.get(g["tipo"], g["tipo"])
        fis_label = idb.FIS_LABEL.get(g["tratamiento_fiscal"], g["tratamiento_fiscal"])
        n_mov = len(g["movimientos"])
        with st.expander(f"{g['instrumento']} — {tipo_label} — {fis_label} ({n_mov} movimiento{'s' if n_mov != 1 else ''})"):
            filas = [{
                "Fecha": m["fecha"],
                "Movimiento": idb.MOV_LABEL.get(m["movimiento"], m["movimiento"]),
                "VN": _fmt_cantidad(m["vn"]),
                "Moneda": m["moneda"],
                "Total ARS": _fmt_pesos(m["total_ars"]),
                "Total USD": _fmt_usd(m["total_usd"]),
                "Div./Renta $": _fmt_pesos(m["div_ars"]),
                "Comisión": _fmt_moneda(m["comision"], m["moneda"]),
                "Resultado realizado $": (_fmt_pesos(m["resultado_ars"]) if m["resultado_ars"] is not None else ""),
                "Rendimiento $": (_fmt_pesos(m["rendimiento_ars"]) if m["rendimiento_ars"] is not None else ""),
                "Dif. de cambio $": (_fmt_pesos(m["diferencia_cambio_ars"]) if m["diferencia_cambio_ars"] is not None else ""),
            } for m in g["movimientos"]]
            st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

            existencia = sum(
                float(l["cant"]) for l in g.get("lotes", []) if float(l["cant"]) > 1e-9
            )
            st.markdown(f"**Existencia al cierre:** {_fmt_cantidad(existencia)}")

            rend_inst = sum(
                m["rendimiento_ars"] for m in g["movimientos"] if m["rendimiento_ars"] is not None
            )
            dif_inst = sum(
                m["diferencia_cambio_ars"] for m in g["movimientos"] if m["diferencia_cambio_ars"] is not None
            )
            tiene_usd = any(m["rendimiento_ars"] is not None for m in g["movimientos"])
            if tiene_usd:
                ci1, ci2 = st.columns(2)
                ci1.markdown(
                    _html_metric(f"Rendimiento total — {g['instrumento']}", rend_inst),
                    unsafe_allow_html=True,
                )
                ci2.markdown(
                    _html_metric(f"Diferencia de cambio total — {g['instrumento']}", dif_inst),
                    unsafe_allow_html=True,
                )
