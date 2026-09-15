"""Persistencia de la herramienta Inversiones (InversionARG) — Etapa 1.

A diferencia de la versión standalone (InversionARG_Final.html + localStorage,
cada usuario/navegador con sus propios datos), esta capa guarda todo en la
MISMA base SQLite que el resto de Estudio Contable, y reutiliza los clientes
ya cargados (``database.listar_clientes()``): no hay alta de cliente propia
para Inversiones.

``inicializar_tablas_inversiones`` se llama desde ``database.inicializar_bd()``
(una línea agregada ahí), igual que ``_inicializar_tablas_sueldos`` o
``_inicializar_tablas_conciliacion``.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Optional
import database as db

BASE_DIR = Path(__file__).resolve().parent
TC_BNA_SEED_PATH = BASE_DIR / "data" / "tc_bna_seed.json"
INSTRUMENTOS_SEED_PATH = BASE_DIR / "data" / "instrumentos_seed.json"
VALUACIONES_CIERRE_SEED_PATH = BASE_DIR / "data" / "valuaciones_cierre_2025.json"

# Alícuota de referencia por defecto para el tratamiento "Gravado a escala"
# (impuesto a las Ganancias, escalas progresivas de persona humana): no hay
# una tasa fija posible acá porque depende de TODOS los ingresos anuales del
# cliente, no solo de Inversiones — se usa el tramo máximo vigente (35%) como
# punto de partida razonable, siempre editable a mano (tanto en la app como
# en el Excel exportado) para ajustarla al caso real de cada cliente.
ALICUOTA_ESCALA_DEFAULT = 0.35

# --- Catálogos (mismos valores que la herramienta original, + fideicomiso/opcion) ---

TIPOS_INSTRUMENTO = (
    "on", "on_exterior", "bono", "titulo_extranjero", "accion", "accion_sin_cotizacion",
    "accion_exterior", "adr", "cedear", "crypto", "letra", "fci", "fci_sin_oferta",
    "fci_extranjero", "caucion", "fideicomiso", "opcion", "plazo_fijo_ars", "plazo_fijo_usd",
    "futuro", "cheque_dif", "fce", "otro",
)
MOVIMIENTOS = ("compra", "venta", "dividendo", "renta", "amort", "apertura", "rescate", "caucion")
TRATAMIENTOS_FISCALES = ("ced15", "ced5", "ret7", "exento", "gravado", "ordinario", "na")
ALCANZA_BP = ("exento", "gravado")
ORIGENES_POOL = ("mep", "oficial", "blue", "ext", "otro")

TIPO_LABEL = {
    "on": "Obligación Negociable (local)", "on_exterior": "Obligación Negociable (exterior)",
    "bono": "Bono / Título público (argentino)", "titulo_extranjero": "Título público extranjero",
    "accion": "Acción (con cotización)", "accion_sin_cotizacion": "Acción (sin cotización)",
    "accion_exterior": "Acción (exterior)", "adr": "ADR de sociedad argentina",
    "cedear": "CEDEAR", "crypto": "Cripto", "letra": "Letra del Tesoro / LECAP",
    "fci": "FCI/Cuotapartes (oferta pública)", "fci_sin_oferta": "FCI (sin oferta pública)",
    "fci_extranjero": "FCI extranjero", "caucion": "Caución",
    "fideicomiso": "Fideicomiso financiero (oferta pública)", "opcion": "Opción / Warrant",
    "plazo_fijo_ars": "Plazo fijo en pesos", "plazo_fijo_usd": "Plazo fijo en dólares",
    "futuro": "Futuro", "cheque_dif": "Cheque de pago diferido",
    "fce": "FCE / Pagaré bursátil", "otro": "Otro",
}
MOV_LABEL = {
    "compra": "Compra", "venta": "Venta", "dividendo": "Dividendo",
    "renta": "Renta", "amort": "Amortización", "apertura": "Suscripción/Alta",
    "rescate": "Rescate", "caucion": "Caución",
}
FIS_LABEL = {
    "ced15": "Cedular 15%", "ced5": "Cedular 5%", "ret7": "Retención 7% (dividendos)",
    "exento": "Exento", "gravado": "Gravado a escala", "ordinario": "Renta ordinaria",
    "na": "N/A",
}
ALCANZA_BP_LABEL = {"exento": "Exento", "gravado": "Gravado"}
ORIGEN_POOL_LABEL = {"mep": "MEP", "oficial": "Oficial", "blue": "Informal", "ext": "Remesa", "otro": "Otro"}

# Movimientos que representan un "rendimiento" (cobro de renta) vs. una
# operación de compra/venta/rescate de capital — la carga tributaria de
# Ganancias suele diferir entre ambos casos (ver CATALOGO_FISCAL).
MOVIMIENTOS_RENDIMIENTO = ("dividendo", "renta")

# Movimientos que fijan un costo base en USD (compra de capital): son los
# únicos donde el TC "de origen" (al que se consiguieron los dólares) es un
# dato a cargar — se usa para el costo fiscal PEPS del pool USD. En el resto
# de los movimientos (venta, rescate, dividendo, renta, amort, caución) el
# TC que corresponde es el de la BNA del día de la operación, que ya se saca
# solo de la base (ver ``tc_para_movimiento``) — no hace falta cargarlo.
MOVIMIENTOS_COSTO_BASE = ("compra", "apertura")

# Catálogo de tratamiento impositivo por tipo de instrumento, según el cuadro
# "Impuesto a las Ganancias (Rendimiento / Compra-Venta) + Bienes Personales"
# provisto por el estudio (cuadro completo, fila por fila):
#
#   Inversión                                  | Rendimiento | Compra/Venta | Bienes Personales
#   Acciones arg. con cotización CNV           | Ret. 7% div | Exento        | Exento
#   Acciones arg. sin cotización                | Ret. 7% div | Gravado 15%   | Exento
#   Acciones en el exterior                     | Grav. escala| Gravado 15%   | Gravado
#   ADRs de sociedades argentinas                | Ret. 7% div | Gravado 15%   | Exento *
#   CEDEARs                                       | Grav. escala| Exento        | Gravado
#   ONs locales en pesos                          | Exento      | Exento        | Exento
#   ONs locales en dólares                        | Exento      | Exento        | Gravado
#   ONs en el exterior                            | Grav. escala| Gravado 15%   | Gravado
#   Bonos argentinos (pesos y dólares)            | Exento      | Exento        | Exento
#   Títulos públicos extranjeros                  | Grav. escala| Gravado 15%   | Gravado
#   FCI (oferta pública, activo subyacente 75%+   | Exento      | Exento        | Exento **
#     en tít. públicos/plazo fijo/ON oferta púb.)
#   FCI sin oferta pública                        | Grav. escala| Gravado 15%   | Gravado
#   FCI extranjeros                               | Grav. escala| Gravado 15%   | Gravado
#   Fideicomisos financieros oferta pública        | Exento      | Exento        | Exento **
#   Plazo fijo en pesos (con o sin cláusula ajuste)| Exento      | —             | Exento
#   Plazo fijo en dólares                          | Grav. escala| —             | Exento
#   Cripto                                        | —           | Gravado 15%   | Gravado
#   Futuros                                        | —           | Grav. escala  | Gravado
#   Opciones                                       | —           | Grav. escala  | Gravado
#   FCE y pagaré bursátil                          | Exento      | Exento        | Gravado
#   Cheque de pago diferido                       | Grav. escala| Exento        | Gravado
#   Cauciones                                      | Grav. escala| —             | Gravado
#
#   * Hay criterios que sostienen que no está gravado, al tributar la
#     sociedad como responsable sustituto por las acciones.
#   ** Condicionado a que esté colocado por oferta pública (CNV) y el activo
#      subyacente principal esté integrado en un 75% por títulos públicos,
#      depósitos a plazo fijo, ONs colocadas por oferta pública e
#      instrumentos emitidos en moneda nacional (ver listado de ARCA).
#
# Cajas de ahorro/cuentas corrientes/cuentas en el exterior/tenencia de
# moneda extranjera no tienen "operaciones" propias en esta herramienta (son
# saldos, no instrumentos con compra/venta) y quedan fuera de este catálogo.
#
# Es un PUNTO DE PARTIDA sugerido — el contador puede (y en los casos con *
# o **, DEBE) corregirlo a mano según el caso particular.
#   rendimiento: tratamiento para dividendo/renta
#   compraventa: tratamiento para compra/venta/rescate/amort/caución/apertura
#   bienes_personales: si la tenencia al 31/12 está alcanzada por BP
CATALOGO_FISCAL: dict[str, dict[str, str]] = {
    "accion":                {"rendimiento": "ret7",    "compraventa": "exento", "bienes_personales": "exento"},
    "accion_sin_cotizacion": {"rendimiento": "ret7",    "compraventa": "ced15",  "bienes_personales": "exento"},
    "accion_exterior":       {"rendimiento": "gravado", "compraventa": "ced15",  "bienes_personales": "gravado"},
    "adr":                   {"rendimiento": "ret7",    "compraventa": "ced15",  "bienes_personales": "exento"},
    "cedear":                {"rendimiento": "gravado", "compraventa": "exento", "bienes_personales": "gravado"},
    "on":                    {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "exento"},
    "on_exterior":           {"rendimiento": "gravado", "compraventa": "ced15",  "bienes_personales": "gravado"},
    "bono":                  {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "exento"},
    "titulo_extranjero":     {"rendimiento": "gravado", "compraventa": "ced15",  "bienes_personales": "gravado"},
    "letra":                 {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "exento"},
    "fci":                   {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "exento"},
    "fci_sin_oferta":        {"rendimiento": "gravado", "compraventa": "ced15",  "bienes_personales": "gravado"},
    "fci_extranjero":        {"rendimiento": "gravado", "compraventa": "ced15",  "bienes_personales": "gravado"},
    "fideicomiso":           {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "exento"},
    "opcion":                {"rendimiento": "na",      "compraventa": "gravado","bienes_personales": "gravado"},
    "plazo_fijo_ars":        {"rendimiento": "exento",  "compraventa": "na",     "bienes_personales": "exento"},
    "plazo_fijo_usd":        {"rendimiento": "gravado", "compraventa": "na",     "bienes_personales": "exento"},
    "crypto":                {"rendimiento": "na",      "compraventa": "ced15",  "bienes_personales": "gravado"},
    "futuro":                {"rendimiento": "na",      "compraventa": "gravado","bienes_personales": "gravado"},
    "fce":                   {"rendimiento": "exento",  "compraventa": "exento", "bienes_personales": "gravado"},
    "cheque_dif":            {"rendimiento": "gravado", "compraventa": "exento", "bienes_personales": "gravado"},
    "caucion":               {"rendimiento": "gravado", "compraventa": "na",     "bienes_personales": "gravado"},
    "otro":                  {"rendimiento": "gravado", "compraventa": "gravado","bienes_personales": "gravado"},
}

# Casos donde la moneda cambia el resultado (según el cuadro: ON local en
# USD queda gravada en Bienes Personales, a diferencia de la emitida en ARS).
_CATALOGO_FISCAL_ON_USD_BP = "gravado"


def sugerir_fiscal(tipo: str, movimiento: str, moneda: str = "ARS") -> dict[str, str]:
    """Sugiere (tratamiento_fiscal, alcanza_bienes_personales) para una operación.

    Es una sugerencia editable, no una determinación definitiva: casos como
    ONs/FCI "sin oferta pública", acciones sin cotización, ADRs o instrumentos
    en el exterior tienen reglas distintas a las del caso general acá cubierto.
    """
    cat = CATALOGO_FISCAL.get(tipo, CATALOGO_FISCAL["otro"])
    es_rendimiento = movimiento in MOVIMIENTOS_RENDIMIENTO
    tratamiento = cat["rendimiento"] if es_rendimiento else cat["compraventa"]
    bp = cat["bienes_personales"]
    if tipo == "on" and moneda == "USD":
        bp = _CATALOGO_FISCAL_ON_USD_BP
    return {"tratamiento_fiscal": tratamiento, "alcanza_bienes_personales": bp}


def default_fis(tipo: str) -> str:
    """Compat: tratamiento fiscal sugerido por defecto para compra/venta según tipo."""
    return CATALOGO_FISCAL.get(tipo, CATALOGO_FISCAL["otro"])["compraventa"]


_catalogo_instrumentos_cache: Optional[list[dict]] = None


_valuaciones_cierre_cache: Optional[dict[str, dict]] = None


def listar_valuaciones_cierre() -> dict[str, dict]:
    """Catálogo de cotizaciones de cierre 2025 (Bienes Personales), indexado
    por ticker en mayúsculas — extraído de las planillas oficiales de
    valuaciones (Acciones, CEDEARs, Fideicomisos Financieros, Opciones,
    Títulos Públicos). Dos de las siete planillas provistas (Obligaciones
    Negociables y Fondos Comunes de Inversión) tienen un defecto de
    generación del PDF que superpone dos capas de texto y hace que la
    extracción automática de la columna Cotización dé valores ilegibles —
    esos instrumentos no están en este catálogo y necesitan carga manual
    (ver ``obtener_valor_cierre`` / ``guardar_valor_cierre_manual``).

    El archivo semilla guarda cada fila en formato compacto
    ``[ticker, cotizacion, moneda]`` (no un objeto por fila) para que el
    JSON pese lo menos posible — acá se expande a un dict por comodidad de
    uso en el resto del módulo."""
    global _valuaciones_cierre_cache
    if _valuaciones_cierre_cache is not None:
        return _valuaciones_cierre_cache
    if not VALUACIONES_CIERRE_SEED_PATH.is_file():
        _valuaciones_cierre_cache = {}
        return _valuaciones_cierre_cache
    try:
        filas = json.loads(VALUACIONES_CIERRE_SEED_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        _valuaciones_cierre_cache = {}
        return _valuaciones_cierre_cache
    _valuaciones_cierre_cache = {
        fila[0]: {"cotizacion": fila[1], "moneda": fila[2]}
        for fila in filas if fila and fila[0]
    }
    return _valuaciones_cierre_cache


def _ticker_de_instrumento(instrumento: str) -> str:
    """El formulario de carga guarda el instrumento como "TICKER - Nombre" (o
    solo el nombre libre si se cargó a mano) — esto recupera la parte de
    ticker para buscarla en el catálogo de cierre."""
    return (instrumento or "").split(" - ")[0].strip().upper()


def obtener_valor_cierre(cliente_id: int, periodo: str, instrumento: str) -> dict:
    """Cotización de cierre del período para Bienes Personales, con la
    siguiente prioridad: 1) valor cargado a mano para ese cliente/período/
    instrumento (siempre gana, es la corrección del contador); 2) catálogo
    de las planillas oficiales 2025 por ticker; 3) sin dato (0, a completar
    a mano). Devuelve también ``fuente`` para que la UI/Excel puedan avisar
    de dónde salió el número."""

    with db.obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT cotizacion, moneda FROM inversiones_valor_cierre_manual "
            "WHERE cliente_id = ? AND periodo = ? AND instrumento = ?",
            (cliente_id, periodo, instrumento),
        ).fetchone()
    if fila:
        return {"cotizacion": float(fila["cotizacion"]), "moneda": fila["moneda"], "fuente": "manual"}

    if periodo == "2025":
        catalogo = listar_valuaciones_cierre()
        entrada = catalogo.get(_ticker_de_instrumento(instrumento))
        if entrada:
            return {"cotizacion": float(entrada["cotizacion"]), "moneda": entrada["moneda"], "fuente": "pdf"}

    return {"cotizacion": 0.0, "moneda": "ARS", "fuente": "sin_dato"}


def guardar_valor_cierre_manual(
    cliente_id: int, periodo: str, instrumento: str, cotizacion: float, moneda: str = "ARS",
) -> None:
    """Guarda (o corrige) a mano la cotización de cierre de un instrumento
    para Bienes Personales — para períodos sin planilla oficial cargada
    (2026 en adelante) o para pisar/completar un dato del catálogo 2025."""

    with db.obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO inversiones_valor_cierre_manual
                (cliente_id, periodo, instrumento, cotizacion, moneda)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(cliente_id, periodo, instrumento)
            DO UPDATE SET cotizacion = excluded.cotizacion, moneda = excluded.moneda
            """,
            (cliente_id, periodo, instrumento.strip(), float(cotizacion), moneda),
        )
        conn.commit()


def listar_catalogo_instrumentos() -> list[dict]:
    """Catálogo de instrumentos conocidos (ticker, nombre, tipo) para autocompletar
    el formulario de carga — extraído de las planillas de valuaciones 2025
    (Acciones/CEDEARs/ONs/Títulos públicos/Fideicomisos/Opciones/FCI)."""
    global _catalogo_instrumentos_cache
    if _catalogo_instrumentos_cache is not None:
        return _catalogo_instrumentos_cache
    if not INSTRUMENTOS_SEED_PATH.is_file():
        _catalogo_instrumentos_cache = []
        return _catalogo_instrumentos_cache
    try:
        _catalogo_instrumentos_cache = json.loads(INSTRUMENTOS_SEED_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        _catalogo_instrumentos_cache = []
    return _catalogo_instrumentos_cache


# --- Esquema ----------------------------------------------------------------

def inicializar_tablas_inversiones(conn: sqlite3.Connection) -> None:
    """Crea las tablas de Inversiones si no existen. Llamado desde database.inicializar_bd()."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS inversiones_operaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            fecha TEXT NOT NULL,
            instrumento TEXT NOT NULL,
            tipo TEXT NOT NULL,
            movimiento TEXT NOT NULL,
            tratamiento_fiscal TEXT NOT NULL DEFAULT 'gravado',
            vn REAL NOT NULL DEFAULT 0,
            precio_ars REAL NOT NULL DEFAULT 0,
            precio_usd REAL NOT NULL DEFAULT 0,
            tc REAL NOT NULL DEFAULT 0,
            tc_origen REAL NOT NULL DEFAULT 0,
            total_ars REAL NOT NULL DEFAULT 0,
            total_usd REAL NOT NULL DEFAULT 0,
            div_usd REAL NOT NULL DEFAULT 0,
            div_ars REAL NOT NULL DEFAULT 0,
            comision REAL NOT NULL DEFAULT 0,
            moneda TEXT NOT NULL DEFAULT 'ARS',
            notas TEXT NOT NULL DEFAULT '',
            costo_fiscal REAL NOT NULL DEFAULT 0,
            asignaciones_usd_json TEXT NOT NULL DEFAULT '[]',
            es_saldo_inicial INTEGER NOT NULL DEFAULT 0,
            alcanza_bienes_personales TEXT NOT NULL DEFAULT 'gravado',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    try:
        conn.execute(
            "ALTER TABLE inversiones_operaciones "
            "ADD COLUMN alcanza_bienes_personales TEXT NOT NULL DEFAULT 'gravado'"
        )
    except sqlite3.OperationalError:
        pass  # columna ya existe (BD creada antes de este agregado)
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_inv_ops_cliente_fecha "
        "ON inversiones_operaciones(cliente_id, fecha)"
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS inversiones_pool_usd (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            fecha TEXT NOT NULL,
            cant_usd REAL NOT NULL DEFAULT 0,
            tc_origen REAL NOT NULL DEFAULT 0,
            disponible_usd REAL NOT NULL DEFAULT 0,
            origen TEXT NOT NULL DEFAULT 'otro',
            descripcion TEXT NOT NULL DEFAULT '',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS inversiones_tc_bna (
            fecha TEXT PRIMARY KEY,
            compra REAL NOT NULL,
            venta REAL NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS inversiones_tenencias_control (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            periodo TEXT NOT NULL,
            instrumento TEXT NOT NULL,
            vn_control REAL NOT NULL DEFAULT 0,
            valor_ars_control REAL NOT NULL DEFAULT 0,
            notas TEXT NOT NULL DEFAULT '',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS inversiones_valor_cierre_manual (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            periodo TEXT NOT NULL,
            instrumento TEXT NOT NULL,
            cotizacion REAL NOT NULL DEFAULT 0,
            moneda TEXT NOT NULL DEFAULT 'ARS',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cliente_id, periodo, instrumento)
        )
        """
    )


def sembrar_tc_bna_default() -> None:
    """Carga el historial BNA 2023-2025 una sola vez (no-op si la tabla ya tiene datos)."""

    if not TC_BNA_SEED_PATH.is_file():
        return
    with db.obtener_conexion() as conn:
        n = conn.execute("SELECT COUNT(*) AS n FROM inversiones_tc_bna").fetchone()["n"]
        if n and int(n) > 0:
            return
        try:
            filas = json.loads(TC_BNA_SEED_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        for fecha, compra, venta in filas:
            conn.execute(
                "INSERT OR IGNORE INTO inversiones_tc_bna (fecha, compra, venta) VALUES (?, ?, ?)",
                (fecha, float(compra), float(venta)),
            )
        conn.commit()


# --- TC BNA ------------------------------------------------------------------

def _obtener_tc_con(conn: sqlite3.Connection, fecha: str, tipo: str = "venta") -> Optional[float]:
    """Igual que ``obtener_tc`` pero reutilizando una conexión ya abierta —
    para no anidar conexiones dentro de una transacción en curso (usado por
    ``asignar_peps_automatico``)."""
    col = "compra" if tipo == "compra" else "venta"
    d = (fecha or "")[:10]
    if not d:
        return None
    fila = conn.execute(
        f"SELECT {col} AS v FROM inversiones_tc_bna WHERE fecha = ?", (d,)
    ).fetchone()
    if fila:
        return float(fila["v"])

    anterior = conn.execute(
        f"SELECT {col} AS v FROM inversiones_tc_bna WHERE fecha <= ? "
        "ORDER BY fecha DESC LIMIT 1",
        (d,),
    ).fetchone()
    siguiente = conn.execute(
        f"SELECT {col} AS v FROM inversiones_tc_bna WHERE fecha > ? "
        "ORDER BY fecha ASC LIMIT 1",
        (d,),
    ).fetchone()

    if tipo == "venta":
        # TC vendedor (compra/venta/amort/apertura/rescate/caución): si la
        # fecha cae en fin de semana o feriado, se toma el día HÁBIL
        # SIGUIENTE. Solo se usa el día anterior como respaldo si la fecha
        # es posterior a toda la serie cargada (no hay "siguiente" posible).
        if siguiente:
            return float(siguiente["v"])
        if anterior:
            return float(anterior["v"])
        return None

    # TC comprador (dividendo/renta): se mantiene la convención original —
    # día hábil ANTERIOR primero.
    if anterior:
        return float(anterior["v"])
    if siguiente:
        return float(siguiente["v"])
    return None


def obtener_tc(fecha: str, tipo: str = "venta") -> Optional[float]:
    """TC BNA para una fecha (columna 'compra' o 'venta').

    Coincidencia exacta si existe. Si no (fin de semana/feriado): para TC
    vendedor (compras/ventas de instrumentos, amortizaciones, etc.) se toma
    el día hábil SIGUIENTE; para TC comprador (dividendos/rentas) se toma el
    día hábil ANTERIOR — misma convención que la herramienta original para
    ese caso. En ambos casos, si no hay dato hacia ese lado (fecha en un
    extremo de la serie cargada), se usa el lado disponible.
    """

    with db.obtener_conexion() as conn:
        return _obtener_tc_con(conn, fecha, tipo)


def tc_para_movimiento(fecha: str, movimiento: str) -> Optional[float]:
    """TC a aplicar según el tipo de movimiento (dividendo/renta -> comprador; resto -> vendedor)."""
    tipo = "compra" if movimiento in ("dividendo", "renta") else "venta"
    return obtener_tc(fecha, tipo)


def listar_tc_bna(limite: int = 400) -> list[dict]:

    with db.obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT * FROM inversiones_tc_bna ORDER BY fecha DESC LIMIT ?", (limite,)
        ).fetchall()
    return [dict(f) for f in filas]


def agregar_tc_bna(fecha: str, compra: float, venta: float) -> None:

    with db.obtener_conexion() as conn:
        conn.execute(
            "INSERT INTO inversiones_tc_bna (fecha, compra, venta) VALUES (?, ?, ?) "
            "ON CONFLICT(fecha) DO UPDATE SET compra = excluded.compra, venta = excluded.venta",
            (fecha, float(compra), float(venta)),
        )
        conn.commit()


# --- Operaciones ---------------------------------------------------------

def crear_operacion(
    cliente_id: int,
    fecha: str,
    instrumento: str,
    tipo: str,
    movimiento: str,
    tratamiento_fiscal: str,
    vn: float = 0,
    precio_ars: float = 0,
    precio_usd: float = 0,
    tc: float = 0,
    tc_origen: float = 0,
    total_ars: float = 0,
    total_usd: float = 0,
    div_usd: float = 0,
    div_ars: float = 0,
    comision: float = 0,
    moneda: str = "ARS",
    notas: str = "",
    costo_fiscal: float = 0,
    asignaciones_usd: Optional[list] = None,
    es_saldo_inicial: bool = False,
    alcanza_bienes_personales: str = "gravado",
) -> int:

    with db.obtener_conexion() as conn:
        cur = conn.execute(
            """
            INSERT INTO inversiones_operaciones (
                cliente_id, fecha, instrumento, tipo, movimiento, tratamiento_fiscal,
                vn, precio_ars, precio_usd, tc, tc_origen, total_ars, total_usd,
                div_usd, div_ars, comision, moneda, notas, costo_fiscal,
                asignaciones_usd_json, es_saldo_inicial, alcanza_bienes_personales
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                cliente_id, fecha, instrumento.strip(), tipo, movimiento, tratamiento_fiscal,
                float(vn), float(precio_ars), float(precio_usd), float(tc), float(tc_origen),
                float(total_ars), float(total_usd), float(div_usd), float(div_ars),
                float(comision), moneda, notas.strip(), float(costo_fiscal),
                json.dumps(asignaciones_usd or [], ensure_ascii=False),
                1 if es_saldo_inicial else 0, alcanza_bienes_personales,
            ),
        )
        conn.commit()
        return int(cur.lastrowid)


def actualizar_operacion(
    operacion_id: int,
    fecha: str,
    instrumento: str,
    tipo: str,
    movimiento: str,
    tratamiento_fiscal: str,
    vn: float,
    precio_ars: float,
    precio_usd: float,
    tc: float,
    tc_origen: float,
    total_ars: float,
    total_usd: float,
    div_usd: float,
    div_ars: float,
    comision: float,
    moneda: str,
    notas: str,
    costo_fiscal: float,
    alcanza_bienes_personales: str,
) -> None:
    """Actualiza una operación ya cargada (edición manual desde la tabla de
    Operaciones cargadas). No toca ``asignaciones_usd_json``: si la edición
    cambia sustancialmente el total en USD, puede quedar desalineada con lo
    ya asignado en Depuración — se ve reflejado ahí como "sin completar" y
    se puede volver a correr "Asignar PEPS automático" o reasignar a mano."""

    with db.obtener_conexion() as conn:
        conn.execute(
            """
            UPDATE inversiones_operaciones SET
                fecha = ?, instrumento = ?, tipo = ?, movimiento = ?, tratamiento_fiscal = ?,
                vn = ?, precio_ars = ?, precio_usd = ?, tc = ?, tc_origen = ?,
                total_ars = ?, total_usd = ?, div_usd = ?, div_ars = ?, comision = ?,
                moneda = ?, notas = ?, costo_fiscal = ?, alcanza_bienes_personales = ?
            WHERE id = ?
            """,
            (
                fecha, instrumento.strip(), tipo, movimiento, tratamiento_fiscal,
                float(vn), float(precio_ars), float(precio_usd), float(tc), float(tc_origen),
                float(total_ars), float(total_usd), float(div_usd), float(div_ars),
                float(comision), moneda, notas.strip(), float(costo_fiscal),
                alcanza_bienes_personales, operacion_id,
            ),
        )
        conn.commit()


def listar_operaciones(
    cliente_id: Optional[int] = None,
    periodo: Optional[str] = None,
    solo_saldos_iniciales: Optional[bool] = None,
) -> list[dict]:

    q = "SELECT * FROM inversiones_operaciones WHERE 1=1"
    args: list = []
    if cliente_id:
        q += " AND cliente_id = ?"
        args.append(cliente_id)
    if periodo:
        q += " AND fecha LIKE ?"
        args.append(f"{periodo}%")
    if solo_saldos_iniciales is True:
        q += " AND es_saldo_inicial = 1"
    elif solo_saldos_iniciales is False:
        q += " AND es_saldo_inicial = 0"
    q += " ORDER BY fecha ASC, id ASC"
    with db.obtener_conexion() as conn:
        filas = conn.execute(q, args).fetchall()
    out = []
    for f in filas:
        d = dict(f)
        try:
            d["asignaciones_usd"] = json.loads(d.get("asignaciones_usd_json") or "[]")
        except json.JSONDecodeError:
            d["asignaciones_usd"] = []
        out.append(d)
    return out


def eliminar_operacion(operacion_id: int) -> None:

    with db.obtener_conexion() as conn:
        conn.execute("DELETE FROM inversiones_operaciones WHERE id = ?", (operacion_id,))
        conn.commit()


def obtener_operacion(operacion_id: int) -> Optional[dict]:

    with db.obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT * FROM inversiones_operaciones WHERE id = ?", (operacion_id,)
        ).fetchone()
    if not fila:
        return None
    d = dict(fila)
    try:
        d["asignaciones_usd"] = json.loads(d.get("asignaciones_usd_json") or "[]")
    except json.JSONDecodeError:
        d["asignaciones_usd"] = []
    return d


def cambiar_tratamiento_fiscal_instrumento(
    cliente_id: int, instrumento: str, periodo: str, nuevo_tratamiento: str
) -> int:
    """Aplica un tratamiento fiscal a TODAS las operaciones de ese instrumento
    (mismo cliente y período) — atajo para corregir de una vez casos
    particulares (ON/FCI sin oferta pública, ADRs, etc.)."""

    with db.obtener_conexion() as conn:
        cur = conn.execute(
            "UPDATE inversiones_operaciones SET tratamiento_fiscal = ? "
            "WHERE cliente_id = ? AND instrumento = ? AND fecha LIKE ?",
            (nuevo_tratamiento, cliente_id, instrumento, f"{periodo}%"),
        )
        conn.commit()
        return cur.rowcount


def cambiar_moneda_operacion(operacion_id: int, nueva_moneda: str) -> None:
    """Corrige la moneda de una operación ya cargada (p. ej. se cargó en ARS
    una compra que en realidad fue en USD, o viceversa). Recalcula total_ars/
    total_usd/costo_fiscal y, si pasa a ARS, devuelve al pool cualquier lote
    USD que tuviera asignado."""

    with db.obtener_conexion() as conn:
        op = conn.execute(
            "SELECT * FROM inversiones_operaciones WHERE id = ?", (operacion_id,)
        ).fetchone()
        if not op:
            return
        tc = float(op["tc"]) or 1.0
        if nueva_moneda == "ARS":
            asignaciones = json.loads(op["asignaciones_usd_json"] or "[]")
            for a in asignaciones:
                lote_id = a.get("lote_id")
                if lote_id:
                    conn.execute(
                        "UPDATE inversiones_pool_usd SET disponible_usd = "
                        "MIN(cant_usd, disponible_usd + ?) WHERE id = ?",
                        (a["cant_usd"], lote_id),
                    )
            total_ars = float(op["total_ars"])
            total_usd = total_ars / tc if tc else 0.0
            conn.execute(
                "UPDATE inversiones_operaciones SET moneda = 'ARS', total_usd = ?, "
                "costo_fiscal = ?, asignaciones_usd_json = '[]' WHERE id = ?",
                (total_usd, total_ars, operacion_id),
            )
        else:
            total_usd = float(op["total_usd"])
            total_ars = total_usd * tc
            conn.execute(
                "UPDATE inversiones_operaciones SET moneda = 'USD', total_ars = ?, "
                "costo_fiscal = ? WHERE id = ?",
                (total_ars, total_ars, operacion_id),
            )
        conn.commit()


# --- Pool USD (Etapa 2 — Depuración) ----------------------------------------

def listar_pool_usd(cliente_id: int) -> list[dict]:

    with db.obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT * FROM inversiones_pool_usd WHERE cliente_id = ? ORDER BY fecha ASC",
            (cliente_id,),
        ).fetchall()
    return [dict(f) for f in filas]


def crear_lote_pool(
    cliente_id: int, fecha: str, cant_usd: float, tc_origen: float,
    origen: str = "otro", descripcion: str = "",
) -> int:

    with db.obtener_conexion() as conn:
        cur = conn.execute(
            """
            INSERT INTO inversiones_pool_usd (
                cliente_id, fecha, cant_usd, tc_origen, disponible_usd, origen, descripcion
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (cliente_id, fecha, float(cant_usd), float(tc_origen), float(cant_usd), origen, descripcion.strip()),
        )
        conn.commit()
        return int(cur.lastrowid)


def eliminar_lote_pool(lote_id: int) -> None:

    with db.obtener_conexion() as conn:
        conn.execute("DELETE FROM inversiones_pool_usd WHERE id = ?", (lote_id,))
        conn.commit()


def _guardar_asignaciones(conn: sqlite3.Connection, operacion_id: int, asignaciones: list, costo_fiscal: float) -> None:
    conn.execute(
        "UPDATE inversiones_operaciones SET asignaciones_usd_json = ?, costo_fiscal = ? WHERE id = ?",
        (json.dumps(asignaciones, ensure_ascii=False), float(costo_fiscal), operacion_id),
    )


def asignar_lote_a_operacion(operacion_id: int, lote_id: int, cant_usd: float) -> None:
    """Asigna (parte de) un lote del pool a una compra en USD, descontando su
    disponibilidad y sumando la asignación al costo fiscal de la operación."""

    with db.obtener_conexion() as conn:
        op = conn.execute(
            "SELECT * FROM inversiones_operaciones WHERE id = ?", (operacion_id,)
        ).fetchone()
        lote = conn.execute(
            "SELECT * FROM inversiones_pool_usd WHERE id = ?", (lote_id,)
        ).fetchone()
        if not op or not lote:
            return
        cant = min(float(cant_usd), float(lote["disponible_usd"]))
        if cant <= 0:
            return
        asignaciones = json.loads(op["asignaciones_usd_json"] or "[]")
        asignaciones.append({
            "lote_id": lote_id, "cant_usd": cant, "tc_origen": float(lote["tc_origen"]),
            "ars": cant * float(lote["tc_origen"]), "origen": lote["origen"], "fecha_lote": lote["fecha"],
        })
        conn.execute(
            "UPDATE inversiones_pool_usd SET disponible_usd = disponible_usd - ? WHERE id = ?",
            (cant, lote_id),
        )
        _guardar_asignaciones(conn, operacion_id, asignaciones, sum(a["ars"] for a in asignaciones))
        conn.commit()


def quitar_asignacion(operacion_id: int, index: int) -> None:
    """Quita una asignación de la operación y devuelve el USD al lote de
    origen en el pool (si tenía uno — no si era el respaldo 'bna')."""

    with db.obtener_conexion() as conn:
        op = conn.execute(
            "SELECT * FROM inversiones_operaciones WHERE id = ?", (operacion_id,)
        ).fetchone()
        if not op:
            return
        asignaciones = json.loads(op["asignaciones_usd_json"] or "[]")
        if index < 0 or index >= len(asignaciones):
            return
        a = asignaciones.pop(index)
        lote_id = a.get("lote_id")
        if lote_id:
            conn.execute(
                "UPDATE inversiones_pool_usd SET disponible_usd = "
                "MIN(cant_usd, disponible_usd + ?) WHERE id = ?",
                (a["cant_usd"], lote_id),
            )
        costo_fiscal = sum(x["ars"] for x in asignaciones) if asignaciones else float(op["total_ars"])
        _guardar_asignaciones(conn, operacion_id, asignaciones, costo_fiscal)
        conn.commit()


def asignar_peps_automatico(operacion_id: int) -> None:
    """Reasigna una compra en USD por orden PEPS: devuelve al pool cualquier
    asignación previa, y toma lotes con fecha <= la de la operación empezando
    por el más antiguo. Si el pool no alcanza a cubrir el total, arma un lote
    de respaldo 'bna' al TC BNA vendedor de la fecha de la operación (mismo
    comportamiento que la herramienta original)."""

    with db.obtener_conexion() as conn:
        op = conn.execute(
            "SELECT * FROM inversiones_operaciones WHERE id = ?", (operacion_id,)
        ).fetchone()
        if not op:
            return

        asignaciones_prev = json.loads(op["asignaciones_usd_json"] or "[]")
        for a in asignaciones_prev:
            lote_id = a.get("lote_id")
            if lote_id:
                conn.execute(
                    "UPDATE inversiones_pool_usd SET disponible_usd = "
                    "MIN(cant_usd, disponible_usd + ?) WHERE id = ?",
                    (a["cant_usd"], lote_id),
                )

        lotes = conn.execute(
            "SELECT * FROM inversiones_pool_usd WHERE cliente_id = ? AND disponible_usd > 0 "
            "AND fecha <= ? ORDER BY fecha ASC, id ASC",
            (op["cliente_id"], op["fecha"]),
        ).fetchall()

        resta = float(op["total_usd"])
        nuevas: list = []
        for lote in lotes:
            if resta <= 0:
                break
            usar = min(float(lote["disponible_usd"]), resta)
            if usar <= 0:
                continue
            conn.execute(
                "UPDATE inversiones_pool_usd SET disponible_usd = disponible_usd - ? WHERE id = ?",
                (usar, lote["id"]),
            )
            nuevas.append({
                "lote_id": lote["id"], "cant_usd": usar, "tc_origen": float(lote["tc_origen"]),
                "ars": usar * float(lote["tc_origen"]), "origen": lote["origen"], "fecha_lote": lote["fecha"],
            })
            resta -= usar

        if resta > 0.0001:
            tc_fallback = _obtener_tc_con(conn, op["fecha"], "venta") or 1.0
            nuevas.append({
                "lote_id": None, "cant_usd": resta, "tc_origen": tc_fallback,
                "ars": resta * tc_fallback, "origen": "bna", "fecha_lote": op["fecha"],
            })

        _guardar_asignaciones(conn, operacion_id, nuevas, sum(a["ars"] for a in nuevas))
        conn.commit()


# --- Posición PEPS y control de cierre (Etapa 3 — Patrimonio) ---------------
#
# Solo Persona Humana por ahora: la tenencia al cierre se valúa "al origen"
# (costo histórico PEPS / costo fiscal según TC de origen del pool). No se
# aplica reexpresión por inflación (RECPAM) — eso queda para cuando el
# estudio necesite soportar Personas Jurídicas en Inversiones.

def calcular_posicion(cliente_id: int, periodo: str) -> list[dict]:
    """Posición PEPS al cierre del período, por instrumento.

    Consume los lotes de compra/apertura en el mismo orden en que se
    cargaron (PEPS) a medida que aparecen ventas/rescates, y devuelve para
    cada instrumento la tenencia remanente (cantidad, costo histórico ARS,
    costo USD al TC de compra, costo fiscal al TC de origen del pool) más
    los dividendos/rentas/amortizaciones cobrados y la comisión acumulada
    en el período. Instrumentos totalmente vendidos quedan con cantidad 0
    (estado 'cerrada') pero conservan lo cobrado en dividendos/rentas.
    """
    ops = listar_operaciones(cliente_id=cliente_id, periodo=periodo)
    grupos: dict[str, dict] = {}
    for op in ops:
        p = grupos.setdefault(op["instrumento"], {
            "instrumento": op["instrumento"], "tipo": op["tipo"],
            "tratamiento_fiscal": op["tratamiento_fiscal"],
            "alcanza_bienes_personales": op.get("alcanza_bienes_personales", "gravado"),
            "lotes": [], "div_ars": 0.0, "div_usd": 0.0,
            "amorts_ars": 0.0, "rentas_ars": 0.0, "comisiones": 0.0,
        })
        # el tratamiento fiscal más reciente manda (por si se corrigió en Depuración)
        p["tratamiento_fiscal"] = op["tratamiento_fiscal"]
        p["alcanza_bienes_personales"] = op.get("alcanza_bienes_personales", p["alcanza_bienes_personales"])

        if op["movimiento"] in ("compra", "apertura"):
            vn = float(op["vn"] or 0)
            cf_unit = (float(op["costo_fiscal"] or op["total_ars"]) / vn) if vn else 0.0
            p["lotes"].append({
                "cant": vn, "p_ars": float(op["precio_ars"] or 0), "p_usd": float(op["precio_usd"] or 0),
                "tc": float(op["tc"] or 1) or 1.0, "cf_unit": cf_unit,
                "tc_origen": float(op["tc_origen"] or 0), "fecha": op["fecha"], "moneda": op["moneda"],
            })
            p["comisiones"] += float(op["comision"] or 0)
        elif op["movimiento"] in ("venta", "rescate"):
            restante = float(op["vn"] or 0)
            for lote in p["lotes"]:
                if restante <= 0:
                    break
                usar = min(lote["cant"], restante)
                lote["cant"] -= usar
                restante -= usar
            p["comisiones"] += float(op["comision"] or 0)
        elif op["movimiento"] in ("dividendo", "renta"):
            p["div_ars"] += float(op["div_ars"] or 0)
            p["div_usd"] += float(op["div_usd"] or 0)
        elif op["movimiento"] == "amort":
            # La amortización es devolución de capital, no una venta: reduce
            # la cantidad de cuotapartes/partes/unidades del instrumento (se
            # consume del mismo lote PEPS que una venta/rescate) pero el
            # importe cobrado no genera un resultado — se registra aparte en
            # amorts_ars.
            p["amorts_ars"] += float(op["div_ars"] or op["total_ars"] or 0)
            restante = float(op["vn"] or 0)
            for lote in p["lotes"]:
                if restante <= 0:
                    break
                usar = min(lote["cant"], restante)
                lote["cant"] -= usar
                restante -= usar
            p["comisiones"] += float(op["comision"] or 0)
        elif op["movimiento"] == "caucion":
            p["rentas_ars"] += float(op["div_ars"] or op["total_ars"] or 0)

    resultado = []
    for p in grupos.values():
        lotes_rem = [l for l in p["lotes"] if l["cant"] > 1e-9]
        cant_total = sum(l["cant"] for l in lotes_rem)
        costo_ars = sum(l["cant"] * (l["p_ars"] or l["p_usd"] * l["tc"]) for l in lotes_rem)
        costo_usd = sum(l["cant"] * l["p_usd"] for l in lotes_rem)
        costo_fiscal = sum(l["cant"] * l["cf_unit"] for l in lotes_rem)
        resultado.append({
            "instrumento": p["instrumento"], "tipo": p["tipo"],
            "tratamiento_fiscal": p["tratamiento_fiscal"],
            "alcanza_bienes_personales": p["alcanza_bienes_personales"],
            "cantidad": cant_total,
            "precio_unit_ars": (costo_ars / cant_total) if cant_total > 1e-9 else 0.0,
            "costo_ars": costo_ars, "costo_usd": costo_usd, "costo_fiscal": costo_fiscal,
            "div_ars": p["div_ars"], "div_usd": p["div_usd"],
            "amorts_ars": p["amorts_ars"], "rentas_ars": p["rentas_ars"],
            "comisiones": p["comisiones"],
            "estado": "mantenida" if cant_total > 1e-9 else "cerrada",
        })
    return resultado


def calcular_movimientos(cliente_id: int, periodo: str) -> list[dict]:
    """Detalle cronológico de movimientos por instrumento (Etapa 3.5 —
    Movimientos), con el resultado realizado de cada venta/rescate contra el
    costo PEPS consumido en ese momento.

    Consume los lotes de compra/apertura en el mismo orden que
    ``calcular_posicion`` (misma lógica PEPS), así que el costo remanente al
    final de esta secuencia coincide siempre con lo que muestra Patrimonio.

    Para ventas/rescates registrados en USD, el resultado realizado de esa
    porción se descompone además en:
      - ``rendimiento_ars``: la ganancia/pérdida propia del instrumento.
        Para la porción de lo vendido cuyo costo se originó en dólares
        (compra/apertura cargada en USD, con su TC de origen), es la
        diferencia de precio en dólares (venta − compra) valuada al TC del
        día de la venta. Para la porción cuyo costo se originó en pesos
        (compra/apertura cargada en ARS — no hay dólares de origen que
        comparar), es directamente el resultado en pesos de esa porción.
      - ``diferencia_cambio_ars``: solo para la porción con costo de origen
        en dólares — la diferencia entre el TC del día de la venta y el TC
        de origen del capital en dólares invertido, aplicada sobre ese
        capital. Es la ganancia/pérdida que viene solo del movimiento del
        tipo de cambio, no de la actividad del cliente. La porción con
        costo de origen en pesos no aporta nada acá (no hay TC de origen
        que comparar).
    Ambas suman exactamente el ``resultado_ars`` (antes de comisión), tal
    como en el ejemplo: comprar USD 1 a TC 1500 y vender USD 2 a TC 1600 da
    USD 1 de rendimiento (a $1600) + $100 de diferencia de cambio.
    """
    ops = listar_operaciones(cliente_id=cliente_id, periodo=periodo)
    grupos: dict[str, dict] = {}
    for op in ops:
        g = grupos.setdefault(op["instrumento"], {
            "instrumento": op["instrumento"], "tipo": op["tipo"],
            "tratamiento_fiscal": op["tratamiento_fiscal"], "lotes": [], "movimientos": [],
        })
        # el tratamiento fiscal más reciente manda (por si se corrigió en Depuración)
        g["tratamiento_fiscal"] = op["tratamiento_fiscal"]

        fila = {
            "id": op["id"], "fecha": op["fecha"], "movimiento": op["movimiento"],
            "vn": float(op["vn"] or 0), "moneda": op["moneda"],
            "total_ars": float(op["total_ars"] or 0), "total_usd": float(op["total_usd"] or 0),
            "div_ars": float(op["div_ars"] or 0), "comision": float(op["comision"] or 0),
            "resultado_ars": None, "rendimiento_ars": None, "diferencia_cambio_ars": None,
        }

        if op["movimiento"] in ("compra", "apertura"):
            vn = float(op["vn"] or 0)
            cf_unit = (float(op["costo_fiscal"] or op["total_ars"]) / vn) if vn else 0.0
            p_usd_unit = (float(op["total_usd"] or 0) / vn) if vn else float(op["precio_usd"] or 0)
            g["lotes"].append({
                "cant": vn, "cf_unit": cf_unit, "p_usd": p_usd_unit, "moneda": op["moneda"],
            })
        elif op["movimiento"] in ("venta", "rescate"):
            restante = float(op["vn"] or 0)
            costo_consumido = 0.0
            tc_venta = float(op["tc"] or 0)
            precio_usd_venta = float(op["precio_usd"] or 0)
            es_usd = op["moneda"] == "USD" and tc_venta > 0
            rendimiento_ars = 0.0
            diferencia_cambio_ars = 0.0
            for lote in g["lotes"]:
                if restante <= 0:
                    break
                usar = min(lote["cant"], restante)
                costo_consumido += usar * lote["cf_unit"]
                if es_usd:
                    if lote.get("moneda") == "USD":
                        p_usd_lote = lote.get("p_usd", 0.0)
                        # ganancia/pérdida en USD del instrumento, valuada al TC de venta
                        rendimiento_ars += usar * (precio_usd_venta - p_usd_lote) * tc_venta
                        # diferencia de cambio: capital en USD del lote, entre su TC
                        # de origen (implícito en cf_unit/p_usd_lote) y el TC de venta
                        diferencia_cambio_ars += usar * p_usd_lote * tc_venta - usar * lote["cf_unit"]
                    else:
                        # el lote se había adquirido en pesos (sin dólares de
                        # origen): todo el resultado de esa porción es
                        # "rendimiento" en pesos — no hay componente de
                        # diferencia de cambio para separar.
                        rendimiento_ars += usar * precio_usd_venta * tc_venta - usar * lote["cf_unit"]
                lote["cant"] -= usar
                restante -= usar
            g["lotes"] = [l for l in g["lotes"] if l["cant"] > 1e-9]
            fila["resultado_ars"] = float(op["total_ars"] or 0) - costo_consumido - float(op["comision"] or 0)
            if es_usd:
                fila["rendimiento_ars"] = rendimiento_ars
                fila["diferencia_cambio_ars"] = diferencia_cambio_ars
        elif op["movimiento"] == "amort":
            # Igual que en calcular_posicion: la amortización reduce la
            # cantidad de cuotapartes/partes/unidades (consume el lote PEPS)
            # pero NO es una venta — no se registra resultado_ars ni se abre
            # en rendimiento/diferencia de cambio, es devolución de capital.
            restante = float(op["vn"] or 0)
            for lote in g["lotes"]:
                if restante <= 0:
                    break
                usar = min(lote["cant"], restante)
                lote["cant"] -= usar
                restante -= usar
            g["lotes"] = [l for l in g["lotes"] if l["cant"] > 1e-9]

        g["movimientos"].append(fila)

    return list(grupos.values())


# --- Resumen patrimonial (Etapa Resumen — Patrimonio) -----------------------

def _calcular_impuesto(
    tratamiento_fiscal: str, base_ars: float, alicuota_escala: float,
) -> tuple[float, float, float]:
    """(base imponible, alícuota, impuesto aproximado) según el tratamiento
    fiscal de la operación. Cedular 15%/5% y Retención 7% tienen tasa fija
    (impuesto exacto). "Gravado a escala" no tiene una tasa única posible —
    depende de TODOS los ingresos anuales del cliente, no solo de
    Inversiones — así que se usa una alícuota de referencia editable
    (``alicuota_escala``, por defecto la escala máxima vigente) y el importe
    resultante es una APROXIMACIÓN a revisar por el contador según el caso
    real. Exento/N.A. no generan impuesto. El impuesto no se calcula sobre
    resultados negativos (una pérdida no genera impuesto a pagar)."""
    base = base_ars if base_ars > 0 else 0.0
    if tratamiento_fiscal == "ced15":
        return base, 0.15, base * 0.15
    if tratamiento_fiscal == "ced5":
        return base, 0.05, base * 0.05
    if tratamiento_fiscal == "ret7":
        return base, 0.07, base * 0.07
    if tratamiento_fiscal in ("gravado", "ordinario"):
        return base, alicuota_escala, base * alicuota_escala
    return base, 0.0, 0.0


def calcular_resumen_patrimonial(
    cliente_id: int, periodo: str, alicuota_escala: float = ALICUOTA_ESCALA_DEFAULT,
) -> dict:
    """Resumen patrimonial de cierre para la solapa Patrimonio: composición de
    la cartera a costo histórico (Ganancias) y a valor de cierre (Bienes
    Personales, con la cotización de las planillas oficiales 2025 o la carga
    manual), variación patrimonial del período descompuesta en rendimientos/
    intereses/rentas vs. compra-venta vs. diferencia de cambio, y el
    impuesto aproximado de cada operación/instrumento según su tratamiento
    fiscal. Todos los totales salen de recorrer ``calcular_posicion`` /
    ``calcular_movimientos`` — no hay ningún número acá que no derive de esos
    dos cálculos, para que el Excel exportado pueda enlazarse con fórmulas a
    los datos puros sin duplicar la lógica fiscal en otro lado."""
    posicion = calcular_posicion(cliente_id, periodo)
    movimientos_por_inst = calcular_movimientos(cliente_id, periodo)
    tc_cierre = _obtener_tc_cierre(periodo)

    tenencias = []
    costo_hist_total = 0.0
    valor_cierre_total = 0.0
    for p in posicion:
        if p["estado"] != "mantenida" or p["cantidad"] <= 1e-9:
            continue
        vc = obtener_valor_cierre(cliente_id, periodo, p["instrumento"])
        cotiz = vc["cotizacion"]
        valor_cierre_unit_ars = cotiz * tc_cierre if vc["moneda"] == "USD" else cotiz
        valor_cierre_total_ars = p["cantidad"] * valor_cierre_unit_ars
        costo_hist_total += p["costo_ars"]
        if p["alcanza_bienes_personales"] == "gravado":
            valor_cierre_total += valor_cierre_total_ars
        tenencias.append({
            "instrumento": p["instrumento"], "tipo": p["tipo"],
            "tratamiento_fiscal": p["tratamiento_fiscal"],
            "alcanza_bienes_personales": p["alcanza_bienes_personales"],
            "cantidad": p["cantidad"],
            "costo_historico_ars": p["costo_ars"],
            "cotizacion_cierre": cotiz, "moneda_cotizacion": vc["moneda"],
            "fuente_valor_cierre": vc["fuente"],
            "valor_cierre_unitario_ars": valor_cierre_unit_ars,
            "valor_cierre_total_ars": valor_cierre_total_ars,
        })

    rendimientos_ars = sum(p["div_ars"] + p["amorts_ars"] + p["rentas_ars"] for p in posicion)
    compraventa_ars = 0.0
    diferencia_cambio_ars = 0.0
    detalle_impuesto = []
    detalle_variacion = []

    for g in movimientos_por_inst:
        for m in g["movimientos"]:
            if m["resultado_ars"] is None:
                continue
            rend = m["rendimiento_ars"] if m["rendimiento_ars"] is not None else m["resultado_ars"]
            dif = m["diferencia_cambio_ars"] or 0.0
            compraventa_ars += rend
            diferencia_cambio_ars += dif
            detalle_variacion.append({
                "fecha": m["fecha"], "instrumento": g["instrumento"],
                "movimiento": MOV_LABEL.get(m["movimiento"], m["movimiento"]),
                "rendimiento_intereses_rentas_ars": 0.0,
                "compraventa_ars": rend, "diferencia_cambio_ars": dif,
            })
            base, alic, impuesto = _calcular_impuesto(g["tratamiento_fiscal"], m["resultado_ars"], alicuota_escala)
            detalle_impuesto.append({
                "fecha": m["fecha"], "instrumento": g["instrumento"],
                "movimiento": MOV_LABEL.get(m["movimiento"], m["movimiento"]),
                "tratamiento_fiscal": g["tratamiento_fiscal"],
                "base_ars": base, "alicuota": alic, "impuesto_ars": impuesto,
            })

    for p in posicion:
        base_rend = p["div_ars"] + p["amorts_ars"] + p["rentas_ars"]
        if base_rend <= 1e-9:
            continue
        detalle_variacion.append({
            "fecha": f"{periodo} (total)", "instrumento": p["instrumento"],
            "movimiento": "Dividendos/rentas/amort.",
            "rendimiento_intereses_rentas_ars": base_rend,
            "compraventa_ars": 0.0, "diferencia_cambio_ars": 0.0,
        })
        base, alic, impuesto = _calcular_impuesto(p["tratamiento_fiscal"], base_rend, alicuota_escala)
        detalle_impuesto.append({
            "fecha": f"{periodo} (total)", "instrumento": p["instrumento"],
            "movimiento": "Dividendos/rentas/amort.",
            "tratamiento_fiscal": p["tratamiento_fiscal"],
            "base_ars": base, "alicuota": alic, "impuesto_ars": impuesto,
        })

    impuesto_exacto_total = sum(d["impuesto_ars"] for d in detalle_impuesto if d["tratamiento_fiscal"] != "gravado")
    impuesto_aproximado_escala = sum(d["impuesto_ars"] for d in detalle_impuesto if d["tratamiento_fiscal"] == "gravado")

    return {
        "periodo": periodo, "alicuota_escala": alicuota_escala, "tc_cierre": tc_cierre,
        "tenencias": tenencias,
        "costo_historico_total_ars": costo_hist_total,
        "valor_cierre_total_ars": valor_cierre_total,
        "variacion_patrimonial": {
            "rendimientos_ars": rendimientos_ars,
            "compraventa_ars": compraventa_ars,
            "diferencia_cambio_ars": diferencia_cambio_ars,
            "total_ars": rendimientos_ars + compraventa_ars + diferencia_cambio_ars,
        },
        "detalle_impuesto": detalle_impuesto,
        "detalle_variacion": detalle_variacion,
        "impuesto_exacto_total_ars": impuesto_exacto_total,
        "impuesto_aproximado_escala_ars": impuesto_aproximado_escala,
    }


def _obtener_tc_cierre(periodo: str) -> float:
    """TC BNA vendedor al 31/12 del período (o el hábil disponible más
    cercano) — se usa para convertir a pesos las cotizaciones de cierre que
    vienen en dólares (Bienes Personales)."""
    tc = obtener_tc(f"{periodo}-12-31", "venta")
    return float(tc) if tc else 1.0


def crear_tenencia_control(
    cliente_id: int, periodo: str, instrumento: str, vn_control: float,
    valor_ars_control: float = 0.0, notas: str = "",
) -> int:
    """Registra la tenencia real de un instrumento según el extracto/comprobante
    de cierre (fin de ejercicio, o 31/12 para Persona Humana), para poder
    compararla contra la posición PEPS calculada y detectar diferencias."""

    with db.obtener_conexion() as conn:
        cur = conn.execute(
            """
            INSERT INTO inversiones_tenencias_control (
                cliente_id, periodo, instrumento, vn_control, valor_ars_control, notas
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (cliente_id, periodo, instrumento.strip(), float(vn_control), float(valor_ars_control), notas.strip()),
        )
        conn.commit()
        return int(cur.lastrowid)


def listar_tenencias_control(cliente_id: int, periodo: str) -> list[dict]:

    with db.obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT * FROM inversiones_tenencias_control WHERE cliente_id = ? AND periodo = ? "
            "ORDER BY instrumento ASC",
            (cliente_id, periodo),
        ).fetchall()
    return [dict(f) for f in filas]


def eliminar_tenencia_control(control_id: int) -> None:

    with db.obtener_conexion() as conn:
        conn.execute("DELETE FROM inversiones_tenencias_control WHERE id = ?", (control_id,))
        conn.commit()
