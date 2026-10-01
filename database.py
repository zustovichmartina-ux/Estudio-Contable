"""Gestión de la base de datos SQLite para clientes del estudio contable."""

import hashlib
import json
import logging
import os
import re
import sqlite3
import threading
import time
import unicodedata
import warnings
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd

try:
    from ddgs import DDGS
except Exception:
    DDGS = None

try:
    import libsql_experimental as _libsql
except Exception:
    _libsql = None

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "estudio_contable.db"
CUITS_AUXILIARES_PATH = BASE_DIR / "Cuits_Auxiliares.xlsx"
LISTA_EMPRESAS_PATH = BASE_DIR / "Lista_empresas.xlsx"
PLAN_CUENTAS_DEFAULT_LEGACY = BASE_DIR / "planes de cuentas" / "Cuentas contables (4).xlsx"
PLAN_CUENTAS_DEFAULT_REPO = BASE_DIR / "data" / "planes_cuentas" / "plan_default.xlsx"
SEED_SOCIEDADES_PJ_PATH = BASE_DIR / "data" / "seed" / "sociedades_pj.json"
DATA_PLANES_DIR = BASE_DIR / "data" / "planes_cuentas"


def _plan_cuentas_default_path() -> Path:
    """Plan genérico: prioriza el del repo (Cloud) y cae al legacy local."""
    if PLAN_CUENTAS_DEFAULT_REPO.is_file():
        return PLAN_CUENTAS_DEFAULT_REPO
    return PLAN_CUENTAS_DEFAULT_LEGACY


# Compat: código legacy importa PLAN_CUENTAS_DEFAULT
PLAN_CUENTAS_DEFAULT = _plan_cuentas_default_path()

TIPOS_PERSONA = ("Persona Jurídica", "Persona Física", "Monotributista")
PREFIJO_CUIT_TEMPORAL = "99"
CUIT_TEMPORAL_INICIO = 99000000001


def _normalizar_nombre(texto: str) -> str:
    """Normaliza nombre para comparaciones fuzzy."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z0-9 ]", " ", texto.upper())
    return re.sub(r"\s+", " ", texto).strip()


def _es_cuit_valido(cuit: str) -> bool:
    """True si el CUIT tiene 11 dígitos y no es temporal."""
    limpio = str(cuit or "").replace("-", "").strip()
    return len(limpio) == 11 and limpio.isdigit() and not limpio.startswith(PREFIJO_CUIT_TEMPORAL)


def _categorizar_tipo(nombre: str) -> str:
    """Clasifica Persona Jurídica o Física según el nombre."""
    nom = nombre.strip().upper()
    if re.search(r"\b(S\.?A\.?|S\.?R\.?L\.?|S\.A\.S\.?|SOCIEDAD|CONSORCIO)\b", nom):
        return "Persona Jurídica"
    return "Persona Física"


def buscar_mes_cierre_web(cuit: str) -> Optional[int]:
    """Busca en internet el mes de cierre de balance para el CUIT dado."""
    if not cuit or len(cuit) < 10:
        return None
    if DDGS is None:
        return None

    query = f"cuit {cuit} cierre de balance"
    try:
        resultados = DDGS().text(query, max_results=3)
        meses = {
            "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
            "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
            "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
        }
        for r in resultados:
            texto = (r.get("title", "") + " " + r.get("body", "")).lower()
            
            # Buscar patrones como "cierre de balance: 12" o "cierre de balance: diciembre"
            match_mes = re.search(r"cierre(?: de balance)?\s*(?::|es el|en)?\s*([a-z]+|\d{1,2})\b", texto)
            if match_mes:
                val = match_mes.group(1)
                if val.isdigit():
                    mes = int(val)
                    if 1 <= mes <= 12:
                        return mes
                elif val in meses:
                    return meses[val]
            
            # Otro intento si aparece explícitamente el mes
            for nombre_mes, num_mes in meses.items():
                if f"cierre {nombre_mes}" in texto or f"cierre de {nombre_mes}" in texto:
                    return num_mes
                    
    except Exception as e:
        print(f"Error al buscar mes de cierre en web para {cuit}: {e}")
        pass
    
    return None


def _credenciales_turso() -> tuple[str, str]:
    """Lee URL y token de Turso desde variables de entorno o Secrets de Streamlit."""
    url = os.environ.get("TURSO_DATABASE_URL", "") or ""
    token = os.environ.get("TURSO_AUTH_TOKEN", "") or ""
    if not url:
        try:
            import streamlit as st  # import perezoso: no depender de Streamlit fuera de la app

            url = str(st.secrets.get("TURSO_DATABASE_URL", "") or "")
            token = str(st.secrets.get("TURSO_AUTH_TOKEN", "") or "")
        except Exception:
            pass
    return url.strip(), token.strip()


class _FilaCompat(tuple):
    """Tupla que además permite acceso por nombre de columna, como sqlite3.Row."""

    _cols: tuple = ()

    def __new__(cls, columnas, valores):
        obj = super().__new__(cls, valores)
        obj._cols = tuple(columnas)
        return obj

    def __getitem__(self, clave):
        if isinstance(clave, str):
            try:
                idx = self._cols.index(clave)
            except ValueError:
                raise KeyError(clave)
            return tuple.__getitem__(self, idx)
        return tuple.__getitem__(self, clave)

    def get(self, clave, default=None):
        try:
            return self[clave]
        except (KeyError, IndexError):
            return default

    def keys(self):
        return self._cols


class _CursorCompatTurso:
    """Envuelve el cursor de libsql para que fetchone/fetchall devuelvan _FilaCompat."""

    def __init__(self, cur):
        self._cur = cur

    def _envolver(self, fila):
        if fila is None:
            return None
        columnas = [d[0] for d in (self._cur.description or [])]
        return _FilaCompat(columnas, fila)

    def fetchone(self):
        return self._envolver(self._cur.fetchone())

    def fetchall(self):
        return [self._envolver(f) for f in self._cur.fetchall()]

    def fetchmany(self, size=None):
        filas = self._cur.fetchmany(size) if size is not None else self._cur.fetchmany()
        return [self._envolver(f) for f in filas]

    def __iter__(self):
        # libsql no hace iterable el cursor. Sin esto, `for fila in conn.execute(...)` revienta.
        return iter(self.fetchall())

    def __getattr__(self, nombre):
        return getattr(self._cur, nombre)


# Streamlit puede atender varias sesiones en threads del mismo proceso; la
# conexión de libsql (réplica local + Turso) se comparte, así que serializamos
# las consultas con un lock global. El refresco de la réplica va en segundo
# plano (sync_interval): un rerun no sale a la red ni toma ese lock para sync.
_turso_conn_lock = threading.Lock()
_turso_conn_obj = None
_turso_es_replica = False
_sync_fondo = False  # True si connect() aceptó sync_interval
_SYNC_INTERVALO_S = 15.0
_ultimo_sync_periodico = 0.0
_sync_periodico_lock = threading.Lock()
# Fuera de Streamlit (tests, scripts) el flag y los contadores son de proceso.
# Con una sesión, viven en st.session_state: ver _bucket_sesion().
_sync_hecho_en_rerun = False
_sync_ok = True
_generacion_datos = 0
_ddl_bloqueado = False
_contadores = {"consultas": 0, "syncs": 0}
_sync_s = 0.0
_t0_rerun = 0.0
_pagina_rerun = "arranque"
_RE_MUTACION = re.compile(r"^\s*(INSERT|UPDATE|DELETE|REPLACE)\b", re.IGNORECASE)
_RE_DESTRUYE_CLIENTES = re.compile(
    r"\b(?:DROP\s+TABLE\s+(?:IF\s+EXISTS\s+)?CLIENTES|CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?CLIENTES_NEW|RENAME\s+TO\s+CLIENTES)\b",
    re.IGNORECASE,
)


def _obtener_conn_turso_singleton():
    global _turso_conn_obj, _turso_es_replica
    if _turso_conn_obj is not None:
        return _turso_conn_obj
    url, token = _credenciales_turso()
    if _libsql is None or not url:
        return None
    with _turso_conn_lock:
        if _turso_conn_obj is None:
            try:
                _turso_conn_obj = _conectar_libsql(url, token)
                _turso_es_replica = _turso_conn_obj is not None
            except Exception:
                _turso_conn_obj = None
                _turso_es_replica = False
    return _turso_conn_obj


def _conectar_libsql(url: str, token: str):
    """Réplica local. sync_interval refresca en segundo plano (segundos).

    Si esta versión del cliente no acepta el argumento, queda el fallback de
    `_sync_periodico_si_toca` (como máximo cada 15s, fuera del lock).
    """
    global _sync_fondo
    try:
        conn = _libsql.connect(
            str(DB_PATH),
            sync_url=url,
            auth_token=token,
            sync_interval=_SYNC_INTERVALO_S,
        )
    except TypeError:
        _sync_fondo = False
        return _libsql.connect(str(DB_PATH), sync_url=url, auth_token=token)
    _sync_fondo = True
    return conn


def _bucket_sesion() -> dict | None:
    """Estado del rerun de esta sesión, o None si no hay contexto de Streamlit.

    Import perezoso: los tests y scripts usan database.py sin Streamlit.
    """
    try:
        from streamlit.runtime.scriptrunner_utils.script_run_context import (
            get_script_run_ctx,
        )
    except Exception:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
        except Exception:
            return None
    try:
        ctx = get_script_run_ctx(suppress_warning=True)
    except TypeError:
        try:
            ctx = get_script_run_ctx()
        except Exception:
            return None
    except Exception:
        return None
    if ctx is None:
        return None
    try:
        import streamlit as st

        ss = st.session_state
    except Exception:
        return None
    estado = ss.get("_ec_rerun")
    if not isinstance(estado, dict):
        estado = {
            "sync_hecho": False,
            "consultas": 0,
            "syncs": 0,
            "sync_s": 0.0,
            "t0": 0.0,
            "pagina": "arranque",
        }
        ss["_ec_rerun"] = estado
    return estado


def _sync_ya_hecho() -> bool:
    bucket = _bucket_sesion()
    if bucket is None:
        return bool(_sync_hecho_en_rerun)
    return bool(bucket.get("sync_hecho"))


def _marcar_sync_hecho(valor: bool) -> None:
    global _sync_hecho_en_rerun
    bucket = _bucket_sesion()
    if bucket is None:
        _sync_hecho_en_rerun = valor
        return
    bucket["sync_hecho"] = valor


def _sumar_consulta() -> None:
    bucket = _bucket_sesion()
    if bucket is None:
        _contadores["consultas"] += 1
        return
    bucket["consultas"] = int(bucket.get("consultas") or 0) + 1


def _sumar_sync(segundos: float) -> None:
    global _sync_s
    bucket = _bucket_sesion()
    if bucket is None:
        _contadores["syncs"] += 1
        _sync_s += float(segundos)
        return
    bucket["syncs"] = int(bucket.get("syncs") or 0) + 1
    bucket["sync_s"] = float(bucket.get("sync_s") or 0.0) + float(segundos)


def metricas_rerun() -> dict[str, float]:
    """Consultas, syncs y segundos de red acumulados en este rerun."""
    bucket = _bucket_sesion()
    if bucket is None:
        return {
            "consultas": int(_contadores["consultas"]),
            "syncs": int(_contadores["syncs"]),
            "sync_s": float(_sync_s),
        }
    return {
        "consultas": int(bucket.get("consultas") or 0),
        "syncs": int(bucket.get("syncs") or 0),
        "sync_s": float(bucket.get("sync_s") or 0.0),
    }


def generacion_datos() -> int:
    """Sube en cada escritura. Las lecturas cacheadas la usan para invalidarse."""
    return _generacion_datos


def contadores() -> dict[str, int]:
    """Consultas y syncs desde el último `comenzar_rerun` (de esta sesión)."""
    metricas = metricas_rerun()
    return {"consultas": int(metricas["consultas"]), "syncs": int(metricas["syncs"])}


def ddl_permitido() -> bool:
    """False si la réplica no se pudo leer: no se manda DDL a ciegas."""
    return not _ddl_bloqueado


def _invalidar_lecturas() -> None:
    global _generacion_datos
    _generacion_datos += 1


def _sync_red() -> bool:
    """Un sync con Turso. No toma `_turso_conn_lock`: la red no frena las otras sesiones."""
    global _sync_ok
    conn = _turso_conn_obj
    if conn is None:
        _sync_ok = False
        return False
    t0 = time.perf_counter()
    try:
        conn.sync()
    except Exception:
        logging.getLogger(__name__).warning(
            "No se pudo sincronizar la réplica; no se toma el archivo local como esquema.",
            exc_info=True,
        )
        _sync_ok = False
        return False
    _sumar_sync(time.perf_counter() - t0)
    _sync_ok = True
    return True


def sincronizar_replica_una_vez() -> bool:
    """Como máximo un sync en este rerun. Lo usa el arranque del esquema, no cada pantalla."""
    if not _turso_es_replica:
        return True
    if _sync_ya_hecho():
        return _sync_ok
    _marcar_sync_hecho(True)
    return _sync_red()


def sincronizar_cola_rutinas() -> bool:
    """Trae el estado de rutina_pedidos. Solo el auto-refresh y el botón Refrescar."""
    if not _turso_es_replica:
        return True
    return _sync_red()


def _sync_periodico_si_toca() -> None:
    """Fallback si no hay sync_interval: como máximo un sync cada 15s, fuera del lock."""
    global _ultimo_sync_periodico
    if not _turso_es_replica or _turso_conn_obj is None:
        return
    ahora = time.monotonic()
    with _sync_periodico_lock:
        if ahora - _ultimo_sync_periodico < _SYNC_INTERVALO_S:
            return
        _ultimo_sync_periodico = ahora
    _sync_red()


def comenzar_rerun() -> None:
    """Arranque de un rerun. No sale a la red: la réplica se refresca sola.

    Si el cliente no aceptó sync_interval, un sync como máximo cada 15s y
    fuera de `_turso_conn_lock`.
    """
    global _sync_hecho_en_rerun, _sync_s, _t0_rerun, _pagina_rerun
    bucket = _bucket_sesion()
    if bucket is None:
        _sync_hecho_en_rerun = False
        _contadores["consultas"] = 0
        _contadores["syncs"] = 0
        _sync_s = 0.0
        _t0_rerun = time.perf_counter()
        _pagina_rerun = "arranque"
    else:
        bucket["sync_hecho"] = False
        bucket["consultas"] = 0
        bucket["syncs"] = 0
        bucket["sync_s"] = 0.0
        bucket["t0"] = time.perf_counter()
        bucket["pagina"] = "arranque"
    if not _sync_fondo:
        _sync_periodico_si_toca()


def anotar_pagina(nombre: str) -> None:
    """Nombre que va a salir en la línea de tiempo del rerun."""
    global _pagina_rerun
    texto = str(nombre or "").strip() or "arranque"
    bucket = _bucket_sesion()
    if bucket is None:
        _pagina_rerun = texto
        return
    bucket["pagina"] = texto


def imprimir_linea_rerun(
    *,
    total: float,
    sync_s: float,
    consultas: int,
    syncs: int,
    pagina: str,
) -> None:
    """Una línea a stdout para Manage app → Logs."""
    print(
        f"RERUN total={total:.3f}s sync={sync_s:.3f}s "
        f"consultas={int(consultas)} syncs={int(syncs)} pagina={pagina}",
        flush=True,
    )


def cerrar_rerun() -> None:
    """Imprime el tiempo del rerun que abrió `comenzar_rerun`."""
    bucket = _bucket_sesion()
    if bucket is None:
        total = time.perf_counter() - float(_t0_rerun or time.perf_counter())
        imprimir_linea_rerun(
            total=total,
            sync_s=_sync_s,
            consultas=_contadores["consultas"],
            syncs=_contadores["syncs"],
            pagina=_pagina_rerun or "arranque",
        )
        return
    t0 = float(bucket.get("t0") or time.perf_counter())
    imprimir_linea_rerun(
        total=time.perf_counter() - t0,
        sync_s=float(bucket.get("sync_s") or 0.0),
        consultas=int(bucket.get("consultas") or 0),
        syncs=int(bucket.get("syncs") or 0),
        pagina=str(bucket.get("pagina") or "arranque"),
    )


def _parametros_libsql(parametros):
    """libsql solo acepta tuple; sqlite3 también acepta list."""
    if isinstance(parametros, list):
        return tuple(parametros)
    return parametros


def _es_error_sql_libsql(exc: BaseException) -> bool:
    """True si el fallo es de SQLite y no de la red.

    Turso remoto lo envuelve: ``Hrana: stream error: ... "SQLite error: ..." SQLITE_...``.
    libsql embebido tira el texto pelado (``duplicate column name: ...``).
    """
    texto = str(exc)
    if "SQLite error" in texto or "SQLITE_" in texto:
        return True
    if "Hrana" in texto or "stream error" in texto:
        return False
    return any(
        marca in texto
        for marca in (
            "duplicate column",
            "already exists",
            "no such table",
            "no such column",
            "constraint failed",
            "syntax error",
        )
    )


def _traducir_error_libsql(exc: BaseException) -> None:
    """Pasa un error SQL de libsql a sqlite3, o tira la conexión si fue de red."""
    if _es_error_sql_libsql(exc):
        texto = str(exc)
        if any(marca in texto for marca in ("UNIQUE", "constraint", "CONSTRAINT")):
            raise sqlite3.IntegrityError(texto) from exc
        raise sqlite3.OperationalError(texto) from exc
    global _turso_conn_obj
    _turso_conn_obj = None  # próxima conexión reintenta desde cero
    raise exc


_ES_SQL_ESCRITURA = re.compile(
    r"^\s*(INSERT|UPDATE|DELETE|REPLACE|CREATE|ALTER|DROP)\b", re.IGNORECASE
)


class _ConexionCompatTurso:
    """
    Envuelve la conexión compartida de libsql con la misma API que usa el resto
    del código para sqlite3.Connection (execute/commit/with ... as conn).

    La réplica se refresca sola (`sync_interval`). Un commit sincroniza con la
    nube solo si hubo una escritura real (``_sucio``). Un commit sin escritura
    no sale a la red. La cola de rutinas pide un sync aparte.
    """

    def __init__(self, conn):
        self._conn = conn
        self._dirty = False  # hubo un INSERT/UPDATE/DELETE: invalidar el cache de lecturas
        self._sucio = False  # hubo una escritura (incluye DDL) desde el último sync()
        self.row_factory = None  # compat: el resto del código no necesita fijarlo

    def _anotar(self, sql) -> None:
        texto = str(sql)
        if _RE_MUTACION.match(texto):
            self._dirty = True
        if _ES_SQL_ESCRITURA.match(texto):
            self._sucio = True

    def _reintentar_si_rota(self, exc):
        _traducir_error_libsql(exc)

    def execute(self, sql, parametros=()):
        _sumar_consulta()
        self._anotar(sql)
        with _turso_conn_lock:
            try:
                return _CursorCompatTurso(self._conn.execute(sql, _parametros_libsql(parametros)))
            except Exception as exc:
                self._reintentar_si_rota(exc)

    def executemany(self, sql, secuencia):
        """Un INSERT de muchas filas sale en un solo execute (varios VALUES).

        El ``executemany`` de libsql hace un ``execute`` por fila y, en Turso,
        cada uno es un viaje de red. Un UPDATE u otro SQL sigue por el driver.
        """
        _sumar_consulta()
        with _turso_conn_lock:
            filas = [_parametros_libsql(fila) for fila in secuencia]
            if not filas:
                return None
            self._anotar(sql)
            self._sucio = True
            masivo = _insert_masivo(sql, filas)
            try:
                if masivo is None:
                    return _CursorCompatTurso(self._conn.executemany(sql, filas))
                ultimo = None
                for sentencia, parametros in masivo:
                    ultimo = self._conn.execute(sentencia, parametros)
                return _CursorCompatTurso(ultimo) if ultimo is not None else None
            except Exception as exc:
                self._reintentar_si_rota(exc)

    def executescript(self, sql):
        """Un script entero en un batch de Hrana (``cursor().executescript``).

        ``Connection.executescript`` traga el error; el del cursor no.
        """
        _sumar_consulta()
        self._anotar(sql)
        if _RE_MUTACION.search(str(sql)):
            self._dirty = True
        self._sucio = True
        with _turso_conn_lock:
            try:
                return self._conn.cursor().executescript(sql)
            except Exception as exc:
                self._reintentar_si_rota(exc)

    def cursor(self):
        with _turso_conn_lock:
            return _CursorCompatTurso(self._conn.cursor())

    def commit(self):
        hizo_sync = False
        sync_s = 0.0
        with _turso_conn_lock:
            try:
                self._conn.commit()
            except Exception as exc:
                self._reintentar_si_rota(exc)
            if self._sucio:
                self._sucio = False
                try:
                    t0 = time.perf_counter()
                    self._conn.sync()
                    sync_s = time.perf_counter() - t0
                    hizo_sync = True
                except Exception:
                    pass
        if hizo_sync:
            _sumar_sync(sync_s)
        if self._dirty:
            _invalidar_lecturas()
            self._dirty = False

    def rollback(self):
        with _turso_conn_lock:
            try:
                self._conn.rollback()
            except Exception:
                pass

    def close(self):
        pass  # conexión compartida: no se cierra por-caller

    def __enter__(self):
        return self

    def __exit__(self, tipo_exc, exc, tb):
        if tipo_exc is None:
            try:
                self.commit()
            except Exception:
                pass
        else:
            self.rollback()
        return False


class _ConexionLocal:
    """SQLite del disco. Misma marca de escritura que la réplica, para invalidar el cache."""

    def __init__(self, conn: sqlite3.Connection):
        self._conn = conn
        self._dirty = False

    def _anotar(self, sql) -> None:
        if _RE_MUTACION.match(str(sql)):
            self._dirty = True

    def execute(self, sql, parametros=()):
        self._anotar(sql)
        return self._conn.execute(sql, parametros)

    def executemany(self, sql, secuencia):
        self._anotar(sql)
        return self._conn.executemany(sql, secuencia)

    def executescript(self, sql):
        if _RE_MUTACION.search(str(sql)):
            self._dirty = True
        return self._conn.executescript(sql)

    def commit(self):
        self._conn.commit()
        if self._dirty:
            _invalidar_lecturas()
            self._dirty = False

    def __enter__(self):
        return self

    def __exit__(self, tipo_exc, exc, tb):
        if tipo_exc is None:
            self.commit()
        else:
            self._conn.rollback()
        return False

    def __getattr__(self, nombre):
        return getattr(self._conn, nombre)


def obtener_conexion():
    """
    Abre conexión con filas accesibles por nombre de columna.
    Si hay credenciales de Turso (Secrets TURSO_DATABASE_URL/TURSO_AUTH_TOKEN o
    variables de entorno), reutiliza una réplica local sincronizada con Turso
    (una sola conexión por proceso) para que los datos sobrevivan a un
    reinicio o "dormida" de la app. Si no hay credenciales, usa el SQLite
    local de siempre.
    """
    conn = _obtener_conn_turso_singleton()
    if conn is not None:
        return _ConexionCompatTurso(conn)
    conn2 = sqlite3.connect(DB_PATH)
    conn2.row_factory = sqlite3.Row
    return _ConexionLocal(conn2)


def _asegurar_app_meta(conn) -> None:
    """Marcas de siembra (versión o hash). Una fila por semilla, no un flag en disco."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_meta (
            clave TEXT PRIMARY KEY,
            valor TEXT NOT NULL
        )
        """
    )


def _valor_fila(fila, clave: str):
    if fila is None:
        return None
    if hasattr(fila, "keys"):
        return fila[clave]
    return fila[0]


def leer_semilla(clave: str, conn=None) -> str | None:
    """Lee la marca de una semilla. Con `conn`, no abre otra conexión."""
    if conn is None:
        with obtener_conexion() as propia:
            return leer_semilla(clave, propia)
    _asegurar_app_meta(conn)
    fila = conn.execute(
        "SELECT valor FROM app_meta WHERE clave = ?",
        (clave,),
    ).fetchone()
    valor = _valor_fila(fila, "valor")
    return None if valor is None else str(valor)


def guardar_semilla(clave: str, valor: str, conn=None) -> None:
    """Guarda la marca. Si no hay `conn`, confirma sola; si hay, la confirma el llamador."""
    if conn is None:
        with obtener_conexion() as propia:
            guardar_semilla(clave, valor, propia)
            propia.commit()
        return
    _asegurar_app_meta(conn)
    conn.execute(
        "INSERT INTO app_meta (clave, valor) VALUES (?, ?) "
        "ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor",
        (clave, valor),
    )


def huella_semilla(*partes: str) -> str:
    return hashlib.sha256("\n".join(partes).encode("utf-8")).hexdigest()


# Subir este número cuando cambie el DDL (columna o tabla nueva). Si coincide
# con app_meta, inicializar_bd no vuelve a mandar CREATE/ALTER ni siembras.
SCHEMA_VERSION = "1"
_CLAVE_SCHEMA = "schema_version"

# Tope de SQLite (32766) y de un request de Turso. Un INSERT masivo se parte.
_MAX_PARAMS_INSERT = 900
_MAX_BYTES_INSERT = 350_000

_RE_INSERT_VALORES = re.compile(
    r"^(?P<head>INSERT\b.+?)\s+VALUES\s*\(\s*\?(?:\s*,\s*\?)*\s*\)\s*(?P<tail>ON\s+CONFLICT\b.*)?$",
    re.IGNORECASE | re.DOTALL,
)


def _peso_parametro(valor) -> int:
    if isinstance(valor, str):
        return len(valor)
    if isinstance(valor, bytes):
        return len(valor)
    return 8


def _insert_masivo(sql: str, filas: list) -> list[tuple[str, tuple]] | None:
    """Arma uno o más INSERT con varios VALUES. None si el SQL no es un INSERT simple."""
    plano = " ".join(str(sql).split()).rstrip(";").strip()
    coincidencia = _RE_INSERT_VALORES.match(plano)
    if not coincidencia:
        return None
    if not filas or not isinstance(filas[0], (tuple, list)):
        return None
    ancho = len(filas[0])
    if ancho < 1 or any(len(fila) != ancho for fila in filas):
        return None
    cabeza = coincidencia.group("head").strip()
    cola = (coincidencia.group("tail") or "").strip()
    marca_fila = "(" + ",".join(["?"] * ancho) + ")"
    lotes: list[tuple[str, tuple]] = []
    lote: list = []
    peso = 0

    def cerrar() -> None:
        nonlocal lote, peso
        if not lote:
            return
        valores = ",".join([marca_fila] * len(lote))
        sentencia = f"{cabeza} VALUES {valores}"
        if cola:
            sentencia = f"{sentencia} {cola}"
        parametros: list = []
        for fila in lote:
            parametros.extend(fila)
        lotes.append((sentencia, tuple(parametros)))
        lote = []
        peso = 0

    for fila in filas:
        suma = sum(_peso_parametro(valor) for valor in fila)
        if lote and (
            (len(lote) + 1) * ancho > _MAX_PARAMS_INSERT or peso + suma > _MAX_BYTES_INSERT
        ):
            cerrar()
        lote.append(fila)
        peso += suma
    cerrar()
    return lotes


def _sql_arranque() -> str:
    """Versión, columnas, CREATE de clientes y conteos de semilla, en una lectura."""
    partes = ["SELECT 'meta' AS tipo, clave AS a, valor AS b FROM app_meta"]
    partes.extend(
        f"SELECT 'col' AS tipo, '{tabla}' AS a, name AS b FROM pragma_table_info('{tabla}')"
        for tabla in _TABLAS_ESQUEMA
    )
    partes.append(
        "SELECT 'sql' AS tipo, 'clientes' AS a, ifnull(sql, '') AS b FROM sqlite_master "
        "WHERE type = 'table' AND name = 'clientes'"
    )
    partes.append(
        "SELECT 'n' AS tipo, 'reglas' AS a, CAST(COUNT(*) AS TEXT) AS b FROM clasificacion_reglas"
    )
    partes.append(
        "SELECT 'n' AS tipo, 'tc' AS a, CAST(COUNT(*) AS TEXT) AS b FROM inversiones_tc_bna"
    )
    return " UNION ALL ".join(partes)


def _duda_esquema(columnas: dict, sql_clientes: str) -> bool:
    """True si clientes dice una cosa y el CREATE otra. Vacío de las dos es una base nueva."""
    cols = set((columnas or {}).get("clientes") or ())
    sql = str(sql_clientes or "").strip()
    if cols and not sql:
        return True
    if sql and not cols:
        return True
    return False


def _leer_arranque(conn):
    """Devuelve version, columnas, sql de clientes, marcas, conteos y si la lectura es dudosa.

    Si todavía no están las tablas, marcas y conteos quedan en None y cada
    semilla mira la base por su cuenta. Si la lectura falla, el último valor
    es True: no se sabe el esquema y no se toca `clientes`.
    """
    try:
        filas = conn.execute(_sql_arranque()).fetchall()
    except sqlite3.OperationalError:
        try:
            columnas, sql_clientes = _leer_estado_esquema(conn)
        except sqlite3.OperationalError:
            return None, {}, "", None, None, True
        return None, columnas, sql_clientes, None, None, _duda_esquema(columnas, sql_clientes)
    version = None
    marcas: dict[str, str] = {}
    columnas: dict[str, set[str]] = {tabla: set() for tabla in _TABLAS_ESQUEMA}
    sql_clientes = ""
    conteos = {"reglas": 0, "tc": 0}
    for fila in filas:
        tipo = str(fila[0])
        if tipo == "meta":
            marcas[str(fila[1])] = "" if fila[2] is None else str(fila[2])
            if str(fila[1]) == _CLAVE_SCHEMA:
                version = marcas[str(fila[1])]
        elif tipo == "col" and fila[2]:
            columnas.setdefault(str(fila[1]), set()).add(str(fila[2]))
        elif tipo == "sql":
            sql_clientes = str(fila[2] or "")
        elif tipo == "n":
            conteos[str(fila[1])] = int(fila[2] or 0)
    return version, columnas, sql_clientes, marcas, conteos, _duda_esquema(columnas, sql_clientes)


_TABLAS_ESQUEMA = (
    "clientes",
    "asientos_generados",
    "usuarios_oficina",
    "bank_transactions",
    "proveedores_pendientes",
    "arca_emisores",
    "rutina_pedidos",
    "inversiones_operaciones",
)


def _leer_estado_esquema(conn) -> tuple[dict[str, set[str]], str]:
    """Columnas reales y el CREATE de clientes, en una sola lectura.

    El nombre de tabla va literal (lista fija): pragma_table_info no toma placeholder.
    """
    partes = [
        f"SELECT 'col' AS tipo, '{tabla}' AS a, name AS b FROM pragma_table_info('{tabla}')"
        for tabla in _TABLAS_ESQUEMA
    ]
    partes.append(
        "SELECT 'sql' AS tipo, name AS a, ifnull(sql, '') AS b FROM sqlite_master "
        "WHERE type = 'table' AND name = 'clientes'"
    )
    columnas: dict[str, set[str]] = {tabla: set() for tabla in _TABLAS_ESQUEMA}
    sql_clientes = ""
    for fila in conn.execute(" UNION ALL ".join(partes)).fetchall():
        tipo = str(fila[0])
        if tipo == "col" and fila[2]:
            columnas.setdefault(str(fila[1]), set()).add(str(fila[2]))
        elif tipo == "sql":
            sql_clientes = str(fila[2] or "")
    return columnas, sql_clientes


def _anotar_parte_columna(parte: str, columnas: set[str]) -> None:
    parte = parte.strip()
    if not parte:
        return
    if parte.upper().startswith(("PRIMARY", "UNIQUE", "FOREIGN", "CHECK", "CONSTRAINT")):
        return
    nombre = parte.split()[0].strip('"`[]')
    if nombre:
        columnas.add(nombre)


def _columnas_de_create(sql: str) -> tuple[str, set[str]] | None:
    coincidencia = re.search(
        r"CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?(\w+)\s*\(",
        sql,
        re.IGNORECASE,
    )
    if not coincidencia:
        return None
    tabla = coincidencia.group(1)
    indice = coincidencia.end()
    profundidad = 1
    cuerpo: list[str] = []
    while indice < len(sql) and profundidad:
        caracter = sql[indice]
        if caracter == "(":
            profundidad += 1
            cuerpo.append(caracter)
        elif caracter == ")":
            profundidad -= 1
            if profundidad:
                cuerpo.append(caracter)
        else:
            cuerpo.append(caracter)
        indice += 1
    columnas: set[str] = set()
    buffer: list[str] = []
    profundidad = 0
    for caracter in "".join(cuerpo):
        if caracter == "(":
            profundidad += 1
            buffer.append(caracter)
        elif caracter == ")":
            profundidad -= 1
            buffer.append(caracter)
        elif caracter == "," and profundidad == 0:
            _anotar_parte_columna("".join(buffer), columnas)
            buffer = []
        else:
            buffer.append(caracter)
    _anotar_parte_columna("".join(buffer), columnas)
    return tabla, columnas


class _CursorGrabado:
    def __init__(self, filas: list):
        self._filas = filas
        self._i = 0

    def fetchone(self):
        if self._i >= len(self._filas):
            return None
        fila = self._filas[self._i]
        self._i += 1
        return fila

    def fetchall(self):
        resto = self._filas[self._i :]
        self._i = len(self._filas)
        return resto


class _GrabadorDDL:
    """Corre el mismo DDL de siempre en memoria y junta el SQL que sí hay que mandar."""

    def __init__(self, columnas: dict[str, set[str]], sql_clientes: str):
        self.columnas = {tabla: set(cols) for tabla, cols in columnas.items()}
        self.sql_clientes = sql_clientes
        self.sentencias: list[str] = []

    def execute(self, sql, parametros=()):
        compacto = " ".join(str(sql).split())
        arriba = compacto.upper()
        if arriba.startswith("SELECT SQL FROM SQLITE_MASTER"):
            if not self.sql_clientes:
                return _CursorGrabado([])
            return _CursorGrabado([_FilaCompat(("sql",), (self.sql_clientes,))])
        if arriba.startswith("PRAGMA TABLE_INFO"):
            hallado = re.search(r"TABLE_INFO\(\s*['\"]?(\w+)", compacto, re.IGNORECASE)
            tabla = hallado.group(1) if hallado else ""
            nombres = sorted(self.columnas.get(tabla, ()))
            filas = [
                _FilaCompat(
                    ("cid", "name", "type", "notnull", "dflt_value", "pk"),
                    (i, nombre, "", 0, None, 0),
                )
                for i, nombre in enumerate(nombres)
            ]
            return _CursorGrabado(filas)
        if arriba.startswith("ALTER"):
            hallado = re.search(
                r"ALTER\s+TABLE\s+(\w+)\s+ADD\s+COLUMN\s+(\w+)",
                compacto,
                re.IGNORECASE,
            )
            if hallado:
                tabla, columna = hallado.group(1), hallado.group(2)
                if columna in self.columnas.get(tabla, set()):
                    raise sqlite3.OperationalError(f"duplicate column name: {columna}")
                self.columnas.setdefault(tabla, set()).add(columna)
            self.sentencias.append(compacto)
            return _CursorGrabado([])
        if arriba.startswith("CREATE"):
            anotado = _columnas_de_create(compacto)
            if anotado:
                tabla, nuevas = anotado
                self.columnas.setdefault(tabla, set()).update(nuevas)
            self.sentencias.append(compacto)
            return _CursorGrabado([])
        self.sentencias.append(compacto)
        return _CursorGrabado([])


def _definir_esquema(conn) -> None:
    """El DDL histórico. Sirve contra la base o contra un grabador en memoria."""
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            cuit TEXT NOT NULL UNIQUE,
            tipo_persona TEXT NOT NULL CHECK (
                tipo_persona IN ('Persona Jurídica', 'Persona Física', 'Monotributista')
            ),
            plan_cuentas_path TEXT,
            mes_cierre_balance INTEGER,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    try:
        conn.execute("ALTER TABLE clientes ADD COLUMN mes_cierre_balance INTEGER")
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE clientes ADD COLUMN plan_cuentas_csv TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE clientes ADD COLUMN balance_devengamiento_bytes BLOB")
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE clientes ADD COLUMN balance_devengamiento_nombre TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE clientes ADD COLUMN balance_devengamiento_actualizado TIMESTAMP")
    except sqlite3.OperationalError:
        pass

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS devengamientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            anio INTEGER NOT NULL,
            datos_json TEXT NOT NULL,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id),
            UNIQUE(cliente_id, mes, anio)
        )
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS asientos_generados (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id   INTEGER NOT NULL REFERENCES clientes(id),
            mes          INTEGER NOT NULL,
            anio         INTEGER NOT NULL,
            tipo         TEXT NOT NULL,
            asiento_json TEXT NOT NULL,
            intentos     INTEGER DEFAULT 1,
            estado       TEXT DEFAULT 'Ingresado',
            creado_en    TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    try:
        conn.execute(
            "ALTER TABLE asientos_generados ADD COLUMN estado TEXT DEFAULT 'Ingresado'"
        )
    except Exception:
        pass

    import auth_oficina  # import local: evita ciclo de imports con database.py

    auth_oficina.inicializar_tabla_usuarios_oficina(conn)
    _inicializar_tablas_sueldos(conn)
    _inicializar_tablas_conciliacion(conn)
    import inversiones_db

    inversiones_db.inicializar_tablas_inversiones(conn)
    import arca.persistencia as arca_persistencia

    arca_persistencia.inicializar_tablas_arca(conn)
    import rutinas as rutinas_cola

    rutinas_cola.inicializar_tablas_rutinas(conn)
    _asegurar_app_meta(conn)


def _sembrar_despues_del_ddl(marcas, conteos) -> None:
    import auth_oficina
    import inversiones_db

    if marcas is None or conteos is None:
        auth_oficina.sembrar_usuarios_oficina_default()
        _sembrar_convenios_sueldos_default()
        _reset_cct_comercio_masivo_si_corresponde()
        sembrar_reglas_conciliacion_default()
        inversiones_db.sembrar_tc_bna_default()
        return
    auth_oficina.sembrar_usuarios_oficina_default(
        marca_leida=marcas.get("usuarios_equipo"),
        conocemos_marca=True,
    )
    _sembrar_convenios_sueldos_default(
        marca_leida=marcas.get("convenios_sueldos"),
        conocemos_marca=True,
    )
    _reset_cct_comercio_masivo_si_corresponde(ya_hecho=marcas.get("cct_reset_v1") == "1")
    sembrar_reglas_conciliacion_default(ya_hay=conteos.get("reglas", 0))
    inversiones_db.sembrar_tc_bna_default(ya_hay=conteos.get("tc", 0))


def _sql_destruye_clientes(sql: str) -> bool:
    return _RE_DESTRUYE_CLIENTES.search(" ".join(str(sql).split())) is not None


def _aplicar_ddl(conn, sentencias: list[str]) -> None:
    """CREATE/ALTER idempotentes entre BEGIN y COMMIT. Si algo falla, ROLLBACK.

    ``executescript`` de libsql no deshace lo ya aplicado: un ALTER que revienta
    dejaba a `clientes` reconstruida. Por eso cada sentencia va en la transacción.
    """
    limpias = []
    for sql in sentencias:
        if _sql_destruye_clientes(sql):
            logging.getLogger(__name__).error(
                "Se descartó DDL que reconstruye clientes: %s", sql[:160]
            )
            continue
        limpias.append(sql)
    if not limpias:
        return
    try:
        conn.execute("BEGIN")
        for sql in limpias:
            conn.execute(sql)
        conn.execute("COMMIT")
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise


def _marcar_modulos_listos() -> None:
    """Las pantallas no vuelven a correr el DDL de rutinas ni de ARCA en cada lectura."""
    # import local: rutinas y arca importan database (ciclo).
    import rutinas as rutinas_cola
    import arca.persistencia as arca_persistencia

    rutinas_cola.marcar_tablas_listas()
    arca_persistencia.marcar_tablas_listas()


def inicializar_bd() -> None:
    """Crea tablas y siembras. Con el esquema ya marcado, una sola lectura.

    En una réplica, primero `sync()`. Si ese sync falla o el esquema queda
    inconsistente, no se manda DDL: una réplica nueva y vacía no es una base nueva.
    """
    global _ddl_bloqueado
    conn = obtener_conexion()
    if isinstance(conn, _ConexionCompatTurso) and _turso_es_replica:
        if not sincronizar_replica_una_vez():
            _ddl_bloqueado = True
            logging.getLogger(__name__).error(
                "Réplica sin sincronizar: no se lee el esquema ni se reconstruye clientes."
            )
            return
    else:
        _ddl_bloqueado = False
    version, columnas, sql_clientes, marcas, conteos, duda = _leer_arranque(conn)
    if duda:
        _ddl_bloqueado = True
        logging.getLogger(__name__).error(
            "Lectura de esquema dudosa: no se aplica DDL ni se reconstruye clientes."
        )
        return
    _ddl_bloqueado = False
    if version == SCHEMA_VERSION:
        _marcar_modulos_listos()
        return
    grabador = _GrabadorDDL(columnas, sql_clientes)
    _definir_esquema(grabador)
    sentencias = [
        sql
        for sql in grabador.sentencias
        if "idx_rutina_pedidos_abierto" not in sql and not _sql_destruye_clientes(sql)
    ]
    _aplicar_ddl(conn, sentencias)
    try:
        conn.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_rutina_pedidos_abierto
            ON rutina_pedidos(rutina)
            WHERE estado IN ('PENDIENTE', 'EN_CURSO')
            """
        )
    except sqlite3.OperationalError:
        pass
    _sembrar_despues_del_ddl(marcas, conteos)
    guardar_semilla(_CLAVE_SCHEMA, SCHEMA_VERSION)
    _marcar_modulos_listos()


def _reglas_cct_basicas() -> dict:
    from cct_escalas import reglas_comercio_julio_2026

    return reglas_comercio_julio_2026()


def _sembrar_convenios_sueldos_default(
    *,
    marca_leida: str | None = None,
    conocemos_marca: bool = False,
) -> None:
    """Catálogo inicial de CCTs. Comercio trae escala FAECYS julio 2026."""
    from cct_escalas import reglas_comercio_julio_2026

    basicas = {
        "antiguedadPorAnioPct": 0.01,
        "presentismoDivisor": 12,
        "horasMensuales": 200,
        "horasExtras50Multiplicador": 1.5,
        "diasMes": 30,
        "jubilacionPct": 0.11,
        "pamiPct": 0.03,
        "obraSocialPct": 0.03,
        "aporteSindicalPct": 0.02,
        "usarEscalaCct": True,
        "escalas": {},
    }
    catalogo = [
        (
            "COMERCIO_130_75",
            "CCT 130/75 Empleados de Comercio",
            reglas_comercio_julio_2026(),
        ),
        ("UOCRA_76", "CCT UOCRA Construcción", basicas),
        ("GASTRONOMICOS_389_04", "CCT 389/04 Gastronómicos", basicas),
        ("SANIDAD_122_75", "CCT 122/75 Sanidad", basicas),
        ("METALURGICOS_260_75", "CCT 260/75 Metalúrgicos", basicas),
        ("OTRO", "Otro / a definir", basicas),
    ]
    filas: list[tuple[str, str, str]] = []
    for codigo, nombre, reglas in catalogo:
        payload = json.dumps(reglas, ensure_ascii=False, sort_keys=True)
        filas.append((codigo, nombre, payload))
    marca = huella_semilla(*(f"{codigo}|{nombre}|{payload}" for codigo, nombre, payload in filas))
    if conocemos_marca and marca_leida == marca:
        return
    with obtener_conexion() as conn:
        if not conocemos_marca and leer_semilla("convenios_sueldos", conn) == marca:
            return
        conn.executemany(
            """
            INSERT INTO convenios_colectivos (codigo, nombre, reglas_json)
            VALUES (?, ?, ?)
            ON CONFLICT(codigo) DO UPDATE SET
                nombre = excluded.nombre,
                reglas_json = excluded.reglas_json
            """,
            filas,
        )
        guardar_semilla("convenios_sueldos", marca, conn)
        conn.commit()


def actualizar_reglas_convenio(codigo: str, reglas: dict) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE convenios_colectivos SET reglas_json = ? WHERE codigo = ?",
            (json.dumps(reglas, ensure_ascii=False), codigo),
        )
        conn.commit()


def _reset_cct_comercio_masivo_si_corresponde(*, ya_hecho: bool = False) -> None:
    """
    Una sola vez: si casi todos quedaron en COMERCIO por el default erróneo
    de la migración, se limpian para forzar asignación real por sociedad.
    """
    if ya_hecho:
        return
    try:
        with obtener_conexion() as conn:
            if leer_semilla("cct_reset_v1", conn) == "1":
                return
            total = int(_valor_fila(conn.execute("SELECT COUNT(*) AS n FROM clientes").fetchone(), "n") or 0)
            comercio = int(
                _valor_fila(
                    conn.execute(
                        "SELECT COUNT(*) AS n FROM clientes WHERE cct_asignado = 'COMERCIO_130_75'"
                    ).fetchone(),
                    "n",
                )
                or 0
            )
            if total > 0 and comercio >= max(1, int(total * 0.8)):
                conn.execute(
                    "UPDATE clientes SET cct_asignado = NULL "
                    "WHERE cct_asignado = 'COMERCIO_130_75'"
                )
            guardar_semilla("cct_reset_v1", "1", conn)
            conn.commit()
    except Exception:
        pass


def _inicializar_tablas_sueldos(conn: sqlite3.Connection) -> None:
    """Tablas de liquidación de sueldos (legajos, novedades, resultados)."""
    try:
        # Sin CCT por defecto: cada sociedad se asigna a mano (no asumir Comercio).
        conn.execute("ALTER TABLE clientes ADD COLUMN cct_asignado TEXT")
    except sqlite3.OperationalError:
        pass

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS convenios_colectivos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL UNIQUE,
            nombre TEXT NOT NULL,
            reglas_json TEXT NOT NULL DEFAULT '{}',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS empleados_sueldos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            cuil TEXT NOT NULL,
            nombre TEXT NOT NULL,
            categoria TEXT NOT NULL DEFAULT '',
            sueldo_basico REAL NOT NULL DEFAULT 0,
            fecha_ingreso TEXT NOT NULL,
            antiguedad_anios INTEGER NOT NULL DEFAULT 0,
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cliente_id, cuil)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS novedades_buzon (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            empleado_id INTEGER NOT NULL REFERENCES empleados_sueldos(id) ON DELETE CASCADE,
            periodo TEXT NOT NULL,
            dias_ausencia INTEGER NOT NULL DEFAULT 0,
            horas_extras_50 REAL NOT NULL DEFAULT 0,
            no_remunerativo_extra REAL NOT NULL DEFAULT 0,
            estado TEXT NOT NULL DEFAULT 'Recibida',
            enviada_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cliente_id, empleado_id, periodo)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS liquidaciones_resultado (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
            empleado_id INTEGER NOT NULL REFERENCES empleados_sueldos(id) ON DELETE CASCADE,
            periodo TEXT NOT NULL,
            total_remunerativo REAL NOT NULL,
            total_no_remunerativo REAL NOT NULL,
            total_descuentos REAL NOT NULL,
            neto_a_percibir REAL NOT NULL,
            detalle_conceptos_json TEXT NOT NULL DEFAULT '[]',
            liquidado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cliente_id, empleado_id, periodo)
        )
        """
    )


# --- Usuarios de oficina (login/roles) -----------------------------------
# La logica vive en auth_oficina.py (ver ese modulo). Estos wrappers quedan
# acá para no tener que tocar cada call site de app.py en este primer corte
# del refactor; el siguiente paso natural es que app.py importe auth_oficina
# directamente y estos wrappers se puedan borrar.

def _aplicar_usuarios_desde_secrets() -> int:
    import auth_oficina

    return auth_oficina._aplicar_usuarios_desde_secrets()


def listar_usuarios_oficina(solo_activos: bool = True) -> list[dict]:
    import auth_oficina

    return auth_oficina.listar_usuarios_oficina(solo_activos=solo_activos)


def obtener_usuario_oficina(usuario: str) -> dict | None:
    import auth_oficina

    return auth_oficina.obtener_usuario_oficina(usuario)


def crear_usuario_oficina(usuario: str, nombre: str, *, pin: str = "", es_admin: bool = False) -> int:
    import auth_oficina

    return auth_oficina.crear_usuario_oficina(usuario, nombre, pin=pin, es_admin=es_admin)


def actualizar_usuario_oficina(
    usuario_id: int,
    *,
    nombre: str | None = None,
    pin: str | None = None,
    es_admin: bool | None = None,
    activo: bool | None = None,
) -> None:
    import auth_oficina

    auth_oficina.actualizar_usuario_oficina(
        usuario_id, nombre=nombre, pin=pin, es_admin=es_admin, activo=activo
    )


def verificar_login_oficina(usuario: str, pin: str = "") -> dict | None:
    import auth_oficina

    return auth_oficina.verificar_login_oficina(usuario, pin)


def verificar_admin_oficina(usuario_admin: str, pin_admin: str = "") -> dict | None:
    import auth_oficina

    return auth_oficina.verificar_admin_oficina(usuario_admin, pin_admin)


def resetear_pin_usuario_oficina(
    usuario_objetivo: str, nuevo_pin: str, *, usuario_admin: str, pin_admin: str = ""
) -> dict:
    import auth_oficina

    return auth_oficina.resetear_pin_usuario_oficina(
        usuario_objetivo, nuevo_pin, usuario_admin=usuario_admin, pin_admin=pin_admin
    )


def usuario_bloqueado_oficina(usuario: str) -> int:
    import auth_oficina

    return auth_oficina.usuario_bloqueado_oficina(usuario)


def cargar_cuits_auxiliares(
    ruta: str | Path | None = None,
) -> dict[str, str]:
    """
    Lee Cuits_Auxiliares.xlsx y devuelve un diccionario nombre_normalizado -> CUIT.
    Columna B = nombre, Columna C = CUIT.
    """
    ruta_final = Path(ruta) if ruta else CUITS_AUXILIARES_PATH
    if not ruta_final.exists():
        return {}

    wb = openpyxl.load_workbook(ruta_final, read_only=True, data_only=True)
    ws = wb.active
    mapeo: dict[str, str] = {}

    for row in ws.iter_rows(min_row=1, max_col=3, values_only=True):
        nombre, cuit = row[1] if len(row) > 1 else None, row[2] if len(row) > 2 else None
        if not nombre or not cuit:
            continue
        cuit_limpio = re.sub(r"\D", "", str(cuit))
        if len(cuit_limpio) != 11:
            continue
        clave = _normalizar_nombre(str(nombre))
        if clave and clave not in ("CUIT", "SOCIEDADES", "RESP INSCRIPTOS", "MONOTRIBUTISTA", "SS DOMESTICO"):
            mapeo[clave] = cuit_limpio

    wb.close()
    return mapeo


def _buscar_cuit_en_auxiliares(nombre: str, auxiliares: dict[str, str]) -> Optional[str]:
    """Busca CUIT en auxiliares con coincidencia exacta o parcial."""
    clave = _normalizar_nombre(nombre)
    if clave in auxiliares:
        return auxiliares[clave]

    for clave_aux, cuit in auxiliares.items():
        if clave in clave_aux or clave_aux in clave:
            return cuit
        # Coincidencia por tokens significativos
        tokens = {t for t in clave.split() if len(t) > 3}
        tokens_aux = {t for t in clave_aux.split() if len(t) > 3}
        if tokens and tokens_aux and len(tokens & tokens_aux) >= min(2, len(tokens)):
            return cuit

    return None


def _siguiente_cuit_temporal(conn: sqlite3.Connection) -> str:
    """Genera el próximo CUIT correlativo temporal (99XXXXXXXXX)."""
    fila = conn.execute(
        "SELECT MAX(CAST(cuit AS INTEGER)) FROM clientes WHERE cuit LIKE ?",
        (f"{PREFIJO_CUIT_TEMPORAL}%",),
    ).fetchone()
    max_actual = int(fila[0]) if fila and fila[0] else CUIT_TEMPORAL_INICIO - 1
    return str(max(max_actual + 1, CUIT_TEMPORAL_INICIO))


def actualizar_cuits_desde_auxiliares(
    ruta: str | Path | None = None,
) -> dict[str, int]:
    """
    Actualiza CUITs de clientes existentes e inserta faltantes desde Lista_empresas.
    Usa Cuits_Auxiliares.xlsx; asigna CUIT temporal correlativo si no hay match.
    """
    inicializar_bd()
    auxiliares = cargar_cuits_auxiliares(ruta)
    stats = {"actualizados": 0, "insertados": 0, "temporales": 0, "sin_cambios": 0}

    with obtener_conexion() as conn:
        clientes = conn.execute("SELECT id, nombre, cuit FROM clientes").fetchall()
        nombres_en_bd = {_normalizar_nombre(c["nombre"]): dict(c) for c in clientes}
        cuits_usados = {str(c["cuit"]).replace("-", "") for c in clientes}

        for cliente in clientes:
            cuit_real = _buscar_cuit_en_auxiliares(cliente["nombre"], auxiliares)
            cuit_actual = str(cliente["cuit"]).replace("-", "")

            if cuit_real and cuit_real != cuit_actual and cuit_real not in cuits_usados:
                conn.execute(
                    "UPDATE clientes SET cuit = ? WHERE id = ?",
                    (cuit_real, cliente["id"]),
                )
                cuits_usados.discard(cuit_actual)
                cuits_usados.add(cuit_real)
                stats["actualizados"] += 1
            else:
                stats["sin_cambios"] += 1

        # Insertar empresas de Lista_empresas que no estén en BD
        if LISTA_EMPRESAS_PATH.exists():
            df = pd.read_excel(LISTA_EMPRESAS_PATH)
            for nombre in df["Nombre"].dropna().astype(str).unique():
                clave = _normalizar_nombre(nombre)
                if clave in nombres_en_bd:
                    continue

                cuit = _buscar_cuit_en_auxiliares(nombre, auxiliares)
                if not cuit or cuit in cuits_usados:
                    cuit = _siguiente_cuit_temporal(conn)
                    stats["temporales"] += 1

                tipo = _categorizar_tipo(nombre)
                mes_cierre = 12 if tipo == "Persona Física" else buscar_mes_cierre_web(cuit)
                try:
                    conn.execute(
                        """
                        INSERT INTO clientes (nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance)
                        VALUES (?, ?, ?, ?, ?)
                        """,
                        (nombre.strip(), cuit, tipo, None, mes_cierre),
                    )
                    cuits_usados.add(cuit)
                    nombres_en_bd[clave] = {"nombre": nombre, "cuit": cuit}
                    stats["insertados"] += 1
                except sqlite3.IntegrityError:
                    pass

        conn.commit()

    return stats


def crear_cliente(
    nombre: str,
    cuit: str,
    tipo_persona: str,
    plan_cuentas_path: Optional[str] = None,
    mes_cierre_balance: Optional[int] = None,
) -> int:
    """Registra un nuevo cliente y devuelve su ID."""
    if tipo_persona not in TIPOS_PERSONA:
        raise ValueError(f"Tipo de persona inválido: {tipo_persona}")

    if tipo_persona in ("Persona Física", "Monotributista"):
        mes_cierre_balance = 12
    elif tipo_persona == "Persona Jurídica" and not mes_cierre_balance:
        mes_cierre_balance = buscar_mes_cierre_web(cuit)

    with obtener_conexion() as conn:
        cursor = conn.execute(
            """
            INSERT INTO clientes (nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance)
            VALUES (?, ?, ?, ?, ?)
            """,
            (nombre.strip(), cuit.strip(), tipo_persona, plan_cuentas_path, mes_cierre_balance),
        )
        conn.commit()
        return int(cursor.lastrowid)


def _listar_clientes_directo() -> list[dict]:
    with obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT * FROM clientes ORDER BY nombre COLLATE NOCASE"
        ).fetchall()
    return [dict(fila) for fila in filas]


def listar_clientes() -> list[dict]:
    """Devuelve todos los clientes ordenados por nombre. Cache de pantalla con TTL."""
    # import local: cache_lecturas importa database (ciclo).
    import cache_lecturas

    return [dict(fila) for fila in cache_lecturas.clientes_de_pantalla()]


def _obtener_cliente_directo(cliente_id: int) -> Optional[dict]:
    with obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT * FROM clientes WHERE id = ?", (cliente_id,)
        ).fetchone()
    return dict(fila) if fila else None


def obtener_cliente(cliente_id: int) -> Optional[dict]:
    """Obtiene un cliente por ID. Cache de pantalla con TTL."""
    if cliente_id is None:
        return None
    # import local: cache_lecturas importa database (ciclo).
    import cache_lecturas

    fila = cache_lecturas.cliente_de_pantalla(int(cliente_id))
    return dict(fila) if fila else None


def actualizar_cliente(
    cliente_id: int,
    nombre: str,
    cuit: str,
    tipo_persona: str,
    plan_cuentas_path: Optional[str] = None,
    mes_cierre_balance: Optional[int] = None,
) -> None:
    """Actualiza los datos de un cliente existente."""
    if tipo_persona not in TIPOS_PERSONA:
        raise ValueError(f"Tipo de persona inválido: {tipo_persona}")

    if tipo_persona in ("Persona Física", "Monotributista"):
        mes_cierre_balance = 12
    elif tipo_persona == "Persona Jurídica" and not mes_cierre_balance:
        mes_cierre_balance = buscar_mes_cierre_web(cuit)

    with obtener_conexion() as conn:
        conn.execute(
            """
            UPDATE clientes
            SET nombre = ?, cuit = ?, tipo_persona = ?, plan_cuentas_path = ?, mes_cierre_balance = ?
            WHERE id = ?
            """,
            (nombre.strip(), cuit.strip(), tipo_persona, plan_cuentas_path, mes_cierre_balance, cliente_id),
        )
        conn.commit()


def guardar_plan_cuentas_csv(cliente_id: int, csv_text: str) -> None:
    """Guarda las cuentas del plan en SQLite (sobrevive si el Excel se pierde)."""
    texto = str(csv_text or "").strip()
    if not texto:
        return
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE clientes SET plan_cuentas_csv = ? WHERE id = ?",
            (texto, int(cliente_id)),
        )
        conn.commit()


def plan_cuentas_csv_cliente(cliente_id: int) -> str:
    with obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT plan_cuentas_csv FROM clientes WHERE id = ?",
            (int(cliente_id),),
        ).fetchone()
    if not fila:
        return ""
    return str(fila["plan_cuentas_csv"] or "").strip()


def guardar_balance_devengamiento(cliente_id: int, nombre: str, contenido: bytes) -> None:
    """Guarda el Excel de Balance de Devengamiento en SQLite (evita resubirlo cada vez)."""
    datos = bytes(contenido or b"")
    if not datos:
        return
    with obtener_conexion() as conn:
        conn.execute(
            """
            UPDATE clientes
            SET balance_devengamiento_bytes = ?,
                balance_devengamiento_nombre = ?,
                balance_devengamiento_actualizado = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (datos, str(nombre or "balance.xlsx"), int(cliente_id)),
        )
        conn.commit()


def balance_devengamiento_cliente(cliente_id: int) -> dict | None:
    """Devuelve {"nombre", "contenido", "actualizado"} o None si no hay nada guardado."""
    with obtener_conexion() as conn:
        fila = conn.execute(
            """
            SELECT balance_devengamiento_bytes, balance_devengamiento_nombre,
                   balance_devengamiento_actualizado
            FROM clientes WHERE id = ?
            """,
            (int(cliente_id),),
        ).fetchone()
    if not fila or not fila["balance_devengamiento_bytes"]:
        return None
    return {
        "nombre": str(fila["balance_devengamiento_nombre"] or "balance.xlsx"),
        "contenido": bytes(fila["balance_devengamiento_bytes"]),
        "actualizado": str(fila["balance_devengamiento_actualizado"] or ""),
    }


def _plan_archivo_con_cuentas(ruta: Path | None) -> bool:
    """True si el archivo tiene al menos una cuenta (no plantilla Tango vacía)."""
    if ruta is None:
        return False
    ruta = Path(ruta)
    if not ruta.is_file():
        return False
    nombre = ruta.name.lower()
    if nombre in ("plan_default.xlsx", "plan_default.xls") or "cuentas contables (4)" in nombre:
        return False
    suf = ruta.suffix.lower()
    if suf == ".csv":
        try:
            lineas = [ln for ln in ruta.read_text(encoding="utf-8").splitlines() if ln.strip()]
        except OSError:
            return False
        return len(lineas) >= 2
    if suf in {".xlsx", ".xls"}:
        try:
            with pd.ExcelFile(ruta) as xl:
                for hoja in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name=hoja, dtype=str)
                    if df is None or df.empty:
                        continue
                    nonempty = df.dropna(how="all")
                    if nonempty.empty:
                        continue
                    col0 = nonempty.iloc[:, 0].astype(str).str.strip()
                    col0 = col0[
                        ~col0.str.lower().isin({"", "nan", "none", "codigo", "código", "cuenta"})
                    ]
                    if bool(col0.ne("").any()):
                        return True
        except Exception:
            return False
        return False
    return False


def _texto_plan_csv_disco(cuit: str, plan_path: str | None = None) -> str:
    candidatos: list[Path] = []
    if plan_path and str(plan_path).lower().endswith(".csv"):
        candidatos.append(Path(plan_path))
    candidatos.append(DATA_PLANES_DIR / f"plan_{cuit}.csv")
    vistos: set[str] = set()
    for cand in candidatos:
        key = str(cand).lower()
        if key in vistos:
            continue
        vistos.add(key)
        if not cand.is_file():
            continue
        try:
            texto = cand.read_text(encoding="utf-8").strip()
        except OSError:
            continue
        if texto.count("\n") >= 1:
            return texto
    return ""


def _resolver_plan_path_catalogo(item: dict, cuit: str) -> str | None:
    """Ruta de plan propio para catálogo. No usa plan_default ni plantillas vacías."""
    explicit = str(item.get("plan_cuentas") or item.get("plan_cuentas_path") or "").strip()
    if explicit:
        cand = Path(explicit)
        if not cand.is_absolute():
            cand = BASE_DIR / cand
        if _plan_archivo_con_cuentas(cand):
            return str(cand)
    propio_csv = DATA_PLANES_DIR / f"plan_{cuit}.csv"
    if _plan_archivo_con_cuentas(propio_csv):
        return str(propio_csv)
    propio = DATA_PLANES_DIR / f"plan_{cuit}.xlsx"
    if _plan_archivo_con_cuentas(propio):
        return str(propio)
    return None


def _fila_catalogo(item: dict) -> tuple[str, str, str, int] | None:
    """Normaliza un ítem de catálogo. None si no se puede insertar."""
    nombre = str(item.get("nombre", "")).strip()
    cuit = re.sub(r"\D", "", str(item.get("cuit", "")))
    tipo = str(item.get("tipo") or item.get("tipo_persona") or "Monotributista").strip()
    if tipo == "Monotributista":
        tipo_persona = "Monotributista"
    elif tipo in TIPOS_PERSONA:
        tipo_persona = tipo
    else:
        tipo_persona = _categorizar_tipo(nombre)
    if not nombre or len(cuit) != 11:
        return None
    mes_raw = item.get("mes_cierre_balance")
    try:
        mes_cierre = int(mes_raw) if mes_raw not in (None, "") else 12
    except (TypeError, ValueError):
        mes_cierre = 12
    if mes_cierre < 1 or mes_cierre > 12:
        mes_cierre = 12
    if tipo_persona in ("Persona Física", "Monotributista"):
        mes_cierre = 12
    return nombre, cuit, tipo_persona, mes_cierre


def _huella_catalogo(catalogo: list[dict]) -> str:
    partes = []
    for item in catalogo:
        nombre = str(item.get("nombre", "")).strip()
        cuit = re.sub(r"\D", "", str(item.get("cuit", "")))
        tipo = str(item.get("tipo") or item.get("tipo_persona") or "").strip()
        mes = item.get("mes_cierre_balance")
        partes.append(f"{nombre}|{cuit}|{tipo}|{mes}")
    return huella_semilla(*partes)


def sincronizar_clientes_catalogo(
    catalogo: list[dict],
    *,
    semilla: str | None = None,
) -> dict[str, int]:
    """Inserta clientes del catálogo estático que aún no existen (por CUIT).

    Con `semilla`, la segunda corrida sale si el contenido no cambió. Esa marca
    no mira planes en disco: un plan nuevo se vincula al cambiar el catálogo
    o al llamar sin semilla.
    """
    inicializar_bd()
    stats = {"insertados": 0, "omitidos": 0, "errores": 0, "planes_vinculados": 0}
    marca = _huella_catalogo(catalogo) if semilla else ""
    if semilla and leer_semilla(f"catalogo:{semilla}") == marca:
        stats["omitidos"] = len(catalogo)
        return stats
    with obtener_conexion() as conn:
        existentes = {
            re.sub(r"\D", "", str(row["cuit"])): {
                "id": row["id"],
                "cuit": row["cuit"],
                "plan_cuentas_path": row["plan_cuentas_path"],
                "plan_cuentas_csv": row["plan_cuentas_csv"],
            }
            for row in conn.execute(
                "SELECT id, cuit, plan_cuentas_path, plan_cuentas_csv FROM clientes"
            ).fetchall()
        }
        inserts: list[tuple] = []
        upd_path_csv: list[tuple] = []
        upd_path: list[tuple] = []
        upd_csv: list[tuple] = []
        vistos = set(existentes)
        for item in catalogo:
            fila = _fila_catalogo(item)
            if fila is None:
                stats["errores"] += 1
                continue
            nombre, cuit, tipo_persona, mes_cierre = fila
            plan_path = _resolver_plan_path_catalogo(item, cuit)
            csv_txt = _texto_plan_csv_disco(cuit, plan_path)
            if cuit in vistos:
                row = existentes.get(cuit)
                if row is not None and row.get("id") is not None:
                    actual = Path(str(row.get("plan_cuentas_path") or ""))
                    csv_bd = str(row.get("plan_cuentas_csv") or "").strip()
                    if plan_path and (
                        not _plan_archivo_con_cuentas(actual) or not csv_bd
                    ):
                        if csv_txt and not csv_bd:
                            upd_path_csv.append((plan_path, csv_txt, row["id"]))
                        else:
                            upd_path.append((plan_path, row["id"]))
                        stats["planes_vinculados"] += 1
                    elif csv_txt and not csv_bd:
                        upd_csv.append((csv_txt, row["id"]))
                        stats["planes_vinculados"] += 1
                stats["omitidos"] += 1
                continue
            vistos.add(cuit)
            inserts.append((nombre, cuit, tipo_persona, plan_path, mes_cierre, csv_txt or None))
            existentes[cuit] = {
                "id": None,
                "cuit": cuit,
                "plan_cuentas_path": plan_path,
                "plan_cuentas_csv": csv_txt,
            }
            stats["insertados"] += 1
        if inserts:
            conn.executemany(
                """
                INSERT INTO clientes (
                    nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance, plan_cuentas_csv
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                inserts,
            )
        if upd_path_csv:
            conn.executemany(
                """
                UPDATE clientes
                SET plan_cuentas_path = ?, plan_cuentas_csv = ?
                WHERE id = ?
                """,
                upd_path_csv,
            )
        if upd_path:
            conn.executemany(
                "UPDATE clientes SET plan_cuentas_path = ? WHERE id = ?",
                upd_path,
            )
        if upd_csv:
            conn.executemany(
                "UPDATE clientes SET plan_cuentas_csv = ? WHERE id = ?",
                upd_csv,
            )
        if semilla:
            guardar_semilla(f"catalogo:{semilla}", marca, conn)
        conn.commit()
    return stats


def cargar_seed_sociedades_pj(ruta: str | Path | None = None) -> dict[str, int]:
    """
    Siembra Personas Jurídicas desde data/seed/sociedades_pj.json (repo / Cloud).
    Vincula plan_{cuit}.xlsx de data/planes_cuentas cuando exista.
    Si el JSON no existe o es inválido: no-op (warning) sin levantar excepción.
    """
    path = Path(ruta) if ruta else SEED_SOCIEDADES_PJ_PATH
    if not path.is_file():
        msg = f"Seed PJ omitido: no existe {path}"
        warnings.warn(msg, UserWarning, stacklevel=2)
        logging.getLogger(__name__).warning(msg)
        return {"insertados": 0, "omitidos": 0, "errores": 0, "planes_vinculados": 0, "sin_archivo": 1}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        msg = f"Seed PJ omitido: no se pudo leer {path}: {exc}"
        warnings.warn(msg, UserWarning, stacklevel=2)
        logging.getLogger(__name__).warning(msg)
        return {"insertados": 0, "omitidos": 0, "errores": 1, "planes_vinculados": 0}
    if not isinstance(data, list):
        msg = f"Seed PJ omitido: {path} no es una lista JSON"
        warnings.warn(msg, UserWarning, stacklevel=2)
        logging.getLogger(__name__).warning(msg)
        return {"insertados": 0, "omitidos": 0, "errores": 1, "planes_vinculados": 0}
    return sincronizar_clientes_catalogo(
        data,
        semilla="sociedades_pj" if ruta is None else None,
    )


def eliminar_cliente(cliente_id: int) -> None:
    """Elimina un cliente por ID."""
    with obtener_conexion() as conn:
        conn.execute("DELETE FROM devengamientos WHERE cliente_id = ?", (cliente_id,))
        conn.execute("DELETE FROM clientes WHERE id = ?", (cliente_id,))
        conn.commit()


def guardar_devengamiento(cliente_id: int, mes: int, anio: int, datos: dict) -> None:
    """Guarda o actualiza los datos de devengamiento de un cliente para un período."""
    with obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO devengamientos (cliente_id, mes, anio, datos_json)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(cliente_id, mes, anio) DO UPDATE SET
                datos_json = excluded.datos_json,
                creado_en = CURRENT_TIMESTAMP
            """,
            (cliente_id, mes, anio, json.dumps(datos, ensure_ascii=False)),
        )
        conn.commit()


def obtener_devengamiento(cliente_id: int, mes: int, anio: int) -> Optional[dict]:
    """Obtiene datos de devengamiento guardados para un período."""
    with obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT datos_json FROM devengamientos WHERE cliente_id = ? AND mes = ? AND anio = ?",
            (cliente_id, mes, anio),
        ).fetchone()
    if not fila:
        return None
    return json.loads(fila["datos_json"])


def listar_devengamientos(cliente_id: int) -> list[dict]:
    """Lista historial de devengamientos de un cliente."""
    with obtener_conexion() as conn:
        filas = conn.execute(
            """
            SELECT id, mes, anio, creado_en
            FROM devengamientos
            WHERE cliente_id = ?
            ORDER BY anio DESC, mes DESC
            """,
            (cliente_id,),
        ).fetchall()
    return [dict(f) for f in filas]


def guardar_asiento_generado(
    cliente_id: int,
    mes: int,
    anio: int,
    tipo: str,
    asiento_json: str,
    intentos: int = 1,
    estado: str = "Ingresado",
) -> None:
    """Persiste un asiento generado por IA para un cliente."""
    with obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO asientos_generados (cliente_id, mes, anio, tipo, asiento_json, intentos, estado)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (cliente_id, mes, anio, tipo, asiento_json, intentos, estado),
        )
        conn.commit()


def listar_asientos_generados(cliente_id: int) -> list[dict]:
    """Lista asientos generados por IA para un cliente, del más reciente al más antiguo."""
    with obtener_conexion() as conn:
        rows = conn.execute(
            """
            SELECT id, mes, anio, tipo, intentos, estado, creado_en
            FROM asientos_generados
            WHERE cliente_id = ?
            ORDER BY creado_en DESC
            """,
            (cliente_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def eliminar_asiento_generado(asiento_id: int) -> None:
    """Elimina físicamente un asiento generado por IA. Sin restricciones por estado."""
    with obtener_conexion() as conn:
        conn.execute("DELETE FROM asientos_generados WHERE id = ?", (asiento_id,))
        conn.commit()


# ─── Liquidación de sueldos ─────────────────────────────────────────────────


def listar_convenios() -> list[dict]:
    with obtener_conexion() as conn:
        rows = conn.execute(
            "SELECT id, codigo, nombre, reglas_json FROM convenios_colectivos ORDER BY codigo"
        ).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        try:
            d["reglas"] = json.loads(d.pop("reglas_json") or "{}")
        except json.JSONDecodeError:
            d["reglas"] = {}
        out.append(d)
    return out


def obtener_convenio(codigo: str) -> dict | None:
    with obtener_conexion() as conn:
        row = conn.execute(
            "SELECT id, codigo, nombre, reglas_json FROM convenios_colectivos WHERE codigo = ?",
            (codigo,),
        ).fetchone()
    if not row:
        return None
    d = dict(row)
    try:
        d["reglas"] = json.loads(d.pop("reglas_json") or "{}")
    except json.JSONDecodeError:
        d["reglas"] = {}
    return d


def actualizar_cct_cliente(cliente_id: int, cct_codigo: str) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE clientes SET cct_asignado = ? WHERE id = ?",
            (cct_codigo, cliente_id),
        )
        conn.commit()


def listar_empleados_sueldos(cliente_id: int, solo_activos: bool = True) -> list[dict]:
    with obtener_conexion() as conn:
        if solo_activos:
            rows = conn.execute(
                """
                SELECT * FROM empleados_sueldos
                WHERE cliente_id = ? AND activo = 1
                ORDER BY nombre COLLATE NOCASE
                """,
                (cliente_id,),
            ).fetchall()
        else:
            rows = conn.execute(
                """
                SELECT * FROM empleados_sueldos
                WHERE cliente_id = ?
                ORDER BY nombre COLLATE NOCASE
                """,
                (cliente_id,),
            ).fetchall()
    return [dict(r) for r in rows]


def upsert_empleado_sueldo(
    cliente_id: int,
    cuil: str,
    nombre: str,
    categoria: str,
    sueldo_basico: float,
    fecha_ingreso: str,
    antiguedad_anios: int,
    activo: bool = True,
    empleado_id: int | None = None,
) -> int:
    with obtener_conexion() as conn:
        if empleado_id:
            conn.execute(
                """
                UPDATE empleados_sueldos SET
                    cuil = ?, nombre = ?, categoria = ?, sueldo_basico = ?,
                    fecha_ingreso = ?, antiguedad_anios = ?, activo = ?
                WHERE id = ? AND cliente_id = ?
                """,
                (
                    cuil,
                    nombre,
                    categoria,
                    float(sueldo_basico),
                    fecha_ingreso,
                    int(antiguedad_anios),
                    1 if activo else 0,
                    empleado_id,
                    cliente_id,
                ),
            )
            conn.commit()
            return int(empleado_id)
        cur = conn.execute(
            """
            INSERT INTO empleados_sueldos (
                cliente_id, cuil, nombre, categoria, sueldo_basico,
                fecha_ingreso, antiguedad_anios, activo
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(cliente_id, cuil) DO UPDATE SET
                nombre = excluded.nombre,
                categoria = excluded.categoria,
                sueldo_basico = excluded.sueldo_basico,
                fecha_ingreso = excluded.fecha_ingreso,
                antiguedad_anios = excluded.antiguedad_anios,
                activo = excluded.activo
            """,
            (
                cliente_id,
                cuil,
                nombre,
                categoria,
                float(sueldo_basico),
                fecha_ingreso,
                int(antiguedad_anios),
                1 if activo else 0,
            ),
        )
        conn.commit()
        if cur.lastrowid:
            return int(cur.lastrowid)
        row = conn.execute(
            "SELECT id FROM empleados_sueldos WHERE cliente_id = ? AND cuil = ?",
            (cliente_id, cuil),
        ).fetchone()
        return int(row["id"])


def eliminar_empleado_sueldo(empleado_id: int) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE empleados_sueldos SET activo = 0 WHERE id = ?", (empleado_id,)
        )
        conn.commit()


def upsert_novedad_buzon(
    cliente_id: int,
    empleado_id: int,
    periodo: str,
    dias_ausencia: int = 0,
    horas_extras_50: float = 0,
    no_remunerativo_extra: float = 0,
) -> int:
    with obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO novedades_buzon (
                cliente_id, empleado_id, periodo, dias_ausencia,
                horas_extras_50, no_remunerativo_extra, estado, enviada_en
            ) VALUES (?, ?, ?, ?, ?, ?, 'Recibida', CURRENT_TIMESTAMP)
            ON CONFLICT(cliente_id, empleado_id, periodo) DO UPDATE SET
                dias_ausencia = excluded.dias_ausencia,
                horas_extras_50 = excluded.horas_extras_50,
                no_remunerativo_extra = excluded.no_remunerativo_extra,
                estado = 'Recibida',
                enviada_en = CURRENT_TIMESTAMP
            """,
            (
                cliente_id,
                empleado_id,
                periodo,
                int(dias_ausencia),
                float(horas_extras_50),
                float(no_remunerativo_extra),
            ),
        )
        conn.commit()
        row = conn.execute(
            """
            SELECT id FROM novedades_buzon
            WHERE cliente_id = ? AND empleado_id = ? AND periodo = ?
            """,
            (cliente_id, empleado_id, periodo),
        ).fetchone()
        return int(row["id"])


def listar_novedades_periodo(cliente_id: int, periodo: str) -> list[dict]:
    with obtener_conexion() as conn:
        rows = conn.execute(
            """
            SELECT n.*, e.nombre AS empleado_nombre, e.cuil
            FROM novedades_buzon n
            JOIN empleados_sueldos e ON e.id = n.empleado_id
            WHERE n.cliente_id = ? AND n.periodo = ?
            ORDER BY e.nombre COLLATE NOCASE
            """,
            (cliente_id, periodo),
        ).fetchall()
    return [dict(r) for r in rows]


def obtener_novedad(cliente_id: int, empleado_id: int, periodo: str) -> dict | None:
    with obtener_conexion() as conn:
        row = conn.execute(
            """
            SELECT * FROM novedades_buzon
            WHERE cliente_id = ? AND empleado_id = ? AND periodo = ?
            """,
            (cliente_id, empleado_id, periodo),
        ).fetchone()
    return dict(row) if row else None


def upsert_liquidacion_resultado(
    cliente_id: int,
    empleado_id: int,
    periodo: str,
    total_remunerativo: float,
    total_no_remunerativo: float,
    total_descuentos: float,
    neto_a_percibir: float,
    conceptos: list,
) -> int:
    payload = json.dumps(conceptos, ensure_ascii=False)
    with obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO liquidaciones_resultado (
                cliente_id, empleado_id, periodo,
                total_remunerativo, total_no_remunerativo,
                total_descuentos, neto_a_percibir, detalle_conceptos_json,
                liquidado_en
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(cliente_id, empleado_id, periodo) DO UPDATE SET
                total_remunerativo = excluded.total_remunerativo,
                total_no_remunerativo = excluded.total_no_remunerativo,
                total_descuentos = excluded.total_descuentos,
                neto_a_percibir = excluded.neto_a_percibir,
                detalle_conceptos_json = excluded.detalle_conceptos_json,
                liquidado_en = CURRENT_TIMESTAMP
            """,
            (
                cliente_id,
                empleado_id,
                periodo,
                float(total_remunerativo),
                float(total_no_remunerativo),
                float(total_descuentos),
                float(neto_a_percibir),
                payload,
            ),
        )
        conn.execute(
            """
            UPDATE novedades_buzon SET estado = 'Liquidada'
            WHERE cliente_id = ? AND empleado_id = ? AND periodo = ?
            """,
            (cliente_id, empleado_id, periodo),
        )
        conn.commit()
        row = conn.execute(
            """
            SELECT id FROM liquidaciones_resultado
            WHERE cliente_id = ? AND empleado_id = ? AND periodo = ?
            """,
            (cliente_id, empleado_id, periodo),
        ).fetchone()
        return int(row["id"])


def listar_liquidaciones_periodo(cliente_id: int, periodo: str) -> list[dict]:
    with obtener_conexion() as conn:
        rows = conn.execute(
            """
            SELECT l.*, e.nombre AS empleado_nombre, e.cuil, e.categoria
            FROM liquidaciones_resultado l
            JOIN empleados_sueldos e ON e.id = l.empleado_id
            WHERE l.cliente_id = ? AND l.periodo = ?
            ORDER BY e.nombre COLLATE NOCASE
            """,
            (cliente_id, periodo),
        ).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        try:
            d["conceptos"] = json.loads(d.get("detalle_conceptos_json") or "[]")
        except json.JSONDecodeError:
            d["conceptos"] = []
        out.append(d)
    return out


def obtener_liquidacion(liquidacion_id: int) -> dict | None:
    with obtener_conexion() as conn:
        row = conn.execute(
            """
            SELECT l.*, e.nombre AS empleado_nombre, e.cuil, e.categoria,
                   c.nombre AS empresa_nombre, c.cuit, c.cct_asignado
            FROM liquidaciones_resultado l
            JOIN empleados_sueldos e ON e.id = l.empleado_id
            JOIN clientes c ON c.id = l.cliente_id
            WHERE l.id = ?
            """,
            (liquidacion_id,),
        ).fetchone()
    if not row:
        return None
    d = dict(row)
    try:
        d["conceptos"] = json.loads(d.get("detalle_conceptos_json") or "[]")
    except json.JSONDecodeError:
        d["conceptos"] = []
    return d


def resumen_sueldos_empresas(periodo: str) -> list[dict]:
    """Panel estudio: empresas + CCT + estado de novedades del período."""
    with obtener_conexion() as conn:
        clientes = conn.execute(
            """
            SELECT id, nombre, cuit,
                   NULLIF(TRIM(COALESCE(cct_asignado, '')), '') AS cct_asignado
            FROM clientes
            ORDER BY nombre COLLATE NOCASE
            """
        ).fetchall()
        out = []
        for c in clientes:
            cid = c["id"]
            n_emp = conn.execute(
                "SELECT COUNT(*) AS n FROM empleados_sueldos WHERE cliente_id = ? AND activo = 1",
                (cid,),
            ).fetchone()["n"]
            n_nov = conn.execute(
                "SELECT COUNT(*) AS n FROM novedades_buzon WHERE cliente_id = ? AND periodo = ?",
                (cid, periodo),
            ).fetchone()["n"]
            n_liq = conn.execute(
                "SELECT COUNT(*) AS n FROM liquidaciones_resultado WHERE cliente_id = ? AND periodo = ?",
                (cid, periodo),
            ).fetchone()["n"]
            if n_liq > 0:
                estado = "Liquidado"
            elif n_nov > 0:
                estado = "Novedades Recibidas"
            else:
                estado = "Pendiente"
            cct = c["cct_asignado"]
            conv = None
            if cct:
                conv = conn.execute(
                    "SELECT nombre FROM convenios_colectivos WHERE codigo = ?",
                    (cct,),
                ).fetchone()
            out.append(
                {
                    "cliente_id": cid,
                    "nombre": c["nombre"],
                    "cuit": c["cuit"],
                    "cct_asignado": cct or "",
                    "convenio_nombre": (
                        conv["nombre"] if conv else ("Sin CCT asignado")
                    ),
                    "empleados": n_emp,
                    "novedades": n_nov,
                    "liquidaciones": n_liq,
                    "estado": estado,
                    "periodo": periodo,
                }
            )
    return out


# ---------------------------------------------------------------------------
# Motor de conciliacion bancaria
# ---------------------------------------------------------------------------

def _inicializar_tablas_conciliacion(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS clasificacion_reglas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patron TEXT NOT NULL,
            categoria TEXT NOT NULL,
            tipo TEXT NOT NULL,
            orden INTEGER NOT NULL DEFAULT 0,
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sociedad_bancos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            banco TEXT NOT NULL,
            orden INTEGER NOT NULL DEFAULT 0,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(cliente_id, banco),
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS bank_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            banco TEXT NOT NULL DEFAULT '',
            periodo TEXT,
            fecha TEXT,
            descripcion TEXT NOT NULL DEFAULT '',
            credito TEXT NOT NULL DEFAULT '0.00',
            debito TEXT NOT NULL DEFAULT '0.00',
            saldo TEXT NOT NULL DEFAULT '0.00',
            categoria TEXT NOT NULL DEFAULT '',
            tipo TEXT NOT NULL DEFAULT '',
            estado TEXT NOT NULL DEFAULT 'PENDIENTE',
            match_detalle TEXT,
            match_ref_id TEXT,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
        """
    )
    for col, spec in (
        ("fuente", "TEXT NOT NULL DEFAULT ''"),
        ("score", "TEXT NOT NULL DEFAULT '0'"),
        ("confianza", "TEXT NOT NULL DEFAULT ''"),
    ):
        try:
            conn.execute(f"ALTER TABLE bank_transactions ADD COLUMN {col} {spec}")
        except sqlite3.OperationalError:
            pass
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS proveedores_pendientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            fecha TEXT,
            tipo_comp TEXT NOT NULL DEFAULT '',
            num_comp TEXT NOT NULL DEFAULT '',
            razon_social TEXT NOT NULL DEFAULT '',
            cuit TEXT NOT NULL DEFAULT '',
            importe TEXT NOT NULL DEFAULT '0.00',
            usado INTEGER NOT NULL DEFAULT 0,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
        """
    )
    try:
        conn.execute("ALTER TABLE proveedores_pendientes ADD COLUMN cuit TEXT NOT NULL DEFAULT ''")
    except sqlite3.OperationalError:
        pass
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS veps_afip (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            numero_vep TEXT NOT NULL DEFAULT '',
            fecha TEXT,
            importe TEXT NOT NULL DEFAULT '0.00',
            impuesto TEXT NOT NULL DEFAULT '',
            periodo_fiscal TEXT NOT NULL DEFAULT '',
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS conciliacion_auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER,
            movimiento_id INTEGER,
            usuario TEXT NOT NULL DEFAULT '',
            accion TEXT NOT NULL DEFAULT '',
            categoria_anterior TEXT,
            categoria_nueva TEXT,
            detalle TEXT,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


def sembrar_reglas_conciliacion_default(*, ya_hay: int | None = None) -> None:
    """Inserta el diccionario del prompt solo si la tabla esta vacia.

    Si ya hay filas (aunque sean reglas editadas a mano) no se tocan.
    """
    if ya_hay is not None and ya_hay > 0:
        return
    from motor_conciliacion import REGLAS_SEED

    with obtener_conexion() as conn:
        n = int(
            _valor_fila(
                conn.execute("SELECT COUNT(*) AS n FROM clasificacion_reglas").fetchone(),
                "n",
            )
            or 0
        )
        if n > 0:
            return
        if not REGLAS_SEED:
            return
        conn.executemany(
            """
            INSERT INTO clasificacion_reglas (patron, categoria, tipo, orden, activo)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                (patron, categoria, tipo, i, 1)
                for i, (patron, categoria, tipo) in enumerate(REGLAS_SEED)
            ],
        )
        conn.commit()


def listar_reglas_clasificacion(solo_activas: bool = False) -> list[dict]:
    q = "SELECT * FROM clasificacion_reglas"
    if solo_activas:
        q += " WHERE activo = 1"
    q += " ORDER BY orden ASC, id ASC"
    with obtener_conexion() as conn:
        return [dict(r) for r in conn.execute(q).fetchall()]


def agregar_regla_clasificacion(patron: str, categoria: str, tipo: str, orden: int | None = None) -> int:
    with obtener_conexion() as conn:
        if orden is None:
            row = conn.execute("SELECT COALESCE(MAX(orden), -1) + 1 AS o FROM clasificacion_reglas").fetchone()
            orden = int(row["o"])
        cur = conn.execute(
            """
            INSERT INTO clasificacion_reglas (patron, categoria, tipo, orden, activo)
            VALUES (?, ?, ?, ?, 1)
            """,
            (patron.strip(), categoria.strip(), tipo.strip(), int(orden)),
        )
        conn.commit()
        return int(cur.lastrowid)


def actualizar_regla_clasificacion(regla_id: int, **campos) -> None:
    allowed = {"patron", "categoria", "tipo", "orden", "activo"}
    parts = []
    vals = []
    for k, v in campos.items():
        if k in allowed:
            parts.append(f"{k} = ?")
            vals.append(v)
    if not parts:
        return
    vals.append(regla_id)
    with obtener_conexion() as conn:
        conn.execute(f"UPDATE clasificacion_reglas SET {', '.join(parts)} WHERE id = ?", vals)
        conn.commit()


def listar_bancos_sociedad(cliente_id: int) -> list[str]:
    """Bancos que efectivamente usa esta sociedad (para filtrar el selector de
    Conciliación y para ofrecerlos como cuenta destino en transferencias entre
    cuentas propias). Vacío si todavía no se configuró ninguno para este
    cliente — en ese caso el resto de la app usa la lista completa como
    respaldo."""
    with obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT banco FROM sociedad_bancos WHERE cliente_id = ? ORDER BY orden ASC, id ASC",
            (int(cliente_id),),
        ).fetchall()
    return [str(f["banco"]) for f in filas]


def guardar_bancos_sociedad(cliente_id: int, bancos: list[str]) -> None:
    """Reemplaza la lista completa de bancos configurados para esta sociedad."""
    with obtener_conexion() as conn:
        conn.execute("DELETE FROM sociedad_bancos WHERE cliente_id = ?", (int(cliente_id),))
        vistos: set[str] = set()
        orden = 0
        for banco in bancos:
            nombre = str(banco or "").strip()
            if not nombre or nombre in vistos:
                continue
            vistos.add(nombre)
            conn.execute(
                "INSERT OR IGNORE INTO sociedad_bancos (cliente_id, banco, orden) VALUES (?, ?, ?)",
                (int(cliente_id), nombre, orden),
            )
            orden += 1
        conn.commit()


def borrar_movimientos_periodo(cliente_id: int, periodo: str | None = None, banco: str | None = None) -> None:
    with obtener_conexion() as conn:
        q = "DELETE FROM bank_transactions WHERE cliente_id = ?"
        args: list = [cliente_id]
        if periodo:
            q += " AND periodo = ?"
            args.append(periodo)
        if banco:
            q += " AND banco = ?"
            args.append(banco)
        conn.execute(q, args)
        conn.commit()


def insertar_movimientos_banco(filas: list[dict]) -> int:
    if not filas:
        return 0
    with obtener_conexion() as conn:
        for f in filas:
            periodo = f.get("periodo")
            if hasattr(periodo, "isoformat"):
                periodo = periodo.isoformat()
            fecha = f.get("fecha")
            if hasattr(fecha, "isoformat"):
                fecha = fecha.isoformat()
            conn.execute(
                """
                INSERT INTO bank_transactions (
                    cliente_id, banco, periodo, fecha, descripcion,
                    credito, debito, saldo, categoria, tipo, estado,
                    match_detalle, match_ref_id, fuente, score, confianza
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    int(f["cliente_id"]),
                    str(f.get("banco") or ""),
                    periodo,
                    fecha,
                    str(f.get("descripcion") or ""),
                    str(f.get("credito") or "0.00"),
                    str(f.get("debito") or "0.00"),
                    str(f.get("saldo") or "0.00"),
                    str(f.get("categoria") or ""),
                    str(f.get("tipo") or ""),
                    str(f.get("estado") or "PENDIENTE"),
                    f.get("match_detalle"),
                    f.get("match_ref_id"),
                    str(f.get("fuente") or ""),
                    str(f.get("score") or "0"),
                    str(f.get("confianza") or ""),
                ),
            )
        conn.commit()
        return len(filas)


def listar_movimientos_banco(
    cliente_id: int,
    periodo: str | None = None,
    banco: str | None = None,
    estado: str | None = None,
) -> list[dict]:
    q = "SELECT * FROM bank_transactions WHERE cliente_id = ?"
    args: list = [cliente_id]
    if periodo:
        q += " AND periodo = ?"
        args.append(periodo)
    if banco:
        q += " AND banco = ?"
        args.append(banco)
    if estado:
        q += " AND estado = ?"
        args.append(estado)
    q += " ORDER BY fecha ASC, id ASC"
    with obtener_conexion() as conn:
        return [dict(r) for r in conn.execute(q, args).fetchall()]


def actualizar_movimiento_banco(mov_id: int, **campos) -> None:
    allowed = {"categoria", "tipo", "estado", "match_detalle", "match_ref_id", "fuente", "score", "confianza"}
    parts, vals = [], []
    for k, v in campos.items():
        if k in allowed:
            parts.append(f"{k} = ?")
            vals.append(v)
    if not parts:
        return
    vals.append(mov_id)
    with obtener_conexion() as conn:
        conn.execute(f"UPDATE bank_transactions SET {', '.join(parts)} WHERE id = ?", vals)
        conn.commit()


def registrar_auditoria_conciliacion(
    *,
    cliente_id: int | None,
    movimiento_id: int | None,
    usuario: str,
    accion: str,
    categoria_anterior: str | None = None,
    categoria_nueva: str | None = None,
    detalle: str | None = None,
) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO conciliacion_auditoria (
                cliente_id, movimiento_id, usuario, accion,
                categoria_anterior, categoria_nueva, detalle
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                cliente_id,
                movimiento_id,
                usuario or "",
                accion,
                categoria_anterior,
                categoria_nueva,
                detalle,
            ),
        )
        conn.commit()


def reemplazar_proveedores_pendientes(cliente_id: int, filas: list[dict]) -> int:
    with obtener_conexion() as conn:
        conn.execute("DELETE FROM proveedores_pendientes WHERE cliente_id = ?", (cliente_id,))
        for f in filas:
            fecha = f.get("fecha")
            if hasattr(fecha, "isoformat"):
                fecha = fecha.isoformat()
            conn.execute(
                """
                INSERT INTO proveedores_pendientes (
                    cliente_id, fecha, tipo_comp, num_comp, razon_social, cuit, importe, usado
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    cliente_id,
                    fecha,
                    str(f.get("tipo_comp") or f.get("tipo") or ""),
                    str(f.get("num_comp") or f.get("comprobante") or ""),
                    str(f.get("razon_social") or f.get("proveedor") or ""),
                    str(f.get("cuit") or ""),
                    str(f.get("importe") or "0.00"),
                    1 if f.get("usado") else 0,
                ),
            )
        conn.commit()
        return len(filas)


def resumen_proveedores_pendientes(cliente_id: int) -> list[dict]:
    """Total pendiente agrupado por CUIT/razón social (para mostrar en la UI)."""
    with obtener_conexion() as conn:
        rows = conn.execute(
            """
            SELECT cuit, razon_social,
                   COUNT(*) AS facturas,
                   SUM(CAST(importe AS REAL)) AS importe_total
            FROM proveedores_pendientes
            WHERE cliente_id = ? AND usado = 0
            GROUP BY cuit, razon_social
            ORDER BY importe_total DESC
            """,
            (cliente_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def listar_proveedores_pendientes(cliente_id: int, solo_libres: bool = True) -> list[dict]:
    q = "SELECT * FROM proveedores_pendientes WHERE cliente_id = ?"
    args: list = [cliente_id]
    if solo_libres:
        q += " AND usado = 0"
    q += " ORDER BY fecha ASC, id ASC"
    with obtener_conexion() as conn:
        rows = [dict(r) for r in conn.execute(q, args).fetchall()]
    for r in rows:
        r["usado"] = bool(r.get("usado"))
    return rows


def marcar_proveedor_usado(factura_id: int, usado: bool = True) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE proveedores_pendientes SET usado = ? WHERE id = ?",
            (1 if usado else 0, factura_id),
        )
        conn.commit()


def reemplazar_veps_afip(cliente_id: int, filas: list[dict]) -> int:
    with obtener_conexion() as conn:
        conn.execute("DELETE FROM veps_afip WHERE cliente_id = ?", (cliente_id,))
        for f in filas:
            fecha = f.get("fecha")
            if hasattr(fecha, "isoformat"):
                fecha = fecha.isoformat()
            conn.execute(
                """
                INSERT INTO veps_afip (
                    cliente_id, numero_vep, fecha, importe, impuesto, periodo_fiscal
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    cliente_id,
                    str(f.get("numero_vep") or ""),
                    fecha,
                    str(f.get("importe") or "0.00"),
                    str(f.get("impuesto") or ""),
                    str(f.get("periodo_fiscal") or ""),
                ),
            )
        conn.commit()
        return len(filas)


def listar_veps_afip(cliente_id: int) -> list[dict]:
    with obtener_conexion() as conn:
        return [
            dict(r)
            for r in conn.execute(
                "SELECT * FROM veps_afip WHERE cliente_id = ? ORDER BY fecha ASC, id ASC",
                (cliente_id,),
            ).fetchall()
        ]
