"""Errores de libsql traducidos a sqlite3, sin red ni Turso real."""
from __future__ import annotations

import socket
import sqlite3
import subprocess
import time
from pathlib import Path

import pytest

import auth_oficina
import database
import ui_version_web as uv

_DUPLICADO = (
    'Hrana: stream error: Error { message: "SQLite error: duplicate column name: '
    'mes_cierre_balance", code: "SQLITE_UNKNOWN" }'
)
_UNICO = "SQLite error: UNIQUE constraint failed: clientes.cuit (SQLITE_CONSTRAINT)"
_RED = "Hrana: stream error: connection reset"


class _ConnFalso:
    def __init__(self, exc: BaseException, *, donde: str = "execute"):
        self.exc = exc
        self.donde = donde
        self.sync_llamado = False

    def execute(self, sql, parametros=()):
        if self.donde == "execute":
            raise self.exc
        return self

    def executemany(self, sql, secuencia):
        if self.donde == "executemany":
            raise self.exc
        return self

    def commit(self):
        if self.donde == "commit":
            raise self.exc

    def sync(self):
        self.sync_llamado = True

    def fetchall(self):
        return []

    def __iter__(self):
        return iter(())


def _envolver(exc: BaseException, donde: str = "execute") -> database._ConexionCompatTurso:
    return database._ConexionCompatTurso(_ConnFalso(exc, donde=donde))


def test_columna_duplicada_es_operational_y_no_tira_la_conexion(monkeypatch):
    sentinel = object()
    monkeypatch.setattr(database, "_turso_conn_obj", sentinel)
    causa = ValueError(_DUPLICADO)
    with pytest.raises(sqlite3.OperationalError) as exc:
        _envolver(causa).execute("ALTER TABLE clientes ADD COLUMN mes_cierre_balance INTEGER")
    assert "duplicate column name: mes_cierre_balance" in str(exc.value)
    assert exc.value.__cause__ is causa
    assert database._turso_conn_obj is sentinel


def test_columna_duplicada_de_libsql_local_tambien(monkeypatch):
    """El cliente embebido no agrega el prefijo 'SQLite error'."""
    sentinel = object()
    monkeypatch.setattr(database, "_turso_conn_obj", sentinel)
    causa = ValueError("duplicate column name: mes_cierre_balance")
    with pytest.raises(sqlite3.OperationalError) as exc:
        _envolver(causa).execute("ALTER TABLE clientes ADD COLUMN mes_cierre_balance INTEGER")
    assert exc.value.__cause__ is causa
    assert database._turso_conn_obj is sentinel


def test_unique_es_integrity_error(monkeypatch):
    sentinel = object()
    monkeypatch.setattr(database, "_turso_conn_obj", sentinel)
    causa = ValueError(_UNICO)
    # El INSERT masivo sale por execute; un UPDATE sigue por executemany.
    with pytest.raises(sqlite3.IntegrityError) as exc:
        _envolver(causa, "execute").executemany(
            "INSERT INTO clientes (cuit) VALUES (?)", [("1",), ("2",)]
        )
    assert "UNIQUE constraint failed" in str(exc.value)
    assert exc.value.__cause__ is causa
    assert database._turso_conn_obj is sentinel
    with pytest.raises(sqlite3.IntegrityError):
        _envolver(causa, "executemany").executemany(
            "UPDATE clientes SET cuit = ? WHERE id = ?", [("1", 1)]
        )
    with pytest.raises(sqlite3.IntegrityError):
        _envolver(causa, "commit").commit()
    assert database._turso_conn_obj is sentinel


def test_error_de_red_tira_la_conexion(monkeypatch):
    sentinel = object()
    monkeypatch.setattr(database, "_turso_conn_obj", sentinel)
    causa = ValueError(_RED)
    with pytest.raises(ValueError) as exc:
        _envolver(causa).execute("SELECT 1")
    assert exc.value is causa
    assert "SQLite error" not in str(causa)
    assert database._turso_conn_obj is None


def test_parametros_en_lista_pasan_como_tupla():
    class _Graba:
        def __init__(self):
            self.params = None
            self.sql = ""
            self.many = None
            self.ejecutados: list = []

        def execute(self, sql, parametros=()):
            self.params = parametros
            self.sql = sql
            self.ejecutados.append(parametros)
            return self

        def executemany(self, sql, secuencia):
            self.many = list(secuencia)
            return self

        def fetchall(self):
            return []

    interno = _Graba()
    conn = database._ConexionCompatTurso(interno)
    conn.execute("UPDATE t SET a = ? WHERE id = ?", ["nombre", 1])
    conn.executemany("UPDATE t SET a = ? WHERE id = ?", [["uno", 1], ["dos", 2]])
    conn.executemany("INSERT INTO t (a) VALUES (?)", [["uno"], ["dos"], ["tres"]])
    assert interno.ejecutados[0] == ("nombre", 1)
    assert interno.params == ("uno", "dos", "tres")
    assert "(?),(?),(?)".replace(" ", "") in interno.sql.replace(" ", "")
    assert interno.many == [("uno", 1), ("dos", 2)]


class _ContadorSentencias:
    """Cuenta viajes: execute, executemany y executescript (el batch cuenta 1)."""

    def __init__(self, conn):
        self._conn = conn
        self.execute_n = 0
        self.executemany_n = 0
        self.executescript_n = 0
        self.insert_execute = 0

    def execute(self, sql, parametros=()):
        self.execute_n += 1
        texto = str(sql).lstrip().upper()
        if texto.startswith("INSERT"):
            self.insert_execute += 1
        return self._conn.execute(sql, parametros)

    def executemany(self, sql, secuencia):
        self.executemany_n += 1
        return self._conn.executemany(sql, secuencia)

    def cursor(self):
        contador = self
        real = self._conn.cursor()

        class _Cursor:
            def executescript(self, sql):
                contador.executescript_n += 1
                return real.executescript(sql)

            def __getattr__(self, nombre):
                return getattr(real, nombre)

        return _Cursor()

    def __getattr__(self, nombre):
        return getattr(self._conn, nombre)

    @property
    def sentencias(self) -> int:
        return self.execute_n + self.executemany_n + self.executescript_n


def test_inicializar_bd_dos_veces_en_libsql_local(tmp_path, monkeypatch, capsys):
    """La misma base libsql local arranca dos veces, y la segunda no resiembra."""
    libsql = pytest.importorskip("libsql_experimental")
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "estudio.db")
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    monkeypatch.setattr(database, "_turso_es_replica", False)
    monkeypatch.setattr(database, "_sync_hecho_en_rerun", False)
    monkeypatch.setattr(database, "_ddl_bloqueado", False)
    contador = _ContadorSentencias(libsql.connect(str(tmp_path / "x.db")))
    monkeypatch.setattr(database, "_turso_conn_obj", contador)
    t0 = time.perf_counter()
    database.inicializar_bd()
    local_1 = time.perf_counter() - t0
    primera = contador.sentencias
    many_1 = contador.executemany_n
    batch_1 = contador.executescript_n
    contador.execute_n = 0
    contador.executemany_n = 0
    contador.executescript_n = 0
    contador.insert_execute = 0
    t1 = time.perf_counter()
    database.inicializar_bd()
    local_2 = time.perf_counter() - t1
    segunda = contador.sentencias
    with capsys.disabled():
        print(
            f"SENTENCIAS inicializar_bd 1={primera} 2={segunda} "
            f"executemany_1={many_1} batch_1={batch_1} "
            f"local_s 1={local_1:.3f} 2={local_2:.3f} "
            f"estimado_turso_s 1={primera * 0.35:.2f} 2={segunda * 0.35:.2f}"
        )
    with database.obtener_conexion() as conn:
        columnas = {fila[1] for fila in conn.execute("PRAGMA table_info(clientes)").fetchall()}
        tablas = {
            fila[0]
            for fila in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        n_tc = conn.execute("SELECT COUNT(*) AS n FROM inversiones_tc_bna").fetchone()["n"]
        n_reglas = conn.execute("SELECT COUNT(*) AS n FROM clasificacion_reglas").fetchone()["n"]
    assert "mes_cierre_balance" in columnas
    for nombre in (
        "usuarios_oficina",
        "rutina_pedidos",
        "rutina_ficha_proyecciones",
        "arca_emisores",
        "inversiones_operaciones",
        "app_meta",
    ):
        assert nombre in tablas
    assert int(n_tc) > 100
    assert int(n_reglas) > 0
    assert many_1 == 0
    assert batch_1 == 0
    assert segunda == 1
    assert primera < 80
    contador.execute_n = 0
    contador.executemany_n = 0
    contador.executescript_n = 0
    catalogo = [
        {"nombre": f"MONO {i}", "cuit": f"20{i:09d}", "tipo": "Monotributista"}
        for i in range(139)
    ]
    t2 = time.perf_counter()
    stats = database.sincronizar_clientes_catalogo(catalogo, semilla="monotributistas")
    local_cat = time.perf_counter() - t2
    with capsys.disabled():
        print(
            f"CATALOGO 139 sentencias={contador.sentencias} executemany={contador.executemany_n} "
            f"local_s={local_cat:.3f} estimado_turso_s={contador.sentencias * 0.35:.2f} stats={stats}"
        )
    assert stats["insertados"] == 139
    assert contador.executemany_n == 0
    assert contador.sentencias < 15


def test_definir_esquema_no_reconstruye_clientes():
    """Aunque el CREATE de clientes venga vacío, no se arma clientes_new."""
    columnas = {tabla: set() for tabla in database._TABLAS_ESQUEMA}
    grabador = database._GrabadorDDL(columnas, "")
    database._definir_esquema(grabador)
    texto = "\n".join(grabador.sentencias)
    assert "clientes_new" not in texto.lower()
    assert "drop table clientes" not in texto.lower()
    assert database._duda_esquema({"clientes": set()}, "") is False
    assert database._duda_esquema({"clientes": {"id"}}, "") is True
    assert database._duda_esquema({"clientes": set()}, "CREATE TABLE clientes (id INTEGER)") is True


def test_ddl_que_falla_no_deja_cambios(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "local.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    monkeypatch.setattr(database, "_turso_es_replica", False)
    with database.obtener_conexion() as conn:
        conn.execute(
            "CREATE TABLE clientes (id INTEGER PRIMARY KEY, nombre TEXT, cct_asignado TEXT, plan_cuentas_csv TEXT)"
        )
        conn.execute(
            "INSERT INTO clientes (nombre, cct_asignado, plan_cuentas_csv) VALUES ('ACME', 'UOCRA_76', 'PLAN-OK')"
        )
        conn.commit()
        with pytest.raises(sqlite3.OperationalError):
            database._aplicar_ddl(
                conn,
                [
                    "CREATE TABLE marca_ddl (id INTEGER)",
                    "ALTER TABLE no_existe ADD COLUMN x INTEGER",
                ],
            )
        marcas = conn.execute(
            "SELECT name FROM sqlite_master WHERE name = 'marca_ddl'"
        ).fetchall()
        fila = conn.execute("SELECT nombre, cct_asignado, plan_cuentas_csv FROM clientes").fetchone()
    assert marcas == []
    assert fila["nombre"] == "ACME"
    assert fila["cct_asignado"] == "UOCRA_76"
    assert fila["plan_cuentas_csv"] == "PLAN-OK"


def test_sync_fallido_no_toca_la_base(monkeypatch):
    class _Replica:
        def __init__(self):
            self.sql: list[str] = []

        def sync(self):
            raise RuntimeError("réplica vacía")

        def execute(self, sql, parametros=()):
            self.sql.append(str(sql))
            raise AssertionError(sql)

        def commit(self):
            raise AssertionError("commit")

        def rollback(self):
            return None

    replica = _Replica()
    monkeypatch.setattr(database, "_turso_conn_obj", replica)
    monkeypatch.setattr(database, "_turso_es_replica", True)
    monkeypatch.setattr(database, "_sync_hecho_en_rerun", False)
    monkeypatch.setattr(database, "_ddl_bloqueado", False)
    database.inicializar_bd()
    assert replica.sql == []
    assert database.ddl_permitido() is False


def _sqld() -> Path | None:
    candidatos = [
        Path("/tmp/sqld/libsql-server-x86_64-unknown-linux-gnu/sqld"),
        Path("/usr/local/bin/sqld"),
    ]
    for ruta in candidatos:
        if ruta.is_file():
            return ruta
    return None


def _puerto_libre() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def test_replica_nueva_sin_sync_no_reconstruye_clientes(tmp_path, monkeypatch):
    """Réplica recién creada, sin sync previo, contra un primario que ya tiene datos."""
    binario = _sqld()
    if binario is None:
        pytest.skip("no está el binario de sqld")
    libsql = pytest.importorskip("libsql_experimental")
    puerto = _puerto_libre()
    primario = tmp_path / "primario"
    proceso = subprocess.Popen(
        [
            str(binario),
            "--db-path",
            str(primario),
            "--http-listen-addr",
            f"127.0.0.1:{puerto}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    url = f"http://127.0.0.1:{puerto}"
    try:
        replica_setup = tmp_path / "setup.db"
        deadline = time.time() + 8
        conn = None
        while time.time() < deadline:
            try:
                conn = libsql.connect(str(replica_setup), sync_url=url)
                conn.sync()
                break
            except Exception:
                time.sleep(0.05)
        assert conn is not None
        conn.execute(
            """
            CREATE TABLE clientes (
                id INTEGER PRIMARY KEY,
                nombre TEXT NOT NULL,
                cuit TEXT NOT NULL UNIQUE,
                tipo_persona TEXT NOT NULL,
                plan_cuentas_csv TEXT,
                cct_asignado TEXT
            )
            """
        )
        conn.execute(
            """
            INSERT INTO clientes (nombre, cuit, tipo_persona, plan_cuentas_csv, cct_asignado)
            VALUES ('ACME SA', '30123456789', 'Persona Jurídica', 'PLAN-OK', 'UOCRA_76')
            """
        )
        conn.execute(
            """
            CREATE TABLE usuarios_oficina (
                id INTEGER PRIMARY KEY,
                usuario TEXT NOT NULL UNIQUE,
                nombre TEXT,
                pin_hash TEXT,
                es_admin INTEGER,
                activo INTEGER,
                intentos_fallidos INTEGER DEFAULT 0,
                bloqueado_hasta TEXT
            )
            """
        )
        conn.execute(
            """
            INSERT INTO usuarios_oficina (usuario, nombre, pin_hash, es_admin, activo, intentos_fallidos)
            VALUES ('existente', 'Existente', 'x', 0, 1, 4)
            """
        )
        conn.commit()
        conn.sync()
        conn.close()

        replica_nueva = tmp_path / "replica-nueva.db"
        monkeypatch.setattr(database, "DB_PATH", replica_nueva)
        monkeypatch.setattr(database, "_turso_conn_obj", None)
        monkeypatch.setattr(database, "_turso_es_replica", False)
        monkeypatch.setattr(database, "_sync_hecho_en_rerun", False)
        monkeypatch.setattr(database, "_ddl_bloqueado", False)
        monkeypatch.setattr(database, "_credenciales_turso", lambda: (url, ""))
        # No se llama sync() acá: lo tiene que hacer inicializar_bd antes de leer.
        database.inicializar_bd()

        control = libsql.connect(str(tmp_path / "control.db"), sync_url=url)
        control.sync()
        fila = control.execute(
            "SELECT nombre, cct_asignado, plan_cuentas_csv FROM clientes WHERE cuit = '30123456789'"
        ).fetchone()
        columnas = {item[1] for item in control.execute("PRAGMA table_info(clientes)").fetchall()}
        tablas = {item[0] for item in control.execute("SELECT name FROM sqlite_master").fetchall()}
        intentos = {
            item[1]
            for item in control.execute("PRAGMA table_info(usuarios_oficina)").fetchall()
        }
        control.close()
    finally:
        proceso.terminate()
        try:
            proceso.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proceso.kill()

    assert "clientes_new" not in tablas
    assert fila is not None
    assert fila[0] == "ACME SA"
    assert fila[1] == "UOCRA_76"
    assert fila[2] == "PLAN-OK"
    assert "cct_asignado" in columnas
    assert "plan_cuentas_csv" in columnas
    assert "intentos_fallidos" in intentos


def test_cache_de_clientes_no_repite_y_se_invalida(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "cache.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    monkeypatch.setattr(database, "_turso_es_replica", False)
    monkeypatch.setattr(database, "_ddl_bloqueado", False)
    database.inicializar_bd()
    llamadas = {"n": 0}
    real = database._listar_clientes_directo

    def cuenta():
        llamadas["n"] += 1
        return real()

    monkeypatch.setattr(database, "_listar_clientes_directo", cuenta)
    database.listar_clientes()
    database.listar_clientes()
    assert llamadas["n"] == 1
    database.crear_cliente("Cache SA", "30777777771", "Persona Jurídica", None, 12)
    database.listar_clientes()
    assert llamadas["n"] == 2
    assert any(fila["nombre"] == "Cache SA" for fila in database.listar_clientes())
    assert llamadas["n"] == 2


def test_commit_sin_escritura_no_sincroniza_y_con_escritura_una_vez():
    interno = _ConnFalso(RuntimeError("no se usa"), donde="ninguno")
    conn = database._ConexionCompatTurso(interno)
    conn.execute("SELECT 1")
    conn.commit()
    assert interno.sync_llamado is False
    generacion = database.generacion_datos()
    conn.execute("UPDATE clientes SET razon_social = ? WHERE id = ?", ("x", 1))
    conn.commit()
    assert interno.sync_llamado is True
    assert database.generacion_datos() == generacion + 1
    interno.sync_llamado = False
    conn.commit()
    assert interno.sync_llamado is False


class _SyncCuenta:
    def __init__(self):
        self.n = 0
        self.dentro_del_lock = False

    def sync(self):
        self.n += 1
        self.dentro_del_lock = database._turso_conn_lock.locked()

    def execute(self, sql, parametros=()):
        return self

    def commit(self):
        return None

    def fetchall(self):
        return []


def test_comenzar_rerun_no_sincroniza(monkeypatch):
    """Con sync_interval la pantalla no sale a la red."""
    replica = _SyncCuenta()
    monkeypatch.setattr(database, "_turso_conn_obj", replica)
    monkeypatch.setattr(database, "_turso_es_replica", True)
    monkeypatch.setattr(database, "_sync_fondo", True)
    monkeypatch.setattr(database, "_bucket_sesion", lambda: None)
    database.comenzar_rerun()
    database.comenzar_rerun()
    assert replica.n == 0


def test_fallback_sincroniza_cada_15s_fuera_del_lock(monkeypatch):
    replica = _SyncCuenta()
    monkeypatch.setattr(database, "_turso_conn_obj", replica)
    monkeypatch.setattr(database, "_turso_es_replica", True)
    monkeypatch.setattr(database, "_sync_fondo", False)
    monkeypatch.setattr(database, "_ultimo_sync_periodico", 0.0)
    monkeypatch.setattr(database, "_bucket_sesion", lambda: None)
    database.comenzar_rerun()
    assert replica.n == 1
    assert replica.dentro_del_lock is False
    database.comenzar_rerun()
    assert replica.n == 1
    monkeypatch.setattr(database, "_ultimo_sync_periodico", time.monotonic() - 16)
    database.comenzar_rerun()
    assert replica.n == 2
    assert replica.dentro_del_lock is False


def test_connect_pide_sync_interval(monkeypatch):
    monkeypatch.setattr(database, "_sync_fondo", False)
    vistos: dict = {}

    class _Lib:
        @staticmethod
        def connect(path, **kwargs):
            vistos["kwargs"] = kwargs
            return object()

    monkeypatch.setattr(database, "_libsql", _Lib)
    conn = database._conectar_libsql("libsql://local", "token")
    assert conn is not None
    assert vistos["kwargs"]["sync_interval"] == database._SYNC_INTERVALO_S
    assert vistos["kwargs"]["sync_url"] == "libsql://local"
    assert vistos["kwargs"]["auth_token"] == "token"
    assert database._sync_fondo is True


def test_sync_interval_ausente_no_lo_manda(monkeypatch):
    monkeypatch.setattr(database, "_sync_fondo", True)
    llamadas: list[dict] = []

    class _Lib:
        @staticmethod
        def connect(path, **kwargs):
            llamadas.append(dict(kwargs))
            if "sync_interval" in kwargs:
                raise TypeError("sync_interval")
            return object()

    monkeypatch.setattr(database, "_libsql", _Lib)
    assert database._conectar_libsql("libsql://local", "token") is not None
    assert database._sync_fondo is False
    assert "sync_interval" not in llamadas[-1]


def test_commit_con_escritura_sincroniza_una_vez(monkeypatch):
    monkeypatch.setattr(database, "_bucket_sesion", lambda: None)
    database.comenzar_rerun()
    interno = _SyncCuenta()
    conn = database._ConexionCompatTurso(interno)
    conn.execute("SELECT 1")
    conn.commit()
    assert interno.n == 0
    conn.execute("INSERT INTO t (a) VALUES (?)", ("x",))
    conn.commit()
    assert interno.n == 1
    assert database.contadores()["syncs"] == 1
    conn.commit()
    assert interno.n == 1


def test_contadores_son_por_sesion(monkeypatch):
    sesiones: dict[str, dict] = {}
    actual = {"id": "a"}

    def bucket():
        return sesiones.setdefault(
            actual["id"],
            {
                "sync_hecho": False,
                "consultas": 0,
                "syncs": 0,
                "sync_s": 0.0,
                "t0": 0.0,
                "pagina": "",
            },
        )

    replica = _SyncCuenta()
    monkeypatch.setattr(database, "_bucket_sesion", bucket)
    monkeypatch.setattr(database, "_turso_conn_obj", replica)
    monkeypatch.setattr(database, "_turso_es_replica", True)
    monkeypatch.setattr(database, "_sync_fondo", True)
    monkeypatch.setattr(database, "_sync_ok", True)

    database.comenzar_rerun()
    assert replica.n == 0
    conn = database._ConexionCompatTurso(replica)
    conn.execute("SELECT 1")
    conn.execute("SELECT 2")
    assert database.contadores() == {"consultas": 2, "syncs": 0}
    assert database.sincronizar_replica_una_vez() is True
    assert replica.n == 1
    assert database.sincronizar_replica_una_vez() is True
    assert replica.n == 1

    actual["id"] = "b"
    database.comenzar_rerun()
    assert database.contadores() == {"consultas": 0, "syncs": 0}
    assert database.sincronizar_replica_una_vez() is True
    assert replica.n == 2
    conn.execute("SELECT 3")
    assert database.contadores()["consultas"] == 1

    actual["id"] = "a"
    assert database.contadores() == {"consultas": 2, "syncs": 1}


def test_linea_de_tiempo_del_rerun(capsys, monkeypatch):
    monkeypatch.setattr(database, "_bucket_sesion", lambda: None)
    monkeypatch.setattr(database, "_turso_es_replica", False)
    database.comenzar_rerun()
    database.anotar_pagina("login")
    database.cerrar_rerun()
    linea = capsys.readouterr().out.strip().splitlines()[-1]
    assert linea.startswith("RERUN ")
    assert "pagina=login" in linea
    assert "consultas=0" in linea
    assert "syncs=0" in linea
    assert "sync=" in linea
    assert "total=" in linea


def test_equipo_con_hash_no_rehashea_y_el_login_es_de_un_usuario(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "oficina.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    monkeypatch.setattr(database, "_turso_es_replica", False)
    monkeypatch.setattr(auth_oficina, "_EQUIPO_PIN_LISTO", False)
    hashes = {"n": 0}

    def cuenta(pin, salt=None):
        hashes["n"] += 1
        return f"pbkdf2${salt or 's'}$ok"

    monkeypatch.setattr(auth_oficina, "_hash_pin", cuenta)
    with database.obtener_conexion() as conn:
        auth_oficina.inicializar_tabla_usuarios_oficina(conn)
        for login, nombre in auth_oficina._EQUIPO_OFICINA:
            conn.execute(
                "INSERT INTO usuarios_oficina (usuario, nombre, pin_hash, es_admin, activo) "
                "VALUES (?, ?, ?, 0, 1)",
                (login, nombre, "pbkdf2$ab$cd"),
            )
        conn.commit()
    marca = auth_oficina._marca_equipo(True)
    assert (
        auth_oficina.sembrar_equipo_oficina(
            forzar_pin=True, conocemos_marca=True, marca_leida=marca
        )
        == 0
    )
    assert hashes["n"] == 0
    assert auth_oficina.equipo_pin_listo() is True
    monkeypatch.setattr(auth_oficina, "_EQUIPO_PIN_LISTO", False)
    auth_oficina.sembrar_equipo_oficina(
        forzar_pin=True, conocemos_marca=True, marca_leida="marca-vieja"
    )
    assert hashes["n"] == 0

    monkeypatch.setattr(
        auth_oficina,
        "obtener_usuario_oficina",
        lambda usuario: {
            "id": 1,
            "usuario": usuario,
            "nombre": "Guadi",
            "pin_hash": "pbkdf2$ab$ok",
            "es_admin": 0,
            "activo": 1,
            "intentos_fallidos": 0,
            "bloqueado_hasta": None,
        },
    )
    monkeypatch.setattr(auth_oficina, "_limpiar_intentos_fallidos", lambda _usuario_id: None)
    assert auth_oficina.verificar_login_oficina("guada", "3278") is not None
    assert auth_oficina.verificar_login_oficina("guada", "3278") is not None
    # Dos intentos, dos hashes: no se cachea el acierto. Nunca los 11 del equipo.
    assert hashes["n"] == 2


def test_version_github_como_maximo_cada_10_minutos(monkeypatch):
    llamadas = {"n": 0}

    def fetch():
        llamadas["n"] += 1
        return "v9"

    monkeypatch.setattr(uv, "_fetch_version_remota", fetch)
    monkeypatch.setattr(uv, "_remoto_id", "")
    monkeypatch.setattr(uv, "_remoto_ts", 0.0)
    assert uv._version_github_cached() == "v9"
    assert uv._version_github_cached() == "v9"
    assert llamadas["n"] == 1
    monkeypatch.setattr(uv, "_remoto_ts", time.monotonic() - 601)
    assert uv._version_github_cached() == "v9"
    assert llamadas["n"] == 2
