"""Errores de libsql traducidos a sqlite3, sin red ni Turso real."""
from __future__ import annotations

import sqlite3

import pytest

import database

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
    with pytest.raises(sqlite3.IntegrityError) as exc:
        _envolver(causa, "executemany").executemany("INSERT INTO clientes (cuit) VALUES (?)", [("1",)])
    assert "UNIQUE constraint failed" in str(exc.value)
    assert exc.value.__cause__ is causa
    assert database._turso_conn_obj is sentinel
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
            self.many = None

        def execute(self, sql, parametros=()):
            self.params = parametros
            return self

        def executemany(self, sql, secuencia):
            self.many = list(secuencia)
            return self

        def fetchall(self):
            return []

    interno = _Graba()
    conn = database._ConexionCompatTurso(interno)
    conn.execute("UPDATE t SET a = ? WHERE id = ?", ["nombre", 1])
    conn.executemany("INSERT INTO t (a) VALUES (?)", [["uno"], ["dos"]])
    assert interno.params == ("nombre", 1)
    assert interno.many == [("uno",), ("dos",)]


class _ContadorSentencias:
    """Cuenta llamadas a execute/executemany, no los commits."""

    def __init__(self, conn):
        self._conn = conn
        self.execute_n = 0
        self.executemany_n = 0
        self.insert_execute = 0

    def execute(self, sql, parametros=()):
        self.execute_n += 1
        if str(sql).lstrip().upper().startswith("INSERT"):
            self.insert_execute += 1
        return self._conn.execute(sql, parametros)

    def executemany(self, sql, secuencia):
        self.executemany_n += 1
        return self._conn.executemany(sql, secuencia)

    def __getattr__(self, nombre):
        return getattr(self._conn, nombre)

    @property
    def sentencias(self) -> int:
        return self.execute_n + self.executemany_n


def test_inicializar_bd_dos_veces_en_libsql_local(tmp_path, monkeypatch, capsys):
    """La misma base libsql local arranca dos veces, y la segunda no resiembra."""
    libsql = pytest.importorskip("libsql_experimental")
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "estudio.db")
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    contador = _ContadorSentencias(libsql.connect(str(tmp_path / "x.db")))
    monkeypatch.setattr(database, "_turso_conn_obj", contador)
    database.inicializar_bd()
    primera = contador.sentencias
    many_1 = contador.executemany_n
    insert_1 = contador.insert_execute
    contador.execute_n = 0
    contador.executemany_n = 0
    contador.insert_execute = 0
    database.inicializar_bd()
    segunda = contador.sentencias
    many_2 = contador.executemany_n
    with capsys.disabled():
        print(f"SENTENCIAS inicializar_bd 1={primera} 2={segunda} executemany_1={many_1} executemany_2={many_2}")
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
    assert many_1 >= 3
    assert many_2 == 0
    assert insert_1 < 30
    assert segunda < primera
