"""Errores de libsql traducidos a sqlite3, sin red ni Turso real."""
from __future__ import annotations

import sqlite3
import time

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
    assert batch_1 == 1
    assert segunda == 1
    assert primera < 40
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
