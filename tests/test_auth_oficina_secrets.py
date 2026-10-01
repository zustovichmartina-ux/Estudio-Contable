"""Usuarios de Secrets: una vez por proceso, sin rehashear ni pisar el bloqueo.

Sin TURSO_*: SQLite temporal. El PBKDF2 real (600k) no corre en estas pruebas.
"""
from __future__ import annotations

import hashlib
import logging
import os

import pytest

import auth_oficina
import database


@pytest.fixture(autouse=True)
def _reset_flag():
    auth_oficina._SECRETS_USUARIOS_APLICADOS = False
    yield
    auth_oficina._SECRETS_USUARIOS_APLICADOS = False


@pytest.fixture
def pbkdf2(monkeypatch):
    llamadas = {"n": 0}

    def fake(hash_name, password, salt, iterations, dklen=None):
        llamadas["n"] += 1
        return hashlib.sha256(password + salt + str(iterations).encode()).digest()

    monkeypatch.setattr(hashlib, "pbkdf2_hmac", fake)
    return llamadas


@pytest.fixture
def db_local(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "estudio.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_turso_es_replica", False)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))
    monkeypatch.delenv("TURSO_DATABASE_URL", raising=False)
    monkeypatch.delenv("TURSO_AUTH_TOKEN", raising=False)
    with database.obtener_conexion() as conn:
        auth_oficina.inicializar_tabla_usuarios_oficina(conn)
        database._asegurar_app_meta(conn)
        conn.commit()
    return tmp_path


def _secrets(monkeypatch, usuarios: list[dict]) -> None:
    monkeypatch.setattr(auth_oficina, "_secrets_oficina_usuarios", lambda: list(usuarios))


def _admin(pin: str = "654321", nombre: str = "Administrador", es_admin: bool = True) -> dict:
    return {"usuario": "admin", "nombre": nombre, "pin": pin, "es_admin": es_admin}


def _fila(usuario: str) -> dict:
    fila = auth_oficina.obtener_usuario_oficina(usuario)
    assert fila is not None
    return fila


def _sqls(monkeypatch) -> list[str]:
    vistos: list[str] = []
    real = database.obtener_conexion

    def espia():
        conn = real()
        original = conn.execute

        def execute(sql, parametros=()):
            vistos.append(" ".join(str(sql).split()))
            return original(sql, parametros)

        conn.execute = execute
        return conn

    monkeypatch.setattr(database, "obtener_conexion", espia)
    return vistos


def test_no_hay_credenciales_turso():
    assert not os.environ.get("TURSO_DATABASE_URL")
    assert not os.environ.get("TURSO_AUTH_TOKEN")


def test_carga_el_pin_vacio_y_no_rehashea(
    db_local, monkeypatch, pbkdf2
):
    with database.obtener_conexion() as conn:
        conn.execute(
            "INSERT INTO usuarios_oficina (usuario, nombre, pin_hash, es_admin, activo, intentos_fallidos) "
            "VALUES ('admin', 'Administrador', '', 1, 1, 3)"
        )
        conn.commit()
    _secrets(monkeypatch, [_admin()])
    antes = pbkdf2["n"]
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 1
    # Hash vacío: no verifica (no hay salt) y hashea una sola vez para guardarlo.
    assert pbkdf2["n"] - antes == 1
    guardado = _fila("admin")["pin_hash"]
    assert guardado.startswith("pbkdf2$")
    assert _fila("admin")["intentos_fallidos"] == 3
    assert auth_oficina.verificar_login_oficina("admin", "654321")["usuario"] == "admin"
    assert _fila("admin")["intentos_fallidos"] == 0

    auth_oficina._limpiar_intentos_fallidos(int(_fila("admin")["id"]))
    with database.obtener_conexion() as conn:
        conn.execute(
            "UPDATE usuarios_oficina SET intentos_fallidos = 3 WHERE usuario = 'admin'"
        )
        conn.commit()
    auth_oficina._SECRETS_USUARIOS_APLICADOS = False
    antes = pbkdf2["n"]
    sql = _sqls(monkeypatch)
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 0
    assert pbkdf2["n"] == antes
    assert _fila("admin")["pin_hash"] == guardado
    assert _fila("admin")["intentos_fallidos"] == 3
    texto = " ".join(sql).upper()
    assert "UPDATE" not in texto
    assert "INSERT" not in texto
    assert "CREATE" not in texto


def test_una_vez_por_proceso_no_vuelve_a_entrar(db_local, monkeypatch, pbkdf2):
    _secrets(monkeypatch, [_admin(), {"usuario": "socio", "nombre": "Socio", "pin": "1111", "es_admin": False}])
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 2
    hasheos = pbkdf2["n"]
    assert hasheos == 2

    def boom():
        raise AssertionError("segunda corrida no debería tocar la base")

    monkeypatch.setattr(database, "obtener_conexion", boom)
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 0
    assert auth_oficina.aplicar_usuarios_desde_secrets_en_login() is None
    assert pbkdf2["n"] == hasheos


def test_secrets_vacios_no_traban_la_carga_siguiente(db_local, monkeypatch, pbkdf2):
    _secrets(monkeypatch, [])
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 0
    assert auth_oficina._SECRETS_USUARIOS_APLICADOS is False
    _secrets(monkeypatch, [_admin()])
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 1
    assert _fila("admin")["pin_hash"].startswith("pbkdf2$")


def test_no_pisa_al_equipo_ni_a_otro_usuario(db_local, monkeypatch, pbkdf2):
    with database.obtener_conexion() as conn:
        conn.execute(
            "INSERT INTO usuarios_oficina (usuario, nombre, pin_hash, es_admin, activo, intentos_fallidos, bloqueado_hasta) "
            "VALUES ('guada', 'Guadi', 'hash-guada', 0, 1, 4, '2099-01-01T00:00:00')"
        )
        conn.commit()
    _secrets(
        monkeypatch,
        [
            _admin(),
            {"usuario": "guada", "nombre": "Otra", "pin": "9999", "es_admin": True},
        ],
    )
    auth_oficina._aplicar_usuarios_desde_secrets()
    guada = _fila("guada")
    assert guada["pin_hash"] == "hash-guada"
    assert guada["nombre"] == "Guadi"
    assert guada["intentos_fallidos"] == 4
    assert guada["bloqueado_hasta"] == "2099-01-01T00:00:00"
    assert guada["es_admin"] == 0


def test_login_solo_cuenta_el_usuario_elegido(db_local, monkeypatch, pbkdf2):
    _secrets(
        monkeypatch,
        [
            _admin(pin="654321"),
            {"usuario": "socio", "nombre": "Socio", "pin": "1111", "es_admin": False},
        ],
    )
    auth_oficina._aplicar_usuarios_desde_secrets()
    assert auth_oficina.verificar_login_oficina("admin", "0000") is None
    assert _fila("admin")["intentos_fallidos"] == 1
    assert _fila("socio")["intentos_fallidos"] == 0
    assert auth_oficina.verificar_login_oficina("socio", "1111")["usuario"] == "socio"
    assert _fila("admin")["intentos_fallidos"] == 1
    assert _fila("socio")["intentos_fallidos"] == 0


def test_actualizar_nombre_no_toca_bloqueo_ni_rehashea(db_local, monkeypatch, pbkdf2):
    auth_oficina.crear_usuario_oficina("admin", "Administrador", pin="654321", es_admin=True)
    auth_oficina.crear_usuario_oficina("socio", "Socio", pin="1111", es_admin=False)
    with database.obtener_conexion() as conn:
        conn.execute(
            "UPDATE usuarios_oficina SET intentos_fallidos = 4, bloqueado_hasta = ?, activo = 0 "
            "WHERE usuario = 'admin'",
            ("2099-06-01T12:00:00",),
        )
        conn.commit()
    pin_hash = _fila("admin")["pin_hash"]
    socio_hash = _fila("socio")["pin_hash"]
    _secrets(monkeypatch, [_admin(nombre="Admin Nuevo"), {"usuario": "socio", "nombre": "Socio", "pin": "1111", "es_admin": False}])
    antes = pbkdf2["n"]
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 1
    # La huella no estaba: verifica el PIN (1 PBKDF2) y no genera otro salt.
    assert pbkdf2["n"] - antes == 2  # admin + socio, solo verificar
    admin = _fila("admin")
    assert admin["nombre"] == "Admin Nuevo"
    assert admin["pin_hash"] == pin_hash
    assert admin["intentos_fallidos"] == 4
    assert admin["bloqueado_hasta"] == "2099-06-01T12:00:00"
    assert admin["activo"] == 1
    assert _fila("socio")["pin_hash"] == socio_hash
    assert _fila("socio")["intentos_fallidos"] == 0
    assert auth_oficina.usuario_bloqueado_oficina("admin") > 0
    assert auth_oficina.verificar_login_oficina("admin", "654321") is None


def test_cambio_de_pin_en_secrets(db_local, monkeypatch, pbkdf2):
    _secrets(monkeypatch, [_admin(pin="654321")])
    auth_oficina._aplicar_usuarios_desde_secrets()
    anterior = _fila("admin")["pin_hash"]
    auth_oficina._SECRETS_USUARIOS_APLICADOS = False
    _secrets(monkeypatch, [_admin(pin="7777")])
    assert auth_oficina._aplicar_usuarios_desde_secrets() == 1
    nuevo = _fila("admin")["pin_hash"]
    assert nuevo != anterior
    assert auth_oficina.verificar_login_oficina("admin", "7777") is not None
    assert auth_oficina.verificar_login_oficina("admin", "654321") is None


def test_fallo_queda_en_el_log_sin_el_pin(monkeypatch, caplog, capsys, pbkdf2):
    pin = "654321"
    _secrets(monkeypatch, [_admin(pin=pin)])

    def boom():
        raise RuntimeError("conexion rechazada")

    monkeypatch.setattr(database, "obtener_conexion", boom)
    with caplog.at_level(logging.ERROR):
        auth_oficina.aplicar_usuarios_desde_secrets_en_login()
    salida = capsys.readouterr()
    completo = caplog.text + salida.out + salida.err
    assert pin not in completo
    assert "oficina_usuarios" not in salida.out
    assert "No se pudieron aplicar los usuarios de Secrets" in caplog.text
    assert "No se pudieron aplicar los usuarios de Secrets" in salida.out
    assert "RuntimeError" in salida.out
    assert "conexion rechazada" not in salida.out
    # Segunda corrida del mismo proceso: no reintenta ni vuelve a hashear.
    assert pbkdf2["n"] == 0
    caplog.clear()
    auth_oficina.aplicar_usuarios_desde_secrets_en_login()
    assert caplog.text == ""
    assert pbkdf2["n"] == 0
