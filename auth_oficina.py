"""Autenticacion y gestion de usuarios de la oficina (login, PIN, roles).

Primer corte de un refactor mas grande: separa el dominio "quien puede
entrar a la app" (login, hash de PIN, rate limiting, alta/baja de usuarios)
del resto de la capa de datos en database.py (clientes, sueldos, balances).

Usa la misma conexion SQLite que database.py (database.obtener_conexion()).
La tabla usuarios_oficina se crea/migra desde inicializar_tabla_usuarios_oficina(),
llamada por database.inicializar_bd().

Import local (no "from database import ...") a proposito: evita import
circular, ya que database.py llama a funciones de este modulo en runtime.
"""

from __future__ import annotations

import re
from pathlib import Path

import database


def inicializar_tabla_usuarios_oficina(conn) -> None:
    """Crea/migra la tabla usuarios_oficina. Llamado desde database.inicializar_bd()."""
    import sqlite3

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS usuarios_oficina (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            nombre TEXT NOT NULL,
            pin_hash TEXT NOT NULL DEFAULT '',
            es_admin INTEGER NOT NULL DEFAULT 0,
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    # Migracion no destructiva: rate limiting de login (intentos fallidos / bloqueo temporal)
    try:
        conn.execute(
            "ALTER TABLE usuarios_oficina ADD COLUMN intentos_fallidos INTEGER NOT NULL DEFAULT 0"
        )
    except sqlite3.OperationalError:
        pass  # columna ya existe
    try:
        conn.execute("ALTER TABLE usuarios_oficina ADD COLUMN bloqueado_hasta TEXT")
    except sqlite3.OperationalError:
        pass  # columna ya existe


def _hash_pin(pin: str, salt: str | None = None) -> str:
    """Hash de PIN con salt aleatorio por usuario (PBKDF2-HMAC-SHA256, 600k iteraciones).

    Formato guardado: "pbkdf2$<salt_hex>$<hash_hex>". Un PIN vacío se guarda
    como "" (sin PIN → login libre solo permitido en entorno local).
    """
    import hashlib
    import secrets as _secrets_mod

    texto = str(pin or "").strip()
    if not texto:
        return ""
    salt_hex = salt or _secrets_mod.token_hex(16)
    derivado = hashlib.pbkdf2_hmac(
        "sha256", texto.encode("utf-8"), bytes.fromhex(salt_hex), 600_000
    )
    return f"pbkdf2${salt_hex}${derivado.hex()}"


def _verificar_pin(pin: str, pin_hash_guardado: str) -> tuple[bool, str | None]:
    """Compara un PIN contra el hash guardado.

    Soporta el formato legacy (sha256 sin salt, sin "$") para no invalidar
    logins existentes: si el PIN matchea el hash legacy, devuelve
    (True, nuevo_hash) para que el llamador migre ese registro al formato
    con salt en el mismo paso (migración transparente, sin pedir reset).
    """
    import hashlib
    import hmac

    guardado = str(pin_hash_guardado or "")
    texto = str(pin or "").strip()
    if not guardado:
        return False, None
    if guardado.startswith("pbkdf2$"):
        try:
            _, salt_hex, _hash_hex = guardado.split("$", 2)
        except ValueError:
            return False, None
        candidato = _hash_pin(texto, salt=salt_hex)
        return hmac.compare_digest(candidato, guardado), None
    # Legado: sha256 sin salt (versiones anteriores). Si matchea, se migra.
    legacy = hashlib.sha256(texto.encode("utf-8")).hexdigest()
    if hmac.compare_digest(legacy, guardado):
        return True, _hash_pin(texto)
    return False, None


_MAX_INTENTOS_LOGIN = 5
_BLOQUEO_LOGIN_MINUTOS = 15


def _segundos_bloqueo_restantes(fila: dict) -> int:
    """Segundos que faltan para que se levante el bloqueo por intentos fallidos (0 si no aplica)."""
    from datetime import datetime

    bloqueado_hasta = fila.get("bloqueado_hasta")
    if not bloqueado_hasta:
        return 0
    try:
        limite = datetime.fromisoformat(str(bloqueado_hasta))
    except ValueError:
        return 0
    return max(0, int((limite - datetime.now()).total_seconds()))


def _registrar_intento_fallido(usuario_id: int) -> None:
    """Suma un intento fallido; bloquea temporalmente al llegar al máximo."""
    from datetime import datetime, timedelta

    with database.obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT intentos_fallidos FROM usuarios_oficina WHERE id = ?",
            (usuario_id,),
        ).fetchone()
        intentos = (int(fila["intentos_fallidos"] or 0) if fila else 0) + 1
        bloqueado_hasta = None
        if intentos >= _MAX_INTENTOS_LOGIN:
            bloqueado_hasta = (
                datetime.now() + timedelta(minutes=_BLOQUEO_LOGIN_MINUTOS)
            ).isoformat()
        conn.execute(
            "UPDATE usuarios_oficina SET intentos_fallidos = ?, bloqueado_hasta = ? WHERE id = ?",
            (intentos, bloqueado_hasta, usuario_id),
        )
        conn.commit()


def _limpiar_intentos_fallidos(usuario_id: int) -> None:
    with database.obtener_conexion() as conn:
        conn.execute(
            "UPDATE usuarios_oficina SET intentos_fallidos = 0, bloqueado_hasta = NULL WHERE id = ?",
            (usuario_id,),
        )
        conn.commit()


def usuario_bloqueado_oficina(usuario: str) -> int:
    """Segundos restantes de bloqueo por intentos fallidos (0 si el usuario puede intentar)."""
    fila = obtener_usuario_oficina(usuario)
    if not fila:
        return 0
    return _segundos_bloqueo_restantes(fila)


def _secrets_oficina_usuarios() -> list[dict]:
    """Lee usuarios desde st.secrets['oficina_usuarios'] (Cloud / local secrets.toml)."""
    try:
        import streamlit as st

        bloque = st.secrets.get("oficina_usuarios")
    except Exception:
        return []
    if bloque is None:
        return []
    out: list[dict] = []
    try:
        items = dict(bloque)
    except Exception:
        return []
    for usuario, meta in items.items():
        try:
            if hasattr(meta, "get"):
                nombre = str(meta.get("nombre") or usuario).strip()
                pin = str(meta.get("pin") or meta.get("password") or "").strip()
                es_admin = bool(meta.get("es_admin", str(usuario).lower() == "admin"))
            else:
                # Forma corta: usuario = "pin_en_texto"
                nombre = str(usuario)
                pin = str(meta or "").strip()
                es_admin = str(usuario).lower() == "admin"
            out.append(
                {
                    "usuario": str(usuario).strip().lower(),
                    "nombre": nombre,
                    "pin": pin,
                    "es_admin": es_admin,
                }
            )
        except Exception:
            continue
    return out


def _aplicar_usuarios_desde_secrets() -> int:
    """
    Crea/actualiza usuarios definidos en Secrets (PIN hasheado en SQLite).
    Devuelve cantidad de usuarios aplicados. No guarda PIN en texto claro en el repo.
    """
    definidos = _secrets_oficina_usuarios()
    if not definidos:
        return 0
    aplicados = 0
    for u in definidos:
        user = u["usuario"]
        if not user:
            continue
        existente = obtener_usuario_oficina(user)
        if existente is None:
            try:
                crear_usuario_oficina(
                    user,
                    u["nombre"],
                    pin=u.get("pin") or "",
                    es_admin=bool(u.get("es_admin")),
                )
                aplicados += 1
                continue
            except ValueError:
                existente = obtener_usuario_oficina(user)
                if existente is None:
                    continue
        kwargs: dict = {
            "nombre": u["nombre"],
            "es_admin": bool(u.get("es_admin")),
            "activo": True,
        }
        if u.get("pin"):
            kwargs["pin"] = u["pin"]
        actualizar_usuario_oficina(int(existente["id"]), **kwargs)
        aplicados += 1
    return aplicados


def sembrar_usuarios_oficina_default() -> None:
    """Crea usuarios iniciales si la tabla está vacía; prioriza Secrets en Cloud."""
    _aplicar_usuarios_desde_secrets()
    existentes = listar_usuarios_oficina(solo_activos=False)
    if existentes:
        return
    defaults = [
        ("admin", "Administrador", "", True),
        ("recepcion", "Recepción", "", False),
        ("contador", "Contador", "", False),
        ("auxiliar", "Auxiliar contable", "", False),
    ]
    for usuario, nombre, pin, admin in defaults:
        try:
            crear_usuario_oficina(usuario, nombre, pin=pin, es_admin=admin)
        except ValueError:
            # Ya existe (carrera con Secrets / otro worker) — no romper el arranque
            continue


def listar_usuarios_oficina(solo_activos: bool = True) -> list[dict]:
    with database.obtener_conexion() as conn:
        if solo_activos:
            filas = conn.execute(
                "SELECT id, usuario, nombre, es_admin, activo FROM usuarios_oficina "
                "WHERE activo = 1 ORDER BY nombre COLLATE NOCASE"
            ).fetchall()
        else:
            filas = conn.execute(
                "SELECT id, usuario, nombre, es_admin, activo FROM usuarios_oficina "
                "ORDER BY nombre COLLATE NOCASE"
            ).fetchall()
    return [dict(f) for f in filas]


def obtener_usuario_oficina(usuario: str) -> dict | None:
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT id, usuario, nombre, pin_hash, es_admin, activo, "
            "intentos_fallidos, bloqueado_hasta "
            "FROM usuarios_oficina WHERE lower(usuario) = lower(?)",
            (str(usuario or "").strip(),),
        ).fetchone()
    return dict(fila) if fila else None


def crear_usuario_oficina(
    usuario: str,
    nombre: str,
    *,
    pin: str = "",
    es_admin: bool = False,
) -> int:
    import sqlite3

    user = re.sub(r"[^a-zA-Z0-9._-]", "", str(usuario or "").strip().lower())
    nom = str(nombre or "").strip()
    if not user or not nom:
        raise ValueError("Usuario y nombre son obligatorios.")
    if obtener_usuario_oficina(user) is not None:
        raise ValueError(
            f"El usuario «{user}» ya existe. Elegí otro login o editá el existente."
        )
    try:
        with database.obtener_conexion() as conn:
            cur = conn.execute(
                "INSERT INTO usuarios_oficina (usuario, nombre, pin_hash, es_admin, activo) "
                "VALUES (?, ?, ?, ?, 1)",
                (user, nom, _hash_pin(pin), 1 if es_admin else 0),
            )
            conn.commit()
            return int(cur.lastrowid)
    except sqlite3.IntegrityError as exc:
        raise ValueError(
            f"El usuario «{user}» ya existe. Elegí otro login o editá el existente."
        ) from exc


def actualizar_usuario_oficina(
    usuario_id: int,
    *,
    nombre: str | None = None,
    pin: str | None = None,
    es_admin: bool | None = None,
    activo: bool | None = None,
) -> None:
    campos: list[str] = []
    vals: list = []
    if nombre is not None:
        campos.append("nombre = ?")
        vals.append(str(nombre).strip())
    if pin is not None:
        campos.append("pin_hash = ?")
        vals.append(_hash_pin(pin))
    if es_admin is not None:
        campos.append("es_admin = ?")
        vals.append(1 if es_admin else 0)
    if activo is not None:
        campos.append("activo = ?")
        vals.append(1 if activo else 0)
    if not campos:
        return
    vals.append(int(usuario_id))
    with database.obtener_conexion() as conn:
        conn.execute(
            f"UPDATE usuarios_oficina SET {', '.join(campos)} WHERE id = ?",
            vals,
        )
        conn.commit()


def _exigir_pin_en_entorno() -> bool:
    """En Cloud (o si hay usuarios en Secrets) no permitir entrar sin PIN configurado."""
    if _secrets_oficina_usuarios():
        return True
    import os

    flags = (
        os.environ.get("STREAMLIT_SHARING_MODE"),
        os.environ.get("STREAMLIT_CLOUD"),
        os.environ.get("IS_STREAMLIT_CLOUD"),
    )
    if any(str(f).strip().lower() in {"1", "true", "yes"} for f in flags if f):
        return True
    if Path("/mount/src").is_dir() or Path("/home/appuser").is_dir():
        return True
    return False


def verificar_login_oficina(usuario: str, pin: str = "") -> dict | None:
    """Valida usuario/PIN. Si el usuario no tiene PIN: local OK vacío; Cloud exige Secrets.

    Bloquea temporalmente el usuario tras varios intentos fallidos seguidos
    (protección básica contra fuerza bruta sobre el PIN).
    """
    fila = obtener_usuario_oficina(usuario)
    if not fila or not fila.get("activo"):
        return None

    if _segundos_bloqueo_restantes(fila) > 0:
        return None

    esperado = str(fila.get("pin_hash") or "")
    if esperado:
        ok, nuevo_hash = _verificar_pin(pin, esperado)
        if not ok:
            _registrar_intento_fallido(int(fila["id"]))
            return None
        if nuevo_hash:
            # Migración transparente: el PIN matcheó un hash legacy sin salt.
            with database.obtener_conexion() as conn:
                conn.execute(
                    "UPDATE usuarios_oficina SET pin_hash = ? WHERE id = ?",
                    (nuevo_hash, int(fila["id"])),
                )
                conn.commit()
    else:
        # Sin PIN en DB: en Cloud / con Secrets no dejar puerta abierta
        if _exigir_pin_en_entorno():
            return None
        # Local: sin PIN → entra (igual que antes)

    _limpiar_intentos_fallidos(int(fila["id"]))
    return {
        "id": fila["id"],
        "usuario": fila["usuario"],
        "nombre": fila["nombre"],
        "es_admin": bool(fila.get("es_admin")),
    }


def verificar_admin_oficina(usuario_admin: str, pin_admin: str = "") -> dict | None:
    """Valida que el usuario sea admin activo y el PIN sea correcto."""
    ok = verificar_login_oficina(usuario_admin, pin_admin)
    if not ok or not ok.get("es_admin"):
        return None
    return ok


def resetear_pin_usuario_oficina(
    usuario_objetivo: str,
    nuevo_pin: str,
    *,
    usuario_admin: str,
    pin_admin: str = "",
) -> dict:
    """
    Resetea el PIN de un usuario activo, autorizado por un administrador.
    Devuelve datos del usuario actualizado.
    """
    admin = verificar_admin_oficina(usuario_admin, pin_admin)
    if not admin:
        raise ValueError("Administrador o PIN incorrecto.")

    objetivo = obtener_usuario_oficina(usuario_objetivo)
    if not objetivo or not objetivo.get("activo"):
        raise ValueError("Usuario a recuperar no encontrado o inactivo.")

    nuevo = str(nuevo_pin or "").strip()
    if len(nuevo) < 4:
        raise ValueError("El nuevo PIN debe tener al menos 4 caracteres.")

    actualizar_usuario_oficina(int(objetivo["id"]), pin=nuevo)
    return {
        "id": objetivo["id"],
        "usuario": objetivo["usuario"],
        "nombre": objetivo["nombre"],
        "reseteado_por": admin["usuario"],
    }

