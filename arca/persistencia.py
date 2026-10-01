"""CAE, ticket WSAA y datos del emisor en la base del estudio.

Usa database.obtener_conexion(): Turso si hay TURSO_DATABASE_URL y
TURSO_AUTH_TOKEN (Streamlit Cloud) y, si no, el SQLite local. Así un redeploy
no borra los números emitidos ni el ticket vigente.
"""
from __future__ import annotations

import datetime as dt

import database

_DDL = (
    """
    CREATE TABLE IF NOT EXISTS arca_emisiones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_hora TEXT,
        ambiente TEXT NOT NULL,
        modo TEXT,
        filas TEXT,
        id_factura TEXT,
        cuit_emisor TEXT,
        pto_vta TEXT,
        tipo TEXT,
        numero TEXT,
        cae TEXT,
        vto_cae TEXT,
        estado TEXT,
        mensajes TEXT,
        total REAL,
        pdf_nombre TEXT,
        huella TEXT,
        fecha_cbte TEXT,
        cbte_asociado TEXT,
        payload_json TEXT,
        creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS arca_emisores (
        cuit TEXT PRIMARY KEY,
        razon_social TEXT,
        domicilio TEXT,
        condicion_iva TEXT,
        iibb TEXT,
        inicio_actividades TEXT,
        actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS arca_ta (
        ambiente TEXT NOT NULL,
        servicio TEXT NOT NULL,
        cert_hash TEXT NOT NULL,
        token TEXT NOT NULL,
        sign TEXT NOT NULL,
        expiration TEXT NOT NULL,
        generation TEXT,
        destination TEXT,
        actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (ambiente, servicio, cert_hash)
    )
    """,
    "CREATE INDEX IF NOT EXISTS idx_arca_emisiones_amb_huella ON arca_emisiones(ambiente, huella)",
    "CREATE INDEX IF NOT EXISTS idx_arca_emisiones_cuit ON arca_emisiones(cuit_emisor, ambiente)",
)


def inicializar_tablas_arca(conn) -> None:
    """Crea las tablas si no existen. Llamado desde database.inicializar_bd()."""
    for sql in _DDL:
        conn.execute(sql)


def asegurar_tablas() -> None:
    with database.obtener_conexion() as conn:
        inicializar_tablas_arca(conn)
        conn.commit()


def _txt(v):
    if v is None:
        return None
    return str(v)


def insertar_emision(fila: list, payload_json: str | None = None) -> None:
    """Fila en el orden de emitir.RES_COLS. No duplica el mismo evento."""
    asegurar_tablas()
    while len(fila) < 18:
        fila = list(fila) + [None]
    fecha_hora, ambiente, modo, filas, id_factura, cuit, pto, tipo, numero, cae, vto, estado, mensajes, total, pdf, huella, fecha_cbte, asoc = fila[:18]
    with database.obtener_conexion() as conn:
        ya = conn.execute(
            """
            SELECT id FROM arca_emisiones
            WHERE ambiente = ? AND ifnull(huella, '') = ? AND ifnull(estado, '') = ?
              AND ifnull(numero, '') = ? AND ifnull(fecha_hora, '') = ?
            """,
            (ambiente, _txt(huella) or "", _txt(estado) or "", _txt(numero) or "", _txt(fecha_hora) or ""),
        ).fetchone()
        if ya:
            return
        conn.execute(
            """
            INSERT INTO arca_emisiones (
                fecha_hora, ambiente, modo, filas, id_factura, cuit_emisor, pto_vta, tipo,
                numero, cae, vto_cae, estado, mensajes, total, pdf_nombre, huella,
                fecha_cbte, cbte_asociado, payload_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                _txt(fecha_hora), ambiente, _txt(modo), _txt(filas), _txt(id_factura), _txt(cuit),
                _txt(pto), _txt(tipo), _txt(numero), _txt(cae), _txt(vto), _txt(estado),
                _txt(mensajes), float(total) if total not in (None, "") else None,
                _txt(pdf), _txt(huella), _txt(fecha_cbte), _txt(asoc), payload_json,
            ),
        )
        conn.commit()


def emisiones_que_bloquean(ambiente: str) -> list[dict]:
    """APROBADO y VERIFICAR de un ambiente, para no reenviar y para asociar NC."""
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        filas = conn.execute(
            """
            SELECT fecha_hora, ambiente, modo, filas, id_factura, cuit_emisor, pto_vta, tipo,
                   numero, cae, vto_cae, estado, mensajes, total, pdf_nombre, huella,
                   fecha_cbte, cbte_asociado
            FROM arca_emisiones
            WHERE ambiente = ? AND (estado LIKE 'APROBADO%' OR estado LIKE 'VERIFICAR%')
            ORDER BY id
            """,
            (ambiente,),
        ).fetchall()
    return [dict(f) for f in filas]


def listar_emisiones(cuit: str = "", ambiente: str = "", limite: int = 200) -> list[dict]:
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        filas = conn.execute(
            """
            SELECT id, fecha_hora, ambiente, modo, filas, id_factura, cuit_emisor, pto_vta, tipo,
                   numero, cae, vto_cae, estado, mensajes, total, huella, fecha_cbte,
                   cbte_asociado, payload_json
            FROM arca_emisiones
            WHERE (? = '' OR cuit_emisor = ?) AND (? = '' OR ambiente = ?)
            ORDER BY id DESC
            LIMIT ?
            """,
            (cuit, cuit, ambiente, ambiente, int(limite)),
        ).fetchall()
    return [dict(f) for f in filas]


def obtener_emision(emision_id: int) -> dict | None:
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute("SELECT * FROM arca_emisiones WHERE id = ?", (int(emision_id),)).fetchone()
    return dict(fila) if fila else None


def guardar_emisor(cuit: str, razon_social: str = "", domicilio: str = "", condicion_iva: str = "",
                   iibb: str = "", inicio_actividades: str = "") -> None:
    asegurar_tablas()
    ahora = dt.datetime.now().isoformat(timespec="seconds")
    with database.obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO arca_emisores (cuit, razon_social, domicilio, condicion_iva, iibb, inicio_actividades, actualizado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(cuit) DO UPDATE SET
                razon_social = excluded.razon_social,
                domicilio = excluded.domicilio,
                condicion_iva = excluded.condicion_iva,
                iibb = excluded.iibb,
                inicio_actividades = excluded.inicio_actividades,
                actualizado_en = excluded.actualizado_en
            """,
            (cuit, razon_social or "", domicilio or "", condicion_iva or "", iibb or "", inicio_actividades or "", ahora),
        )
        conn.commit()


def cargar_emisor(cuit: str) -> dict | None:
    if not cuit:
        return None
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute("SELECT * FROM arca_emisores WHERE cuit = ?", (cuit,)).fetchone()
    return dict(fila) if fila else None


def listar_emisores() -> list[dict]:
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT cuit, razon_social, domicilio, condicion_iva, iibb, inicio_actividades FROM arca_emisores ORDER BY razon_social"
        ).fetchall()
    return [dict(f) for f in filas]


def leer_ta(ambiente: str, servicio: str, cert_hash: str) -> dict | None:
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            SELECT token, sign, expiration, generation, destination
            FROM arca_ta WHERE ambiente = ? AND servicio = ? AND cert_hash = ?
            """,
            (ambiente, servicio, cert_hash),
        ).fetchone()
    return dict(fila) if fila else None


def guardar_ta(ambiente: str, servicio: str, cert_hash: str, ta: dict) -> None:
    asegurar_tablas()
    ahora = dt.datetime.now().isoformat(timespec="seconds")
    with database.obtener_conexion() as conn:
        conn.execute(
            """
            INSERT INTO arca_ta (ambiente, servicio, cert_hash, token, sign, expiration, generation, destination, actualizado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(ambiente, servicio, cert_hash) DO UPDATE SET
                token = excluded.token,
                sign = excluded.sign,
                expiration = excluded.expiration,
                generation = excluded.generation,
                destination = excluded.destination,
                actualizado_en = excluded.actualizado_en
            """,
            (
                ambiente, servicio, cert_hash, ta.get("token") or "", ta.get("sign") or "",
                ta.get("expiration") or "", ta.get("generation"), ta.get("destination"), ahora,
            ),
        )
        conn.commit()
