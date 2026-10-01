"""Cola de rutinas de oficina.

La web solo anota el pedido. Lo ejecuta un asistente externo (scripts/cola_rutinas.py)
que lee la misma base: Turso si hay TURSO_DATABASE_URL y TURSO_AUTH_TOKEN, y si no
el SQLite local.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import database

# Catálogo editable. `codigo` es lo que se guarda en rutina_pedidos.rutina.
RUTINAS: tuple[dict[str, str], ...] = (
    {
        "codigo": "seguimiento_balances_urgencia",
        "nombre": "Seguimiento balances urgencia",
        "descripcion": "Actualiza Seguimiento_Balances.xlsx en el servidor marcando urgencias.",
    },
    {
        "codigo": "control_fcc_portal_iva",
        "nombre": "Control FCC Portal IVA",
        "descripcion": "Revisa en ARCA Portal IVA las facturas de los clientes FCC.",
    },
    {
        "codigo": "fcc_monotributistas",
        "nombre": "FCC monotributistas",
        "descripcion": "Control mensual de monotributistas FCC.",
    },
    {
        "codigo": "aviso_bazan_bajar_archivos",
        "nombre": "Aviso Bazan bajar archivos",
        "descripcion": "Recordatorio para bajar archivos de Bazan.",
    },
    {
        "codigo": "bazan_detalle_items",
        "nombre": "Bazan Detalle Items",
        "descripcion": "Arma el detalle de ítems de Bazan.",
    },
)

ESTADOS = ("PENDIENTE", "EN_CURSO", "OK", "ERROR", "CANCELADO")
ESTADOS_ABIERTOS = ("PENDIENTE", "EN_CURSO")
ESTADOS_CIERRE = ("OK", "ERROR")

_POR_CODIGO = {item["codigo"]: item for item in RUTINAS}

try:
    _TZ_AR = ZoneInfo("America/Buenos_Aires")
except Exception:
    _TZ_AR = timezone(timedelta(hours=-3))

_DDL = (
    """
    CREATE TABLE IF NOT EXISTS rutina_pedidos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rutina TEXT NOT NULL,
        parametros TEXT,
        solicitado_por TEXT NOT NULL,
        creado_en TEXT NOT NULL,
        estado TEXT NOT NULL CHECK (
            estado IN ('PENDIENTE', 'EN_CURSO', 'OK', 'ERROR', 'CANCELADO')
        ),
        tomado_en TEXT,
        terminado_en TEXT,
        resultado TEXT,
        archivos TEXT
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS idx_rutina_pedidos_estado
    ON rutina_pedidos(estado, id DESC)
    """,
)


class ErrorRutina(Exception):
    """Error de la cola (rutina desconocida, nombre vacío, estado inválido)."""


class PedidoDuplicado(ErrorRutina):
    """Ya hay un pedido PENDIENTE o EN_CURSO de esa rutina."""

    def __init__(self, rutina: str, pedido_id: int, estado: str) -> None:
        self.rutina = rutina
        self.pedido_id = int(pedido_id)
        self.estado = estado
        nombre = nombre_rutina(rutina)
        super().__init__(
            f"Ya hay un pedido {estado} (n.º {pedido_id}) de {nombre}. No se creó otro."
        )


class PedidoNoEncontrado(ErrorRutina):
    """No existe ese id."""


class EstadoInvalido(ErrorRutina):
    """El pedido no está en el estado que exige la operación."""


def nombre_rutina(codigo: str) -> str:
    item = _POR_CODIGO.get(str(codigo or "").strip())
    if not item:
        return str(codigo or "").strip()
    return item["nombre"]


def rutina_por_codigo(codigo: str) -> dict[str, str] | None:
    item = _POR_CODIGO.get(str(codigo or "").strip())
    return dict(item) if item else None


def ahora_utc() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def formatear_fecha_ar(valor: str | None) -> str:
    """ISO (UTC) → DD/MM/YYYY HH:MM en America/Buenos_Aires."""
    texto = str(valor or "").strip()
    if not texto:
        return ""
    try:
        momento = datetime.fromisoformat(texto.replace("Z", "+00:00"))
    except ValueError:
        return texto
    if momento.tzinfo is None:
        momento = momento.replace(tzinfo=timezone.utc)
    return momento.astimezone(_TZ_AR).strftime("%d/%m/%Y %H:%M")


def inicializar_tablas_rutinas(conn) -> None:
    """Crea la tabla y los índices si no existen. Llamado desde database.inicializar_bd()."""
    for sql in _DDL:
        conn.execute(sql)
    _asegurar_indice_abierto(conn)


def asegurar_tablas() -> None:
    with database.obtener_conexion() as conn:
        inicializar_tablas_rutinas(conn)
        conn.commit()


def _asegurar_indice_abierto(conn) -> None:
    """Un solo pedido abierto por rutina. Si ya hay duplicados, no rompe el arranque."""
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


def _a_dict(fila) -> dict:
    datos = {clave: fila[clave] for clave in fila.keys()}
    datos["nombre"] = nombre_rutina(str(datos.get("rutina") or ""))
    return datos


def _validar_codigo(rutina: str) -> str:
    codigo = str(rutina or "").strip()
    if codigo not in _POR_CODIGO:
        raise ErrorRutina(f"Rutina desconocida: {codigo or '(vacía)'}.")
    return codigo


def _validar_quien(solicitado_por: str) -> str:
    quien = str(solicitado_por or "").strip()
    if not quien:
        raise ErrorRutina("Indicá quién pide la rutina.")
    return quien


def pedido_abierto(rutina: str) -> dict | None:
    """Pedido PENDIENTE o EN_CURSO de esa rutina, si hay uno."""
    codigo = str(rutina or "").strip()
    if not codigo:
        return None
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            SELECT * FROM rutina_pedidos
            WHERE rutina = ? AND estado IN ('PENDIENTE', 'EN_CURSO')
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        ).fetchone()
    return _a_dict(fila) if fila else None


def crear_pedido(rutina: str, solicitado_por: str, parametros: str | None = None) -> dict:
    """Encola un pedido. Falla si esa rutina ya tiene uno PENDIENTE o EN_CURSO."""
    codigo = _validar_codigo(rutina)
    quien = _validar_quien(solicitado_por)
    params = str(parametros or "").strip() or None
    abierto = pedido_abierto(codigo)
    if abierto:
        raise PedidoDuplicado(codigo, int(abierto["id"]), str(abierto["estado"]))
    asegurar_tablas()
    try:
        with database.obtener_conexion() as conn:
            otra = conn.execute(
                """
                SELECT id, estado FROM rutina_pedidos
                WHERE rutina = ? AND estado IN ('PENDIENTE', 'EN_CURSO')
                """,
                (codigo,),
            ).fetchone()
            if otra:
                raise PedidoDuplicado(codigo, int(otra["id"]), str(otra["estado"]))
            fila = conn.execute(
                """
                INSERT INTO rutina_pedidos (
                    rutina, parametros, solicitado_por, creado_en, estado
                ) VALUES (?, ?, ?, ?, 'PENDIENTE')
                RETURNING *
                """,
                (codigo, params, quien, ahora_utc()),
            ).fetchone()
            conn.commit()
    except sqlite3.IntegrityError as exc:
        abierto = pedido_abierto(codigo)
        if abierto:
            raise PedidoDuplicado(codigo, int(abierto["id"]), str(abierto["estado"])) from exc
        raise ErrorRutina("No se pudo crear el pedido.") from exc
    if fila is None:
        raise ErrorRutina("No se pudo crear el pedido.")
    return _a_dict(fila)


def obtener_pedido(pedido_id: int) -> dict | None:
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT * FROM rutina_pedidos WHERE id = ?",
            (int(pedido_id),),
        ).fetchone()
    return _a_dict(fila) if fila else None


def listar_pedidos(*, estado: str | None = None, limite: int = 50) -> list[dict]:
    if estado is not None and estado not in ESTADOS:
        raise ErrorRutina(f"Estado desconocido: {estado}.")
    tope = max(1, int(limite))
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        if estado:
            filas = conn.execute(
                """
                SELECT * FROM rutina_pedidos
                WHERE estado = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (estado, tope),
            ).fetchall()
        else:
            filas = conn.execute(
                """
                SELECT * FROM rutina_pedidos
                ORDER BY id DESC
                LIMIT ?
                """,
                (tope,),
            ).fetchall()
    return [_a_dict(fila) for fila in filas]


def cancelar_pedido(pedido_id: int) -> dict:
    """Pasa a CANCELADO solo si estaba PENDIENTE."""
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            UPDATE rutina_pedidos
            SET estado = 'CANCELADO',
                terminado_en = ?,
                resultado = COALESCE(NULLIF(resultado, ''), 'Cancelado desde la web.')
            WHERE id = ? AND estado = 'PENDIENTE'
            RETURNING *
            """,
            (ahora_utc(), int(pedido_id)),
        ).fetchone()
        if fila is None:
            _raise_estado(conn, int(pedido_id), "Solo se puede cancelar un pedido PENDIENTE.")
        conn.commit()
    return _a_dict(fila)


def tomar_pedido(pedido_id: int) -> dict:
    """Pasa a EN_CURSO solo si estaba PENDIENTE. El UPDATE es la operación atómica."""
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            UPDATE rutina_pedidos
            SET estado = 'EN_CURSO', tomado_en = ?
            WHERE id = ? AND estado = 'PENDIENTE'
            RETURNING *
            """,
            (ahora_utc(), int(pedido_id)),
        ).fetchone()
        if fila is None:
            _raise_estado(conn, int(pedido_id), "Solo se puede tomar un pedido PENDIENTE.")
        conn.commit()
    return _a_dict(fila)


def terminar_pedido(
    pedido_id: int,
    estado: str,
    resultado: str,
    archivos: str | None = None,
) -> dict:
    """Cierra un pedido EN_CURSO en OK o ERROR."""
    cierre = str(estado or "").strip().upper()
    if cierre not in ESTADOS_CIERRE:
        raise EstadoInvalido("El cierre tiene que ser OK o ERROR.")
    texto = str(resultado or "").strip()
    if not texto:
        raise ErrorRutina("El resultado no puede estar vacío.")
    adjuntos = str(archivos or "").strip() or None
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            UPDATE rutina_pedidos
            SET estado = ?, resultado = ?, archivos = ?, terminado_en = ?
            WHERE id = ? AND estado = 'EN_CURSO'
            RETURNING *
            """,
            (cierre, texto, adjuntos, ahora_utc(), int(pedido_id)),
        ).fetchone()
        if fila is None:
            _raise_estado(conn, int(pedido_id), "Solo se puede terminar un pedido EN_CURSO.")
        conn.commit()
    return _a_dict(fila)


def _raise_estado(conn, pedido_id: int, mensaje: str) -> None:
    existe = conn.execute(
        "SELECT estado FROM rutina_pedidos WHERE id = ?",
        (pedido_id,),
    ).fetchone()
    if existe is None:
        raise PedidoNoEncontrado(f"No existe el pedido {pedido_id}.")
    raise EstadoInvalido(f"{mensaje} El pedido {pedido_id} está en {existe['estado']}.")
