"""Cola de rutinas de oficina.

La web solo anota el pedido. Lo ejecuta un asistente externo (scripts/cola_rutinas.py)
que lee la misma base: Turso si hay TURSO_DATABASE_URL y TURSO_AUTH_TOKEN, y si no
el SQLite local.
"""
from __future__ import annotations

import json
import re
import sqlite3
import unicodedata
from datetime import datetime, timedelta, timezone
from io import BytesIO
from zoneinfo import ZoneInfo

import openpyxl

import database

# Catálogo editable. `codigo` es lo que se guarda en rutina_pedidos.rutina.
# `requisitos` es la checklist de la web: hay que tildarlos todos para ejecutar.
# Van marcados A CONFIRMAR hasta que la oficina ajuste el texto.
# Proyecciones no usa checklist: `requisitos` va vacío y el texto está en `ayuda`.
RUTINAS: tuple[dict, ...] = (
    {
        "codigo": "seguimiento_balances_urgencia",
        "nombre": "Seguimiento balances urgencia",
        "descripcion": "Actualiza Seguimiento_Balances.xlsx en el servidor marcando urgencias.",
        "requisitos": (
            "A CONFIRMAR: Seguimiento_Balances.xlsx está en el servidor y nadie lo tiene abierto.",
            "A CONFIRMAR: ya está definido qué urgencias hay que marcar.",
        ),
    },
    {
        "codigo": "control_fcc_portal_iva",
        "nombre": "Control FCC Portal IVA",
        "descripcion": "Revisa en ARCA Portal IVA las facturas de los clientes FCC.",
        "requisitos": (
            "A CONFIRMAR: el listado de clientes FCC del período está al día.",
            "A CONFIRMAR: el asistente puede entrar a ARCA Portal IVA.",
        ),
    },
    {
        "codigo": "fcc_monotributistas",
        "nombre": "FCC monotributistas",
        "descripcion": "Control mensual de monotributistas FCC.",
        "requisitos": (
            "A CONFIRMAR: el mes a controlar es el que corresponde.",
            "A CONFIRMAR: el padrón de monotributistas FCC está actualizado.",
        ),
    },
    {
        "codigo": "aviso_bazan_bajar_archivos",
        "nombre": "Aviso Bazan bajar archivos",
        "descripcion": "Recordatorio para bajar archivos de Bazan.",
        "requisitos": (
            "A CONFIRMAR: corresponde avisar ahora para bajar los archivos de Bazan.",
            "A CONFIRMAR: hay alguien en la oficina para bajarlos.",
        ),
    },
    {
        "codigo": "bazan_detalle_items",
        "nombre": "Bazan Detalle Items",
        "descripcion": "Arma el detalle de ítems de Bazan.",
        "requisitos": (
            "A CONFIRMAR: los archivos de Bazan de este período ya se bajaron.",
            "A CONFIRMAR: el período del detalle está definido.",
        ),
    },
    {
        "codigo": "proyecciones_ganancias_iva",
        "nombre": "Proyecciones Ganancias por IVA",
        "descripcion": "Arma o actualiza la proyección de Ganancias de cada cliente desde las compras y ventas descargadas.",
        "requisitos": (),
        "ayuda": (
            "Por cada cliente y mes se verifica en el servidor:",
            "1. Listado de imputación contable resumido de compras y ventas del mes.",
            "2. Excel de proyección vigente en su carpeta.",
            "3. PDF DDJJ IIBB del mes.",
            "4. PDF F931 del mes (no aplica si 'NO TIENE EMPLEADOS').",
            "5. TISH del mes solo si TISH=Sí.",
            "6. Si el mes es posterior al cierre del ejercicio, hace falta papel nuevo del ejercicio siguiente.",
        ),
    },
)

CODIGO_PROYECCION = "proyecciones_ganancias_iva"
OPCION_TODOS = "Todos"
ESTADOS_CONTROL = ("OK", "FALTA", "NO APLICA", "FALTA DATO")
COLUMNAS_RESULTADO_PROYECCION = (
    "Cliente",
    "Imputación compras y ventas",
    "Proyección vigente",
    "DDJJ IIBB",
    "F931",
    "TISH",
    "Papel ejercicio siguiente",
    "Estado",
)
# Campo interno, título de la columna en la ficha xlsx / en el JSON exportado.
COLUMNAS_FICHA = (
    ("sociedad", "Sociedad"),
    ("cuit", "CUIT"),
    ("activa", "Activa"),
    ("tish", "TISH"),
    ("mes_inicio", "Mes inicio"),
    ("mes_cierre", "Mes cierre"),
    ("dia_revision", "Día revisión"),
    ("proyeccion_vigente", "Proyección vigente"),
    ("carpeta_proyecciones", "Carpeta Proyecciones"),
    ("carpeta_iibb", "Carpeta IIBB"),
    ("carpeta_f931", "Carpeta F931"),
    ("ultimo_mes_cargado", "Último mes cargado"),
    ("notas", "Notas"),
)
_RE_PERIODO = re.compile(r"^(0[1-9]|1[0-2])-(\d{4})$")
_INACTIVOS = frozenset({"no", "n", "0", "false", "inactiva", "inactivo"})

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
        archivos TEXT,
        preview TEXT
    )
    """,
    """
    CREATE INDEX IF NOT EXISTS idx_rutina_pedidos_estado
    ON rutina_pedidos(estado, id DESC)
    """,
    """
    CREATE TABLE IF NOT EXISTS rutina_ficha_proyecciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sociedad TEXT NOT NULL,
        cuit TEXT,
        activa TEXT,
        tish TEXT,
        mes_inicio TEXT,
        mes_cierre TEXT,
        dia_revision TEXT,
        proyeccion_vigente TEXT,
        carpeta_proyecciones TEXT,
        carpeta_iibb TEXT,
        carpeta_f931 TEXT,
        ultimo_mes_cargado TEXT,
        notas TEXT,
        orden INTEGER NOT NULL DEFAULT 0,
        actualizado_en TEXT
    )
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
    _agregar_columna_preview(conn)
    _asegurar_indice_abierto(conn)


def _agregar_columna_preview(conn) -> None:
    """Suma `preview` en bases que ya tenían la tabla (ALTER idempotente)."""
    columnas = {fila[1] for fila in conn.execute("PRAGMA table_info(rutina_pedidos)")}
    if "preview" in columnas:
        return
    conn.execute("ALTER TABLE rutina_pedidos ADD COLUMN preview TEXT")


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
    datos["preview"] = _preview_desde_texto(datos.get("preview"))
    return datos


def _preview_desde_texto(valor) -> dict | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, dict):
        return valor
    try:
        data = json.loads(valor)
    except (TypeError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    return data


def _celda_preview(valor):
    if isinstance(valor, (dict, list)):
        return json.dumps(valor, ensure_ascii=False)
    return valor


def normalizar_preview(data) -> dict:
    """Deja el preview en {resumen_md, tablas, archivos}. Lanza ErrorRutina si no cierra."""
    if not isinstance(data, dict):
        raise ErrorRutina("El preview tiene que ser un objeto JSON.")
    resumen = data.get("resumen_md")
    if resumen is None:
        resumen = ""
    if not isinstance(resumen, str):
        raise ErrorRutina("resumen_md tiene que ser texto.")
    tablas = data.get("tablas") or []
    if not isinstance(tablas, list):
        raise ErrorRutina("tablas tiene que ser una lista.")
    limpias = []
    for tabla in tablas:
        if not isinstance(tabla, dict):
            raise ErrorRutina("Cada tabla del preview tiene que ser un objeto.")
        columnas = tabla.get("columnas") or []
        filas = tabla.get("filas") or []
        if not isinstance(columnas, list) or not isinstance(filas, list):
            raise ErrorRutina("columnas y filas tienen que ser listas.")
        filas_limpias = []
        for fila in filas:
            if isinstance(fila, (list, tuple)):
                filas_limpias.append([_celda_preview(celda) for celda in fila])
            else:
                filas_limpias.append(fila)
        limpias.append(
            {
                "titulo": str(tabla.get("titulo") or ""),
                "columnas": [str(col) for col in columnas],
                "filas": filas_limpias,
            }
        )
    archivos = data.get("archivos") or []
    if not isinstance(archivos, list):
        raise ErrorRutina("archivos tiene que ser una lista de rutas.")
    return {
        "resumen_md": resumen,
        "tablas": limpias,
        "archivos": [str(ruta) for ruta in archivos if str(ruta or "").strip()],
    }


def filas_tabla_preview(tabla: dict) -> list[dict]:
    """Pasa una tabla del preview a filas para st.dataframe."""
    columnas = [str(col) for col in (tabla.get("columnas") or [])]
    salida: list[dict] = []
    for fila in tabla.get("filas") or []:
        if isinstance(fila, dict):
            if columnas:
                salida.append({col: fila.get(col, "") for col in columnas})
            else:
                salida.append({str(clave): valor for clave, valor in fila.items()})
            continue
        if isinstance(fila, (list, tuple)):
            headers = columnas or [str(i + 1) for i in range(len(fila))]
            valores = list(fila) + [""] * max(0, len(headers) - len(fila))
            salida.append({col: valores[i] if i < len(valores) else "" for i, col in enumerate(headers)})
            continue
        salida.append({"valor": fila})
    return salida


def es_proyeccion(codigo: str) -> bool:
    return str(codigo or "").strip() == CODIGO_PROYECCION


def _norm_header(texto) -> str:
    plano = unicodedata.normalize("NFKD", str(texto or ""))
    plano = "".join(c for c in plano if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", plano).strip().casefold()


def _texto_celda(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%m-%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, int) and not isinstance(valor, bool):
        return str(valor)
    return str(valor).strip()


def cliente_activo(valor) -> bool:
    return _norm_header(valor) not in _INACTIVOS


def periodo_valido(periodo: str) -> bool:
    return bool(_RE_PERIODO.match(str(periodo or "").strip()))


def periodo_sugerido() -> str:
    hoy = datetime.now(_TZ_AR)
    return f"{hoy.month:02d}-{hoy.year}"


def leer_ficha_xlsx(contenido: bytes) -> list[dict]:
    """Lee la ficha. La primera fila son los títulos de COLUMNAS_FICHA."""
    if not contenido:
        raise ErrorRutina("El archivo de la ficha está vacío.")
    try:
        libro = openpyxl.load_workbook(BytesIO(contenido), data_only=True, read_only=True)
    except Exception as exc:
        raise ErrorRutina("No se pudo leer el Excel de la ficha.") from exc
    try:
        hoja = libro.active
        filas = hoja.iter_rows(values_only=True)
        encabezados = next(filas, None)
        if not encabezados:
            raise ErrorRutina("La ficha no tiene encabezados.")
        alias = {_norm_header(etiqueta): campo for campo, etiqueta in COLUMNAS_FICHA}
        columnas: list[str | None] = []
        for titulo in encabezados:
            columnas.append(alias.get(_norm_header(titulo)))
        if "sociedad" not in columnas:
            raise ErrorRutina("La ficha tiene que tener la columna Sociedad.")
        salida: list[dict] = []
        for cruda in filas:
            registro = {campo: "" for campo, _etiqueta in COLUMNAS_FICHA}
            for indice, campo in enumerate(columnas):
                if campo is None or indice >= len(cruda):
                    continue
                registro[campo] = _texto_celda(cruda[indice])
            if not registro["sociedad"]:
                continue
            salida.append(registro)
        return salida
    finally:
        libro.close()


def listar_ficha() -> list[dict]:
    asegurar_tablas()
    campos = ", ".join(campo for campo, _etiqueta in COLUMNAS_FICHA)
    with database.obtener_conexion() as conn:
        filas = conn.execute(
            f"SELECT {campos} FROM rutina_ficha_proyecciones ORDER BY orden, id"
        ).fetchall()
    return [{campo: fila[campo] or "" for campo, _etiqueta in COLUMNAS_FICHA} for fila in filas]


def reemplazar_ficha(filas: list[dict]) -> int:
    """Reemplaza la ficha entera. La planilla es la fuente."""
    asegurar_tablas()
    momento = ahora_utc()
    campos = [campo for campo, _etiqueta in COLUMNAS_FICHA]
    marcas = ", ".join("?" for _ in campos)
    with database.obtener_conexion() as conn:
        conn.execute("DELETE FROM rutina_ficha_proyecciones")
        for orden, fila in enumerate(filas):
            sociedad = str(fila.get("sociedad") or "").strip()
            if not sociedad:
                continue
            valores = [sociedad if campo == "sociedad" else str(fila.get(campo) or "").strip() for campo in campos]
            conn.execute(
                f"""
                INSERT INTO rutina_ficha_proyecciones ({", ".join(campos)}, orden, actualizado_en)
                VALUES ({marcas}, ?, ?)
                """,
                (*valores, orden, momento),
            )
        conn.commit()
    return len([fila for fila in filas if str(fila.get("sociedad") or "").strip()])


def importar_ficha_xlsx(contenido: bytes) -> int:
    return reemplazar_ficha(leer_ficha_xlsx(contenido))


def exportar_ficha() -> list[dict]:
    """Ficha con los títulos de la planilla, para `ficha --json`."""
    return [
        {etiqueta: fila.get(campo) or "" for campo, etiqueta in COLUMNAS_FICHA}
        for fila in listar_ficha()
    ]


def etiquetas_clientes(filas: list[dict] | None = None) -> list[tuple[str, dict]]:
    """Clientes activos con una etiqueta única para el multiselect."""
    activos = [fila for fila in (filas if filas is not None else listar_ficha()) if cliente_activo(fila.get("activa"))]
    repetidos: dict[str, int] = {}
    for fila in activos:
        repetidos[fila["sociedad"]] = repetidos.get(fila["sociedad"], 0) + 1
    salida: list[tuple[str, dict]] = []
    for fila in activos:
        nombre = fila["sociedad"]
        if repetidos[nombre] > 1 and fila.get("cuit"):
            etiqueta = f"{nombre} ({fila['cuit']})"
        else:
            etiqueta = nombre
        salida.append((etiqueta, fila))
    return salida


def resolver_clientes(elegidos: list[str], filas: list[dict] | None = None) -> tuple[list[dict], bool]:
    """Devuelve (clientes, todos). 'Todos' expande a los activos."""
    pares = etiquetas_clientes(filas)
    if not elegidos:
        return [], False
    if OPCION_TODOS in elegidos:
        return [fila for _etiqueta, fila in pares], True
    por_etiqueta = {etiqueta: fila for etiqueta, fila in pares}
    return [por_etiqueta[nombre] for nombre in elegidos if nombre in por_etiqueta], False


def parametros_proyeccion(periodo: str, clientes: list[dict], *, todos: bool) -> str:
    periodo_txt = str(periodo or "").strip()
    if not periodo_valido(periodo_txt):
        raise ErrorRutina("El período tiene que ser MM-AAAA.")
    if not clientes:
        raise ErrorRutina("Elegí al menos un cliente.")
    payload = {
        "periodo": periodo_txt,
        "todos": bool(todos),
        "clientes": [
            {"sociedad": str(fila.get("sociedad") or "").strip(), "cuit": str(fila.get("cuit") or "").strip()}
            for fila in clientes
            if str(fila.get("sociedad") or "").strip()
        ],
    }
    if not payload["clientes"]:
        raise ErrorRutina("Elegí al menos un cliente.")
    return json.dumps(payload, ensure_ascii=False)


def resumen_parametros(texto: str | None) -> str:
    """Texto corto para el historial. Si no es el JSON de proyecciones, se deja igual."""
    crudo = str(texto or "").strip()
    if not crudo:
        return ""
    try:
        data = json.loads(crudo)
    except json.JSONDecodeError:
        return crudo
    if not isinstance(data, dict) or "periodo" not in data:
        return crudo
    nombres = []
    for cliente in data.get("clientes") or []:
        if isinstance(cliente, dict):
            nombres.append(str(cliente.get("sociedad") or cliente.get("cuit") or "").strip())
        else:
            nombres.append(str(cliente).strip())
    nombres = [nombre for nombre in nombres if nombre]
    if data.get("todos"):
        quienes = "Todos"
    elif len(nombres) <= 3:
        quienes = ", ".join(nombres)
    else:
        quienes = ", ".join(nombres[:3]) + f" y {len(nombres) - 3} más"
    return f"{data.get('periodo')} · {quienes}".strip(" ·")


def tabla_control_proyeccion(filas: list[dict]) -> dict:
    """Tabla de preview: una fila por cliente y una columna por requisito, más el estado."""
    columnas = list(COLUMNAS_RESULTADO_PROYECCION)
    salida = []
    for fila in filas:
        salida.append([fila.get(columna, "") for columna in columnas])
    return {"titulo": "Por cliente", "columnas": columnas, "filas": salida}


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


def ultimo_pedido(rutina: str) -> dict | None:
    """El pedido más nuevo de esa rutina, en cualquier estado."""
    codigo = str(rutina or "").strip()
    if not codigo:
        return None
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            SELECT * FROM rutina_pedidos
            WHERE rutina = ?
            ORDER BY id DESC
            LIMIT 1
            """,
            (codigo,),
        ).fetchone()
    return _a_dict(fila) if fila else None


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
    preview: dict | None = None,
) -> dict:
    """Cierra un pedido EN_CURSO en OK o ERROR. `preview` es el JSON de la vista previa."""
    cierre = str(estado or "").strip().upper()
    if cierre not in ESTADOS_CIERRE:
        raise EstadoInvalido("El cierre tiene que ser OK o ERROR.")
    texto = str(resultado or "").strip()
    if not texto:
        raise ErrorRutina("El resultado no puede estar vacío.")
    adjuntos = str(archivos or "").strip() or None
    preview_txt = None
    if preview is not None:
        normal = normalizar_preview(preview)
        preview_txt = json.dumps(normal, ensure_ascii=False)
        if adjuntos is None and normal["archivos"]:
            adjuntos = "\n".join(normal["archivos"])
    asegurar_tablas()
    with database.obtener_conexion() as conn:
        fila = conn.execute(
            """
            UPDATE rutina_pedidos
            SET estado = ?, resultado = ?, archivos = ?, preview = ?, terminado_en = ?
            WHERE id = ? AND estado = 'EN_CURSO'
            RETURNING *
            """,
            (cierre, texto, adjuntos, preview_txt, ahora_utc(), int(pedido_id)),
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
