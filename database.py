"""Gestión de la base de datos SQLite para clientes del estudio contable."""

import json
import logging
import re
import sqlite3
import unicodedata
import warnings
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd
from ddgs import DDGS

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


def obtener_conexion() -> sqlite3.Connection:
    """Abre conexión a SQLite con filas accesibles por nombre de columna."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _migrar_tipo_monotributista(conn: sqlite3.Connection) -> None:
    """Amplía el CHECK de tipo_persona para admitir Monotributista."""
    fila = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='clientes'"
    ).fetchone()
    ddl = fila[0] if fila else ""
    if "Monotributista" in ddl:
        return
    conn.execute(
        """
        CREATE TABLE clientes_new (
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
    conn.execute(
        """
        INSERT INTO clientes_new (id, nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance, creado_en)
        SELECT id, nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance, creado_en
        FROM clientes
        """
    )
    conn.execute("DROP TABLE clientes")
    conn.execute("ALTER TABLE clientes_new RENAME TO clientes")


def inicializar_bd() -> None:
    """Crea la tabla de clientes si no existe."""
    with obtener_conexion() as conn:
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
        _migrar_tipo_monotributista(conn)
        # Migrate schema if needed
        try:
            conn.execute("ALTER TABLE clientes ADD COLUMN mes_cierre_balance INTEGER")
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
        # Migración no destructiva: agregar columna estado a tablas existentes
        try:
            conn.execute(
                "ALTER TABLE asientos_generados ADD COLUMN estado TEXT DEFAULT 'Ingresado'"
            )
        except Exception:
            pass  # columna ya existe

        import auth_oficina  # import local: evita ciclo de imports con database.py

        auth_oficina.inicializar_tabla_usuarios_oficina(conn)
        _inicializar_tablas_sueldos(conn)
        _inicializar_tablas_conciliacion(conn)
        conn.commit()
    auth_oficina.sembrar_usuarios_oficina_default()
    _sembrar_convenios_sueldos_default()
    _reset_cct_comercio_masivo_si_corresponde()
    sembrar_reglas_conciliacion_default()


def _reglas_cct_basicas() -> dict:
    from cct_escalas import reglas_comercio_julio_2026

    return reglas_comercio_julio_2026()


def _sembrar_convenios_sueldos_default() -> None:
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
    with obtener_conexion() as conn:
        for codigo, nombre, reglas in catalogo:
            payload = json.dumps(reglas, ensure_ascii=False)
            row = conn.execute(
                "SELECT id, reglas_json FROM convenios_colectivos WHERE codigo = ?",
                (codigo,),
            ).fetchone()
            if not row:
                conn.execute(
                    """
                    INSERT INTO convenios_colectivos (codigo, nombre, reglas_json)
                    VALUES (?, ?, ?)
                    """,
                    (codigo, nombre, payload),
                )
            elif codigo == "COMERCIO_130_75":
                # Mantener escala Comercio al día en cada init
                conn.execute(
                    "UPDATE convenios_colectivos SET nombre = ?, reglas_json = ? WHERE codigo = ?",
                    (nombre, payload, codigo),
                )
        conn.commit()


def actualizar_reglas_convenio(codigo: str, reglas: dict) -> None:
    with obtener_conexion() as conn:
        conn.execute(
            "UPDATE convenios_colectivos SET reglas_json = ? WHERE codigo = ?",
            (json.dumps(reglas, ensure_ascii=False), codigo),
        )
        conn.commit()


def _reset_cct_comercio_masivo_si_corresponde() -> None:
    """
    Una sola vez: si casi todos quedaron en COMERCIO por el default erróneo
    de la migración, se limpian para forzar asignación real por sociedad.
    """
    flag = BASE_DIR / "logs" / "sueldos_cct_reset_v1.flag"
    if flag.exists():
        return
    try:
        with obtener_conexion() as conn:
            total = conn.execute("SELECT COUNT(*) AS n FROM clientes").fetchone()["n"]
            comercio = conn.execute(
                "SELECT COUNT(*) AS n FROM clientes WHERE cct_asignado = 'COMERCIO_130_75'"
            ).fetchone()["n"]
            if total > 0 and comercio >= max(1, int(total * 0.8)):
                conn.execute(
                    "UPDATE clientes SET cct_asignado = NULL "
                    "WHERE cct_asignado = 'COMERCIO_130_75'"
                )
                conn.commit()
        flag.parent.mkdir(parents=True, exist_ok=True)
        flag.write_text("ok", encoding="utf-8")
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


def listar_clientes() -> list[dict]:
    """Devuelve todos los clientes ordenados por nombre."""
    with obtener_conexion() as conn:
        filas = conn.execute(
            "SELECT * FROM clientes ORDER BY nombre COLLATE NOCASE"
        ).fetchall()
    return [dict(fila) for fila in filas]


def obtener_cliente(cliente_id: int) -> Optional[dict]:
    """Obtiene un cliente por ID."""
    with obtener_conexion() as conn:
        fila = conn.execute(
            "SELECT * FROM clientes WHERE id = ?", (cliente_id,)
        ).fetchone()
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


def _resolver_plan_path_catalogo(item: dict, cuit: str) -> str | None:
    """Ruta de plan propio para catálogo. No usa plan_default (eso no es vínculo real)."""
    explicit = str(item.get("plan_cuentas") or item.get("plan_cuentas_path") or "").strip()
    if explicit:
        cand = Path(explicit)
        if not cand.is_absolute():
            cand = BASE_DIR / cand
        if cand.is_file() and cand.name.lower() not in ("plan_default.xlsx", "plan_default.xls"):
            if "cuentas contables (4)" not in cand.name.lower():
                return str(cand)
    propio = DATA_PLANES_DIR / f"plan_{cuit}.xlsx"
    if propio.is_file():
        return str(propio)
    return None


def sincronizar_clientes_catalogo(catalogo: list[dict]) -> dict[str, int]:
    """Inserta clientes del catálogo estático que aún no existen (por CUIT)."""
    inicializar_bd()
    stats = {"insertados": 0, "omitidos": 0, "errores": 0, "planes_vinculados": 0}
    with obtener_conexion() as conn:
        existentes = {
            re.sub(r"\D", "", str(row["cuit"])): dict(row)
            for row in conn.execute(
                "SELECT id, cuit, plan_cuentas_path FROM clientes"
            ).fetchall()
        }
        for item in catalogo:
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
                stats["errores"] += 1
                continue
            mes_raw = item.get("mes_cierre_balance")
            try:
                mes_cierre = int(mes_raw) if mes_raw not in (None, "") else 12
            except (TypeError, ValueError):
                mes_cierre = 12
            if mes_cierre < 1 or mes_cierre > 12:
                mes_cierre = 12
            if tipo_persona in ("Persona Física", "Monotributista"):
                mes_cierre = 12
            plan_path = _resolver_plan_path_catalogo(item, cuit)
            if cuit in existentes:
                # En Cloud/repo: si hay plan propio y la BD apunta a un path inexistente, actualizar.
                row = existentes[cuit]
                actual = Path(str(row.get("plan_cuentas_path") or ""))
                propio = DATA_PLANES_DIR / f"plan_{cuit}.xlsx"
                if propio.is_file() and (not actual.is_file()):
                    conn.execute(
                        "UPDATE clientes SET plan_cuentas_path = ? WHERE id = ?",
                        (str(propio), row["id"]),
                    )
                    stats["planes_vinculados"] += 1
                stats["omitidos"] += 1
                continue
            try:
                conn.execute(
                    """
                    INSERT INTO clientes (nombre, cuit, tipo_persona, plan_cuentas_path, mes_cierre_balance)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (nombre, cuit, tipo_persona, plan_path, mes_cierre),
                )
                existentes[cuit] = {"cuit": cuit, "plan_cuentas_path": plan_path}
                stats["insertados"] += 1
            except sqlite3.IntegrityError:
                stats["omitidos"] += 1
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
    return sincronizar_clientes_catalogo(data)


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
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS proveedores_pendientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER NOT NULL,
            fecha TEXT,
            tipo_comp TEXT NOT NULL DEFAULT '',
            num_comp TEXT NOT NULL DEFAULT '',
            razon_social TEXT NOT NULL DEFAULT '',
            importe TEXT NOT NULL DEFAULT '0.00',
            usado INTEGER NOT NULL DEFAULT 0,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
        """
    )
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


def sembrar_reglas_conciliacion_default() -> None:
    """Inserta el diccionario del prompt solo si la tabla esta vacia."""
    from motor_conciliacion import REGLAS_SEED

    with obtener_conexion() as conn:
        n = conn.execute("SELECT COUNT(*) AS n FROM clasificacion_reglas").fetchone()["n"]
        if n and int(n) > 0:
            return
        for i, (patron, categoria, tipo) in enumerate(REGLAS_SEED):
            conn.execute(
                """
                INSERT INTO clasificacion_reglas (patron, categoria, tipo, orden, activo)
                VALUES (?, ?, ?, ?, 1)
                """,
                (patron, categoria, tipo, i),
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
                    match_detalle, match_ref_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    allowed = {"categoria", "tipo", "estado", "match_detalle", "match_ref_id"}
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
                    cliente_id, fecha, tipo_comp, num_comp, razon_social, importe, usado
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    cliente_id,
                    fecha,
                    str(f.get("tipo_comp") or f.get("tipo") or ""),
                    str(f.get("num_comp") or f.get("comprobante") or ""),
                    str(f.get("razon_social") or f.get("proveedor") or ""),
                    str(f.get("importe") or "0.00"),
                    1 if f.get("usado") else 0,
                ),
            )
        conn.commit()
        return len(filas)


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
