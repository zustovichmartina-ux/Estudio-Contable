"""Lectura de facturas/proveedores desde SQL Server de Tango Estudios Contables."""
from __future__ import annotations

import os
import re
from datetime import date, datetime
from functools import lru_cache
from typing import Any, Optional

import pandas as pd
from rapidfuzz import fuzz, process

try:
    import pyodbc
except ImportError:  # pragma: no cover
    pyodbc = None  # type: ignore[assignment]

# Defaults de la oficina (override con env TANGO_SQL_*)
TANGO_SQL_SERVER = os.environ.get("TANGO_SQL_SERVER", r"192.168.1.55,63288")
TANGO_SQL_USER = os.environ.get("TANGO_SQL_USER", "Axoft")
TANGO_SQL_PASSWORD = os.environ.get("TANGO_SQL_PASSWORD", "Axoft")
TANGO_SQL_DICCIONARIO = os.environ.get("TANGO_SQL_DICCIONARIO", "Diccionario_039846_001")
TANGO_SQL_DRIVER = os.environ.get("TANGO_SQL_DRIVER", "SQL Server")
CUENTA_PROVEEDORES_DEFAULT = "21101"


class TangoSQLError(RuntimeError):
    """Error de conexión o consulta a Tango SQL."""


def _normalizar_clave(texto: str) -> str:
    t = (texto or "").upper()
    t = (
        t.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
    )
    t = re.sub(r"[^A-Z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _cuit_digitos(cuit: str | None) -> str:
    return re.sub(r"\D", "", str(cuit or ""))


def tango_sql_disponible() -> tuple[bool, str]:
    """True si pyodbc está instalado y el servidor responde."""
    if pyodbc is None:
        return False, "Falta el paquete pyodbc. Instalalo con: pip install pyodbc"
    try:
        with conectar_tango_sql(TANGO_SQL_DICCIONARIO) as cn:
            cn.cursor().execute("SELECT 1")
        return True, f"OK · {TANGO_SQL_SERVER}"
    except Exception as exc:
        return False, str(exc).split("\n")[0][:200]


def _connection_string(database: str) -> str:
    return (
        f"DRIVER={{{TANGO_SQL_DRIVER}}};"
        f"SERVER={TANGO_SQL_SERVER};"
        f"DATABASE={database};"
        f"UID={TANGO_SQL_USER};"
        f"PWD={TANGO_SQL_PASSWORD};"
        "TrustServerCertificate=yes;"
        "Connection Timeout=12;"
    )


def conectar_tango_sql(database: str):
    """Abre conexión ODBC a una base Tango."""
    if pyodbc is None:
        raise TangoSQLError("Falta pyodbc. Ejecutá: pip install pyodbc")
    try:
        return pyodbc.connect(_connection_string(database))
    except Exception as exc:
        raise TangoSQLError(f"No se pudo conectar a Tango SQL ({database}): {exc}") from exc


def _fetch_df(cn, sql: str, params: tuple | list | None = None) -> pd.DataFrame:
    cur = cn.cursor()
    if params:
        cur.execute(sql, params)
    else:
        cur.execute(sql)
    if cur.description is None:
        return pd.DataFrame()
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    if not rows:
        return pd.DataFrame(columns=cols)
    data = [tuple(r) for r in rows]
    return pd.DataFrame.from_records(data, columns=cols)


@lru_cache(maxsize=1)
def listar_empresas_tango(solo_habilitadas: bool = True) -> list[dict[str, Any]]:
    """Empresas del diccionario Tango (NombreEmpresa → NombreBD)."""
    with conectar_tango_sql(TANGO_SQL_DICCIONARIO) as cn:
        sql = """
            SELECT IDEmpresa, NombreEmpresa, NombreBD, Habilita, ESTADO
            FROM Empresa
            WHERE NombreBD IS NOT NULL AND LTRIM(RTRIM(NombreBD)) <> ''
        """
        if solo_habilitadas:
            sql += " AND (Habilita = 'S' OR Habilita = 'C')"
        sql += " ORDER BY NombreEmpresa"
        df = _fetch_df(cn, sql)
    out: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        out.append({
            "id": int(row["IDEmpresa"]),
            "nombre": str(row["NombreEmpresa"] or "").strip(),
            "base": str(row["NombreBD"] or "").strip(),
            "habilita": str(row.get("Habilita") or ""),
            "estado": str(row.get("ESTADO") or ""),
        })
    return out


def leer_cuit_empresa_tango(nombre_bd: str) -> str:
    """CUIT de la ficha EMPRESA dentro de la base de la sociedad."""
    with conectar_tango_sql(nombre_bd) as cn:
        df = _fetch_df(cn, "SELECT TOP 1 CUIT FROM EMPRESA")
    if df.empty:
        return ""
    return str(df.iloc[0]["CUIT"] or "").strip()


def resolver_empresa_tango(
    *,
    cuit: str | None = None,
    nombre: str | None = None,
    nombre_bd: str | None = None,
) -> dict[str, Any]:
    """
    Resuelve la base SQL de la empresa.
    Prioridad: nombre_bd explícito → CUIT → fuzzy por nombre.
    """
    if nombre_bd:
        return {
            "nombre": nombre_bd,
            "base": nombre_bd,
            "cuit": _cuit_digitos(cuit),
            "score": 100.0,
            "origen": "base_explicita",
        }

    empresas = listar_empresas_tango(solo_habilitadas=True)
    if not empresas:
        raise TangoSQLError("No hay empresas en el diccionario Tango.")

    cuit_n = _cuit_digitos(cuit)
    nombre_n = _normalizar_clave(nombre or "")

    # 1) Fuzzy por nombre (rápido)
    if nombre_n:
        opciones = {_normalizar_clave(e["nombre"]): e for e in empresas}
        match = process.extractOne(
            nombre_n,
            list(opciones.keys()),
            scorer=fuzz.token_set_ratio,
        )
        if match and match[1] >= 70:
            emp = opciones[match[0]]
            cuit_emp = ""
            try:
                cuit_emp = _cuit_digitos(leer_cuit_empresa_tango(emp["base"]))
            except Exception:
                cuit_emp = ""
            if cuit_n and cuit_emp and cuit_n != cuit_emp and match[1] < 92:
                pass  # seguir a búsqueda por CUIT
            else:
                return {
                    "nombre": emp["nombre"],
                    "base": emp["base"],
                    "cuit": cuit_emp or cuit_n,
                    "score": float(match[1]),
                    "origen": "nombre",
                }

    # 2) Por CUIT (recorre empresas; cacheable)
    if cuit_n:
        for emp in empresas:
            try:
                cuit_emp = _cuit_digitos(leer_cuit_empresa_tango(emp["base"]))
            except Exception:
                continue
            if cuit_emp == cuit_n:
                return {
                    "nombre": emp["nombre"],
                    "base": emp["base"],
                    "cuit": cuit_emp,
                    "score": 100.0,
                    "origen": "cuit",
                }

    raise TangoSQLError(
        f"No se encontró empresa Tango para nombre={nombre!r} cuit={cuit!r}. "
        "Elegí la base manualmente."
    )


def cargar_facturas_proveedores_tango_sql(
    *,
    nombre_bd: str | None = None,
    cuit: str | None = None,
    nombre_empresa: str | None = None,
    cuenta: str = CUENTA_PROVEEDORES_DEFAULT,
    fecha_desde: date | datetime | None = None,
    fecha_hasta: date | datetime | None = None,
) -> pd.DataFrame:
    """
    Listado por imputación contable (cuenta proveedores, Haber > 0)
    desde V_SubdiarioAsientosIV — equivalente al Excel de Tango.
    """
    emp = resolver_empresa_tango(cuit=cuit, nombre=nombre_empresa, nombre_bd=nombre_bd)
    base = emp["base"]
    cuenta = str(cuenta or CUENTA_PROVEEDORES_DEFAULT).strip()

    sql = """
        SELECT
            Codigo_de_Cuenta AS cuenta,
            Fecha_Asiento AS fecha,
            Codigo_de_Tipo_Comprobante AS tipo,
            Numero_de_Comprobante AS comprobante,
            CodProvClie AS codigo_prov,
            RazonSocial AS proveedor,
            Concepto_Movimiento AS descripcion,
            CAST(0 AS float) AS debe,
            CAST(Importe_Renglon AS float) AS haber,
            Descripcion_de_Cuenta AS desc_cuenta
        FROM V_SubdiarioAsientosIV
        WHERE Codigo_de_Cuenta LIKE ?
          AND D_H = 'H'
          AND Importe_Renglon > 0.009
    """
    params: list[Any] = [f"{cuenta}%"]
    if fecha_desde is not None:
        sql += " AND Fecha_Asiento >= ?"
        params.append(fecha_desde if isinstance(fecha_desde, datetime) else datetime.combine(fecha_desde, datetime.min.time()))
    if fecha_hasta is not None:
        sql += " AND Fecha_Asiento < DATEADD(day, 1, ?)"
        params.append(fecha_hasta if isinstance(fecha_hasta, datetime) else datetime.combine(fecha_hasta, datetime.min.time()))
    sql += " ORDER BY Fecha_Asiento, RazonSocial, Numero_de_Comprobante"

    with conectar_tango_sql(base) as cn:
        raw = _fetch_df(cn, sql, params)

    if raw.empty:
        return pd.DataFrame(
            columns=[
                "factura_id", "fecha", "proveedor", "proveedor_norm", "importe",
                "comprobante", "tipo", "cuenta", "descripcion", "codigo_prov", "cuit",
                "debe", "haber",
            ]
        )

    # Reusar normalización del cargador Excel
    from procesador import _normalizar_texto

    work = raw.copy()
    work["fecha"] = pd.to_datetime(work["fecha"], errors="coerce").dt.date
    work["haber"] = pd.to_numeric(work["haber"], errors="coerce").fillna(0.0).round(2)
    work["debe"] = 0.0
    work["importe"] = work["haber"]
    work["proveedor"] = work["proveedor"].astype(str).fillna("").str.strip()
    work["proveedor_norm"] = work["proveedor"].map(_normalizar_texto)
    work["comprobante"] = work["comprobante"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    work["tipo"] = work["tipo"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    work["cuenta"] = work["cuenta"].astype(str)
    work["descripcion"] = work["descripcion"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    work["codigo_prov"] = work["codigo_prov"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    work["cuit"] = ""
    work = work[work["proveedor_norm"].str.len() > 1].copy()
    work = work.reset_index(drop=True)
    work["factura_id"] = work.index.astype(int)
    work.attrs["tango_empresa"] = emp
    work.attrs["tango_base"] = base
    work.attrs["origen"] = f"Tango SQL · {base} · cuenta {cuenta}"
    return work[
        [
            "factura_id", "fecha", "proveedor", "proveedor_norm", "importe",
            "comprobante", "tipo", "cuenta", "descripcion", "codigo_prov", "cuit",
            "debe", "haber",
        ]
    ]


def probar_carga_proveedores_rele() -> dict[str, Any]:
    """Smoke test rápido para RELE."""
    df = cargar_facturas_proveedores_tango_sql(
        nombre_empresa="OFTALMOLOGIA RELE MAR DEL PLATA S.R.L.",
        cuit="30-71802274-2",
    )
    return {
        "filas": len(df),
        "importe": float(df["importe"].sum()) if len(df) else 0.0,
        "origen": df.attrs.get("origen"),
    }
