"""Ventas Mercado Pago / Mercado Libre de una sociedad.

Separa venta, MELI Envíos y comisiones. No mezclar con gastos_mp (cuenta personal).
"""

from __future__ import annotations

import csv
import io
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from excel_formato_estudio import guardar_informe_excel, exportar_informe_excel

DIR_LOCAL = Path.home() / ".estudio-contable" / "ventas_mp"
ARCHIVO_SOCIEDADES = DIR_LOCAL / "sociedades.json"

COLUMNAS_REPORTE = [
    "TRANSACTION_DATE",
    "SETTLEMENT_DATE",
    "SOURCE_ID",
    "ORDER_ID",
    "EXTERNAL_REFERENCE",
    "TRANSACTION_TYPE",
    "DESCRIPTION",
    "SALE_DETAIL",
    "STORE_NAME",
    "PAYER_NAME",
    "PAYER_ID_NUMBER",
    "PAYMENT_METHOD",
    "PAYMENT_METHOD_TYPE",
    "TRANSACTION_AMOUNT",
    "SETTLEMENT_NET_AMOUNT",
    "REAL_AMOUNT",
    "FEE_AMOUNT",
    "TAXES_AMOUNT",
    "SHIPPING_ID",
    "BUSINESS_UNIT",
    "SUB_UNIT",
    "METADATA",
]

COLS = [
    "Id",
    "Fecha",
    "Hora",
    "Orden",
    "Concepto",
    "Contraparte",
    "Detalle",
    "Bruto",
    "Comision",
    "Neto",
    "Tipo",
    "Medio",
]

TIPOS_ENVIO = {
    "SETTLEMENT_SHIPPING",
    "REFUND_SHIPPING",
    "CHARGEBACK_SHIPPING",
    "DISPUTE_SHIPPING",
}
TIPOS_DEVOLUCION = {"REFUND", "CHARGEBACK", "DISPUTE", "CASHBACK"}
TIPOS_RETIRO = {"WITHDRAWAL", "PAYOUT", "WITHDRAWAL_CANCEL"}


def asegurar_dir() -> Path:
    DIR_LOCAL.mkdir(parents=True, exist_ok=True)
    return DIR_LOCAL


def _slug(texto: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", (texto or "").lower()).strip("-")
    return s[:60] or "sociedad"


def dir_sociedad(slug: str) -> Path:
    path = DIR_LOCAL / _slug(slug)
    path.mkdir(parents=True, exist_ok=True)
    return path


def listar_sociedades() -> list[dict[str, str]]:
    if not ARCHIVO_SOCIEDADES.exists():
        return []
    try:
        data = json.loads(ARCHIVO_SOCIEDADES.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    return [x for x in data if isinstance(x, dict) and x.get("slug")]


def guardar_sociedad(nombre: str, cuit: str = "") -> dict[str, str]:
    asegurar_dir()
    slug = _slug(nombre or cuit or "sociedad")
    actual = listar_sociedades()
    fila = {"slug": slug, "nombre": (nombre or slug).strip(), "cuit": re.sub(r"\D", "", cuit or "")}
    resto = [s for s in actual if s.get("slug") != slug]
    ARCHIVO_SOCIEDADES.write_text(
        json.dumps([fila, *resto], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    dir_sociedad(slug)
    return fila


def archivo_token(slug: str) -> Path:
    return dir_sociedad(slug) / "token.txt"


def cargar_token(slug: str) -> str:
    path = archivo_token(slug)
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    return ""


def guardar_token(slug: str, token: str) -> None:
    archivo_token(slug).write_text(token.strip(), encoding="utf-8")


def _limpio(valor: object) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = " ".join(str(valor).strip().split())
    if s.lower() in {"", "nan", "none", "-", "null"}:
        return ""
    return s


def _a_float(valor: object) -> float:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip().replace(" ", "").replace("$", "")
    if not s or s.lower() in {"-", "nan"}:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _a_fecha_hora(valor: object) -> tuple[date | None, str]:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None, ""
    if isinstance(valor, datetime):
        return valor.date(), valor.strftime("%H:%M")
    if isinstance(valor, date):
        return valor, ""
    texto = str(valor).strip()
    dayfirst = "T" not in texto and texto.count("/") >= 2
    ts = pd.to_datetime(texto, dayfirst=dayfirst, errors="coerce")
    if pd.isna(ts):
        return None, ""
    hora = ts.strftime("%H:%M")
    if hora == "00:00" and "T" not in texto and ":" not in texto:
        return ts.date(), ""
    return ts.date(), hora


def _buscar_col(df: pd.DataFrame, candidatos: tuple[str, ...]) -> str | None:
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidatos:
        if cand.lower() in cols:
            return cols[cand.lower()]
    for cand in candidatos:
        for low, orig in cols.items():
            if cand.lower() in low:
                return orig
    return None


def df_vacio() -> pd.DataFrame:
    return pd.DataFrame(columns=COLS)


def concepto_de_tipo(tipo: str, descripcion: str = "") -> str:
    t = str(tipo or "").strip().upper()
    low = f"{tipo} {descripcion}".lower()
    if t in TIPOS_ENVIO or "shipping" in low or "meli envio" in low or "mercado envíos" in low or "mercado envios" in low:
        return "MELI Envíos"
    if t in TIPOS_RETIRO or "retiro" in low or "payout" in low:
        return "Retiro"
    if t in TIPOS_DEVOLUCION or "reintegro" in low or "devolucion" in low:
        return "Devolución"
    if "fee" in low or "comision" in low or "comisión" in low:
        return "Comisión MP/ML"
    if t in {"SETTLEMENT", "PAYMENT", "REGULAR_PAYMENT"}:
        return "Venta"
    return "Otro"


def _mp_request(method: str, url: str, token: str, body: dict | None = None) -> tuple[int, Any]:
    data = None
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/csv, */*",
        "User-Agent": "EstudioContable-VentasMP/1.0",
    }
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            raw = resp.read()
            if not raw:
                return resp.status, {}
            ctype = (resp.headers.get("Content-Type") or "").lower()
            if "csv" in ctype or url.rstrip("/").endswith(".csv"):
                return resp.status, raw
            try:
                return resp.status, json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                return resp.status, raw
    except urllib.error.HTTPError as exc:
        err = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Mercado Pago HTTP {exc.code}: {err[:500]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"No se pudo conectar con Mercado Pago: {exc.reason}") from exc


def _mp_try(method: str, url: str, token: str, body: dict | None = None) -> Any | None:
    try:
        _status, payload = _mp_request(method, url, token, body)
        return payload
    except RuntimeError:
        return None


def usuario_mp(token: str) -> dict[str, Any]:
    data = _mp_try("GET", "https://api.mercadopago.com/users/me", token)
    return data if isinstance(data, dict) else {}


def conectar(token: str) -> dict[str, Any]:
    """Valida el token de producción de la sociedad y devuelve quién es."""
    me = usuario_mp(token)
    if not me or not me.get("id"):
        raise RuntimeError(
            "El token no sirve. Tiene que ser el Access Token de PRODUCCIÓN "
            "(APP_USR-…) de la cuenta Mercado Pago de la sociedad, no el personal."
        )
    nombre = " ".join(
        p for p in (_limpio(me.get("first_name")), _limpio(me.get("last_name"))) if p
    ) or _limpio(me.get("nickname")) or str(me.get("id"))
    return {
        "id": me.get("id"),
        "nombre": nombre,
        "nickname": _limpio(me.get("nickname")),
        "email": _limpio(me.get("email")),
        "cuit": _limpio((me.get("identification") or {}).get("number"))
        if isinstance(me.get("identification"), dict)
        else "",
        "site": _limpio(me.get("site_id")),
    }


def _configurar_reporte(token: str) -> None:
    actual = _mp_try("GET", "https://api.mercadopago.com/v1/account/settlement_report/config", token)
    actual = actual if isinstance(actual, dict) else {}
    freq = actual.get("frequency") or {"hour": 0, "type": "daily", "value": None, "format": "CSV"}
    body = {
        "file_name_prefix": actual.get("file_name_prefix") or "ventas-sociedad",
        "include_withdraw": True,
        "header_language": "es",
        "display_timezone": actual.get("display_timezone") or "GMT-03",
        "separator": ";",
        "frequency": freq,
        "columns": [{"key": k, "alias": ""} for k in COLUMNAS_REPORTE],
    }
    _mp_try("PUT", "https://api.mercadopago.com/v1/account/settlement_report/config", token, body)


def _col(df: pd.DataFrame, row: pd.Series, candidatos: tuple[str, ...]) -> str:
    col = _buscar_col(df, candidatos)
    if not col:
        return ""
    return _limpio(row.get(col))


def parsear_reporte(contenido: bytes | str) -> pd.DataFrame:
    if isinstance(contenido, bytes):
        texto = contenido.decode("utf-8-sig", errors="replace")
    else:
        texto = contenido
    sample = texto[:4000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        sep = dialect.delimiter
    except csv.Error:
        sep = ";" if sample.count(";") > sample.count(",") else ","
    df_raw = pd.read_csv(io.StringIO(texto), sep=sep)
    col_fecha = _buscar_col(df_raw, ("transaction_date", "settlement_date", "fecha", "date"))
    col_id = _buscar_col(df_raw, ("source_id", "id", "order_id"))
    col_bruto = _buscar_col(df_raw, ("transaction_amount", "importe", "monto"))
    col_neto = _buscar_col(df_raw, ("settlement_net_amount", "real_amount", "net_received_amount"))
    col_fee = _buscar_col(df_raw, ("fee_amount", "comision", "comisión", "mktp_fee_amount"))
    filas: list[dict[str, Any]] = []
    for i, row in df_raw.iterrows():
        fecha, hora = _a_fecha_hora(row[col_fecha]) if col_fecha else (None, "")
        tipo = _col(df_raw, row, ("transaction_type", "tipo"))
        desc = _col(df_raw, row, ("description", "descripcion", "descripción", "sale_detail"))
        contrap = (
            _col(df_raw, row, ("payer_name", "store_name", "origen/destino", "contraparte"))
            or desc
            or "(sin dato)"
        )
        bruto = _a_float(row[col_bruto]) if col_bruto else 0.0
        neto = _a_float(row[col_neto]) if col_neto else bruto
        fee = abs(_a_float(row[col_fee]) if col_fee else 0.0)
        concepto = concepto_de_tipo(tipo, desc)
        if concepto == "MELI Envíos" and neto == 0 and bruto != 0:
            neto = bruto
        mid = _col(df_raw, row, ("source_id", "id")) or f"fila-{i}"
        orden = _col(df_raw, row, ("order_id", "external_reference", "shipping_id")) or mid
        filas.append(
            {
                "Id": mid,
                "Fecha": fecha or date.today(),
                "Hora": hora,
                "Orden": orden,
                "Concepto": concepto,
                "Contraparte": contrap,
                "Detalle": " · ".join(p for p in (desc, tipo, _col(df_raw, row, ("sale_detail",))) if p and p != contrap),
                "Bruto": round(bruto, 2),
                "Comision": round(fee, 2),
                "Neto": round(neto, 2),
                "Tipo": tipo or "MOVIMIENTO",
                "Medio": _col(df_raw, row, ("payment_method", "payment_method_type")),
            }
        )
    if not filas:
        return df_vacio()
    return pd.DataFrame(filas)


def _cobros_a_filas(pagos: list[dict[str, Any]], my_id: object) -> list[dict[str, Any]]:
    yo = str(my_id or "")
    filas: list[dict[str, Any]] = []
    for p in pagos:
        if not isinstance(p, dict):
            continue
        if str(p.get("status") or "") not in {"approved", "accredited"}:
            continue
        collector = str(p.get("collector_id") or "")
        if yo and collector and collector != yo:
            continue
        payer = p.get("payer") if isinstance(p.get("payer"), dict) else {}
        nombre = " ".join(
            x for x in (_limpio(payer.get("first_name")), _limpio(payer.get("last_name"))) if x
        )
        email = _limpio(payer.get("email"))
        desc = _limpio(p.get("description"))
        op = str(p.get("operation_type") or "payment")
        if op in {"account_fund", "partition_transfer", "withdraw"}:
            continue
        td = p.get("transaction_details") if isinstance(p.get("transaction_details"), dict) else {}
        try:
            bruto = float(p.get("transaction_amount") or 0)
        except (TypeError, ValueError):
            bruto = 0.0
        try:
            neto = float(td.get("net_received_amount") if td.get("net_received_amount") is not None else bruto)
        except (TypeError, ValueError):
            neto = bruto
        fee = round(abs(bruto) - abs(neto), 2) if abs(bruto) >= abs(neto) else 0.0
        fecha, hora = _a_fecha_hora(p.get("date_created") or p.get("date_approved"))
        orden = _limpio(p.get("order") if not isinstance(p.get("order"), dict) else p.get("order", {}).get("id"))
        if not orden and isinstance(p.get("order"), dict):
            orden = _limpio(p["order"].get("id"))
        filas.append(
            {
                "Id": str(p.get("id") or ""),
                "Fecha": fecha or date.today(),
                "Hora": hora,
                "Orden": orden or str(p.get("id") or ""),
                "Concepto": "Venta",
                "Contraparte": nombre or email or desc or "Comprador",
                "Detalle": " · ".join(x for x in (desc, email, op) if x),
                "Bruto": round(abs(bruto), 2),
                "Comision": fee,
                "Neto": round(abs(neto), 2),
                "Tipo": op,
                "Medio": _limpio(p.get("payment_type_id") or p.get("payment_method_id")),
            }
        )
    return filas


def buscar_cobros(token: str, desde: date, hasta: date, my_id: object) -> pd.DataFrame:
    begin = datetime(desde.year, desde.month, desde.day).strftime("%Y-%m-%dT00:00:00.000-03:00")
    end = datetime(hasta.year, hasta.month, hasta.day, 23, 59, 59).strftime("%Y-%m-%dT23:59:59.000-03:00")
    filas: list[dict[str, Any]] = []
    offset = 0
    while offset < 4000:
        params = urllib.parse.urlencode(
            {
                "sort": "date_created",
                "criteria": "desc",
                "range": "date_created",
                "begin_date": begin,
                "end_date": end,
                "offset": str(offset),
                "limit": "50",
            }
        )
        payload = _mp_try("GET", f"https://api.mercadopago.com/v1/payments/search?{params}", token)
        if not isinstance(payload, dict):
            break
        results = payload.get("results") or []
        if not results:
            break
        filas.extend(_cobros_a_filas(results, my_id))
        paging = payload.get("paging") if isinstance(payload.get("paging"), dict) else {}
        total = int(paging.get("total") or 0)
        offset += len(results)
        if offset >= total:
            break
    if not filas:
        return df_vacio()
    df = pd.DataFrame(filas)
    return df[df["Id"].astype(str) != ""].drop_duplicates(subset=["Id"])


def bajar_reporte(token: str, desde: date, hasta: date) -> pd.DataFrame:
    _configurar_reporte(token)
    begin = datetime(desde.year, desde.month, desde.day).strftime("%Y-%m-%dT00:00:00Z")
    end_dt = datetime(hasta.year, hasta.month, hasta.day) + timedelta(days=1) - timedelta(seconds=1)
    end = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    _mp_request(
        "POST",
        "https://api.mercadopago.com/v1/account/settlement_report",
        token,
        {"begin_date": begin, "end_date": end},
    )
    file_name = None
    for _ in range(25):
        time.sleep(2)
        _status, lista = _mp_request("GET", "https://api.mercadopago.com/v1/account/settlement_report/list", token)
        items = lista if isinstance(lista, list) else (lista.get("results") or [])
        for item in items:
            if not isinstance(item, dict):
                continue
            st = str(item.get("status") or "").lower()
            fn = item.get("file_name") or item.get("filename")
            if st in {"processed", "ready", "available"} and fn:
                file_name = fn
                break
        if file_name:
            break
    if not file_name:
        raise RuntimeError(
            "Mercado Pago aceptó el pedido pero el reporte no terminó a tiempo. "
            "Probá de nuevo o subí el CSV de dinero en cuenta."
        )
    _status, raw = _mp_request(
        "GET",
        f"https://api.mercadopago.com/v1/account/settlement_report/{urllib.parse.quote(str(file_name))}",
        token,
    )
    if isinstance(raw, dict):
        raise RuntimeError(f"Respuesta inesperada al descargar el reporte: {raw}")
    if isinstance(raw, bytes):
        return parsear_reporte(raw)
    return parsear_reporte(str(raw))


def _merge(a: pd.DataFrame, b: pd.DataFrame) -> pd.DataFrame:
    partes = [x for x in (a, b) if x is not None and not x.empty]
    if not partes:
        return df_vacio()
    df = pd.concat(partes, ignore_index=True)
    df["_q"] = df["Contraparte"].astype(str).ne("(sin dato)").astype(int) * 50 + df["Detalle"].astype(str).str.len().clip(0, 40)
    df.loc[df["Concepto"].eq("MELI Envíos"), "_q"] += 30
    df = df.sort_values("_q", ascending=False).drop_duplicates(subset=["Id"], keep="first")
    return df.drop(columns=["_q"]).reset_index(drop=True)


def sincronizar_ventas(
    slug: str,
    token: str,
    desde: date,
    hasta: date,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    cuenta = conectar(token)
    avisos: list[str] = []
    reporte = df_vacio()
    try:
        reporte = bajar_reporte(token, desde, hasta)
    except RuntimeError as exc:
        avisos.append(str(exc))
    cobros = buscar_cobros(token, desde, hasta, cuenta.get("id"))
    df = _merge(reporte, cobros)
    if df.empty:
        raise RuntimeError(
            "La API no devolvió ventas en ese período. "
            "Confirmá que el token es de la cuenta vendedora de la sociedad. "
            + (" ".join(avisos) if avisos else "")
        )
    meta = {
        **cuenta,
        "avisos": avisos,
        "movimientos": int(len(df)),
        "filas_reporte": int(len(reporte)),
        "filas_cobros": int(len(cobros)),
        "desde": str(desde),
        "hasta": str(hasta),
    }
    guardar_ledger(slug, df, meta)
    return df, meta


def guardar_ledger(slug: str, df: pd.DataFrame, meta: dict[str, Any] | None = None) -> None:
    carpeta = dir_sociedad(slug)
    cols = [c for c in COLS if c in df.columns]
    df[cols].to_csv(carpeta / "ledger.csv", index=False, encoding="utf-8-sig")
    if meta is not None:
        (carpeta / "meta.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )


def cargar_ledger(slug: str) -> tuple[pd.DataFrame, dict[str, Any]]:
    path = dir_sociedad(slug) / "ledger.csv"
    if not path.exists():
        return df_vacio(), {}
    df = pd.read_csv(path)
    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
    if "Id" in df.columns:
        df["Id"] = df["Id"].astype(str)
    meta: dict[str, Any] = {}
    meta_path = dir_sociedad(slug) / "meta.json"
    if meta_path.exists():
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            meta = {}
    return df, meta


def kpis(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {"ventas": 0.0, "envios": 0.0, "comisiones": 0.0, "neto": 0.0, "devoluciones": 0.0, "movimientos": 0}
    conc = df["Concepto"].astype(str)
    bruto = pd.to_numeric(df["Bruto"], errors="coerce").fillna(0.0)
    comi = pd.to_numeric(df["Comision"], errors="coerce").fillna(0.0)
    neto = pd.to_numeric(df["Neto"], errors="coerce").fillna(0.0)
    ventas = bruto[conc.eq("Venta")].sum()
    envios = neto[conc.eq("MELI Envíos")].abs().sum()
    if envios == 0:
        envios = bruto[conc.eq("MELI Envíos")].abs().sum()
    return {
        "ventas": float(ventas),
        "envios": float(envios),
        "comisiones": float(comi.sum()),
        "neto": float(neto.sum()),
        "devoluciones": float(neto[conc.eq("Devolución")].abs().sum()),
        "movimientos": int(len(df)),
    }


def cuadro_por_orden(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Orden", "Ventas", "MELI_Envios", "Comision", "Neto", "Movimientos"])
    g = df.copy()
    g["Bruto"] = pd.to_numeric(g["Bruto"], errors="coerce").fillna(0.0)
    g["Comision"] = pd.to_numeric(g["Comision"], errors="coerce").fillna(0.0)
    g["Neto"] = pd.to_numeric(g["Neto"], errors="coerce").fillna(0.0)
    filas = []
    for orden, part in g.groupby("Orden", dropna=False):
        conc = part["Concepto"].astype(str)
        filas.append(
            {
                "Orden": str(orden),
                "Ventas": float(part.loc[conc.eq("Venta"), "Bruto"].sum()),
                "MELI_Envios": float(part.loc[conc.eq("MELI Envíos"), "Neto"].abs().sum()),
                "Comision": float(part["Comision"].sum()),
                "Neto": float(part["Neto"].sum()),
                "Movimientos": int(len(part)),
            }
        )
    out = pd.DataFrame(filas)
    return out.sort_values("Ventas", ascending=False).reset_index(drop=True)


def exportar_excel_ventas(
    df: pd.DataFrame,
    ruta: str | Path,
    *,
    sociedad: str,
    periodo: str,
) -> Path:
    metricas = kpis(df)
    por_orden = cuadro_por_orden(df)
    por_concepto = (
        df.groupby("Concepto", dropna=False)
        .agg(Movimientos=("Id", "count"), Bruto=("Bruto", "sum"), Comision=("Comision", "sum"), Neto=("Neto", "sum"))
        .reset_index()
        if not df.empty
        else pd.DataFrame()
    )
    envios = df[df["Concepto"] == "MELI Envíos"].copy() if not df.empty else df_vacio()
    return guardar_informe_excel(
        ruta,
        titulo=f"Ventas Mercado Pago · {sociedad}",
        subtitulo="Venta, MELI Envíos y comisiones (cuenta de la sociedad)",
        periodo=periodo,
        kpis=[
            ("Ventas brutas", metricas["ventas"], "money"),
            ("MELI Envíos", metricas["envios"], "money"),
            ("Comisiones MP/ML", metricas["comisiones"], "money"),
            ("Neto acreditado", metricas["neto"], "money"),
            ("Devoluciones", metricas["devoluciones"], "money"),
            ("Movimientos", metricas["movimientos"], "int"),
        ],
        resumenes=[("Por concepto", por_concepto), ("Por orden / venta", por_orden.head(80))],
        detalle=df,
        hoja_detalle="Movimientos",
        hojas_adicionales=[("Por orden", por_orden), ("MELI Envios", envios)],
        col_moneda=["Bruto", "Comision", "Neto", "Ventas", "MELI_Envios"],
        col_fecha=["Fecha"],
        col_texto=["Id", "Orden", "Hora", "Concepto", "Contraparte", "Detalle", "Tipo", "Medio"],
        total_col="Neto",
    )


def excel_ventas_bytes(df: pd.DataFrame, *, sociedad: str, periodo: str) -> bytes:
    metricas = kpis(df)
    por_orden = cuadro_por_orden(df)
    return exportar_informe_excel(
        titulo=f"Ventas Mercado Pago · {sociedad}",
        subtitulo="Venta, MELI Envíos y comisiones (cuenta de la sociedad)",
        periodo=periodo,
        kpis=[
            ("Ventas brutas", metricas["ventas"], "money"),
            ("MELI Envíos", metricas["envios"], "money"),
            ("Comisiones MP/ML", metricas["comisiones"], "money"),
            ("Neto acreditado", metricas["neto"], "money"),
        ],
        resumenes=[("Por orden / venta", por_orden.head(80))],
        detalle=df,
        hoja_detalle="Movimientos",
        hojas_adicionales=[("Por orden", por_orden)],
        col_moneda=["Bruto", "Comision", "Neto", "Ventas", "MELI_Envios"],
        col_fecha=["Fecha"],
        col_texto=["Id", "Orden", "Hora", "Concepto", "Contraparte", "Detalle", "Tipo", "Medio"],
        total_col="Neto",
    )
