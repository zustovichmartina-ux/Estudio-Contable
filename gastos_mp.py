"""Gastos personales de Mercado Pago (cuenta propia).

Libro permanente: cada sync suma movimientos, no borra el historial.
No mezclar con sociedades, ventas de Mercado Libre ni MELI Envíos.
"""

from __future__ import annotations

import csv
import io
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

DIR_LOCAL = Path.home() / ".estudio-contable" / "gastos_mp"
ARCHIVO_TOKEN = DIR_LOCAL / "token.txt"
ARCHIVO_REGLAS = DIR_LOCAL / "reglas.json"
ARCHIVO_LEDGER = DIR_LOCAL / "movimientos.csv"
ARCHIVO_META = DIR_LOCAL / "ultimo_sync_meta.json"

ESTADOS_CON_PLATA = {"approved", "accredited", "refunded", "charged_back"}
MAX_PAGOS_API = 8000

CATEGORIAS: list[str] = [
    "Alquiler y expensas",
    "Supermercado",
    "Delivery",
    "Restaurantes y cafés",
    "Nafta",
    "Uber y taxis",
    "Transporte (SUBE, peajes)",
    "Farmacia",
    "Salud",
    "Celular e internet",
    "Luz, gas y agua",
    "Streaming",
    "Indumentaria",
    "Mercado Libre / compras online",
    "Entretenimiento",
    "Hogar",
    "Educación",
    "Impuestos",
    "Transferencias enviadas",
    "Transferencias recibidas",
    "Ingreso desde banco",
    "Sueldo y honorarios",
    "Otros ingresos",
    "Devoluciones",
    "Retiro a banco",
    "Reservas Mercado Pago",
    "Movimiento interno MP",
    "Mercado Crédito",
    "Comisiones Mercado Pago",
    "Otros",
    "Sin clasificar",
]

CATEGORIAS_NO_GASTO = {
    "Sueldo y honorarios",
    "Otros ingresos",
    "Transferencias recibidas",
    "Ingreso desde banco",
    "Devoluciones",
    "Retiro a banco",
    "Reservas Mercado Pago",
    "Movimiento interno MP",
}

REGLAS_PALABRAS: list[tuple[tuple[str, ...], str]] = [
    (("pedidosya", "pedido ya", "dlo*pedidosya", "rappi", "uber eats", "ubereats"), "Delivery"),
    (("apple.com", "google one", "google *google", "youtubepremium", "youtube premium", "wl*google"), "Streaming"),
    (("mercado credito", "mercado crédito", "pago de cuotas", "pago de creditos de mercado", "pago de créditos de mercado"), "Mercado Crédito"),
    (("mostaza", "mcdonald", "burger king", "starbucks", "cafe martinez", "café martinez", "havanna", "starbucks", "fonte d", "gastronomia", "moma mdp", "sr kiosquero", "fina dulce", "la exitosa"), "Restaurantes y cafés"),
    (("ypf", "shell", "axion", "puma energy", "nafta", "combustible"), "Nafta"),
    (("uber", "cabify", "didi"), "Uber y taxis"),
    (("sube", "autovia", "peaje", "ausol", "autovía", "pasajes"), "Transporte (SUBE, peajes)"),
    (("carrefour", "coto", "jumbo", "vea", "disco", "changomas", "makro", "dia argentina", "supermercado dia", "compra en dia", "granja campo verde", "atlantic market", "drugstore argentina", "granja el buen sabor", "toledomp"), "Supermercado"),
    (("farmacity", "dr ahorro", "simplicity", "farmacia", "farmaonline"), "Farmacia"),
    (("osde", "swiss medical", "galeno", "medicus", "hospital", "clinica", "clínica"), "Salud"),
    (("netflix", "spotify", "youtube premium", "disney", "hbo", "prime video", "apple.com/bill", "icloud", "max oficial"), "Streaming"),
    (("personal flow", "telecom personal", "movistar", "claro", "fibertel", "telecentro", "tuenti"), "Celular e internet"),
    (("edenor", "edesur", "metrogas", "naturgy", "aysa", "edesal", "edelap"), "Luz, gas y agua"),
    (("mercadolibre", "mercado libre"), "Mercado Libre / compras online"),
    (("cinemark", "hoyts", "teatro", "steam", "playstation", "xbox", "basquet", "basket"), "Entretenimiento"),
    (("zara", "nike", "adidas", "levi"), "Indumentaria"),
    (("sodimac", "easy", "musimundo", "fravega", "garbarino"), "Hogar"),
    (("alquiler", "expensas", "inmobiliaria"), "Alquiler y expensas"),
    (("afip", "arca", "arba", "municipalidad", "abl ", "inmobiliario"), "Impuestos"),
    (("comision", "comisión", "cargo por", "fee mercadopago"), "Comisiones Mercado Pago"),
]

TIPOS_RETIRO = {"WITHDRAWAL", "PAYOUT", "WITHDRAWAL_CANCEL"}
TIPOS_DEVOLUCION = {"REFUND", "CHARGEBACK", "DISPUTE", "CASHBACK"}

COLUMNAS_REPORTE = [
    "TRANSACTION_DATE",
    "TRANSACTION_DATE_SHORT",
    "SETTLEMENT_DATE",
    "SOURCE_ID",
    "TRANSACTION_TYPE",
    "DESCRIPTION",
    "STORE_NAME",
    "PAYER_NAME",
    "PAYER_ID_NUMBER",
    "PAYER_ID_TYPE",
    "SALE_DETAIL",
    "EXTERNAL_REFERENCE",
    "PAYMENT_METHOD",
    "PAYMENT_METHOD_TYPE",
    "SETTLEMENT_NET_AMOUNT",
    "TRANSACTION_AMOUNT",
    "REAL_AMOUNT",
    "FEE_AMOUNT",
    "POI_WALLET_NAME",
    "POI_BANK_NAME",
    "POS_NAME",
    "BUSINESS_UNIT",
    "SUB_UNIT",
    "METADATA",
    "ISSUER_NAME",
    "FRANCHISE",
    "LAST_FOUR_DIGITS",
]

COLS_MOV = [
    "Id",
    "Fecha",
    "Hora",
    "Contraparte",
    "Detalle",
    "Categoria",
    "Importe",
    "Saldo",
    "Tipo",
    "Medio",
]


def asegurar_dir_local() -> Path:
    DIR_LOCAL.mkdir(parents=True, exist_ok=True)
    return DIR_LOCAL


def cargar_token() -> str:
    env = (os.environ.get("MP_ACCESS_TOKEN") or "").strip()
    if env:
        return env
    if ARCHIVO_TOKEN.exists():
        return ARCHIVO_TOKEN.read_text(encoding="utf-8").strip()
    return ""


def guardar_token(token: str) -> None:
    asegurar_dir_local()
    ARCHIVO_TOKEN.write_text(token.strip(), encoding="utf-8")


def cargar_reglas() -> dict[str, Any]:
    if not ARCHIVO_REGLAS.exists():
        return {"mapeos": {}, "extras": []}
    try:
        data = json.loads(ARCHIVO_REGLAS.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"mapeos": {}, "extras": []}
    if not isinstance(data, dict):
        return {"mapeos": {}, "extras": []}
    mapeos = data.get("mapeos") or {}
    extras = data.get("extras") or []
    if not isinstance(mapeos, dict):
        mapeos = {}
    if not isinstance(extras, list):
        extras = []
    return {"mapeos": {str(k).lower(): str(v) for k, v in mapeos.items()}, "extras": extras}


def guardar_reglas(reglas: dict[str, Any]) -> None:
    asegurar_dir_local()
    ARCHIVO_REGLAS.write_text(json.dumps(reglas, ensure_ascii=False, indent=2), encoding="utf-8")


def _norm(texto: object) -> str:
    return " ".join(
        str(texto or "")
        .lower()
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .split()
    )


def _limpio(valor: object) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = " ".join(str(valor).strip().split())
    if s.lower() in {"", "nan", "none", "-", "null"}:
        return ""
    return s


def _join_unicos(*partes: object) -> str:
    vistos: list[str] = []
    vistos_n: set[str] = set()
    for p in partes:
        t = _limpio(p)
        if not t:
            continue
        n = _norm(t)
        if n in vistos_n:
            continue
        vistos_n.add(n)
        vistos.append(t)
    return " · ".join(vistos)


def _clave_comercio(comercio: str) -> str:
    return _norm(comercio).strip()[:80]


def _textos_partition(branch: object, extra_ref: str = "") -> tuple[str, str]:
    b = str(branch or "")
    if "AM-to-POT" in b:
        contrap = "A reservas Mercado Pago"
        det = "Dinero disponible → reservas MP"
    elif "POT-to-AM" in b:
        contrap = "Desde reservas Mercado Pago"
        det = "Reservas MP → dinero disponible"
    else:
        contrap = "Movimiento interno MP"
        det = "Traspaso interno de Mercado Pago"
    return contrap, _join_unicos(det, b, extra_ref)


def _etiquetar_traspasos_internos(df: pd.DataFrame) -> pd.DataFrame:
    """Pone nombre claro a traspasos AM↔reservas y los que fondean Mercado Crédito."""
    if df.empty or "Tipo" not in df.columns:
        return df
    out = df.copy()
    tipo = out["Tipo"].astype(str).str.lower()
    part = tipo.eq("partition_transfer")
    if not part.any():
        return out
    blob = out["Contraparte"].fillna("").astype(str) + " " + out["Detalle"].fillna("").astype(str)
    am = blob.str.contains("AM-to-POT", case=False, regex=False)
    pot = blob.str.contains("POT-to-AM", case=False, regex=False)
    out.loc[part & am, "Contraparte"] = "A reservas Mercado Pago"
    out.loc[part & am, "Detalle"] = "Dinero disponible → reservas MP · AM-to-POT"
    out.loc[part & pot & ~am, "Contraparte"] = "Desde reservas Mercado Pago"
    out.loc[part & pot & ~am, "Detalle"] = "Reservas MP → dinero disponible · POT-to-AM"

    fechas = pd.to_datetime(out["Fecha"], errors="coerce")
    importes = pd.to_numeric(out["Importe"], errors="coerce").fillna(0.0).abs().round(2)
    texto_pago = (out["Contraparte"].fillna("").astype(str) + " " + out["Detalle"].fillna("").astype(str)).map(_norm)
    es_pago_credito = (~part) & texto_pago.str.contains("mercado credito", regex=False)
    if es_pago_credito.any():
        claves = set(zip(fechas[es_pago_credito].dt.strftime("%Y-%m-%d"), importes[es_pago_credito]))
        for idx in out.index[part]:
            fecha_txt = fechas.at[idx]
            if pd.isna(fecha_txt):
                continue
            clave = (pd.Timestamp(fecha_txt).strftime("%Y-%m-%d"), float(importes.at[idx]))
            if clave not in claves:
                continue
            if float(pd.to_numeric(out.at[idx, "Importe"], errors="coerce") or 0) < 0:
                out.at[idx, "Contraparte"] = "A Mercado Crédito"
                out.at[idx, "Detalle"] = "Dinero disponible → Mercado Crédito · AM-to-POT"
            else:
                out.at[idx, "Contraparte"] = "Desde Mercado Crédito"
                out.at[idx, "Detalle"] = "Mercado Crédito → dinero disponible · POT-to-AM"
    return out


def clasificar_movimiento(
    *,
    contraparte: str,
    detalle: str,
    tipo: str,
    importe: float,
    reglas: dict[str, Any] | None = None,
) -> str:
    reglas = reglas or cargar_reglas()
    tipo_u = str(tipo or "").strip().upper()
    tipo_l = str(tipo or "").strip().lower()
    imp = float(importe or 0)
    texto = f"{contraparte} {detalle} {tipo}"
    low = _norm(texto)

    if tipo_l == "partition_transfer":
        if "mercado credito" in low:
            return "Mercado Crédito"
        if "reserva" in low:
            return "Reservas Mercado Pago"
        return "Movimiento interno MP"
    if tipo_l == "account_fund":
        return "Ingreso desde banco"
    if tipo_l == "money_transfer" and imp > 0:
        return "Transferencias recibidas"
    if tipo_l == "money_transfer" and imp < 0:
        return "Transferencias enviadas"

    if tipo_u in TIPOS_RETIRO or "retiro" in low or "transferencia a banco" in low or "transferencia a tu banco" in low:
        return "Retiro a banco"
    if tipo_u in TIPOS_DEVOLUCION or "reintegro" in low or "devolucion" in low:
        return "Devoluciones"
    if "comision" in low or "comisión" in low:
        return "Comisiones Mercado Pago"

    clave = _clave_comercio(contraparte)
    mapeos: dict[str, str] = reglas.get("mapeos") or {}
    if clave and clave in mapeos and mapeos[clave] in CATEGORIAS:
        return mapeos[clave]

    for extra in reglas.get("extras") or []:
        if not isinstance(extra, dict):
            continue
        palabra = _norm(extra.get("contiene") or "")
        cat = str(extra.get("categoria") or "")
        if palabra and palabra in low and cat in CATEGORIAS:
            return cat

    for palabras, cat in REGLAS_PALABRAS:
        if any(p in low for p in palabras):
            return cat

    if "transfer" in low or "envio de dinero" in low or "envío de dinero" in low:
        return "Transferencias enviadas" if imp < 0 else "Otros ingresos"
    if imp > 0:
        return "Otros ingresos"
    if imp < 0:
        return "Sin clasificar"
    return "Sin clasificar"


def aplicar_clasificacion(df: pd.DataFrame, reglas: dict[str, Any] | None = None) -> pd.DataFrame:
    reglas = reglas or cargar_reglas()
    out = _etiquetar_traspasos_internos(df.copy())
    cats: list[str] = []
    for _, row in out.iterrows():
        cats.append(
            clasificar_movimiento(
                contraparte=str(row.get("Contraparte") or row.get("Comercio") or ""),
                detalle=str(row.get("Detalle") or ""),
                tipo=str(row.get("Tipo") or ""),
                importe=float(row.get("Importe") or 0),
                reglas=reglas,
            )
        )
    out["Categoria"] = cats
    return out


def aplicar_saldo(df: pd.DataFrame, saldo_final: float | None = None) -> pd.DataFrame:
    """Saldo después de cada movimiento. Si hay saldo_final (API o actual), encadena hacia atrás."""
    if df.empty:
        out = df.copy()
        out["Saldo"] = pd.Series(dtype=float)
        return out
    out = df.copy()
    out["_ord"] = range(len(out))
    cron = out.sort_values(["Fecha", "Hora", "Id"], ascending=[True, True, True], kind="mergesort")
    importes = pd.to_numeric(cron["Importe"], errors="coerce").fillna(0.0)
    if saldo_final is None:
        cron["Saldo"] = importes.cumsum()
    else:
        neto = float(importes.sum())
        inicial = float(saldo_final) - neto
        cron["Saldo"] = inicial + importes.cumsum()
    cron = cron.sort_values(["Fecha", "Hora", "Id"], ascending=[False, False, False], kind="mergesort")
    cron = cron.drop(columns=["_ord"], errors="ignore")
    return cron.reset_index(drop=True)


def aprender_desde_edicion(df_original: pd.DataFrame, df_editado: pd.DataFrame) -> int:
    if df_original.empty or df_editado.empty:
        return 0
    reglas = cargar_reglas()
    mapeos: dict[str, str] = dict(reglas.get("mapeos") or {})
    orig = df_original.set_index("Id", drop=False)
    nuevos = 0
    for _, row in df_editado.iterrows():
        mid = str(row.get("Id") or "")
        if not mid or mid not in orig.index:
            continue
        cat_nueva = str(row.get("Categoria") or "").strip()
        cat_vieja = str(orig.loc[mid, "Categoria"] or "").strip()
        if cat_nueva == cat_vieja or cat_nueva not in CATEGORIAS:
            continue
        clave = _clave_comercio(str(row.get("Contraparte") or orig.loc[mid].get("Contraparte") or ""))
        if not clave:
            continue
        if mapeos.get(clave) != cat_nueva:
            mapeos[clave] = cat_nueva
            nuevos += 1
    if nuevos:
        reglas["mapeos"] = mapeos
        guardar_reglas(reglas)
    return nuevos


def _buscar_col(df: pd.DataFrame, candidatos: tuple[str, ...]) -> str | None:
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidatos:
        if cand in norm_map:
            return norm_map[cand]
    for cand in candidatos:
        for n, orig in norm_map.items():
            if cand in n:
                return orig
    return None


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


def _col(df: pd.DataFrame, row: pd.Series, candidatos: tuple[str, ...]) -> str:
    col = _buscar_col(df, candidatos)
    if not col:
        return ""
    return _limpio(row.get(col))


def normalizar_csv_mp(df_raw: pd.DataFrame, saldo_final: float | None = None) -> pd.DataFrame:
    """Unifica CSV de Account Money (API) o de Actividad (web) con contraparte completa."""
    col_fecha = _buscar_col(
        df_raw,
        ("transaction_date", "settlement_date", "transaction_date_short", "fecha", "date"),
    )
    col_hora = _buscar_col(df_raw, ("hora", "hour", "time"))
    col_tipo = _buscar_col(df_raw, ("transaction_type", "tipo de operacion", "tipo de operación", "tipo"))
    col_importe = _buscar_col(
        df_raw,
        ("settlement_net_amount", "real_amount", "transaction_amount", "importe", "monto"),
    )
    col_id = _buscar_col(df_raw, ("source_id", "codigo de operacion", "código de operación", "order_id", "id"))
    col_medio = _buscar_col(df_raw, ("payment_method", "payment_method_type", "medio de pago", "producto"))

    filas: list[dict[str, Any]] = []
    for i, row in df_raw.iterrows():
        fecha, hora_dt = _a_fecha_hora(row[col_fecha]) if col_fecha else (None, "")
        hora = _limpio(row[col_hora]) if col_hora else hora_dt
        if not hora:
            hora = hora_dt

        destino = _col(df_raw, row, ("origen/destino", "destino", "origen", "contraparte", "beneficiario"))
        store = _col(df_raw, row, ("store_name", "pos_name", "comercio"))
        payer = _col(df_raw, row, ("payer_name", "nombre"))
        desc = _col(df_raw, row, ("description", "descripcion", "descripción", "concepto"))
        sale = _col(df_raw, row, ("sale_detail", "detalle de venta", "detalle"))
        banco_qr = _col(df_raw, row, ("poi_bank_name", "bank of origin"))
        wallet_qr = _col(df_raw, row, ("poi_wallet_name", "digital wallet"))
        ident = _col(df_raw, row, ("payer_id_number", "cuit", "dni", "cuil"))
        ext = _col(df_raw, row, ("external_reference", "referencia"))
        meta = _col(df_raw, row, ("metadata",))

        contrap = destino or store or payer or desc or "(sin dato)"
        detalle = _join_unicos(
            desc if _norm(desc) != _norm(contrap) else "",
            sale if _norm(sale) != _norm(contrap) else "",
            f"CUIT/DNI {ident}" if ident else "",
            banco_qr,
            wallet_qr,
            ext,
            meta,
        )
        tipo = _col(df_raw, row, ("transaction_type", "tipo de operacion", "tipo de operación", "tipo")) or "MOVIMIENTO"
        if col_tipo:
            tipo = _limpio(row[col_tipo]) or tipo
        importe = _a_float(row[col_importe]) if col_importe else 0.0
        mid = _limpio(row[col_id]) if col_id else f"fila-{i}"
        medio = _limpio(row[col_medio]) if col_medio else ""
        if fecha is None and contrap == "(sin dato)" and importe == 0:
            continue
        filas.append(
            {
                "Id": mid or f"fila-{i}",
                "Fecha": fecha or date.today(),
                "Hora": hora,
                "Contraparte": contrap,
                "Detalle": detalle,
                "Importe": round(importe, 2),
                "Tipo": tipo,
                "Medio": medio,
            }
        )
    df = pd.DataFrame(filas)
    if df.empty:
        return df_vacio()
    df = df.drop_duplicates(subset=["Id"], keep="last")
    df = aplicar_clasificacion(df)
    return aplicar_saldo(df, saldo_final=saldo_final)


def df_vacio() -> pd.DataFrame:
    return pd.DataFrame(columns=COLS_MOV)


def leer_csv_mp(contenido: bytes | str, nombre: str = "actividad.csv", saldo_final: float | None = None) -> pd.DataFrame:
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
    buf = io.StringIO(texto)
    df_raw = pd.read_csv(buf, sep=sep)
    df_raw.attrs["archivo"] = nombre
    return normalizar_csv_mp(df_raw, saldo_final=saldo_final)


def movimientos_demo(hoy: date | None = None) -> pd.DataFrame:
    hoy = hoy or date.today()
    raw = [
        (0, "14:32", "WITHDRAWAL", "BBVA Francés ****455", "Retiro a caja de ahorro CBU 0170", "account_money", -85000),
        (0, "21:10", "SETTLEMENT", "PedidosYa · Burger 54", "Pago QR · hamburguesa combo", "account_money", -18450),
        (1, "08:40", "SETTLEMENT", "YPF Av. Libertador", "Nafta súper · tarjeta MP", "debit_card", -42300),
        (1, "12:01", "SETTLEMENT", "Spotify", "Suscripción Premium mensual", "account_money", -2599),
        (2, "19:20", "SETTLEMENT", "Carrefour Palmares", "Super semanal", "qr", -67840),
        (2, "09:15", "SETTLEMENT", "Uber", "Viaje casa-oficina · Juan P.", "account_money", -3200),
        (3, "18:00", "SETTLEMENT", "Farmacity", "Farmacia · ibuprofeno", "account_money", -8900),
        (3, "11:30", "SETTLEMENT", "Starbucks Alto Palermo", "Café", "qr", -4500),
        (4, "10:00", "SETTLEMENT", "Netflix", "Suscripción estándar", "account_money", -9999),
        (5, "09:00", "TRANSFER", "María López", "Transferencia enviada · alquiler agosto · CVU 00000031000", "account_money", -280000),
        (5, "20:40", "SETTLEMENT", "Coto Digital", "Super", "account_money", -21450),
        (6, "08:00", "SETTLEMENT", "Personal Flow", "Internet + celular", "debit_card", -38999),
        (7, "13:20", "SETTLEMENT", "Rappi", "Almuerzo · sushi", "account_money", -12500),
        (7, "21:00", "SETTLEMENT", "Cinemark Palermo", "Cine · 2 entradas", "qr", -15600),
        (8, "17:10", "SETTLEMENT", "Shell", "Nafta", "debit_card", -38100),
        (9, "16:00", "SETTLEMENT", "Mercado Libre", "Auriculares XYZ · vendedor TECHAR", "account_money", -24999),
        (10, "08:30", "SETTLEMENT", "Edenor", "Luz · servicio agosto", "account_money", -45200),
        (11, "13:05", "SETTLEMENT", "McDonalds Santa Fe", "Almuerzo", "qr", -8900),
        (12, "15:00", "REFUND", "PedidosYa", "Reintegro pedido cancelado", "account_money", 18450),
        (13, "19:50", "SETTLEMENT", "Jumbo Unicenter", "Super", "qr", -54120),
        (14, "14:00", "SETTLEMENT", "Nike Store", "Zapatillas", "debit_card", -79999),
        (15, "10:00", "TRANSFER", "Estudio / haberes", "Transferencia recibida · honorarios", "account_money", 850000),
        (16, "11:40", "SETTLEMENT", "Easy", "Lámpara y cables", "account_money", -32800),
        (17, "08:10", "SETTLEMENT", "SUBE", "Carga", "account_money", -5000),
        (18, "16:45", "SETTLEMENT", "Havanna", "Café y medialunas", "qr", -6200),
        (19, "18:20", "SETTLEMENT", "Axion energy", "Nafta", "debit_card", -35500),
        (20, "20:00", "SETTLEMENT", "Dia Argentina", "Super", "qr", -18700),
        (21, "09:00", "SETTLEMENT", "YouTube Premium", "Suscripción", "account_money", -1599),
        (22, "13:30", "SETTLEMENT", "Mostaza", "Almuerzo", "qr", -7400),
        (23, "11:00", "SETTLEMENT", "ARBA", "IIBB", "account_money", -12000),
        (24, "19:00", "PAYOUT", "Cajero BBVA", "Extracción efectivo", "atm", -40000),
        (25, "15:20", "SETTLEMENT", "Fravega", "Auriculares", "account_money", -45999),
        (26, "08:50", "SETTLEMENT", "Cabify", "Viaje", "account_money", -4100),
        (27, "21:10", "SETTLEMENT", "Disco", "Super", "qr", -29800),
        (28, "09:15", "SETTLEMENT", "Metrogas", "Gas", "account_money", -22100),
        (29, "17:00", "SETTLEMENT", "Zara", "Ropa", "debit_card", -62300),
    ]
    filas = []
    for i, (hace, hora, tipo, contrap, detalle, medio, importe) in enumerate(raw):
        filas.append(
            {
                "Id": f"demo-{i+1:03d}",
                "Fecha": hoy - timedelta(days=hace),
                "Hora": hora,
                "Contraparte": contrap,
                "Detalle": detalle,
                "Importe": float(importe),
                "Tipo": tipo,
                "Medio": medio,
            }
        )
    df = aplicar_clasificacion(pd.DataFrame(filas))
    return aplicar_saldo(df, saldo_final=245800.0)


def _mp_request(method: str, url: str, token: str, body: dict | None = None) -> tuple[int, Any]:
    data = None
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json, text/csv, */*",
        "User-Agent": "EstudioContable-GastosMP/1.0",
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


def saldo_mp(token: str, user_id: object) -> float | None:
    if not user_id:
        return None
    data = _mp_try(
        "GET",
        f"https://api.mercadopago.com/users/{user_id}/mercadopago_account/balance",
        token,
    )
    if not isinstance(data, dict):
        return None
    for key in ("available_balance", "available", "total_amount", "amount"):
        if key in data and data[key] is not None:
            try:
                return float(data[key])
            except (TypeError, ValueError):
                continue
    av = data.get("available")
    if isinstance(av, dict) and av.get("amount") is not None:
        try:
            return float(av["amount"])
        except (TypeError, ValueError):
            return None
    return None


def _configurar_reporte(token: str) -> None:
    actual = _mp_try("GET", "https://api.mercadopago.com/v1/account/settlement_report/config", token)
    actual = actual if isinstance(actual, dict) else {}
    freq = actual.get("frequency") or {"hour": 0, "type": "daily", "value": None, "format": "CSV"}
    body = {
        "file_name_prefix": actual.get("file_name_prefix") or "gastos-personales",
        "include_withdraw": True,
        "header_language": "es",
        "display_timezone": actual.get("display_timezone") or "GMT-03",
        "separator": ";",
        "frequency": freq,
        "columns": [{"key": k, "alias": ""} for k in COLUMNAS_REPORTE],
    }
    _mp_try("PUT", "https://api.mercadopago.com/v1/account/settlement_report/config", token, body)


def _importe_pago(p: dict[str, Any], my_id: object) -> float:
    try:
        bruto = float(p.get("transaction_amount") or 0)
    except (TypeError, ValueError):
        bruto = 0.0
    op = str(p.get("operation_type") or "").lower()
    poi = p.get("point_of_interaction") if isinstance(p.get("point_of_interaction"), dict) else {}
    biz = poi.get("business_info") if isinstance(poi.get("business_info"), dict) else {}
    branch = str(biz.get("branch") or "")
    collector = str(p.get("collector_id") or "")
    yo = str(my_id or "")

    if op in {"regular_payment", "recurring_payment"}:
        return -abs(bruto)
    if op == "account_fund":
        return abs(bruto)
    if op == "partition_transfer":
        if "AM-to-POT" in branch:
            return -abs(bruto)
        return abs(bruto)
    if op == "money_transfer":
        if yo and collector == yo:
            return abs(bruto)
        return -abs(bruto)
    if collector == yo:
        return abs(bruto)
    return -abs(bruto)


def _pagos_a_filas(pagos: list[dict[str, Any]], my_id: object = None) -> list[dict[str, Any]]:
    filas: list[dict[str, Any]] = []
    for p in pagos:
        if not isinstance(p, dict):
            continue
        if str(p.get("status") or "") not in ESTADOS_CON_PLATA:
            continue
        payer = p.get("payer") if isinstance(p.get("payer"), dict) else {}
        nombre = _join_unicos(payer.get("first_name"), payer.get("last_name"))
        email = _limpio(payer.get("email"))
        ident_obj = payer.get("identification") if isinstance(payer.get("identification"), dict) else {}
        ident = _limpio(ident_obj.get("number"))
        desc = _limpio(p.get("description"))
        poi = p.get("point_of_interaction") if isinstance(p.get("point_of_interaction"), dict) else {}
        biz = poi.get("business_info") if isinstance(poi.get("business_info"), dict) else {}
        add = p.get("additional_info") if isinstance(p.get("additional_info"), dict) else {}
        items = add.get("items") if isinstance(add.get("items"), list) else []
        titulos = ", ".join(
            t
            for t in (
                _limpio(i.get("title") or i.get("description"))
                for i in items
                if isinstance(i, dict)
            )
            if t
        )
        op = str(p.get("operation_type") or "payment")
        importe = round(_importe_pago(p, my_id), 2)
        fecha, hora = _a_fecha_hora(p.get("date_created") or p.get("date_approved"))

        if op in {"regular_payment", "recurring_payment"}:
            contrap = desc or titulos or _limpio(p.get("statement_descriptor")) or "Compra Mercado Pago"
        elif op == "account_fund":
            contrap = desc or "Transferencia bancaria"
        elif op == "partition_transfer":
            contrap, detalle = _textos_partition(biz.get("branch"), _limpio(p.get("external_reference")))
        elif op == "money_transfer":
            contrap = nombre or email or desc or "Transferencia Mercado Pago"
        else:
            contrap = desc or titulos or nombre or email or "(sin dato)"

        if op != "partition_transfer":
            detalle = _join_unicos(
                desc if _norm(desc) != _norm(contrap) else "",
                titulos if _norm(titulos) != _norm(contrap) else "",
                f"CUIT/DNI {ident}" if ident else "",
                email if email and _norm(email) not in _norm(contrap) else "",
                p.get("payment_method_id"),
                p.get("statement_descriptor"),
                biz.get("branch"),
                poi.get("type"),
            )
        filas.append(
            {
                "Id": str(p.get("id") or ""),
                "Fecha": fecha or date.today(),
                "Hora": hora,
                "Contraparte": contrap,
                "Detalle": detalle,
                "Importe": importe,
                "Tipo": op,
                "Medio": _limpio(p.get("payment_type_id") or p.get("payment_method_id")),
            }
        )
        if str(p.get("status") or "") in {"refunded", "charged_back"} and str(p.get("id") or ""):
            fecha_r, hora_r = _a_fecha_hora(p.get("date_last_updated") or p.get("date_created"))
            filas.append(
                {
                    "Id": f"{p.get('id')}-reintegro",
                    "Fecha": fecha_r or fecha or date.today(),
                    "Hora": hora_r or hora,
                    "Contraparte": f"Reintegro · {contrap}",
                    "Detalle": _join_unicos("Devolución Mercado Pago", detalle),
                    "Importe": round(-importe, 2),
                    "Tipo": "REFUND" if str(p.get("status")) == "refunded" else "CHARGEBACK",
                    "Medio": _limpio(p.get("payment_type_id") or p.get("payment_method_id")),
                }
            )
    return filas


def buscar_pagos(token: str, desde: date, hasta: date, my_id: object = None) -> pd.DataFrame:
    begin = datetime(desde.year, desde.month, desde.day).strftime("%Y-%m-%dT00:00:00.000-03:00")
    end = datetime(hasta.year, hasta.month, hasta.day, 23, 59, 59).strftime("%Y-%m-%dT23:59:59.000-03:00")
    filas: list[dict[str, Any]] = []
    offset = 0
    while offset < MAX_PAGOS_API:
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
        filas.extend(_pagos_a_filas(results, my_id=my_id))
        paging = payload.get("paging") if isinstance(payload.get("paging"), dict) else {}
        total = int(paging.get("total") or 0)
        offset += len(results)
        if offset >= total:
            break
    if not filas:
        return df_vacio()
    df = pd.DataFrame(filas)
    df = df[df["Id"].astype(str) != ""].drop_duplicates(subset=["Id"])
    return aplicar_clasificacion(df)


def bajar_reporte_account_money(token: str, desde: date, hasta: date) -> pd.DataFrame:
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
            "Probá de nuevo en un minuto o subí el CSV de Actividad."
        )
    _status, raw = _mp_request(
        "GET",
        f"https://api.mercadopago.com/v1/account/settlement_report/{urllib.parse.quote(str(file_name))}",
        token,
    )
    if isinstance(raw, dict):
        raise RuntimeError(f"Respuesta inesperada al descargar el reporte: {raw}")
    if isinstance(raw, bytes):
        return leer_csv_mp(raw, str(file_name))
    return leer_csv_mp(str(raw), str(file_name))


def _calidad_fila(row: pd.Series) -> int:
    contrap = str(row.get("Contraparte") or "").strip()
    detalle = str(row.get("Detalle") or "").strip()
    tipo = str(row.get("Tipo") or "").strip().lower()
    score = 0
    if contrap and contrap not in {"(sin dato)", "Pago recibido"}:
        score += 100
    score += min(len(detalle), 80)
    if tipo in {"settlement", "settlement_report", "movimiento", ""}:
        score -= 50
    if tipo in {
        "refund",
        "chargeback",
        "regular_payment",
        "recurring_payment",
        "money_transfer",
        "account_fund",
        "partition_transfer",
    }:
        score += 40
    return score


def _merge_movimientos(a: pd.DataFrame, b: pd.DataFrame) -> pd.DataFrame:
    partes = [x for x in (a, b) if x is not None and not x.empty]
    if not partes:
        return df_vacio()
    cats_prev: dict[str, str] = {}
    for parte in partes:
        if "Id" in parte.columns and "Categoria" in parte.columns:
            for mid, cat in zip(parte["Id"].astype(str), parte["Categoria"].astype(str)):
                if cat and cat not in {"", "Sin clasificar", "nan"} and cat in CATEGORIAS:
                    cats_prev[str(mid)] = cat
    df = pd.concat(partes, ignore_index=True)
    df["_q"] = df.apply(_calidad_fila, axis=1)
    df = df.sort_values("_q", ascending=False).drop_duplicates(subset=["Id"], keep="first")
    df = df.drop(columns=["_q"])
    df = aplicar_clasificacion(df)
    if cats_prev:
        ids = df["Id"].astype(str)
        for idx, mid in ids.items():
            prev = cats_prev.get(str(mid))
            if prev:
                df.at[idx, "Categoria"] = prev
    return df


def sincronizar_mercadopago(
    token: str,
    desde: date,
    hasta: date,
    *,
    rapido: bool = False,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Suma al libro permanente los movimientos del período. No borra el historial anterior."""
    avisos: list[str] = []
    me = usuario_mp(token)
    if not me:
        raise RuntimeError(
            "El token no sirve o no tiene permiso. Generá uno de producción en "
            "Tus integraciones (APP_USR-…) con tu propia cuenta."
        )
    nombre = _join_unicos(me.get("first_name"), me.get("last_name"), me.get("nickname")) or str(me.get("id"))
    saldo_actual = saldo_mp(token, me.get("id"))
    reporte = df_vacio()
    if not rapido:
        try:
            reporte = bajar_reporte_account_money(token, desde, hasta)
        except RuntimeError as exc:
            avisos.append(str(exc))
    pagos = buscar_pagos(token, desde, hasta, my_id=me.get("id"))
    nuevos = _merge_movimientos(reporte, pagos)
    ledger = cargar_ledger()
    df = _merge_movimientos(ledger, nuevos) if not ledger.empty else nuevos
    if df.empty:
        raise RuntimeError(
            "La API no devolvió movimientos en ese período. "
            "En Mercado Pago: Actividad → Descargar CSV e importalo. "
            + (" ".join(avisos) if avisos else "")
        )
    if saldo_actual is None:
        avisos.append("No pude leer el saldo actual; el saldo de la planilla arranca en 0 al inicio del período.")
    df = aplicar_saldo(df, saldo_final=saldo_actual)
    meta = {
        "usuario": nombre,
        "email": _limpio(me.get("email")),
        "saldo_actual": saldo_actual,
        "avisos": avisos,
        "movimientos": int(len(df)),
        "filas_reporte": int(len(reporte)),
        "filas_pagos": int(len(pagos)),
        "nuevos_periodo": int(len(nuevos)),
        "desde": str(desde),
        "hasta": str(hasta),
        "rapido": rapido,
    }
    guardar_ledger(df, meta)
    return df, meta


def guardar_ledger(df: pd.DataFrame, meta: dict[str, Any] | None = None) -> None:
    asegurar_dir_local()
    cols = [c for c in COLS_MOV if c in df.columns]
    df[cols].to_csv(ARCHIVO_LEDGER, index=False, encoding="utf-8-sig")
    df[cols].to_csv(DIR_LOCAL / "ultimo_sync.csv", index=False, encoding="utf-8-sig")
    if meta is not None:
        ARCHIVO_META.write_text(
            json.dumps(meta, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )


def cargar_ledger() -> pd.DataFrame:
    for path in (ARCHIVO_LEDGER, DIR_LOCAL / "ultimo_sync.csv"):
        if not path.exists():
            continue
        df = pd.read_csv(path)
        if df.empty or "Contraparte" not in df.columns:
            continue
        if "Fecha" in df.columns:
            df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.date
        if "Id" in df.columns:
            df["Id"] = df["Id"].astype(str)
        return df
    return df_vacio()


def incorporar_al_libro(nuevos: pd.DataFrame, saldo_final: float | None = None) -> pd.DataFrame:
    """Suma filas (CSV o API) al libro permanente."""
    df = _merge_movimientos(cargar_ledger(), nuevos) if not nuevos.empty else cargar_ledger()
    if df.empty:
        return df_vacio()
    df = aplicar_saldo(df, saldo_final=saldo_final)
    meta = {"movimientos": int(len(df)), "fuente": "csv"}
    if ARCHIVO_META.exists():
        try:
            meta = {**json.loads(ARCHIVO_META.read_text(encoding="utf-8")), **meta}
        except json.JSONDecodeError:
            pass
    guardar_ledger(df, meta)
    return df


def _guardar_ultimo_sync(df: pd.DataFrame, meta: dict[str, Any]) -> None:
    guardar_ledger(df, meta)


def cargar_ultimo_sync() -> tuple[pd.DataFrame, dict[str, Any]] | None:
    df = cargar_ledger()
    if df.empty:
        return None
    meta: dict[str, Any] = {}
    if ARCHIVO_META.exists():
        try:
            meta = json.loads(ARCHIVO_META.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            meta = {}
    return df, meta


def _mascara_traspaso(df: pd.DataFrame) -> pd.Series:
    if "Tipo" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["Tipo"].astype(str).str.lower().eq("partition_transfer")


def resumen_por_categoria(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Categoria", "Gastado", "Ingresos", "Movimientos"])
    imp = pd.to_numeric(df["Importe"], errors="coerce")
    interno = _mascara_traspaso(df)
    g = (
        df.loc[~interno & ~df["Categoria"].isin(CATEGORIAS_NO_GASTO) & (imp < 0)]
        .groupby("Categoria", dropna=False)["Importe"]
        .sum()
        .abs()
        .rename("Gastado")
    )
    i = (
        df.loc[~interno & (imp > 0)]
        .groupby("Categoria", dropna=False)["Importe"]
        .sum()
        .rename("Ingresos")
    )
    n = df.groupby("Categoria", dropna=False)["Id"].count().rename("Movimientos")
    out = pd.concat([g, i, n], axis=1).fillna(0).reset_index()
    out["Gastado"] = out["Gastado"].astype(float)
    out["Ingresos"] = out["Ingresos"].astype(float)
    out["Movimientos"] = out["Movimientos"].astype(int)
    return out.sort_values("Gastado", ascending=False)


def kpis(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {
            "gastado": 0.0,
            "ingresos": 0.0,
            "neto": 0.0,
            "saldo": 0.0,
            "sin_clasificar": 0,
            "movimientos": 0,
        }
    imp = pd.to_numeric(df["Importe"], errors="coerce").fillna(0.0)
    interno = _mascara_traspaso(df)
    gastos = df.loc[~interno & ~df["Categoria"].isin(CATEGORIAS_NO_GASTO) & (imp < 0), "Importe"]
    ingresos = df.loc[~interno & (imp > 0), "Importe"]
    saldo = 0.0
    if "Saldo" in df.columns and df["Saldo"].notna().any():
        # Planilla en orden más reciente primero
        saldo = float(pd.to_numeric(df["Saldo"], errors="coerce").dropna().iloc[0])
    return {
        "gastado": float(gastos.abs().sum()) if len(gastos) else 0.0,
        "ingresos": float(ingresos.sum()) if len(ingresos) else 0.0,
        "neto": float(imp.sum()),
        "saldo": saldo,
        "sin_clasificar": int((df["Categoria"] == "Sin clasificar").sum()),
        "movimientos": int(len(df)),
    }
