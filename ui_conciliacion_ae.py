# -*- coding: utf-8 -*-
"""Conciliación bancaria: extracto completo → imputación en la línea → asiento Tango."""
from __future__ import annotations

import calendar
import copy
import html
import json
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import database as db
from capa_revision import gate_asiento, resolver_codigo_plan
from conceptos_bancos import CACHE_PATH, cargar_instructivo
from motor_conciliacion import (
    CATEGORIA_A_CUENTA_HINT,
    cuenta_banco_del_plan,
    correr_motor,
    df_extracto_a_filas,
    es_fila_saldo_bancario,
    money,
    origen_linea_extracto,
    renglones_asiento_banco_mes,
    validar_saldos_corridos,
)
from procesador import (
    AsientoDevengamiento,
    ExportacionTangoError,
    RenglonAsiento,
    cargar_plan_cuentas,
    clasificar_movimiento_extracto,
    generar_excel_tango_nativo,
    guardar_biblioteca_persistida,
    normalizar_codigo_cuenta_tango,
    procesar_extractos_bancarios_pdfs,
)


class _UploadMemoria:
    """Bytes del PDF guardados entre reruns de Streamlit."""

    def __init__(self, name: str, data: bytes) -> None:
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data

_CSS = """
<style>
.ce-top{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin:0 0 12px 0;}
.ce-title{font-size:1.35rem;font-weight:700;color:#1F4E79;margin:0;}
.ce-sub{color:#5c6379;font-size:0.9rem;margin:2px 0 0 0;}
.ce-chip{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:999px;margin:8px 6px 0 0;}
.ce-chip.regla{background:#e6f4ea;color:#137333;}
.ce-chip.sugerido{background:#fff4e5;color:#b06000;}
.ce-chip.a_clasificar{background:#fce8e6;color:#c5221f;}
.ce-wrap{border:1px solid #e6eaf0;border-radius:10px;overflow:auto;max-height:420px;background:#fff;}
.ce-table{width:100%;border-collapse:collapse;font-family:"Segoe UI",Calibri,sans-serif;}
.ce-table th{background:#1F4E79;color:#fff;font-size:12px;font-weight:600;text-align:left;padding:10px;white-space:nowrap;}
.ce-table td{padding:8px 10px;font-size:13px;border-bottom:1px solid #e6eaf0;}
.ce-table tr:nth-child(even) td{background:#f7f9fc;}
.ce-table tr.pend td{background:#fff6f5;}
.ce-table .num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;}
div[data-testid="stCustomComponentV1"] iframe{border:0 !important;}
</style>
"""

_EXTRACTO_GRID = components.declare_component(
    "extracto_grid",
    path=str(Path(__file__).resolve().parent / "extracto_grid"),
)

_CLASIFS_BASE = (
    "Pago ARBA",
    "SIRCREB",
    "IIBB",
    "Ingresos brutos Tucuman",
    "Impuestos a los débitos y créditos",
    "Impuesto a los sellos",
    "Percepción IVA",
    "IVA",
    "Rescate FIMA",
    "Suscripción FCI",
    "Gastos Bancarios",
    "Inversiones",
    "Intereses",
    "Depositos en efvo",
    "Cheques recibidos",
    "Cheques emitidos",
    "Pago de haberes",
    "Pagos AFIP",
    "Pagos tarjeta corporativa",
    "Compras",
    "Pago de Servicios",
    "Transferencias recibidas",
    "Transferencias emitidas",
    "Pagos recibidos",
    "Acreditaciones comercios",
)
_CLASIF_SIN_CUENTA = {
    "",
    "sin clasificar",
    "movimientos a identificar",
    "identificar",
    "otro",
    "otro…",
    "otro...",
}
_MAPA_CLASIF_PATH = Path(__file__).resolve().parent / "data" / "clasif_extracto.json"


def _norm_clasif(texto: str) -> str:
    t = str(texto or "").strip().lower()
    t = re.sub(r"\s+", " ", t)
    return t


def _clasif_tiene_cuenta(nombre: str) -> bool:
    return _norm_clasif(nombre) not in _CLASIF_SIN_CUENTA


def _cargar_mapa_clasif(sociedad_id: int) -> dict[str, str]:
    sid = str(int(sociedad_id))
    cache = st.session_state.get(f"clasif_map_{sid}")
    if isinstance(cache, dict):
        return dict(cache)
    try:
        raw = json.loads(_MAPA_CLASIF_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        raw = {}
    mapa = raw.get(sid) if isinstance(raw, dict) else {}
    if not isinstance(mapa, dict):
        mapa = {}
    clean = {
        str(k).strip(): str(v).strip()
        for k, v in mapa.items()
        if _clasif_tiene_cuenta(str(k)) and str(v).strip() not in {"", "99999"}
    }
    st.session_state[f"clasif_map_{sid}"] = clean
    return clean


def _guardar_mapa_clasif(sociedad_id: int, mapa: dict[str, str]) -> None:
    sid = str(int(sociedad_id))
    clean = {
        str(k).strip(): str(v).strip()
        for k, v in (mapa or {}).items()
        if _clasif_tiene_cuenta(str(k)) and str(v).strip() not in {"", "99999"}
    }
    st.session_state[f"clasif_map_{sid}"] = clean
    try:
        raw = json.loads(_MAPA_CLASIF_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        raw = {}
    if not isinstance(raw, dict):
        raw = {}
    raw[sid] = clean
    _MAPA_CLASIF_PATH.parent.mkdir(parents=True, exist_ok=True)
    _MAPA_CLASIF_PATH.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")


def _codigo_de_clasif(nombre: str, mapa: dict[str, str]) -> str:
    if not _clasif_tiene_cuenta(nombre):
        return ""
    n = _norm_clasif(nombre)
    for k, v in (mapa or {}).items():
        if _norm_clasif(k) == n and str(v).strip() not in {"", "99999"}:
            return str(v).strip()
    return ""


def _aplicar_mapa_clasif(movs: list[dict], mapa: dict[str, str]) -> list[dict]:
    if not mapa:
        return movs
    out = list(movs)
    for m in out:
        nombre = str(m.get("categoria") or m.get("extracto_label") or "")
        cod = _codigo_de_clasif(nombre, mapa)
        actual = str(m.get("cuenta_codigo") or "").strip()
        if not cod or actual not in {"", "99999"}:
            continue
        m["cuenta_codigo"] = cod
        if str(m.get("origen") or "") == "a_clasificar":
            m["origen"] = "sugerido"
    return out


def _mapa_desde_movimientos(movs: list[dict], base: dict[str, str] | None = None) -> dict[str, str]:
    mapa = dict(base or {})
    for m in movs:
        nombre = str(m.get("categoria") or "").strip()
        codigo = str(m.get("cuenta_codigo") or "").strip()
        if _clasif_tiene_cuenta(nombre) and codigo not in {"", "99999"}:
            mapa[nombre] = codigo
    return mapa


def _clasifs_componente(sociedad_id: int, movs: list[dict]) -> list[dict]:
    mapa = _mapa_desde_movimientos(movs, _cargar_mapa_clasif(sociedad_id))
    nombres: list[str] = []
    vistos: set[str] = set()
    for n in list(_CLASIFS_BASE) + [str(m.get("categoria") or "") for m in movs] + list(mapa.keys()):
        nom = str(n or "").strip()
        key = _norm_clasif(nom)
        if not nom or key in vistos or key in {"otro", "otro…", "otro..."}:
            continue
        vistos.add(key)
        nombres.append(nom)
    cola = [n for n in nombres if not _clasif_tiene_cuenta(n)]
    cuerpo = [n for n in nombres if _clasif_tiene_cuenta(n)]
    cuerpo.sort(key=lambda x: x.lower())
    return [{"nombre": n, "codigo": mapa.get(n) or ""} for n in cuerpo + cola]
    if not d:
        d = date.today()
    return d.strftime("%m/%Y")


def _periodo_mm_yyyy(d: date | None) -> str:
    if not d:
        d = date.today()
    return d.strftime("%m/%Y")


def _fmt_fecha(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, date):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if "T" in s:
        s = s.split("T", 1)[0]
    if len(s) >= 10 and s[4] == "-":
        try:
            y, m, d = s[:10].split("-")
            return f"{d}/{m}/{y}"
        except ValueError:
            return s
    return s


def _fmt_money(n: float) -> str:
    return f"$ {n:,.2f}"


def _label_origen(origen: str) -> str:
    return {
        "regla": "Regla",
        "sugerido": "Sugerido",
        "a_clasificar": "A clasificar",
    }.get(origen, origen or "A clasificar")


def _col_en_plan(plan_df: pd.DataFrame | None, nombres: tuple[str, ...]) -> str | None:
    if plan_df is None or getattr(plan_df, "empty", True):
        return None
    mapa = {str(c).strip().lower(): c for c in plan_df.columns}
    for n in nombres:
        if n in mapa:
            return mapa[n]
    return None


def _plan_df_sociedad(sociedad_id: int) -> pd.DataFrame | None:
    """Plan de cuentas de la sociedad activa (session o disco)."""
    df = st.session_state.get("plan_cuentas_df")
    cid = st.session_state.get("plan_cuentas_cliente_id")
    if (
        df is not None
        and not getattr(df, "empty", True)
        and (cid in (None, sociedad_id) or int(cid) == int(sociedad_id))
    ):
        return df
    df2 = st.session_state.get(f"plan_cuentas_df_{int(sociedad_id)}")
    if df2 is not None and not getattr(df2, "empty", True):
        return df2
    try:
        cliente = db.obtener_cliente(int(sociedad_id))
    except Exception:
        return df
    if not cliente:
        return df
    candidatos: list[Path] = []
    raw = str(cliente.get("plan_cuentas_path") or "").strip()
    if raw:
        candidatos.append(Path(raw))
    cuit = re.sub(r"\D", "", str(cliente.get("cuit") or ""))
    base = Path(__file__).resolve().parent / "data" / "planes_cuentas"
    if cuit:
        candidatos.append(base / f"plan_{cuit}.xlsx")
        candidatos.append(base / f"plan_{cuit}.csv")
    candidatos.append(base / f"plan_id_{int(sociedad_id)}.xlsx")
    for p in candidatos:
        if not p.is_file():
            continue
        try:
            loaded = cargar_plan_cuentas(p)
        except Exception:
            continue
        if loaded is None or getattr(loaded, "empty", True):
            continue
        st.session_state["plan_cuentas_df"] = loaded
        st.session_state[f"plan_cuentas_df_{int(sociedad_id)}"] = loaded
        st.session_state["plan_cuentas_cliente_id"] = int(sociedad_id)
        return loaded
    return df


def _opciones_plan(plan_df: pd.DataFrame | None) -> list[str]:
    opts = ["99999 — A clasificar"]
    if plan_df is None or getattr(plan_df, "empty", True):
        return opts
    col_cod = _col_en_plan(plan_df, ("codigo", "código", "cuenta", "cod")) or plan_df.columns[0]
    col_desc = _col_en_plan(
        plan_df, ("descripcion", "descripción", "nombre", "desc")
    )
    if col_desc is None and len(plan_df.columns) > 1:
        col_desc = plan_df.columns[1]
    vistos: set[str] = set()
    filas: list[tuple[str, str]] = []
    for _, row in plan_df.iterrows():
        cod = normalizar_codigo_cuenta_tango(row.get(col_cod))
        if not cod or cod in vistos or cod.lower() in {"nan", "codigo", "código", "cuenta"}:
            continue
        vistos.add(cod)
        desc = str(row.get(col_desc) or "").strip() if col_desc is not None else ""
        filas.append((cod, desc))
    filas.sort(key=lambda x: x[0])
    for cod, desc in filas:
        opts.append(f"{cod} — {desc}" if desc else cod)
    return opts


def _clave_mov_repetido(m: dict) -> str:
    t = str(
        m.get("categoria")
        or m.get("extracto_label")
        or m.get("clasif")
        or m.get("descripcion")
        or m.get("desc")
        or ""
    )
    t = t.lower().strip()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\b\d{5,}\b", "", t)
    return t.strip()


def _propagar_cuentas_repetidas(movs: list[dict]) -> list[dict]:
    """La misma descripción/clasificación queda en la misma cuenta Tango."""
    grupos: dict[str, list[int]] = {}
    for i, m in enumerate(movs):
        clave = _clave_mov_repetido(m)
        if not clave:
            continue
        grupos.setdefault(clave, []).append(i)
    out = list(movs)
    for idxs in grupos.values():
        if len(idxs) < 2:
            continue
        codigos = [
            str(out[i].get("cuenta_codigo") or "").strip()
            for i in idxs
            if str(out[i].get("cuenta_codigo") or "").strip() not in {"", "99999"}
        ]
        if not codigos:
            continue
        elegido = max(set(codigos), key=codigos.count)
        modelo = next(out[i] for i in idxs if str(out[i].get("cuenta_codigo") or "").strip() == elegido)
        for i in idxs:
            actual = str(out[i].get("cuenta_codigo") or "").strip()
            if actual not in {"", "99999"}:
                continue
            out[i]["cuenta_codigo"] = elegido
            out[i]["cuenta_plan"] = modelo.get("cuenta_plan") or out[i].get("cuenta_plan")
            out[i]["origen"] = modelo.get("origen") or "sugerido"
    return out


def _codigo_desde_opcion(texto: str) -> str:
    s = str(texto or "").strip()
    if "—" in s:
        s = s.split("—", 1)[0].strip()
    elif " - " in s:
        s = s.split(" - ", 1)[0].strip()
    return s or "99999"


def _opcion_desde_codigo(codigo: str, opciones: list[str]) -> str:
    cod = str(codigo or "99999").strip() or "99999"
    for op in opciones:
        if _codigo_desde_opcion(op) == cod:
            return op
    return opciones[0] if opciones else "99999 — A clasificar"


def _enriquecer(movs: list[dict], plan_df, mapa_clasif: dict[str, str] | None = None) -> list[dict]:
    out = []
    for i, m in enumerate(movs):
        fila = dict(m)
        cred = float(money(m.get("credito")))
        deb = float(money(m.get("debito")))
        label = clasificar_movimiento_extracto(
            str(m.get("descripcion") or ""),
            importe=cred - deb,
        )
        fila["extracto_label"] = label
        categoria = str(m.get("categoria") or "")
        codigo, desc_plan, score = resolver_codigo_plan(
            categoria if categoria and "identificar" not in categoria.lower() else label,
            plan_df,
            hints=CATEGORIA_A_CUENTA_HINT,
        )
        fuente = str(m.get("fuente") or "")
        identificado = "identificar" not in categoria.lower()
        if fuente in {"conceptos_bancos", "regla_local"} and identificado and codigo != "99999":
            fila["cuenta_codigo"] = codigo
            fila["cuenta_plan"] = desc_plan
            fila["origen"] = "regla"
        elif codigo != "99999" and score >= 78:
            fila["cuenta_codigo"] = codigo
            fila["cuenta_plan"] = desc_plan or label
            fila["categoria"] = categoria if identificado else (label or categoria)
            fila["origen"] = "sugerido" if fuente not in {"conceptos_bancos", "regla_local"} else "regla"
        else:
            fila["cuenta_codigo"] = "99999"
            fila["cuenta_plan"] = desc_plan or label or categoria or "A clasificar"
            fila["origen"] = "a_clasificar"
        fila["origen"] = origen_linea_extracto(fila)
        fila["monto"] = round(cred - deb, 2)
        fila["_idx"] = i
        out.append(fila)
    out = _propagar_cuentas_repetidas(out)
    return _aplicar_mapa_clasif(out, mapa_clasif or {})


def _sin_filas_saldo(movs: list[dict]) -> list[dict]:
    """Saca saldo anterior / inicial / final y el 'nan' pegado al concepto."""
    out: list[dict] = []
    for m in movs:
        desc = re.sub(r"\s+\bnan\b", "", str(m.get("descripcion") or ""), flags=re.I)
        desc = re.sub(r"\s+", " ", desc).strip()
        blob = f"{desc} {m.get('categoria') or ''} {m.get('extracto_label') or ''}"
        if es_fila_saldo_bancario(blob):
            continue
        fila = dict(m)
        fila["descripcion"] = desc
        out.append(fila)
    return out


def _df_extracto(movs: list[dict], opciones: list[str]) -> pd.DataFrame:
    filas = []
    for i, m in enumerate(movs):
        filas.append(
            {
                "_i": int(m.get("_idx", i)),
                "Fecha": _fmt_fecha(m.get("fecha")),
                "Descripción": str(m.get("descripcion") or ""),
                "Débito": float(money(m.get("debito"))),
                "Crédito": float(money(m.get("credito"))),
                "Clasificación": str(m.get("categoria") or m.get("extracto_label") or ""),
                "Cuenta": _opcion_desde_codigo(str(m.get("cuenta_codigo") or "99999"), opciones),
                "Origen": _label_origen(str(m.get("origen") or "")),
            }
        )
    return pd.DataFrame(filas)


def _filas_componente(movs: list[dict]) -> list[dict]:
    filas = []
    for i, m in enumerate(_sin_filas_saldo(movs)):
        filas.append(
            {
                "i": int(m.get("_idx", i)),
                "fecha": _fmt_fecha(m.get("fecha")),
                "desc": str(m.get("descripcion") or ""),
                "deb": float(money(m.get("debito"))),
                "cred": float(money(m.get("credito"))),
                "clasif": str(m.get("categoria") or m.get("extracto_label") or ""),
                "codigo": str(m.get("cuenta_codigo") or "99999"),
                "origen": str(m.get("origen") or "a_clasificar"),
            }
        )
    return filas


def _cuentas_componente(opciones: list[str]) -> list[dict]:
    out = []
    vistos: set[str] = set()
    for op in opciones:
        cod = _codigo_desde_opcion(op)
        if not cod or cod in vistos:
            continue
        vistos.add(cod)
        out.append({"codigo": cod, "label": op})
    return out


def _aplicar_filas_componente(
    movs: list[dict],
    filas: list[dict],
    plan_df,
    mapa_clasif: dict[str, str] | None = None,
) -> list[dict]:
    col_cod = "codigo" if plan_df is not None and "codigo" in plan_df.columns else None
    col_desc = "descripcion" if plan_df is not None and "descripcion" in plan_df.columns else None
    desc_por_cod: dict[str, str] = {}
    if plan_df is not None and col_cod and col_desc:
        for _, row in plan_df.iterrows():
            desc_por_cod[str(row.get(col_cod) or "").strip()] = str(row.get(col_desc) or "").strip()
    by_i: dict[int, dict] = {}
    for f in filas or []:
        try:
            by_i[int(f.get("i"))] = f
        except (TypeError, ValueError):
            continue
    out = list(movs)
    for pos, m in enumerate(out):
        try:
            idx = int(m.get("_idx", pos))
        except (TypeError, ValueError):
            idx = pos
        f = by_i.get(idx) or by_i.get(pos)
        if not f:
            continue
        codigo = str(f.get("codigo") or "99999").strip() or "99999"
        clasif = str(f.get("clasif") or "").strip()
        origen = str(f.get("origen") or "")
        out[pos]["cuenta_codigo"] = codigo
        out[pos]["categoria"] = clasif or out[pos].get("categoria")
        if codigo != "99999" and codigo in desc_por_cod:
            out[pos]["cuenta_plan"] = desc_por_cod[codigo]
        elif clasif:
            out[pos]["cuenta_plan"] = clasif
        if origen in {"regla", "sugerido", "a_clasificar"}:
            out[pos]["origen"] = origen
        elif codigo != "99999":
            if str(out[pos].get("origen") or "") == "a_clasificar":
                out[pos]["origen"] = "sugerido"
        else:
            out[pos]["origen"] = "a_clasificar"
    return out


def _html_tabla_asiento(rows: list[dict]) -> str:
    if not rows:
        return "<p class='ce-sub'>Sin renglones.</p>"
    body = []
    for r in rows:
        cod = str(r.get("Código") or "")
        cls = "pend" if cod == "99999" else ""
        debe = float(r.get("Debe") or 0)
        haber = float(r.get("Haber") or 0)
        body.append(
            "<tr class='"
            + cls
            + "'><td>"
            + html.escape(cod)
            + "</td><td>"
            + html.escape(str(r.get("Descripción") or ""))
            + "</td><td class='num'>"
            + html.escape(_fmt_money(debe) if abs(debe) > 0.005 else "")
            + "</td><td class='num'>"
            + html.escape(_fmt_money(haber) if abs(haber) > 0.005 else "")
            + "</td></tr>"
        )
    return (
        "<div class='ce-wrap'><table class='ce-table'><thead>"
        "<tr><th>Código</th><th>Descripción</th><th class='num'>Debe</th>"
        "<th class='num'>Haber</th></tr></thead><tbody>"
        + "".join(body)
        + "</tbody></table></div>"
    )


def _aplicar_edicion(movs: list[dict], edited: pd.DataFrame, plan_df) -> list[dict]:
    if edited is None or edited.empty:
        return movs
    col_cod = "codigo" if plan_df is not None and "codigo" in plan_df.columns else None
    col_desc = "descripcion" if plan_df is not None and "descripcion" in plan_df.columns else None
    desc_por_cod: dict[str, str] = {}
    if plan_df is not None and col_cod and col_desc:
        for _, row in plan_df.iterrows():
            desc_por_cod[str(row.get(col_cod) or "").strip()] = str(row.get(col_desc) or "").strip()
    out = list(movs)
    for pos, (_, row) in enumerate(edited.iterrows()):
        try:
            idx = int(row.get("_i"))
        except (TypeError, ValueError):
            idx = pos
        if idx < 0 or idx >= len(out):
            if pos < len(out):
                idx = pos
            else:
                continue
        codigo = _codigo_desde_opcion(str(row.get("Cuenta") or ""))
        clasif = str(row.get("Clasificación") or "").strip()
        out[idx]["cuenta_codigo"] = codigo
        out[idx]["categoria"] = clasif or out[idx].get("categoria")
        if codigo != "99999" and codigo in desc_por_cod:
            out[idx]["cuenta_plan"] = desc_por_cod[codigo]
        elif clasif:
            out[idx]["cuenta_plan"] = clasif
        if codigo != "99999":
            if str(out[idx].get("origen") or "") == "a_clasificar":
                out[idx]["origen"] = "sugerido"
        else:
            out[idx]["origen"] = "a_clasificar"
    return out


def _asiento_desde_rows(
    rows: list[dict],
    *,
    banco: str,
    periodo: str,
    fecha_asiento: date,
) -> AsientoDevengamiento:
    renglones = [
        RenglonAsiento(
            codigo_cuenta=str(r.get("Código") or "99999"),
            descripcion_cuenta=str(r.get("Descripción") or ""),
            debe=float(r.get("Debe") or 0),
            haber=float(r.get("Haber") or 0),
        )
        for r in rows
        if abs(float(r.get("Debe") or 0)) > 0.005 or abs(float(r.get("Haber") or 0)) > 0.005
    ]
    asiento = AsientoDevengamiento(
        identificador=1,
        concepto=f"{banco} {periodo}".strip(),
        fecha=fecha_asiento,
        renglones=renglones,
    )
    asiento.tipo = banco  # type: ignore[attr-defined]
    asiento.periodo = periodo  # type: ignore[attr-defined]
    asiento.fecha_tango_str = fecha_asiento.strftime("%d/%m/%Y")  # type: ignore[attr-defined]
    return asiento


def _guardar_biblioteca_banco(
    sociedad_id: int,
    asientos: list,
    rows: list[dict],
    *,
    banco: str,
    periodo: str,
) -> str:
    etiqueta = str(periodo).replace("/", "-")
    existentes = st.session_state.get("biblioteca_bancos") or []
    if any(
        e.get("sociedad_id") == sociedad_id
        and e.get("banco") == banco
        and str(e.get("periodo") or "").replace("/", "-") == etiqueta
        for e in existentes
    ):
        raise ValueError(f"El período {etiqueta} ya está archivado para {banco}.")
    st.session_state.setdefault("biblioteca_bancos", []).append(
        {
            "sociedad_id": sociedad_id,
            "banco": banco,
            "periodo": periodo,
            "asientos": copy.deepcopy(asientos),
            "rows": copy.deepcopy(rows),
        }
    )
    st.session_state.setdefault("periodos_bancos_procesados", []).append(
        {"sociedad_id": sociedad_id, "banco": banco, "periodo": periodo}
    )
    guardar_biblioteca_persistida(
        biblioteca_asientos=st.session_state.get("biblioteca_asientos") or [],
        biblioteca_bancos=st.session_state.get("biblioteca_bancos") or [],
        periodos_procesados=st.session_state.get("periodos_procesados") or [],
        periodos_bancos_procesados=st.session_state.get("periodos_bancos_procesados") or [],
        usuario=str(st.session_state.get("oficina_usuario") or ""),
    )
    return etiqueta


def _paso_subir(
    *,
    sociedad_id: int,
    banco_elegido: str,
    nombre_activo: str | None,
    plan_vinculado: bool,
    periodo_key: str,
    preview_key: str,
    paso_key: str,
) -> None:
    st.markdown(
        '<div class="ce-top"><div><p class="ce-title">Extracto bancario</p>'
        "<p class='ce-sub'>Subí el PDF. La web lo lee y te muestra el extracto "
        "completo, movimiento por movimiento, con la imputación en la misma línea.</p></div></div>",
        unsafe_allow_html=True,
    )
    if not nombre_activo:
        st.error("Seleccioná el cliente arriba.")
        return
    if not plan_vinculado:
        st.error("Vinculá el plan de cuentas de esta sociedad antes de clasificar.")

    archivos = st.file_uploader(
        "PDF del extracto (Galicia, Macro, Santander, Provincia, BBVA, BNA, Mercado Pago)",
        type=["pdf", "xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key=f"ce_files_{sociedad_id}",
    )
    cache_files = f"ce_files_bytes_{sociedad_id}"
    if archivos:
        st.session_state[cache_files] = [(a.name, a.getvalue()) for a in archivos]
    periodo_ui = st.date_input(
        "Mes del extracto",
        value=st.session_state[periodo_key],
        key=f"ce_periodo_ui_{sociedad_id}",
    )
    st.session_state[periodo_key] = periodo_ui.replace(day=1)

    if st.button("Leer extracto", type="primary", key=f"ce_run_{sociedad_id}"):
        pares = [(a.name, a.getvalue()) for a in archivos] if archivos else list(
            st.session_state.get(cache_files) or []
        )
        if not pares:
            st.error("Subí el PDF o un Excel del extracto.")
            return
        if not plan_vinculado:
            st.error("Falta el plan de cuentas.")
            return
        fuentes = [_UploadMemoria(n, b) for n, b in pares]
        with st.spinner("Leyendo extracto…"):
            df, meta, errores = procesar_extractos_bancarios_pdfs(
                fuentes, banco_hint=banco_elegido
            )
        if errores:
            st.warning(
                "Algunos archivos tuvieron problemas: "
                + "; ".join(f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5])
            )
        if df is None or getattr(df, "empty", True):
            st.error("Este archivo no se pudo leer como extracto.")
            return
        filas = df_extracto_a_filas(df)
        filas = _sin_filas_saldo(filas)
        ok_saldo, msg_saldo = validar_saldos_corridos(filas)
        if not ok_saldo:
            st.warning(f"El saldo corrido no cierra del todo: {msg_saldo}. Igual se muestra para imputar.")
        banco = str((meta or {}).get("banco") or "") or banco_elegido or ""
        resultados = correr_motor(
            filas,
            db.listar_reglas_clasificacion(solo_activas=True),
            db.listar_proveedores_pendientes(sociedad_id, solo_libres=False),
            db.listar_veps_afip(sociedad_id),
            cliente_id=sociedad_id,
            banco=banco,
            periodo=st.session_state[periodo_key],
            saldo_ok=True,
        )
        movs = _enriquecer(
            resultados,
            _plan_df_sociedad(sociedad_id),
            mapa_clasif=_cargar_mapa_clasif(sociedad_id),
        )
        st.session_state[preview_key] = {
            "movimientos": movs,
            "banco": banco,
            "periodo": _periodo_mm_yyyy(st.session_state[periodo_key]),
            "periodo_date": st.session_state[periodo_key],
            "meta": meta or {},
        }
        st.session_state[paso_key] = "extracto"
        st.rerun()

    with st.expander("Instructivo Conceptos Bancos", expanded=False):
        try:
            info = cargar_instructivo()
            st.caption(
                f"{len(info.get('cuentas') or [])} cuentas · "
                f"{len(info.get('bancos') or {})} bancos · "
                f"`{info.get('origen') or CACHE_PATH}`"
            )
        except Exception as exc:
            st.warning(f"Sin instructivo: {exc}")
        up_cb = st.file_uploader(
            "Actualizar Conceptos Bancos 2.xlsx",
            type=["xlsx"],
            key=f"ce_uploader_conceptos_{sociedad_id}",
        )
        if up_cb is not None:
            data = cargar_instructivo(up_cb.getvalue())
            st.success(
                f"Cache actualizado: {len(data.get('cuentas') or [])} cuentas, "
                f"{len(data.get('bancos') or {})} bancos."
            )


def _paso_extracto(
    *,
    sociedad_id: int,
    preview_key: str,
    paso_key: str,
    nombre_activo: str | None,
) -> None:
    preview = st.session_state.get(preview_key) or {}
    movs = _sin_filas_saldo(list(preview.get("movimientos") or []))
    preview["movimientos"] = movs
    st.session_state[preview_key] = preview
    banco = str(preview.get("banco") or "")
    periodo = str(preview.get("periodo") or "")
    plan_df = _plan_df_sociedad(sociedad_id)
    opciones = _opciones_plan(plan_df)
    token_key = f"ce_grid_token_{sociedad_id}"
    token = int(st.session_state.get(token_key) or 0)
    n_plan = max(0, len(opciones) - 1)
    subtitulo = (
        f"{nombre_activo or ''} — {len(movs)} movimientos · {n_plan} cuentas del plan. "
        "Las que tienen regla quedan tomadas; el resto, sugeridas o a clasificar. "
        "Cambiá la clasificación en la línea: la cuenta no se mueve. "
        "El asiento engloba después por clasificación."
    )
    if n_plan <= 0:
        st.warning("No está el plan de cuentas de esta sociedad. Vinculalo y volvé a leer el extracto.")

    out = _EXTRACTO_GRID(
        titulo=f"Extracto {banco} · {periodo}",
        subtitulo=subtitulo,
        grid_id=f"ce_grid_{sociedad_id}_{token}_v3",
        filas=_filas_componente(movs),
        cuentas_json=json.dumps(_cuentas_componente(opciones), ensure_ascii=False),
        clasifs_json=json.dumps(_clasifs_componente(sociedad_id, movs), ensure_ascii=False),
        key=f"ce_grid_{sociedad_id}_{token}_v3",
        default={"action": "idle", "filas": []},
    )
    accion = str((out or {}).get("action") or "idle")
    if accion in {"siguiente", "volver"}:
        mapa = _cargar_mapa_clasif(sociedad_id)
        for item in (out or {}).get("clasifs") or []:
            nom = str((item or {}).get("nombre") or "").strip()
            cod = str((item or {}).get("codigo") or "").strip()
            if _clasif_tiene_cuenta(nom) and cod not in {"", "99999"}:
                mapa[nom] = cod
        actualizados = _aplicar_filas_componente(
            movs, (out or {}).get("filas") or [], plan_df, mapa_clasif=mapa
        )
        mapa = _mapa_desde_movimientos(actualizados, mapa)
        _guardar_mapa_clasif(sociedad_id, mapa)
        preview["movimientos"] = actualizados
        st.session_state[preview_key] = preview
        st.session_state[token_key] = token + 1
        if accion == "siguiente":
            st.session_state[paso_key] = "asiento"
        else:
            st.session_state.pop(preview_key, None)
            st.session_state[paso_key] = "subir"
        st.rerun()


def _paso_asiento(
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    preview_key: str,
    paso_key: str,
) -> None:
    preview = st.session_state.get(preview_key) or {}
    movs = _sin_filas_saldo(list(preview.get("movimientos") or []))
    preview["movimientos"] = movs
    st.session_state[preview_key] = preview
    banco = str(preview.get("banco") or banco_elegido or "Banco")
    periodo = str(preview.get("periodo") or _periodo_mm_yyyy(date.today()))
    plan_df = _plan_df_sociedad(sociedad_id)
    opciones = _opciones_plan(plan_df)
    sug_cod, sug_desc = cuenta_banco_del_plan(plan_df, banco)

    head_l, head_r = st.columns([4, 1])
    with head_l:
        st.markdown('<p class="ce-title">Asiento Tango</p>', unsafe_allow_html=True)
        st.caption("Así quedaría el asiento del mes. Si cierra, guardalo o bajá el Excel.")
    with head_r:
        if st.button("Volver", use_container_width=True, key=f"ce_back_{sociedad_id}"):
            st.session_state[paso_key] = "extracto"
            st.rerun()

    cuenta_banco_opt = st.selectbox(
        "Cuenta banco (contrapartida del asiento)",
        opciones,
        index=max(0, opciones.index(_opcion_desde_codigo(sug_cod, opciones)))
        if _opcion_desde_codigo(sug_cod, opciones) in opciones
        else 0,
        key=f"ce_cta_banco_{sociedad_id}",
    )
    codigo_banco = _codigo_desde_opcion(cuenta_banco_opt)
    desc_banco = cuenta_banco_opt.split("—", 1)[-1].strip() if "—" in cuenta_banco_opt else sug_desc

    periodo_date = preview.get("periodo_date")
    if isinstance(periodo_date, date):
        ultimo = calendar.monthrange(periodo_date.year, periodo_date.month)[1]
        fecha_asiento = date(periodo_date.year, periodo_date.month, ultimo)
    else:
        fecha_asiento = date.today()
    fecha_str = fecha_asiento.strftime("%d/%m/%Y")

    rows = renglones_asiento_banco_mes(
        movs,
        codigo_banco=codigo_banco,
        descripcion_banco=desc_banco or banco,
        periodo=periodo,
        fecha_str=fecha_str,
    )
    asiento = _asiento_desde_rows(rows, banco=banco, periodo=periodo, fecha_asiento=fecha_asiento)
    g = gate_asiento(rows)
    debe = sum(float(r.get("Debe") or 0) for r in rows)
    haber = sum(float(r.get("Haber") or 0) for r in rows)
    k1, k2, k3 = st.columns(3)
    k1.metric("Debe", _fmt_money(debe))
    k2.metric("Haber", _fmt_money(haber))
    k3.metric("Diferencia", _fmt_money(debe - haber))

    st.markdown(_html_tabla_asiento(rows), unsafe_allow_html=True)

    if g.get("bloqueantes"):
        st.error("Todavía no se puede guardar ni exportar:")
        for item in g["bloqueantes"]:
            st.markdown(f"- {item}")
        st.caption("Volvé al extracto y completá las líneas en 99999 / A clasificar.")
        return

    c_bib, c_xls = st.columns(2)
    with c_bib:
        if st.button(
            "Guardar en biblioteca",
            type="primary",
            use_container_width=True,
            key=f"ce_bib_{sociedad_id}",
        ):
            try:
                etiqueta = _guardar_biblioteca_banco(
                    sociedad_id, [asiento], rows, banco=banco, periodo=periodo,
                )
                db.insertar_movimientos_banco(movs)
                st.success(f"Asiento {banco} {etiqueta} guardado en la biblioteca.")
            except ValueError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"No se pudo guardar: {exc}")
    with c_xls:
        try:
            parts = [int(p) for p in periodo.replace("-", "/").split("/") if p]
            if len(parts) >= 2 and parts[0] > 12:
                anio, mes = parts[0], parts[1]
            elif len(parts) >= 2:
                mes, anio = parts[0], parts[1]
            else:
                mes, anio = fecha_asiento.month, fecha_asiento.year
            fp = (
                periodo,
                codigo_banco,
                tuple((r.get("Código"), r.get("Debe"), r.get("Haber")) for r in rows),
            )
            cache_fp = f"ce_xlsx_fp_{sociedad_id}"
            cache_bytes = f"ce_xlsx_bytes_{sociedad_id}"
            cache_name = f"ce_xlsx_name_{sociedad_id}"
            if st.session_state.get(cache_fp) != fp:
                ruta = generar_excel_tango_nativo(
                    [asiento],
                    nombre_activo or banco,
                    str(cuit_activo or ""),
                    mes,
                    anio,
                    plan_cuentas=plan_df,
                )
                st.session_state[cache_bytes] = ruta.read_bytes()
                st.session_state[cache_name] = ruta.name
                st.session_state[cache_fp] = fp
            st.download_button(
                "Excel para importar a Tango",
                data=st.session_state[cache_bytes],
                file_name=st.session_state.get(cache_name) or "asiento_banco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"ce_xlsx_{sociedad_id}",
            )
        except ExportacionTangoError as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(f"No se pudo armar el Excel Tango: {exc}")


def render_conciliacion_ae(
    *,
    sociedad_id: int,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
) -> None:
    st.markdown(_CSS, unsafe_allow_html=True)
    periodo_key = f"ae_periodo_{sociedad_id}"
    preview_key = f"ae_preview_{sociedad_id}"
    paso_key = f"ce_paso_{sociedad_id}"
    if periodo_key not in st.session_state:
        st.session_state[periodo_key] = date.today().replace(day=1)
    if paso_key not in st.session_state:
        st.session_state[paso_key] = "extracto" if st.session_state.get(preview_key) else "subir"

    paso = st.session_state.get(paso_key) or "subir"
    if paso == "extracto" and not st.session_state.get(preview_key):
        paso = "subir"
        st.session_state[paso_key] = "subir"

    if paso == "subir":
        _paso_subir(
            sociedad_id=sociedad_id,
            banco_elegido=banco_elegido,
            nombre_activo=nombre_activo,
            plan_vinculado=plan_vinculado,
            periodo_key=periodo_key,
            preview_key=preview_key,
            paso_key=paso_key,
        )
    elif paso == "extracto":
        _paso_extracto(
            sociedad_id=sociedad_id,
            preview_key=preview_key,
            paso_key=paso_key,
            nombre_activo=nombre_activo,
        )
    else:
        _paso_asiento(
            sociedad_id=sociedad_id,
            banco_elegido=banco_elegido,
            cuit_activo=cuit_activo,
            nombre_activo=nombre_activo,
            preview_key=preview_key,
            paso_key=paso_key,
        )
