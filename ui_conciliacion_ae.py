# -*- coding: utf-8 -*-
"""Conciliación Bancaria: UI de AE_Studio.html, reglas y persistencia de la web."""
from __future__ import annotations

import base64
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import database as db
from capa_revision import resolver_codigo_plan
from motor_conciliacion import (
    CATEGORIA_A_CUENTA_HINT,
    bucket_ae,
    correr_motor,
    df_extracto_a_filas,
    money,
    validar_saldos_corridos,
)
from procesador import clasificar_movimiento_extracto, procesar_extractos_bancarios_pdfs

_COMPONENT_DIR = Path(__file__).resolve().parent / "ae_conciliacion_component"
_ae_ui = components.declare_component("ae_conciliacion", path=str(_COMPONENT_DIR))

_SUB_AE = (
    ("impuestos a los débitos y créditos", "ICDB Ley 25.413"),
    ("impuesto ley 25.413", "ICDB Ley 25.413"),
    ("pago de haberes", "Haberes / Sueldo"),
    ("pago haberes", "Haberes / Sueldo"),
    ("honorario", "Honorarios profesionales"),
    ("transferencia de tercero", "Transferencia de tercero"),
    ("transferencias recibidas", "Transferencia de tercero"),
    ("transferencia a tercero", "Transferencia a tercero"),
    ("transferencias emitidas", "Transferencia a tercero"),
    ("rescate fima", "FCI - Rescate"),
    ("rescate fci", "FCI - Rescate"),
    ("suscripción fci", "FCI - Suscripcion"),
    ("suscripcion fci", "FCI - Suscripcion"),
    ("percepción iva", "Percepcion IVA"),
    ("percepcion iva", "Percepcion IVA"),
    ("iibb", "Ret. IIBB - ARBA"),
    ("pago arba", "Ret. IIBB - ARBA"),
    ("sircreb", "Ret. IIBB - ARBA"),
    ("pagos tarjeta", "Pago tarjeta"),
    ("compras", "Compra debito"),
    ("compra con tarjeta", "Compra debito"),
    ("pagos afip", "Pago AFIP"),
    ("intereses", "Intereses"),
    ("autonom", "Aportes autonomos"),
    ("monotributo", "Monotributo"),
    ("prepaga", "Obra social / Prepaga"),
    ("obra social", "Obra social / Prepaga"),
    ("usd", "Compra/venta USD"),
)


class _Up:
    def __init__(self, name: str, data: bytes) -> None:
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def _periodo_str(d: date | None) -> str:
    if not d:
        return date.today().strftime("%Y-%m-01")
    return d.replace(day=1).isoformat()


def _sub_ae(label: str, desc: str) -> str:
    blob = f"{label} {desc}".lower()
    compacto = blob.replace(" ", "")
    if "25413" in compacto or "icdb" in blob:
        return "ICDB Ley 25.413"
    for src, dst in _SUB_AE:
        if src in blob:
            return dst
    return (label or desc or "-")[:80] or "-"


def _gravado(sub: str, vista: str) -> str:
    if sub in {"Intereses"} or "exento" in sub.lower():
        return "Exento"
    if vista == "retencion" and sub == "ICDB Ley 25.413":
        return "No computable"
    if vista == "ingreso":
        return "Gravado"
    return ""


def _enriquecer(movs: list[dict]) -> list[dict]:
    out = []
    for m in movs:
        cred = float(money(m.get("credito")))
        deb = float(money(m.get("debito")))
        label = clasificar_movimiento_extracto(
            str(m.get("descripcion") or ""),
            importe=cred - deb,
        )
        fila = dict(m)
        fila["extracto_label"] = label
        fila["vista"] = bucket_ae(m, extracto_label=label)
        fila["sub_categoria"] = _sub_ae(label, str(m.get("descripcion") or ""))
        if fila["sub_categoria"] in {"-", ""}:
            fila["sub_categoria"] = str(
                m.get("concepto_instructivo") or m.get("categoria") or "-"
            )
        fila["monto"] = round(cred - deb, 2)
        fila["gravado"] = _gravado(str(fila["sub_categoria"]), str(fila["vista"]))
        out.append(fila)
    return out


def _aplicar_plan(movs: list[dict], plan_df) -> list[dict]:
    out = []
    for m in movs:
        fila = dict(m)
        codigo, desc_plan, score = resolver_codigo_plan(
            str(m.get("categoria") or ""),
            plan_df,
            hints=CATEGORIA_A_CUENTA_HINT,
        )
        fila["cuenta_codigo"] = codigo
        fila["cuenta_plan"] = desc_plan
        fila["score_plan"] = score
        out.append(fila)
    return out


def _fila_ae(m: dict) -> dict:
    cred = float(money(m.get("credito")))
    deb = float(money(m.get("debito")))
    monto = float(m.get("monto") if m.get("monto") is not None else cred - deb)
    vista = m.get("vista") or bucket_ae(m, extracto_label=str(m.get("extracto_label") or ""))
    fecha = str(m.get("fecha") or "")
    return {
        "id": m.get("id"),
        "fecha": fecha,
        "fecha_texto": fecha,
        "descripcion": str(m.get("descripcion") or ""),
        "monto": monto,
        "credito": cred,
        "debito": deb,
        "categoria": vista,
        "sub_categoria": str(m.get("sub_categoria") or "-"),
        "cuenta_codigo": str(m.get("cuenta_codigo") or "") or None,
        "gravado": str(m.get("gravado") or ""),
        "periodo": str(m.get("periodo") or ""),
        "banco": str(m.get("banco") or ""),
        "estado": str(m.get("estado") or ""),
        "cuit_vinculado": "",
        "observacion": str(m.get("match_detalle") or ""),
    }


def _plan_ae(plan_df) -> list[dict]:
    if plan_df is None or getattr(plan_df, "empty", True):
        return []
    out = []
    cols = {str(c).lower(): c for c in plan_df.columns}
    c_cod = cols.get("codigo") or cols.get("código") or cols.get("cuenta")
    c_desc = cols.get("descripcion") or cols.get("descripción") or cols.get("nombre")
    if not c_cod:
        return []
    for _, row in plan_df.head(4000).iterrows():
        codigo = str(row.get(c_cod) or "").strip()
        if not codigo:
            continue
        out.append(
            {
                "codigo": codigo,
                "descripcion": str(row.get(c_desc) or "") if c_desc else "",
            }
        )
    return out


def _clientes_ae(clientes: list[dict]) -> list[dict]:
    return [
        {
            "id": str(c.get("id")),
            "nombre": str(c.get("nombre") or c.get("razon_social") or ""),
            "cuit": str(c.get("cuit") or ""),
            "tipo": str(c.get("tipo_persona") or ""),
        }
        for c in clientes
    ]


def _parse_texto(texto: str, banco: str, periodo: str) -> pd.DataFrame:
    filas = []
    fre = re.compile(r"^(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})")
    for raw in texto.splitlines():
        line = raw.strip()
        if len(line) < 4:
            continue
        m = fre.match(line)
        if not m:
            continue
        rest = line[len(m.group(0)) :].strip()
        tokens = re.split(r"\t|  {2,}", rest)
        nums: list[float] = []
        desc: list[str] = []
        for t in tokens:
            tc = t.strip().replace(".", "").replace(",", ".").rstrip("-")
            if re.fullmatch(r"-?[\d]+\.?\d*", tc) and len(tc) > 1:
                nums.append(float(tc))
            elif t.strip():
                desc.append(t.strip())
        cred = 0.0
        deb = 0.0
        if len(nums) >= 2:
            v = nums[-2]
            if v >= 0:
                cred = v
            else:
                deb = abs(v)
        elif len(nums) == 1:
            if nums[0] >= 0:
                cred = nums[0]
            else:
                deb = abs(nums[0])
        filas.append(
            {
                "Fecha": m.group(1),
                "Descripcion": " ".join(desc)[:200],
                "Credito": cred,
                "Debito": deb,
                "Banco": banco,
            }
        )
    return pd.DataFrame(filas)


def _correr_import(
    *,
    sociedad_id: int,
    banco_elegido: str,
    archivos,
    periodo: str,
) -> tuple[list[dict], dict]:
    df, meta, errores = procesar_extractos_bancarios_pdfs(archivos)
    aviso = ""
    if errores:
        aviso = "; ".join(
            f"{e.get('archivo')}: {e.get('motivo')}" for e in errores[:5]
        )
    if df is None or getattr(df, "empty", True):
        return [], {"error": aviso or "Este archivo no se pudo leer.", "meta": meta or {}}
    filas = df_extracto_a_filas(df)
    ok_saldo, msg_saldo = validar_saldos_corridos(filas)
    banco = str((meta or {}).get("banco") or "") or banco_elegido or ""
    resultados = correr_motor(
        filas,
        db.listar_reglas_clasificacion(solo_activas=True),
        db.listar_proveedores_pendientes(sociedad_id, solo_libres=False),
        db.listar_veps_afip(sociedad_id),
        cliente_id=sociedad_id,
        banco=banco,
        periodo=periodo,
        saldo_ok=ok_saldo,
    )
    movs = _aplicar_plan(_enriquecer(resultados), st.session_state.get("plan_cuentas_df"))
    return movs, {
        "banco": banco,
        "periodo": periodo,
        "aviso": aviso,
        "saldo": msg_saldo if not ok_saldo else "",
        "meta": meta or {},
    }


def _guardar(sociedad_id: int, preview: dict) -> int:
    movs = preview.get("movimientos") or []
    periodo = str(preview.get("periodo") or _periodo_str(date.today()))
    db.borrar_movimientos_periodo(sociedad_id, periodo=periodo, banco=None)
    n = db.insertar_movimientos_banco(movs)
    db.registrar_auditoria_conciliacion(
        cliente_id=sociedad_id,
        movimiento_id=None,
        usuario=str(st.session_state.get("oficina_usuario") or "sistema"),
        accion="guardar_extracto_ae",
        detalle=f"{n} movimientos | banco={preview.get('banco')}",
    )
    st.session_state["motor_last_periodo"] = periodo
    return n


def render_conciliacion_ae(
    *,
    sociedad_id: int | None,
    banco_elegido: str,
    cuit_activo: str | None,
    nombre_activo: str | None,
    plan_vinculado: bool,
    clientes: list[dict] | None = None,
) -> None:
    clientes = clientes or db.listar_clientes()
    clientes_pj = [c for c in clientes if c.get("tipo_persona") == "Persona Jurídica"]
    preview_key = f"ae_preview_{sociedad_id or 0}"
    view_key = f"ae_view_{sociedad_id or 0}"
    periodo = st.session_state.get("motor_last_periodo") or _periodo_str(date.today())

    preview = st.session_state.get(preview_key) or {}
    if preview.get("movimientos"):
        movs = preview["movimientos"]
    elif sociedad_id:
        guardados = db.listar_movimientos_banco(int(sociedad_id), periodo=periodo)
        movs = _aplicar_plan(_enriquecer(guardados), st.session_state.get("plan_cuentas_df"))
    else:
        movs = []

    banco_filtro = str(banco_elegido or "").strip()
    if banco_filtro:
        filtrados = [
            m
            for m in movs
            if not m.get("banco") or banco_filtro.lower() in str(m.get("banco") or "").lower()
        ]
        if filtrados:
            movs = filtrados

    cliente = None
    if sociedad_id:
        cliente = {
            "id": str(sociedad_id),
            "nombre": nombre_activo or "",
            "cuit": cuit_activo or "",
        }

    provs = []
    if sociedad_id:
        for p in db.listar_proveedores_pendientes(int(sociedad_id), solo_libres=False):
            provs.append(
                {
                    "id": p.get("id"),
                    "cuit": str(p.get("cuit") or p.get("num_comp") or ""),
                    "razon_social": str(p.get("razon_social") or ""),
                    "categoria_contable": "PROVEEDORES",
                    "rubro": str(p.get("tipo_comp") or ""),
                    "es_deducible": False,
                    "activo": True,
                }
            )

    usuarios = []
    try:
        for u in db.listar_usuarios_oficina(solo_activos=False):
            usuarios.append(
                {
                    "usuario": u.get("usuario"),
                    "nombre": u.get("nombre"),
                    "es_admin": bool(u.get("es_admin")),
                    "activo": bool(u.get("activo", True)),
                }
            )
    except Exception:
        usuarios = []

    event = _ae_ui(
        movs=[_fila_ae(m) for m in movs],
        clientes=_clientes_ae(clientes_pj or clientes),
        cliente=cliente,
        banco=banco_elegido or "",
        preview=[_fila_ae(m) for m in (preview.get("movimientos") or [])],
        plan=_plan_ae(st.session_state.get("plan_cuentas_df")),
        proveedores=provs,
        empleados=[],
        usuarios=usuarios,
        usuario=str(st.session_state.get("oficina_usuario") or "Estudio"),
        view=st.session_state.get(view_key) or "dashboard",
        notify=st.session_state.pop("ae_notify", "") or "",
        key=f"ae_ui_{sociedad_id or 0}",
        default=None,
    )

    if not event or not isinstance(event, dict):
        return

    action = str(event.get("action") or "")
    if action == "pick_client":
        try:
            st.session_state["sociedad_activa"] = int(event.get("id"))
        except (TypeError, ValueError):
            pass
        st.rerun()
        return
    if action == "set_banco":
        st.session_state["selector_banco_conciliar_v2"] = str(event.get("banco") or "")
        st.rerun()
        return
    if action == "cancel_preview":
        st.session_state.pop(preview_key, None)
        st.rerun()
        return
    if action == "clear" and sociedad_id:
        db.borrar_movimientos_periodo(int(sociedad_id), periodo=None, banco=None)
        st.session_state.pop(preview_key, None)
        st.session_state["ae_notify"] = "Movimientos eliminados"
        st.rerun()
        return
    if action == "save" and sociedad_id and preview.get("movimientos"):
        n = _guardar(int(sociedad_id), preview)
        st.session_state.pop(preview_key, None)
        st.session_state[view_key] = "dashboard"
        st.session_state["ae_notify"] = f"{n} movimientos guardados"
        st.rerun()
        return
    if action == "reclass" and event.get("id"):
        try:
            db.actualizar_movimiento_banco(
                int(event["id"]),
                categoria=str(event.get("sub_categoria") or event.get("categoria") or ""),
                tipo=str(event.get("categoria") or ""),
            )
        except Exception:
            pass
        st.session_state["ae_notify"] = "Reclasificado"
        st.rerun()
        return
    if action == "import":
        if not sociedad_id:
            st.session_state["ae_notify"] = "Selecciona el cliente primero"
            st.rerun()
            return
        kind = str(event.get("kind") or "pdf")
        with st.spinner("Leyendo extracto (OCR si es escaneo)…"):
            if kind == "text":
                df = _parse_texto(
                    str(event.get("text") or ""),
                    banco_elegido or "Otro",
                    periodo,
                )
                if df.empty:
                    st.session_state["ae_notify"] = "No se detectaron movimientos"
                    st.rerun()
                    return
                filas = df_extracto_a_filas(df)
                ok_saldo, _msg = validar_saldos_corridos(filas)
                resultados = correr_motor(
                    filas,
                    db.listar_reglas_clasificacion(solo_activas=True),
                    db.listar_proveedores_pendientes(int(sociedad_id), solo_libres=False),
                    db.listar_veps_afip(int(sociedad_id)),
                    cliente_id=int(sociedad_id),
                    banco=banco_elegido or "",
                    periodo=periodo,
                    saldo_ok=ok_saldo,
                )
                movs_imp = _aplicar_plan(
                    _enriquecer(resultados), st.session_state.get("plan_cuentas_df")
                )
                meta = {"banco": banco_elegido or ""}
            else:
                raw = event.get("b64") or ""
                try:
                    data = base64.b64decode(raw)
                except Exception:
                    st.session_state["ae_notify"] = "No se pudo leer el archivo"
                    st.rerun()
                    return
                nombre = str(event.get("name") or ("extracto.pdf" if kind == "pdf" else "extracto.xlsx"))
                movs_imp, meta = _correr_import(
                    sociedad_id=int(sociedad_id),
                    banco_elegido=banco_elegido,
                    archivos=[_Up(nombre, data)],
                    periodo=periodo,
                )
                if meta.get("error"):
                    st.session_state["ae_notify"] = str(meta["error"])
                    st.session_state[view_key] = "import"
                    st.rerun()
                    return
        st.session_state[preview_key] = {
            "movimientos": movs_imp,
            "banco": meta.get("banco") or banco_elegido,
            "periodo": periodo,
        }
        st.session_state[view_key] = "import"
        n_ing = sum(1 for m in movs_imp if m.get("vista") == "ingreso")
        st.session_state["ae_notify"] = f"{len(movs_imp)} movimientos ({n_ing} ingresos)"
        st.rerun()
        return
