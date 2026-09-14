"""UI Streamlit ARCA: solo encola jobs. Nunca claves. Badge de acceso por CUIT."""
from __future__ import annotations

import base64
import re
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

import streamlit as st

from afip_worker.auth import admin_mark_cuit_ready
from afip_worker.catalogo import catalogo_lookup, load_catalogo
from afip_worker.client import RemoteError, RemoteWorker
from afip_worker.jobs import (
    create_job,
    enqueue_job,
    list_jobs,
    normalizar_cuit,
    uploads_dir,
)
from afip_worker.registry import (
    badge_label,
    ensure_cuit_registered,
    list_cuits,
    mark_needs_admin,
)


_ACTION_LABELS = {
    "bajar_comprobantes": "Descargar Comprobantes en Línea",
    "bajar_portal_iva": "Descargar Portal IVA (compras / ventas)",
    "emitir_fcc": "Emitir facturas (FCC)",
    "bajar_veps": "Descargar VEPs",
}
_ACTIONS_UI = ("bajar_comprobantes", "bajar_portal_iva", "emitir_fcc", "bajar_veps")
_ACTIONS_LIVE = frozenset(_ACTIONS_UI)

_STATUS_BADGE = {
    "pending": "En cola",
    "running": "En curso",
    "needs_auth": "Falta 2FA",
    "done": "Listo",
    "error": "Error",
}

_HEALTH_TTL_SEC = 20.0
_RE_DIGITS = re.compile(r"\D")


def _secret_str(key: str) -> str:
    try:
        return str(st.secrets.get(key) or "").strip().strip('"').strip("'")
    except Exception:
        return ""


def _digits(cuit: str) -> str:
    return _RE_DIGITS.sub("", cuit or "")


def _parse_cuit_label(raw: str) -> tuple[str, str]:
    text = str(raw or "").strip()
    if not text or "—" not in text:
        return "", ""
    cuit, razon = text.split("—", 1)
    cuit = cuit.strip()
    razon = razon.strip()
    if razon == "(sin nombre)":
        razon = ""
    return cuit, razon


def _label_cuit(cuit: str, razon: str) -> str:
    return f"{cuit} — {razon or '(sin nombre)'}"


def _solicitado_por() -> str:
    return str(
        st.session_state.get("usuario_oficina_nombre")
        or st.session_state.get("usuario_oficina")
        or "oficina"
    ).strip()


def _carpeta_cliente(razon: str) -> str:
    t = re.sub(r'[\\/:*?"<>|]+', " ", razon or "").strip()
    return re.sub(r"\s+", " ", t)


def _ruta_sugerida(razon: str, hasta: date | None = None, action: str = "") -> str:
    mes = (hasta or date.today()).strftime("%m-%Y")
    slug = _carpeta_cliente(razon)
    if action == "bajar_portal_iva":
        if not slug:
            return rf"\\TANGOSRV\Compartido\CLIENTES\...\Impuestos\Portal IVA\{mes}"
        return rf"\\TANGOSRV\Compartido\CLIENTES\{slug}\Impuestos\Portal IVA\{mes}"
    if action == "bajar_veps":
        if not slug:
            return rf"\\TANGOSRV\Compartido\CLIENTES\...\Impuestos\VEPs\{mes}"
        return rf"\\TANGOSRV\Compartido\CLIENTES\{slug}\Impuestos\VEPs\{mes}"
    if not slug:
        return rf"\\TANGOSRV\Compartido\CLIENTES\...\Facturas\{mes}"
    return rf"\\TANGOSRV\Compartido\CLIENTES\{slug}\Facturas\{mes}"


def _fmt_dt(iso: str) -> str:
    raw = str(iso or "").strip()
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return raw


def _remote() -> RemoteWorker | None:
    url = _secret_str("AFIP_WORKER_URL")
    token = _secret_str("AFIP_WORKER_TOKEN")
    if url.startswith("http://127.0.0.1") or url.startswith("http://localhost"):
        return None
    if url and token:
        return RemoteWorker(url, token)
    return None


def _remote_health(remote: RemoteWorker) -> bool:
    now = time.monotonic()
    cached = st.session_state.get("_arca_health")
    ts = float(st.session_state.get("_arca_health_ts") or 0)
    if cached is not None and (now - ts) < _HEALTH_TTL_SEC:
        return bool(cached)
    ok = remote.health()
    st.session_state["_arca_health"] = ok
    st.session_state["_arca_health_ts"] = now
    return ok


def _lookup_acceso(cuit: str, rows: list[dict[str, Any]]) -> tuple[str, str]:
    """Solo lectura: no registra el CUIT al tipear."""
    digits = _digits(cuit)
    if len(digits) != 11:
        return "", ""
    for row in rows:
        if _digits(str(row.get("cuit") or "")) == digits:
            label = str(row.get("acceso") or badge_label(str(row.get("status") or "")))
            return label, str(row.get("note") or "")
    hit = catalogo_lookup(cuit)
    if hit:
        return str(hit.get("acceso") or "Listo"), str(hit.get("note") or "")
    return "Pedir acceso", "CUIT nuevo: hace falta 2FA una vez en la PC RECEPCION."


def _cuit_rows(remote: RemoteWorker | None) -> list[dict[str, Any]]:
    """Lista para la UI: worker si responde; si no, catálogo del repo (la nube no tiene jobs/)."""
    seed = load_catalogo()
    by_cuit = {str(r.get("cuit")): r for r in seed}
    if remote:
        try:
            for row in remote.list_cuits():
                cuit = str(row.get("cuit") or "")
                if cuit:
                    by_cuit[cuit] = row
        except RemoteError:
            pass
    elif not remote:
        for e in list_cuits():
            by_cuit[e.cuit] = {
                "cuit": e.cuit,
                "razon_social": e.razon_social,
                "status": e.status,
                "acceso": badge_label(e.status),
                "note": e.note,
                "updated_at": e.updated_at,
                "chrome_profile": e.chrome_profile,
            }
    return sorted(by_cuit.values(), key=lambda r: str(r.get("cuit") or ""))


def render_arca_module() -> None:
    """Módulo top-level ARCA: encolar + cola + registry (sin ejecutar AFIP)."""
    st.caption(
        "Cualquiera del estudio encola acá. AFIP lo navega RECEPCION sola "
        "(Chrome + clave ya guardada). Nadie más necesita permiso en AFIP. "
        "Las claves nunca van a Excel ni a esta web."
    )
    remote = _remote()
    if remote:
        if _remote_health(remote):
            st.success("Conectado al worker de RECEPCION.")
        else:
            st.error(
                "No se llega al worker. En RECEPCION dejá abierto "
                "`iniciar_afip_worker.bat` y, si el túnel cambió, actualizá "
                "`AFIP_WORKER_URL` en Secrets."
            )
    else:
        st.caption(
            "Cola en esta PC. En la nube hace falta `AFIP_WORKER_URL` en Secrets "
            "para que RECEPCION tome los trabajos."
        )

    tab_encolar, tab_cola, tab_cuits = st.tabs(["Encolar", "Cola", "CUITs / acceso"])

    with tab_encolar:
        _render_encolar(remote)
    with tab_cola:
        _render_cola(remote)
    with tab_cuits:
        _render_registry(remote)


def render_afip_cola_admin() -> None:
    """Alias retrocompatible."""
    render_arca_module()


def _on_cuit_catalogo() -> None:
    cuit, razon = _parse_cuit_label(str(st.session_state.get("arca_cuit_select") or ""))
    if cuit:
        st.session_state["arca_job_cuit"] = cuit
    if razon:
        st.session_state["arca_job_razon"] = razon


def _prefijar_desde_sociedad(conocidos: list[tuple[str, str]]) -> None:
    cuit = str(st.session_state.get("cuit_activo") or "").strip()
    razon = str(st.session_state.get("nombre_activo") or "").strip()
    sig = _digits(cuit)
    if st.session_state.get("_arca_prefijado_cuit") == sig:
        return
    st.session_state["_arca_prefijado_cuit"] = sig
    if not sig:
        return
    matched_c = ""
    catalog_r = ""
    for c, r in conocidos:
        if _digits(c) == sig:
            matched_c = c
            catalog_r = r
            break
    if matched_c:
        st.session_state["arca_cuit_select"] = _label_cuit(matched_c, catalog_r)
        st.session_state["arca_job_cuit"] = matched_c
    else:
        st.session_state["arca_cuit_select"] = ""
        st.session_state["arca_job_cuit"] = normalizar_cuit(cuit)
    nombre = catalog_r or razon
    if nombre:
        st.session_state["arca_job_razon"] = nombre


def _render_encolar(remote: RemoteWorker | None) -> None:
    rows = _cuit_rows(remote)
    conocidos = [
        (str(e.get("cuit") or ""), str(e.get("razon_social") or ""))
        for e in rows
    ]
    cuit_opts = [""] + [_label_cuit(c, r) for c, r in conocidos if c]
    if not st.session_state.get("_arca_default_live"):
        st.session_state["arca_job_action"] = "bajar_comprobantes"
        st.session_state["_arca_default_live"] = True
    _prefijar_desde_sociedad(conocidos)

    c1, c2 = st.columns(2)
    with c1:
        st.selectbox(
            "CUIT registrado",
            options=cuit_opts,
            key="arca_cuit_select",
            on_change=_on_cuit_catalogo,
            help="Si no está, escribí el CUIT abajo. Sale de la sociedad activa.",
        )
        cuit = st.text_input("CUIT", placeholder="27-42043034-0", key="arca_job_cuit")
        razon = st.text_input(
            "Razón social",
            placeholder="Nombre del cliente",
            key="arca_job_razon",
        )
    with c2:
        action = st.selectbox(
            "Acción",
            options=list(_ACTIONS_UI),
            format_func=lambda a: _ACTION_LABELS.get(a, a),
            key="arca_job_action",
        )
        requested_by = _solicitado_por()
        st.caption(f"Lo pide **{requested_by}**.")

    label, note = _lookup_acceso(cuit, rows)
    if label == "Listo":
        st.success(f"Acceso: **{label}** — {note}")
    elif label == "Pedir acceso":
        st.info(
            f"Acceso: **primera vez** — RECEPCION entra sola a AFIP. "
            "Nadie del estudio necesita permiso fiscal. "
            "Solo se frena si AFIP pide 2FA (lo ves en Cola)."
        )
    elif label:
        st.info(f"Acceso: **{label}** — {note}")

    params: dict[str, Any] = {}
    plantilla_b64 = ""
    plantilla_name = ""
    hasta_ruta = date.today()
    if action == "emitir_fcc":
        plantilla = st.file_uploader(
            "Plantilla Excel (montos / receptor / concepto — SIN claves)",
            type=["xlsx", "xls"],
            key="arca_job_xlsx",
        )
        fecha_emision = st.date_input(
            "Fecha emisión",
            value=date.today(),
            format="DD/MM/YYYY",
            key="arca_job_emision",
        )
        params["fecha_emision"] = fecha_emision.isoformat()
        hasta_ruta = fecha_emision
        if plantilla is not None:
            raw = plantilla.getvalue()
            if remote:
                plantilla_b64 = base64.b64encode(raw).decode("ascii")
                plantilla_name = plantilla.name
                st.caption(f"Plantilla se envía al worker: `{plantilla.name}`")
            else:
                dest = uploads_dir() / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}_{plantilla.name}"
                dest.write_bytes(raw)
                params["plantilla_excel"] = str(dest.relative_to(Path(__file__).resolve().parents[1]))
                st.caption(f"Plantilla guardada: `{params['plantilla_excel']}`")
    else:
        d1, d2 = st.columns(2)
        with d1:
            desde = st.date_input(
                "Período desde",
                value=date.today().replace(day=1),
                format="DD/MM/YYYY",
                key="arca_job_desde",
            )
        with d2:
            hasta = st.date_input(
                "Período hasta",
                value=date.today(),
                format="DD/MM/YYYY",
                key="arca_job_hasta",
            )
        params["periodo_desde"] = desde.isoformat()
        params["periodo_hasta"] = hasta.isoformat()
        hasta_ruta = hasta
        if action == "bajar_comprobantes":
            params["analizar_monotributo"] = st.checkbox(
                "Al terminar, analizar recategorización (período facturado, recibos, NC)",
                value=True,
                key="arca_job_analizar_mono",
                help="El worker arma el papel de trabajo de monotributo con los PDF bajados.",
            )
        elif action == "bajar_portal_iva":
            st.caption(
                "Baja CSV/PDF de compras y ventas (Portal IVA; si hace falta, Mis Comprobantes) "
                r"a Impuestos\Portal IVA\MM-YYYY. No presenta la DDJJ."
            )
        elif action == "bajar_veps":
            st.caption(
                r"Baja los VEPs del período a Impuestos\VEPs\MM-YYYY."
            )

    sugerida = _ruta_sugerida(razon, hasta_ruta, action)
    ruta = st.text_input(
        "Ruta destino (UNC)",
        placeholder=sugerida,
        key="arca_job_ruta",
        help="Carpeta de red donde RECEPCION guarda los PDF. Si lo dejás vacío, usa la sugerida.",
    )
    params["ruta_destino"] = (ruta.strip() or sugerida).strip().strip('"').strip("'")
    if not ruta.strip():
        st.caption(f"Si encolás ahora, guarda en `{sugerida}`")

    if st.button("Encolar", type="primary", key="arca_job_enqueue"):
        if len(_digits(cuit)) != 11:
            st.error("El CUIT tiene que tener 11 dígitos.")
        elif not razon.strip():
            st.error("Falta la razón social.")
        elif action == "emitir_fcc" and not (params.get("plantilla_excel") or plantilla_b64):
            st.error("Subí la plantilla Excel para emitir FCC.")
        else:
            try:
                if remote:
                    job = remote.enqueue(
                        cuit=cuit,
                        razon_social=razon,
                        action=action,
                        params=params,
                        requested_by=requested_by,
                        plantilla_b64=plantilla_b64,
                        plantilla_name=plantilla_name,
                    )
                    jid = str(job.get("id") or "")
                    st.success(f"Encolado en RECEPCION: {jid}")
                    st.caption(
                        "Lo toma RECEPCION sola. El resto del estudio no abre AFIP ni necesita "
                        "permiso fiscal. Si AFIP pide 2FA, aparece en Cola."
                    )
                else:
                    ensure_cuit_registered(cuit, razon)
                    job_obj = create_job(
                        cuit=cuit,
                        razon_social=razon,
                        action=action,  # type: ignore[arg-type]
                        params=params,
                        requested_by=requested_by,
                    )
                    enqueue_job(job_obj)
                    st.success(f"Encolado: {job_obj.id}")
                    st.caption(
                        "En esta PC queda en la cola local. En la nube, RECEPCION lo toma sola. "
                        "Si AFIP pide 2FA, aparece en Cola."
                    )
            except (ValueError, RemoteError) as exc:
                st.error(str(exc))


def _render_cola(remote: RemoteWorker | None) -> None:
    if st.button("Actualizar cola", key="arca_cola_refresh"):
        st.session_state.pop("_arca_health", None)
        st.rerun()
    rows_src: list[dict[str, Any]] = []
    try:
        if remote:
            rows_src = remote.list_jobs()
        else:
            rows_src = [j.to_dict() for j in list_jobs()]
    except RemoteError as exc:
        st.error(str(exc))
        return
    if not rows_src:
        st.info("Sin trabajos todavía.")
        return
    rows = []
    for j in reversed(rows_src):
        result = j.get("result") or {}
        rows.append(
            {
                "Estado": _STATUS_BADGE.get(str(j.get("status")), j.get("status")),
                "CUIT": j.get("cuit"),
                "Cliente": j.get("razon_social"),
                "Acción": _ACTION_LABELS.get(str(j.get("action")), j.get("action")),
                "Lo pidió": j.get("requested_by"),
                "Creado": _fmt_dt(str(j.get("created_at") or "")),
                "Mensaje": str(result.get("message") or "")[:140] if isinstance(result, dict) else "",
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)

    done_mono = [
        j for j in rows_src
        if j.get("status") == "done"
        and str(j.get("action")) == "bajar_comprobantes"
        and isinstance(j.get("result"), dict)
        and (j.get("result") or {}).get("extra", {}).get("monotributo")
    ]
    if done_mono:
        st.markdown("##### Recategorización lista (jobs ARCA)")
        for j in done_mono[:8]:
            info = ((j.get("result") or {}).get("extra") or {}).get("monotributo") or {}
            xlsx = info.get("xlsx") or ""
            st.info(
                f"**{j.get('razon_social')}** ({j.get('cuit')}): "
                f"{info.get('cantidad', 0)} comprobante(s) · neto ${float(info.get('neto') or 0):,.2f}"
                + (f" · Excel `{xlsx}`" if xlsx else "")
            )
            errs = info.get("errores") or []
            if errs:
                with st.expander(f"Advertencias {j.get('id')}", expanded=False):
                    st.dataframe(errs, use_container_width=True, hide_index=True)

    needs = [j for j in rows_src if j.get("status") == "needs_auth"]
    if needs:
        st.warning(
            f"**{len(needs)} trabajo(s) esperan 2FA.** "
            "En RECEPCION abrí AFIP una vez y marcá el CUIT como Listo."
        )
        for j in needs:
            auth = j.get("auth") or {}
            result = j.get("result") or {}
            note = ""
            if isinstance(auth, dict):
                note = str(auth.get("note") or "")
            if isinstance(result, dict) and not note:
                note = str(result.get("message") or "")
            st.code(f"{j.get('id')}\nCUIT {j.get('cuit')}\n{note}")


def _render_registry(remote: RemoteWorker | None) -> None:
    st.markdown("##### Registry de CUITs")
    st.caption("Solo estado de acceso. Sin contraseñas ni tokens.")
    entries = _cuit_rows(remote)
    if entries:
        st.dataframe(
            [
                {
                    "CUIT": e.get("cuit"),
                    "Razón social": e.get("razon_social"),
                    "Acceso": e.get("acceso") or badge_label(str(e.get("status") or "")),
                    "Nota": e.get("note"),
                    "Actualizado": _fmt_dt(str(e.get("updated_at") or "")),
                }
                for e in entries
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Todavía no hay CUITs registrados.")

    st.divider()
    st.markdown("##### Handoff admin (después del 2FA)")
    hc1, hc2 = st.columns(2)
    with hc1:
        cuit_h = st.text_input("CUIT a marcar", key="arca_hand_cuit")
        razon_h = st.text_input("Razón social", key="arca_hand_razon")
    with hc2:
        if st.button("Marcar Listo (Chrome autofill OK)", type="primary", key="arca_hand_ready"):
            if not cuit_h.strip():
                st.error("Indicá el CUIT.")
            else:
                try:
                    if remote:
                        remote.mark_ready(cuit_h, razon_h)
                    else:
                        admin_mark_cuit_ready(cuit_h, razon_h)
                    st.success(f"{cuit_h} → Listo")
                    st.rerun()
                except RemoteError as exc:
                    st.error(str(exc))
        if st.button("Marcar Pedir acceso", key="arca_hand_need"):
            if not cuit_h.strip():
                st.error("Indicá el CUIT.")
            else:
                try:
                    if remote:
                        remote.mark_needs_admin(cuit_h, razon_h)
                    else:
                        mark_needs_admin(cuit_h, razon_social=razon_h)
                    st.warning(f"{cuit_h} → Pedir acceso")
                    st.rerun()
                except RemoteError as exc:
                    st.error(str(exc))
