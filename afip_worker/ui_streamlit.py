"""UI Streamlit ARCA: solo encola jobs. Nunca claves. Badge de acceso por CUIT."""
from __future__ import annotations

import base64
from datetime import date, datetime
from pathlib import Path
from typing import Any

import streamlit as st

from afip_worker.auth import admin_mark_cuit_ready
from afip_worker.catalogo import catalogo_lookup, load_catalogo
from afip_worker.client import RemoteError, RemoteWorker
from afip_worker.jobs import (
    ACTIONS,
    create_job,
    enqueue_job,
    jobs_root,
    list_jobs,
    uploads_dir,
)
from afip_worker.registry import (
    badge_label,
    ensure_cuit_registered,
    list_cuits,
    mark_needs_admin,
)


_ACTION_LABELS = {
    "emitir_fcc": "Emitir facturas (FCC)",
    "bajar_veps": "Descargar VEPs",
    "bajar_comprobantes": "Descargar Comprobantes en Línea",
}

_STATUS_BADGE = {
    "pending": "pending",
    "running": "running",
    "needs_auth": "needs_auth",
    "done": "done",
    "error": "error",
}


def _secret_str(key: str) -> str:
    try:
        return str(st.secrets.get(key) or "").strip().strip('"').strip("'")
    except Exception:
        return ""


def _remote() -> RemoteWorker | None:
    url = _secret_str("AFIP_WORKER_URL")
    token = _secret_str("AFIP_WORKER_TOKEN")
    if url.startswith("http://127.0.0.1") or url.startswith("http://localhost"):
        return None
    if url and token:
        return RemoteWorker(url, token)
    return None


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


def _badge_acceso(cuit: str, razon: str, remote: RemoteWorker | None) -> tuple[str, str]:
    if remote:
        try:
            entry = remote.ensure_cuit(cuit, razon)
            label = str(entry.get("acceso") or badge_label(str(entry.get("status") or "")))
            return label, str(entry.get("note") or "")
        except RemoteError:
            pass
    hit = catalogo_lookup(cuit)
    if hit:
        return str(hit.get("acceso") or "Listo"), str(hit.get("note") or "")
    if not remote:
        return _badge_acceso_local(cuit, razon)
    return "Pedir acceso", "Sin confirmación del worker; si es CUIT nuevo hace falta 2FA en RECEPCION."


def _badge_acceso_local(cuit: str, razon: str = "") -> tuple[str, str]:
    entry = ensure_cuit_registered(cuit, razon)
    return badge_label(entry.status), entry.note


def render_arca_module() -> None:
    """Módulo top-level ARCA: encolar + cola + registry (sin ejecutar AFIP)."""
    st.caption(
        "Esta web es para **todo el estudio**: cualquiera encola desde acá. "
        "AFIP se ejecuta solo en la PC RECEPCION (Chrome + autofill). "
        "Las claves nunca van a Excel ni a esta web."
    )
    remote = _remote()
    if remote:
        if remote.health():
            st.success("Conectado al worker de RECEPCION (nube → túnel → tu PC).")
        else:
            st.error(
                "No se llega al worker. En la PC RECEPCION dejá abierto "
                "`iniciar_afip_worker.bat` y actualizá `AFIP_WORKER_URL` en Secrets "
                "si el túnel cambió de URL."
            )
        st.caption(f"Cola remota: `{st.secrets.get('AFIP_WORKER_URL')}`")
    else:
        root = jobs_root()
        st.caption(f"Cola local: `{root}` — para la web en la nube configurá AFIP_WORKER_URL en Secrets.")

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


def _render_encolar(remote: RemoteWorker | None) -> None:
    conocidos = [
        (str(e.get("cuit") or ""), str(e.get("razon_social") or ""))
        for e in _cuit_rows(remote)
    ]
    cuit_opts = [""] + [f"{c} — {r or '(sin nombre)'}" for c, r in conocidos if c]

    c1, c2 = st.columns(2)
    with c1:
        elegido = st.selectbox(
            "CUIT registrado",
            options=cuit_opts,
            key="arca_cuit_select",
            help="O escribí un CUIT nuevo abajo.",
        )
        cuit_default = ""
        razon_default = ""
        if elegido and "—" in elegido:
            cuit_default = elegido.split("—", 1)[0].strip()
            razon_default = elegido.split("—", 1)[1].strip()
            if razon_default == "(sin nombre)":
                razon_default = ""
        cuit = st.text_input("CUIT", value=cuit_default, placeholder="27-42043034-0", key="arca_job_cuit")
        razon = st.text_input(
            "Razón social",
            value=razon_default,
            placeholder="Camila Rocio Albarello",
            key="arca_job_razon",
        )
    with c2:
        action = st.selectbox(
            "Acción",
            options=list(ACTIONS),
            format_func=lambda a: _ACTION_LABELS.get(a, a),
            key="arca_job_action",
        )
        requested_by = st.text_input("Solicitado por", value="admin", key="arca_job_by")

    if cuit.strip():
        label, note = _badge_acceso(cuit, razon, remote)
        if label == "Listo":
            st.success(f"Acceso: **{label}** — {note}")
        elif label == "Pedir acceso":
            st.warning(f"Acceso: **{label}** — {note}")
            st.info(
                "CUIT nuevo o sin sesión: el admin abre AFIP en la PC del worker (2FA una vez), "
                "después marca el CUIT como Listo en la pestaña **CUITs / acceso**."
            )
        elif label:
            st.info(f"Acceso: **{label}** — {note}")

    params: dict[str, Any] = {}
    ruta = st.text_input(
        "Ruta destino (UNC)",
        placeholder=r"\\TANGOSRV\Compartido\CLIENTES\...\Facturas\08-2026",
        key="arca_job_ruta",
    )
    params["ruta_destino"] = ruta.strip().strip('"').strip("'")

    plantilla_b64 = ""
    plantilla_name = ""
    if action == "emitir_fcc":
        plantilla = st.file_uploader(
            "Plantilla Excel (montos / receptor / concepto — SIN claves)",
            type=["xlsx", "xls"],
            key="arca_job_xlsx",
        )
        fecha_emision = st.date_input("Fecha emisión", value=date.today(), key="arca_job_emision")
        params["fecha_emision"] = fecha_emision.isoformat()
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
                key="arca_job_desde",
            )
        with d2:
            hasta = st.date_input("Período hasta", value=date.today(), key="arca_job_hasta")
        params["periodo_desde"] = desde.isoformat()
        params["periodo_hasta"] = hasta.isoformat()
        if action == "bajar_comprobantes":
            params["analizar_monotributo"] = st.checkbox(
                "Al terminar, analizar recategorización (período facturado, recibos, NC)",
                value=True,
                key="arca_job_analizar_mono",
                help="El worker arma el papel de trabajo de monotributo con los PDF bajados.",
            )

    if st.button("Encolar", type="primary", key="arca_job_enqueue"):
        if not cuit.strip() or not razon.strip():
            st.error("CUIT y razón social son obligatorios.")
        elif not ruta.strip():
            st.error("Ruta destino UNC es obligatoria.")
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
                        requested_by=requested_by or "admin",
                        plantilla_b64=plantilla_b64,
                        plantilla_name=plantilla_name,
                    )
                    st.success(f"Encolado en RECEPCION `{job.get('id')}`")
                    st.caption("Lo toma la PC RECEPCION sola. El resto del estudio no tiene que abrir AFIP. Solo frena si el CUIT dice **Pedir acceso**.")
                    st.json(job)
                else:
                    ensure_cuit_registered(cuit, razon)
                    job_obj = create_job(
                        cuit=cuit,
                        razon_social=razon,
                        action=action,  # type: ignore[arg-type]
                        params=params,
                        requested_by=requested_by or "admin",
                    )
                    path = enqueue_job(job_obj)
                    st.success(f"Encolado `{job_obj.id}` → `{path.name}`")
                    st.caption(
                        "Lo toma la PC RECEPCION sola. El resto del estudio no tiene que abrir AFIP. "
                        "Solo frena si el CUIT dice **Pedir acceso**."
                    )
                    st.json(job_obj.to_dict())
            except (ValueError, RemoteError) as exc:
                st.error(str(exc))


def _render_cola(remote: RemoteWorker | None) -> None:
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
        auth = j.get("auth") or {}
        result = j.get("result") or {}
        rows.append(
            {
                "Estado": _STATUS_BADGE.get(str(j.get("status")), j.get("status")),
                "ID": j.get("id"),
                "CUIT": j.get("cuit"),
                "Cliente": j.get("razon_social"),
                "Acción": _ACTION_LABELS.get(str(j.get("action")), j.get("action")),
                "Auth": auth.get("status") if isinstance(auth, dict) else "",
                "Creado": j.get("created_at"),
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
            f"**{len(needs)} job(s) needs_auth.** "
            "Abrí AFIP conmigo en la PC del worker (2FA una vez) y marcá el CUIT como Listo."
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
                    "Actualizado": e.get("updated_at"),
                    "Perfil Chrome": e.get("chrome_profile"),
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
