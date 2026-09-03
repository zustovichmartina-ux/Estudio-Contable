"""UI Streamlit ARCA: solo encola jobs. Nunca claves. Badge de acceso por CUIT."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

import streamlit as st

from afip_worker.auth import admin_mark_cuit_ready
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
    "bajar_comprobantes": "Descargar comprobantes",
}

_STATUS_BADGE = {
    "pending": "pending",
    "running": "running",
    "needs_auth": "needs_auth",
    "done": "done",
    "error": "error",
}


def _badge_acceso(cuit: str, razon: str = "") -> tuple[str, str]:
    entry = ensure_cuit_registered(cuit, razon)
    return badge_label(entry.status), entry.note


def render_arca_module() -> None:
    """Módulo top-level ARCA: encolar + cola + registry (sin ejecutar AFIP)."""
    st.caption(
        "ARCA **solo encola** trabajos. El worker local (PC RECEPCION) abre Chrome/AFIP. "
        "Las claves viven solo en el autofill de Chrome — nunca en Excel ni en esta web."
    )
    root = jobs_root()
    st.caption(f"Cola: `{root}`")

    tab_encolar, tab_cola, tab_cuits = st.tabs(["Encolar", "Cola", "CUITs / acceso"])

    with tab_encolar:
        _render_encolar()
    with tab_cola:
        _render_cola()
    with tab_cuits:
        _render_registry()


def render_afip_cola_admin() -> None:
    """Alias retrocompatible."""
    render_arca_module()


def _render_encolar() -> None:
    conocidos = list_cuits()
    cuit_opts = [""] + [f"{e.cuit} — {e.razon_social or '(sin nombre)'}" for e in conocidos]

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
        label, note = _badge_acceso(cuit, razon)
        if label == "Listo":
            st.success(f"Acceso: **{label}** — {note}")
        elif label == "Pedir acceso":
            st.warning(f"Acceso: **{label}** — {note}")
            st.info(
                "CUIT nuevo o sin sesión: el admin abre AFIP en la PC del worker (2FA una vez), "
                "después marca el CUIT como Listo en la pestaña **CUITs / acceso**."
            )
        else:
            st.info(f"Acceso: **{label}** — {note}")

    params: dict[str, Any] = {}
    ruta = st.text_input(
        "Ruta destino (UNC)",
        placeholder=r"\\TANGOSRV\Compartido\CLIENTES\...\Facturas\08-2026",
        key="arca_job_ruta",
    )
    params["ruta_destino"] = ruta.strip()

    if action == "emitir_fcc":
        plantilla = st.file_uploader(
            "Plantilla Excel (montos / receptor / concepto — SIN claves)",
            type=["xlsx", "xls"],
            key="arca_job_xlsx",
        )
        fecha_emision = st.date_input("Fecha emisión", value=date.today(), key="arca_job_emision")
        params["fecha_emision"] = fecha_emision.isoformat()
        if plantilla is not None:
            dest = uploads_dir() / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}_{plantilla.name}"
            dest.write_bytes(plantilla.getvalue())
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

    if st.button("Encolar", type="primary", key="arca_job_enqueue"):
        if not cuit.strip() or not razon.strip():
            st.error("CUIT y razón social son obligatorios.")
        elif not ruta.strip():
            st.error("Ruta destino UNC es obligatoria.")
        elif action == "emitir_fcc" and not params.get("plantilla_excel"):
            st.error("Subí la plantilla Excel para emitir FCC.")
        else:
            ensure_cuit_registered(cuit, razon)
            job = create_job(
                cuit=cuit,
                razon_social=razon,
                action=action,  # type: ignore[arg-type]
                params=params,
                requested_by=requested_by or "admin",
            )
            path = enqueue_job(job)
            st.success(f"Encolado `{job.id}` → `{path.name}`")
            st.json(job.to_dict())


def _render_cola() -> None:
    jobs = list_jobs()
    if not jobs:
        st.info("Sin trabajos todavía.")
        return
    rows = []
    for j in reversed(jobs):
        rows.append(
            {
                "Estado": _STATUS_BADGE.get(j.status, j.status),
                "ID": j.id,
                "CUIT": j.cuit,
                "Cliente": j.razon_social,
                "Acción": _ACTION_LABELS.get(j.action, j.action),
                "Auth": j.auth.status,
                "Creado": j.created_at,
                "Mensaje": (j.result.message or "")[:140],
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)

    needs = [j for j in jobs if j.status == "needs_auth"]
    if needs:
        st.warning(
            f"**{len(needs)} job(s) needs_auth.** "
            "Abrí AFIP conmigo en la PC del worker (2FA una vez) y marcá el CUIT como Listo."
        )
        for j in needs:
            st.code(f"{j.id}\nCUIT {j.cuit}\n{j.auth.note or j.result.message}")


def _render_registry() -> None:
    st.markdown("##### Registry de CUITs")
    st.caption("Solo estado de acceso. Sin contraseñas ni tokens.")
    entries = list_cuits()
    if entries:
        st.dataframe(
            [
                {
                    "CUIT": e.cuit,
                    "Razón social": e.razon_social,
                    "Acceso": badge_label(e.status),
                    "Nota": e.note,
                    "Actualizado": e.updated_at,
                    "Perfil Chrome": e.chrome_profile,
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
                admin_mark_cuit_ready(cuit_h, razon_h)
                st.success(f"{cuit_h} → Listo")
                st.rerun()
        if st.button("Marcar Pedir acceso", key="arca_hand_need"):
            if not cuit_h.strip():
                st.error("Indicá el CUIT.")
            else:
                mark_needs_admin(cuit_h, razon_social=razon_h)
                st.warning(f"{cuit_h} → Pedir acceso")
                st.rerun()
