"""UI Streamlit: encolar jobs AFIP y ver estados (sin claves)."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

import streamlit as st

from afip_worker.jobs import (
    ACTIONS,
    create_job,
    enqueue_job,
    jobs_root,
    list_jobs,
    uploads_dir,
)


_ACTION_LABELS = {
    "emitir_fcc": "Emitir facturas (FCC)",
    "bajar_veps": "Descargar VEPs",
    "bajar_comprobantes": "Descargar comprobantes",
}

_STATUS_BADGE = {
    "pending": "🟡 pending",
    "running": "🔵 running",
    "needs_auth": "🟠 needs_auth",
    "done": "🟢 done",
    "error": "🔴 error",
}


def _auth_badge(cuit: str) -> str:
    """v1: sin Chrome check en Cloud — siempre pedir acceso si no hay jobs done del CUIT."""
    jobs = [j for j in list_jobs() if j.cuit.replace("-", "") == cuit.replace("-", "")]
    if any(j.status == "done" and j.auth.status == "ready" for j in jobs):
        return "Listo (hubo corridas OK)"
    if any(j.status == "needs_auth" for j in jobs):
        return "Pedir acceso (needs_auth)"
    return "Desconocido — Pedir acceso en 1er login"


def render_afip_cola_admin() -> None:
    st.markdown("#### AFIP — Cola de trabajos")
    st.caption(
        "La web **solo encola**. El worker local (PC RECEPCION / Cursor) ejecuta AFIP. "
        "Nunca se guardan claves acá ni en el Excel."
    )
    root = jobs_root()
    st.caption(f"Cola: `{root}`")

    with st.expander("Encolar trabajo", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            cuit = st.text_input("CUIT", placeholder="27-42043034-0", key="afip_job_cuit")
            razon = st.text_input(
                "Razón social",
                placeholder="Camila Rocio Albarello",
                key="afip_job_razon",
            )
        with c2:
            action = st.selectbox(
                "Acción",
                options=list(ACTIONS),
                format_func=lambda a: _ACTION_LABELS.get(a, a),
                key="afip_job_action",
            )
            requested_by = st.text_input("Solicitado por", value="admin", key="afip_job_by")

        if cuit.strip():
            st.info(f"Acceso CUIT: **{_auth_badge(cuit)}**")

        params: dict[str, Any] = {}
        ruta = st.text_input(
            "Ruta destino (UNC)",
            placeholder=r"\\TANGOSRV\Compartido\CLIENTES\...\Facturas\08-2026",
            key="afip_job_ruta",
        )
        params["ruta_destino"] = ruta.strip()

        if action == "emitir_fcc":
            plantilla = st.file_uploader(
                "Plantilla Excel (montos / receptor / concepto — SIN claves)",
                type=["xlsx", "xls"],
                key="afip_job_xlsx",
            )
            fecha_emision = st.date_input(
                "Fecha emisión",
                value=date.today(),
                key="afip_job_emision",
            )
            params["fecha_emision"] = fecha_emision.isoformat()
            if plantilla is not None:
                dest = uploads_dir() / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}_{plantilla.name}"
                dest.write_bytes(plantilla.getvalue())
                params["plantilla_excel"] = str(dest.relative_to(Path(__file__).resolve().parents[1]))
                st.caption(f"Plantilla guardada: `{params['plantilla_excel']}`")
        else:
            d1, d2 = st.columns(2)
            with d1:
                desde = st.date_input("Período desde", value=date.today().replace(day=1), key="afip_job_desde")
            with d2:
                hasta = st.date_input("Período hasta", value=date.today(), key="afip_job_hasta")
            params["periodo_desde"] = desde.isoformat()
            params["periodo_hasta"] = hasta.isoformat()

        force_auth = st.checkbox(
            "Simular needs_auth (solo prueba dry-run)",
            value=False,
            key="afip_job_force_auth",
        )
        if force_auth:
            params["force_needs_auth"] = True

        if st.button("Encolar", type="primary", key="afip_job_enqueue"):
            if not cuit.strip() or not razon.strip():
                st.error("CUIT y razón social son obligatorios.")
            elif not ruta.strip():
                st.error("Ruta destino UNC es obligatoria.")
            elif action == "emitir_fcc" and not params.get("plantilla_excel"):
                st.error("Subí la plantilla Excel para emitir FCC.")
            else:
                job = create_job(
                    cuit=cuit,
                    razon_social=razon,
                    action=action,  # type: ignore[arg-type]
                    params=params,
                    requested_by=requested_by or "admin",
                )
                path = enqueue_job(job)
                st.success(f"Encolado `{job.id}` → `{path}`")
                st.json(job.to_dict())

    st.divider()
    st.markdown("##### Cola")
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
                "Mensaje": (j.result.message or "")[:120],
            }
        )
    st.dataframe(rows, use_container_width=True, hide_index=True)

    needs = [j for j in jobs if j.status == "needs_auth"]
    if needs:
        st.warning(
            f"**{len(needs)} job(s) needs_auth.** "
            "Abrí AFIP conmigo: el admin hace login/2FA una vez en la PC del worker "
            "y luego reencolá o reintentá el job."
        )
        with st.expander("Jobs needs_auth"):
            for j in needs:
                st.code(f"{j.id}\n{j.auth.note or j.result.message}")
