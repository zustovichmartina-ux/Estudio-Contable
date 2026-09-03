# -*- coding: utf-8 -*-
"""Auth AFIP: nunca claves. Registry CUIT + Chrome autofill / handoff 2FA."""
from __future__ import annotations

from dataclasses import dataclass

from .jobs import Job, JobAuth, mark_needs_auth
from .registry import ensure_cuit_registered, get_cuit, mark_needs_admin, mark_ready


@dataclass
class AuthCheck:
    ready: bool
    note: str = ""
    status: str = "unknown"


def check_session_ready(job: Job, *, dry_run: bool = True) -> AuthCheck:
    """
    Registry:
      - CUIT nuevo → needs_admin (UI muestra Pedir acceso).
      - ready → Chrome autofill OK.
    Worker (dry-run o live):
      - ready → actúa solo.
      - needs_admin / unknown / force_needs_auth → needs_auth (solo ahí pide admin).
    """
    ensure_cuit_registered(job.cuit, job.razon_social)

    if job.params.get("force_needs_auth"):
        mark_needs_admin(
            job.cuit,
            razon_social=job.razon_social,
            note="force_needs_auth: admin debe abrir AFIP y completar 2FA",
        )
        return AuthCheck(
            ready=False,
            status="needs_admin",
            note="force_needs_auth: el admin debe abrir AFIP y completar 2FA una vez",
        )

    entry = get_cuit(job.cuit)

    # Autónomo: si el CUIT está Listo, sigue solo (dry-run o live).
    # Solo frena y pide admin cuando no hay acceso AFIP.
    if entry and entry.status == "ready":
        note = entry.note or "Registry: ready (Chrome autofill)"
        if dry_run:
            note = f"dry-run + {note}"
        return AuthCheck(ready=True, status="ready", note=note)

    status = entry.status if entry else "needs_admin"
    note = (
        (entry.note if entry else "")
        or "CUIT sin sesión AFIP: handoff admin / 2FA en PC worker"
    )
    mark_needs_admin(job.cuit, razon_social=job.razon_social, note=note)
    return AuthCheck(ready=False, status=status, note=note)


def apply_auth(job: Job, *, dry_run: bool = True) -> Job:
    chk = check_session_ready(job, dry_run=dry_run)
    if chk.ready:
        job.auth = JobAuth(status="ready", note=chk.note)
        return job
    mark_needs_auth(job, chk.note)
    return job


def admin_mark_cuit_ready(cuit: str, razon_social: str = "") -> None:
    """Tras handoff 2FA exitoso en la PC del worker (sin guardar claves)."""
    mark_ready(
        cuit,
        razon_social=razon_social,
        note="Admin confirmó login AFIP / Chrome autofill listo",
    )


def lookup_status(cuit: str) -> str:
    entry = get_cuit(cuit)
    return entry.status if entry else "unknown"
