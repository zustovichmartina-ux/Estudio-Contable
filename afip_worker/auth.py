# -*- coding: utf-8 -*-
"""Auth AFIP: nunca claves. Solo estado ready / needs_admin (handoff 2FA)."""
from __future__ import annotations

from dataclasses import dataclass

from .jobs import Job, JobAuth, mark_needs_auth


@dataclass
class AuthCheck:
    ready: bool
    note: str = ""


def check_session_ready(job: Job, *, dry_run: bool = True) -> AuthCheck:
    """
    En dry-run: auth ready salvo que params.force_needs_auth=True.
    En producción (v2): inspeccionar perfil Chrome / cookies del CUIT.
    """
    if job.params.get("force_needs_auth"):
        return AuthCheck(
            ready=False,
            note="force_needs_auth: el admin debe abrir AFIP y completar 2FA una vez",
        )
    if dry_run:
        return AuthCheck(ready=True, note="dry-run: auth simulada OK")
    # Stub real: sin Chrome/session check todavía
    return AuthCheck(
        ready=False,
        note="Auth real no implementada: marcar needs_auth hasta Playwright + perfil Chrome",
    )


def apply_auth(job: Job, *, dry_run: bool = True) -> Job:
    chk = check_session_ready(job, dry_run=dry_run)
    if chk.ready:
        job.auth = JobAuth(status="ready", note=chk.note)
        return job
    mark_needs_auth(job, chk.note)
    return job
