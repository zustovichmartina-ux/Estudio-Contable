# -*- coding: utf-8 -*-
"""Stubs de acciones AFIP. Dry-run: no abre Chrome ni escribe secretos."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from ..jobs import Job
from ..naming import nombre_comprobante, nombre_fcc, nombre_vep


@dataclass
class ActionResult:
    ok: bool
    message: str = ""
    files: list[str] = field(default_factory=list)
    needs_auth: bool = False


def _destino(job: Job) -> Path:
    raw = str(job.params.get("ruta_destino") or "").strip()
    if not raw:
        raise ValueError("params.ruta_destino es obligatorio")
    return Path(raw)


def run_emitir_fcc(job: Job, *, dry_run: bool = True) -> ActionResult:
    dest = _destino(job)
    if dry_run:
        fake = dest / nombre_fcc(cliente=job.razon_social or "Cliente", pv=1, nro=1)
        return ActionResult(
            ok=True,
            message=f"[dry-run] emitir_fcc → {fake} (sin AFIP)",
            files=[str(fake)],
        )
    return ActionResult(ok=False, message="emitir_fcc real: pendiente Playwright")


def run_bajar_veps(job: Job, *, dry_run: bool = True) -> ActionResult:
    dest = _destino(job)
    periodo = str(job.params.get("periodo_hasta") or job.params.get("periodo_desde") or "")[:7]
    if dry_run:
        fake = dest / nombre_vep(
            nro="000000",
            concepto="dryrun",
            importe="0.00",
            periodo_aaaa_mm=periodo or "0000-00",
        )
        return ActionResult(
            ok=True,
            message=f"[dry-run] bajar_veps período {periodo} → {fake}",
            files=[str(fake)],
        )
    return ActionResult(ok=False, message="bajar_veps real: pendiente Playwright")


def run_bajar_comprobantes(job: Job, *, dry_run: bool = True) -> ActionResult:
    dest = _destino(job)
    desde = str(job.params.get("periodo_desde") or "")
    hasta = str(job.params.get("periodo_hasta") or "")
    if dry_run:
        fake = dest / nombre_comprobante(
            pv=1,
            nro=1,
            emision=hasta or desde or "0000-00-00",
            desde=desde or "0000-00-00",
            hasta=hasta or "0000-00-00",
        )
        return ActionResult(
            ok=True,
            message=f"[dry-run] bajar_comprobantes {desde}→{hasta} → {fake}",
            files=[str(fake)],
        )
    return ActionResult(ok=False, message="bajar_comprobantes real: pendiente Playwright")


DISPATCH = {
    "emitir_fcc": run_emitir_fcc,
    "bajar_veps": run_bajar_veps,
    "bajar_comprobantes": run_bajar_comprobantes,
}


def run_action(job: Job, *, dry_run: bool = True) -> ActionResult:
    fn = DISPATCH.get(job.action)
    if not fn:
        return ActionResult(ok=False, message=f"action desconocida: {job.action}")
    try:
        return fn(job, dry_run=dry_run)
    except Exception as exc:
        return ActionResult(ok=False, message=f"{type(exc).__name__}: {exc}")
