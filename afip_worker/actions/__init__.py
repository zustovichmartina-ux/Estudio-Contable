# -*- coding: utf-8 -*-
"""Acciones AFIP: dry-run (sin Chrome) o live Playwright."""
from __future__ import annotations

from datetime import date as _date
from pathlib import Path

from ..jobs import Job, limpiar_ruta
from ..naming import nombre_comprobante, nombre_fcc, nombre_vep
from .comprobantes import run_bajar_comprobantes_live
from .result import ActionResult

__all__ = ["ActionResult", "run_action", "DISPATCH"]


def _destino(job: Job) -> Path:
    raw = limpiar_ruta(str(job.params.get("ruta_destino") or ""))
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


def _analizar_monotributo_si_pide(job: Job, result: ActionResult) -> ActionResult:
    if not result.ok:
        return result
    flag = job.params.get("analizar_monotributo")
    if not flag or str(flag).lower() in {"0", "false", "no"}:
        return result
    rutas = list(result.files or [])
    dest = _destino(job)
    if dest.is_dir():
        for pdf in dest.glob("*.pdf"):
            if str(pdf) not in rutas:
                rutas.append(str(pdf))
    pdfs = [p for p in rutas if str(p).lower().endswith(".pdf")]
    if not pdfs:
        extra = dict(result.extra or {})
        extra["monotributo"] = {
            "cantidad": 0,
            "neto": 0.0,
            "xlsx": "",
            "errores": [{"archivo": "(job)", "motivo": "No hay PDFs para analizar."}],
        }
        result.extra = extra
        result.message = f"{result.message} · recategorización: sin PDFs"
        return result
    # procesador es pesado; se importa acá para no cargarlo en dry-run sin PDFs.
    from procesador import analizar_comprobantes_monotributo_rutas

    digits = "".join(ch for ch in str(job.cuit or "") if ch.isdigit()) or "00000000000"
    xlsx = dest / f"Recategorizacion_Monotributo_{digits}_{_date.today().strftime('%Y%m%d')}.xlsx"
    info = analizar_comprobantes_monotributo_rutas(
        pdfs,
        cuit_cliente=job.cuit,
        dest_xlsx=xlsx,
    )
    extra = dict(result.extra or {})
    extra["monotributo"] = info
    result.extra = extra
    files = list(result.files or [])
    if info.get("xlsx") and info["xlsx"] not in files:
        files.append(info["xlsx"])
        result.files = files
    result.message = (
        f"{result.message} · recategorización: {info.get('cantidad', 0)} cmpte(s), "
        f"neto ${float(info.get('neto') or 0):,.2f}"
    )
    return result


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
        result = ActionResult(
            ok=True,
            message=f"[dry-run] bajar_comprobantes {desde}→{hasta} → {fake}",
            files=[str(fake)],
        )
        if job.params.get("analizar_monotributo"):
            extra = dict(result.extra or {})
            extra["monotributo"] = {
                "cantidad": 0,
                "neto": 0.0,
                "xlsx": "",
                "errores": [{"archivo": "(dry-run)", "motivo": "Sin PDFs reales; el análisis corre en --live."}],
            }
            result.extra = extra
        return result
    result = run_bajar_comprobantes_live(job)
    return _analizar_monotributo_si_pide(job, result)


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
