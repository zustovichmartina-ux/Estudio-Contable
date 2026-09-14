# -*- coding: utf-8 -*-
"""Live: bajar VEPs (Volantes Electrónicos de Pago) del período."""
from __future__ import annotations

import logging
import re
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout

from ..browser import chrome_context
from ..jobs import Job, jobs_root, limpiar_ruta
from ..naming import nombre_vep, safe_name
from .result import ActionResult
from .session import (
    _active,
    _click_first,
    _completar_fechas,
    _elegir_empresa,
    _is_forbidden,
    _login_afip,
    _screenshot,
    abrir_servicio,
    copiar_si_nuevo,
    guardar_descarga,
)

LOG = logging.getLogger("afip_worker.veps")

_RE_VEP = re.compile(
    r"VEP|Volante Electr[oó]nico|Administraci[oó]n de VEP|Consulta de VEP",
    re.I,
)
_RE_BUSCAR = re.compile(r"^Buscar$|^Consultar$|^Filtrar$", re.I)
_RE_PDF = re.compile(r"pdf|imprimir|ver vep|descargar|constancia", re.I)
_RE_NO = re.compile(r"no se encontr|sin vep|0 vep|no existen", re.I)

_VEP_URLS = (
    "https://serviciosweb.afip.gob.ar/clavefiscal/veps/consulta.aspx",
    "https://serviciosweb.afip.gob.ar/clavefiscal/veps/",
    "https://seti.afip.gob.ar/padron-puc-constancia-internet/ConsultaVEPAction.do",
)


def run_bajar_veps_live(job: Job) -> ActionResult:
    dest = Path(limpiar_ruta(str(job.params.get("ruta_destino") or "")))
    if not str(dest):
        return ActionResult(ok=False, message="params.ruta_destino es obligatorio")
    try:
        dest.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return ActionResult(ok=False, message=f"No se puede crear destino {dest}: {exc}")

    desde = str(job.params.get("periodo_desde") or "")[:10]
    hasta = str(job.params.get("periodo_hasta") or "")[:10]
    if not desde or not hasta:
        return ActionResult(ok=False, message="periodo_desde y periodo_hasta son obligatorios")

    tmp_dl = jobs_root() / "running" / f"{job.id}_dl"
    tmp_dl.mkdir(parents=True, exist_ok=True)
    saved: list[str] = []

    try:
        with chrome_context(download_dir=tmp_dl) as (_ctx, page, _attached):
            auth = _login_afip(page, job)
            if auth is not None:
                _screenshot(page, job, "login")
                return auth

            host = _abrir_veps(page, job)
            if host is None:
                _screenshot(_active(page), job, "portal")
                return ActionResult(
                    ok=False,
                    message="No encontré el servicio de VEPs. ¿Está adherido a este CUIT?",
                )

            err = _descargar_veps(host, job, dest, tmp_dl, desde, hasta, saved)
            if err and not saved:
                _screenshot(host, job, "veps")
                return ActionResult(ok=False, message=err)
    except PlaywrightTimeout as exc:
        return ActionResult(ok=False, message=f"Timeout AFIP: {exc}")
    except Exception as exc:
        return ActionResult(ok=False, message=f"{type(exc).__name__}: {exc}")

    extra = f" ({err})" if saved and err else ""
    return ActionResult(
        ok=True,
        message=f"VEPs {desde}→{hasta}: {len(saved)} archivo(s){extra}",
        files=saved,
    )


def _abrir_veps(page: Page, job: Job) -> Page | None:
    host = abrir_servicio(page, job, query="VEP", name_re=_RE_VEP)
    if host is not None and not _is_forbidden(host):
        return host
    for url in _VEP_URLS:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(1_200)
            host = _active(page)
            if _is_forbidden(host):
                continue
            _elegir_empresa(host, job)
            return host
        except Exception as exc:
            LOG.warning("VEP goto %s: %s", url, exc)
    return None


def _descargar_veps(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    desde: str,
    hasta: str,
    saved: list[str],
) -> str:
    if not _completar_fechas(page, desde, hasta):
        labeled = page.get_by_label(re.compile(r"fecha|per[ií]odo", re.I))
        if labeled.count() == 0 and not _click_first(
            page,
            [
                page.get_by_role("button", name=_RE_BUSCAR),
                page.get_by_text(_RE_BUSCAR),
            ],
        ):
            return "Entré a VEPs pero no pude cargar fechas ni Buscar"
    _click_first(
        page,
        [
            page.get_by_role("button", name=_RE_BUSCAR),
            page.locator("input[value='Buscar'], input[value='Consultar']"),
            page.get_by_text(_RE_BUSCAR),
        ],
    )
    page.wait_for_timeout(2_000)
    body = ""
    try:
        body = page.locator("body").inner_text(timeout=5_000)
    except Exception:
        pass
    if body and _RE_NO.search(body):
        return ""

    n = _bajar_filas(page, job, dest, tmp_dl, saved)
    if n == 0:
        bulk = page.get_by_text(_RE_PDF)
        if bulk.count():
            files = guardar_descarga(page, bulk.first, tmp_dl)
            for src in files:
                dest_path = dest / (src.name if src.suffix else f"{src.name}.pdf")
                copied = copiar_si_nuevo(src, dest_path)
                if copied:
                    saved.append(copied)
                    n += 1
    if n == 0 and body and not _RE_NO.search(body):
        return f"{desde}: vi VEPs pero no pude bajar los PDF"
    return ""


def _bajar_filas(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    saved: list[str],
) -> int:
    n = 0
    rows = page.locator("table tbody tr, table tr")
    count = min(rows.count(), 80)
    for i in range(count):
        row = rows.nth(i)
        try:
            text = (row.inner_text(timeout=1_500) or "").strip()
        except Exception:
            continue
        if not text or ("fecha" in text.lower() and "importe" in text.lower()):
            continue
        icon = row.get_by_role("link", name=_RE_PDF)
        if icon.count() == 0:
            icon = row.locator("a[href*='pdf' i], a[title*='pdf' i], input[type='image']")
        if icon.count() == 0:
            continue
        nro = _nro_vep(text)
        filename = nombre_vep(
            nro=nro or f"{i + 1:04d}",
            concepto=safe_name(job.razon_social or "VEP", 24),
            importe=_importe(text),
            periodo_aaaa_mm=(str(job.params.get("periodo_hasta") or "")[:7] or "0000-00"),
        )
        dest_path = dest / filename
        files = guardar_descarga(page, icon.first, tmp_dl)
        for src in files:
            copied = copiar_si_nuevo(src, dest_path if src.suffix.lower() == ".pdf" else dest / src.name)
            if copied:
                saved.append(copied)
                n += 1
        page.wait_for_timeout(300)
    return n


def _nro_vep(text: str) -> str:
    hit = re.search(r"\b(\d{8,14})\b", text.replace(".", ""))
    return hit.group(1) if hit else ""


def _importe(text: str) -> str:
    hit = re.search(r"\$?\s*([\d\.]+,\d{2})", text)
    return hit.group(1) if hit else "0.00"
