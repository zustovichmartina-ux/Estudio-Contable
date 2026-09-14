# -*- coding: utf-8 -*-
"""Live: login AFIP/ARCA y bajar PDFs de Comprobantes en Línea (RCEL)."""
from __future__ import annotations

import calendar
import logging
import re
import shutil
from datetime import date
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeout

from ..browser import chrome_context
from ..jobs import Job, jobs_root, limpiar_ruta
from ..naming import nombre_fcc, safe_name
from .result import ActionResult
from .session import (
    _active,
    _click_first,
    _completar_fechas,
    _elegir_empresa,
    _is_forbidden,
    _login_afip,
    _screenshot,
)

LOG = logging.getLogger("afip_worker.comprobantes")

_RE_RCEL = re.compile(r"^Comprobantes en l[ií]nea", re.I)
_RE_CONSULTAS = re.compile(r"^Consultas$|Consultar Comprobantes", re.I)
_RE_BUSCAR = re.compile(r"^Buscar$|^Consultar$", re.I)
_RE_NO_RES = re.compile(r"no se encontr|sin comprobante|0 comprobante|no existen", re.I)
_RE_PV_NRO = re.compile(r"(\d{1,5})\s*[-/]\s*(\d{1,8})")
_RE_MENU = re.compile(r"Consultas|Generar Comprobantes", re.I)


def run_bajar_comprobantes_live(job: Job) -> ActionResult:
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

            if not _abrir_comprobantes_en_linea(page, job):
                _screenshot(_active(page), job, "portal")
                return ActionResult(
                    ok=False,
                    message="No encontré Comprobantes en Línea. ¿El servicio está adherido a este CUIT?",
                )

            host = _active(page)
            if not _abrir_consultas(host):
                _screenshot(host, job, "menu")
                return ActionResult(
                    ok=False,
                    message="Entré a Comprobantes en Línea pero no encontré Consultas.",
                )

            host = _active(page)
            err = _descargar_consultas(host, job, dest, tmp_dl, desde, hasta, saved)
            if err and not saved:
                _screenshot(host, job, "consulta")
                return ActionResult(ok=False, message=err)
    except PlaywrightTimeout as exc:
        return ActionResult(ok=False, message=f"Timeout AFIP: {exc}")
    except Exception as exc:
        return ActionResult(ok=False, message=f"{type(exc).__name__}: {exc}")

    extra = f" ({err})" if saved and err else ""
    return ActionResult(
        ok=True,
        message=f"Comprobantes en Línea {desde}→{hasta}: {len(saved)} PDF(s){extra}",
        files=saved,
    )


def _meses(desde: str, hasta: str) -> list[tuple[str, str]]:
    d0 = date.fromisoformat(desde[:10])
    d1 = date.fromisoformat(hasta[:10])
    out: list[tuple[str, str]] = []
    year, month = d0.year, d0.month
    while (year, month) <= (d1.year, d1.month):
        last = calendar.monthrange(year, month)[1]
        start = date(year, month, 1)
        end = date(year, month, last)
        if start < d0:
            start = d0
        if end > d1:
            end = d1
        out.append((start.isoformat(), end.isoformat()))
        if month == 12:
            year, month = year + 1, 1
        else:
            month += 1
    return out


def _scope(page: Page) -> Any:
    for frame in page.frames:
        url = (frame.url or "").lower()
        if "rcel" in url or "fe.afip" in url:
            return frame
    return page


def _abrir_comprobantes_en_linea(page: Page, job: Job) -> bool:
    if _en_rcel(page):
        return True
    buscador = page.locator(
        "input[type='search'], input[placeholder*='Buscar' i], #buscador, #inputBuscar, "
        "input[aria-label*='Buscar' i], input[id*='buscar' i], input[name*='buscar' i]"
    )
    if buscador.count():
        try:
            buscador.first.click()
            buscador.first.fill("Comprobantes en línea")
            buscador.first.press("Enter")
            page.wait_for_timeout(1_500)
        except Exception:
            LOG.warning("buscador del portal no respondió")

    pages_before = list(page.context.pages)
    if _click_rcel_link(page):
        page.wait_for_timeout(2_000)
        host = _active(page)
        for extra in page.context.pages:
            if extra not in pages_before:
                extra.bring_to_front()
                host = extra
                break
        if _is_forbidden(host):
            LOG.warning("RCEL 403 tras el click del portal")
            return False
        _elegir_empresa(host, job)
        return _en_rcel(host) or _wait_rcel(host)

    LOG.warning("no encontré el servicio Comprobantes en línea en el portal")
    return False


def _click_rcel_link(page: Page) -> bool:
    return _click_first(
        page,
        [
            page.get_by_role("link", name=_RE_RCEL),
            page.get_by_role("button", name=_RE_RCEL),
            page.get_by_text(_RE_RCEL),
        ],
    )


def _en_rcel(page: Page) -> bool:
    if _is_forbidden(page):
        return False
    try:
        body = page.locator("body").inner_text(timeout=3_000)
    except Exception:
        return False
    if re.search(r"Mis\s+Comprobantes", body, re.I) and "Generar Comprobantes" not in body:
        return False
    return bool(_RE_MENU.search(body))


def _wait_rcel(page: Page) -> bool:
    try:
        page.get_by_text(_RE_MENU).first.wait_for(state="visible", timeout=20_000)
        return True
    except PlaywrightTimeout:
        return _en_rcel(page)


def _abrir_consultas(page: Page) -> bool:
    scope = _scope(page)
    return _click_first(
        page,
        [
            scope.get_by_role("link", name=_RE_CONSULTAS),
            scope.get_by_role("button", name=_RE_CONSULTAS),
            scope.get_by_text(_RE_CONSULTAS),
        ],
    )


def _descargar_consultas(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    desde: str,
    hasta: str,
    saved: list[str],
) -> str:
    scope = _scope(page)
    pvs = _valores_punto_venta(scope)
    LOG.info("puntos de venta: %s", pvs or "(único)")
    notes: list[str] = []
    for mes_desde, mes_hasta in _meses(desde, hasta):
        for pv in pvs or [""]:
            n_before = len(saved)
            err = _buscar_y_bajar(
                page, job, dest, tmp_dl, mes_desde, mes_hasta, pv, saved
            )
            if err:
                notes.append(err)
            else:
                LOG.info("mes %s pv=%s +%s PDF", mes_desde[:7], pv or "-", len(saved) - n_before)
    if not saved:
        return notes[0] if notes else "No bajó ningún PDF de Comprobantes en Línea"
    return "; ".join(notes[:3]) if notes else ""


def _valores_punto_venta(scope: Any) -> list[str]:
    sel = scope.get_by_label(re.compile(r"punto de venta", re.I))
    if sel.count() == 0:
        labeled = scope.locator("select")
        sel = labeled.first if labeled.count() else sel
    if sel.count() == 0:
        return [""]
    options = sel.locator("option")
    vals: list[str] = []
    todos = ""
    for i in range(options.count()):
        opt = options.nth(i)
        val = (opt.get_attribute("value") or "").strip()
        text = (opt.inner_text() or "").strip().lower()
        if not val or val in {"-1", "0"} or "seleccione" in text:
            continue
        if "todos" in text:
            todos = val
            break
        vals.append(val)
    if todos:
        return [todos]
    return vals or [""]


def _buscar_y_bajar(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    desde: str,
    hasta: str,
    pv: str,
    saved: list[str],
) -> str:
    scope = _scope(_active(page))
    if pv:
        try:
            sel = scope.get_by_label(re.compile(r"punto de venta", re.I))
            if sel.count() == 0:
                sel = scope.locator("select").first
            if sel.count():
                sel.select_option(pv)
        except Exception as exc:
            LOG.warning("select PV %s: %s", pv, exc)

    if not _completar_fechas(scope, desde, hasta):
        return f"{desde}: no pude cargar fechas"

    if not _click_first(
        page,
        [
            scope.get_by_role("button", name=_RE_BUSCAR),
            scope.locator("input[value='Buscar'], input[value='Consultar']"),
            scope.get_by_text(_RE_BUSCAR),
        ],
    ):
        return f"{desde}: no encontré Buscar"
    page.wait_for_timeout(2_000)
    scope = _scope(_active(page))

    body = ""
    try:
        body = scope.locator("body").inner_text(timeout=5_000)
    except Exception:
        pass
    if body and _RE_NO_RES.search(body):
        return ""

    n = _bajar_pdfs_listado(_active(page), scope, job, dest, tmp_dl, saved)
    if n == 0 and body and not _RE_NO_RES.search(body):
        return f"{desde}: vi el listado pero no pude bajar PDFs"
    return ""


def _bajar_pdfs_listado(
    page: Page,
    scope: Any,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    saved: list[str],
) -> int:
    n = 0
    cliente = job.razon_social or job.cuit
    rows = scope.locator("table tbody tr, table tr")
    count = rows.count()
    LOG.info("filas consulta: %s", count)
    for i in range(count):
        row = rows.nth(i)
        text = ""
        try:
            text = row.inner_text(timeout=2_000)
        except Exception:
            continue
        if not text.strip() or "fecha" in text.lower() and "tipo" in text.lower():
            continue
        icon = _icono_pdf(row)
        if icon is None:
            continue
        pv, nro = _pv_nro_de_fila(text)
        if pv is not None and nro is not None:
            filename = nombre_fcc(cliente=cliente, pv=pv, nro=nro)
        else:
            filename = f"{safe_name(cliente, 40)} consulta_{i + 1:03d}.pdf"
        path = dest / filename
        if path.exists() and path.stat().st_size > 500:
            saved.append(str(path))
            n += 1
            continue
        if _click_save_pdf(page, icon, path, tmp_dl):
            saved.append(str(path))
            n += 1
            LOG.info("PDF %s", path.name)
        page.wait_for_timeout(400)
    return n


def _pv_nro_de_fila(text: str) -> tuple[int | None, int | None]:
    hit = _RE_PV_NRO.search(text.replace(".", ""))
    if not hit:
        return None, None
    try:
        return int(hit.group(1)), int(hit.group(2))
    except ValueError:
        return None, None


def _icono_pdf(row: Locator) -> Locator | None:
    candidates = [
        row.locator("a[title*='Imprimir' i], a[title*='Ver' i], a[title*='PDF' i]"),
        row.locator("img[alt*='Imprimir' i], img[alt*='Ver' i], img[title*='Imprimir' i]"),
        row.locator("input[type='image'], input[src*='imprimir' i], input[src*='printer' i]"),
        row.get_by_role("link", name=re.compile(r"imprimir|ver|pdf", re.I)),
    ]
    for loc in candidates:
        try:
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def _click_save_pdf(page: Page, clickable: Locator, dest: Path, tmp_dl: Path) -> bool:
    before = {p.name for p in tmp_dl.glob("*")} if tmp_dl.exists() else set()
    pages_before = list(page.context.pages)
    try:
        clickable.click(timeout=8_000)
    except Exception:
        return False
    page.wait_for_timeout(2_200)
    for f in tmp_dl.glob("*"):
        if f.name in before:
            continue
        if f.suffix.lower() == ".pdf" and f.stat().st_size > 200:
            shutil.copyfile(f, dest)
            return dest.exists() and dest.stat().st_size > 200
    for extra in page.context.pages:
        if extra in pages_before:
            continue
        url = extra.url or ""
        try:
            extra.wait_for_load_state("domcontentloaded", timeout=10_000)
            url = extra.url or url
            if "pdf" in url.lower() or url.lower().endswith(".pdf"):
                body = extra.context.request.get(url).body()
                dest.write_bytes(body)
                extra.close()
                return dest.exists() and dest.stat().st_size > 200
            extra.close()
        except Exception:
            try:
                extra.close()
            except Exception:
                pass
    return dest.exists() and dest.stat().st_size > 200

