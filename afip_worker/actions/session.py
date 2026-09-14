# -*- coding: utf-8 -*-
"""Login AFIP/ARCA reutilizable: mismo Chrome persistente, sin claves en jobs."""
from __future__ import annotations

import logging
import re
import shutil
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeout

from ..browser import login_timeout_ms
from ..jobs import Job, jobs_root, normalizar_cuit
from ..naming import iso_to_ddmmyyyy
from .result import ActionResult

LOG = logging.getLogger("afip_worker.session")

LOGIN_URL = "https://auth.afip.gob.ar/contribuyente_/login.xhtml"
PORTAL_URL = "https://portalcf.cloud.afip.gob.ar/portal/app/"
PORTAL_HINTS = (
    "portalcf.cloud.afip.gob.ar",
    "portalcf.afip.gob.ar",
    "menuPrincipal",
    "serviciosjava",
    "fe.afip.gob.ar",
    "rcel",
    "arca.gob.ar",
    "iva.afip.gob.ar",
    "portaliva",
    "niva",
    "miscomprobantes",
)
LOGIN_HINTS = ("login.xhtml", "contribuyente_/login", "clave fiscal")
_RE_CUIT = re.compile(r"CUIT", re.I)
_RE_INGRESAR = re.compile(r"Ingresar|Siguiente|Continuar", re.I)


def _cuit_digits(job: Job) -> str:
    return re.sub(r"\D", "", normalizar_cuit(job.cuit))


def _active(page: Page) -> Page:
    pages = page.context.pages
    if not pages:
        return page
    nxt = pages[-1]
    try:
        nxt.bring_to_front()
    except Exception:
        pass
    return nxt


def _sesion_expirada(page: Page) -> bool:
    try:
        text = (page.locator("body").inner_text(timeout=3_000) or "").lower()
    except Exception:
        return False
    return "sesión ha expirado" in text or "sesion ha expirado" in text


def _is_login(page: Page) -> bool:
    if _sesion_expirada(page):
        return True
    url = (page.url or "").lower()
    return any(h in url for h in LOGIN_HINTS)


def _is_portal(page: Page) -> bool:
    if _sesion_expirada(page) or _is_login(page):
        return False
    url = (page.url or "").lower()
    if any(h.lower() in url for h in PORTAL_HINTS) and not _is_login(page):
        return True
    try:
        body = page.locator("body").inner_text(timeout=3_000)
    except Exception:
        return False
    low = body.lower()
    if "ingresar con clave" in low or "sesión ha expirado" in low:
        return False
    return "mis servicios" in low or "salir" in low


def _is_blocker(page: Page) -> bool:
    try:
        text = page.locator("body").inner_text(timeout=3_000).lower()
    except Exception:
        return False
    keys = (
        "captcha",
        "recaptcha",
        "token",
        "aplicación ciud",
        "codigo de 6",
        "código de 6",
        "autenticación",
        "segundo factor",
        "clave temporal",
    )
    return any(k in text for k in keys)


def _login_afip(page: Page, job: Job) -> ActionResult | None:
    """Reusa cookies. En segundo plano no espera a que alguien tipee la clave."""
    LOG.info("login AFIP cuit=%s", job.cuit)
    try:
        page.goto(PORTAL_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(800)
    except Exception:
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(800)
    if _sesion_expirada(page):
        LOG.info("sesión AFIP expirada; voy al login")
        page.wait_for_timeout(5_000)
        if not _is_login(page):
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
    if _is_portal(page) and not _is_login(page):
        LOG.info("sesión AFIP ya abierta (segundo plano)")
        return None

    digits = _cuit_digits(job)
    user = page.locator("#F1\\:username")
    if user.count() == 0:
        user = page.get_by_label(_RE_CUIT)
    if user.count() == 0:
        user = page.locator("input[name='F1:username'], input[type='text']").first
    try:
        user.first.wait_for(state="attached", timeout=15_000)
        user.first.click(force=True)
        user.first.fill(digits)
        user.first.press("Tab")
    except PlaywrightTimeout:
        if _is_portal(page):
            return None
        return ActionResult(
            needs_auth=True,
            ok=False,
            message="No apareció el campo CUIT de AFIP. Completá el login en Chrome.",
        )

    _click_first(
        page,
        [
            page.locator("#F1\\:btnSiguiente"),
            page.get_by_role("button", name=_RE_INGRESAR),
            page.locator("input[type='submit']"),
        ],
    )
    page.wait_for_timeout(1_200)

    pwd = page.locator("#F1\\:password, input[type='password']").first
    try:
        if pwd.count():
            pwd.click(force=True)
            page.wait_for_timeout(400)
            page.keyboard.press("ArrowDown")
            page.keyboard.press("Enter")
            LOG.info("reingreso automático (autofill, %ss)", login_timeout_ms() // 1000)
    except Exception:
        pass

    steps = max(8, login_timeout_ms() // 1000)
    logged_in = False
    for _ in range(steps):
        if _is_portal(page) and not _is_login(page):
            logged_in = True
            break
        if _is_blocker(page):
            break
        _try_submit_password(page)
        page.wait_for_timeout(1_000)
    if not logged_in:
        if _is_blocker(page):
            return ActionResult(
                needs_auth=True,
                ok=False,
                message="AFIP pidió 2FA. Completalo una vez con iniciar_afip_sesion.bat; después reingresa solo.",
            )
        return ActionResult(
            needs_auth=True,
            ok=False,
            message=(
                "No pude reingresar solo (falta clave guardada en el Chrome del worker). "
                "Una vez en RECEPCION: iniciar_afip_sesion.bat y aceptá 'guardar contraseña'."
            ),
        )
    _maybe_elegir_representado(page, job)
    return None


def _try_submit_password(page: Page) -> None:
    pwd = page.locator("#F1\\:password, input[type='password']").first
    try:
        if pwd.count() == 0:
            return
        if not (pwd.input_value() or "").strip():
            return
        _click_first(
            page,
            [
                page.locator("#F1\\:btnIngresar"),
                page.get_by_role("button", name=re.compile(r"^Ingresar$", re.I)),
            ],
        )
    except Exception:
        return


def _maybe_elegir_representado(page: Page, job: Job) -> None:
    for text in (normalizar_cuit(job.cuit), _cuit_digits(job), job.razon_social):
        if not text:
            continue
        loc = page.get_by_text(str(text), exact=False)
        try:
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=3_000)
                page.wait_for_timeout(800)
                return
        except Exception:
            continue


def _elegir_empresa(page: Page, job: Job) -> None:
    for text in (job.razon_social, normalizar_cuit(job.cuit), _cuit_digits(job)):
        if not (text or "").strip():
            continue
        loc = page.get_by_text(str(text), exact=False)
        try:
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=4_000)
                page.wait_for_timeout(1_000)
                return
        except Exception:
            continue


def _is_forbidden(page: Page) -> bool:
    try:
        text = (page.locator("body").inner_text(timeout=2_000) or "").lower()
    except Exception:
        text = ""
    return "forbidden" in text or "don't have permission" in text or "no tiene permiso" in text


def _click_first(page: Page, locators: list[Locator]) -> bool:
    for loc in locators:
        try:
            if loc.count() == 0:
                continue
            loc.first.click(timeout=8_000, force=True)
            page.wait_for_timeout(400)
            return True
        except Exception:
            continue
    return False


def _screenshot(page: Page, job: Job, tag: str) -> None:
    path = jobs_root() / "error" / f"{job.id}_{tag}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        LOG.info("screenshot %s", path)
    except Exception:
        return


def buscar_en_portal(page: Page, query: str) -> None:
    buscador = page.locator(
        "input[type='search'], input[placeholder*='Buscar' i], #buscador, #inputBuscar, "
        "input[aria-label*='Buscar' i], input[id*='buscar' i], input[name*='buscar' i]"
    )
    if buscador.count() == 0:
        return
    try:
        buscador.first.click()
        buscador.first.fill(query)
        buscador.first.press("Enter")
        page.wait_for_timeout(1_500)
    except Exception:
        LOG.warning("buscador del portal no respondió (%s)", query)


def abrir_servicio(
    page: Page,
    job: Job,
    *,
    query: str,
    name_re: re.Pattern[str],
) -> Page | None:
    """Busca un tile del portal, lo abre (nueva pestaña) y elige el representado."""
    buscar_en_portal(page, query)
    pages_before = list(page.context.pages)
    if not _click_first(
        page,
        [
            page.get_by_role("link", name=name_re),
            page.get_by_role("button", name=name_re),
            page.get_by_text(name_re),
        ],
    ):
        return None
    page.wait_for_timeout(2_000)
    host = _active(page)
    for extra in page.context.pages:
        if extra not in pages_before:
            extra.bring_to_front()
            host = extra
            break
    if _is_forbidden(host):
        LOG.warning("servicio %s: 403 / sin permiso", query)
        return None
    _elegir_empresa(host, job)
    return host


def _fill_date(loc: Locator, value: str) -> None:
    loc.click()
    loc.fill("")
    loc.fill(value)
    loc.press("Tab")


def _completar_fechas(scope: Any, desde: str, hasta: str) -> bool:
    ddmmyyyy_d = iso_to_ddmmyyyy(desde)
    ddmmyyyy_h = iso_to_ddmmyyyy(hasta)
    desde_loc = scope.locator(
        "input[id*='fecha' i][id*='desde' i], input[name*='desde' i], "
        "input[placeholder*='desde' i], input[aria-label*='desde' i]"
    )
    hasta_loc = scope.locator(
        "input[id*='fecha' i][id*='hasta' i], input[name*='hasta' i], "
        "input[placeholder*='hasta' i], input[aria-label*='hasta' i]"
    )
    ok = False
    if desde_loc.count():
        _fill_date(desde_loc.first, ddmmyyyy_d)
        ok = True
    else:
        labeled = scope.get_by_label(re.compile(r"desde", re.I))
        if labeled.count():
            _fill_date(labeled.first, ddmmyyyy_d)
            ok = True
    if hasta_loc.count():
        _fill_date(hasta_loc.first, ddmmyyyy_h)
        ok = True
    else:
        labeled = scope.get_by_label(re.compile(r"hasta", re.I))
        if labeled.count():
            _fill_date(labeled.first, ddmmyyyy_h)
            ok = True
    if not ok:
        texts = scope.get_by_role("textbox")
        if texts.count() >= 2:
            _fill_date(texts.nth(0), ddmmyyyy_d)
            _fill_date(texts.nth(1), ddmmyyyy_h)
            ok = True
    return ok


def guardar_descarga(
    page: Page,
    clickable: Locator,
    tmp_dl: Path,
    timeout_ms: int = 25_000,
) -> list[Path]:
    """Click + Playwright download, o archivos nuevos en tmp_dl."""
    tmp_dl.mkdir(parents=True, exist_ok=True)
    before = {p.name for p in tmp_dl.glob("*")}
    found: list[Path] = []
    try:
        with page.expect_download(timeout=timeout_ms) as info:
            clickable.click(timeout=8_000, force=True)
        dl = info.value
        name = Path(dl.suggested_filename or "download.bin").name
        dest = tmp_dl / name
        n = 1
        while dest.exists():
            dest = tmp_dl / f"{dest.stem}_{n}{dest.suffix}"
            n += 1
        dl.save_as(str(dest))
        if dest.exists() and dest.stat().st_size > 50:
            found.append(dest)
    except Exception:
        try:
            clickable.click(timeout=8_000, force=True)
        except Exception:
            pass
        page.wait_for_timeout(2_200)
    if not found:
        page.wait_for_timeout(800)
        for f in tmp_dl.glob("*"):
            if f.name in before or not f.is_file():
                continue
            if f.stat().st_size > 50:
                found.append(f)
    return found


def copiar_si_nuevo(src: Path, dest: Path) -> str | None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 200:
        return str(dest)
    try:
        shutil.copyfile(src, dest)
    except OSError as exc:
        LOG.warning("no pude copiar %s → %s: %s", src, dest, exc)
        return None
    if dest.exists() and dest.stat().st_size > 50:
        return str(dest)
    return None
