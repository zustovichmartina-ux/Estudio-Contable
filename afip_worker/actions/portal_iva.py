# -*- coding: utf-8 -*-
"""Live: Portal IVA (compras/ventas CSV·PDF) y fallback Mis Comprobantes.

No presenta DDJJ. Si ARCA cambió el SPA, deja screenshot + fallback manual.
"""
from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeout

from ..browser import chrome_context
from ..jobs import Job, jobs_root, limpiar_ruta
from ..naming import (
    dest_portal_iva,
    etiqueta_mm_yyyy,
    meses_del_periodo,
    nombre_fcc_compra,
    nombre_mis_comprobantes,
    nombre_portal_iva_csv,
    nombre_portal_iva_libro,
    safe_name,
)
from .result import ActionResult
from .session import (
    PORTAL_URL,
    _active,
    _click_first,
    _completar_fechas,
    _login_afip,
    _screenshot,
    abrir_servicio,
    copiar_si_nuevo,
    guardar_descarga,
)

LOG = logging.getLogger("afip_worker.portal_iva")

MANUAL_FALLBACK = (
    "Fallback manual en RECEPCION (Chrome, sin claves en la web): "
    "Portal IVA → período → Libro IVA Compras/Ventas → Importar comprobantes desde ARCA → CSV "
    "(si viene ZIP, descomprimirlo). "
    "Si no está el tile: Mis Comprobantes → Recibidos/Emitidos → XLS/CSV."
)

_RE_PORTAL_IVA = re.compile(r"Portal IVA", re.I)
_RE_MIS_CMP = re.compile(r"^Mis [Cc]omprobantes$|Mis [Cc]omprobantes", re.I)
_RE_NUEVA_DDJJ = re.compile(
    r"Nueva declaraci[oó]n jurada|Nueva presentaci[oó]n", re.I
)
_RE_DDJJ_PRESENTADAS = re.compile(
    r"declaraciones juradas presentadas|Mis presentaciones|presentaciones realizadas",
    re.I,
)
_RE_LIBRO_IVA = re.compile(
    r"Libro de IVA|Libro IVA|Registraci[oó]n electr[oó]nica",
    re.I,
)
_RE_COMPRAS = re.compile(r"^Compras$|Libro IVA Compras|Libro de IVA Compras", re.I)
_RE_VENTAS = re.compile(r"^Ventas$|Libro IVA Ventas|Libro de IVA Ventas", re.I)
_RE_IMPORTAR = re.compile(
    r"Importar comprobantes desde ARCA|Importar desde ARCA|^Importar$",
    re.I,
)
_RE_EXPORT = re.compile(
    r"^CSV$|^XLSX?$|Exportar CSV|Descargar CSV|Excel|Exportar",
    re.I,
)
_RE_PDF_LIBRO = re.compile(
    r"vista previa|descargar.*pdf|acuse de presentaci|planilla",
    re.I,
)
_RE_RECIBIDOS = re.compile(r"^Recibidos$", re.I)
_RE_EMITIDOS = re.compile(r"^Emitidos$", re.I)
_RE_BUSCAR = re.compile(r"^Buscar$|^Consultar$|^Filtrar$", re.I)
_RE_PRESENTAR = re.compile(
    r"presentar libro|presentar declaraci|confirmar presentaci|presentar ddjj",
    re.I,
)
_RE_NO_PRESENTAR_OK = re.compile(r"acuse|vista previa|descargar", re.I)
_RE_PV_NRO = re.compile(r"(\d{1,5})\s*[-/]\s*(\d{1,8})")

_LADOS = ("compras", "ventas")


def run_bajar_portal_iva_live(job: Job) -> ActionResult:
    dest = Path(limpiar_ruta(str(job.params.get("ruta_destino") or "")))
    if not str(dest):
        return ActionResult(ok=False, message="params.ruta_destino es obligatorio")
    try:
        dest.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return ActionResult(ok=False, message=f"No se puede crear destino {dest}: {exc}")

    desde = str(job.params.get("periodo_desde") or "")[:10]
    hasta = str(job.params.get("periodo_hasta") or job.params.get("periodo") or "")[:10]
    if not desde:
        periodo = str(job.params.get("periodo") or "")[:7]
        if len(periodo) == 7 and periodo[4] == "-":
            desde = f"{periodo}-01"
            hasta = desde
    if not desde or not hasta:
        return ActionResult(
            ok=False,
            message="periodo_desde y periodo_hasta son obligatorios (o params.periodo AAAA-MM)",
        )

    tmp_dl = jobs_root() / "running" / f"{job.id}_dl"
    tmp_dl.mkdir(parents=True, exist_ok=True)
    saved: list[str] = []
    notas: list[str] = []
    fuentes: list[str] = []

    try:
        with chrome_context(download_dir=tmp_dl) as (_ctx, page, _attached):
            auth = _login_afip(page, job)
            if auth is not None:
                _screenshot(page, job, "login")
                return auth

            iva_host = _abrir_portal_iva(page, job)
            if iva_host is None:
                notas.append("No encontré el tile Portal IVA (¿servicio adherido a este CUIT?).")
                _screenshot(_active(page), job, "portal_iva")
            else:
                fuentes.append("portal_iva")
                err = _recorrer_portal_iva(iva_host, job, dest, tmp_dl, desde, hasta, saved)
                if err:
                    notas.append(err)

            if not saved:
                mis_host = _abrir_mis_comprobantes(page, job)
                if mis_host is None:
                    notas.append("Tampoco encontré Mis Comprobantes.")
                    _screenshot(_active(page), job, "mis_comprobantes")
                else:
                    fuentes.append("mis_comprobantes")
                    err = _recorrer_mis_comprobantes(
                        mis_host, job, dest, tmp_dl, desde, hasta, saved
                    )
                    if err:
                        notas.append(err)
    except PlaywrightTimeout as exc:
        return ActionResult(ok=False, message=f"Timeout AFIP: {exc}. {MANUAL_FALLBACK}")
    except Exception as exc:
        return ActionResult(ok=False, message=f"{type(exc).__name__}: {exc}. {MANUAL_FALLBACK}")

    extra = {
        "fuente": "+".join(fuentes) if fuentes else "",
        "manual_fallback": MANUAL_FALLBACK,
    }
    if not saved:
        detail = "; ".join(notas[:4]) if notas else "No bajó ningún archivo de Portal IVA"
        return ActionResult(
            ok=False,
            message=f"{detail}. {MANUAL_FALLBACK}",
            extra=extra,
        )
    extra_msg = f" ({'; '.join(notas[:2])})" if notas else ""
    return ActionResult(
        ok=True,
        message=(
            f"Portal IVA {desde}→{hasta}: {len(saved)} archivo(s)"
            f"{extra_msg}"
        ),
        files=saved,
        extra=extra,
    )


def inferir_lado_archivo(nombre: str, lado_hint: str = "") -> str:
    hint = (lado_hint or "").strip().lower()
    if hint.startswith("vent") or "emitid" in hint:
        return "ventas"
    if hint.startswith("comp") or "recibid" in hint:
        return "compras"
    low = (nombre or "").lower()
    if any(k in low for k in ("venta", "emitid", "sales")):
        return "ventas"
    return "compras"


def es_boton_presentar(texto: str) -> bool:
    t = (texto or "").strip().lower()
    if _RE_NO_PRESENTAR_OK.search(t):
        return False
    return bool(_RE_PRESENTAR.search(t))


def nombre_destino_descarga(
    original: str,
    *,
    lado: str,
    periodo_mm_yyyy: str,
    cliente: str,
    desde: str = "",
    hasta: str = "",
) -> str:
    name = Path(original).name
    ext = Path(name).suffix.lower().lstrip(".") or "bin"
    lado_n = inferir_lado_archivo(name, lado)
    if ext == "csv":
        return nombre_portal_iva_csv(lado_n)
    if ext in {"xls", "xlsx"}:
        tipo = "Recibidos" if lado_n == "compras" else "Emitidos"
        return nombre_mis_comprobantes(
            cliente=cliente,
            tipo=tipo,
            desde=desde or periodo_mm_yyyy,
            hasta=hasta or periodo_mm_yyyy,
            ext=ext,
        )
    if ext == "pdf":
        low = name.lower()
        if any(k in low for k in ("libro", "acuse", "vista", "f2051", "ddjj", "iva")):
            return nombre_portal_iva_libro(lado_n, periodo_mm_yyyy, "pdf")
        return f"{safe_name(cliente or 'Cliente', 40)} PortalIVA {lado_n} {periodo_mm_yyyy}.pdf"
    if ext == "zip":
        return f"PortalIVA_{lado_n}_{periodo_mm_yyyy}.zip"
    return name


def materializar_descarga(
    src: Path,
    dest: Path,
    *,
    lado: str,
    periodo_mm_yyyy: str,
    cliente: str,
    desde: str = "",
    hasta: str = "",
) -> list[str]:
    out: list[str] = []
    if src.suffix.lower() == ".zip":
        try:
            with zipfile.ZipFile(src) as zf:
                for info in zf.infolist():
                    if info.is_dir() or info.filename.endswith("/"):
                        continue
                    inner = Path(info.filename).name
                    if inner.startswith(".") or inner.startswith("__"):
                        continue
                    target_name = nombre_destino_descarga(
                        inner,
                        lado=lado,
                        periodo_mm_yyyy=periodo_mm_yyyy,
                        cliente=cliente,
                        desde=desde,
                        hasta=hasta,
                    )
                    target = dest / target_name
                    target.parent.mkdir(parents=True, exist_ok=True)
                    if target.exists() and target.stat().st_size > 0:
                        out.append(str(target))
                        continue
                    target.write_bytes(zf.read(info))
                    if target.exists() and target.stat().st_size > 0:
                        out.append(str(target))
            return out
        except zipfile.BadZipFile:
            LOG.warning("ZIP inválido %s; lo copio entero", src)
    target = dest / nombre_destino_descarga(
        src.name,
        lado=lado,
        periodo_mm_yyyy=periodo_mm_yyyy,
        cliente=cliente,
        desde=desde,
        hasta=hasta,
    )
    copied = copiar_si_nuevo(src, target)
    if copied:
        out.append(copied)
    return out


def _abrir_portal_iva(page: Page, job: Job) -> Page | None:
    host = abrir_servicio(page, job, query="Portal IVA", name_re=_RE_PORTAL_IVA)
    if host is None:
        return None
    if _en_portal_iva(host):
        return host
    host.wait_for_timeout(2_000)
    host = _active(page)
    return host if _en_portal_iva(host) or not _parece_rcel(host) else None


def _en_portal_iva(page: Page) -> bool:
    try:
        body = page.locator("body").inner_text(timeout=4_000) or ""
    except Exception:
        body = ""
    url = (page.url or "").lower()
    if any(h in url for h in ("portaliva", "/niva", "iva.afip", "iva.arca")):
        return True
    return bool(
        re.search(r"Portal IVA|Libro de IVA|IVA Simple|declaraci[oó]n jurada de IVA", body, re.I)
    )


def _parece_rcel(page: Page) -> bool:
    try:
        body = page.locator("body").inner_text(timeout=2_000) or ""
    except Exception:
        return False
    return bool(re.search(r"Generar Comprobantes", body, re.I)) and "Portal IVA" not in body


def _abrir_mis_comprobantes(page: Page, job: Job) -> Page | None:
    # Volver al portal de servicios si estamos en otra app.
    try:
        page.goto(PORTAL_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1_000)
    except Exception:
        pass
    host = abrir_servicio(page, job, query="Mis Comprobantes", name_re=_RE_MIS_CMP)
    if host is None:
        return None
    host.wait_for_timeout(1_500)
    return _active(page)


def _scope_iva(page: Page) -> Any:
    for frame in page.frames:
        url = (frame.url or "").lower()
        if any(h in url for h in ("iva", "niva", "portaliva", "miscomprobantes", "libro")):
            if "login" not in url:
                return frame
    return page


def _recorrer_portal_iva(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    desde: str,
    hasta: str,
    saved: list[str],
) -> str:
    notas: list[str] = []
    cliente = job.razon_social or job.cuit
    for mes_desde, mes_hasta in meses_del_periodo(desde, hasta):
        tag = etiqueta_mm_yyyy(mes_desde)
        dest_mes = dest_portal_iva(str(dest), mes_desde)
        dest_mes.mkdir(parents=True, exist_ok=True)
        n_before = len(saved)
        if not _entrar_periodo_portal_iva(page, mes_desde):
            notas.append(f"{tag}: no pude abrir el período en Portal IVA")
            continue
        host = _active(page)
        _click_si_no_presentar(host, _RE_LIBRO_IVA)
        host.wait_for_timeout(800)
        for lado in _lados_pedido(job):
            if not _abrir_lado(host, lado):
                notas.append(f"{tag}: no encontré {lado}")
                continue
            host = _active(page)
            n = _exportar_lado_portal_iva(
                host, dest_mes, tmp_dl, lado, tag, cliente, mes_desde, mes_hasta, saved
            )
            LOG.info("Portal IVA %s %s +%s", tag, lado, n)
        if len(saved) == n_before:
            notas.append(f"{tag}: sin archivos")
    if not saved:
        return notas[0] if notas else "Portal IVA no entregó CSV/PDF"
    return "; ".join(notas[:3]) if notas else ""


def _lados_pedido(job: Job) -> tuple[str, ...]:
    raw = str(job.params.get("lados") or job.params.get("lado") or "").strip().lower()
    if not raw:
        return _LADOS
    picked = []
    if "vent" in raw:
        picked.append("ventas")
    if "comp" in raw or "compra" in raw:
        picked.append("compras")
    return tuple(picked) or _LADOS


def _entrar_periodo_portal_iva(page: Page, mes_desde: str) -> bool:
    host = _active(page)
    year, month = mes_desde[:4], mes_desde[5:7]
    label = f"{month}/{year}"
    alt = f"{month}-{year}"
    # Presentaciones ya hechas: más seguro (no crea borrador).
    _click_si_no_presentar(host, _RE_DDJJ_PRESENTADAS)
    host.wait_for_timeout(600)
    if _click_texto(host, (label, alt, f"{year}-{month}")):
        host.wait_for_timeout(1_000)
        return True
    _click_si_no_presentar(host, _RE_NUEVA_DDJJ)
    host.wait_for_timeout(600)
    if _seleccionar_periodo_controles(host, year, month):
        _click_si_no_presentar(host, re.compile(r"^Continuar$|^Ingresar$|^Aceptar$", re.I))
        host.wait_for_timeout(1_000)
        LOG.info("abrí período %s (posible borrador; no presento)", label)
        return True
    if _click_texto(host, (label, alt)):
        host.wait_for_timeout(800)
        return True
    return False


def _seleccionar_periodo_controles(page: Page, year: str, month: str) -> bool:
    scope = _scope_iva(page)
    ok = False
    for sel, value in (
        (scope.get_by_label(re.compile(r"a[nñ]o|year", re.I)), year),
        (scope.get_by_label(re.compile(r"mes|month", re.I)), month),
        (scope.get_by_label(re.compile(r"per[ií]odo", re.I)), f"{month}/{year}"),
    ):
        try:
            if sel.count():
                try:
                    sel.first.select_option(value)
                except Exception:
                    sel.first.fill(value)
                ok = True
        except Exception:
            continue
    if ok:
        return True
    selects = scope.locator("select")
    if selects.count() >= 2:
        try:
            selects.nth(0).select_option(year)
            selects.nth(1).select_option(month)
            return True
        except Exception:
            try:
                selects.nth(0).select_option(month)
                selects.nth(1).select_option(year)
                return True
            except Exception:
                return False
    return False


def _abrir_lado(page: Page, lado: str) -> bool:
    pattern = _RE_VENTAS if lado == "ventas" else _RE_COMPRAS
    return _click_si_no_presentar(page, pattern)


def _exportar_lado_portal_iva(
    page: Page,
    dest: Path,
    tmp_dl: Path,
    lado: str,
    periodo_mm_yyyy: str,
    cliente: str,
    desde: str,
    hasta: str,
    saved: list[str],
) -> int:
    n0 = len(saved)
    scope = _scope_iva(_active(page))
    _click_si_no_presentar(page, _RE_IMPORTAR)
    page.wait_for_timeout(600)
    scope = _scope_iva(_active(page))
    _bajar_exports(
        page,
        scope,
        dest,
        tmp_dl,
        lado,
        periodo_mm_yyyy,
        cliente,
        desde,
        hasta,
        saved,
        extra_re=_RE_PDF_LIBRO,
    )
    return len(saved) - n0


def _recorrer_mis_comprobantes(
    page: Page,
    job: Job,
    dest: Path,
    tmp_dl: Path,
    desde: str,
    hasta: str,
    saved: list[str],
) -> str:
    notas: list[str] = []
    cliente = job.razon_social or job.cuit
    host = _active(page)
    for mes_desde, mes_hasta in meses_del_periodo(desde, hasta):
        tag = etiqueta_mm_yyyy(mes_desde)
        dest_mes = dest_portal_iva(str(dest), mes_desde)
        dest_mes.mkdir(parents=True, exist_ok=True)
        for lado, tab_re in (("compras", _RE_RECIBIDOS), ("ventas", _RE_EMITIDOS)):
            if lado not in _lados_pedido(job):
                continue
            if not _click_si_no_presentar(host, tab_re):
                notas.append(f"{tag}: sin pestaña {lado}")
                continue
            host = _active(page)
            scope = _scope_iva(host)
            if not _completar_fechas(scope, mes_desde, mes_hasta):
                notas.append(f"{tag}: no pude cargar fechas Mis Comprobantes ({lado})")
                continue
            _click_first(
                host,
                [
                    scope.get_by_role("button", name=_RE_BUSCAR),
                    scope.locator("input[value='Buscar'], input[value='Consultar']"),
                    scope.get_by_text(_RE_BUSCAR),
                ],
            )
            host.wait_for_timeout(1_500)
            scope = _scope_iva(_active(host))
            n = _bajar_exports(
                host,
                scope,
                dest_mes,
                tmp_dl,
                lado,
                tag,
                cliente,
                mes_desde,
                mes_hasta,
                saved,
            )
            _bajar_pdfs_filas(host, scope, dest_mes, tmp_dl, saved)
            LOG.info("Mis Comprobantes %s %s +%s", tag, lado, n)
    if not saved:
        return notas[0] if notas else "Mis Comprobantes no entregó archivos"
    return "; ".join(notas[:3]) if notas else ""


def _bajar_exports(
    page: Page,
    scope: Any,
    dest: Path,
    tmp_dl: Path,
    lado: str,
    periodo_mm_yyyy: str,
    cliente: str,
    desde: str,
    hasta: str,
    saved: list[str],
    extra_re: re.Pattern[str] | None = None,
) -> int:
    n0 = len(saved)
    locators = [
        scope.get_by_role("link", name=_RE_EXPORT),
        scope.get_by_role("button", name=_RE_EXPORT),
        scope.get_by_text(_RE_EXPORT),
        scope.locator("a[title*='CSV' i], a[title*='Excel' i], a[title*='XLS' i], img[alt*='CSV' i]"),
    ]
    if extra_re is not None:
        locators.extend(
            [
                scope.get_by_role("link", name=extra_re),
                scope.get_by_role("button", name=extra_re),
                scope.get_by_text(extra_re),
            ]
        )
    seen: set[str] = set()
    for loc in locators:
        try:
            count = loc.count()
        except Exception:
            continue
        for i in range(min(count, 6)):
            btn = loc.nth(i)
            label = ""
            try:
                if not btn.is_visible():
                    continue
                label = (btn.inner_text(timeout=1_000) or btn.get_attribute("title") or "")[:80]
            except Exception:
                pass
            if es_boton_presentar(label):
                LOG.info("salteo botón presentar: %s", label)
                continue
            key = f"{label}:{i}:{id(loc)}"
            if key in seen:
                continue
            seen.add(key)
            files = guardar_descarga(page, btn, tmp_dl)
            for src in files:
                for path in materializar_descarga(
                    src,
                    dest,
                    lado=lado,
                    periodo_mm_yyyy=periodo_mm_yyyy,
                    cliente=cliente,
                    desde=desde,
                    hasta=hasta,
                ):
                    if path not in saved:
                        saved.append(path)
    return len(saved) - n0


def _bajar_pdfs_filas(
    page: Page,
    scope: Any,
    dest: Path,
    tmp_dl: Path,
    saved: list[str],
) -> int:
    n0 = len(saved)
    rows = scope.locator("table tbody tr, table tr")
    try:
        count = rows.count()
    except Exception:
        return 0
    for i in range(min(count, 80)):
        row = rows.nth(i)
        try:
            text = row.inner_text(timeout=1_500)
        except Exception:
            continue
        if not text.strip() or ("fecha" in text.lower() and "tipo" in text.lower()):
            continue
        icon = _icono_pdf(row)
        if icon is None:
            continue
        pv, nro = _pv_nro_de_fila(text)
        proveedor = _proveedor_de_fila(text)
        if pv is not None and nro is not None:
            filename = nombre_fcc_compra(
                tipo="FCC",
                letra="A",
                pv=pv,
                nro=nro,
                proveedor=proveedor or "COMPROBANTE",
            )
        else:
            filename = f"comprobante_{i + 1:03d}.pdf"
        target = dest / filename
        if target.exists() and target.stat().st_size > 0:
            if str(target) not in saved:
                saved.append(str(target))
            continue
        files = guardar_descarga(page, icon, tmp_dl)
        for src in files:
            if src.suffix.lower() != ".pdf":
                continue
            copied = copiar_si_nuevo(src, target)
            if copied and copied not in saved:
                saved.append(copied)
    return len(saved) - n0


def _icono_pdf(row: Locator) -> Locator | None:
    candidates = [
        row.locator("a[title*='PDF' i], a[title*='Ver' i], a[title*='Imprimir' i]"),
        row.locator("img[alt*='PDF' i], img[title*='PDF' i]"),
        row.get_by_role("link", name=re.compile(r"pdf|ver|imprimir", re.I)),
    ]
    for loc in candidates:
        try:
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def _pv_nro_de_fila(text: str) -> tuple[int | None, int | None]:
    hit = _RE_PV_NRO.search(text.replace(".", ""))
    if not hit:
        return None, None
    try:
        return int(hit.group(1)), int(hit.group(2))
    except ValueError:
        return None, None


def _proveedor_de_fila(text: str) -> str:
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    for ln in lines:
        if re.search(r"[A-Za-zÁÉÍÓÚÑáéíóú]{3,}", ln) and not _RE_PV_NRO.search(ln.replace(".", "")):
            if not re.search(r"fecha|tipo|cuit|total", ln, re.I):
                return safe_name(ln, 40)
    return ""


def _click_si_no_presentar(page: Page, pattern: re.Pattern[str]) -> bool:
    scope = _scope_iva(_active(page))
    candidates = [
        scope.get_by_role("link", name=pattern),
        scope.get_by_role("button", name=pattern),
        scope.get_by_text(pattern),
    ]
    for loc in candidates:
        try:
            if loc.count() == 0:
                continue
            label = (loc.first.inner_text(timeout=800) or "")[:120]
            if es_boton_presentar(label):
                continue
            loc.first.click(timeout=8_000, force=True)
            page.wait_for_timeout(400)
            return True
        except Exception:
            continue
    return False


def _click_texto(page: Page, textos: tuple[str, ...]) -> bool:
    for text in textos:
        loc = page.get_by_text(text, exact=False)
        try:
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=4_000)
                page.wait_for_timeout(400)
                return True
        except Exception:
            continue
    return False
