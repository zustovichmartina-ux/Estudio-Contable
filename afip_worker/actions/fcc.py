# -*- coding: utf-8 -*-
"""Live: emitir FCC desde plantilla Excel en Comprobantes en Línea (RCEL)."""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeout

from ..browser import chrome_context
from ..jobs import Job, jobs_root, limpiar_ruta
from ..naming import nombre_fcc, safe_name
from .comprobantes import _abrir_comprobantes_en_linea
from .result import ActionResult
from .session import (
    _active,
    _click_first,
    _login_afip,
    _screenshot,
)

LOG = logging.getLogger("afip_worker.fcc")

_RE_GENERAR = re.compile(r"Generar Comprobantes", re.I)
_RE_CONFIRMAR = re.compile(r"^Confirmar$|^Aceptar$|^Emitir$|Obtener CAE", re.I)
_RE_CONTINUAR = re.compile(r"^Continuar$|^Siguiente$|^Aceptar datos", re.I)
_RE_CAE = re.compile(r"\bCAE\b[:\s]*(\d{10,14})", re.I)

_COL_CUIT = ("cuit", "cuit receptor", "nro doc", "documento", "doc receptor", "cuit_receptor")
_COL_RAZON = ("razon", "razón", "denominacion", "denominación", "receptor", "cliente")
_COL_PV = ("pv", "punto de venta", "pto", "punto_venta")
_COL_TIPO = ("tipo", "letra", "tipo cbte", "comprobante")
_COL_CONCEPTO = ("concepto", "descripcion", "descripción", "detalle", "item")
_COL_IMPORTE = ("importe", "neto", "monto", "total", "precio")
_COL_CANT = ("cantidad", "cant", "qty")
_COL_ALIC = ("alicuota", "alícuota", "iva", "% iva")


def run_emitir_fcc_live(job: Job) -> ActionResult:
    dest = Path(limpiar_ruta(str(job.params.get("ruta_destino") or "")))
    if not str(dest):
        return ActionResult(ok=False, message="params.ruta_destino es obligatorio")
    try:
        dest.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return ActionResult(ok=False, message=f"No se puede crear destino {dest}: {exc}")

    plantilla = Path(str(job.params.get("plantilla_excel") or ""))
    if not plantilla.is_file():
        return ActionResult(ok=False, message="Falta la plantilla Excel (sin claves).")
    filas = _leer_plantilla(plantilla)
    if not filas:
        return ActionResult(ok=False, message="La plantilla no tiene filas para emitir.")

    tmp_dl = jobs_root() / "running" / f"{job.id}_dl"
    tmp_dl.mkdir(parents=True, exist_ok=True)
    saved: list[str] = []
    errores: list[str] = []

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
                    message="No encontré Comprobantes en Línea para emitir FCC.",
                )
            host = _active(page)
            if not _click_first(
                host,
                [
                    host.get_by_role("link", name=_RE_GENERAR),
                    host.get_by_role("button", name=_RE_GENERAR),
                    host.get_by_text(_RE_GENERAR),
                ],
            ):
                _screenshot(host, job, "menu")
                return ActionResult(
                    ok=False,
                    message="Entré a RCEL pero no encontré Generar Comprobantes.",
                )
            host = _active(page)
            for i, fila in enumerate(filas, start=1):
                try:
                    ok_msg = _emitir_una(host, job, fila, dest, tmp_dl, saved)
                    if ok_msg:
                        LOG.info("FCC fila %s: %s", i, ok_msg)
                    else:
                        errores.append(f"fila {i}: no pude emitir")
                        _screenshot(host, job, f"fcc_{i}")
                except Exception as exc:
                    errores.append(f"fila {i}: {type(exc).__name__}: {exc}")
                    _screenshot(_active(page), job, f"fcc_{i}")
                host = _active(page)
    except PlaywrightTimeout as exc:
        return ActionResult(ok=False, message=f"Timeout AFIP: {exc}")
    except Exception as exc:
        return ActionResult(ok=False, message=f"{type(exc).__name__}: {exc}")

    extra = {"errores": errores, "filas": len(filas)}
    if not saved:
        return ActionResult(
            ok=False,
            message="No se emitió ninguna FCC. " + (errores[0] if errores else "Revisá la plantilla y el punto de venta."),
            extra=extra,
        )
    nota = f" · {len(errores)} fila(s) con error" if errores else ""
    return ActionResult(
        ok=True,
        message=f"Emití {len(saved)} FCC{nota}",
        files=saved,
        extra=extra,
    )


def _norm_col(name: str) -> str:
    t = str(name or "").strip().lower()
    t = t.replace("_", " ")
    t = re.sub(r"\s+", " ", t)
    return t


def _celda(row: dict[str, Any], aliases: tuple[str, ...]) -> str:
    keys = {_norm_col(k): k for k in row}
    for alias in aliases:
        if alias in keys:
            val = row.get(keys[alias])
            if val is None or (isinstance(val, float) and pd.isna(val)):
                return ""
            return str(val).strip()
    return ""


def _leer_plantilla(path: Path) -> list[dict[str, Any]]:
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception:
        df = pd.read_csv(path, dtype=str)
    df = df.fillna("")
    if df.empty:
        return []
    out: list[dict[str, Any]] = []
    for rec in df.to_dict(orient="records"):
        plain = {str(k): v for k, v in rec.items()}
        if not any(str(v).strip() for v in plain.values()):
            continue
        out.append(plain)
    return out


def _emitir_una(
    page: Page,
    job: Job,
    fila: dict[str, Any],
    dest: Path,
    tmp_dl: Path,
    saved: list[str],
) -> str:
    cuit = re.sub(r"\D", "", _celda(fila, _COL_CUIT))
    razon = _celda(fila, _COL_RAZON)
    pv = _celda(fila, _COL_PV)
    tipo = _celda(fila, _COL_TIPO) or "C"
    concepto = _celda(fila, _COL_CONCEPTO)
    importe = _celda(fila, _COL_IMPORTE)
    cant = _celda(fila, _COL_CANT) or "1"
    if pv:
        _elegir_punto_venta(page, pv)
    if tipo:
        _click_first(
            page,
            [
                page.get_by_role("link", name=re.compile(rf"Factura\s*{re.escape(tipo[0])}", re.I)),
                page.get_by_text(re.compile(rf"Factura\s*{re.escape(tipo[0])}", re.I)),
            ],
        )
        page.wait_for_timeout(600)
    if cuit:
        _fill_label(page, r"CUIT|Doc\.|Documento|Nro\.\s*Doc", cuit)
    if razon:
        _fill_label(page, r"Raz[oó]n|Denominaci[oó]n|Apellido|Nombre", razon)
    _click_first(
        page,
        [
            page.get_by_role("button", name=_RE_CONTINUAR),
            page.get_by_text(_RE_CONTINUAR),
        ],
    )
    page.wait_for_timeout(800)
    if concepto:
        _fill_label(page, r"Descripci[oó]n|Detalle|Concepto|Producto", concepto)
    if cant:
        _fill_label(page, r"Cantidad", cant)
    if importe:
        _fill_label(page, r"Precio|Importe|Unitario|Neto", importe)
    _click_first(
        page,
        [
            page.get_by_role("button", name=_RE_CONTINUAR),
            page.get_by_text(_RE_CONTINUAR),
        ],
    )
    page.wait_for_timeout(800)
    _click_first(
        page,
        [
            page.get_by_role("button", name=_RE_CONFIRMAR),
            page.get_by_text(_RE_CONFIRMAR),
        ],
    )
    page.wait_for_timeout(2_000)
    body = ""
    try:
        body = page.locator("body").inner_text(timeout=5_000)
    except Exception:
        pass
    cae = ""
    hit = _RE_CAE.search(body or "")
    if hit:
        cae = hit.group(1)
    nro = _nro_emitido(body)
    pv_n = _int_or(pv, 1)
    nro_n = nro or (len(saved) + 1)
    filename = nombre_fcc(cliente=job.razon_social or razon or "Cliente", pv=pv_n, nro=nro_n)
    dest_path = dest / filename
    if dest_path.exists() and dest_path.stat().st_size > 200:
        saved.append(str(dest_path))
        return f"ya estaba {dest_path.name}"
    pdf_link = page.get_by_role("link", name=re.compile(r"pdf|imprimir|ver comprobante", re.I))
    if pdf_link.count():
        try:
            with page.expect_download(timeout=20_000) as info:
                pdf_link.first.click(timeout=8_000, force=True)
            dl = info.value
            dl.save_as(str(dest_path))
        except Exception:
            pass
    if dest_path.exists() and dest_path.stat().st_size > 200:
        saved.append(str(dest_path))
        return f"CAE {cae}" if cae else dest_path.name
    if cae:
        dest_path.write_text(f"CAE {cae}\n{body[:800]}", encoding="utf-8")
        saved.append(str(dest_path.with_suffix(".txt")))
        return f"CAE {cae} (sin PDF)"
    return ""


def _elegir_punto_venta(page: Page, pv: str) -> None:
    digits = re.sub(r"\D", "", pv)
    if not digits:
        return
    _click_first(
        page,
        [
            page.get_by_role("link", name=re.compile(rf"\b{digits}\b")),
            page.get_by_text(re.compile(rf"Punto de Venta\s*{digits}", re.I)),
        ],
    )
    sel = page.get_by_label(re.compile(r"punto de venta", re.I))
    if sel.count():
        try:
            sel.first.select_option(label=re.compile(digits))
        except Exception:
            try:
                sel.first.select_option(digits)
            except Exception:
                pass
    page.wait_for_timeout(500)


def _fill_label(page: Page, label_re: str, value: str) -> None:
    loc = page.get_by_label(re.compile(label_re, re.I))
    if loc.count() == 0:
        return
    try:
        loc.first.click(force=True)
        loc.first.fill(value)
        loc.first.press("Tab")
    except Exception:
        return


def _nro_emitido(body: str) -> int:
    hit = re.search(r"(?:Nro|N[úu]mero|Comprobante)\s*[:\s]*(\d{1,8})", body or "", re.I)
    if not hit:
        return 0
    try:
        return int(hit.group(1))
    except ValueError:
        return 0


def _int_or(raw: str, default: int) -> int:
    digits = re.sub(r"\D", "", raw or "")
    if not digits:
        return default
    try:
        return int(digits)
    except ValueError:
        return default
