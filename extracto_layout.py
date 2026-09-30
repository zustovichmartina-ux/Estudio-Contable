# -*- coding: utf-8 -*-
"""Parseo de extractos AR por coordenada Y: últimos 2 números = importe + saldo."""
from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from typing import Any

import fitz

RE_FECHA = re.compile(r"^(\d{2}/\d{2}/\d{2,4})\s+")
RE_FECHA_SOLA = re.compile(r"^(\d{2}/\d{2}/\d{2,4})$")
RE_NUM = re.compile(r"(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})-?|-?\d+,\d{2}-?)")
RE_SKIP = re.compile(
    r"^(Fecha|Movimientos|Resumen|Total|Los dep|Dispon|Canales|Ingres|"
    r"Llaman|Usted|Al comp|Banco de|Chatea|http|Pagina|Página)",
    re.I,
)

# Pie del PDF (todos los bancos): totales de impuestos / legal. No son movimientos.
# Galicia pega este recap a la fecha del último día; hay que matchearlo en cualquier parte del texto.
_MARCAS_FIN_CUERPO_EXTRACTO = (
    "detalle impositivo",
    "tipo de impuesto",
    "totales de retencion",
    "totales mensuales de retencion",
    "totales de percepcion",
    "informacion impositiva",
    "resumen impositivo",
    "resumen de impuestos",
    "detalle de impuestos",
    "consolidado de retencion",
    "consolidado de retenciones",
    "consolidado importe",
    "consolidado de impuestos",
    "credito fiscal discriminado",
    "el credito fiscal",
    "credito computable como pago a cuenta",
    "pueden ser modificados por devoluciones",
    "los totales mensuales",
    "total retencion impuesto",
    "total mensual retencion",
    "total retencion i.v.a",
    "total retencion iva",
    "total retencion imp",
    "concepto valor tipo",
    "itf-ley25413",
    "itf-ley 25413",
    "s/dec. 380",
    "s/dec 380",
    "cambio de comisiones",
    "tasas de acuerdos",
    "los precios no incluyen iva",
    "fin del extracto",
    "resumen del periodo",
    "total creditos del periodo",
    "total debitos del periodo",
    "canales de atencion",
    "los depositos estan garantizados",
    "los depositos efectuados",
    "los depositos en pesos",
    "garantia de hasta",
    "ponemos en tu conocimiento",
    "saldo total detalle",
    "tasa extraordinaria",
)
_PREFIJOS_FIN_CUERPO_EXTRACTO = (
    "total retencion",
    "total mensual retencion",
    "total impuesto",
    "total mensual",
    "consolidado de",
    "consolidado importe",
    "itf-ley",
)


def _norm_marca_extracto(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t.lower()).strip()


def es_fin_cuerpo_movimientos_extracto(texto: str) -> bool:
    """True si el texto es el recap de impuestos / pie legal: cortar acá."""
    n = _norm_marca_extracto(texto)
    if not n:
        return False
    if any(k in n for k in _MARCAS_FIN_CUERPO_EXTRACTO):
        return True
    if n.startswith("total $") or n.startswith("total$"):
        return True
    if any(n.startswith(p) or f" {p}" in n for p in _PREFIJOS_FIN_CUERPO_EXTRACTO):
        return True
    # Galicia: el consolidado lista débito y crédito de 25.413 en el mismo renglón.
    if (
        ("25413" in n or "25.413" in n or "25 413" in n)
        and "sobre credito" in n
        and "sobre debito" in n
    ):
        return True
    if n.count("total retencion") >= 2 or n.count("total mensual") >= 2:
        return True
    return False


def truncar_paginas_cuerpo_extracto(
    paginas: list[tuple[int, str]],
) -> list[tuple[int, str]]:
    """Deja el extracto del primer movimiento al último: tira el recap y lo que sigue."""
    out: list[tuple[int, str]] = []
    for num, texto in paginas or []:
        lineas: list[str] = []
        corto = False
        for ln in str(texto or "").splitlines():
            if es_fin_cuerpo_movimientos_extracto(ln):
                corto = True
                break
            lineas.append(ln)
        out.append((num, "\n".join(lineas)))
        if corto:
            break
    return out


def pnum(s: str) -> float:
    t = (s or "").strip().replace("$", "").replace(" ", "")
    if not t:
        return 0.0
    neg = t.endswith("-") or t.startswith("-")
    t = t.replace("-", "").replace(".", "").replace(",", ".")
    try:
        return (float(t) or 0.0) * (-1 if neg else 1)
    except ValueError:
        return 0.0


Y_BUCKET = 6


def _lineas_spans_pagina(page, bucket: int = Y_BUCKET) -> list[str]:
    """Junta fecha + concepto + importes de la misma fila visual."""
    by_y: dict[int, list[tuple[float, str]]] = defaultdict(list)
    d = page.get_text("dict") if page is not None else {}
    for b in (d or {}).get("blocks") or []:
        if b.get("type") != 0:
            continue
        for line in b.get("lines") or []:
            for sp in line.get("spans") or []:
                t = (sp.get("text") or "").strip()
                if not t:
                    continue
                y = int(float(sp["bbox"][1]) / bucket) * bucket
                by_y[y].append((float(sp["bbox"][0]), t))
    if not by_y:
        for w in page.get_text("words") or []:
            x0, y0, _x1, _y1, t = w[:5]
            t = str(t or "").strip()
            if not t:
                continue
            y = int(float(y0) / bucket) * bucket
            by_y[y].append((float(x0), t))
    lineas: list[str] = []
    for y in sorted(by_y):
        txt = " ".join(t for _, t in sorted(by_y[y], key=lambda x: x[0])).strip()
        if txt:
            lineas.append(txt)
    return lineas


def paginas_por_y(data: bytes, bucket: int = Y_BUCKET) -> list[tuple[int, str]]:
    """Texto por página reconstruido por Y (Santander overlay / columnas)."""
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        return [
            (i + 1, "\n".join(_lineas_spans_pagina(doc[i], bucket)))
            for i in range(doc.page_count)
        ]
    finally:
        doc.close()


def lineas_por_y(data: bytes) -> tuple[list[str], int]:
    """Reconstruye renglones agrupando spans de la misma altura Y."""
    paginas = paginas_por_y(data)
    lineas = lineas_desde_paginas(paginas)
    chars = sum(len(t) for _, t in paginas)
    return lineas, chars


def lineas_desde_paginas(paginas: list[tuple[int, str]]) -> list[str]:
    out: list[str] = []
    for _n, txt in paginas or []:
        for ln in str(txt or "").splitlines():
            s = ln.strip()
            if s:
                out.append(s)
    return out


def parse_lineas(lineas: list[str]) -> list[dict[str, Any]]:
    movs: list[dict[str, Any]] = []
    i = 0
    while i < len(lineas):
        raw = re.sub(r"[ \t]+", " ", lineas[i]).strip()
        i += 1
        if not raw:
            continue
        if es_fin_cuerpo_movimientos_extracto(raw):
            break
        if RE_SKIP.match(raw):
            continue
        mf = RE_FECHA.match(raw)
        if not mf:
            continue
        rest = raw[mf.end():]
        nums = [pnum(x) for x in RE_NUM.findall(rest)]
        if len(nums) < 1:
            continue
        saldo = nums[-1] if len(nums) >= 2 else 0.0
        valor = nums[-2] if len(nums) >= 2 else nums[0]
        credito = valor if valor >= 0 else 0.0
        debito = abs(valor) if valor < 0 else 0.0
        desc = RE_NUM.sub("", rest)
        desc = re.sub(r"\s+", " ", desc).strip()
        extras: list[str] = []
        while i < len(lineas):
            nxt = re.sub(r"[ \t]+", " ", lineas[i]).strip()
            if es_fin_cuerpo_movimientos_extracto(nxt):
                break
            if RE_FECHA.match(nxt) or RE_FECHA_SOLA.match(nxt) or RE_SKIP.match(nxt):
                break
            if nxt.lower().startswith("total") or re.search(r"p[aá]gina", nxt, re.I):
                break
            extras.append(nxt)
            i += 1
        extra_txt = " ".join(extras)
        extra_txt = re.sub(r"\s+", " ", RE_NUM.sub("", extra_txt)).strip()
        if len(desc) < 14 and extra_txt:
            desc = f"{desc} {extra_txt}".strip()
            det = ""
        else:
            det = extra_txt
        blob = f"{desc} {det}".lower()
        if any(
            k in blob
            for k in (
                "saldo anterior",
                "saldo inicial",
                "saldo al inicio",
                "saldo de apertura",
                "saldo final",
                "saldo al cierre",
            )
        ):
            continue
        if len(desc) < 3 or (credito < 0.005 and debito < 0.005):
            continue
        if desc.lower() == "total":
            continue
        movs.append(
            {
                "fecha": mf.group(1),
                "descripcion": desc[:240],
                "detalle": det[:240],
                "credito": round(credito, 2),
                "debito": round(debito, 2),
                "monto": round(credito - debito, 2),
                "saldo": round(abs(saldo), 2),
            }
        )
    return movs


def parse_bloques(lineas: list[str]) -> list[dict[str, Any]]:
    movs: list[dict[str, Any]] = []
    i = 0
    while i < len(lineas):
        raw = re.sub(r"[ \t]+", " ", lineas[i]).strip()
        i += 1
        if es_fin_cuerpo_movimientos_extracto(raw):
            break
        mf = RE_FECHA_SOLA.match(raw)
        if not mf:
            continue
        extras: list[str] = []
        montos: list[float] = []
        while i < len(lineas) and len(montos) < 2:
            nxt = re.sub(r"[ \t]+", " ", lineas[i]).strip()
            if RE_FECHA_SOLA.match(nxt) and (montos or extras):
                break
            if es_fin_cuerpo_movimientos_extracto(nxt):
                break
            if RE_SKIP.match(nxt) or nxt.lower().startswith("total"):
                break
            nums = RE_NUM.findall(nxt)
            if nums and re.fullmatch(
                r"-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+,\d{2}",
                nxt.replace(" ", "").replace("$", ""),
            ):
                montos.append(pnum(nxt))
                i += 1
                continue
            extras.append(nxt)
            i += 1
        if len(montos) < 2:
            continue
        valor, saldo = montos[0], abs(montos[1])
        desc = " ".join(extras).strip()
        desc = re.sub(r"\s+", " ", RE_NUM.sub("", desc)).strip()
        blob = desc.lower()
        if any(
            k in blob
            for k in (
                "saldo anterior",
                "saldo inicial",
                "saldo al inicio",
                "saldo de apertura",
                "saldo final",
            )
        ):
            continue
        if len(desc) < 3:
            continue
        credito = valor if valor >= 0 else 0.0
        debito = abs(valor) if valor < 0 else 0.0
        movs.append(
            {
                "fecha": mf.group(1),
                "descripcion": desc[:240],
                "detalle": "",
                "credito": round(credito, 2),
                "debito": round(debito, 2),
                "monto": round(credito - debito, 2),
                "saldo": round(saldo, 2),
            }
        )
    return movs


def score_cadena(movs: list[dict[str, Any]]) -> int:
    ok = 0
    for a, b in zip(movs, movs[1:]):
        esperado = round(float(a["saldo"]) + float(b["monto"]), 2)
        if abs(esperado - float(b["saldo"])) <= 0.08:
            ok += 1
    return ok


def elegir_mejor(cands: list[list[dict[str, Any]]]) -> list[dict[str, Any]]:
    mejor: list[dict[str, Any]] = []
    sc = (-1, -1)
    for c in cands:
        t = (score_cadena(c), len(c))
        if t > sc:
            sc = t
            mejor = c
    return mejor


def parsear_lineas(lineas: list[str]) -> list[dict[str, Any]]:
    """Elige el parseo (línea completa vs bloques OCR) con mejor cadena de saldos."""
    cands: list[list[dict[str, Any]]] = []
    for fn in (parse_lineas, parse_bloques):
        try:
            cands.append(fn(lineas))
        except Exception:
            continue
    return elegir_mejor(cands) if cands else []
