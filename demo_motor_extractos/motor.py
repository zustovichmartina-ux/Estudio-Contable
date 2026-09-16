# -*- coding: utf-8 -*-
"""Parser Galicia (Y + últimos 2 importes) con OCR del estudio si el PDF es escaneo."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from procesador import _paginas_texto_extracto_pdf

CACHE = ROOT / "_cache_extractos_sunny_galicia"

RE_FECHA = re.compile(r"^(\d{2}/\d{2}/\d{2,4})\s+")
RE_FECHA_SOLA = re.compile(r"^(\d{2}/\d{2}/\d{2,4})$")
RE_NUM = re.compile(r"(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+,\d{2})")
RE_SKIP = re.compile(
    r"^(Fecha|Movimientos|Resumen|Total|Los dep|Dispon|Canales|Ingres|"
    r"Llaman|Usted|Al comp|Banco de|Chatea|http|Pagina|Página)",
    re.I,
)
RULES = [
    ("inter-cta", "FCI - Suscripcion", re.compile(r"SUSCRIPCION FIMA|SUSCRIPCION FCI", re.I)),
    ("inter-cta", "FCI - Rescate", re.compile(r"RESCATE FIMA|RESCATE FCI", re.I)),
    ("inter-cta", "Transferencia propia", re.compile(r"TRANSF\. CTAS PROPIAS|TRANSFERENCIA DE CUENTA PROPIA", re.I)),
    ("retencion", "Ret. IIBB - ARBA", re.compile(r"ARBA AUTOM|PAGO.*ARBA|ARBA IIBB", re.I)),
    ("retencion", "ICDB Ley 25.413", re.compile(r"LEY 25\.413|IMP\. DEB\. LEY 25413|IMP\. CRE\. LEY 25413|IMP\.LEY 25413|DEB\.LEY", re.I)),
    ("retencion", "Percepcion IVA", re.compile(r"PERCEP\. IVA|PERCEPCION IVA|PERC\.IVA", re.I)),
    ("impuesto", "IVA banco", re.compile(r"(^|\s)IVA(\s|$)", re.I)),
    ("impuesto", "Pago AFIP", re.compile(r"TRANSF\. AFIP|\bVEP\b|PAGO DE SERVICIOS.*AFIP", re.I)),
    ("egreso", "Comisiones y gastos bancarios", re.compile(r"COMISION SERVICIO|COMISION BANCO|COM\. GESTION|COM\. DEPOSITO", re.I)),
    ("impuesto", "IIBB", re.compile(r"ING\. BRUTOS|IIBB", re.I)),
    ("ingreso", "Transferencia de tercero", re.compile(r"TRANSFERENCIA DE TERCEROS|CREDITO TRANSFERENCIA", re.I)),
    ("egreso", "Pago tarjeta VISA", re.compile(r"PAGO TARJETA VISA|PAGO VISA", re.I)),
    ("egreso", "Cheque / ECHEQ", re.compile(r"ECHEQ|DEPOSITO DE CHEQUE|G\.DE ECHEQ", re.I)),
    ("egreso", "Transferencia a tercero", re.compile(r"TRANSFERENCIA A TERCEROS|TRF INMED PROVEED|TRANSFERENCIAS CASH|SERVICIO PAGO A PROVEEDORES", re.I)),
    ("egreso", "Pago de servicios", re.compile(r"PAGO DE SERVICIOS", re.I)),
    ("ingreso", "Anulacion / devolucion", re.compile(r"ANULACION|DEV\.IMP", re.I)),
]


def pnum(s: str) -> float:
    t = (s or "").strip()
    neg = t.endswith("-") or t.startswith("-")
    t = t.replace("-", "").replace(".", "").replace(",", ".")
    return (float(t) or 0.0) * (-1 if neg else 1)


def classify(desc: str) -> tuple[str, str]:
    for cat, sub, rx in RULES:
        if rx.search(desc or ""):
            return cat, sub
    return "sin-cat", "-"


def lineas_por_y(data: bytes) -> tuple[list[str], int]:
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        lineas: list[str] = []
        chars = 0
        for i in range(doc.page_count):
            chars += len((doc[i].get_text("text") or "").strip())
            d = doc[i].get_text("dict")
            by_y: dict[int, list[tuple[float, str]]] = defaultdict(list)
            for b in d.get("blocks") or []:
                if b.get("type") != 0:
                    continue
                for line in b.get("lines") or []:
                    for sp in line.get("spans") or []:
                        t = (sp.get("text") or "").strip()
                        if not t:
                            continue
                        y = round(sp["bbox"][1])
                        by_y[y].append((sp["bbox"][0], t))
            for y in sorted(by_y):
                txt = " ".join(t for _, t in sorted(by_y[y], key=lambda x: x[0])).strip()
                if txt:
                    lineas.append(txt)
        return lineas, chars
    finally:
        doc.close()


def cache_ocr(data: bytes) -> Path:
    h = hashlib.sha1(data).hexdigest()[:16]
    CACHE.mkdir(parents=True, exist_ok=True)
    return CACHE / f"ocr_{h}.json"


def lineas_ocr(data: bytes) -> list[str]:
    cp = cache_ocr(data)
    if cp.exists():
        payload = json.loads(cp.read_text(encoding="utf-8"))
        return list(payload.get("lineas") or [])
    paginas = _paginas_texto_extracto_pdf(data, dpi_ocr=170, forzar_ocr=True)
    lineas: list[str] = []
    for _n, txt in paginas:
        for ln in str(txt).splitlines():
            s = ln.strip()
            if s:
                lineas.append(s)
    cp.write_text(json.dumps({"lineas": lineas}, ensure_ascii=False), encoding="utf-8")
    return lineas


def parse_lineas(lineas: list[str]) -> list[dict]:
    movs: list[dict] = []
    i = 0
    while i < len(lineas):
        raw = re.sub(r"[ \t]+", " ", lineas[i]).strip()
        i += 1
        if not raw or RE_SKIP.match(raw):
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
        desc = re.sub(r"\b\d{6,}\b", "", desc)
        desc = re.sub(r"\s+", " ", desc).strip()
        extras: list[str] = []
        while i < len(lineas):
            nxt = re.sub(r"[ \t]+", " ", lineas[i]).strip()
            if RE_FECHA.match(nxt) or RE_FECHA_SOLA.match(nxt) or RE_SKIP.match(nxt):
                break
            if nxt.lower().startswith("total") or re.search(r"p[aá]gina", nxt, re.I):
                break
            extras.append(nxt)
            i += 1
        det = " ".join(extras)
        full = f"{desc} {det}".strip()
        if len(desc) < 3 or (credito < 0.005 and debito < 0.005):
            continue
        if desc.lower() == "total":
            continue
        cat, sub = classify(full)
        movs.append(
            {
                "fecha": mf.group(1),
                "descripcion": desc[:120],
                "detalle": det[:160],
                "credito": round(credito, 2),
                "debito": round(debito, 2),
                "monto": round(credito - debito, 2),
                "saldo": round(abs(saldo), 2),
                "categoria": cat,
                "sub": sub,
            }
        )
    return movs


def parse_bloques(lineas: list[str]) -> list[dict]:
    movs: list[dict] = []
    i = 0
    while i < len(lineas):
        raw = re.sub(r"[ \t]+", " ", lineas[i]).strip()
        i += 1
        mf = RE_FECHA_SOLA.match(raw)
        if not mf:
            continue
        extras: list[str] = []
        montos: list[float] = []
        while i < len(lineas) and len(montos) < 2:
            nxt = re.sub(r"[ \t]+", " ", lineas[i]).strip()
            if RE_FECHA_SOLA.match(nxt) and (montos or extras):
                break
            if RE_SKIP.match(nxt) or nxt.lower().startswith("total"):
                break
            nums = RE_NUM.findall(nxt)
            if nums and re.fullmatch(r"-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+,\d{2}", nxt.replace(" ", "").replace("$", "")):
                montos.append(pnum(nxt))
                i += 1
                continue
            extras.append(nxt)
            i += 1
        if len(montos) < 2:
            continue
        valor, saldo = montos[0], abs(montos[1])
        desc = " ".join(extras).strip()
        if len(desc) < 3:
            continue
        credito = valor if valor >= 0 else 0.0
        debito = abs(valor) if valor < 0 else 0.0
        cat, sub = classify(desc)
        movs.append(
            {
                "fecha": mf.group(1),
                "descripcion": desc[:120],
                "detalle": "",
                "credito": round(credito, 2),
                "debito": round(debito, 2),
                "monto": round(credito - debito, 2),
                "saldo": round(saldo, 2),
                "categoria": cat,
                "sub": sub,
            }
        )
    return movs


def score_cadena(movs: list[dict]) -> int:
    ok = 0
    for a, b in zip(movs, movs[1:]):
        esperado = round(float(a["saldo"]) + float(b["monto"]), 2)
        if abs(esperado - float(b["saldo"])) <= 0.08:
            ok += 1
    return ok


def elegir(cands: list[list[dict]]) -> list[dict]:
    mejor: list[dict] = []
    sc = (-1, -1)
    for c in cands:
        t = (score_cadena(c), len(c))
        if t > sc:
            sc = t
            mejor = c
    return mejor


def parsear_pdf(data: bytes, nombre: str = "") -> dict:
    lineas_nat, chars = lineas_por_y(data)
    uso_ocr = chars < 50
    lineas = lineas_ocr(data) if uso_ocr else lineas_nat
    movs = elegir([parse_lineas(lineas), parse_bloques(lineas)])
    cred = round(sum(m["credito"] for m in movs), 2)
    deb = round(sum(m["debito"] for m in movs), 2)
    cadena = score_cadena(movs)
    return {
        "archivo": nombre,
        "ocr": uso_ocr,
        "chars_nativos": chars,
        "movimientos": movs,
        "n": len(movs),
        "cadena_ok": cadena,
        "cadena_total": max(len(movs) - 1, 0),
        "creditos": cred,
        "debitos": deb,
        "saldo_final": movs[-1]["saldo"] if movs else None,
    }
