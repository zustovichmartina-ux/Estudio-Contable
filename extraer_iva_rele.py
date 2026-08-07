"""Extrae IVA / percepciones IVA de rele.pdf (Santander escaneado) agrupado por mes."""
from __future__ import annotations

import json
import re
import time
from collections import defaultdict
from pathlib import Path

import fitz
import pandas as pd

from procesador import _obtener_lector_ocr, _lector_ocr_run_lock

PDF = Path(r"c:\Users\recep\Desktop\rele.pdf")
OUT_DIR = Path(r"C:\Users\recep\Desktop\Estudio Contable\logs")
CKPT = OUT_DIR / "rele_iva_ocr_checkpoint.json"
XLSX = OUT_DIR / "rele_iva_por_mes.xlsx"
JSON_OUT = OUT_DIR / "rele_iva_por_mes.json"

# IVA de comisiones (ley 27743) + posibles "Percepcion IVA"
RE_IVA = re.compile(
    r"(?i)\biva\b|percepci[oó]n\s*iva|perc\.?\s*iva|trans\s*fisc|transfisc|27743"
)
RE_EXCL = re.compile(
    r"(?i)impuesto\s+ley\s*25\.?413|retenci[oó]n\s+arba|sellos|iibb(?!\s*iva)"
)
RE_FECHA = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
RE_MONTO = re.compile(r"(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?!\d)")


def _parse_monto(s: str) -> float:
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _mes_de_fecha(f: str) -> str:
    m = RE_FECHA.search(f)
    if not m:
        return "sin_fecha"
    d, mo, y = m.group(1).split("/")
    yi = int(y)
    if yi < 100:
        yi += 2000
    return f"{yi:04d}-{int(mo):02d}"


def _ocr_pagina(page, dpi: int = 150) -> str:
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    from PIL import Image
    import io
    import numpy as np

    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    arr = np.array(img)
    lector = _obtener_lector_ocr()
    with _lector_ocr_run_lock:
        res = lector.readtext(arr, detail=0, paragraph=False)
    return "\n".join(str(x) for x in (res or []))


def _extraer_iva_de_texto(texto: str, pagina: int) -> list[dict]:
    lineas = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    out: list[dict] = []
    fecha_ctx = ""
    for i, ln in enumerate(lineas):
        mf = RE_FECHA.search(ln)
        if mf:
            fecha_ctx = mf.group(1)
        if not RE_IVA.search(ln):
            continue
        if RE_EXCL.search(ln) and not RE_IVA.search(ln):
            continue
        # ventana: línea + siguiente (a veces monto abajo)
        ventana = ln
        if i + 1 < len(lineas):
            ventana = f"{ln} {lineas[i + 1]}"
        montos = RE_MONTO.findall(ventana)
        if not montos:
            continue
        # tomar el primer monto razonable (débito típico IVA)
        monto = _parse_monto(montos[0])
        if monto <= 0:
            continue
        # filtrar montos enormes que son saldos (heurística)
        if monto > 500_000:
            # probar último monto chico
            candidatos = [_parse_monto(x) for x in montos]
            chicos = [c for c in candidatos if 0 < c < 50_000]
            if not chicos:
                continue
            monto = chicos[0]
        fecha = fecha_ctx or (mf.group(1) if mf else "")
        out.append(
            {
                "pagina": pagina,
                "fecha": fecha,
                "mes": _mes_de_fecha(fecha) if fecha else "sin_fecha",
                "descripcion": ln[:200],
                "importe": round(monto, 2),
            }
        )
    return out


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    ckpt = {"pagina_hasta": 0, "items": []}
    if CKPT.exists():
        ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        print(f"RESUME from page {ckpt['pagina_hasta']+1}, items={len(ckpt['items'])}", flush=True)

    doc = fitz.open(str(PDF))
    total = doc.page_count
    print(f"PDF pages={total}", flush=True)
    t0 = time.time()

    # warmup OCR
    _obtener_lector_ocr()
    print("OCR ready", flush=True)

    start = int(ckpt.get("pagina_hasta") or 0)
    for i in range(start, total):
        t1 = time.time()
        try:
            texto = _ocr_pagina(doc[i], dpi=150)
        except Exception as exc:
            print(f"PAGE {i+1}/{total} OCR FAIL {exc}", flush=True)
            texto = ""
        nuevos = _extraer_iva_de_texto(texto, i + 1)
        ckpt["items"].extend(nuevos)
        ckpt["pagina_hasta"] = i + 1
        CKPT.write_text(json.dumps(ckpt, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            f"PAGE {i+1}/{total} +{len(nuevos)} iva  total={len(ckpt['items'])}  "
            f"{time.time()-t1:.1f}s",
            flush=True,
        )

    items = ckpt["items"]
    # dedupe por fecha+desc+importe
    seen = set()
    uniq = []
    for it in items:
        key = (it.get("fecha"), it.get("descripcion"), it.get("importe"))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(it)

    por_mes: dict[str, dict] = defaultdict(lambda: {"cantidad": 0, "total": 0.0})
    for it in uniq:
        m = it["mes"]
        por_mes[m]["cantidad"] += 1
        por_mes[m]["total"] += float(it["importe"])

    resumen = [
        {"mes": k, "cantidad": v["cantidad"], "total": round(v["total"], 2)}
        for k, v in sorted(por_mes.items())
    ]
    payload = {
        "archivo": str(PDF),
        "elapsed_s": round(time.time() - t0, 1),
        "paginas": total,
        "total_general": round(sum(r["total"] for r in resumen), 2),
        "por_mes": resumen,
        "detalle": uniq,
        "nota": (
            "Movimientos de extracto Santander con IVA (incluye "
            "'Iva 21%/10,5% reg de transfisc ley 27743' y 'Percepcion IVA' si aparece)."
        ),
    }
    JSON_OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    df = pd.DataFrame(uniq)
    with pd.ExcelWriter(XLSX, engine="openpyxl") as w:
        pd.DataFrame(resumen).to_excel(w, sheet_name="Por_mes", index=False)
        df.to_excel(w, sheet_name="Detalle", index=False)
    print("DONE", payload["total_general"], XLSX, flush=True)
    for r in resumen:
        print(r, flush=True)


if __name__ == "__main__":
    main()
