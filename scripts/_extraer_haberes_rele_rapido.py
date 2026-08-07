"""Completa OCR de rele.pdf con RapidOCR (rápido) y genera Excel de débitos haberes."""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import fitz
import numpy as np
import pandas as pd
from PIL import Image
import io

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from _extraer_haberes_rele import (  # noqa: E402
    CHECKPOINT,
    OUT_DIR,
    OUT_JSON,
    OUT_XLSX,
    PDF_PATH,
    _cargar_checkpoint,
    _guardar_checkpoint,
    _parsear_lineas_haberes,
)

DPI = 150


def _ocr_rapid(pagina_fitz, engine, dpi: int = DPI) -> list[str]:
    pix = pagina_fitz.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    rot = getattr(pagina_fitz, "rotation", 0) or 0
    if rot:
        img = img.rotate(-rot, expand=True)
    w, h = img.size
    if rot == 0 and w > h:
        img = img.rotate(90, expand=True)
    result, _ = engine(np.array(img))
    if not result:
        return []
    filas: dict[int, list[tuple[float, str]]] = {}
    for item in result:
        # RapidOCR: [box, text, score]
        box, texto, _score = item[0], item[1], item[2]
        y = (box[0][1] + box[2][1]) / 2
        x = box[0][0]
        clave = int(y / 16) * 16
        filas.setdefault(clave, []).append((x, str(texto)))
    return [
        " ".join(t for _, t in sorted(filas[y], key=lambda z: z[0])).strip()
        for y in sorted(filas)
        if any(t.strip() for _, t in filas[y])
    ]


def main() -> None:
    from rapidocr_onnxruntime import RapidOCR

    t0 = time.time()
    engine = RapidOCR()
    doc = fitz.open(PDF_PATH)
    total = doc.page_count
    data = _cargar_checkpoint()
    ok = {int(k) for k in data.get("paginas", {}).keys()}
    print(f"PDF {total} págs · ya OK: {len(ok)} · faltan: {total - len(ok)}", flush=True)

    for i in range(total):
        num = i + 1
        if num in ok:
            continue
        lineas = _ocr_rapid(doc[i], engine)
        data.setdefault("paginas", {})[str(num)] = lineas
        nuevos = _parsear_lineas_haberes(num, lineas)
        ok.add(num)
        _guardar_checkpoint(data)
        print(
            f"[{num}/{total}] líneas={len(lineas)} haberes+={len(nuevos)} "
            f"({time.time() - t0:.0f}s)",
            flush=True,
        )

    items = []
    for num_s, lineas in sorted(data.get("paginas", {}).items(), key=lambda x: int(x[0])):
        items.extend(_parsear_lineas_haberes(int(num_s), lineas))
    data["items"] = items
    _guardar_checkpoint(data)
    doc.close()

    df = pd.DataFrame(items)
    if df.empty:
        print("Sin débitos de haberes.")
        return

    df = df.drop_duplicates(subset=["Fecha", "Comprobante", "Debito"], keep="first")
    df["_f"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
    df = df.sort_values(["_f", "Pagina"], kind="stable").drop(columns=["_f"]).reset_index(drop=True)

    total_deb = round(float(df["Debito"].sum()), 2)
    df_out = df.rename(
        columns={
            "Debito": "Débito",
            "Pagina": "Página",
            "Movimiento": "Concepto",
        }
    )
    fechas = pd.to_datetime(df_out["Fecha"], dayfirst=True, errors="coerce")
    resumen = (
        df_out.assign(Mes=fechas.dt.to_period("M").astype(str))
        .groupby("Mes", dropna=False)["Débito"]
        .agg(Cantidad="count", Total="sum")
        .reset_index()
    )
    resumen["Total"] = resumen["Total"].round(2)

    from excel_formato_estudio import guardar_informe_excel

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fmin, fmax = fechas.min(), fechas.max()
    periodo = ""
    if pd.notna(fmin) and pd.notna(fmax):
        periodo = f"Período: {fmin.strftime('%d/%m/%Y')} a {fmax.strftime('%d/%m/%Y')}"

    guardar_informe_excel(
        OUT_XLSX,
        titulo="Débitos de haberes",
        subtitulo="Extracto bancario · Pago haberes",
        periodo=periodo,
        kpis=[
            ("Total débitos", total_deb, "money"),
            ("Cantidad de movimientos", len(df_out), "int"),
        ],
        resumenes=[("Por mes", resumen)],
        detalle=df_out,
        hoja_detalle="Movimientos",
        col_moneda=["Débito", "Total"],
        col_fecha=["Fecha"],
        total_col="Débito",
    )

    OUT_JSON.write_text(
        json.dumps(
            {"total": total_deb, "cantidad": len(df), "items": df.to_dict(orient="records")},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"OK: {len(df)} mov · ${total_deb:,.2f} · {time.time() - t0:.0f}s")
    print(f"Excel: {OUT_XLSX}")
    print(resumen.to_string(index=False))


if __name__ == "__main__":
    main()
