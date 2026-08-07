"""Extrae débitos 'Pago haberes' del extracto Santander rele.pdf (OCR + Excel)."""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import fitz
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from procesador import _ocr_pagina_rapida, _limpiar_monto  # noqa: E402

PDF_PATH = Path(r"c:\Users\recep\Downloads\rele.pdf")
OUT_DIR = ROOT / "logs"
CHECKPOINT = OUT_DIR / "rele_haberes_ocr_checkpoint.json"
OUT_XLSX = OUT_DIR / "rele_debitos_haberes.xlsx"
OUT_JSON = OUT_DIR / "rele_debitos_haberes.json"

_RE_FECHA = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
_RE_MONTO = re.compile(r"(-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2})")
_RE_HABERES = re.compile(
    r"(?i)pago\s+haberes|acreditamiento\s+de\s+haberes|debito\s+haberes|"
    r"d[eé]bito\s+de\s+haberes|liquidaci[oó]n\s+de\s+haberes"
)


def _norm_fecha(txt: str) -> str:
    m = _RE_FECHA.search(txt or "")
    if not m:
        return ""
    raw = m.group(1)
    candidatos = [raw]
    # OCR frecuente: 11→71, 14→74 (el 1 se lee como 7)
    partes = raw.split("/")
    if len(partes) == 3 and partes[0].isdigit() and int(partes[0]) > 31 and partes[0].startswith("7"):
        candidatos.append("1" + partes[0][1:] + "/" + partes[1] + "/" + partes[2])
    for cand in candidatos:
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try:
                d = datetime.strptime(cand, fmt).date()
                if 1 <= d.day <= 31 and 1 <= d.month <= 12:
                    return d.strftime("%d/%m/%Y")
            except ValueError:
                continue
    return ""


def _cargar_checkpoint() -> dict:
    if CHECKPOINT.exists():
        return json.loads(CHECKPOINT.read_text(encoding="utf-8"))
    return {"paginas": {}, "items": []}


def _guardar_checkpoint(data: dict) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHECKPOINT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _parsear_lineas_haberes(pagina: int, lineas: list[str]) -> list[dict]:
    items: list[dict] = []
    fecha_contexto = ""
    for i, ln in enumerate(lineas):
        f = _norm_fecha(ln)
        if f and not _RE_HABERES.search(ln):
            # línea solo fecha / fecha + otros movimientos
            fecha_contexto = f
        if not _RE_HABERES.search(ln):
            continue

        fecha = f or fecha_contexto
        if not fecha and i > 0:
            fecha = _norm_fecha(lineas[i - 1]) or fecha_contexto

        montos = [_limpiar_monto(m.group(1)) for m in _RE_MONTO.finditer(ln)]
        # En extracto Santander el débito suele ser el primer monto "grande"
        # tras la descripción; si hay varios, tomar el mayor (saldo suele ser último).
        debito = 0.0
        if montos:
            # Preferir el primero si hay 1-2; si hay 3+ (debito/credito/saldo), el primero
            # si la línea no parece crédito.
            if len(montos) >= 3:
                debito = montos[0]
            elif len(montos) == 2:
                # a veces comprobante numérico no es monto; montos ya filtrados por ,dd
                debito = montos[0]
            else:
                debito = montos[0]

        if debito <= 0:
            # buscar en línea siguiente (OCR partido)
            if i + 1 < len(lineas):
                montos2 = [_limpiar_monto(m.group(1)) for m in _RE_MONTO.finditer(lineas[i + 1])]
                if montos2:
                    debito = montos2[0]

        m_comp = re.search(r"\b(\d{6,10})\b", ln)
        comprobante = m_comp.group(1) if m_comp else ""
        m_ref = re.search(r"(?i)pago\s+haberes\s+(\d+)", ln)
        referencia = m_ref.group(1) if m_ref else ""

        items.append(
            {
                "Pagina": pagina,
                "Fecha": fecha,
                "Comprobante": comprobante,
                "Referencia": referencia,
                "Movimiento": re.sub(r"\s+", " ", ln).strip(),
                "Debito": round(float(debito), 2),
            }
        )
    return items


def main() -> None:
    if not PDF_PATH.exists():
        raise SystemExit(f"No existe {PDF_PATH}")

    doc = fitz.open(PDF_PATH)
    total = doc.page_count
    data = _cargar_checkpoint()
    paginas_ok = {int(k) for k in data.get("paginas", {}).keys()}
    items = list(data.get("items") or [])

    print(f"PDF {PDF_PATH.name}: {total} páginas. Ya OCR: {len(paginas_ok)}")

    for i in range(total):
        num = i + 1
        if num in paginas_ok:
            continue
        lineas = _ocr_pagina_rapida(doc[i], dpi=160)
        data.setdefault("paginas", {})[str(num)] = lineas
        nuevos = _parsear_lineas_haberes(num, lineas)
        items.extend(nuevos)
        data["items"] = items
        paginas_ok.add(num)
        if num % 5 == 0 or nuevos:
            _guardar_checkpoint(data)
            print(
                f"[{num}/{total}] líneas={len(lineas)} haberes+={len(nuevos)} total={len(items)}",
                flush=True,
            )

    # Reparse completo por si mejoramos regex
    items = []
    for num_s, lineas in sorted(data.get("paginas", {}).items(), key=lambda x: int(x[0])):
        items.extend(_parsear_lineas_haberes(int(num_s), lineas))
    data["items"] = items
    _guardar_checkpoint(data)
    doc.close()

    df = pd.DataFrame(items)
    if df.empty:
        print("No se encontraron débitos de haberes.")
        return

    # Deduplicar por fecha+comprobante+debito
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

    fmin = fechas.min()
    fmax = fechas.max()
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
    print(f"OK: {len(df)} movimientos · total ${total_deb:,.2f}")
    print(f"Excel: {OUT_XLSX}")


if __name__ == "__main__":
    main()
