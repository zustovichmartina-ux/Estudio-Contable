"""ARBA débitos desde cache OCR → Excel formato Estudio."""

from __future__ import annotations

import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from excel_formato_estudio import guardar_informe_excel  # noqa: E402
from procesador import _limpiar_monto  # noqa: E402

CKPT = ROOT / "logs" / "rele_haberes_ocr_checkpoint.json"
OUT = ROOT / "logs" / "ARBA_movimientos.xlsx"

RE_FECHA = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b")
RE_MONTO = re.compile(r"(-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2})")
RE_RET = re.compile(r"(?i)retenci[oó]n\s*arba")
RE_TOTAL = re.compile(r"(?i)total\s+retenci")
RE_REF = re.compile(r"(?i)^\s*arba\s*(plan\s*pag|iibb|autom)")
RE_PAGO = re.compile(r"(?i)pago\s+de\s+servicios")


def norm_fecha(txt: str, fb: date | None = None) -> date | None:
    m = RE_FECHA.search(txt or "")
    if not m:
        return fb
    raw = m.group(1)
    dd, mm, yy = raw.split("/")
    cands = [raw]
    if dd.isdigit() and int(dd) > 31 and dd.startswith("7"):
        cands.append("1" + dd[1:] + "/" + mm + "/" + yy)
    if mm.isdigit() and int(mm) > 12 and mm.startswith("7"):
        cands.append(dd + "/1" + mm[1:] + "/" + yy)
    for cand in cands:
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try:
                return datetime.strptime(cand, fmt).date()
            except ValueError:
                pass
    return fb


def debito(ln: str) -> float:
    ms = [_limpiar_monto(m.group(1)) for m in RE_MONTO.finditer(ln)]
    return float(ms[0]) if ms else 0.0


def extraer(paginas: dict) -> pd.DataFrame:
    rows: list[dict] = []
    for num_s, lineas in sorted(paginas.items(), key=lambda x: int(x[0])):
        pagina = int(num_s)
        fecha_ctx: date | None = None
        for i, ln in enumerate(lineas):
            f = norm_fecha(ln)
            if f:
                fecha_ctx = f

            if RE_RET.search(ln) and not RE_TOTAL.search(ln):
                deb = debito(ln)
                if deb <= 0:
                    continue
                es_u = bool(re.search(r"(?i)alicuota\s*u\b|alicuotau\b", ln))
                concepto = (
                    "Retención ARBA alícuota U" if es_u else "Retención ARBA alícuota"
                )
                rows.append(
                    {
                        "Fecha": norm_fecha(ln, fecha_ctx),
                        "Concepto": concepto,
                        "Detalle": "",
                        "Débito": round(deb, 2),
                        "Página": pagina,
                    }
                )
                continue

            if RE_REF.search(ln) and i > 0:
                prev = lineas[i - 1]
                if not RE_PAGO.search(prev):
                    if i > 1 and RE_PAGO.search(lineas[i - 2]):
                        prev = lineas[i - 2]
                    else:
                        continue
                deb = debito(prev)
                if deb <= 0:
                    continue
                ref = re.sub(r"\s+", " ", ln).strip()
                if re.search(r"(?i)plan\s*pag", ln):
                    concepto = "ARBA — Plan de pagos"
                elif re.search(r"(?i)iibb", ln):
                    concepto = "ARBA — IIBB"
                elif re.search(r"(?i)autom", ln):
                    concepto = "ARBA — Débito automático"
                else:
                    concepto = "ARBA — Pago"
                detalle = re.sub(r"(?i)^\s*arba\s*", "", ref).strip(" :.-")
                rows.append(
                    {
                        "Fecha": norm_fecha(prev, fecha_ctx),
                        "Concepto": concepto,
                        "Detalle": detalle,
                        "Débito": round(deb, 2),
                        "Página": pagina,
                    }
                )

    df = pd.DataFrame(rows)
    df = df.dropna(subset=["Fecha"])
    df = df.drop_duplicates(
        subset=["Fecha", "Concepto", "Detalle", "Débito"], keep="first"
    )
    return df.sort_values(["Fecha", "Página"], kind="stable").reset_index(drop=True)


def main() -> None:
    data = json.loads(CKPT.read_text(encoding="utf-8"))
    df = extraer(data["paginas"])
    df_det = df.copy()
    df_det["Mes"] = pd.to_datetime(df_det["Fecha"]).dt.strftime("%Y-%m")

    por_tipo = (
        df_det.groupby("Concepto")["Débito"]
        .agg(Cantidad="count", Total="sum")
        .reset_index()
        .sort_values("Total", ascending=False)
    )
    por_mes = (
        df_det.groupby("Mes")["Débito"].agg(Cantidad="count", Total="sum").reset_index()
    )

    fmin, fmax = df["Fecha"].min(), df["Fecha"].max()
    guardar_informe_excel(
        OUT,
        titulo="Movimientos ARBA",
        subtitulo="Oftalmología Rele · Extracto Santander · Débitos ARBA",
        periodo=f"Período: {fmin.strftime('%d/%m/%Y')} a {fmax.strftime('%d/%m/%Y')}",
        kpis=[
            ("Total débitos", float(df["Débito"].sum()), "money"),
            ("Cantidad de movimientos", len(df), "int"),
        ],
        resumenes=[("Por tipo", por_tipo), ("Por mes", por_mes)],
        detalle=df[["Fecha", "Concepto", "Detalle", "Débito", "Página"]],
        hoja_detalle="Movimientos",
        col_moneda=["Débito", "Total"],
        col_fecha=["Fecha"],
        total_col="Débito",
    )
    print(f"OK: {len(df)} movimientos · ${df['Débito'].sum():,.2f}")
    print(OUT)


if __name__ == "__main__":
    main()
