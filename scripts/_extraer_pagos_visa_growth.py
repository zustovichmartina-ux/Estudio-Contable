# -*- coding: utf-8 -*-
"""Pagos de tarjeta Visa (pesos y USD) — Growth Marketing SA."""
from __future__ import annotations

import io
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
import fitz
import numpy as np
from PIL import Image

try:
    import easyocr
except ImportError:
    easyocr = None

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_formato_estudio import guardar_informe_excel

BANCOS = Path(
    r"T:\CLIENTES\GROWTH MARKETING SA\Cierre 2025 01-07-2025 al 30-06-2026"
    r"\Papeles de Trabajo\Bancos"
)
OUT_T = BANCOS / "Pagos_Visa_TC_Growth_2026.xlsx"
OUT_DESK = Path(r"C:\Users\recep\Desktop\Pagos_Visa_TC_Growth_2026.xlsx")

RE_PAGO = re.compile(
    r"^(?P<fecha>\d{2}-\d{2}-\d{2})\s+SU PAGO EN (?P<moneda>PESOS|USD)\s+"
    r"(?P<signo>-)?(?P<importe>\d{1,3}(?:\.\d{3})*,\d{2})\s*$",
    re.I,
)
RE_DEUDA = re.compile(
    r"^(?P<fecha>\d{2}-\d{2}-\d{2})\s+TRANSFERENCIA DEUDA\s+"
    r"(?P<usd>\d{1,3}(?:\.\d{3})*,\d{2})\s+TC\s*(?P<tc>[\d.,]+)\s+"
    r"(?P<ars>\d{1,3}(?:\.\d{3})*,\d{2})",
    re.I,
)
RE_SALDO_ANT = re.compile(
    r"SALDO ANTERIOR\s+(?P<pesos>-?\d{1,3}(?:\.\d{3})*,\d{2})\s+"
    r"(?P<usd>-?\d{1,3}(?:\.\d{3})*,\d{2})",
    re.I,
)
RE_TOTAL = re.compile(
    r"TOTAL A PAGAR\s+(?P<pesos>-?\d{1,3}(?:\.\d{3})*,\d{2})\s+"
    r"(?P<usd>-?\d{1,3}(?:\.\d{3})*,\d{2})",
    re.I,
)
RE_NRO = re.compile(r"Resumen N[°ºo]?\s*(?P<nro>\S+)", re.I)
MESES = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04",
    "MAY": "05", "JUN": "06", "JUL": "07", "AGO": "08",
    "SEP": "09", "OCT": "10", "NOV": "11", "DIC": "12",
}


def _money(s) -> float:
    t = str(s or "").strip().replace("$", "").replace(" ", "")
    if not t:
        return 0.0
    neg = t.startswith("-")
    t = t.lstrip("-")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        val = round(float(t), 2)
    except ValueError:
        return 0.0
    return -val if neg else val


def _fecha(raw: str):
    raw = (raw or "").strip().replace("/", "-")
    m = re.match(r"^(\d{2})-([A-Za-z]{3})-(\d{2,4})$", raw)
    if m:
        mes = MESES.get(m.group(2).upper())
        if mes:
            raw = f"{m.group(1)}-{mes}-{m.group(3)}"
    for fmt in ("%d-%m-%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def _fechas_cabecera(texto: str) -> tuple:
    """Cierre y vencimiento actuales (3ª y 4ª fecha del encabezado Galicia)."""
    for line in texto.splitlines()[:12]:
        fechas = re.findall(
            r"\d{2}-(?:Ene|Feb|Mar|Abr|May|Jun|Jul|Ago|Sep|Oct|Nov|Dic)-\d{2}",
            line,
            flags=re.I,
        )
        if len(fechas) >= 4:
            return _fecha(fechas[2]), _fecha(fechas[3])
    return None, None


def _texto_pdf(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join((pg.extract_text() or "") for pg in pdf.pages)


def _ocr_pagina1(path: Path) -> str:
    """OCR solo página 1 (resumen consolidado) con EasyOCR."""
    if easyocr is None:
        return ""
    doc = fitz.open(str(path))
    try:
        pix = doc[0].get_pixmap(dpi=140)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        lector = easyocr.Reader(["es", "en"], gpu=False, verbose=False)
        filas: dict[int, list[tuple[float, str]]] = {}
        for bbox, txt, _conf in lector.readtext(np.array(img)):
            y = int((bbox[0][1] + bbox[2][1]) / 2 / 14) * 14
            filas.setdefault(y, []).append((bbox[0][0], txt))
        lineas = []
        for y in sorted(filas):
            partes = sorted(filas[y], key=lambda x: x[0])
            lineas.append(" ".join(t for _, t in partes))
        return "\n".join(lineas)
    except Exception:
        return ""
    finally:
        doc.close()


def pdfs_visa() -> list[Path]:
    out = []
    seen: set[str] = set()
    extra = [
        Path(
            r"\\TANGOSRV\Compartido\CLIENTES\GROWTH MARKETING SA"
            r"\Cierre 2025 01-07-2025 al 30-06-2026\202507\Resumenes"
            r"\Resumen tarj credito Julio 2025.pdf"
        ),
        Path(
            r"\\TANGOSRV\Compartido\CLIENTES\GROWTH MARKETING SA"
            r"\Cierre 2025 01-07-2025 al 30-06-2026\202605\Bancos"
            r"\Resumen Tarjeta Cred Mayo 2026.pdf"
        ),
        Path(
            r"\\TANGOSRV\Compartido\CLIENTES\GROWTH MARKETING SA"
            r"\Cierre 2025 01-07-2025 al 30-06-2026\202606\Bancos"
            r"\Resumen_Tarjetas Cred Junio 2026.pdf"
        ),
    ]
    for p in extra + sorted(BANCOS.glob("*.pdf")):
        n = p.name.upper()
        if not p.exists():
            continue
        if p.suffix.lower() != ".pdf":
            continue
        if "VISA" not in n and "TARJETA" not in n and "TARJ" not in n:
            continue
        key = str(p.resolve()).lower() if p.exists() else p.name.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def extraer() -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    pagos: list[dict] = []
    resúmenes: list[dict] = []
    avisos: list[dict] = []

    for p in pdfs_visa():
        texto = ""
        origen = "texto"
        try:
            texto = _texto_pdf(p)
        except Exception as exc:
            avisos.append({"Archivo": p.name, "Motivo": f"Error lectura: {exc}"})
            continue
        if len(texto.strip()) < 80:
            texto = _ocr_pagina1(p)
            origen = "OCR pág. 1"
            if len(texto.strip()) < 40:
                avisos.append(
                    {
                        "Archivo": p.name,
                        "Motivo": "PDF escaneado sin texto — no se pudo OCR",
                    }
                )
                resúmenes.append(
                    {
                        "Nro resumen": "",
                        "Archivo": p.name,
                        "Cierre": None,
                        "Vencimiento": None,
                        "Saldo anterior ARS": None,
                        "Saldo anterior USD": None,
                        "Pago ARS": None,
                        "Pago USD": None,
                        "Total a pagar ARS": None,
                        "Total a pagar USD": None,
                        "Nota": "PDF escaneado — pendiente OCR",
                        "Origen": "pendiente OCR",
                    }
                )
                continue

        cierre, venc = _fechas_cabecera(texto)
        nro = ""
        m_n = RE_NRO.search(texto)
        if m_n:
            nro = m_n.group("nro")

        saldo_ars = saldo_usd = None
        total_ars = total_usd = None
        m_s = RE_SALDO_ANT.search(texto)
        m_t = RE_TOTAL.search(texto)
        if m_s:
            saldo_ars = _money(m_s.group("pesos"))
            saldo_usd = _money(m_s.group("usd"))
        if m_t:
            total_ars = _money(m_t.group("pesos"))
            total_usd = _money(m_t.group("usd"))

        pago_ars = 0.0
        pago_usd = 0.0
        n_pagos = 0
        nota = ""
        for line in texto.splitlines():
            s = line.strip()
            m = RE_PAGO.match(s)
            if m:
                fecha = _fecha(m.group("fecha"))
                moneda = m.group("moneda").upper()
                importe = _money((m.group("signo") or "") + m.group("importe"))
                importe_abs = abs(importe)
                if moneda == "PESOS":
                    pago_ars += importe_abs
                else:
                    pago_usd += importe_abs
                n_pagos += 1
                pagos.append(
                    {
                        "Fecha": fecha,
                        "Concepto": f"SU PAGO EN {moneda}",
                        "Moneda": "ARS" if moneda == "PESOS" else "USD",
                        "Importe": importe_abs,
                        "Importe (como figura)": importe,
                        "Contravalor ARS": None,
                        "Tipo de cambio": None,
                        "Nro resumen": nro,
                        "Cierre resumen": cierre,
                        "Vencimiento": venc,
                        "Archivo": p.name,
                        "Origen": origen,
                    }
                )
                continue
            md = RE_DEUDA.match(s)
            if not md:
                continue
            fecha = _fecha(md.group("fecha"))
            usd = abs(_money(md.group("usd")))
            ars_eq = abs(_money(md.group("ars")))
            tc = _money(md.group("tc").replace(".", "").replace(",", ".")) if md.group("tc") else None
            pago_usd += usd
            n_pagos += 1
            nota = f"Transferencia deuda USD {usd:,.2f} → ARS {ars_eq:,.2f}"
            pagos.append(
                {
                    "Fecha": fecha,
                    "Concepto": "TRANSFERENCIA DEUDA (USD a pesos)",
                    "Moneda": "USD",
                    "Importe": usd,
                    "Importe (como figura)": -usd,
                    "Contravalor ARS": ars_eq,
                    "Tipo de cambio": tc,
                    "Nro resumen": nro,
                    "Cierre resumen": cierre,
                    "Vencimiento": venc,
                    "Archivo": p.name,
                    "Origen": origen,
                }
            )

        if n_pagos == 0:
            nota = "Sin líneas SU PAGO"
        elif pago_usd == 0 and saldo_usd is not None and saldo_usd <= 0:
            nota = "Sin pago USD (saldo anterior en dólares a favor)"

        resúmenes.append(
            {
                "Nro resumen": nro,
                "Archivo": p.name,
                "Cierre": cierre,
                "Vencimiento": venc,
                "Saldo anterior ARS": saldo_ars,
                "Saldo anterior USD": saldo_usd,
                "Pago ARS": pago_ars if n_pagos else None,
                "Pago USD": pago_usd if n_pagos else None,
                "Total a pagar ARS": total_ars,
                "Total a pagar USD": total_usd,
                "Nota": nota,
                "Origen": origen,
            }
        )
        if n_pagos == 0 and origen != "pendiente OCR":
            avisos.append(
                {
                    "Archivo": p.name,
                    "Motivo": "Sin líneas SU PAGO (puede ser crédito / sin pago USD)",
                }
            )

    df_pagos = pd.DataFrame(pagos)
    if not df_pagos.empty:
        df_pagos = df_pagos.sort_values(["Fecha", "Moneda"], ignore_index=True)
    df_res = pd.DataFrame(resúmenes)
    return df_pagos, df_res, avisos


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    df_pagos, df_res, avisos = extraer()
    print("pagos", len(df_pagos), "resumenes", len(df_res), "avisos", len(avisos))
    if not df_pagos.empty:
        print(df_pagos.to_string(index=False))

    ars = df_pagos[df_pagos["Moneda"] == "ARS"] if not df_pagos.empty else pd.DataFrame()
    usd = df_pagos[df_pagos["Moneda"] == "USD"] if not df_pagos.empty else pd.DataFrame()
    tot_ars = float(ars["Importe"].sum()) if not ars.empty else 0.0
    tot_usd = float(usd["Importe"].sum()) if not usd.empty else 0.0

    kpis = [
        ("Pagos en pesos", len(ars)),
        ("Total ARS", tot_ars),
        ("Pagos en USD", len(usd)),
        ("Total USD", tot_usd),
        ("Resúmenes", len(df_res)),
    ]
    hojas = [("Por resumen", df_res)]
    if avisos:
        hojas.append(("Avisos", pd.DataFrame(avisos)))

    for dest in (OUT_T, OUT_DESK):
        guardar_informe_excel(
            dest,
            titulo="Pagos Visa / tarjeta Galicia",
            subtitulo="GROWTH MARKETING SA — SU PAGO EN PESOS y SU PAGO EN USD",
            periodo="Cierre 01/07/2025 al 30/06/2026",
            kpis=kpis,
            resumenes=[
                ("Pagos en pesos", ars.drop(columns=["Moneda"], errors="ignore") if not ars.empty else pd.DataFrame()),
                ("Pagos en USD", usd.drop(columns=["Moneda"], errors="ignore") if not usd.empty else pd.DataFrame()),
            ],
            detalle=df_pagos,
            hoja_detalle="Todos los pagos",
            hojas_adicionales=hojas,
            col_moneda=["Importe", "Importe (como figura)", "Contravalor ARS", "Tipo de cambio",
                        "Saldo anterior ARS", "Saldo anterior USD",
                        "Pago ARS", "Pago USD", "Total a pagar ARS", "Total a pagar USD"],
            col_fecha=["Fecha", "Cierre", "Vencimiento", "Cierre resumen"],
            total_col="Importe",
        )
        print("OK", dest)


if __name__ == "__main__":
    main()
