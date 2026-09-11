# -*- coding: utf-8 -*-
"""Consumos (compras) Visa Galicia — Growth Marketing SA, línea por línea."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_formato_estudio import guardar_informe_excel
from _extraer_pagos_visa_growth import (
    BANCOS,
    RE_NRO,
    _fecha,
    _money,
    _texto_pdf,
    pdfs_visa,
    easyocr,
    fitz,
    np,
    Image,
    io,
)

OUT_T = BANCOS / "Consumos_Visa_TC_Growth_2026.xlsx"
OUT_DESK = Path(r"C:\Users\recep\Desktop\Consumos_Visa_TC_Growth_2026.xlsx")

RE_LINEA = re.compile(r"^(\d{2}-\d{2}-\d{2})\s+(?:([KEF*Q])\s+)?(.+)$")
RE_MONEY = r"-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}"
RE_FX = re.compile(
    rf"^(?P<ref>.+?)\s+(?P<mon>USD|EUR)\s+(?P<orig>{RE_MONEY})"
    rf"\s+(?P<comp>\d{{4,8}})\s+(?P<usd>{RE_MONEY})\s*$",
    re.I,
)
RE_USD_PEGADO = re.compile(
    rf"^(?P<ref>.+USD)\s+(?P<orig>{RE_MONEY})\s+(?P<comp>\d{{4,8}})"
    rf"\s+(?P<usd>{RE_MONEY})\s*$",
    re.I,
)
RE_ARS_CUOTA = re.compile(
    rf"^(?P<ref>.+?)\s+(?P<cuota>\d{{2}}/\d{{2}})\s+(?P<comp>\d{{4,8}})"
    rf"\s+(?P<ars>{RE_MONEY})\s*$",
)
RE_ARS = re.compile(
    rf"^(?P<ref>.+?)\s+(?P<comp>\d{{4,8}})\s+(?P<ars>{RE_MONEY})\s*$",
)
RE_ARS_SOLO = re.compile(rf"^(?P<ref>.+?)\s+(?P<ars>{RE_MONEY})\s*$")
RE_TITULAR = re.compile(
    r"TARJETA\s+(\d+)\s+Total Consumos de\s+(.+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})",
    re.I,
)

SKIP = re.compile(
    r"SU PAGO|TRANSFERENCIA DEUDA|SALDO ANTERIOR|TOTAL A PAGAR|"
    r"DETALLE DEL CONSUMO|FECHA REFERENCIA|IMPUESTO|COMISI[OÓ]N|"
    r"PERCEP|IIBB|DB\.?\s*IVA|DB\.?\s*RG|IVA \$|DEV\.IMP|"
    r"INTERESES FINANCIACION|Cuotas a vencer|Página \d",
    re.I,
)

CATS: list[tuple[str, str]] = [
    (r"OPENAI|CHATGPT", "IA / OpenAI"),
    (r"GOOGLE|GSUITE|WORKSPACE", "Google Workspace / Google"),
    (r"FACEBK|FACEBOOK|META ", "Publicidad Meta"),
    (r"MADGICX", "Publicidad / Madgicx"),
    (r"ADOBE", "Adobe"),
    (r"FREEPIK", "Diseño / Freepik"),
    (r"ARTLIST", "Contenido / Artlist"),
    (r"OPUS CLIP|CAPTIONS|CAPCUT|MOVAVI", "Video / edición"),
    (r"DOCUSIGN", "Software / DocuSign"),
    (r"PADDLE|N8N", "Software / n8n"),
    (r"NOTION", "Software / Notion"),
    (r"MANYCHAT", "Software / Manychat"),
    (r"COURSERA|UDEMY|CODERHOUSE|VA360|ACADEMY", "Capacitación"),
    (r"MASTER METRICS", "Analytics / Master Metrics"),
    (r"MAGNIFIC", "Software / Magnific"),
    (r"FAST\s*CO", "Software / Fastco"),
    (r"GODADDY", "Dominios / GoDaddy"),
    (r"PERFIT", "Email / Perfit"),
    (r"TIENDUP", "Ecommerce / Tiendup"),
    (r"DONWEB", "Hosting / Donweb"),
    (r"NIC ARGENTINA", "Dominios / NIC"),
    (r"MERPAGO|MERCADOPAGO", "Mercado Pago"),
    (r"KLINKO", "Klinko"),
]


def _clasificar(ref: str) -> str:
    n = (ref or "").upper()
    for pat, cat in CATS:
        if re.search(pat, n):
            return cat
    return "Otros"


def _limpiar_ref(ref: str) -> str:
    t = re.sub(r"\s+", " ", ref or "").strip()
    t = re.sub(r"(USD|EUR)\s*$", "", t, flags=re.I).strip()
    t = re.sub(r"in[0-9A-Za-z]{5,}\s*$", "", t).strip()
    return t


def parsear_consumo(resto: str) -> dict | None:
    s = resto.strip()
    m = RE_FX.match(s) or RE_USD_PEGADO.match(s)
    if m:
        gd = m.groupdict()
        return {
            "Referencia": _limpiar_ref(gd["ref"]),
            "Cuota": "",
            "Comprobante": gd.get("comp") or "",
            "Moneda origen": (gd.get("mon") or "USD").upper(),
            "Importe origen": _money(gd["orig"]),
            "Moneda": "USD",
            "Importe ARS": 0.0,
            "Importe USD": _money(gd["usd"]),
        }
    m = RE_ARS_CUOTA.match(s)
    if m:
        return {
            "Referencia": _limpiar_ref(m.group("ref")),
            "Cuota": m.group("cuota"),
            "Comprobante": m.group("comp"),
            "Moneda origen": "ARS",
            "Importe origen": _money(m.group("ars")),
            "Moneda": "ARS",
            "Importe ARS": _money(m.group("ars")),
            "Importe USD": 0.0,
        }
    m = RE_ARS.match(s)
    if m:
        return {
            "Referencia": _limpiar_ref(m.group("ref")),
            "Cuota": "",
            "Comprobante": m.group("comp"),
            "Moneda origen": "ARS",
            "Importe origen": _money(m.group("ars")),
            "Moneda": "ARS",
            "Importe ARS": _money(m.group("ars")),
            "Importe USD": 0.0,
        }
    m = RE_ARS_SOLO.match(s)
    if m and not re.search(r"TARJETA\s+\d+\s+Total", m.group("ref"), re.I):
        return {
            "Referencia": _limpiar_ref(m.group("ref")),
            "Cuota": "",
            "Comprobante": "",
            "Moneda origen": "ARS",
            "Importe origen": _money(m.group("ars")),
            "Moneda": "ARS",
            "Importe ARS": _money(m.group("ars")),
            "Importe USD": 0.0,
        }
    return None


def _ocr_paginas(path: Path, max_paginas: int = 3) -> str:
    if easyocr is None:
        return ""
    doc = fitz.open(str(path))
    try:
        lector = easyocr.Reader(["es", "en"], gpu=False, verbose=False)
        bloques: list[str] = []
        for i in range(min(max_paginas, doc.page_count)):
            pix = doc[i].get_pixmap(dpi=140)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            filas: dict[int, list[tuple[float, str]]] = {}
            for bbox, txt, _conf in lector.readtext(np.array(img)):
                y = int((bbox[0][1] + bbox[2][1]) / 2 / 14) * 14
                filas.setdefault(y, []).append((bbox[0][0], txt))
            lineas = []
            for y in sorted(filas):
                partes = sorted(filas[y], key=lambda x: x[0])
                lineas.append(" ".join(t for _, t in partes))
            bloques.append("\n".join(lineas))
        return "\n".join(bloques)
    except Exception:
        return ""
    finally:
        doc.close()


def extraer_archivo(path: Path) -> list[dict]:
    texto = _texto_pdf(path)
    if len(texto.strip()) < 80:
        print("  OCR", path.name)
        texto = _ocr_paginas(path, 3)
    nro = ""
    m_n = RE_NRO.search(texto)
    if m_n:
        nro = m_n.group("nro")
        nro = nro.replace("o", "0").replace("O", "0")
        if nro.lower().startswith("vi"):
            nro = "VI" + nro[2:]

    filas: list[dict] = []
    titular = ""
    tarjeta = ""
    pendientes: list[dict] = []

    def flush(tit: str, tarj: str) -> None:
        for row in pendientes:
            row["Titular"] = tit
            row["Tarjeta"] = tarj
            filas.append(row)
        pendientes.clear()

    for raw in texto.splitlines():
        line = raw.strip()
        mt = RE_TITULAR.search(line)
        if mt:
            flush(mt.group(2).strip(), mt.group(1))
            titular, tarjeta = mt.group(2).strip(), mt.group(1)
            continue
        if SKIP.search(line) or not RE_LINEA.match(line):
            continue
        m = RE_LINEA.match(line)
        fecha, marca, resto = m.group(1), m.group(2) or "", m.group(3)
        parsed = parsear_consumo(resto)
        if not parsed:
            continue
        parsed.update(
            {
                "Fecha": _fecha(fecha),
                "Marca": marca,
                "Clasificación": _clasificar(parsed["Referencia"] + " " + resto),
                "Nro resumen": nro,
                "Archivo": path.name,
                "Línea original": line,
            }
        )
        pendientes.append(parsed)
    flush(titular, tarjeta)
    return filas


CUENTA_PLAT = ("Gastos Plataformas Online U$S", "42301")
CUENTA_MKT = ("Gastos de Marketing Digital U$S", "42443")

# (regex sobre referencia, descripción, tipo, cuenta)
MAPA_VISTA: list[tuple[str, str, str, tuple[str, str]]] = [
    (r"FACEBK|FACEBOOK", "Pauta publicitaria en Facebook/Instagram", "Publicidad (pauta paga)", CUENTA_MKT),
    (r"MADGICX", "Pauta / optimización de ads (Madgicx)", "Publicidad (pauta paga)", CUENTA_MKT),
    (r"CAPCUT", "Editor de video | ByteDance", "Software - Edición de video/Contenido", CUENTA_PLAT),
    (r"OPUS CLIP", "Editor de video con IA (Opus Clip)", "Software - Edición de video/Contenido", CUENTA_PLAT),
    (r"CAPTIONS", "Subtítulos / edición de video (Captions)", "Software - Edición de video/Contenido", CUENTA_PLAT),
    (r"MOVAVI", "Editor de video (Movavi)", "Software - Edición de video/Contenido", CUENTA_PLAT),
    (r"NOTION", "Gestión de proyectos y documentos", "Software - Productividad", CUENTA_PLAT),
    (r"OPENAI|CHATGPT", "Suscripción a ChatGPT", "Software - Productividad/IA", CUENTA_PLAT),
    (r"DOCUSIGN", "Firma electrónica de documentos", "Software - Administración", CUENTA_PLAT),
    (r"ADOBE", "Software de diseño (Adobe)", "Software - Diseño/Contenido", CUENTA_PLAT),
    (r"FREEPIK", "Banco de imágenes y diseño", "Software - Diseño/Contenido", CUENTA_PLAT),
    (r"ARTLIST", "Música y stock de video (Artlist)", "Software - Diseño/Contenido", CUENTA_PLAT),
    (r"GOOGLE|GSUITE|WORKSPACE", "Correo y herramientas de Google", "Software - Productividad", CUENTA_PLAT),
    (r"MASTER METRICS", "Analítica de marketing", "Software - Marketing/Analítica", CUENTA_PLAT),
    (r"PADDLE|N8N", "Automatización de flujos (n8n)", "Software - Automatización", CUENTA_PLAT),
    (r"KLINKO", "Software / plataforma Klinko", "Software - Marketing", CUENTA_PLAT),
    (r"MANYCHAT", "Automatización de WhatsApp/IG (Manychat)", "Software - Marketing", CUENTA_PLAT),
    (r"COURSERA|UDEMY|CODERHOUSE|VA360|ACADEMY", "Capacitación online", "Software - Capacitación", CUENTA_PLAT),
    (r"MAGNIFIC", "Imágenes con IA (Magnific)", "Software - Productividad/IA", CUENTA_PLAT),
    (r"FAST\s*CO", "Software Fastco", "Software - Productividad", CUENTA_PLAT),
    (r"GODADDY", "Dominios / hosting (GoDaddy)", "Software - Infraestructura", CUENTA_PLAT),
]


def _vista_comercio(ref: str) -> tuple[str, str, str, str]:
    n = (ref or "").upper()
    for pat, desc, tipo, cuenta in MAPA_VISTA:
        if re.search(pat, n):
            return desc, tipo, cuenta[0], cuenta[1]
    return "Servicio / plataforma online", "Software - Otros", CUENTA_PLAT[0], CUENTA_PLAT[1]


def armar_vista_usd(df: pd.DataFrame) -> pd.DataFrame:
    usd = df[df["Moneda"] == "USD"].copy()
    filas = []
    for _, r in usd.iterrows():
        desc, tipo, cuenta, codigo = _vista_comercio(str(r.get("Referencia") or ""))
        filas.append(
            {
                "Fecha": r["Fecha"],
                "Comercio": r["Referencia"],
                "Qué es": desc,
                "Tipo": tipo,
                "Moneda": "USD",
                "Importe": r["Importe USD"],
                "Cotiz.": None,
                "Pesos": None,
                "Cuenta": cuenta,
                "Cód.": codigo,
            }
        )
    out = pd.DataFrame(filas)
    if not out.empty:
        out = out.sort_values(["Fecha", "Comercio"], ignore_index=True)
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    rows: list[dict] = []
    for p in pdfs_visa():
        got = extraer_archivo(p)
        print(p.name, "lineas", len(got))
        rows.extend(got)
    df = pd.DataFrame(rows)
    if df.empty:
        raise SystemExit("Sin consumos")
    df = df.sort_values(["Fecha", "Moneda", "Referencia"], ignore_index=True)
    vista = armar_vista_usd(df)
    por_cuenta = (
        vista.groupby(["Cuenta", "Cód."], as_index=False)["Importe"]
        .sum()
        .sort_values("Importe", ascending=False)
        if not vista.empty
        else pd.DataFrame()
    )
    por_tipo = (
        vista.groupby("Tipo", as_index=False)["Importe"]
        .sum()
        .sort_values("Importe", ascending=False)
        if not vista.empty
        else pd.DataFrame()
    )
    kpis = [
        ("Líneas USD", len(vista), "int"),
        ("Total USD", float(vista["Importe"].sum()) if not vista.empty else 0.0, "money"),
        ("Plataformas 42301", float(vista.loc[vista["Cód."] == "42301", "Importe"].sum()) if not vista.empty else 0.0, "money"),
        ("Marketing 42443", float(vista.loc[vista["Cód."] == "42443", "Importe"].sum()) if not vista.empty else 0.0, "money"),
    ]
    for dest in (OUT_T, OUT_DESK):
        guardar_informe_excel(
            dest,
            titulo="Consumos USD Visa Galicia",
            subtitulo="GROWTH MARKETING SA — compras en dólares (sin pagos)",
            periodo="Cierre 01/07/2025 al 30/06/2026",
            kpis=kpis,
            resumenes=[
                ("Por cuenta contable", por_cuenta),
                ("Por tipo", por_tipo),
            ],
            detalle=vista,
            hoja_detalle="Detalle USD",
            hojas_adicionales=[("Todas las líneas", df)],
            col_moneda=["Importe", "Importe USD", "Importe ARS", "Importe origen", "Pesos"],
            col_fecha=["Fecha"],
            total_col="Importe",
        )
        print("OK", dest)
    print("USD", float(vista["Importe"].sum()), "lineas", len(vista))
    print(por_tipo.to_string(index=False))


if __name__ == "__main__":
    main()
