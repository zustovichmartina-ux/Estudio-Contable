#!/usr/bin/env python3
"""Clasifica facturas Oftalmología Rele jul/2026 → gasto vs activo."""

from __future__ import annotations

import re
import sys
from pathlib import Path

import fitz
import pandas as pd
import pdfplumber

from excel_formato_estudio import guardar_informe_excel

BASE = Path(r"C:\Users\recep\Desktop\Estudio Contable\_tmp_rele_facturas")
OUT = Path(r"C:\Users\recep\Desktop\Compras_Oftalmologia_Rele_Julio2026.xlsx")

MEDICOS = (
    "bravo",
    "lamare",
    "olsina",
    "cisneros",
    "aleman",
    "rebora",
    "barros",
    "lara",
    "suarez",
    "rodriguez",
    "gonzalez",
)

# (keywords, clase, tipo, detalle)
RULES: list[tuple[tuple[str, ...], str, str, str]] = [
    (
        ("alcon",),
        "Activo",
        "Insumos quirúrgicos inventariables",
        "Productos Alcon (lentes/insumos cirugía)",
    ),
    (
        ("implantec", "viscoelast", "viscotec"),
        "Activo",
        "Insumos quirúrgicos inventariables",
        "Sustancia viscoelástica / insumos cirugía",
    ),
    (
        ("lampara", "lámpara", "mesa electr", "ls-4", "med srl"),
        "Activo",
        "Equipamiento médico",
        "Equipamiento oftalmológico",
    ),
    (("payway",), "Gasto", "Comisiones medios de pago", "Aranceles Payway"),
    (
        ("via cargo", "via bariloche", "encomienda"),
        "Gasto",
        "Fletes y encomiendas",
        "Encomiendas / flete",
    ),
    (("soda",), "Gasto", "Servicios / agua", "Agua / soda"),
    (
        ("cityssan", "quality clean", "limpa", "limpieza"),
        "Gasto",
        "Limpieza y sanitización",
        "Limpieza",
    ),
    (
        ("vetra", "reparacion", "reparación", "ascensor", "cerradura"),
        "Gasto",
        "Mantenimiento edilicio",
        "Reparaciones edificio / ascensor",
    ),
    (
        ("diferencia de cambio",),
        "Gasto",
        "Diferencia de cambio",
        "Ajuste cambiario",
    ),
    (
        ("oxford", "libreria", "librería", "papelera", "camoga"),
        "Gasto",
        "Útiles de oficina / librería",
        "Papelería",
    ),
    (
        ("farmacia", "magister", "gaona"),
        "Gasto",
        "Farmacia / medicamentos",
        "Farmacia",
    ),
    (
        ("oaci", "gsj", "integral pack"),
        "Gasto",
        "Insumos varios",
        "Insumos / packaging",
    ),
    (("biomat",), "Gasto", "Proveedor médico / ajuste", "Biomat"),
    (("rabbione", "msz"), "Gasto", "Proveedores varios", "Revisar detalle"),
    (
        ("honorario",),
        "Gasto",
        "Honorarios profesionales médicos",
        "Honorarios médicos",
    ),
]


def parse_ar(s: str) -> float | None:
    t = str(s).strip().replace("$", "").replace(" ", "").replace("USD", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def texto_pdf(path: Path) -> str:
    try:
        with pdfplumber.open(str(path)) as pdf:
            t = "\n".join((p.extract_text() or "") for p in pdf.pages)
        if len(t.strip()) >= 40:
            return t
    except Exception:
        pass
    try:
        doc = fitz.open(str(path))
        t = "\n".join((page.get_text() or "") for page in doc)
        doc.close()
        return t
    except Exception:
        return ""


def total_y_fecha(t: str) -> tuple[str, float | None]:
    fecha = ""
    m = re.search(
        r"Fecha(?: de Emisi[oó]n)?:\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        t,
        re.I,
    )
    if m:
        fecha = m.group(1).replace("-", "/")
    else:
        m = re.search(r"FECHA\s+(\d{2})\s+(\d{2})\s+(\d{2})", t, re.I)
        if m:
            fecha = f"{m.group(1)}/{m.group(2)}/20{m.group(3)}"
        else:
            m = re.search(r"\b(\d{2}/\d{2}/20\d{2})\b", t)
            if m:
                fecha = m.group(1)

    total = None
    pats = [
        r"Importe Total:\s*[$]?\s*([\d.]+,\d{2})",
        r"Total:\s*[$]?\s*([\d.]+,\d{2})",
        r"TOTAL\s*[$]\s*([\d.]+,\d{2})",
        r"TOTAL\s+[$]?\s*([\d.]+,\d{2})",
        r"Total\s+USD\s*([\d.,]+)",
    ]
    for pat in pats:
        ms = list(re.finditer(pat, t, re.I))
        if ms:
            total = parse_ar(ms[-1].group(1))
            break
    return fecha, total


def proveedor(t: str, name: str) -> str:
    for pat in (
        r"Raz[oó]n Social:\s*([^\n]+)",
        r"PAYWAY S\.A\.U\.",
        r"VIA BARILOCHE S\.A\.",
        r"IMPLANTEC S\.A\.",
        r"Med SRL",
        r"CITYSSAN S\.A\.",
        r"CITYSSAN",
        r"\bALCON\b",
        r"\bBIOMAT\b",
    ):
        m = re.search(pat, t, re.I)
        if m:
            return (m.group(1) if m.lastindex else m.group(0)).strip()[:80]
    af = name.upper()
    for tag in (
        "ALCON",
        "BIOMAT",
        "VETRA",
        "OACI",
        "GSJ",
        "OXFORD",
        "CAMOGA",
        "MAGISTER",
        "RABBIONE",
        "MSZ",
        "LIMPA",
        "IMPLANTEC",
        "INTEGRAL",
        "PAYWAY",
        "CITYSSAN",
        "SODA",
    ):
        if tag in af:
            return tag
    stem = re.sub(r"^\d+[_\- ]*", "", Path(name).stem)
    stem = re.sub(r"[-_ ]*(FC|FA|FAC|HOJA).*$", "", stem, flags=re.I)
    return stem[:80].strip(" -_") or name


def clasificar(name: str, t: str, prov: str) -> tuple[str, str, str]:
    blob = f"{name}\n{prov}\n{t}".lower()
    if any(m in name.lower() for m in MEDICOS) or "honorario" in blob:
        return "Gasto", "Honorarios profesionales médicos", "Honorarios médicos"
    for keys, clase, tipo, det in RULES:
        if any(k in blob for k in keys):
            return clase, tipo, det
    return "Gasto", "Sin clasificar — revisar", "Revisar comprobante"


def detalle(t: str) -> str:
    hits: list[str] = []
    for ln in t.splitlines():
        low = ln.lower()
        if any(
            k in low
            for k in (
                "honorario",
                "lámpara",
                "lampara",
                "mesa",
                "visco",
                "arancel",
                "encomienda",
                "reparacion",
                "reparación",
                "diferencia de cambio",
                "limpieza",
                "servicio",
                "lente",
                "pack",
                "cerradura",
                "soda",
            )
        ):
            hits.append(ln.strip())
    return " | ".join(hits[:3])[:250]


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    files = sorted([p for p in BASE.iterdir() if p.is_file()], key=lambda p: p.name.lower())
    grupos: dict[str, list[Path]] = {}
    for p in files:
        key = re.sub(r"\s*[-_]?\s*HOJA\s*\d+", "", p.stem, flags=re.I).strip().upper()
        grupos.setdefault(key, []).append(p)

    filas: list[dict] = []
    for key, paths in sorted(grupos.items()):
        paths = sorted(paths, key=lambda p: p.name)
        print(f">> {key}", flush=True)
        textos: list[str] = []
        for p in paths:
            if p.suffix.lower() == ".pdf":
                t = texto_pdf(p)
            else:
                t = ""  # imagen: clasifica por nombre de archivo
            textos.append(t)
            print(f"   {p.name} len={len(t)}", flush=True)
        texto = "\n".join(textos)
        archivo = " + ".join(p.name for p in paths)
        prov = proveedor(texto, paths[0].name)
        fecha, total = total_y_fecha(texto)
        clase, tipo, det0 = clasificar(archivo, texto, prov)
        det = detalle(texto) or det0
        mon = "USD" if re.search(r"\bUSD\b|D[oó]lares", texto, re.I) else "ARS"
        m = re.search(r"(\d{4,5}[-_]\d{5,8})", texto)
        nro = m.group(1) if m else ""
        if not nro:
            m = re.search(r"FC\s*(\d+)", archivo, re.I)
            nro = m.group(1) if m else ""
        obs = []
        if not texto and paths[0].suffix.lower() != ".pdf":
            obs.append("Imagen — clasificado por nombre/proveedor")
        if len(paths) > 1:
            obs.append("Multi-hoja")
        filas.append(
            {
                "Fecha": fecha,
                "Proveedor": prov,
                "Nro comprobante": nro,
                "Archivo": archivo,
                "Clase (Gasto/Activo)": clase,
                "Tipo": tipo,
                "Detalle / conceptos": det[:300],
                "Importe total": total if total is not None else "",
                "Moneda": mon,
                "Observaciones": " | ".join(obs),
            }
        )

    df = pd.DataFrame(filas)
    print("filas", len(df))
    print(df.groupby(["Clase (Gasto/Activo)", "Tipo"]).size().to_string())

    resumen = (
        df.groupby(["Clase (Gasto/Activo)", "Tipo"]).size().reset_index(name="Cantidad")
    )
    dfn = df.copy()
    dfn["Imp"] = pd.to_numeric(dfn["Importe total"], errors="coerce")
    por = (
        dfn[dfn["Moneda"] == "ARS"]
        .groupby(["Clase (Gasto/Activo)", "Tipo"])["Imp"]
        .agg(Cantidad="count", Importe_ARS="sum")
        .reset_index()
    )
    kpis = [
        ("Comprobantes", float(len(df))),
        ("Gastos", float((df["Clase (Gasto/Activo)"] == "Gasto").sum())),
        ("Activos", float((df["Clase (Gasto/Activo)"] == "Activo").sum())),
        (
            "Total ARS parseado",
            float(dfn.loc[dfn["Moneda"] == "ARS", "Imp"].sum(skipna=True) or 0),
        ),
    ]
    guardar_informe_excel(
        OUT,
        titulo="Compras — Oftalmología Rele Mar del Plata SRL",
        subtitulo="Clasificación gasto vs activo — Facturas julio 2026",
        periodo="07/2026",
        kpis=kpis,
        resumenes=[
            ("Cantidad por tipo", resumen),
            ("Totales ARS por tipo", por),
        ],
        detalle=df,
        hoja_detalle="Detalle compras",
        col_moneda=["Importe total", "Importe_ARS"],
        col_fecha=["Fecha"],
        total_col="Importe total",
    )
    print("Excel", OUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
