# -*- coding: utf-8 -*-
"""Detalle percepciones IVA / IIBB en FCC Rele jul-2026."""
from __future__ import annotations

import re
import sys
import unicodedata
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from excel_formato_estudio import guardar_informe_excel

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026\07-2026"
)
OUT = FOLDER / "Percepciones_IVA_IIBB_Rele_07-2026.xlsx"

_MONEY = Decimal("0.01")

RE_MONEY = re.compile(
    r"\$?\s*(-?\d{1,3}(?:\.\d{3})*(?:,\d{2})|-?\d+,\d{2})"
)

JURIS = [
    ("CABA", re.compile(r"\bCABA\b|CIUDAD\s+AUT[OÓ]NOMA|CAPITAL\s+FEDERAL|C\.?A\.?B\.?A", re.I)),
    ("BUENOS AIRES", re.compile(r"BUENOS\s+AIRES|\bPBA\b|\bBA\b|ARBA|PROVINCIA\s+DE\s+BUENOS", re.I)),
    ("CORDOBA", re.compile(r"C[OÓ]RDOBA", re.I)),
    ("SANTA FE", re.compile(r"SANTA\s+FE", re.I)),
    ("MENDOZA", re.compile(r"MENDOZA", re.I)),
    ("TUCUMAN", re.compile(r"TUCUM[AÁ]N", re.I)),
    ("ENTRE RIOS", re.compile(r"ENTRE\s+R[IÍ]OS", re.I)),
    ("NEUQUEN", re.compile(r"NEUQU[EÉ]N", re.I)),
    ("RIO NEGRO", re.compile(r"R[IÍ]O\s+NEGRO", re.I)),
    ("CHACO", re.compile(r"\bCHACO\b", re.I)),
    ("MISIONES", re.compile(r"MISIONES", re.I)),
    ("SALTA", re.compile(r"\bSALTA\b", re.I)),
    ("JUJUY", re.compile(r"JUJUY", re.I)),
    ("SAN JUAN", re.compile(r"SAN\s+JUAN", re.I)),
    ("SAN LUIS", re.compile(r"SAN\s+LUIS", re.I)),
    ("LA PAMPA", re.compile(r"LA\s+PAMPA", re.I)),
    ("LA RIOJA", re.compile(r"LA\s+RIOJA", re.I)),
    ("CATAMARCA", re.compile(r"CATAMARCA", re.I)),
    ("SANTIAGO DEL ESTERO", re.compile(r"SANTIAGO\s+DEL\s+ESTERO", re.I)),
    ("CHUBUT", re.compile(r"CHUBUT", re.I)),
    ("SANTA CRUZ", re.compile(r"SANTA\s+CRUZ", re.I)),
    ("TIERRA DEL FUEGO", re.compile(r"TIERRA\s+DEL\s+FUEGO", re.I)),
    ("FORMOSA", re.compile(r"FORMOSA", re.I)),
    ("CORRIENTES", re.compile(r"CORRIENTES", re.I)),
]


def _num(raw: str) -> Decimal:
    s = str(raw or "").strip().replace("$", "").replace(" ", "")
    if not s:
        return Decimal("0")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s).quantize(_MONEY, rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0")


def _norm(s: str) -> str:
    t = unicodedata.normalize("NFKD", str(s or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.upper()


def texto_archivo(path: Path) -> str:
    if path.suffix.lower() == ".pdf":
        try:
            import pdfplumber

            with pdfplumber.open(path) as pdf:
                return "\n".join((p.extract_text() or "") for p in pdf.pages)
        except Exception:
            try:
                import fitz

                doc = fitz.open(path)
                t = "\n".join(page.get_text("text") or "" for page in doc)
                doc.close()
                return t
            except Exception:
                return ""
    return ""


def jurisdiccion(blob: str) -> str:
    n = _norm(blob)
    for nombre, pat in JURIS:
        if pat.search(n):
            return nombre
    if re.search(r"RG\s*5329|PER\.?\s*RG", n):
        return "NACIONAL (RG 5329)"
    return "S/D"


def extraer_percepciones(texto: str) -> list[dict]:
    """Solo percepciones sufridas con importe > 0 (no nomenclador IIBB ni IVA 21%)."""
    if not texto.strip():
        return []
    t = texto.replace("\r", "\n")
    t = re.sub(r"\(cid:\d+\)", "", t)
    lineas = [re.sub(r"\s+", " ", ln).strip() for ln in t.splitlines() if ln.strip()]
    hallados: list[dict] = []

    def add(tipo: str, jur: str, concepto: str, importe: Decimal) -> None:
        if importe <= 0:
            return
        hallados.append(
            {
                "Tipo percepción": tipo,
                "Jurisdicción": jur,
                "Concepto en comprobante": concepto[:180],
                "Importe": float(importe),
            }
        )

    # 1) Perc. IIBB ARBA / CABA: encabezado + fila de importes
    for i, ln in enumerate(lineas):
        n = _norm(ln)
        if "PERC" not in n and "PERCEPC" not in n:
            continue
        tiene_arba = bool(re.search(r"ARBA|PBA|BUENOS\s+AIRES", n))
        tiene_caba = bool(re.search(r"CABA|CAPITAL", n))
        if not (tiene_arba or tiene_caba):
            continue
        if i + 1 >= len(lineas):
            continue
        nums = [_num_us_or_ar(x) for x in re.findall(
            r"\d{1,3}(?:[.,]\d{3})*[.,]\d{2}", lineas[i + 1]
        )]
        # OACI: subtotal, iva, arba, caba, total
        # Quality: gravado, gravado, exento, arba, caba, total
        if tiene_arba and tiene_caba and len(nums) >= 4:
            # penúltimos no-total: ARBA y CABA antes del total
            arba, caba = nums[-3], nums[-2]
            add("Percepción IIBB", "BUENOS AIRES (ARBA)", ln, arba)
            add("Percepción IIBB", "CABA", ln, caba)
        elif tiene_arba and len(nums) >= 2:
            add("Percepción IIBB", "BUENOS AIRES (ARBA)", ln, nums[-2] if len(nums) >= 3 else nums[0])

    # 2) "Percepción de IIBB - BUENOS AIRES $826,45"
    for ln in lineas:
        n = _norm(ln)
        if re.search(r"PERCEPCION\s+DE\s+IIBB|PERCEPCION\s+IIBB|PERC\.?\s*IIBB", n):
            if "ARBA" in n or "CABA" in n:
                continue  # ya tomado en (1) si tenía importe
            m = RE_MONEY.search(ln)
            if not m:
                continue
            add("Percepción IIBB", jurisdiccion(ln), ln, _num(m.group(1)))

    # 3) Mar del Plata Soda: "4.00 % I.I.B.B" + fila US
    for i, ln in enumerate(lineas):
        n = _norm(ln)
        if "I.I.B.B" not in n and "I I B B" not in n.replace(".", " "):
            continue
        if "4.00" not in n and "4,00" not in n:
            continue
        if i + 1 >= len(lineas):
            continue
        nums = [_num_us_or_ar(x) for x in re.findall(
            r"\d+\.\d{2}", lineas[i + 1]
        )]
        # subtotal intern subtotal iibb iva total
        if len(nums) >= 6:
            add("Percepción IIBB", "BUENOS AIRES", ln, nums[3])
        elif len(nums) >= 4:
            add("Percepción IIBB", "BUENOS AIRES", ln, nums[-3])

    # 4) Percepción IVA (no alícuota 21%)
    for ln in lineas:
        n = _norm(ln)
        if not re.search(r"PERCEPCION\s+IVA|PERC\.?\s*IVA|PER\.?\s*RG\s*5329|RG\s*5329", n):
            continue
        if re.search(r"IVA:\s*21|IVA\s+21|21,00\s*%", n) and "RG" not in n:
            continue
        m = re.search(r"(\d+[.,]\d{2})", ln)
        if not m:
            continue
        raw = m.group(1)
        imp = _num_us_or_ar(raw)
        add(
            "Percepción IVA RG 5329" if "5329" in n else "Percepción IVA",
            "NACIONAL",
            ln,
            imp,
        )

    # Dedupe same tipo+jur+importe
    uniq = []
    seen = set()
    for h in hallados:
        k = (h["Tipo percepción"], h["Jurisdicción"], round(h["Importe"], 2))
        if k in seen:
            continue
        seen.add(k)
        uniq.append(h)
    return uniq


def _num_us_or_ar(raw: str) -> Decimal:
    s = str(raw or "").strip().replace("$", "").replace(" ", "")
    if not s:
        return Decimal("0")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return Decimal(s).quantize(_MONEY, rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0")


def datos_archivo(nombre: str) -> dict:
    m = re.match(
        r"(FCC|NCC|NDC)([ABC])(?:(\d+)-)?(\d+)\s+(.+?)\.(pdf|jpeg|jpg)$",
        nombre,
        re.I,
    )
    if not m:
        return {"Comprobante": nombre, "Proveedor": "", "Tipo": ""}
    tipo, letra, pv, nro, prov, _ = m.groups()
    nro_txt = f"{letra} {pv}-{nro}" if pv else f"{letra} {nro}"
    return {
        "Comprobante": f"{tipo}{letra}{pv + '-' if pv else ''}{nro}",
        "Proveedor": prov.replace(" HOJA 1", "").replace(" HOJA 2", "").strip(),
        "Tipo": tipo,
        "Nro": nro_txt,
    }


def main() -> None:
    filas: list[dict] = []
    sin_texto = []
    for p in sorted(FOLDER.iterdir(), key=lambda x: x.name.lower()):
        if p.suffix.lower() not in {".pdf", ".jpeg", ".jpg"}:
            continue
        if p.name.startswith("Percepciones"):
            continue
        texto = texto_archivo(p)
        if p.suffix.lower() == ".pdf" and not texto.strip():
            sin_texto.append(p.name)
        percs = extraer_percepciones(texto)
        meta = datos_archivo(p.name)
        for perc in percs:
            filas.append({**meta, "Archivo": p.name, **perc})

    df = pd.DataFrame(filas)
    if df.empty:
        print("Sin percepciones detectadas en PDF")
        df = pd.DataFrame(
            columns=[
                "Comprobante",
                "Proveedor",
                "Tipo",
                "Nro",
                "Tipo percepción",
                "Jurisdicción",
                "Importe",
                "Concepto en comprobante",
                "Archivo",
            ]
        )

    # una fila por percepción; resumen por factura
    if not df.empty:
        por_fct = (
            df.groupby(["Comprobante", "Proveedor"], dropna=False)
            .agg(
                Cant_percepciones=("Importe", "size"),
                Jurisdicciones=("Jurisdicción", lambda s: " + ".join(sorted(set(s)))),
                Total_percepciones=("Importe", "sum"),
                Tipos=("Tipo percepción", lambda s: " + ".join(sorted(set(s)))),
            )
            .reset_index()
        )
        por_tipo = (
            df.groupby(["Tipo percepción", "Jurisdicción"], dropna=False)["Importe"]
            .sum()
            .reset_index()
            .sort_values(["Tipo percepción", "Jurisdicción"])
        )
        multi = por_fct[por_fct["Cant_percepciones"] > 1]
    else:
        por_fct = pd.DataFrame()
        por_tipo = pd.DataFrame()
        multi = pd.DataFrame()

    kpis = [
        ("Comprobantes con percepción", int(df["Comprobante"].nunique()) if not df.empty else 0, "int"),
        ("Líneas de percepción", len(df), "int"),
        ("Total percepciones", float(df["Importe"].sum()) if not df.empty else 0.0, "money"),
        ("Más de 1 jurisdicción / línea", int(len(multi)), "int"),
    ]
    guardar_informe_excel(
        OUT,
        titulo="Percepciones IVA e IIBB — Rele",
        subtitulo="Oftalmología Rele Mar del Plata SRL — facturas de compra",
        periodo="07/2026",
        kpis=kpis,
        resumenes=[
            ("Por tipo y jurisdicción", por_tipo),
            ("Por comprobante", por_fct),
            ("Comprobantes con más de una percepción", multi),
        ],
        detalle=df,
        hoja_detalle="Detalle percepciones",
        col_moneda=["Importe", "Total_percepciones"],
        total_col="Importe",
    )
    desktop = Path(r"C:\Users\recep\Desktop") / OUT.name
    guardar_informe_excel(
        desktop,
        titulo="Percepciones IVA e IIBB — Rele",
        subtitulo="Oftalmología Rele Mar del Plata SRL — facturas de compra",
        periodo="07/2026",
        kpis=kpis,
        resumenes=[
            ("Por tipo y jurisdicción", por_tipo),
            ("Por comprobante", por_fct),
            ("Comprobantes con más de una percepción", multi),
        ],
        detalle=df,
        hoja_detalle="Detalle percepciones",
        col_moneda=["Importe", "Total_percepciones"],
        total_col="Importe",
    )
    print("OUT", OUT)
    print("DESKTOP", desktop)
    print("filas", len(df))
    if not df.empty:
        print(df[["Comprobante", "Proveedor", "Tipo percepción", "Jurisdicción", "Importe"]].to_string(index=False))
    if sin_texto:
        print("PDF sin texto", sin_texto)


if __name__ == "__main__":
    main()
