# -*- coding: utf-8 -*-
"""
Segregación contable de facturas de venta (TusFacturasAPP PDF → Excel).

Una fila por línea de detalle. Factura A discrimina IVA por línea;
Factura B / NC B traen subtotal bruto y se desagregan con alícuota
aprendida de códigos vistos en Facturas A (o 21% estimado).
"""
from __future__ import annotations

import io
import json
import re
import unicodedata
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Sequence

import pandas as pd

from excel_formato_estudio import exportar_informe_excel, guardar_informe_excel

_MONEY = Decimal("0.01")
TOL = Decimal("0.05")
UMBRAL_CLASIF = 85.0
ALICUOTA_DEFAULT = Decimal("21")

_DIR_DATOS = Path(__file__).resolve().parent / "data"
_PATH_ALICUOTAS = _DIR_DATOS / "codigo_alicuota_iva.json"

COLUMNAS = [
    "Fecha",
    "Tipo Comprobante",
    "Nro. Comprobante",
    "Receptor",
    "Condición IVA Receptor",
    "Concepto",
    "Código Producto",
    "Neto",
    "IVA",
    "Alícuota IVA",
    "Bruto",
    "Servicio/Mercadería",
    "Alícuota Estimada",
    "Comprobante Asociado",
    "Archivo",
    "Revisar",
]

RE_MONEY = re.compile(
    r"(?:-?\d{1,3}(?:\.\d{3})*(?:,\d{2,3})|-?\d+(?:,\d{2,3}))"
)


def money(v: Any) -> Decimal:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return Decimal("0.00")
    if isinstance(v, Decimal):
        return v.quantize(_MONEY, rounding=ROUND_HALF_UP)
    raw = str(v).strip().replace("$", "").replace(" ", "")
    if not raw or raw.lower() in {"nan", "none", "-"}:
        return Decimal("0.00")
    neg = raw.startswith("-")
    raw = raw.lstrip("-")
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    try:
        d = Decimal(raw)
    except Exception:
        return Decimal("0.00")
    if neg:
        d = -d
    return d.quantize(_MONEY, rounding=ROUND_HALF_UP)


def _f(d: Decimal) -> float:
    return float(d)


def _norm(texto: str) -> str:
    if not texto:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper()
    t = re.sub(r"[^A-Z0-9\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


# ── Persistencia alícuotas ────────────────────────────────────────────────────

def cargar_mapa_alicuotas(path: Path | None = None) -> dict[str, Decimal]:
    p = path or _PATH_ALICUOTAS
    if not p.exists():
        return {}
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}
    out: dict[str, Decimal] = {}
    for k, v in (raw or {}).items():
        codigo = str(k).strip()
        if not codigo:
            continue
        try:
            out[codigo] = Decimal(str(v).replace(",", ".").replace("%", ""))
        except Exception:
            continue
    return out


def guardar_mapa_alicuotas(mapa: dict[str, Decimal], path: Path | None = None) -> Path:
    p = path or _PATH_ALICUOTAS
    p.parent.mkdir(parents=True, exist_ok=True)
    serial = {str(k): str(v) for k, v in sorted(mapa.items())}
    p.write_text(json.dumps(serial, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def _fmt_alicuota(raw: Any) -> str:
    if raw is None or raw == "":
        return ""
    if isinstance(raw, Decimal):
        v = raw
    else:
        s = str(raw).strip().replace("%", "").replace(",", ".")
        try:
            v = Decimal(s)
        except Exception:
            return f"{raw}".strip()
    if v == Decimal("10.5") or v == Decimal("10.50"):
        return "10,5%"
    if v == int(v):
        return f"{int(v)}%"
    return f"{str(v).replace('.', ',')}%"


def _parse_alicuota_pct(ali: str | Decimal | None) -> Decimal:
    if isinstance(ali, Decimal):
        return ali
    if not ali:
        return Decimal("0")
    s = str(ali).strip().replace("%", "").replace(",", ".")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0")


def _alicuota_desde_nombre(concepto: str) -> Decimal | None:
    m = re.search(r"\(\s*(10[.,]5|21|27)\s*%\s*\)", concepto or "", re.I)
    if not m:
        return None
    return _parse_alicuota_pct(m.group(1))


# ── PDF texto ─────────────────────────────────────────────────────────────────

def extraer_texto_pdf(data: bytes) -> str:
    try:
        import pdfplumber

        with pdfplumber.open(io.BytesIO(data)) as pdf:
            partes = [(page.extract_text() or "") for page in pdf.pages]
        texto = "\n".join(partes)
        if texto.strip():
            return texto
    except Exception:
        pass
    try:
        import fitz

        doc = fitz.open(stream=data, filetype="pdf")
        partes = [page.get_text("text") or "" for page in doc]
        doc.close()
        return "\n".join(partes)
    except Exception:
        return ""


def _solo_original(texto: str) -> str:
    m = re.search(r"(?m)^\s*DUPLICADO\s*$", texto)
    if m and m.start() > 200:
        return texto[: m.start()]
    m = re.search(r"\bDUPLICADO\b", texto)
    if m and m.start() > 400:
        return texto[: m.start()]
    return texto


# ── Cabecera ──────────────────────────────────────────────────────────────────

def detectar_tipo_desde_nombre(nombre: str) -> str | None:
    """
    Prioridad 1: el nombre del PDF de TusFacturas suele ser inequívoco
    (ej. FACTURA_B-00010-00000093.pdf, NOTA_DE_CREDITO_B-00010-00000001.pdf).
    """
    if not nombre:
        return None
    n = Path(str(nombre)).name.upper()
    n = n.replace(" ", "_").replace("-", "_")

    if "NOTA_DE_CREDITO_A" in n or "NOTA_CREDITO_A" in n or re.search(r"\bNCA\b", n):
        return "Nota de Crédito A"
    if "NOTA_DE_CREDITO_B" in n or "NOTA_CREDITO_B" in n or re.search(r"\bNCB\b", n):
        return "Nota de Crédito B"
    if "NOTA_DE_CREDITO_C" in n or "NOTA_CREDITO_C" in n or re.search(r"\bNCC\b", n):
        return "Nota de Crédito C"
    if "NOTA_DE_CREDITO" in n or "NOTA_CREDITO" in n:
        # Letra suelta tras NOTA_DE_CREDITO_X
        m = re.search(r"NOTA_DE_CREDITO_([ABC])\b", n)
        if m:
            return f"Nota de Crédito {m.group(1)}"
        return "Nota de Crédito"

    if re.search(r"FACTURA_A\b|FACTURAA\b|_FA_|__FA-", n) or n.startswith("FA_"):
        return "Factura A"
    if re.search(r"FACTURA_B\b|FACTURAB\b|_FB_|__FB-", n) or n.startswith("FB_"):
        return "Factura B"
    if re.search(r"FACTURA_C\b|FACTURAC\b|_FC_|__FC-", n):
        return "Factura C"

    # Patrones tipo 3071…__FACTURA_B-00010-…
    m = re.search(r"(?:^|[_\W])FACTURA[_ ]?([ABC])(?:[_\W]|$)", n)
    if m:
        return f"Factura {m.group(1)}"
    m = re.search(r"(?:^|[_\W])FC[_ ]?([ABC])(?:[_\W]|$)", n)
    if m:
        return f"Factura {m.group(1)}"
    m = re.search(r"(?:^|[_\W])NC[_ ]?([ABC])(?:[_\W]|$)", n)
    if m:
        return f"Nota de Crédito {m.group(1)}"
    return None


def detectar_tipo(texto: str, nombre_archivo: str = "") -> str:
    """Tipo del comprobante: primero nombre del PDF, después texto."""
    desde_nombre = detectar_tipo_desde_nombre(nombre_archivo)
    if desde_nombre:
        return desde_nombre

    t = _norm(texto[:2500])
    head = texto[:1200]
    if re.search(r"NOTA\s+DE\s+CREDITO|NOTA\s+DE\s+CR[EÉ]DITO", t):
        if re.search(r"\bCOD\.?\s*0*3\b|\bNCA\b|NOTA DE CREDITO A", t):
            return "Nota de Crédito A"
        if re.search(r"\bCOD\.?\s*0*8\b|\bNCB\b|NOTA DE CREDITO B", t):
            return "Nota de Crédito B"
        if re.search(r"\bCOD\.?\s*013\b|\bNCC\b", t):
            return "Nota de Crédito C"
        return "Nota de Crédito"
    # Orden estricto: C → B → A (evitar que COD.1 mate antes que COD.6)
    if re.search(r"\bCOD\.?\s*011\b|\bFACTURA\s+C\b", t) or re.search(
        r"FACTURA C\b", head, re.I
    ):
        return "Factura C"
    if re.search(r"\bCOD\.?\s*0*6\b|\bFACTURA\s+B\b", t) or re.search(
        r"FACTURA B\b", head, re.I
    ):
        return "Factura B"
    if re.search(r"\bCOD\.?\s*0*1\b|\bFACTURA\s+A\b", t) or re.search(
        r"FACTURA A\b", head, re.I
    ):
        return "Factura A"
    m = re.search(r"COD\.?\s*(\d{1,3})", head, re.I)
    if m:
        cod = int(m.group(1))
        return {
            1: "Factura A",
            6: "Factura B",
            11: "Factura C",
            3: "Nota de Crédito A",
            8: "Nota de Crédito B",
            13: "Nota de Crédito C",
        }.get(cod, f"Código {cod}")
    return "Factura"


def es_tipo_a(tipo: str) -> bool:
    t = (tipo or "").upper()
    return "FACTURA A" in t or "CRÉDITO A" in t or "CREDITO A" in t


def es_tipo_b_nc(tipo: str) -> bool:
    t = (tipo or "").upper()
    return any(
        x in t
        for x in (
            "FACTURA B",
            "FACTURA C",
            "CRÉDITO B",
            "CREDITO B",
            "CRÉDITO C",
            "CREDITO C",
        )
    )


def es_nota_credito(tipo: str) -> bool:
    t = (tipo or "").upper()
    return "CRÉDITO" in t or "CREDITO" in t


def extraer_fecha(texto: str) -> str:
    m = re.search(r"Fecha\s*:?\s*(\d{2}/\d{2}/\d{4})", texto[:2000], re.I)
    if m:
        return m.group(1)
    m = re.search(r"Fecha\s*(?:de\s*)?Emisi[oó]n\s*:?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto[:2000])
    return m.group(1) if m else ""


def extraer_pv_numero(texto: str) -> tuple[str, str]:
    m = re.search(r"Nro\.?\s*:?\s*(\d{4,5})\s*[-/]\s*(\d{6,8})", texto[:2500], re.I)
    if m:
        return f"{int(m.group(1)):05d}", f"{int(m.group(2)):08d}"
    m = re.search(
        r"Punto\s*de\s*Venta\s*:?\s*(\d+)\s*.{0,80}?Comp\.?\s*Nro\.?\s*:?\s*(\d+)",
        texto[:3000],
        re.I | re.S,
    )
    if m:
        return f"{int(m.group(1)):05d}", f"{int(m.group(2)):08d}"
    m = re.search(r"\b(\d{5})\s*[-/]\s*(\d{8})\b", texto[:2500])
    if m:
        return m.group(1), m.group(2)
    return "", ""


def nro_comprobante(pv: str, nro: str) -> str:
    if pv and nro:
        return f"{pv}-{nro}"
    return pv or nro or ""


def extraer_condicion_iva(texto: str) -> str:
    head = _solo_original(texto)[:3500]
    matches = list(
        re.finditer(
            r"(?:CUIT|DNI|CUIL)\s*:?\s*[\d\-]+\s*\(([^)]{3,60})\)",
            head,
            re.I,
        )
    )
    if matches:
        # El primero suele ser emisor sin paréntesis; el del receptor trae (Condición)
        return re.sub(r"\s+", " ", matches[-1].group(1)).strip()
    matches2 = list(
        re.finditer(
            r"Cond(?:ici[oó]n)?\.?\s*frente\s*al\s*IVA\s*:?\s*([^\n\r]{3,60})",
            head,
            re.I,
        )
    )
    if matches2:
        raw = matches2[-1].group(1).strip()
        raw = re.split(r"\s{2,}|CUIT|DNI|Inicio|Per[ií]odo", raw)[0].strip()
        if raw:
            return raw
    return ""


def extraer_receptor(texto: str) -> str:
    """Razón social del comprador (no el emisor)."""
    head = _solo_original(texto)[:4000]
    # TusFacturas: "Apellido y nombre / Razón Social: LOPEZ DIEGO ANDRES"
    # (no confundir con "Razón social: FACTURA A" del encabezado)
    m = re.search(
        r"Apellido\s+y\s+[Nn]ombre\s*/\s*Raz[oó]n\s+Social\s*:?\s*([^\n\r]{3,120})",
        head,
        re.I,
    )
    if m:
        nom = re.sub(r"\s+", " ", m.group(1)).strip()
        nom = re.split(r"\s{2,}|(?:CUIT|DNI|CUIL|Cond\.|Domicilio)", nom)[0].strip()
        if nom and "FACTURA" not in nom.upper():
            return nom
    # Fallback: línea previa al domicilio del receptor + CUIT con condición
    m = re.search(
        r"Apellido\s+y\s+[Nn]ombre\s*/\s*Raz[oó]n\s+Social\s*:?\s*([^\n\r]+)\s*\n\s*Domicilio\s*:",
        head,
        re.I,
    )
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    return ""


def extraer_comprobante_asociado(texto: str) -> str:
    m = re.search(
        r"COMPROBANTES?\s+ASOCIADOS?\s*:?\s*([^\n\r]{5,120})",
        _solo_original(texto),
        re.I,
    )
    if not m:
        return ""
    raw = re.sub(r"\s+", " ", m.group(1)).strip()
    # Recortar tras CUIT/Fecha si vienen en la misma línea
    raw = re.split(r"\s+CUIT\s*:|\s+Fecha\s*:", raw, maxsplit=1)[0].strip()
    return raw


def es_prueba_tecnica(texto: str, receptor: str) -> bool:
    blob = _norm(f"{texto[:2000]} {receptor}")
    return any(
        k in blob
        for k in (
            "PRUEBA TECNICA",
            "PRUEBA TECNICA DE FACTURACION",
            "PRUEBA DE FACTURACION ELECTRONICA",
            "TESTING",
        )
    )


def bonificacion_general(texto: str) -> Decimal:
    # TusFacturas imprime "Bonificación General - $ 48.760,33" (guión antes del $)
    m = re.search(
        r"Bonificaci[oó]n\s+General\s*:?\s*-?\s*\$?\s*-?\s*("
        + RE_MONEY.pattern
        + r")",
        _solo_original(texto),
        re.I,
    )
    return abs(money(m.group(1))) if m else Decimal("0.00")


def totales_pie(texto: str) -> tuple[Decimal, Decimal, Decimal, dict[str, Decimal]]:
    cuerpo = _solo_original(texto)
    ivas: dict[str, Decimal] = {}

    def _pick(labels: tuple[str, ...], *, excluir: tuple[str, ...] = ()) -> Decimal:
        for lab in labels:
            for m in re.finditer(
                lab + r"\s*:?\s*\$?\s*(" + RE_MONEY.pattern + r")",
                cuerpo,
                re.I,
            ):
                ctx = _norm(m.group(0))
                if any(e in ctx for e in excluir):
                    continue
                return money(m.group(1))
        return Decimal("0.00")

    excl = ("PERCEP", "RETENC", "SIRCREB", "IIBB", "ING BRUTOS")
    neto = _pick(
        (
            r"Importe\s+neto\s+gravado",
            r"Importe\s+Neto\s+Gravado",
            r"Neto\s+Gravado",
            r"Importe\s+Subtotal",
        ),
        excluir=excl,
    )
    iva = Decimal("0.00")
    for m in re.finditer(
        r"IVA\s*(10[.,]5|21|27)\s*%\s*:?\s*\$?\s*(" + RE_MONEY.pattern + r")",
        cuerpo,
        re.I,
    ):
        ali = _fmt_alicuota(m.group(1))
        mon = money(m.group(2))
        if mon <= 0:
            continue
        ivas[ali] = ivas.get(ali, Decimal("0.00")) + mon
        iva += mon
    if iva <= 0:
        iva = _pick((r"IVA\s+Contenido", r"D[eé]bito\s+Fiscal"), excluir=excl + ("IVA 0",))
        if iva > 0:
            ivas["Contenido"] = iva
    total = _pick((r"Importe\s+Total", r"Total\s+a\s+Pagar"), excluir=excl)
    if total <= 0 and neto > 0:
        total = (neto + iva).quantize(_MONEY)
    if neto <= 0 and total > 0 and iva > 0 and total > iva:
        neto = (total - iva).quantize(_MONEY)
    elif neto <= 0 and total > 0:
        neto = total
    return neto, iva, total, ivas


# ── Líneas de detalle ─────────────────────────────────────────────────────────

_UNIDADES = (
    r"unidades|unidad|kg|kgs|mts?|mtrs?|horas?|hs|litros?|lt|u\.?m\.?|pack|caja|cajas|metros?"
)

_RE_LINEA_A = re.compile(
    rf"^(?P<cant>\d+(?:[.,]\d+)?)\s+(?:{_UNIDADES})\s+"
    rf"(?P<concepto>.*?)\s*"
    rf"\$\s*(?P<precio>{RE_MONEY.pattern})\s+"
    rf"(?P<pct>\d+(?:[.,]\d+)?)\s*%\s+"
    rf"\$\s*(?P<bonif>{RE_MONEY.pattern})\s+"
    rf"(?P<iva_pct>\d+(?:[.,]\d+)?)\s*%\s+"
    rf"\$\s*(?P<subtotal>{RE_MONEY.pattern})\s*$",
    re.I,
)

_RE_LINEA_B = re.compile(
    rf"^(?P<cant>\d+(?:[.,]\d+)?)\s+(?:{_UNIDADES})\s+"
    rf"(?P<concepto>.*?)\s*"
    rf"\$\s*(?P<precio>{RE_MONEY.pattern})\s+"
    rf"(?P<pct>\d+(?:[.,]\d+)?)\s*%\s+"
    rf"\$\s*(?P<bonif>{RE_MONEY.pattern})\s+"
    rf"\$\s*(?P<subtotal>{RE_MONEY.pattern})\s*$",
    re.I,
)

_RE_LINEA_SIMPLE = re.compile(
    rf"^(?P<cant>\d+(?:[.,]\d+)?)\s+(?:{_UNIDADES})\s+"
    rf"(?P<concepto>.*?)\s*"
    rf"\$\s*(?P<precio>{RE_MONEY.pattern})\s+"
    rf"\$\s*(?P<subtotal>{RE_MONEY.pattern})\s*$",
    re.I,
)

_BLOQUEADOS = (
    "SUBTOTAL", "IMPORTE NETO", "IMPORTE TOTAL", "IMPORTE SUBTOTAL",
    "BONIFICACION GENERAL", "BONIFICACIÓN GENERAL", "CAE", "PAGINA",
    "CODIGO PRODUCTO", "PRODUCTO / SERVICIO", "PRODUCTO SERVICIO",
    "CANTIDAD", "PRECIO UNIT", "PERCEP", "RETENC", "IVA CONTENIDO",
    "OTROS TRIBUTOS", "REGIMEN DE TRANSPARENCIA", "COMPROBANTE AUTORIZADO",
)


def _concepto_limpio(concepto: str) -> str:
    c = re.sub(r"\s+", " ", (concepto or "").strip())
    c = re.sub(r"\$\s*$", "", c).strip()
    return c


def _codigo_producto(concepto: str) -> str:
    m = re.search(r"\[(\d+)\]", concepto or "")
    return m.group(1) if m else ""


def extraer_lineas_concepto(texto: str) -> list[dict]:
    cuerpo = _solo_original(texto)
    for corte in (
        r"Bonificaci[oó]n General",
        r"Regimen de Transparencia",
        r"Importe\s+neto\s+gravado",
        r"Importe Neto Gravado",
        r"Importe Subtotal",
        r"Importe Total",
        r"IVA Contenido",
        r"CAE N",
        r"Comprobante Autorizado",
        r"COMPROBANTES?\s+ASOCIADOS?",
    ):
        m = re.search(corte, cuerpo, re.I)
        if m and m.start() > 150:
            cuerpo = cuerpo[: m.start()]
            break

    m_tab = re.search(
        r"Cantidad\s+Producto\s*/\s*Servicio|Producto\s*/\s*Servicio|C[oó]digo\s+Producto",
        cuerpo,
        re.I,
    )
    if m_tab:
        cuerpo = cuerpo[m_tab.end() :]

    raw_lines = [ln.strip() for ln in cuerpo.splitlines() if ln.strip()]
    lineas: list[dict] = []
    pending: list[str] = []

    def _flush_pending() -> str:
        nonlocal pending
        txt = _concepto_limpio(" ".join(pending))
        pending = []
        return txt

    i = 0
    while i < len(raw_lines):
        line = raw_lines[i]
        if any(b in _norm(line) for b in ("PERCEP", "RETENC", "SIRCREB", "IIBB")):
            i += 1
            continue

        if re.match(r"^\[\d+\]", line) and not re.search(
            rf"\d\s+(?:{_UNIDADES})\s+", line, re.I
        ):
            if not re.search(r"\$\s*" + RE_MONEY.pattern, line):
                pending.append(line)
                i += 1
                continue

        matched = None
        for rex in (_RE_LINEA_A, _RE_LINEA_B, _RE_LINEA_SIMPLE):
            m = rex.match(line)
            if m:
                matched = m
                break

        if matched:
            concepto = _concepto_limpio(matched.group("concepto") or "")
            if not concepto and pending:
                concepto = _flush_pending()
            elif pending:
                pref = _flush_pending()
                if pref:
                    concepto = _concepto_limpio(f"{pref} {concepto}".strip())

            if i + 1 < len(raw_lines):
                nxt = raw_lines[i + 1]
                if (
                    not re.match(rf"^\d+(?:[.,]\d+)?\s+(?:{_UNIDADES})\b", nxt, re.I)
                    and not re.match(r"^\[\d+\]", nxt)
                    and "$" not in nxt
                    and not re.search(
                        r"Bonificaci|Importe|IVA|CAE|Subtotal|Regimen|Comprobante",
                        nxt,
                        re.I,
                    )
                    and len(nxt) <= 80
                ):
                    concepto = _concepto_limpio(f"{concepto} {nxt}".strip())
                    i += 1

            subtotal = money(matched.group("subtotal"))
            alicuota = ""
            if "iva_pct" in matched.groupdict() and matched.group("iva_pct"):
                alicuota = _fmt_alicuota(matched.group("iva_pct"))
            if (
                subtotal > 0
                and len(concepto) >= 3
                and not any(b in _norm(concepto) for b in _BLOQUEADOS)
            ):
                lineas.append(
                    {
                        "concepto": concepto,
                        "codigo": _codigo_producto(concepto),
                        "subtotal": subtotal,
                        "alicuota": alicuota,
                    }
                )
            pending = []
            i += 1
            continue

        if (
            not re.search(r"Bonificaci|Importe|IVA|CAE|Regimen|Comprobante|P[aá]g\.", line, re.I)
            and len(line) >= 3
            and not re.match(r"^\d{2}/\d{2}/\d{4}", line)
        ):
            if "$" not in line or re.match(r"^\[\d+\]", line):
                pending.append(re.sub(r"\$\s*" + RE_MONEY.pattern, "", line).strip())
        i += 1

    out: list[dict] = []
    for it in lineas:
        if (
            out
            and out[-1]["concepto"] == it["concepto"]
            and abs(out[-1]["subtotal"] - it["subtotal"]) < Decimal("0.02")
        ):
            continue
        out.append(it)
    return out


# ── Cálculo neto / IVA / bruto ────────────────────────────────────────────────

def _aplicar_bonificacion(
    lineas: list[dict], bonif: Decimal
) -> list[dict]:
    """Prorratea Bonificación General sobre subtotales impresos."""
    if not lineas:
        return []
    suma = sum((x["subtotal"] for x in lineas), Decimal("0.00"))
    if bonif <= 0 or suma <= 0:
        for ln in lineas:
            ln["subtotal_calc"] = ln["subtotal"]
        return lineas
    ratio = bonif / suma
    for ln in lineas:
        ln["subtotal_calc"] = (
            ln["subtotal"] * (Decimal("1") - ratio)
        ).quantize(_MONEY, rounding=ROUND_HALF_UP)
    # Ajuste centavos: suma calc = suma - bonif
    objetivo = (suma - bonif).quantize(_MONEY)
    actual = sum((x["subtotal_calc"] for x in lineas), Decimal("0.00"))
    delta = (objetivo - actual).quantize(_MONEY)
    if delta != 0 and lineas:
        lineas[-1]["subtotal_calc"] = (lineas[-1]["subtotal_calc"] + delta).quantize(
            _MONEY
        )
    return lineas


def calcular_importes_lineas(
    lineas: list[dict],
    *,
    tipo: str,
    bonif: Decimal,
    mapa_alicuotas: dict[str, Decimal],
) -> list[dict]:
    """
    Factura A: subtotal = neto; IVA = neto × alícuota.
    Factura B/NC: subtotal = bruto; neto = bruto / (1+ali/100).
    """
    lineas = _aplicar_bonificacion([dict(x) for x in lineas], bonif)
    es_a = es_tipo_a(tipo)

    for ln in lineas:
        sub = ln.get("subtotal_calc", ln["subtotal"])
        codigo = ln.get("codigo") or _codigo_producto(ln.get("concepto", ""))
        ln["codigo"] = codigo
        estimada = "NO"

        if es_a and ln.get("alicuota"):
            ali_pct = _parse_alicuota_pct(ln["alicuota"])
        else:
            ali_pct = Decimal("0")
            if codigo and codigo in mapa_alicuotas:
                ali_pct = mapa_alicuotas[codigo]
            else:
                hint = _alicuota_desde_nombre(ln.get("concepto", ""))
                if hint is not None:
                    ali_pct = hint
                else:
                    ali_pct = ALICUOTA_DEFAULT
                    estimada = "SI"
            ln["alicuota"] = _fmt_alicuota(ali_pct)

        if es_a:
            neto = sub
            iva = (neto * ali_pct / Decimal(100)).quantize(
                _MONEY, rounding=ROUND_HALF_UP
            )
            bruto = (neto + iva).quantize(_MONEY, rounding=ROUND_HALF_UP)
            estimada = "NO"
        else:
            # B / NC: subtotal es bruto con IVA incluido
            bruto = sub
            denom = Decimal("1") + (ali_pct / Decimal(100))
            neto = (bruto / denom).quantize(_MONEY, rounding=ROUND_HALF_UP)
            iva = (bruto - neto).quantize(_MONEY, rounding=ROUND_HALF_UP)

        ln["neto"] = neto
        ln["iva"] = iva
        ln["bruto"] = bruto
        ln["alicuota_estimada"] = estimada
    return lineas


def aprender_alicuotas_de_lineas(
    lineas: list[dict], mapa: dict[str, Decimal]
) -> dict[str, Decimal]:
    for ln in lineas:
        codigo = ln.get("codigo") or _codigo_producto(ln.get("concepto", ""))
        ali = _parse_alicuota_pct(ln.get("alicuota"))
        if codigo and ali > 0:
            mapa[codigo] = ali
        hint = _alicuota_desde_nombre(ln.get("concepto", ""))
        if codigo and hint is not None and codigo not in mapa:
            mapa[codigo] = hint
    return mapa


# ── Clasificación ─────────────────────────────────────────────────────────────

def _lineas_concepto_lista(text: str) -> list[str]:
    out: list[str] = []
    for line in (text or "").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if _norm(s) in {
            "DETALLE DE ITEMS VENDIDOS",
            "CONCEPTO",
            "PRODUCTO",
            "MERCADERIA",
            "SERVICIOS",
        }:
            continue
        s = re.sub(
            r"^(m:|s:|mercader[ií]a|servicios?)\s*:?\s*",
            "",
            s,
            flags=re.I,
        ).strip()
        if s:
            out.append(s)
    return out


def _leer_lista(fuente: Any) -> list[str]:
    if fuente is None:
        return []
    if isinstance(fuente, (str, Path)):
        p = Path(fuente)
        if not p.exists():
            return []
        data = p.read_bytes()
        nombre = p.name.lower()
    else:
        nombre = str(getattr(fuente, "name", "") or "").lower()
        data = fuente.read() if hasattr(fuente, "read") else fuente
        if hasattr(fuente, "seek"):
            try:
                fuente.seek(0)
            except Exception:
                pass
        if isinstance(data, str):
            data = data.encode("utf-8")
    if not isinstance(data, (bytes, bytearray)):
        return []

    if nombre.endswith(".pdf") or data[:4] == b"%PDF":
        return _lineas_concepto_lista(extraer_texto_pdf(bytes(data)))

    if nombre.endswith(".txt") or (
        not nombre.endswith((".xlsx", ".xls", ".csv")) and b"\x00" not in data[:20]
    ):
        return _lineas_concepto_lista(data.decode("utf-8", errors="replace"))

    bio = io.BytesIO(bytes(data))
    df = pd.read_csv(bio) if nombre.endswith(".csv") else pd.read_excel(bio)
    col = df.columns[0]
    for c in df.columns:
        if any(
            k in _norm(str(c))
            for k in ("CONCEPTO", "PRODUCTO", "DETALLE", "NOMBRE", "SERVIC", "MERC")
        ):
            col = c
            break
    return [str(x).strip() for x in df[col].dropna().tolist() if str(x).strip()]


def clasificar(
    concepto: str,
    mercaderia: Sequence[str],
    servicios: Sequence[str],
    *,
    umbral: float = UMBRAL_CLASIF,
) -> tuple[str, float, float]:
    """
    Retorna (etiqueta, score_merc, score_serv).
    Etiqueta: Mercadería | Servicio | PENDIENTE (probable X score%).
    """
    try:
        from rapidfuzz import fuzz
    except Exception:
        fuzz = None

    nc = _norm(concepto)
    nc_sin = _norm(re.sub(r"\[\d+\]", "", concepto))

    def best(lista: Sequence[str]) -> float:
        score = 0.0
        for item in lista:
            ni = _norm(item)
            if not ni:
                continue
            if ni in nc or ni in nc_sin or nc_sin in ni or nc in ni:
                return 100.0
            if fuzz:
                score = max(
                    score,
                    float(fuzz.token_set_ratio(nc_sin or nc, ni)),
                    float(fuzz.partial_ratio(nc_sin or nc, ni)),
                )
        return score

    sm = best(mercaderia)
    ss = best(servicios)
    if sm >= umbral and sm >= ss:
        return "Mercadería", sm, ss
    if ss >= umbral and ss > sm:
        return "Servicio", sm, ss
    if sm >= ss and sm > 0:
        return f"PENDIENTE (probable Mercadería {sm:.0f}%)", sm, ss
    if ss > 0:
        return f"PENDIENTE (probable Servicio {ss:.0f}%)", sm, ss
    return "PENDIENTE", sm, ss


def _es_lista_maestra(nombre: str) -> bool:
    nom_u = (nombre or "").upper()
    if nom_u in {"MERCADERIAS.PDF", "MERCADERÍAS.PDF", "SERVICIOS.PDF"}:
        return True
    if "MERCADER" in nom_u and nom_u.endswith(".PDF") and "FACTURA" not in nom_u:
        return True
    if nom_u.startswith("SERVICIOS") and nom_u.endswith((".PDF", ".XLSX", ".CSV", ".TXT")):
        return True
    return False


def _bucket_tipo_nombre(nombre: str) -> str:
    t = detectar_tipo_desde_nombre(nombre) or ""
    tu = t.upper()
    if "CRÉDITO" in tu or "CREDITO" in tu:
        return "NC"
    if "FACTURA B" in tu or tu.endswith(" B"):
        return "B"
    if "FACTURA A" in tu or tu.endswith(" A"):
        return "A"
    if "FACTURA C" in tu or tu.endswith(" C"):
        return "C"
    return "OTRO"


def _muestra_piloto_mixta(
    packed: list[tuple[str, bytes]], limite: int
) -> list[tuple[str, bytes]]:
    """Toma hasta ``limite`` PDFs mezclando A/B/NC (no los primeros N alfabéticos)."""
    if limite <= 0 or len(packed) <= limite:
        return packed
    buckets: dict[str, list[tuple[str, bytes]]] = {
        "A": [],
        "B": [],
        "C": [],
        "NC": [],
        "OTRO": [],
    }
    for item in packed:
        buckets[_bucket_tipo_nombre(item[0])].append(item)

    # Cupos proporcionales con piso para no dejar afuera B/NC
    orden = ["A", "B", "C", "NC", "OTRO"]
    no_vacios = [k for k in orden if buckets[k]]
    if not no_vacios:
        return packed[:limite]
    base = max(1, limite // len(no_vacios))
    cupos = {k: min(len(buckets[k]), base) for k in no_vacios}
    usados = sum(cupos.values())
    # Repartir sobrante priorizando B (mayoría real) y luego A
    for k in ("B", "A", "NC", "C", "OTRO"):
        if usados >= limite or k not in cupos:
            continue
        extra = min(len(buckets[k]) - cupos[k], limite - usados)
        if extra > 0:
            cupos[k] += extra
            usados += extra
    out: list[tuple[str, bytes]] = []
    for k in orden:
        out.extend(buckets[k][: cupos.get(k, 0)])
    # Si aún falta (poco de algún tipo), completar con el resto en orden
    if len(out) < limite:
        ya = {id(x) for x in out}
        for item in packed:
            if len(out) >= limite:
                break
            if id(item) in ya:
                continue
            out.append(item)
    return sorted(out, key=lambda x: x[0].lower())[:limite]


# ── Pipeline ──────────────────────────────────────────────────────────────────

def _leer_entrada(item: Any) -> tuple[str, bytes]:
    if isinstance(item, (str, Path)):
        p = Path(item)
        return p.name, p.read_bytes()
    if isinstance(item, tuple) and len(item) == 2:
        return str(item[0]), item[1] if isinstance(item[1], (bytes, bytearray)) else bytes(item[1])
    nombre = str(getattr(item, "name", "factura.pdf"))
    data = item.read() if hasattr(item, "read") else item
    if hasattr(item, "seek"):
        try:
            item.seek(0)
        except Exception:
            pass
    if isinstance(data, str):
        data = data.encode("utf-8")
    return nombre, bytes(data)


def procesar_un_pdf(
    nombre: str,
    data: bytes,
    *,
    mercaderia: Sequence[str] | None = None,
    servicios: Sequence[str] | None = None,
    mapa_alicuotas: dict[str, Decimal] | None = None,
    umbral_clasif: float = UMBRAL_CLASIF,
) -> dict:
    mapa = dict(mapa_alicuotas or {})
    texto = extraer_texto_pdf(data)
    if not texto.strip():
        return {
            "ok": False,
            "requiere_ocr": True,
            "prueba_tecnica": False,
            "filas": [],
            "error": "PDF sin texto (escaneado) — requiere OCR/manual",
            "meta": {"archivo": nombre},
            "lineas_raw": [],
            "tipo": "",
        }

    tipo = detectar_tipo(texto, nombre_archivo=nombre)
    fecha = extraer_fecha(texto)
    pv, nro = extraer_pv_numero(texto)
    nro_comp = nro_comprobante(pv, nro)
    condicion = extraer_condicion_iva(texto)
    receptor = extraer_receptor(texto)
    asociado = extraer_comprobante_asociado(texto) if es_nota_credito(tipo) else ""
    bonif = bonificacion_general(texto)
    _neto_doc, _iva_doc, total_doc, _ivas = totales_pie(texto)
    items = extraer_lineas_concepto(texto)
    prueba = es_prueba_tecnica(texto, receptor)

    meta_base = {
        "archivo": nombre,
        "tipo": tipo,
        "fecha": fecha,
        "pv": pv,
        "nro": nro,
        "nro_comprobante": nro_comp,
        "condicion_iva": condicion,
        "receptor": receptor,
        "prueba_tecnica": prueba,
        "bonificacion": _f(bonif),
        "total_pdf": _f(total_doc),
    }

    if not items:
        return {
            "ok": False,
            "requiere_ocr": False,
            "prueba_tecnica": prueba,
            "filas": [],
            "error": "No se leyeron líneas de detalle",
            "meta": meta_base,
            "lineas_raw": [],
            "tipo": tipo,
        }

    if es_tipo_a(tipo):
        aprender_alicuotas_de_lineas(items, mapa)

    items = calcular_importes_lineas(
        items, tipo=tipo, bonif=bonif, mapa_alicuotas=mapa
    )
    if es_tipo_a(tipo):
        aprender_alicuotas_de_lineas(items, mapa)

    merc = list(mercaderia or [])
    serv = list(servicios or [])
    signo = Decimal("-1") if es_nota_credito(tipo) else Decimal("1")

    suma_bruto = sum((x["bruto"] for x in items), Decimal("0.00"))
    # Comparar contra total impreso en valor absoluto (NC se validan antes del signo)
    revisar = ""
    if total_doc > 0 and abs(suma_bruto - total_doc) > TOL:
        revisar = "revisar"

    filas = []
    for it in items:
        etiqueta, _sm, _ss = (
            clasificar(it["concepto"], merc, serv, umbral=umbral_clasif)
            if (merc or serv)
            else ("PENDIENTE", 0.0, 0.0)
        )
        filas.append(
            {
                "Fecha": fecha,
                "Tipo Comprobante": tipo,
                "Nro. Comprobante": nro_comp,
                "Receptor": receptor,
                "Condición IVA Receptor": condicion,
                "Concepto": it["concepto"],
                "Código Producto": it.get("codigo") or "",
                "Neto": _f((it["neto"] * signo).quantize(_MONEY)),
                "IVA": _f((it["iva"] * signo).quantize(_MONEY)),
                "Alícuota IVA": it.get("alicuota") or "",
                "Bruto": _f((it["bruto"] * signo).quantize(_MONEY)),
                "Servicio/Mercadería": etiqueta,
                "Alícuota Estimada": it.get("alicuota_estimada") or "NO",
                "Comprobante Asociado": asociado,
                "Archivo": nombre,
                "Revisar": revisar,
                "Prueba técnica": "SI" if prueba else "",
            }
        )

    return {
        "ok": True,
        "requiere_ocr": False,
        "prueba_tecnica": prueba,
        "filas": filas,
        "error": None,
        "meta": {
            **meta_base,
            "lineas": len(filas),
            "suma_lineas": _f(suma_bruto),
            "revisar": bool(revisar),
        },
        "lineas_raw": items,
        "tipo": tipo,
        "mapa_alicuotas": mapa,
    }


def procesar_carpeta_o_uploads(
    entradas: Iterable[Any],
    *,
    lista_mercaderia: Any = None,
    lista_servicios: Any = None,
    mapa_alicuotas: dict[str, Decimal] | None = None,
    persistir_alicuotas: bool = True,
    path_alicuotas: Path | None = None,
    limite: int | None = None,
    umbral_clasif: float = UMBRAL_CLASIF,
) -> dict:
    """
    Procesa PDFs en dos pasadas:
      1) Facturas A → aprende código→alícuota
      2) Todos los comprobantes con el mapa completo
    ``limite`` recorta el lote (piloto 30-50) de forma estable por nombre.
    """
    merc = _leer_lista(lista_mercaderia)
    serv = _leer_lista(lista_servicios)
    mapa = dict(mapa_alicuotas) if mapa_alicuotas is not None else cargar_mapa_alicuotas(path_alicuotas)

    packed: list[tuple[str, bytes]] = []
    for item in entradas or []:
        nombre, data = _leer_entrada(item)
        if _es_lista_maestra(nombre):
            continue
        packed.append((nombre, data))

    # Idempotencia: un archivo = una vez (último gana si hay duplicados de nombre)
    by_name: dict[str, bytes] = {}
    for nom, data in packed:
        by_name[nom] = data
    packed = sorted(by_name.items(), key=lambda x: x[0].lower())

    if limite is not None and limite > 0:
        # Piloto mezclado por tipo (si se corta solo por nombre, salen todas A:
        # FACTURA_A-… ordena antes que FACTURA_B-…).
        packed = _muestra_piloto_mixta(packed, int(limite))

    encontrados = len(packed)

    # Pasada 1: aprender alícuotas de Facturas A
    for nombre, data in packed:
        texto = extraer_texto_pdf(data)
        if not texto.strip():
            continue
        tipo = detectar_tipo(texto, nombre_archivo=nombre)
        if not es_tipo_a(tipo):
            continue
        items = extraer_lineas_concepto(texto)
        aprender_alicuotas_de_lineas(items, mapa)

    if persistir_alicuotas and mapa:
        guardar_mapa_alicuotas(mapa, path_alicuotas)

    # Pasada 2: procesar todo
    todas: list[dict] = []
    fallidos: list[dict] = []
    ocr: list[dict] = []
    pruebas: list[dict] = []
    ok_n = revisar_n = 0
    vistos_comp: set[str] = set()

    for nombre, data in packed:
        res = procesar_un_pdf(
            nombre,
            data,
            mercaderia=merc,
            servicios=serv,
            mapa_alicuotas=mapa,
            umbral_clasif=umbral_clasif,
        )
        if res.get("mapa_alicuotas"):
            mapa.update(res["mapa_alicuotas"])

        if res.get("requiere_ocr"):
            ocr.append({"archivo": nombre, "motivo": res.get("error")})
            continue
        if not res.get("ok"):
            fallidos.append({"archivo": nombre, "motivo": res.get("error") or "Error"})
            continue

        meta = res.get("meta") or {}
        # Idempotencia por tipo+número (A y B pueden compartir el mismo PV-nro)
        nro_c = meta.get("nro_comprobante") or ""
        tipo_c = meta.get("tipo") or ""
        clave = f"{tipo_c}|{nro_c}" if nro_c else nombre
        if clave in vistos_comp:
            continue
        vistos_comp.add(clave)

        ok_n += 1
        if meta.get("revisar"):
            revisar_n += 1
        if res.get("prueba_tecnica"):
            pruebas.append(
                {
                    "archivo": nombre,
                    "nro": meta.get("nro_comprobante"),
                    "receptor": meta.get("receptor"),
                    "total": meta.get("total_pdf"),
                }
            )
        todas.extend(res["filas"])

    if persistir_alicuotas and mapa:
        guardar_mapa_alicuotas(mapa, path_alicuotas)

    df = pd.DataFrame(todas)
    if not df.empty:
        cols = [c for c in COLUMNAS if c in df.columns]
        extras = [c for c in df.columns if c not in cols]
        df = df[cols + extras]

    n_estimada = int((df["Alícuota Estimada"] == "SI").sum()) if not df.empty and "Alícuota Estimada" in df.columns else 0
    n_pendiente = (
        int(df["Servicio/Mercadería"].astype(str).str.startswith("PENDIENTE").sum())
        if not df.empty and "Servicio/Mercadería" in df.columns
        else 0
    )
    tipos_filas = (
        df["Tipo Comprobante"].astype(str).value_counts().to_dict()
        if not df.empty and "Tipo Comprobante" in df.columns
        else {}
    )

    return {
        "detalle": df,
        "resumen": {
            "pdfs_encontrados": encontrados,
            "pdfs_ok": ok_n,
            "pdfs_fallidos": len(fallidos),
            "pdfs_requiere_ocr": len(ocr),
            "facturas_revisar": revisar_n,
            "filas_detalle": len(df),
            "alicuota_estimada": n_estimada,
            "clasif_pendiente": n_pendiente,
            "pruebas_tecnicas": len(pruebas),
            "codigos_alicuota": len(mapa),
            "limite": limite,
            "tipos": tipos_filas,
        },
        "fallidos": fallidos,
        "requiere_ocr": ocr,
        "pruebas_tecnicas": pruebas,
        "mapa_alicuotas": {k: str(v) for k, v in sorted(mapa.items())},
    }


def exportar_excel_bytes(resultado: dict, *, titulo: str = "", subtitulo: str = "") -> bytes:
    df = resultado.get("detalle")
    if df is None or df.empty:
        detalle = pd.DataFrame(columns=COLUMNAS)
    else:
        cols = [c for c in COLUMNAS if c in df.columns]
        extras = [c for c in df.columns if c not in cols and c != "Prueba técnica"]
        detalle = df[cols + extras].copy()

    r = resultado.get("resumen") or {}
    resumen_df = pd.DataFrame(
        [
            {"Métrica": "PDFs encontrados", "Valor": r.get("pdfs_encontrados", 0)},
            {"Métrica": "PDFs OK", "Valor": r.get("pdfs_ok", 0)},
            {"Métrica": "Filas de detalle", "Valor": r.get("filas_detalle", 0)},
            {"Métrica": "No cerraron vs total", "Valor": r.get("facturas_revisar", 0)},
            {"Métrica": "Alícuota estimada (SI)", "Valor": r.get("alicuota_estimada", 0)},
            {"Métrica": "Clasificación PENDIENTE", "Valor": r.get("clasif_pendiente", 0)},
            {"Métrica": "Pruebas técnicas", "Valor": r.get("pruebas_tecnicas", 0)},
            {"Métrica": "Códigos en tabla alícuota", "Valor": r.get("codigos_alicuota", 0)},
            {"Métrica": "Requieren OCR", "Valor": r.get("pdfs_requiere_ocr", 0)},
            {"Métrica": "No leídos", "Valor": r.get("pdfs_fallidos", 0)},
        ]
    )

    hojas: list[tuple[str, pd.DataFrame]] = []
    if not detalle.empty and "Servicio/Mercadería" in detalle.columns:
        clase = detalle["Servicio/Mercadería"].astype(str)
        for nombre, mask in (
            ("Mercadería", clase == "Mercadería"),
            ("Servicios", clase == "Servicio"),
            ("PENDIENTE", clase.str.startswith("PENDIENTE")),
        ):
            sub = detalle.loc[mask].copy()
            hojas.append((nombre, sub if not sub.empty else pd.DataFrame(columns=COLUMNAS)))

    if resultado.get("fallidos"):
        hojas.append(("No leídos", pd.DataFrame(resultado["fallidos"])))
    if resultado.get("requiere_ocr"):
        hojas.append(("Requiere OCR", pd.DataFrame(resultado["requiere_ocr"])))
    if resultado.get("pruebas_tecnicas"):
        hojas.append(("Pruebas técnicas", pd.DataFrame(resultado["pruebas_tecnicas"])))

    mapa = resultado.get("mapa_alicuotas") or {}
    if mapa:
        hojas.append(
            (
                "Código→Alícuota",
                pd.DataFrame(
                    [{"Código Producto": k, "Alícuota %": v} for k, v in mapa.items()]
                ),
            )
        )

    return exportar_informe_excel(
        titulo=titulo or "Detalle de ítems — Facturas de venta",
        subtitulo=subtitulo
        or "Fecha | Tipo | Nro | Receptor | Concepto | Neto | IVA | Bruto | Clasificación",
        periodo=date.today().strftime("%d/%m/%Y"),
        kpis=[
            ("PDFs OK", r.get("pdfs_ok", 0)),
            ("Filas", r.get("filas_detalle", 0)),
            ("Revisar", r.get("facturas_revisar", 0)),
            ("Pendiente", r.get("clasif_pendiente", 0)),
            ("Aliq. est.", r.get("alicuota_estimada", 0)),
        ],
        resumenes=[("Resumen", resumen_df)],
        detalle=detalle,
        hoja_detalle="Detalle",
        hojas_adicionales=hojas or None,
        col_moneda=["Neto", "IVA", "Bruto"],
        col_fecha=["Fecha"],
        col_texto=["Nro. Comprobante", "Código Producto", "Alícuota IVA"],
        total_col="Bruto",
    )


def guardar_excel(
    resultado: dict,
    ruta: str | Path,
    *,
    titulo: str = "",
    subtitulo: str = "",
) -> Path:
    """Guarda el informe en disco con formato del estudio."""
    ruta = Path(ruta)
    df = resultado.get("detalle")
    if df is None or df.empty:
        detalle = pd.DataFrame(columns=COLUMNAS)
    else:
        cols = [c for c in COLUMNAS if c in df.columns]
        extras = [c for c in df.columns if c not in cols and c != "Prueba técnica"]
        detalle = df[cols + extras].copy()

    r = resultado.get("resumen") or {}
    resumen_df = pd.DataFrame(
        [
            {"Métrica": k, "Valor": v}
            for k, v in [
                ("PDFs encontrados", r.get("pdfs_encontrados", 0)),
                ("PDFs OK", r.get("pdfs_ok", 0)),
                ("Filas", r.get("filas_detalle", 0)),
                ("Revisar", r.get("facturas_revisar", 0)),
                ("Alícuota estimada", r.get("alicuota_estimada", 0)),
                ("PENDIENTE", r.get("clasif_pendiente", 0)),
            ]
        ]
    )
    hojas: list[tuple[str, pd.DataFrame]] = []
    if not detalle.empty and "Servicio/Mercadería" in detalle.columns:
        clase = detalle["Servicio/Mercadería"].astype(str)
        for nombre, mask in (
            ("Mercadería", clase == "Mercadería"),
            ("Servicios", clase == "Servicio"),
            ("PENDIENTE", clase.str.startswith("PENDIENTE")),
        ):
            sub = detalle.loc[mask].copy()
            hojas.append((nombre, sub if not sub.empty else pd.DataFrame(columns=COLUMNAS)))
    if resultado.get("pruebas_tecnicas"):
        hojas.append(("Pruebas técnicas", pd.DataFrame(resultado["pruebas_tecnicas"])))
    mapa = resultado.get("mapa_alicuotas") or {}
    if mapa:
        hojas.append(
            (
                "Código→Alícuota",
                pd.DataFrame(
                    [{"Código Producto": k, "Alícuota %": v} for k, v in mapa.items()]
                ),
            )
        )

    return guardar_informe_excel(
        ruta=ruta,
        titulo=titulo or "Detalle de ítems — Facturas de venta",
        subtitulo=subtitulo or "",
        periodo=date.today().strftime("%d/%m/%Y"),
        kpis=[
            ("PDFs OK", r.get("pdfs_ok", 0)),
            ("Filas", r.get("filas_detalle", 0)),
            ("Revisar", r.get("facturas_revisar", 0)),
            ("Pendiente", r.get("clasif_pendiente", 0)),
        ],
        resumenes=[("Resumen", resumen_df)],
        detalle=detalle,
        hoja_detalle="Detalle",
        hojas_adicionales=hojas or None,
        col_moneda=["Neto", "IVA", "Bruto"],
        col_fecha=["Fecha"],
        col_texto=["Nro. Comprobante", "Código Producto", "Alícuota IVA"],
        total_col="Bruto",
    )
