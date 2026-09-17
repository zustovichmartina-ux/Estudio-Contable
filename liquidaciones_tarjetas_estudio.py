# -*- coding: utf-8 -*-
"""Liquidaciones de tarjetas en el formato del estudio (Recife / plantilla).

PDFs Naranja, Favacard, Cabal y First Data/Fiserv → Excel con:
- LIQUIDACIONES (resumen por marca, fórmulas del estudio)
- First Data (una fila por liquidación)
- Movimientos (cada VENTA/cupón + cada LIQUIDACIÓN, todas las columnas)
- Control (neto calculado vs PDF)
"""
from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from procesadores.tarjeta_parser import (
    detectar_entidad_por_texto,
    extraer_con_plantilla,
)

COLOR_TITULO = "1F4E79"
COLOR_HEADER = "1F4E79"
COLOR_ZEBRA = "F2F2F2"
COLOR_TOTAL = "DCE6F1"
COLOR_BORDE = "BFBFBF"
TOL = 0.05
_FONT = "Calibri"
_NUM = "#,##0.00"
_THIN = Border(
    left=Side(style="thin", color=COLOR_BORDE),
    right=Side(style="thin", color=COLOR_BORDE),
    top=Side(style="thin", color=COLOR_BORDE),
    bottom=Side(style="thin", color=COLOR_BORDE),
)
_FILL_HDR = PatternFill("solid", fgColor=COLOR_HEADER)
_FILL_TTL = PatternFill("solid", fgColor=COLOR_TITULO)
_FILL_ZEBRA = PatternFill("solid", fgColor=COLOR_ZEBRA)
_FILL_TOT = PatternFill("solid", fgColor=COLOR_TOTAL)
_FONT_TTL = Font(name=_FONT, bold=True, color="FFFFFF", size=12)
_FONT_HDR = Font(name=_FONT, bold=True, color="FFFFFF", size=10)
_FONT_BODY = Font(name=_FONT, size=10)
_FONT_BOLD = Font(name=_FONT, bold=True, size=10)

RE_PAGO = re.compile(
    r"IMPORTE\s*NETO\s*DE\s*PAGOS\s*\$\s*([\d.]+,\d{2})(-)?"
    r".{0,250}?"
    r"el\s+d[ií]a\s+(\d{2}/\d{2}/\d{4})"
    r".{0,180}?"
    r"Nro\.?\s*Liq:\s*(\d+)",
    re.I | re.S,
)
RE_VENTA_CTDO = re.compile(
    r"Venta\s+ctdo\s+(\d{2}/\d{2}/\d{2})\s+\S+\s+\S+\s+\S+\s+\S+\s+"
    r"[\d,.]+\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})",
    re.I,
)
RE_VTA_CUO = re.compile(
    r"vta\s+\d+\s+cuo\s+(\d{2}/\d{2}/\d{2})\s+\S+\s+\S+\s+\S+\s+\S+\s+"
    r"Cup[^\s]*\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})",
    re.I,
)

HEADERS_MOV = [
    "Transacción",
    "Ventas c/descuento",
    "VENTAS C/DTO PLAN CUOTAS",
    "Dto. Arancel",
    "ARANCEL CUOTAS",
    "DTO S/VENTAS FIN ADQ CONT",
    "DTO S/VENTAS FIN ADQ CUOTA",
    "IVA ARANCEL CUOTAS 21,00%",
    "IVA S/DTO FIN ADQ CUOTA 21,00%",
    "IVA S/DTO FIN ADQ CONT 21,00%",
    "IVA CRED.FISC.COMERCIO S/ARANC 21,00%",
    "PERCEPCION IVA R.G. 2408 1,50 %",
    "PERCEPCION IVA R.G. 2408 3,00 %",
    "RETENCION ING.BRUTOS SIRTAC",
    "PER B.A.I.BR.DN.01/04",
    "QR PERC. IIBB. CABA REG GRAL",
    "QR PERCEPCION IVA 2408",
    "SERVICIO OPER. INTERNAC.",
    "IVA RI SERV.OPER. INT.",
    "IMPORTE NETO DE PAGOS",
    "Nro Liquidación",
    "Fecha pago",
    "Archivo",
]


def money(s: str | None, neg: str | None = None) -> float:
    if s is None:
        return 0.0
    t = str(s).strip().replace("$", "").replace(" ", "")
    if not t:
        return 0.0
    signo = -1 if (neg or t.endswith("-") or t.startswith("-") or t.startswith("(")) else 1
    t = t.replace("-", "").replace("(", "").replace(")", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return round(signo * float(t), 2)
    except ValueError:
        return 0.0


def texto_pdf(path: Path) -> str:
    with pdfplumber.open(str(path)) as pdf:
        return "\n".join((pg.extract_text() or "") for pg in pdf.pages)


def clasificar_pdf(nombre: str, texto: str) -> str:
    n = (nombre or "").upper()
    t = (texto or "").upper()
    blob = n + "\n" + t[:5000]
    if "FAVACARD" in blob or re.search(r"\bFAVA\b", n):
        return "FAVA"
    if "NARANJA" in blob:
        return "NARANJA"
    if "LIQUIDACION PARA PAGO A COMERCIOS" in blob or (
        "CABAL" in blob and "ARANCEL DE DESCUENTO" in blob
    ):
        return "CABAL"
    if "IMPORTE NETO DE PAGOS" in blob and (
        "VENTAS C/DESCUENTO" in blob
        or "VENTAS C/DTO" in blob
        or "FIRST DATA" in blob
        or "FISERV" in blob
        or "TARJETA DE CREDITO PESOS" in blob
        or "TARJETA DE DEBITO PESOS" in blob
    ):
        return "FIRST_DATA"
    return "OTRO"


@dataclass
class VentaCupon:
    fecha: str
    ventas: float
    arancel: float
    dto_fin: float


@dataclass
class LiqFD:
    archivo: str
    nro: str
    fecha: str
    ventas_contado: float = 0.0
    ventas_cuotas: float = 0.0
    arancel: float = 0.0
    arancel_cuotas: float = 0.0
    dto_cont: float = 0.0
    dto_cuota: float = 0.0
    iva_arancel: float = 0.0
    iva_arancel_cuotas: float = 0.0
    iva_dto_cont: float = 0.0
    iva_dto_cuota: float = 0.0
    perc_iva_15: float = 0.0
    perc_iva_30: float = 0.0
    sirtac: float = 0.0
    per_iibb_ba: float = 0.0
    perc_caba: float = 0.0
    perc_iva_qr: float = 0.0
    serv_int: float = 0.0
    iva_serv_int: float = 0.0
    otras: float = 0.0
    neto: float = 0.0
    cupones: list[VentaCupon] = field(default_factory=list)
    otras_det: list[str] = field(default_factory=list)

    @property
    def ventas(self) -> float:
        return round(self.ventas_contado + self.ventas_cuotas, 2)

    @property
    def bi21(self) -> float:
        return round(
            self.arancel + self.arancel_cuotas + self.dto_cont + self.dto_cuota + self.serv_int,
            2,
        )

    @property
    def perc_iva(self) -> float:
        return round(self.perc_iva_15 + self.perc_iva_30 + self.perc_iva_qr, 2)

    @property
    def ret_iibb(self) -> float:
        return round(self.sirtac, 2)

    @property
    def deducciones(self) -> float:
        return round(
            self.arancel
            + self.arancel_cuotas
            + self.dto_cont
            + self.dto_cuota
            + self.iva_arancel
            + self.iva_arancel_cuotas
            + self.iva_dto_cont
            + self.iva_dto_cuota
            + self.perc_iva_15
            + self.perc_iva_30
            + self.sirtac
            + self.per_iibb_ba
            + self.perc_caba
            + self.perc_iva_qr
            + self.serv_int
            + self.iva_serv_int
            + self.otras,
            2,
        )

    @property
    def neto_calc(self) -> float:
        return round(self.ventas - self.deducciones, 2)


@dataclass
class ResultadoTarjetasEstudio:
    ok: bool
    excel_bytes: bytes = b""
    nombre_archivo: str = "Liquidaciones_de_Tarjetas.xlsx"
    mensajes: list[str] = field(default_factory=list)
    advertencias: list[str] = field(default_factory=list)
    n_pdf: int = 0
    n_fd: int = 0
    n_ventas: int = 0
    n_control_ok: int = 0
    n_control_diff: int = 0
    comercio: str = ""
    periodo: str = ""
    error: str = ""
    hojas: list[str] = field(default_factory=list)


def _aplicar_deduccion(liq: LiqFD, concepto: str, monto: float) -> None:
    u = re.sub(r"\s+", " ", concepto.upper())
    if "ARANCEL CUOTAS" in u:
        liq.arancel_cuotas = round(liq.arancel_cuotas + monto, 2)
    elif re.search(r"\bARANCEL\b", u) and "IVA" not in u:
        liq.arancel = round(liq.arancel + monto, 2)
    elif "IVA CRED" in u or "S/ARANC" in u:
        liq.iva_arancel = round(liq.iva_arancel + monto, 2)
    elif "IVA ARANCEL CUOTAS" in u:
        liq.iva_arancel_cuotas = round(liq.iva_arancel_cuotas + monto, 2)
    elif "IVA S/DTO FIN ADQ CUOTA" in u:
        liq.iva_dto_cuota = round(liq.iva_dto_cuota + monto, 2)
    elif "IVA S/DTO FIN ADQ CONT" in u:
        liq.iva_dto_cont = round(liq.iva_dto_cont + monto, 2)
    elif "DTO S/VENTAS FIN ADQ CUOTA" in u:
        liq.dto_cuota = round(liq.dto_cuota + monto, 2)
    elif "DTO S/VENTAS FIN ADQ CONT" in u:
        liq.dto_cont = round(liq.dto_cont + monto, 2)
    elif "PERCEPCION IVA" in u and "1,50" in u:
        liq.perc_iva_15 = round(liq.perc_iva_15 + monto, 2)
    elif "PERCEPCION IVA" in u and "3,00" in u:
        liq.perc_iva_30 = round(liq.perc_iva_30 + monto, 2)
    elif "PERC IVA 2408" in u or "PERCEPCION IVA 2408" in u:
        liq.perc_iva_qr = round(liq.perc_iva_qr + monto, 2)
    elif "SIRTAC" in u:
        liq.sirtac = round(liq.sirtac + monto, 2)
    elif "PER B.A" in u or "PER BA" in u.replace(".", ""):
        liq.per_iibb_ba = round(liq.per_iibb_ba + monto, 2)
    elif "PERC.IB.CABA" in u or "IB.CABA" in u:
        liq.perc_caba = round(liq.perc_caba + monto, 2)
    elif "SERV.OPER" in u or "SERVICIO OPER" in u:
        if "IVA" in u:
            liq.iva_serv_int = round(liq.iva_serv_int + monto, 2)
        else:
            liq.serv_int = round(liq.serv_int + monto, 2)
    else:
        liq.otras = round(liq.otras + monto, 2)
        liq.otras_det.append(f"{concepto} {monto}")


def parsear_fd(path: Path, texto: str | None = None) -> list[LiqFD]:
    text = texto if texto is not None else texto_pdf(path)
    out: list[LiqFD] = []
    matches = list(RE_PAGO.finditer(text))
    for i, m in enumerate(matches):
        start = matches[i - 1].end() if i else 0
        bloque = text[start : m.start()]
        liq = LiqFD(
            archivo=path.name,
            nro=m.group(4),
            fecha=m.group(3),
            neto=money(m.group(1), m.group(2)),
        )
        for mv in re.finditer(
            r"\+\s*VENTAS\s+C/?DESCUENTO\s+CONTADO\s*\$\s*([\d.]+,\d{2})(-)?",
            bloque,
            re.I,
        ):
            liq.ventas_contado = round(liq.ventas_contado + money(mv.group(1), mv.group(2)), 2)
        for mv in re.finditer(
            r"\+\s*VENTAS\s+C/DTO\s+ANTIC POR FINANC ADQUIREN\s*\$\s*([\d.]+,\d{2})(-)?",
            bloque,
            re.I,
        ):
            liq.ventas_cuotas = round(liq.ventas_cuotas + money(mv.group(1), mv.group(2)), 2)
        for mv in re.finditer(
            r"\+\s*VENTAS\s+C/DTO\s+PLAN CUOTAS\s*\$\s*([\d.]+,\d{2})(-)?",
            bloque,
            re.I,
        ):
            liq.ventas_cuotas = round(liq.ventas_cuotas + money(mv.group(1), mv.group(2)), 2)
        for md in re.finditer(r"-\s*((?:(?!\$).){3,90}?)\$\s*([\d.]+,\d{2})(-)?", bloque):
            conc = re.sub(r"\s+", " ", md.group(1)).strip()
            if re.search(r"0800|www\.|Rep[uú]blica|Estimado|Centro de Atenci", conc, re.I):
                continue
            _aplicar_deduccion(liq, conc, money(md.group(2), md.group(3)))
        for mc in RE_VENTA_CTDO.finditer(bloque):
            liq.cupones.append(
                VentaCupon(mc.group(1), money(mc.group(2)), money(mc.group(3)), money(mc.group(4)))
            )
        for mc in RE_VTA_CUO.finditer(bloque):
            liq.cupones.append(
                VentaCupon(mc.group(1), money(mc.group(2)), money(mc.group(3)), money(mc.group(4)))
            )
        out.append(liq)
    return out


def parsear_naranja(path: Path, texto: str | None = None) -> dict[str, float | str]:
    t = texto if texto is not None else texto_pdf(path)

    def grab(pat: str) -> float:
        m = re.search(pat, t, re.I | re.S)
        return money(m.group(1)) if m else 0.0

    etiqueta = path.stem
    if re.search(r"\bGUE\b|GUEMES", path.name + t[:800], re.I):
        etiqueta = "Güemes"
    elif re.search(r"\bJBJ\b|JUSTO", path.name + t[:800], re.I):
        etiqueta = "JB Justo"
    return {
        "archivo": path.name,
        "etiqueta": etiqueta,
        "total": grab(r"Totales\s+1\s+\$\s*([\d.]+,\d{2})"),
        "arancel": grab(r"Arancel\s+-\s*\$\s*([\d.]+,\d{2})"),
        "interes": grab(r"Inter[eé]s\s+Plan[^\n]*\$\s*([\d.]+,\d{2})"),
        "iva": grab(r"IVA\s+21\.0\s*%\s+\$\s*([\d.]+,\d{2})"),
        "perc_iva": grab(r"Percepci[oó]n\s+de\s+IVA[^\n]*\$\s*([\d.]+,\d{2})"),
        "perc_iibb": grab(r"Percepci[oó]n\s+Ingresos?\s+Brutos[^\n]*\$\s*([\d.]+,\d{2})"),
        "sirtac": grab(r"Sirtac[^\n]*\$\s*([\d.]+,\d{2})"),
        "neto": grab(r"Neto Liquidado.*?Importe\s+\$\s*([\d.]+,\d{2})"),
    }


def parsear_fava(path: Path, texto: str | None = None) -> dict[str, float | str]:
    t = texto if texto is not None else texto_pdf(path)

    def grab(pat: str) -> float:
        m = re.search(pat, t, re.I)
        return money(m.group(1)) if m else 0.0

    etiqueta = path.stem
    if re.search(r"\bGUE\b|GUEMES", path.name + t[:800], re.I):
        etiqueta = "Güemes"
    elif re.search(r"\bJBJ\b|JUSTO", path.name + t[:800], re.I):
        etiqueta = "JB Justo"
    com = grab(r"Cargos y Bonific\.\s*Comisi[oó]n\s*:\s*-?\s*([\d.]+,\d{2})")
    dto = grab(r"Descuento Plan C\s*:\s*-?\s*([\d.]+,\d{2})")
    bon = grab(r"Bonificacion Plan C\s*:\s*([\d.]+,\d{2})")
    return {
        "archivo": path.name,
        "etiqueta": etiqueta,
        "total": grab(r"Total de la presentaci[oó]n\s*:\s*([\d.]+,\d{2})"),
        "comision": com,
        "dto": dto,
        "bonif": bon,
        "iva": grab(r"Impuestos Iva 21,00[^\n]*:\s*-?\s*([\d.]+,\d{2})"),
        "perc_iva": grab(r"Percepcion de IVA\s*:\s*-?\s*([\d.]+,\d{2})"),
        "perc_iibb": grab(r"Percepcion I\.Brutos PBA[^\n]*:\s*-?\s*([\d.]+,\d{2})"),
        "sirtac": grab(r"Ret\.\s*IIBB SIRTAC[^\n]*:\s*-?\s*([\d.]+,\d{2})"),
        "neto": grab(r"NETO A PAGAR PESOS\s*\$\s*([\d.]+,\d{2})"),
        "bi": round(com + dto - bon, 2),
    }


def parsear_cabal(path: Path, texto: str | None = None) -> list[dict[str, float | str]]:
    t = texto if texto is not None else texto_pdf(path)
    bloques = re.split(r"FECHA DE PAGO\s*:", t)
    out: list[dict[str, float | str]] = []
    for b in bloques[1:]:
        m_nro = re.search(r"LIQUIDACION NRO\.\s*:\s*(\d+)", b, re.I)
        if not m_nro:
            continue

        def grab(pat: str) -> float:
            m = re.search(pat, b, re.I)
            return money(m.group(1)) if m else 0.0

        m_ar = re.search(r"ARANCEL DE DESCUENTO\s+[\d.,]+%\s+([\d.]+,\d{2})(-)?", b, re.I)
        arancel = money(m_ar.group(1), m_ar.group(2)) if m_ar else 0.0
        cft_m = re.search(r"COSTO FINANCIERO TOTAL\s+([\d.]+,\d{2})(-)?", b, re.I)
        cft = money(cft_m.group(1), cft_m.group(2)) if cft_m else 0.0
        sirtac = 0.0
        m_sirtac = re.search(r"RETENCION IIBB SIRTAC[^\n]+", b, re.I)
        if m_sirtac:
            nums_s = re.findall(r"([\d.]+,\d{2})(-)?", m_sirtac.group(0))
            if nums_s:
                sirtac = money(nums_s[-1][0], nums_s[-1][1])
        tipo = "DEBITO" if "CABAL DEBITO" in b.upper() else "CREDITO"
        out.append(
            {
                "archivo": path.name,
                "nro": m_nro.group(1),
                "tipo": tipo,
                "etiqueta": f"Cabal {tipo} {m_nro.group(1)}",
                "total": grab(r"TOTAL DE VENTAS\.[^\n]*?([\d.]+,\d{2})"),
                "arancel": arancel,
                "cft": cft,
                "sirtac": sirtac,
                "neto": grab(r"IMPORTE NETO FINAL A LIQUIDAR\s*\.*\s*([\d.]+,\d{2})"),
            }
        )
    return out


def _fmt(cell) -> None:
    cell.number_format = _NUM
    cell.font = _FONT_BODY
    cell.border = _THIN


def _bi_formula(*montos: float) -> float | str | None:
    parts = [round(float(x), 2) for x in montos if abs(float(x or 0)) > 0.001]
    if not parts:
        return None
    if len(parts) == 1:
        return parts[0]
    return "=" + "+".join(str(p) for p in parts)


def _titulo_seccion(ws: Worksheet, row: int, texto: str, last_col: int = 10) -> None:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
    cell = ws.cell(row, 2, texto)
    cell.font = _FONT_TTL
    cell.fill = _FILL_TTL
    cell.alignment = Alignment(horizontal="left", vertical="center")


def _headers(ws: Worksheet, row: int, titulos: list[str], start_col: int = 2) -> None:
    for i, h in enumerate(titulos):
        cell = ws.cell(row, start_col + i, h)
        cell.font = _FONT_HDR
        cell.fill = _FILL_HDR
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = _THIN


def _escribir_liquidaciones(
    ws: Worksheet,
    naranja: list[dict],
    fava: list[dict],
    cabal: list[dict],
    liqs: list[LiqFD],
    comercio: str,
) -> None:
    ws["A1"] = comercio or "Liquidaciones de tarjetas"
    ws["A1"].font = Font(name=_FONT, bold=True, size=14, color=COLOR_TITULO)
    row = 2
    hdr_n = [
        "Total Liquidacion",
        "BI 21%",
        "IVA 21%",
        "Perc. IVA",
        "Perc. IIBB",
        "Ret. IIBB",
        "TOTAL",
        "Neto a Cobrar",
        "Dif",
    ]
    if naranja:
        _titulo_seccion(ws, row, "TARJETA NARANJA")
        row += 1
        _headers(ws, row, hdr_n)
        row += 1
        first = row
        for item in naranja:
            ws.cell(row, 1, item.get("etiqueta") or item.get("archivo"))
            ws.cell(row, 2, item["total"])
            interes = float(item.get("interes") or 0)
            arancel = float(item.get("arancel") or 0)
            ws.cell(row, 3, f"={arancel}+{interes}" if interes else arancel)
            ws.cell(row, 4, f"=+C{row}*0.21")
            ws.cell(row, 5, f"=+C{row}*0.03" if float(item.get("perc_iva") or 0) else 0)
            ws.cell(row, 6, f"=+C{row}*0.04" if float(item.get("perc_iibb") or 0) else 0)
            ws.cell(row, 7, item.get("sirtac") or 0)
            ws.cell(row, 8, item.get("neto") or 0)
            ws.cell(row, 9, f"=+B{row}-SUM(C{row}:G{row})")
            ws.cell(row, 10, f"=+H{row}-I{row}")
            for c in range(2, 11):
                _fmt(ws.cell(row, c))
            row += 1
        last = row - 1
        ws.cell(row, 2, f"=SUM(B{first}:B{last})")
        for col, letter in enumerate("CDEFGHIJ", 3):
            ws.cell(row, col, f"=SUM({letter}{first}:{letter}{last})")
        for c in range(2, 11):
            _fmt(ws.cell(row, c))
            ws.cell(row, c).font = _FONT_BOLD
            ws.cell(row, c).fill = _FILL_TOT
        row += 2

    if fava:
        _titulo_seccion(ws, row, "FAVACARD")
        row += 1
        _headers(ws, row, hdr_n)
        row += 1
        first = row
        for item in fava:
            ws.cell(row, 1, item.get("etiqueta") or item.get("archivo"))
            ws.cell(row, 2, item["total"])
            ws.cell(row, 3, f"={item['comision']}+{item['dto']}-{item['bonif']}")
            ws.cell(row, 4, f"=+C{row}*0.21")
            ws.cell(row, 5, item.get("perc_iva") or 0)
            ws.cell(row, 6, f"=+C{row}*0.04")
            ws.cell(row, 7, item.get("sirtac") or 0)
            ws.cell(row, 8, item.get("neto") or 0)
            ws.cell(row, 9, f"=+B{row}-SUM(C{row}:G{row})")
            ws.cell(row, 10, f"=+H{row}-I{row}")
            for c in range(2, 11):
                _fmt(ws.cell(row, c))
            row += 1
        last = row - 1
        ws.cell(row, 2, f"=SUM(B{first}:B{last})")
        for col, letter in enumerate("CDEFGHIJ", 3):
            ws.cell(row, col, f"=SUM({letter}{first}:{letter}{last})")
        for c in range(2, 11):
            _fmt(ws.cell(row, c))
            ws.cell(row, c).font = _FONT_BOLD
            ws.cell(row, c).fill = _FILL_TOT
        row += 2

    if cabal:
        hdr_c = [
            "Total Liquidacion",
            "BI 21%",
            "IVA 21%",
            "BI 10,5%",
            "IVA 10,5%",
            "RET. IIBB",
            "TOTAL",
            "Neto a Cobrar",
            "Dif",
        ]
        _titulo_seccion(ws, row, "CABAL SA")
        row += 1
        _headers(ws, row, hdr_c)
        row += 1
        first = row
        for item in cabal:
            ws.cell(row, 1, item.get("etiqueta") or item.get("nro"))
            ws.cell(row, 2, item["total"])
            ws.cell(row, 3, _bi_formula(float(item.get("arancel") or 0), float(item.get("cft") or 0)))
            ws.cell(row, 4, f"=+C{row}*0.21")
            ws.cell(row, 6, f"=+E{row}*0.105")
            ws.cell(row, 7, item.get("sirtac") or 0)
            ws.cell(row, 8, item.get("neto") or 0)
            ws.cell(row, 9, f"=+B{row}-SUM(C{row}:G{row})")
            ws.cell(row, 10, f"=+H{row}-I{row}")
            for c in range(2, 11):
                _fmt(ws.cell(row, c))
            row += 1
        last = row - 1
        ws.cell(row, 2, f"=SUM(B{first}:B{last})")
        for col, letter in enumerate("CDEFGH", 3):
            ws.cell(row, col, f"=SUM({letter}{first}:{letter}{last})")
        ws.cell(row, 9, f"=+B{row}-SUM(C{row}:G{row})")
        ws.cell(row, 10, f"=+H{row}-I{row}")
        for c in range(2, 11):
            _fmt(ws.cell(row, c))
            ws.cell(row, c).font = _FONT_BOLD
            ws.cell(row, c).fill = _FILL_TOT
        row += 2

    if liqs:
        hdr_nacion = [
            "Total Liquidacion",
            "BI 21%",
            "IVA 21%",
            "Perc. IVA",
            "Perc. IIBB",
            "Perc. CABA",
            "Ret. IIBB",
            "Neto a Cobrar",
            "Dif",
        ]
        _titulo_seccion(ws, row, "BANCO NACION / FIRST DATA")
        row += 1
        _headers(ws, row, hdr_nacion)
        row += 1
        ws.cell(row, 2, round(sum(x.ventas for x in liqs), 2))
        ws.cell(row, 3, round(sum(x.bi21 for x in liqs), 2))
        ws.cell(row, 4, f"=+C{row}*0.21")
        ws.cell(row, 5, round(sum(x.perc_iva for x in liqs), 2))
        ws.cell(row, 6, round(sum(x.per_iibb_ba for x in liqs), 2))
        ws.cell(row, 7, round(sum(x.perc_caba for x in liqs), 2))
        ws.cell(row, 8, round(sum(x.ret_iibb for x in liqs), 2))
        ws.cell(row, 9, round(sum(x.neto for x in liqs), 2))
        ws.cell(row, 10, f"=+B{row}-SUM(C{row}:H{row})-I{row}")
        for c in range(2, 11):
            _fmt(ws.cell(row, c))
            ws.cell(row, c).font = _FONT_BOLD

    ws.column_dimensions["A"].width = 22
    for i in range(2, 11):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.freeze_panes = "A2"


def _escribir_first_data(ws: Worksheet, liqs: list[LiqFD], comercio: str, periodo: str) -> None:
    ws.merge_cells("A2:K2")
    ws["A2"] = comercio or "FIRST DATA"
    ws["A2"].font = _FONT_TTL
    ws["A2"].fill = _FILL_TTL
    ws.merge_cells("A3:L3")
    ws["A3"] = f"FIRST DATA {periodo}".strip()
    ws["A3"].font = _FONT_HDR
    ws["A3"].fill = _FILL_HDR
    hdr = [
        "TOTAL",
        "BI 21",
        "IVA 21",
        "BI 10,5%",
        "IVA 10,5%",
        "P IIBB",
        "P IBB CABA",
        "R IIBB",
        "P IVA",
        "NETO",
        "Neto PDF",
        "Dif",
        "Nro Liq",
        "Fecha",
        "Archivo",
    ]
    for i, h in enumerate(hdr, 1):
        cell = ws.cell(4, i, h)
        cell.font = _FONT_HDR
        cell.fill = _FILL_HDR
        cell.alignment = Alignment(wrap_text=True, horizontal="center")
        cell.border = _THIN
    start = 5
    for i, liq in enumerate(liqs):
        r = start + i
        ws.cell(r, 1, liq.ventas)
        ws.cell(r, 2, _bi_formula(liq.arancel, liq.arancel_cuotas, liq.dto_cont, liq.dto_cuota, liq.serv_int))
        ws.cell(r, 3, f"=+B{r}*21%")
        ws.cell(r, 5, f"=+D{r}*10.5%")
        if liq.per_iibb_ba:
            ws.cell(r, 6, liq.per_iibb_ba)
        if liq.perc_caba:
            ws.cell(r, 7, liq.perc_caba)
        if liq.ret_iibb:
            ws.cell(r, 8, liq.ret_iibb)
        bits = [x for x in (liq.perc_iva_15, liq.perc_iva_30, liq.perc_iva_qr) if abs(x) > 0.001]
        if len(bits) == 1:
            ws.cell(r, 9, bits[0])
        elif bits:
            ws.cell(r, 9, "=" + "+".join(str(x) for x in bits))
        ws.cell(r, 10, f"=+A{r}-SUM(B{r}:I{r})")
        ws.cell(r, 11, liq.neto)
        ws.cell(r, 12, f"=+K{r}-J{r}")
        ws.cell(r, 13, liq.nro)
        ws.cell(r, 14, liq.fecha)
        ws.cell(r, 15, liq.archivo)
        for c in range(1, 13):
            _fmt(ws.cell(r, c))
            if i % 2:
                ws.cell(r, c).fill = _FILL_ZEBRA
        for c in range(13, 16):
            ws.cell(r, c).font = _FONT_BODY
            ws.cell(r, c).border = _THIN
    last = start + len(liqs) - 1
    tot = last + 1
    for col, letter in enumerate("ABCDEFGHIJKL", 1):
        if letter == "L":
            ws[f"L{tot}"] = f"=+K{tot}-J{tot}"
        else:
            ws[f"{letter}{tot}"] = f"=SUM({letter}{start}:{letter}{last})"
        _fmt(ws[f"{letter}{tot}"])
        ws[f"{letter}{tot}"].font = _FONT_BOLD
        ws[f"{letter}{tot}"].fill = _FILL_TOT
    ws.auto_filter.ref = f"A4:O{last}"
    ws.freeze_panes = "A5"
    ws.row_dimensions[4].height = 30
    for i, w in enumerate([12, 14, 12, 12, 12, 12, 12, 12, 12, 12, 12, 10, 12, 12, 36], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _escribir_movimientos(ws: Worksheet, liqs: list[LiqFD]) -> list[tuple]:
    for c, h in enumerate(HEADERS_MOV, 1):
        cell = ws.cell(1, c, h)
        cell.font = _FONT_HDR
        cell.fill = _FILL_HDR
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        cell.border = _THIN
    row = 2
    control: list[tuple] = []
    for liq in liqs:
        for cup in liq.cupones:
            ws.cell(row, 1, "VENTA")
            ws.cell(row, 2, cup.ventas)
            ws.cell(row, 4, -cup.arancel if cup.arancel else 0)
            ws.cell(row, 21, liq.nro)
            ws.cell(row, 22, liq.fecha)
            ws.cell(row, 23, liq.archivo)
            row += 1
        ws.cell(row, 1, "LIQUIDACIÓN")
        vals = {
            2: liq.ventas_contado,
            3: liq.ventas_cuotas,
            4: -liq.arancel,
            5: -liq.arancel_cuotas,
            6: -liq.dto_cont,
            7: -liq.dto_cuota,
            8: -liq.iva_arancel_cuotas,
            9: -liq.iva_dto_cuota,
            10: -liq.iva_dto_cont,
            11: -liq.iva_arancel,
            12: -liq.perc_iva_15,
            13: -liq.perc_iva_30,
            14: -liq.sirtac,
            15: -liq.per_iibb_ba,
            16: -liq.perc_caba,
            17: -liq.perc_iva_qr,
            18: -liq.serv_int,
            19: -liq.iva_serv_int,
            20: liq.neto,
            21: liq.nro,
            22: liq.fecha,
            23: liq.archivo,
        }
        for col, val in vals.items():
            ws.cell(row, col, val)
        ok = abs(liq.neto_calc - liq.neto) <= TOL
        control.append(
            (liq.nro, liq.ventas, -liq.deducciones, liq.neto_calc, liq.neto, "OK" if ok else "DIFF", liq.archivo)
        )
        row += 1
    for r in ws.iter_rows(min_row=2, max_row=max(2, row - 1), min_col=2, max_col=20):
        for cell in r:
            if isinstance(cell.value, (int, float)):
                _fmt(cell)
    ws.auto_filter.ref = f"A1:W{max(1, row - 1)}"
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32
    for i in range(1, 24):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions["W"].width = 36
    return control


def _escribir_control(ws: Worksheet, control: list[tuple]) -> None:
    hdr = [
        "Nro Liquidación",
        "Ventas",
        "Total Deducciones",
        "Neto Calculado",
        "Neto según PDF",
        "Control",
        "Archivo",
    ]
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(1, c, h)
        cell.font = _FONT_HDR
        cell.fill = _FILL_HDR
        cell.border = _THIN
    for i, fila in enumerate(control, 2):
        for c, val in enumerate(fila, 1):
            ws.cell(i, c, val)
            if c in (2, 3, 4, 5):
                _fmt(ws.cell(i, c))
            else:
                ws.cell(i, c).font = _FONT_BODY
                ws.cell(i, c).border = _THIN
    ws.auto_filter.ref = f"A1:G{len(control) + 1}"
    ws.freeze_panes = "A2"
    for i, w in enumerate([14, 14, 16, 16, 16, 12, 36], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _escribir_otros(ws: Worksheet, otros: list[dict]) -> None:
    cols = [
        "Archivo",
        "Entidad",
        "Fecha",
        "Nro_Liquidacion",
        "Neto_Gravado",
        "IVA_21",
        "Percepcion_IVA",
        "Retencion_IVA",
        "Retencion_IIBB",
        "Percepcion_IIBB",
        "Total_Descontado",
    ]
    for c, h in enumerate(cols, 1):
        cell = ws.cell(1, c, h)
        cell.font = _FONT_HDR
        cell.fill = _FILL_HDR
        cell.border = _THIN
    for i, d in enumerate(otros, 2):
        for c, k in enumerate(cols, 1):
            ws.cell(i, c, d.get(k))
            if c >= 5:
                _fmt(ws.cell(i, c))
    ws.auto_filter.ref = f"A1:K{len(otros) + 1}"
    ws.freeze_panes = "A2"


def generar_excel_estudio(
    *,
    naranja: list[dict],
    fava: list[dict],
    cabal: list[dict],
    liqs: list[LiqFD],
    otros: list[dict],
    comercio: str,
    periodo: str,
) -> bytes:
    wb = Workbook()
    ws_liq = wb.active
    ws_liq.title = "LIQUIDACIONES"
    _escribir_liquidaciones(ws_liq, naranja, fava, cabal, liqs, comercio)
    if liqs:
        ws_fd = wb.create_sheet("First Data")
        _escribir_first_data(ws_fd, liqs, comercio, periodo)
        ws_mov = wb.create_sheet("Movimientos")
        control = _escribir_movimientos(ws_mov, liqs)
        ws_ctrl = wb.create_sheet("Control")
        _escribir_control(ws_ctrl, control)
    if otros:
        _escribir_otros(wb.create_sheet("Otros"), otros)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _comercio_periodo(texto: str, nombre: str) -> tuple[str, str]:
    comercio = ""
    m = re.search(r"GLOBAL RECIFE[^\n]*|RAZ[OÓ]N SOCIAL[:\s]+([^\n]+)", texto, re.I)
    if m:
        comercio = (m.group(0) if m.lastindex is None else m.group(1)).strip()[:80]
    if "GLOBAL RECIFE" in texto.upper():
        comercio = "GLOBAL RECIFE"
    per = ""
    m2 = re.search(r"PESOS\s+([A-ZÁÉÍÓÚ]+)\s+(\d{4})", texto, re.I)
    if m2:
        mes = m2.group(1).title()
        per = f"{mes} {m2.group(2)}"
    m3 = re.search(r"(20\d{2})(0[1-9]|1[0-2])", nombre)
    if m3 and not per:
        meses = "Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic".split()
        per = f"{meses[int(m3.group(2)) - 1]} {m3.group(1)}"
    return comercio, per


def procesar_pdfs_tarjetas_estudio(archivos: list[tuple[str, bytes]]) -> ResultadoTarjetasEstudio:
    if not archivos:
        return ResultadoTarjetasEstudio(ok=False, error="Subí al menos un PDF.")

    naranja: list[dict] = []
    fava: list[dict] = []
    cabal: list[dict] = []
    liqs: list[LiqFD] = []
    otros: list[dict] = []
    mensajes: list[str] = []
    advertencias: list[str] = []
    comercio = ""
    periodo = ""

    with tempfile.TemporaryDirectory(prefix="liq_tarjetas_") as tmp:
        tmp_path = Path(tmp)
        for i, (nombre, data) in enumerate(archivos):
            safe = Path(nombre or f"liq_{i + 1}.pdf").name
            dest = tmp_path / safe
            dest.write_bytes(data)
            try:
                texto = texto_pdf(dest)
            except Exception as exc:
                advertencias.append(f"{safe}: no se pudo leer ({exc})")
                continue
            kind = clasificar_pdf(safe, texto)
            com, per = _comercio_periodo(texto, safe)
            comercio = comercio or com
            periodo = periodo or per
            try:
                if kind == "NARANJA":
                    naranja.append(parsear_naranja(dest, texto))
                    mensajes.append(f"{safe}: Naranja")
                elif kind == "FAVA":
                    fava.append(parsear_fava(dest, texto))
                    mensajes.append(f"{safe}: Favacard")
                elif kind == "CABAL":
                    parsed = parsear_cabal(dest, texto)
                    cabal.extend(parsed)
                    mensajes.append(f"{safe}: Cabal ({len(parsed)} liquidación/es)")
                elif kind == "FIRST_DATA":
                    parsed = parsear_fd(dest, texto)
                    liqs.extend(parsed)
                    mensajes.append(f"{safe}: First Data ({len(parsed)} liquidación/es)")
                else:
                    entidad = detectar_entidad_por_texto(f"{texto}\n{safe}")
                    fila = extraer_con_plantilla(texto, entidad)
                    fila["Archivo"] = safe
                    otros.append(fila)
                    mensajes.append(f"{safe}: {entidad} (hoja Otros)")
            except Exception as exc:
                advertencias.append(f"{safe}: {exc}")

    if not (naranja or fava or cabal or liqs or otros):
        return ResultadoTarjetasEstudio(
            ok=False,
            error="No se pudo extraer ninguna liquidación de los PDF.",
            mensajes=mensajes,
            advertencias=advertencias,
        )

    n_ok = sum(1 for x in liqs if abs(x.neto_calc - x.neto) <= TOL)
    n_diff = len(liqs) - n_ok
    n_ventas = sum(len(x.cupones) for x in liqs)
    for x in liqs:
        if x.otras_det:
            advertencias.append(f"Liq {x.nro}: deducción no mapeada: {x.otras_det}")

    xlsx = generar_excel_estudio(
        naranja=naranja,
        fava=fava,
        cabal=cabal,
        liqs=liqs,
        otros=otros,
        comercio=comercio,
        periodo=periodo,
    )
    hojas = ["LIQUIDACIONES"]
    if liqs:
        hojas.extend(["First Data", "Movimientos", "Control"])
    if otros:
        hojas.append("Otros")
    slug = re.sub(r"[^A-Za-z0-9]+", "_", (comercio or "Tarjetas"))[:40].strip("_")
    per_slug = re.sub(r"\s+", "_", periodo or "")
    nombre = f"Liquidaciones_de_Tarjetas_{slug}_{per_slug}.xlsx".replace("__", "_")
    return ResultadoTarjetasEstudio(
        ok=True,
        excel_bytes=xlsx,
        nombre_archivo=nombre,
        mensajes=mensajes,
        advertencias=advertencias,
        n_pdf=len(archivos),
        n_fd=len(liqs),
        n_ventas=n_ventas,
        n_control_ok=n_ok,
        n_control_diff=n_diff,
        comercio=comercio,
        periodo=periodo,
        hojas=hojas,
    )
