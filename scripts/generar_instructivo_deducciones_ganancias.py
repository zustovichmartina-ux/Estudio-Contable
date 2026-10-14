#!/usr/bin/env python3
"""Genera Instructivo_Deducciones_Ganancias_1hoja.pdf (checklist documentación).

Formato pedido:
- Sin topes ni montos ("sin explicación de cuánto").
- Si tenés esta deducción → necesitás esta documentación.
- Estética Estudio Zona Güemes (#1F4E79, wordmark, 1 hoja A4).
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

REPO = Path(__file__).resolve().parents[1]
OUT_DIR = REPO / "plantillas"
OUT_PDF = OUT_DIR / "Instructivo_Deducciones_Ganancias_1hoja.pdf"
WORDMARK = REPO / "assets" / "estudio-zona-guemes-wordmark-oscuro.png"

# Identidad visual (excel_formato_estudio + marca)
AZUL = colors.HexColor("#1F4E79")
AZUL_SUAVE = colors.HexColor("#2E75B6")
TEXTO = colors.HexColor("#1A1A1A")
SUAVE = colors.HexColor("#666666")
ZEBRA = colors.HexColor("#F2F2F2")
LINEA = colors.HexColor("#D0D7DE")
BLANCO = colors.white

# (Si tenés…, Documentación a presentar) — sin montos ni topes
FILAS: list[tuple[str, str]] = [
    (
        "Cónyuge o unión convivencial",
        "DNI/CUIL y partida de matrimonio o constancia de unión convivencial.",
    ),
    (
        "Hijos / hijastros a cargo",
        "DNI/CUIL y partida de nacimiento (o sentencia de adopción / tenencia).",
    ),
    (
        "Otras cargas de familia",
        "Documentación del vínculo y que acredite dependencia económica.",
    ),
    (
        "Alquiler de vivienda (inquilino)",
        "Contrato + facturas/recibos del año a tu nombre + CUIT del locador.",
    ),
    (
        "Intereses de crédito hipotecario",
        "Certificado bancario de intereses pagados en el año fiscal.",
    ),
    (
        "Seguros de vida / sepelio",
        "Póliza vigente + comprobantes de pago del período.",
    ),
    (
        "Cuota de medicina prepaga / obra social",
        "Facturas o resumen anual a nombre tuyo o de tus cargas.",
    ),
    (
        "Gastos médicos, odontológicos o paramédicos",
        "Facturas a tu nombre (o de cargas) del profesional/institución.",
    ),
    (
        "Donaciones",
        "Recibos de entidades reconocidas por AFIP (con CUIT y concepto).",
    ),
    (
        "Personal de casas particulares",
        "Alta AFIP + comprobantes de sueldos y aportes del año.",
    ),
    (
        "Aportes a seguros de retiro / voluntarios",
        "Comprobantes de aporte emitidos por la compañía o entidad.",
    ),
    (
        "Gastos de sepelio",
        "Facturas del servicio a tu nombre.",
    ),
    (
        "Indumentaria o equipamiento de trabajo",
        "Facturas + nota del empleador si corresponde (relación de dependencia).",
    ),
    (
        "Viajes / educación de hijos",
        "Facturas o recibos del establecimiento o prestador.",
    ),
]


def _styles() -> dict[str, ParagraphStyle]:
    return {
        "titulo": ParagraphStyle(
            "titulo",
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=17,
            textColor=AZUL,
            alignment=TA_CENTER,
            spaceAfter=2,
        ),
        "subtitulo": ParagraphStyle(
            "subtitulo",
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=SUAVE,
            alignment=TA_CENTER,
            spaceAfter=6,
        ),
        "nota": ParagraphStyle(
            "nota",
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=SUAVE,
            alignment=TA_LEFT,
        ),
        "th": ParagraphStyle(
            "th",
            fontName="Helvetica-Bold",
            fontSize=8.5,
            leading=10,
            textColor=BLANCO,
            alignment=TA_LEFT,
        ),
        "celda": ParagraphStyle(
            "celda",
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=TEXTO,
            alignment=TA_LEFT,
        ),
        "celda_bold": ParagraphStyle(
            "celda_bold",
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=TEXTO,
            alignment=TA_LEFT,
        ),
        "pie": ParagraphStyle(
            "pie",
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            textColor=SUAVE,
            alignment=TA_CENTER,
        ),
    }


def construir_documento(destino: Path = OUT_PDF) -> Path:
    destino.parent.mkdir(parents=True, exist_ok=True)
    estilos = _styles()

    doc = SimpleDocTemplate(
        str(destino),
        pagesize=A4,
        leftMargin=1.3 * cm,
        rightMargin=1.3 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
        title="Instructivo Deducciones Ganancias — Documentación",
        author="Estudio Zona Güemes",
    )

    story: list = []

    if WORDMARK.exists():
        logo = Image(str(WORDMARK), width=7.2 * cm, height=2.4 * cm)
        logo.hAlign = "CENTER"
        story.append(logo)
        story.append(Spacer(1, 2 * mm))

    story.append(Paragraph("INSTRUCTIVO — DEDUCCIONES GANANCIAS", estilos["titulo"]))
    story.append(
        Paragraph(
            "Checklist de documentación · 1 hoja · Sin montos ni topes",
            estilos["subtitulo"],
        )
    )
    story.append(
        Paragraph(
            "<b>Regla:</b> si tenés la deducción de la izquierda, "
            "necesitamos la documentación de la derecha. "
            "No hace falta explicar montos: con el comprobante alcanza.",
            estilos["nota"],
        )
    )
    story.append(Spacer(1, 3 * mm))

    header = [
        Paragraph("Si tenés…", estilos["th"]),
        Paragraph("Documentación a presentar", estilos["th"]),
    ]
    data = [header]
    for izquierda, derecha in FILAS:
        data.append(
            [
                Paragraph(izquierda, estilos["celda_bold"]),
                Paragraph(derecha, estilos["celda"]),
            ]
        )

    col_w = [7.2 * cm, 11.0 * cm]
    tabla = Table(data, colWidths=col_w, repeatRows=1)
    style_cmds: list = [
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("TEXTCOLOR", (0, 0), (-1, 0), BLANCO),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("BOX", (0, 0), (-1, -1), 0.6, AZUL),
        ("LINEBELOW", (0, 0), (-1, 0), 1.2, AZUL),
        ("LINEBELOW", (0, 1), (-1, -2), 0.3, LINEA),
        ("LINEAFTER", (0, 0), (0, -1), 0.4, LINEA),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    tabla.setStyle(TableStyle(style_cmds))
    story.append(tabla)

    story.append(Spacer(1, 4 * mm))
    story.append(
        Paragraph(
            "Enviá todo digitalizado (PDF o foto legible) al estudio. "
            "Si falta un comprobante, esa deducción no se puede cargar.",
            estilos["pie"],
        )
    )
    story.append(
        Paragraph(
            "Estudio Zona Güemes · Contador Hernán Trujillo",
            estilos["pie"],
        )
    )

    doc.build(story)
    return destino


if __name__ == "__main__":
    path = construir_documento()
    print(path)
