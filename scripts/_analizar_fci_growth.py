"""Growth FCI: cuadros FIFO (intereses, lotes, tenencia) desde FCI GROWTH.xlsx."""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_formato_estudio import (
    BODY_FONT,
    BOLD_FONT,
    DATE_FMT,
    HDR_FONT,
    MONEY_FMT,
    SECTION_FONT,
    SUB_FONT,
    ZEBRA,
    _escribir_encabezado_hoja,
    _escribir_tabla,
    _pintar_header_fila,
)
from inversiones import aplicar_fifo, identificar_movimientos

HOJA = Path(r"C:\Users\recep\Desktop\FCI GROWTH.xlsx")
OUT = Path(r"C:\Users\recep\Desktop\Intereses_FCI_Growth.xlsx")
FONDO = "FIMA PREMIUM CLASE A"
FECHA_INI = date(2025, 6, 30)
FECHA_CIERRE = date(2026, 6, 30)
# Valor cuota del último extracto de fondos del mes (posición 30/06/2026).
VC_CIERRE_EXTRACTO = 81.856194

CUOTAS_FMT = "#,##0.00"
VC_FMT = "0.000000"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="D6E3F0")
NOTA_FONT = Font(name="Calibri", size=10, italic=True, color="666666")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def leer_hoja() -> tuple[float, float, float, list[dict], date | None, float | None]:
    wb = load_workbook(HOJA, data_only=False)
    ws = wb["FCI"]
    ini_cant = float(ws["C7"].value)
    ini_vc = float(ws["D7"].value)
    ini_tot = float(ws["E7"].value)
    movs: list[dict] = []
    ultima_fecha: date | None = None
    ultimo_vc: float | None = None
    for r in range(10, (ws.max_row or 10) + 1):
        tipo = str(ws.cell(r, 3).value or "").strip().upper()
        if tipo not in {"RESCATE", "SUSCRIPCION"}:
            continue
        fecha = _as_date(ws.cell(r, 2).value)
        if fecha is None:
            continue
        cant = float(ws.cell(r, 4).value)
        vc = float(ws.cell(r, 5).value)
        tot = ws.cell(r, 6).value
        tot_f = float(tot) if tot not in (None, "") else round(cant * vc, 2)
        movs.append(
            {
                "Fecha": fecha.strftime("%d/%m/%Y"),
                "_fecha": fecha,
                "Especie": FONDO,
                "Tipo_Operacion": "Rescate" if tipo == "RESCATE" else "Suscripcion",
                "Cantidad": cant,
                "Precio": vc,
                "Monto_Total": tot_f,
                "Moneda": "ARS",
                "Descripcion": tipo,
                "Archivo origen": HOJA.name,
                "Nueva_Clasificacion": None,
            }
        )
        ultima_fecha = fecha
        ultimo_vc = vc
    return ini_cant, ini_vc, ini_tot, movs, ultima_fecha, ultimo_vc


def consolidar_lotes(lotes: list[dict], fecha_inicial: date) -> list[dict]:
    agrupado: dict[tuple, dict] = {}
    for lote in lotes:
        fecha = lote.get("Fecha") or fecha_inicial
        origen = str(lote.get("Origen") or "")
        if "inicial" in origen.lower() or lote.get("Fecha") is None:
            origen_txt = "Saldo inicial"
            fecha = fecha_inicial
        else:
            origen_txt = "Suscripción"
        clave = (fecha, round(float(lote["Costo_Unitario"]), 8), origen_txt)
        if clave not in agrupado:
            agrupado[clave] = {
                "Origen": origen_txt,
                "Fecha origen": fecha,
                "Cuotas": 0.0,
                "Valor cuota origen": float(lote["Costo_Unitario"]),
                "Costo de origen": 0.0,
            }
        agrupado[clave]["Cuotas"] += float(lote["Cantidad"])
        agrupado[clave]["Costo de origen"] += float(lote["Costo_Total"])
    filas = list(agrupado.values())
    for fila in filas:
        fila["Cuotas"] = round(float(fila["Cuotas"]), 4)
        fila["Costo de origen"] = round(float(fila["Costo de origen"]), 2)
    filas.sort(key=lambda r: r["Fecha origen"])
    return filas


def _borde_rango(ws: Worksheet, r1: int, c1: int, r2: int, c2: int) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = THIN


def _pintar_total(ws: Worksheet, fila: int, n_cols: int) -> None:
    for c in range(1, n_cols + 1):
        cell = ws.cell(fila, c)
        cell.fill = TOTAL_FILL
        cell.font = BOLD_FONT


def _fmt_celda(cell, tipo: str) -> None:
    cell.font = BODY_FONT
    if tipo == "money":
        cell.number_format = MONEY_FMT
        cell.alignment = Alignment(horizontal="right")
    elif tipo == "qty":
        cell.number_format = CUOTAS_FMT
        cell.alignment = Alignment(horizontal="right")
    elif tipo == "vc":
        cell.number_format = VC_FMT
        cell.alignment = Alignment(horizontal="right")
    elif tipo == "date":
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(horizontal="center")
    elif tipo == "int":
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right")


def escribir_cuadro(
    ws: Worksheet,
    fila: int,
    titulo: str,
    headers: list[str],
    rows: list[list],
    tipos: list[str],
    *,
    total_cols: list[int] | None = None,
    nota: str = "",
) -> tuple[int, int | None]:
    """Escribe un cuadro. Devuelve (próxima fila, fila TOTAL o None)."""
    n = len(headers)
    ws.cell(fila, 1, titulo).font = SECTION_FONT
    fila += 1
    if nota:
        ws.cell(fila, 1, nota).font = NOTA_FONT
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=n)
        fila += 1

    header_row = fila
    for c, h in enumerate(headers, 1):
        ws.cell(fila, c, h)
    _pintar_header_fila(ws, fila, n)
    fila += 1

    first_data = fila
    for i, row in enumerate(rows):
        for c, (val, tipo) in enumerate(zip(row, tipos, strict=True), 1):
            cell = ws.cell(fila, c, val)
            _fmt_celda(cell, tipo)
            if i % 2 == 1:
                cell.fill = ZEBRA
        fila += 1
    last_data = fila - 1 if rows else header_row

    total_row = None
    if total_cols and rows:
        total_row = fila
        ws.cell(fila, 1, "TOTAL").font = BOLD_FONT
        for c in total_cols:
            letra = get_column_letter(c)
            cell = ws.cell(fila, c, f"=SUM({letra}{first_data}:{letra}{last_data})")
            cell.font = BOLD_FONT
            _fmt_celda(cell, tipos[c - 1])
            cell.font = BOLD_FONT
        _pintar_total(ws, fila, n)
        fila += 1

    _borde_rango(ws, header_row, 1, (total_row or last_data), n)
    return fila + 1, total_row


def armar_excel(
    *,
    mensual: pd.DataFrame,
    lotes: list[dict],
    detalle: pd.DataFrame,
    ini_cant: float,
    ini_vc: float,
    ini_tot: float,
    interes_total: float,
    f_fin: date | None,
    vc_fin: float | None,
) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"

    fila = _escribir_encabezado_hoja(
        ws,
        "Growth Marketing S.A. — FCI FIMA Premium Clase A",
        (
            "FIFO: interés de cada rescate = (valor cuota del rescate − valor cuota del lote) × cuotas. "
            "Saldo inicial 30/06/2025 tomado de FCI GROWTH.xlsx."
        ),
        "Ejercicio 01/07/2025 al 30/06/2026",
    )

    cuotas_fin = round(sum(l["Cuotas"] for l in lotes), 4)
    costo_fin = round(sum(l["Costo de origen"] for l in lotes), 2)

    ws.cell(fila, 1, "Saldo inicial (30/06/2025)").font = BOLD_FONT
    ws.cell(fila, 2, ini_cant)
    _fmt_celda(ws.cell(fila, 2), "qty")
    ws.cell(fila, 3, "cuotas  @")
    ws.cell(fila, 3).font = BODY_FONT
    ws.cell(fila, 4, ini_vc)
    _fmt_celda(ws.cell(fila, 4), "vc")
    ws.cell(fila, 5, ini_tot)
    _fmt_celda(ws.cell(fila, 5), "money")
    fila += 1

    ws.cell(fila, 1, "Último movimiento del extracto").font = BOLD_FONT
    ws.cell(fila, 2, f_fin)
    _fmt_celda(ws.cell(fila, 2), "date")
    ws.cell(fila, 3, "VC")
    ws.cell(fila, 3).font = BODY_FONT
    ws.cell(fila, 4, vc_fin)
    _fmt_celda(ws.cell(fila, 4), "vc")
    fila += 2

    ws.cell(fila, 1, "Interés realizado (rescates)").font = BOLD_FONT
    cell_int = ws.cell(fila, 2, interes_total)
    _fmt_celda(cell_int, "money")
    cell_int.font = BOLD_FONT
    fila += 1

    ws.cell(fila, 1, "Cuotas saldo final").font = BOLD_FONT
    cell_c = ws.cell(fila, 2, cuotas_fin)
    _fmt_celda(cell_c, "qty")
    cell_c.font = BOLD_FONT
    fila += 1

    ws.cell(fila, 1, "Costo de origen del saldo").font = BOLD_FONT
    cell_k = ws.cell(fila, 2, costo_fin)
    _fmt_celda(cell_k, "money")
    cell_k.font = BOLD_FONT
    fila += 1

    kpi_ten_row = fila
    ws.cell(fila, 1, "Resultado por tenencia al 30/06/2026").font = BOLD_FONT
    fila += 2

    mensual_rows = []
    for _, r in mensual.iterrows():
        mensual_rows.append(
            [
                r["Mes"],
                int(r["Suscripciones"]),
                int(r["Rescates"]),
                float(r["Cuotas suscriptas"]),
                float(r["Cuotas rescatadas"]),
                float(r["Interés ganado"]),
            ]
        )
    fila, _tot_int = escribir_cuadro(
        ws,
        fila,
        "1. Intereses realizados por mes (rescates FIFO)",
        [
            "Mes",
            "Suscripciones",
            "Rescates",
            "Cuotas suscriptas",
            "Cuotas rescatadas",
            "Interés ganado",
        ],
        mensual_rows,
        ["text", "int", "int", "qty", "qty", "money"],
        total_cols=[4, 5, 6],
        nota="El interés de cada rescate se calcula lote por lote: (VC rescate − VC de origen) × cuotas aplicadas.",
    )

    lotes_rows = [
        [
            l["Origen"],
            l["Fecha origen"],
            l["Cuotas"],
            l["Valor cuota origen"],
            l["Costo de origen"],
        ]
        for l in lotes
    ]
    fila, _tot_lotes = escribir_cuadro(
        ws,
        fila,
        "2. Composición del saldo final (lotes FIFO que quedan)",
        [
            "Origen",
            "Fecha origen",
            "Cuotas",
            "Valor cuota origen",
            "Costo de origen",
        ],
        lotes_rows,
        ["text", "date", "qty", "vc", "money"],
        total_cols=[3, 5],
        nota=(
            "Son las suscripciones (y el saldo inicial, si quedara) que FIFO no consumió. "
            "El costo de origen es cuotas × valor cuota de esa suscripción."
        ),
    )

    ws.cell(fila, 1, "3. Resultado por tenencia al 30/06/2026").font = SECTION_FONT
    fila += 1
    ws.cell(
        fila,
        1,
        "Valuación al cierre = cuotas × valor cuota al 30/06/2026.  "
        "Resultado por tenencia = valuación al 30/06/2026 − costo de origen de las suscripciones del saldo.",
    ).font = NOTA_FONT
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=6)
    fila += 1

    ws.cell(fila, 1, "Valor cuota al 30/06/2026").font = BOLD_FONT
    vc_row = fila
    vc_cell = ws.cell(fila, 2)
    vc_cell.fill = INPUT_FILL
    vc_cell.number_format = VC_FMT
    vc_cell.font = BOLD_FONT
    vc_cell.comment = Comment(
        "Completar con la cotización oficial de FIMA Premium Clase A al 30/06/2026 "
        "(CAFCI / Galicia). No figura en FCI GROWTH.xlsx.",
        "Estudio",
    )
    ws.cell(
        fila,
        3,
        f"Celda amarilla: cotización oficial de cierre. Último VC del extracto "
        f"({f_fin.strftime('%d/%m/%Y') if f_fin else 's/f'}): {vc_fin}",
    ).font = NOTA_FONT
    fila += 2

    ten_headers = [
        "Origen",
        "Fecha origen",
        "Cuotas",
        "Valor cuota origen",
        "Costo de origen",
        "Valuación 30/06/2026",
        "Resultado por tenencia",
    ]
    header_row = fila
    for c, h in enumerate(ten_headers, 1):
        ws.cell(fila, c, h)
    _pintar_header_fila(ws, fila, len(ten_headers))
    fila += 1

    first_ten = fila
    vc_ref = f"$B${vc_row}"
    for i, lote in enumerate(lotes):
        ws.cell(fila, 1, lote["Origen"]).font = BODY_FONT
        c_fecha = ws.cell(fila, 2, lote["Fecha origen"])
        _fmt_celda(c_fecha, "date")
        c_cuotas = ws.cell(fila, 3, lote["Cuotas"])
        _fmt_celda(c_cuotas, "qty")
        c_vc = ws.cell(fila, 4, lote["Valor cuota origen"])
        _fmt_celda(c_vc, "vc")
        c_costo = ws.cell(fila, 5, lote["Costo de origen"])
        _fmt_celda(c_costo, "money")
        c_val = ws.cell(fila, 6, f'=IF({vc_ref}="","",C{fila}*{vc_ref})')
        _fmt_celda(c_val, "money")
        c_res = ws.cell(fila, 7, f'=IF({vc_ref}="","",F{fila}-E{fila})')
        _fmt_celda(c_res, "money")
        if i % 2 == 1:
            for c in range(1, 8):
                ws.cell(fila, c).fill = ZEBRA
        fila += 1
    last_ten = fila - 1

    tot_ten = fila
    ws.cell(fila, 1, "TOTAL").font = BOLD_FONT
    for col, tipo in ((3, "qty"), (5, "money"), (6, "money"), (7, "money")):
        letra = get_column_letter(col)
        cell = ws.cell(fila, col, f"=SUM({letra}{first_ten}:{letra}{last_ten})")
        _fmt_celda(cell, tipo)
        cell.font = BOLD_FONT
    _pintar_total(ws, fila, 7)
    _borde_rango(ws, header_row, 1, tot_ten, 7)
    fila += 2

    kpi_ten = ws.cell(kpi_ten_row, 2, f'=IF({vc_ref}="","(completar VC 30/06)",G{tot_ten})')
    kpi_ten.font = BOLD_FONT
    kpi_ten.number_format = MONEY_FMT

    ws.cell(
        fila,
        1,
        "Si se carga el valor cuota del 30/06/2026 en la celda amarilla, valuación y resultado por tenencia se recalculan solos.",
    ).font = NOTA_FONT

    anchos = {1: 42, 2: 18, 3: 20, 4: 22, 5: 22, 6: 24, 7: 26}
    for c, w in anchos.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    wb.defined_names.add(DefinedName(name="VC_CIERRE", attr_text=f"Resumen!$B${vc_row}"))
    wb.defined_names.add(DefinedName(name="TENENCIA_TOTAL", attr_text=f"Resumen!$G${tot_ten}"))

    ws_det = wb.create_sheet("Tramos FIFO", 1)
    _escribir_tabla(
        ws_det,
        detalle,
        1,
        col_moneda=["Costo lote", "Interés"],
        col_fecha=["Fecha", "Origen lote"],
        col_texto=["Tipo"],
        zebra=True,
        total_col="Interés",
    )
    ws_det.column_dimensions["A"].width = 14
    ws_det.column_dimensions["B"].width = 14
    ws_det.column_dimensions["C"].width = 14
    ws_det.column_dimensions["D"].width = 16
    ws_det.column_dimensions["E"].width = 16
    ws_det.column_dimensions["F"].width = 14
    ws_det.column_dimensions["G"].width = 16
    ws_det.column_dimensions["H"].width = 16
    return wb


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    ini_cant, ini_vc, ini_tot, movs, f_fin, vc_fin = leer_hoja()
    ini = pd.DataFrame(
        [
            {
                "Especie": FONDO,
                "Cantidad": ini_cant,
                "Costo_Unitario": ini_vc,
                "Moneda": "ARS",
            }
        ]
    )
    df = identificar_movimientos(pd.DataFrame(movs))
    res = aplicar_fifo(df, ini)
    avisos = [a for a in res.avisos if "sin stock" in a.lower()]
    if avisos:
        print("AVISOS STOCK")
        for a in avisos:
            print(" ", a)

    apps = pd.DataFrame(res.aplicaciones)
    apps["Fecha"] = pd.to_datetime(apps["Fecha_salida"], dayfirst=True, errors="coerce")
    apps["Interes"] = pd.to_numeric(apps["Interes"], errors="coerce").fillna(0)
    apps["Cantidad"] = pd.to_numeric(apps["Cantidad"], errors="coerce").fillna(0)

    df["_fecha"] = pd.to_datetime(df["_fecha"])
    mensual_rows = []
    for (y, m), g in apps.groupby([apps["Fecha"].dt.year, apps["Fecha"].dt.month]):
        sus = df[
            (df["Tipo_Operacion"] == "Suscripcion")
            & (df["_fecha"].dt.year == y)
            & (df["_fecha"].dt.month == m)
        ]
        rec = df[
            (df["Tipo_Operacion"] == "Rescate")
            & (df["_fecha"].dt.year == y)
            & (df["_fecha"].dt.month == m)
        ]
        mensual_rows.append(
            {
                "Mes": f"{MESES[int(m)]} {int(y)}",
                "Orden": f"{int(y):04d}-{int(m):02d}",
                "Suscripciones": int(len(sus)),
                "Rescates": int(len(rec)),
                "Cuotas suscriptas": round(float(sus["Cantidad"].sum()), 2) if len(sus) else 0.0,
                "Cuotas rescatadas": round(float(g["Cantidad"].sum()), 2),
                "Interés ganado": round(float(g["Interes"].sum()), 2),
            }
        )
    mensual = pd.DataFrame(mensual_rows).sort_values("Orden")
    interes_total = round(float(mensual["Interés ganado"].sum()), 2)

    lotes = consolidar_lotes(res.lotes_abiertos, FECHA_INI)

    detalle = apps[
        [
            "Fecha_salida",
            "Evento",
            "Cantidad",
            "Valor_CC_lote",
            "Valor_CC_salida",
            "Fecha_lote",
            "Costo_parcial",
            "Interes",
        ]
    ].copy()
    detalle = detalle.rename(
        columns={
            "Fecha_salida": "Fecha",
            "Evento": "Tipo",
            "Cantidad": "Cuotas",
            "Valor_CC_lote": "VC lote",
            "Valor_CC_salida": "VC rescate",
            "Fecha_lote": "Origen lote",
            "Costo_parcial": "Costo lote",
            "Interes": "Interés",
        }
    )

    wb = armar_excel(
        mensual=mensual,
        lotes=lotes,
        detalle=detalle,
        ini_cant=ini_cant,
        ini_vc=ini_vc,
        ini_tot=ini_tot,
        interes_total=interes_total,
        f_fin=f_fin,
        vc_fin=vc_fin,
    )
    destino = OUT
    try:
        wb.save(destino)
    except PermissionError:
        destino = OUT.with_name("Intereses_FCI_Growth_cuadro.xlsx")
        wb.save(destino)

    print("Saldo inicial", ini_cant, "cuotas @", ini_vc, "=", ini_tot)
    print("Movimientos", len(movs), "rescates", sum(1 for m in movs if m["Tipo_Operacion"] == "Rescate"))
    print()
    print(f"{'Mes':<18} {'Interés':>16}")
    for _, r in mensual.iterrows():
        print(f"{r['Mes']:<18} {r['Interés ganado']:>16,.2f}")
    print(f"{'TOTAL realizado':<18} {interes_total:>16,.2f}")
    print()
    print("Lotes saldo final")
    for l in lotes:
        print(
            f"  {l['Origen']:<16} {l['Fecha origen'].strftime('%d/%m/%Y')}  "
            f"{l['Cuotas']:>12,.4f}  VC {l['Valor cuota origen']:.6f}  "
            f"$ {l['Costo de origen']:,.2f}"
        )
    print(
        f"  {'TOTAL':<16} {'':<10}  {sum(l['Cuotas'] for l in lotes):>12,.4f}  "
        f"{'':<14}  $ {sum(l['Costo de origen'] for l in lotes):,.2f}"
    )
    print()
    print("Cierre contable", FECHA_CIERRE.strftime("%d/%m/%Y"), "— VC a completar en celda amarilla")
    print("OK", destino)


if __name__ == "__main__":
    main()
