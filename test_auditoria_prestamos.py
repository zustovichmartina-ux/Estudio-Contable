"""Prueba mínima de auditoría de préstamos con datos simulados (sin PDF real)."""
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd

from procesador import (
    CuotaPrestamo,
    MovimientoBanco,
    cruzar_extractos_prestamos,
    constituir_saldos_finales,
    ejecutar_auditoria_prestamos,
    generar_reporte_auditoria_prestamos,
    validar_contra_mayor,
)

SALIDA = Path("exportaciones/Auditoria_Prestamos_test_mock.xlsx")


def _crear_mayor_mock() -> BytesIO:
    """Genera mayor Tango simulado en memoria."""
    filas = [
        {"fecha": "05/03/2025", "codigo_cuenta": "22101", "descripcion": "Cuota capital", "debe": 0, "haber": 80000, "saldo": 920000},
        {"fecha": "05/03/2025", "codigo_cuenta": "42502", "descripcion": "Intereses prestamo", "debe": 20000, "haber": 0, "saldo": 20000},
        {"fecha": "05/04/2025", "codigo_cuenta": "22101", "descripcion": "Cuota capital", "debe": 0, "haber": 81000, "saldo": 839000},
        {"fecha": "05/04/2025", "codigo_cuenta": "42502", "descripcion": "Intereses prestamo", "debe": 19000, "haber": 0, "saldo": 39000},
    ]
    df = pd.DataFrame(filas)
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    buf.name = "mayor_mock.xlsx"
    return buf


def main():
    cuotas = [
        CuotaPrestamo(1, date(2025, 3, 5), 80000.0, 20000.0, 100000.0, banco="santander"),
        CuotaPrestamo(2, date(2025, 4, 5), 81000.0, 19000.0, 100000.0, banco="santander"),
    ]
    movs = [
        MovimientoBanco(date(2025, 3, 4), "001", "Debito cuota prestamo", 100000.0, 0, banco="santander"),
        MovimientoBanco(date(2025, 4, 6), "002", "Debito cuota prestamo", 100000.0, 0, banco="santander"),
    ]

    cruces = cruzar_extractos_prestamos(movs, cuotas)
    assert len(cruces) == 2, "Deben generarse 2 cruces"
    assert all(c.coincidencia for c in cruces), "Ambos cruces deben coincidir"
    print("OK cruzar_extractos_prestamos:", sum(c.coincidencia for c in cruces), "cruces")

    mayor_buf = _crear_mayor_mock()
    from procesador import cargar_mayor_tango
    mayor_df = cargar_mayor_tango(mayor_buf)
    validaciones = validar_contra_mayor(mayor_df, cuotas, movs, cruces)
    print("OK validar_contra_mayor:", len(validaciones), "validaciones")

    saldos = constituir_saldos_finales(
        [{"banco": "Banco Santander", "saldo_inicial": 1000000.0}],
        cuotas,
        movs,
        mayor_df,
    )
    assert len(saldos) == 1
    print("OK constituir_saldos_finales: calculado=", saldos[0].saldo_calculado)

    # Pipeline completo sin PDFs (cuotas vacías, solo extractos simulados vía movs directos)
    resultado = ejecutar_auditoria_prestamos(
        pdf_prestamos=[],
        pdf_extractos=[],
        ruta_mayor=mayor_buf,
        saldos_iniciales=[{"banco": "Banco Santander", "saldo_inicial": 1000000.0}],
        nombre_cliente="TEST MOCK PJ",
    )
    resultado.cuotas = cuotas
    resultado.cruces = cruces
    resultado.movimientos_banco = movs
    resultado.validaciones_mayor = validaciones
    resultado.saldos_por_banco = saldos

    excel = generar_reporte_auditoria_prestamos(resultado)
    SALIDA.parent.mkdir(exist_ok=True)
    SALIDA.write_bytes(excel)
    print("OK generar_reporte ->", SALIDA.resolve())
    print("TODAS LAS PRUEBAS PASARON")


if __name__ == "__main__":
    main()
