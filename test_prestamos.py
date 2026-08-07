import sys
import os

sys.path.insert(0, r"C:\Users\recep\Desktop\Estudio Contable")
from procesador import generar_excel_prestamos_auditoria

bancos_data = {
    "Banco Galicia": [
        {
            "prestamo_n": 1,
            "capital_original": 5000000,
            "sistema": "Francés",
            "cuotas": [
                {
                    "cuota": i + 1,
                    "vencimiento": f"2026-{i + 1:02d}-15",
                    "capital": 50000 * (i + 1),
                    "intereses": 30000,
                    "iva_gastos": 6300,
                    "monto_abonar": 86300,
                    "saldo_restante": 5000000 - 50000 * (i + 1),
                }
                for i in range(12)
            ],
        }
    ],
    "Banco Santander": [
        {
            "prestamo_n": 1,
            "capital_original": 2000000,
            "sistema": "Alemán",
            "cuotas": [
                {
                    "cuota": i + 1,
                    "vencimiento": f"2026-{i + 1:02d}-20",
                    "capital": 20000,
                    "intereses": 15000,
                    "iva_gastos": 3150,
                    "monto_abonar": 38150,
                    "saldo_restante": 2000000 - 20000 * (i + 1),
                }
                for i in range(12)
            ],
        }
    ],
}

saldos = {"Banco Galicia": 5000000, "Banco Santander": 2000000}

ruta = generar_excel_prestamos_auditoria(bancos_data, saldos, modo="completo")
print(f"Excel generado: {ruta}")

tamano = os.path.getsize(ruta)
print(f"Tamaño archivo: {tamano:,} bytes")

import openpyxl
wb = openpyxl.load_workbook(ruta)
print(f"Hojas creadas: {wb.sheetnames}")
print("TEST EXITOSO")
