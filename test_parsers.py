# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import generar_auditoria as ga

# Test _limpiar_monto edge cases
print("=== _limpiar_monto tests ===")
casos = [
    ("993.995,55", 993995.55),
    ("3.006.617.31", 3006617.31),
    ("65.439.196.00", 65439196.0),
    ("2,756,034.83", 2756034.83),
    ("0,00", 0.0),
    ("24.800.000,00", 24800000.0),
    ("388141,79", 388141.79),
    ("0.0o", 0.0),
    ("5.006.617,31", 5006617.31),
]
for s, esperado in casos:
    got = ga._limpiar_monto(s)
    ok = abs(got - esperado) < 0.02
    status = "OK" if ok else f"FAIL got={got} esperado={esperado}"
    print(f"  {s!r} -> {got} ({status})")

# Test parsear_fecha_flexible
print()
print("=== _parsear_fecha_flexible tests ===")
for txt in ["18/06'2025", "18/08*2025", "8,09/2025", "18/10,2025", "18/03.2026", "18/04/2025"]:
    f = ga._parsear_fecha_flexible(txt)
    print(f"  {txt!r} -> {f}")

# Test Santander OCR parser with simulated text
print()
print("=== Santander OCR parser test ===")
santander_ocr_sim = ["""Desarrollo del Prestamo
(peracion: 0032-0212-039100426399 18/03/2026
Escha _YIe Capital Cuota Mnt Connpens Per Ivlal Cupta Saldo Deuda Capital
18/04/2025 0,00 5.006.617,31 3.006.617.31 65.439.196.00
18/05/2025 0,00
2.097.639,98 2.097.639,98 65.439.196,00
18/06'2025 0,00 2.167.561,31 2.167.561,31 65.439.196.00
18/03.2026 65.439.196,00 1.957.797,32 67.396.993,32 0.00"""]
c = ga._parsear_santander_ocr(santander_ocr_sim)
print(f"Cuotas extraidas: {len(c)}")
for cuota in c:
    print(f"  {cuota}")

# Test Nacion OCR parser
print()
print("=== Nacion OCR parser test ===")
nacion_ocr_sim = ["""BANCO NACION ARGENTINA
Fecha Ir Vto  08-09-2025
Ult.Vto 08-09-2025
Capital 24 .800 .000 , 00
24 .800 .000 00 Importe Neto sin Cargos ni IVA"""]
c2 = ga._parsear_nacion_ocr(nacion_ocr_sim)
print(f"Cuotas extraidas: {len(c2)}")
for cuota in c2:
    print(f"  {cuota}")
