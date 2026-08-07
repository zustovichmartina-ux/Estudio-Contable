import sys
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from procesador import escanear_carpeta_cliente, extraer_movimientos_anuales, RUTA_RAIZ_CLIENTES

cuit = "30717847810"
print(f"Escaneando carpeta para CUIT: {cuit}")
rutas_pdf, ruta_compras, ruta_cuentas = escanear_carpeta_cliente(cuit, RUTA_RAIZ_CLIENTES)

print(f"PDFs encontrados: {len(rutas_pdf)}")
print(f"Ruta compras: {ruta_compras}")
print(f"Ruta cuentas: {ruta_cuentas}")

if rutas_pdf:
    print("Extrayendo movimientos...")
    mov_banco, bancos, saldos_mes = extraer_movimientos_anuales(rutas_pdf)
    print(f"Movimientos extraídos: {len(mov_banco)}")
    print(f"Bancos detectados: {bancos}")
    print(f"Saldos mes: {saldos_mes}")
else:
    print("No se encontraron PDFs.")
