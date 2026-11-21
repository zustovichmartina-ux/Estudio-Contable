"""Genera datos INVENTADOS para la demo de punto de equilibrio.

Crea en datos_ejemplo/:
  - ventas_bistrosoft_noviembre.csv  -> simula un export de ventas ticket por ticket
  - gastos_fijos.csv                 -> gastos fijos del mes que carga el estudio
  - costos_productos.csv             -> precio y costo de receta por artículo
  - parametros.csv                   -> % variables y días de apertura

El formato del export es una suposición: cuando tengamos un export real de
Bistrosoft se ajusta el lector (punto_equilibrio.py), no este archivo.
"""
from __future__ import annotations

import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

CARPETA = Path(__file__).parent / "datos_ejemplo"

# artículo, rubro, precio, costo de receta, peso en la venta (popularidad)
CATALOGO = [
    ("Espresso", "Cafetería", 3200, 700, 40),
    ("Cortado", "Cafetería", 3800, 900, 45),
    ("Café con leche", "Cafetería", 4500, 1200, 35),
    ("Capuchino", "Cafetería", 5200, 1500, 30),
    ("Medialuna", "Panadería", 1500, 450, 90),
    ("Tostado J&Q", "Cocina", 8500, 3000, 15),
    ("Avocado toast", "Cocina", 11000, 4200, 8),
    ("Porción de torta", "Pastelería", 6500, 2200, 12),
    ("Jugo de naranja", "Bebidas", 5500, 1900, 15),
    ("Agua mineral", "Bebidas", 2500, 1400, 20),
    ("Brunch completo", "Cocina", 18000, 7000, 5),
]

GASTOS_FIJOS = [
    ("Alquiler", 3_500_000),
    ("Sueldos y cargas sociales", 9_000_000),
    ("Luz, gas, agua e internet", 1_200_000),
    ("Honorarios del estudio", 600_000),
    ("Sistema de gestión", 150_000),
    ("Seguros", 250_000),
    ("Mantenimiento y limpieza", 400_000),
    ("Tasas municipales y monotributos", 500_000),
]

PARAMETROS = [
    ("comision_medios_pago_pct", "3.5"),
    ("ingresos_brutos_pct", "3.0"),
    ("dias_apertura_mes", "30"),
]

# Venta objetivo por día de la semana (lunes=0 ... domingo=6), en pesos
OBJETIVO_DIA = {0: 760_000, 1: 800_000, 2: 860_000, 3: 900_000, 4: 1_100_000, 5: 1_350_000, 6: 1_250_000}
DIAS_LLUVIA = {10, 17}
MEDIOS = [("Efectivo", 30), ("Débito", 25), ("Crédito", 15), ("QR Mercado Pago", 30)]


def _fmt_ar(n: float) -> str:
    """1234567.5 -> '1.234.567,50' (como exporta un sistema en castellano)."""
    entero, dec = f"{n:,.2f}".split(".")
    return entero.replace(",", ".") + "," + dec


def generar_ventas(anio: int = 2026, mes: int = 11, hasta_dia: int = 21, semilla: int = 2611) -> list[dict]:
    rnd = random.Random(semilla)
    pesos = [c[4] for c in CATALOGO]
    filas: list[dict] = []
    ticket = 10_000
    for dia in range(1, hasta_dia + 1):
        f = date(anio, mes, dia)
        objetivo = OBJETIVO_DIA[f.weekday()] * (0.8 if dia in DIAS_LLUVIA else 1) * rnd.uniform(0.9, 1.1)
        vendido = 0.0
        while vendido < objetivo:
            ticket += 1
            hora = datetime(anio, mes, dia, 7, 30) + timedelta(minutes=rnd.randint(0, 12 * 60 + 30))
            medio = rnd.choices([m[0] for m in MEDIOS], [m[1] for m in MEDIOS])[0]
            items: dict[str, int] = {}
            for _ in range(rnd.choice([1, 1, 2, 2, 2, 3, 4])):
                art = rnd.choices(CATALOGO, pesos)[0]
                items[art[0]] = items.get(art[0], 0) + 1
            for art in CATALOGO:
                if art[0] not in items:
                    continue
                cant = items[art[0]]
                importe = cant * art[2]
                vendido += importe
                filas.append({
                    "Fecha": f.strftime("%d/%m/%Y"),
                    "Hora": hora.strftime("%H:%M"),
                    "Nro Ticket": ticket,
                    "Artículo": art[0],
                    "Rubro": art[1],
                    "Cantidad": cant,
                    "Precio Unit.": _fmt_ar(art[2]),
                    "Importe": _fmt_ar(importe),
                    "Medio de Pago": medio,
                })
    return filas


def escribir(ruta: Path, filas: list[dict] | list[tuple], encabezado: list[str]) -> None:
    with ruta.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(encabezado)
        for fila in filas:
            w.writerow(fila.values() if isinstance(fila, dict) else fila)


def main() -> None:
    CARPETA.mkdir(exist_ok=True)
    ventas = generar_ventas()
    escribir(CARPETA / "ventas_bistrosoft_noviembre.csv", ventas, list(ventas[0].keys()))
    escribir(CARPETA / "gastos_fijos.csv", GASTOS_FIJOS, ["concepto", "importe_mensual"])
    escribir(CARPETA / "costos_productos.csv", [(c[0], c[2], c[3]) for c in CATALOGO], ["articulo", "precio", "costo_receta"])
    escribir(CARPETA / "parametros.csv", PARAMETROS, ["parametro", "valor"])
    print(f"{len(ventas)} líneas de venta generadas en {CARPETA}")


if __name__ == "__main__":
    main()
