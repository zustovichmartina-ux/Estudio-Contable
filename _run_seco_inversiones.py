# -*- coding: utf-8 -*-
"""Procesa lote Seco Matias → Excel FIFO (CA USD + FIMA + BIENES)."""
from __future__ import annotations

from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd

from inversiones import (
    aplicar_fifo,
    exportar_inversiones_excel,
    extraer_saldo_inicial_ddjj_pdf,
    procesar_archivos_inversiones,
    reclasificar_movimientos,
)

BASE_GAL = Path(r"T:/CLIENTES/SECO MATIAS/Ganancias/GANANCIAS/2025/Banco Galicia")
USD_DIR = next(p for p in (BASE_GAL / "Caja de Ahorro").iterdir() if p.is_dir() and "4017039" in p.name)
INV_DIR = BASE_GAL / "Inversiones"
BIENES = Path(r"C:/Users/recep/Desktop/SECO/BIENES.pdf")
OUT = Path(r"C:/Users/recep/Desktop/SECO/Seco_Inversiones_FIFO_2025.xlsx")


class U:
    def __init__(self, path: Path):
        self.name = path.name
        self._b = path.read_bytes()

    def getvalue(self):
        return self._b


def main() -> None:
    pdfs = sorted(USD_DIR.glob("*.pdf")) + sorted(INV_DIR.glob("RESUMEN*.pdf"))
    df_mov, errs = procesar_archivos_inversiones([U(p) for p in pdfs])
    df_mov = reclasificar_movimientos(df_mov)

    # Unificar nombres de especie
    if not df_mov.empty:
        df_mov.loc[
            df_mov["Especie"].astype(str).str.contains("PREMIUM", case=False, na=False),
            "Especie",
        ] = "FIMA PREMIUM CLASE A"
        df_mov.loc[
            df_mov["Especie"].astype(str).str.contains("RENTA FIJA DOLAR", case=False, na=False),
            "Especie",
        ] = "FIMA RENTA FIJA DOLARES CLASE A"

    print("Movimientos:", len(df_mov))
    print(df_mov.groupby(["Grupo", "Especie", "Tipo_Operacion"]).size().to_string())
    for e in errs:
        print("aviso:", e)

    df_ini, avisos = extraer_saldo_inicial_ddjj_pdf(BIENES.read_bytes(), BIENES.name)
    print("Avisos BIENES:", avisos)
    # Quedarnos solo con FCI reales / USD
    if not df_ini.empty:
        mask = df_ini["Especie"].astype(str).str.contains("FIMA|USD", case=False, na=False)
        df_ini = df_ini.loc[mask].copy()
    print("Saldo inicial filtrado:\n", df_ini.to_string(index=False) if not df_ini.empty else "(vacío)")

    rows = df_ini.to_dict("records") if not df_ini.empty else []

    # Apertura CA U$D Galicia 4017039 = 42.65 (01-2025 sin movimientos)
    if not any(str(r.get("Especie")) == "USD" for r in rows):
        rows.append({
            "Especie": "USD",
            "Grupo": "Dólar / MEP",
            "Cantidad": 42.65,
            "Costo_Unitario": 1.0,
            "Costo_Total": 42.65,
            "Moneda": "USD",
            "Origen": "Apertura CA U$D 4017039 (01-2025)",
            "Fecha": "01/01/1900",
        })
    else:
        # Preferir 42.65 de la cuenta operativa si BIENES trajo otra cifra chica
        for r in rows:
            if str(r.get("Especie")) == "USD" and float(r.get("Cantidad") or 0) < 10:
                r["Cantidad"] = 42.65
                r["Costo_Total"] = 42.65
                r["Costo_Unitario"] = 1.0
                r["Origen"] = "Apertura CA U$D 4017039 (ajustado desde extracto)"

    # Estimar / corregir cuotas FIMA Premium:
    # BIENES trae valuación sin cuotas; stock 31/12/2024 = back-calc desde RESUMEN_31-01-2025
    # (posición 103.717,83 − suscripciones + rescates de enero) = 35.808,63
    FIMA_OPEN_CUOTAS = 35808.63
    for r in rows:
        if str(r.get("Especie")) == "FIMA PREMIUM CLASE A":
            costo = float(r.get("Costo_Total") or 0)
            cant = float(r.get("Cantidad") or 0)
            if cant <= 1.01 and costo > 1000:
                r["Cantidad"] = FIMA_OPEN_CUOTAS
                r["Costo_Unitario"] = round(costo / FIMA_OPEN_CUOTAS, 6)
                r["Origen"] = f"{r.get('Origen')} (cuotas {FIMA_OPEN_CUOTAS} vía RESUMEN ene-2025)"
                print(f"FIMA Premium: {FIMA_OPEN_CUOTAS} cuotas, costo {costo}")
            elif cant < 1000 and costo > 1000:
                # ID de formulario AFIP mal leído como cantidad
                r["Cantidad"] = FIMA_OPEN_CUOTAS
                r["Costo_Unitario"] = round(costo / FIMA_OPEN_CUOTAS, 6)
                r["Origen"] = f"{r.get('Origen')} (cuotas {FIMA_OPEN_CUOTAS} vía RESUMEN ene-2025)"
                print(f"FIMA Premium corregido: {FIMA_OPEN_CUOTAS} cuotas, costo {costo}")
    if not any(str(r.get("Especie")) == "FIMA PREMIUM CLASE A" for r in rows):
        rows.append({
            "Especie": "FIMA PREMIUM CLASE A",
            "Grupo": "FCI",
            "Cantidad": FIMA_OPEN_CUOTAS,
            "Costo_Unitario": round(2024803.85 / FIMA_OPEN_CUOTAS, 6),
            "Costo_Total": 2024803.85,
            "Moneda": "ARS",
            "Origen": "BIENES + RESUMEN ene-2025 (manual)",
            "Fecha": "01/01/1900",
        })
        print(f"FIMA Premium agregado: {FIMA_OPEN_CUOTAS} cuotas")

    df_ini = pd.DataFrame(rows)

    # 11-2025.pdf está escaneado (sin texto). Puente por saldos de extracto:
    # cierre 10-2025 = 83.205,12 → apertura 12-2025 = 84.220,52 → net Nov = 1.015,40
    NOV_USD = 1015.40
    if not df_mov.empty:
        puente = pd.DataFrame([{
            "Fecha": "30/11/2025",
            "_fecha": pd.Timestamp("2025-11-30").date(),
            "Especie": "USD",
            "Grupo": "Dólar / MEP",
            "Tipo_Operacion": "Compra",
            "Cantidad": NOV_USD,
            "Precio": 1.0,
            "Monto_Total": NOV_USD,
            "Moneda": "USD",
            "Descripcion": "AJUSTE saldo 11-2025 (PDF escaneado; puente 83205.12→84220.52)",
            "Archivo origen": "11-2025.pdf (ajuste)",
            "Nueva_Clasificacion": None,
        }])
        df_mov = pd.concat([df_mov, puente], ignore_index=True)
        df_mov["_fecha"] = pd.to_datetime(df_mov["_fecha"], errors="coerce")
        df_mov = df_mov.sort_values("_fecha", kind="mergesort").reset_index(drop=True)
        print(f"Puente Nov USD: +{NOV_USD}")
        errs.append({
            "archivo": "11-2025.pdf",
            "motivo": f"PDF escaneado: se insertó ajuste USD +{NOV_USD} por diferencia de saldos extracto.",
        })

    res = aplicar_fifo(df_mov, df_ini)

    print("Aplicaciones:", len(res.aplicaciones), "Avisos FIFO:", len(res.avisos))
    for a in res.avisos[:20]:
        print(" ", a)
    saldos = pd.DataFrame(res.saldos)
    if not saldos.empty:
        print("Saldos relevantes:")
        print(
            saldos[saldos["Especie"].astype(str).str.contains("FIMA|USD", case=False, na=False)]
            .to_string(index=False)
        )

    xlsx = exportar_inversiones_excel(
        df_mov,
        res,
        df_inicial=df_ini,
        meta={"nota": "SECO MATIAS — Galicia CA U$D + FIMA 2025"},
    )
    OUT.write_bytes(xlsx)
    print("OK →", OUT, len(xlsx), "bytes")


if __name__ == "__main__":
    main()
