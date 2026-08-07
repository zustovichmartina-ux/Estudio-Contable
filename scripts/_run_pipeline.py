"""Ejecuta el pipeline completo de conciliación con todos los PDFs disponibles."""
from __future__ import annotations

import hashlib
import os
import sys
import tempfile
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import database as db
from procesador import (
    aplicar_saldos_al_resultado,
    cargar_compras_tango,
    cargar_plan_cuentas,
    conciliar_movimientos,
    detectar_banco_pdf,
    extraer_movimientos_anuales,
    generar_planilla_conciliacion,
    generar_txt_tango,
    notificar_conciliacion_completada,
)

EXTRACTOS_DIR = BASE_DIR / "extractos bancarios"
EXPORT_DIR = BASE_DIR / "exportaciones"
CUIT_OBJETIVO = "30717847810"
NOMBRE_OBJETIVO = "Gastroenterolog"


def _hash_archivo(ruta: Path) -> str:
    h = hashlib.md5()
    with open(ruta, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _es_extracto_bancario(ruta: Path) -> bool:
    """Excluye PDFs de documentación que no son extractos."""
    nombre = ruta.name.lower()
    if "asientos contables" in nombre:
        return False
    if "informacion sobre tango" in str(ruta).lower():
        return False
    try:
        banco = detectar_banco_pdf(ruta)
        return banco in ("santander", "galicia", "frances", "credicoop", "provincia", "macro")
    except Exception:
        return True


def buscar_pdfs() -> list[Path]:
    """Recopila PDFs de extractos bancarios, proyecto y uploads Streamlit (temp)."""
    candidatos: list[Path] = []

    if EXTRACTOS_DIR.exists():
        candidatos.extend(sorted(EXTRACTOS_DIR.glob("*.pdf")))

    for pdf in BASE_DIR.rglob("*.pdf"):
        if pdf not in candidatos and _es_extracto_bancario(pdf):
            candidatos.append(pdf)

    import time
    temp_dir = Path(tempfile.gettempdir())
    cutoff = time.time() - 86400 * 2  # últimos 2 días
    for pdf in temp_dir.glob("tmp*.pdf"):
        try:
            if pdf.stat().st_mtime >= cutoff:
                candidatos.append(pdf)
        except OSError:
            continue

    # Deduplicar por hash de contenido
    vistos: dict[str, Path] = {}
    for pdf in candidatos:
        if not pdf.exists():
            continue
        try:
            h = _hash_archivo(pdf)
            if h not in vistos:
                vistos[h] = pdf
        except OSError:
            continue

    return sorted(vistos.values(), key=lambda p: p.name.lower())


def resolver_cliente() -> dict:
    db.inicializar_bd()
    clientes = db.listar_clientes()

    for c in clientes:
        if CUIT_OBJETIVO in str(c.get("cuit", "")).replace("-", ""):
            return c
        if NOMBRE_OBJETIVO.lower() in c.get("nombre", "").lower():
            return c

    for c in clientes:
        if c.get("tipo_persona") == "Persona Jurídica":
            return c

    if clientes:
        return clientes[0]

    raise RuntimeError("No hay clientes en la base de datos.")


def main() -> None:
    print("=" * 60)
    print("PIPELINE DE CONCILIACIÓN — Estudio Contable")
    print("=" * 60)

    pdfs = buscar_pdfs()
    print(f"\nPDFs encontrados ({len(pdfs)}):")
    for p in pdfs:
        print(f"   - {p}")

    if not pdfs:
        print("\nERROR: No se encontraron PDFs de extractos bancarios.")
        sys.exit(1)

    cliente = resolver_cliente()
    cuit = str(cliente["cuit"]).replace("-", "")
    nombre = cliente["nombre"]
    print(f"\nCliente: {nombre} (CUIT {cuit}) - {cliente['tipo_persona']}")

    print("\nExtrayendo movimientos anuales...")
    movimientos, bancos, saldos_por_mes = extraer_movimientos_anuales(pdfs)
    print(f"   Movimientos extraídos: {len(movimientos)}")
    print(f"   Bancos detectados: {', '.join(bancos)}")

    print("\nConciliando con Compras_Tango...")
    compras = cargar_compras_tango()
    resultado = conciliar_movimientos(movimientos, [], compras=compras)
    resultado = aplicar_saldos_al_resultado(resultado, saldos_por_mes)

    print("\nValidacion de balances:")
    balance_ok = True
    for clave in sorted(saldos_por_mes.keys()):
        b = saldos_por_mes[clave]
        estado = "CIERRA" if b.balance_cierra else "NO CIERRA"
        if not b.balance_cierra:
            balance_ok = False
        print(
            f"   {b.mes:02d}/{b.anio}: SI={b.saldo_inicial} | "
            f"Ingresos={b.ingresos} | Egresos={b.egresos} | "
            f"SF={b.saldo_final} | {estado}"
            + (f" (dif={b.diferencia_balance})" if not b.balance_cierra else "")
        )

    print(f"\n   Saldo inicial global: {resultado.saldo_inicial}")
    print(f"   Saldo final global:   {resultado.saldo_final}")
    print(f"   Balance general:      {'CIERRA' if resultado.balance_cierra else 'NO CIERRA'}")

    plan_path = cliente.get("plan_cuentas_path") or None
    plan_cuentas = cargar_plan_cuentas(plan_path)

    EXPORT_DIR.mkdir(exist_ok=True)
    ruta_xlsx = EXPORT_DIR / f"Conciliacion_{cuit}.xlsx"
    ruta_txt = EXPORT_DIR / f"Asientos_Tango_{cuit}.txt"

    print("\nGenerando planilla de conciliacion...")
    xlsx_bytes = generar_planilla_conciliacion(resultado, nombre)
    ruta_xlsx.write_bytes(xlsx_bytes)
    print(f"   OK {ruta_xlsx}")

    print("\nGenerando TXT Tango...")
    txt_content = generar_txt_tango(resultado, plan_cuentas)
    ruta_txt.write_text(txt_content, encoding="utf-8")
    print(f"   OK {ruta_txt} ({txt_content.count(chr(10))} lineas)")

    notif = notificar_conciliacion_completada(
        mensaje=f"Conciliación de {nombre} completada. {len(movimientos)} movimientos procesados."
    )
    print(f"\nNotificacion: {'enviada' if notif else 'no disponible'}")

    print("\n" + "=" * 60)
    print("RESUMEN FINAL")
    print("=" * 60)
    print(f"PDFs procesados:     {len(pdfs)}")
    for p in pdfs:
        print(f"  - {p.name}")
    print(f"Movimientos:         {len(movimientos)}")
    print(f"Conciliados:         {len(resultado.conciliados)}")
    print(f"Solo banco:          {len(resultado.solo_banco)}")
    print(f"Anomalías:           {len(resultado.anomalias)}")
    print(f"Saldo inicial:       {resultado.saldo_inicial}")
    print(f"Saldo final:         {resultado.saldo_final}")
    print(f"Balance validado:    {'Sí' if balance_ok else 'No — revisar diferencias'}")
    print(f"Planilla Excel:      {ruta_xlsx}")
    print(f"Asientos Tango TXT:  {ruta_txt}")
    print("=" * 60)


if __name__ == "__main__":
    main()
