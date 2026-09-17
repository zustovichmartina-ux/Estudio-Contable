# -*- coding: utf-8 -*-
"""Prueba del convertidor de liquidaciones (formato estudio)."""
from __future__ import annotations

from pathlib import Path

from liquidaciones_tarjetas_estudio import procesar_pdfs_tarjetas_estudio


def test_recife_agosto_si_hay_red() -> None:
    carpeta = Path(
        r"\\TANGOSRV\Compartido\CLIENTES\GLOBAL RECIFE SA"
        r"\Cierre 2026 del 01-01-2026 al 31-01-2026\202608\LIQUIDACIONES DE TARJETAS"
    )
    if not carpeta.is_dir():
        return
    pdfs = []
    seen: set[str] = set()
    for p in sorted(carpeta.iterdir()):
        if p.suffix.lower() != ".pdf":
            continue
        if p.name.lower() in seen:
            continue
        seen.add(p.name.lower())
        pdfs.append((p.name, p.read_bytes()))
    assert pdfs, "hay carpeta pero no PDF"
    res = procesar_pdfs_tarjetas_estudio(pdfs)
    assert res.ok, res.error
    assert res.n_fd == 69, res.n_fd
    assert res.n_control_ok == 69, (res.n_control_ok, res.n_control_diff)
    assert res.n_ventas >= 200, res.n_ventas
    assert res.excel_bytes[:2] == b"PK"
    assert "First Data" in res.hojas
    assert "Movimientos" in res.hojas
    assert "Control" in res.hojas


if __name__ == "__main__":
    test_recife_agosto_si_hay_red()
    print("OK test_recife_agosto_si_hay_red")
