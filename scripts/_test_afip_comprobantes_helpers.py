# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from afip_worker.naming import (
    codigo_afip_a_fcc,
    dest_portal_iva,
    etiqueta_mm_yyyy,
    iso_to_ddmmyyyy,
    meses_del_periodo,
    nombre_fcc_compra,
    nombre_mis_comprobantes,
    nombre_portal_iva_csv,
    nombre_portal_iva_libro,
)


def main() -> int:
    assert iso_to_ddmmyyyy("2026-01-01") == "01/01/2026"
    assert iso_to_ddmmyyyy("2026-09-04") == "04/09/2026"
    name = nombre_mis_comprobantes(
        cliente="Martina Zustovich",
        tipo="Emitidos",
        desde="2026-01-01",
        hasta="2026-09-04",
        ext="xlsx",
    )
    assert name.startswith("Martina Zustovich MisComprobantes Emitidos")
    assert name.endswith(".xlsx")
    assert meses_del_periodo("2026-07-15", "2026-08-02") == [
        ("2026-07-15", "2026-07-31"),
        ("2026-08-01", "2026-08-02"),
    ]
    assert etiqueta_mm_yyyy("2026-08-01") == "08-2026"
    posix = dest_portal_iva("/tmp/Portal IVA", "2026-08-01")
    assert posix.name == "08-2026"
    assert dest_portal_iva("/tmp/Portal IVA/08-2026", "2026-08-01").name == "08-2026"
    unc = str(dest_portal_iva(r"\\TANGOSRV\x\Impuestos\Portal IVA", "2026-08-01"))
    assert unc.replace("/", "\\").endswith(r"\08-2026")
    already = str(dest_portal_iva(r"\\TANGOSRV\x\Impuestos\Portal IVA\08-2026", "2026-08-01"))
    assert already.replace("/", "\\").rstrip("\\").endswith("08-2026")
    assert already.count("08-2026") == 1
    assert nombre_portal_iva_csv("compras") == "Compras.csv"
    assert nombre_portal_iva_csv("ventas") == "Ventas.csv"
    assert nombre_portal_iva_libro("compras", "08-2026") == "LibroIVA_Compras_08-2026.pdf"
    assert codigo_afip_a_fcc(1) == ("FCC", "A")
    assert codigo_afip_a_fcc("003") == ("NCC", "A")
    assert nombre_fcc_compra(tipo="FCC", letra="A", pv=13, nro=157, proveedor="PROVEEDOR") == (
        "FCCA13-157 PROVEEDOR.pdf"
    )
    print("PASS helpers", name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
