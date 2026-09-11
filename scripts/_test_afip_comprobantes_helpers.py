# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from afip_worker.naming import iso_to_ddmmyyyy, nombre_mis_comprobantes


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
    print("PASS helpers", name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
