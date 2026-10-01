# -*- coding: utf-8 -*-
"""Tests Portal IVA: dry-run sin Chrome, nombres, ZIP y ACTIONS."""
from __future__ import annotations

import io
import sys
import tempfile
import zipfile
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from afip_worker.actions import run_action, run_bajar_portal_iva
from afip_worker.actions.portal_iva import (
    es_boton_presentar,
    inferir_lado_archivo,
    materializar_descarga,
    nombre_destino_descarga,
)
from afip_worker.jobs import ACTIONS, create_job


def _job(tmp: Path, **params):
    base = {
        "ruta_destino": str(tmp),
        "periodo_desde": "2026-08-01",
        "periodo_hasta": "2026-08-31",
    }
    base.update(params)
    return create_job(
        cuit="27-42043034-0",
        razon_social="Camila Rocio Albarello",
        action="bajar_portal_iva",
        params=base,
        requested_by="test",
    )


def test_action_en_tuple() -> None:
    assert "bajar_portal_iva" in ACTIONS
    assert "bajar_comprobantes" in ACTIONS


def test_nombres_y_boton_presentar() -> None:
    assert inferir_lado_archivo("LIBRO_IVA_DIGITAL_VENTAS.csv") == "ventas"
    assert inferir_lado_archivo("foo.csv", "recibidos") == "compras"
    assert nombre_destino_descarga(
        "algo.csv",
        lado="compras",
        periodo_mm_yyyy="08-2026",
        cliente="Test",
    ) == "Compras.csv"
    assert "MisComprobantes Recibidos" in nombre_destino_descarga(
        "Mis Comprobantes Recibidos - CUIT 27420430340.xlsx",
        lado="compras",
        periodo_mm_yyyy="08-2026",
        cliente="Camila",
        desde="2026-08-01",
        hasta="2026-08-31",
    )
    assert es_boton_presentar("Presentar Libro IVA")
    assert es_boton_presentar("Confirmar presentación")
    assert not es_boton_presentar("Acuse de presentación")
    assert not es_boton_presentar("Vista previa")
    assert not es_boton_presentar("CSV")


def test_materializar_zip_csv(tmp: Path) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("LIBRO_COMPRAS.csv", "fecha;total\n")
    zpath = tmp / "in.zip"
    zpath.write_bytes(buf.getvalue())
    dest = tmp / "out"
    dest.mkdir()
    files = materializar_descarga(
        zpath,
        dest,
        lado="compras",
        periodo_mm_yyyy="08-2026",
        cliente="Test",
    )
    assert any(Path(f).name == "Compras.csv" for f in files)
    assert (dest / "Compras.csv").read_text(encoding="utf-8").startswith("fecha")


def test_dry_run_no_abre_chrome(tmp: Path) -> None:
    job = _job(tmp)
    with patch("afip_worker.actions.portal_iva.chrome_context") as mock_ctx:
        mock_ctx.side_effect = AssertionError("dry-run no debe abrir Chrome")
        result = run_bajar_portal_iva(job, dry_run=True)
    mock_ctx.assert_not_called()
    assert result.ok
    assert "[dry-run] bajar_portal_iva" in result.message
    assert any(f.endswith("Compras.csv") for f in result.files)
    assert any(f.endswith("Ventas.csv") for f in result.files)
    assert "08-2026" in result.files[0]
    dispatched = run_action(job, dry_run=True)
    assert dispatched.ok


def test_create_job_rechaza_action_inventada() -> None:
    try:
        create_job(
            cuit="27-42043034-0",
            razon_social="X",
            action="hackear_afip",  # type: ignore[arg-type]
            params={"ruta_destino": "C:\\tmp"},
        )
    except ValueError as exc:
        assert "inválida" in str(exc)
        return
    raise AssertionError("esperaba ValueError")


def main() -> int:
    test_action_en_tuple()
    test_nombres_y_boton_presentar()
    with tempfile.TemporaryDirectory(prefix="portal_iva_") as raw:
        tmp = Path(raw)
        test_materializar_zip_csv(tmp)
        test_dry_run_no_abre_chrome(tmp)
    test_create_job_rechaza_action_inventada()
    print("PASS portal_iva helpers + dry-run")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
