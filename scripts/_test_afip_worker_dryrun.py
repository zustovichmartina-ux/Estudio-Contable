# -*- coding: utf-8 -*-
"""Test dry-run: encola 4 jobs (una por acción) y verifica movimiento de carpetas."""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from afip_worker.jobs import (
    create_job,
    enqueue_job,
    ensure_job_dirs,
    list_jobs,
)
from afip_worker.main import process_one
from afip_worker.registry import mark_ready


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="afip_jobs_"))
    ensure_job_dirs(tmp)
    dest = tmp / "destino_fake"
    dest.mkdir()
    registry = tmp / "cuit_registry.json"
    os.environ["AFIP_JOBS_ROOT"] = str(tmp)
    os.environ["AFIP_CUIT_REGISTRY"] = str(registry)

    # CUIT con acceso: worker actúa solo
    mark_ready("27-42043034-0", razon_social="Camila Rocio Albarello", note="test ready")

    specs = [
        ("emitir_fcc", {"ruta_destino": str(dest), "plantilla_excel": "uploads/fake.xlsx", "fecha_emision": "2026-08-31"}),
        ("bajar_veps", {"ruta_destino": str(dest), "periodo_desde": "2026-08-01", "periodo_hasta": "2026-08-31"}),
        ("bajar_comprobantes", {
            "ruta_destino": str(dest),
            "periodo_desde": "2026-08-01",
            "periodo_hasta": "2026-08-31",
            "analizar_monotributo": True,
        }),
        ("bajar_portal_iva", {
            "ruta_destino": str(dest),
            "periodo_desde": "2026-08-01",
            "periodo_hasta": "2026-08-31",
        }),
    ]
    ids = []
    for action, params in specs:
        job = create_job(
            cuit="27-42043034-0",
            razon_social="Camila Rocio Albarello",
            action=action,  # type: ignore[arg-type]
            params=params,
            requested_by="test",
        )
        enqueue_job(job, tmp)
        ids.append(job.id)
        print("enqueued", job.id, action)

    assert len(list_jobs("pending", tmp)) == 4

    for _ in range(4):
        ok = process_one(dry_run=True, root=tmp)
        assert ok, "esperaba un job pending"

    pending = list_jobs("pending", tmp)
    running = list_jobs("running", tmp)
    done = list_jobs("done", tmp)
    error = list_jobs("error", tmp)
    needs = list_jobs("needs_auth", tmp)

    print("pending", len(pending), "running", len(running), "done", len(done), "error", len(error), "needs_auth", len(needs))
    assert len(pending) == 0, pending
    assert len(running) == 0, running
    assert len(error) == 0, [e.result.message for e in error]
    assert len(needs) == 0
    assert len(done) == 4
    done_ids = {j.id for j in done}
    assert set(ids) == done_ids
    cmpte = next(j for j in done if j.action == "bajar_comprobantes")
    extra_mono = (cmpte.result.extra or {}).get("monotributo") or {}
    assert extra_mono.get("cantidad") == 0
    assert extra_mono.get("errores")

    portal = next(j for j in done if j.action == "bajar_portal_iva")
    assert "[dry-run] bajar_portal_iva" in portal.result.message
    assert any("Compras.csv" in f for f in portal.result.files)

    # needs_auth: CUIT sin acceso (nuevo) → frena solo
    job_auth = create_job(
        cuit="20-00000000-0",
        razon_social="Test Auth",
        action="bajar_veps",
        params={
            "ruta_destino": str(dest),
            "periodo_desde": "2026-01-01",
            "periodo_hasta": "2026-01-31",
        },
    )
    enqueue_job(job_auth, tmp)
    assert process_one(dry_run=True, root=tmp)
    assert len(list_jobs("needs_auth", tmp)) == 1
    print("needs_auth OK (CUIT sin acceso)", list_jobs("needs_auth", tmp)[0].id)

    # force_needs_auth explícito también frena
    job_force = create_job(
        cuit="27-42043034-0",
        razon_social="Camila Rocio Albarello",
        action="bajar_veps",
        params={
            "ruta_destino": str(dest),
            "force_needs_auth": True,
            "periodo_desde": "2026-01-01",
            "periodo_hasta": "2026-01-31",
        },
    )
    enqueue_job(job_force, tmp)
    assert process_one(dry_run=True, root=tmp)
    assert len(list_jobs("needs_auth", tmp)) == 2
    print("needs_auth OK (force)", job_force.id)

    shutil.rmtree(tmp, ignore_errors=True)
    print("PASS dry-run 4 jobs + needs_auth")

    from afip_worker.auth import check_session_ready
    from afip_worker.estudio import cuit_login_recepcion
    from afip_worker.jobs import create_job as _cj

    live_job = _cj(
        cuit="20-11111111-1",
        razon_social="Sin permiso previo",
        action="bajar_comprobantes",
        params={"ruta_destino": "/tmp", "periodo_desde": "2026-08-01", "periodo_hasta": "2026-08-31"},
    )
    live_chk = check_session_ready(live_job, dry_run=False)
    assert live_chk.ready, live_chk
    print("PASS live intenta AFIP aunque el CUIT no esté Listo")

    login = cuit_login_recepcion()
    assert len(login) == 11
    assert login != "20111111111"
    print("PASS login RECEPCION distinto al CUIT del cliente")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
