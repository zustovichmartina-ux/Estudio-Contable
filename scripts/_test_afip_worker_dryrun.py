# -*- coding: utf-8 -*-
"""Test dry-run: encola 3 jobs (una por acción) y verifica movimiento de carpetas."""
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


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="afip_jobs_"))
    ensure_job_dirs(tmp)
    dest = tmp / "destino_fake"
    dest.mkdir()
    registry = tmp / "cuit_registry.json"
    os.environ["AFIP_JOBS_ROOT"] = str(tmp)
    os.environ["AFIP_CUIT_REGISTRY"] = str(registry)

    specs = [
        ("emitir_fcc", {"ruta_destino": str(dest), "plantilla_excel": "uploads/fake.xlsx", "fecha_emision": "2026-08-31"}),
        ("bajar_veps", {"ruta_destino": str(dest), "periodo_desde": "2026-08-01", "periodo_hasta": "2026-08-31"}),
        ("bajar_comprobantes", {"ruta_destino": str(dest), "periodo_desde": "2026-08-01", "periodo_hasta": "2026-08-31"}),
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

    assert len(list_jobs("pending", tmp)) == 3

    for _ in range(3):
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
    assert len(done) == 3
    done_ids = {j.id for j in done}
    assert set(ids) == done_ids

    # needs_auth path
    job_auth = create_job(
        cuit="20-00000000-0",
        razon_social="Test Auth",
        action="bajar_veps",
        params={"ruta_destino": str(dest), "force_needs_auth": True, "periodo_desde": "2026-01-01", "periodo_hasta": "2026-01-31"},
    )
    enqueue_job(job_auth, tmp)
    assert process_one(dry_run=True, root=tmp)
    assert len(list_jobs("needs_auth", tmp)) == 1
    print("needs_auth OK", list_jobs("needs_auth", tmp)[0].id)

    shutil.rmtree(tmp, ignore_errors=True)
    print("PASS dry-run 3 jobs + needs_auth")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
