# -*- coding: utf-8 -*-
"""Worker AFIP: lee jobs/pending, ejecuta (dry-run por defecto), mueve a done/error/needs_auth."""
from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

# Permitir `python -m afip_worker.main` desde la raíz del repo
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from afip_worker.actions import run_action
from afip_worker.auth import check_session_ready
from afip_worker.jobs import (
    claim_next,
    ensure_job_dirs,
    finish_job,
    jobs_root,
    login_in_progress,
    mark_needs_auth,
    write_job,
)
from afip_worker.server import start_api_thread
from afip_worker.token import load_or_create_token
from afip_worker.tunnel import start_cloudflared_tunnel

LOG = logging.getLogger("afip_worker")


def process_one(*, dry_run: bool = True, root: Path | None = None) -> bool:
    """Procesa un job. True si tomó uno; False si la cola pending está vacía."""
    root = root or jobs_root()
    ensure_job_dirs(root)
    if login_in_progress(root):
        LOG.info("login visible en curso; pauso la cola")
        return False
    job = claim_next(root)
    if not job:
        return False

    LOG.info("running %s action=%s cuit=%s", job.id, job.action, job.cuit)
    write_job(job, root)

    chk = check_session_ready(job, dry_run=dry_run)
    if not chk.ready:
        mark_needs_auth(job, chk.note, root)
        LOG.warning("needs_auth %s: %s", job.id, chk.note)
        return True

    job.auth.status = "ready"
    job.auth.note = chk.note
    write_job(job, root)

    result = run_action(job, dry_run=dry_run)
    if result.needs_auth:
        mark_needs_auth(job, result.message, root)
        LOG.warning("needs_auth %s: %s", job.id, result.message)
        return True

    finish_job(
        job,
        ok=result.ok,
        message=result.message,
        files=result.files,
        extra=getattr(result, "extra", None) or {},
        root=root,
    )
    LOG.info(
        "%s %s: %s",
        "done" if result.ok else "error",
        job.id,
        result.message,
    )
    return True


def loop(*, dry_run: bool = True, interval: float = 3.0, once: bool = False, root: Path | None = None) -> None:
    root = root or jobs_root()
    ensure_job_dirs(root)
    LOG.info("worker start dry_run=%s root=%s", dry_run, root)
    while True:
        worked = process_one(dry_run=dry_run, root=root)
        if once:
            break
        if not worked:
            time.sleep(interval)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Worker AFIP (cola jobs/)")
    parser.add_argument("--dry-run", action="store_true", default=True, help="Sin AFIP real (default si no hay --live)")
    parser.add_argument("--live", action="store_true", help="Comprobantes en Línea en segundo plano (sesión persistente)")
    parser.add_argument("--once", action="store_true", help="Procesar un solo job y salir")
    parser.add_argument("--interval", type=float, default=3.0, help="Segundos entre polls")
    parser.add_argument("--jobs-root", type=str, default="", help="Override AFIP_JOBS_ROOT")
    parser.add_argument("--serve", action="store_true", help="API HTTP para Streamlit Cloud")
    parser.add_argument("--no-tunnel", action="store_true", help="API local sin cloudflared")
    parser.add_argument("--host", type=str, default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    dry_run = not args.live
    root = Path(args.jobs_root) if args.jobs_root else None
    if root:
        ensure_job_dirs(root)

    if args.serve:
        token = load_or_create_token()
        start_api_thread(host=args.host, port=args.port, token=token)
        LOG.info("AFIP_WORKER_TOKEN listo (jobs/.worker_token). No lo subas a git.")
        if not args.no_tunnel:
            start_cloudflared_tunnel(port=args.port, token=token)

    loop(dry_run=dry_run, interval=args.interval, once=args.once, root=root)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
