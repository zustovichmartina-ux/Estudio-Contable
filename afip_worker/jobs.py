# -*- coding: utf-8 -*-
"""Modelo Job JSON + carpetas pending/running/done/error/needs_auth."""
from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Literal
from zoneinfo import ZoneInfo

TZ = ZoneInfo("America/Argentina/Buenos_Aires")

Action = Literal["emitir_fcc", "bajar_veps", "bajar_comprobantes"]
Status = Literal["pending", "running", "needs_auth", "done", "error"]
AuthStatus = Literal["unknown", "ready", "needs_admin", "failed"]

ACTIONS: tuple[Action, ...] = ("emitir_fcc", "bajar_veps", "bajar_comprobantes")
STATUSES: tuple[Status, ...] = ("pending", "running", "needs_auth", "done", "error")

_INVALID = re.compile(r'[\\/:*?"<>|]+')


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


def jobs_root() -> Path:
    """Raíz de la cola. Override con env AFIP_JOBS_ROOT (UNC o local)."""
    raw = (os.environ.get("AFIP_JOBS_ROOT") or "").strip()
    root = Path(raw) if raw else _repo_root() / "jobs"
    ensure_job_dirs(root)
    return root


def ensure_job_dirs(root: Path | None = None) -> Path:
    root = root or jobs_root()
    for name in ("pending", "running", "needs_auth", "done", "error"):
        (root / name).mkdir(parents=True, exist_ok=True)
    return root


def uploads_dir() -> Path:
    d = _repo_root() / "uploads" / "afip"
    d.mkdir(parents=True, exist_ok=True)
    return d


@dataclass
class JobAuth:
    status: AuthStatus = "unknown"
    note: str = ""


@dataclass
class JobResult:
    files: list[str] = field(default_factory=list)
    message: str = ""


@dataclass
class Job:
    id: str
    created_at: str
    requested_by: str
    cuit: str
    razon_social: str
    action: Action
    params: dict[str, Any] = field(default_factory=dict)
    auth: JobAuth = field(default_factory=JobAuth)
    status: Status = "pending"
    result: JobResult = field(default_factory=JobResult)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "created_at": self.created_at,
            "requested_by": self.requested_by,
            "cuit": self.cuit,
            "razon_social": self.razon_social,
            "action": self.action,
            "params": dict(self.params or {}),
            "auth": asdict(self.auth) if isinstance(self.auth, JobAuth) else dict(self.auth or {}),
            "status": self.status,
            "result": asdict(self.result) if isinstance(self.result, JobResult) else dict(self.result or {}),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Job":
        auth_raw = data.get("auth") or {}
        res_raw = data.get("result") or {}
        action = str(data.get("action") or "")
        if action not in ACTIONS:
            raise ValueError(f"action inválida: {action}")
        status = str(data.get("status") or "pending")
        if status not in STATUSES:
            raise ValueError(f"status inválido: {status}")
        return cls(
            id=str(data["id"]),
            created_at=str(data.get("created_at") or ""),
            requested_by=str(data.get("requested_by") or ""),
            cuit=normalizar_cuit(str(data.get("cuit") or "")),
            razon_social=str(data.get("razon_social") or ""),
            action=action,  # type: ignore[arg-type]
            params=dict(data.get("params") or {}),
            auth=JobAuth(
                status=str(auth_raw.get("status") or "unknown"),  # type: ignore[arg-type]
                note=str(auth_raw.get("note") or ""),
            ),
            status=status,  # type: ignore[arg-type]
            result=JobResult(
                files=list(res_raw.get("files") or []),
                message=str(res_raw.get("message") or ""),
            ),
        )


def normalizar_cuit(cuit: str) -> str:
    digits = re.sub(r"\D", "", cuit or "")
    if len(digits) == 11:
        return f"{digits[:2]}-{digits[2:10]}-{digits[10]}"
    return (cuit or "").strip()


def require_cuit(cuit: str) -> str:
    """Normaliza y exige 11 dígitos (evita CUITs truncados al encolar)."""
    norm = normalizar_cuit(cuit)
    digits = re.sub(r"\D", "", norm)
    if len(digits) != 11:
        raise ValueError(
            f"CUIT inválido ({cuit!r}): se necesitan 11 dígitos, hay {len(digits)}"
        )
    return norm


def limpiar_ruta(ruta: str) -> str:
    """Quita comillas y espacios que suelen pegarse desde Excel/UI."""
    return (ruta or "").strip().strip('"').strip("'").strip()


def new_job_id(action: str, razon: str = "") -> str:
    now = datetime.now(TZ)
    slug = _INVALID.sub("", (razon or "job").lower().replace(" ", "-"))[:24].strip("-") or "job"
    act = (action or "job").replace("_", "-")[:16]
    return f"{now.strftime('%Y%m%d-%H%M%S')}-{act}-{slug}"


def create_job(
    *,
    cuit: str,
    razon_social: str,
    action: Action,
    params: dict[str, Any] | None = None,
    requested_by: str = "admin",
    job_id: str | None = None,
) -> Job:
    if action not in ACTIONS:
        raise ValueError(f"action inválida: {action}")
    now = datetime.now(TZ)
    clean_params = dict(params or {})
    if "ruta_destino" in clean_params:
        clean_params["ruta_destino"] = limpiar_ruta(str(clean_params["ruta_destino"]))
    return Job(
        id=job_id or new_job_id(action, razon_social),
        created_at=now.isoformat(timespec="seconds"),
        requested_by=requested_by,
        cuit=require_cuit(cuit),
        razon_social=(razon_social or "").strip(),
        action=action,
        params=clean_params,
        auth=JobAuth(status="unknown"),
        status="pending",
        result=JobResult(),
    )


def job_path(job: Job | str, status: Status, root: Path | None = None) -> Path:
    root = root or jobs_root()
    jid = job.id if isinstance(job, Job) else job
    return root / status / f"{jid}.json"


def write_job(job: Job, root: Path | None = None) -> Path:
    root = root or jobs_root()
    path = job_path(job, job.status, root)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(job.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)
    return path


def read_job(path: Path) -> Job:
    data = json.loads(path.read_text(encoding="utf-8"))
    return Job.from_dict(data)


def enqueue_job(job: Job, root: Path | None = None) -> Path:
    job.status = "pending"
    return write_job(job, root)


def list_jobs(
    status: Status | None = None,
    root: Path | None = None,
) -> list[Job]:
    root = root or jobs_root()
    folders = [status] if status else list(STATUSES)
    out: list[Job] = []
    for folder in folders:
        d = root / folder
        if not d.is_dir():
            continue
        for p in sorted(d.glob("*.json")):
            try:
                out.append(read_job(p))
            except Exception:
                continue
    out.sort(key=lambda j: j.created_at)
    return out


def next_pending(root: Path | None = None) -> Job | None:
    pending = list_jobs("pending", root)
    return pending[0] if pending else None


def move_job(job: Job, new_status: Status, root: Path | None = None) -> Path:
    """Mueve el JSON a la carpeta del nuevo status (lock simple por rename)."""
    root = root or jobs_root()
    old = job_path(job, job.status, root)
    job.status = new_status
    dest = job_path(job, new_status, root)
    dest.parent.mkdir(parents=True, exist_ok=True)
    write_job(job, root)
    if old.exists() and old.resolve() != dest.resolve():
        try:
            old.unlink()
        except OSError:
            pass
    # Limpiar copias huérfanas en otras carpetas
    for st in STATUSES:
        if st == new_status:
            continue
        orphan = job_path(job, st, root)
        if orphan.exists() and orphan.resolve() != dest.resolve():
            try:
                orphan.unlink()
            except OSError:
                pass
    return dest


def claim_next(root: Path | None = None) -> Job | None:
    """Toma el primer pending y lo pasa a running. None si no hay."""
    root = root or jobs_root()
    job = next_pending(root)
    if not job:
        return None
    move_job(job, "running", root)
    return job


def finish_job(
    job: Job,
    *,
    ok: bool,
    message: str = "",
    files: list[str] | None = None,
    root: Path | None = None,
) -> Path:
    job.result = JobResult(files=list(files or []), message=message)
    return move_job(job, "done" if ok else "error", root)


def mark_needs_auth(job: Job, note: str = "", root: Path | None = None) -> Path:
    job.auth = JobAuth(status="needs_admin", note=note)
    job.result = JobResult(message=note or "Se requiere login AFIP / 2FA del admin")
    return move_job(job, "needs_auth", root)


def clear_queue(root: Path | None = None, *, statuses: tuple[Status, ...] | None = None) -> int:
    """Solo para tests. Borra JSON en carpetas indicadas."""
    root = root or jobs_root()
    n = 0
    for st in statuses or STATUSES:
        for p in (root / st).glob("*.json"):
            p.unlink(missing_ok=True)
            n += 1
    return n


def copy_tree_empty(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)
