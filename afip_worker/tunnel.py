# -*- coding: utf-8 -*-
"""Túnel Cloudflare para exponer la API local a Streamlit Cloud.

Por defecto usa un túnel rápido (`trycloudflare.com`), cuya URL cambia al
reiniciar. Si hay un túnel con nombre configurado, arranca
`cloudflared tunnel --config … run` y publica el hostname estable.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from .jobs import _repo_root
from .token import bridge_info_path

LOG = logging.getLogger("afip_worker.tunnel")
URL_RE = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com", re.I)
HOSTNAME_RE = re.compile(
    r"(?im)^\s*-?\s*hostname:\s*[\"']?([^\"'#\n]+)"
)
SERVICE_RE = re.compile(
    r"(?im)^\s*-?\s*service:\s*[\"']?(https?://[^\s\"'#]+)"
)
CREDENTIALS_RE = re.compile(
    r"(?im)^\s*-?\s*credentials-file:\s*[\"']?([^\"'#\n]+)"
)
LOCAL_ORIGIN_RE = re.compile(r"https?://127\.0\.0\.1:(\d+)", re.I)

NamedKind = Literal["config", "token"]


@dataclass(frozen=True)
class NamedTunnelSpec:
    """Plan para un túnel Cloudflare con hostname estable."""

    kind: NamedKind
    public_url: str | None
    config_path: Path | None = None
    token: str | None = None
    error: str | None = None


def cloudflared_dir() -> Path:
    return _repo_root() / "jobs" / "cloudflared"


def find_cloudflared() -> str | None:
    env = (os.environ.get("CLOUDFLARED") or "").strip()
    if env and Path(env).exists():
        return env
    found = shutil.which("cloudflared")
    if found:
        return found
    for cand in (
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WinGet" / "Links" / "cloudflared.exe",
        Path(r"C:\Program Files\cloudflared\cloudflared.exe"),
        Path(r"C:\Program Files (x86)\cloudflared\cloudflared.exe"),
        _repo_root() / "tools" / "cloudflared.exe",
    ):
        if cand.exists():
            return str(cand)
    return None


def normalize_public_url(value: str) -> str:
    v = (value or "").strip().rstrip("/")
    if not v:
        return ""
    if not re.match(r"^https?://", v, re.I):
        v = "https://" + v
    return v


def parse_ingress_hostname(config_text: str) -> str | None:
    """Primer hostname público de un config.yml de cloudflared."""
    for m in HOSTNAME_RE.finditer(config_text or ""):
        host = m.group(1).strip().strip("\"'")
        if not host:
            continue
        if host.lower() in {"localhost", "127.0.0.1"}:
            continue
        return host
    return None


def parse_local_service(config_text: str) -> str | None:
    m = SERVICE_RE.search(config_text or "")
    if not m:
        return None
    return m.group(1).strip()


def parse_credentials_file(config_text: str) -> str | None:
    m = CREDENTIALS_RE.search(config_text or "")
    if not m:
        return None
    return m.group(1).strip().strip("\"'")


def hostname_override() -> str:
    for key in ("AFIP_TUNNEL_HOSTNAME", "AFIP_WORKER_PUBLIC_URL"):
        raw = (os.environ.get(key) or "").strip()
        if raw:
            return raw
    return ""


def named_config_path() -> Path | None:
    """Ruta al config.yml: env `AFIP_CLOUDFLARED_CONFIG` o `jobs/cloudflared/config.yml`."""
    env = (os.environ.get("AFIP_CLOUDFLARED_CONFIG") or "").strip()
    if env:
        return Path(env)
    for name in ("config.yml", "config.yaml"):
        cand = cloudflared_dir() / name
        if cand.is_file():
            return cand
    return None


def named_tunnel_token() -> str:
    for key in ("AFIP_CLOUDFLARED_TOKEN", "TUNNEL_TOKEN"):
        raw = (os.environ.get(key) or "").strip()
        if raw:
            return raw
    return ""


def detect_named_tunnel() -> NamedTunnelSpec | None:
    """Si hay config/token de túnel con nombre, devolver el spec; si no, None (túnel rápido)."""
    cfg = named_config_path()
    env_cfg = (os.environ.get("AFIP_CLOUDFLARED_CONFIG") or "").strip()
    if env_cfg and cfg is not None and not cfg.is_file():
        return NamedTunnelSpec(
            kind="config",
            public_url=normalize_public_url(hostname_override()) or None,
            config_path=cfg,
            error=f"AFIP_CLOUDFLARED_CONFIG no existe: {cfg}",
        )
    if cfg is not None and cfg.is_file():
        text = cfg.read_text(encoding="utf-8", errors="replace")
        host = hostname_override() or parse_ingress_hostname(text) or ""
        url = normalize_public_url(host) or None
        cred = parse_credentials_file(text)
        error = None
        if not url:
            error = (
                "Túnel con nombre: no hay hostname en config.yml. "
                "Definí AFIP_TUNNEL_HOSTNAME (ej. https://afip-worker.tudominio.com)."
            )
        if cred:
            cred_path = Path(cred)
            if not cred_path.is_absolute():
                for base in (cfg.parent, _repo_root(), cloudflared_dir()):
                    cand = base / cred
                    if cand.is_file():
                        cred_path = cand
                        break
            if not cred_path.is_file():
                warn = f"credentials-file no encontrado: {cred}"
                error = f"{error} {warn}".strip() if error else warn
        return NamedTunnelSpec(
            kind="config",
            public_url=url,
            config_path=cfg,
            error=error,
        )

    token = named_tunnel_token()
    if token:
        url = normalize_public_url(hostname_override()) or None
        error = None
        if not url:
            error = (
                "Túnel con token: definí AFIP_TUNNEL_HOSTNAME "
                "(hostname público de Zero Trust)."
            )
        return NamedTunnelSpec(kind="token", public_url=url, token=token, error=error)
    return None


def named_tunnel_command(exe: str, spec: NamedTunnelSpec) -> list[str]:
    cmd = [exe, "tunnel", "--no-autoupdate"]
    if spec.kind == "config" and spec.config_path is not None:
        cmd.extend(["--config", str(spec.config_path), "run"])
        return cmd
    if spec.kind == "token" and spec.token:
        cmd.extend(["run", "--token", spec.token])
        return cmd
    raise ValueError("NamedTunnelSpec incompleto para armar el comando")


def quick_tunnel_command(exe: str, port: int) -> list[str]:
    return [exe, "tunnel", "--no-autoupdate", "--url", f"http://127.0.0.1:{port}"]


def cmd_for_log(cmd: list[str]) -> str:
    """Oculta valores de flags sensibles al loguear el comando."""
    redacted: list[str] = []
    hide_next = False
    for part in cmd:
        if hide_next:
            redacted.append("<redacted>")
            hide_next = False
            continue
        if part in {"--token", "--credentials-file"}:
            redacted.append(part)
            hide_next = True
            continue
        redacted.append(part)
    return " ".join(redacted)


def write_bridge_info(*, url: str, token: str, port: int, stable: bool = False) -> Path:
    path = bridge_info_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    if stable:
        note = (
            "Túnel con nombre (URL estable). No hace falta actualizar Secrets "
            "al reiniciar el worker, salvo que cambie el token."
        )
    else:
        note = (
            "El túnel rápido cambia de URL si reiniciás el worker: actualizá Secrets."
        )
    path.write_text(
        "\n".join(
            [
                "Pegá esto en Streamlit Cloud → Manage app → Settings → Secrets:",
                "",
                f'AFIP_WORKER_URL = "{url}"',
                f'AFIP_WORKER_TOKEN = "{token}"',
                "",
                f"API local: http://127.0.0.1:{port}",
                note,
                "",
            ]
        ),
        encoding="utf-8",
    )
    return path


def _drain_stdout(proc: subprocess.Popen[str]) -> None:
    def _run() -> None:
        if not proc.stdout:
            return
        try:
            for line in proc.stdout:
                if line.strip():
                    LOG.info("cloudflared: %s", line.strip())
        except Exception:
            return

    threading.Thread(target=_run, name="cloudflared-log", daemon=True).start()


def _spawn(cmd: list[str]) -> subprocess.Popen[str]:
    return subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _read_until(*, proc: subprocess.Popen[str], deadline: float, url_from_logs: bool) -> str | None:
    buf = ""
    found: str | None = None
    while time.time() < deadline:
        if proc.poll() is not None:
            rest = proc.stdout.read() if proc.stdout else ""
            LOG.error("cloudflared salió: %s%s", buf, rest)
            if url_from_logs:
                m = URL_RE.search(buf + (rest or ""))
                if m:
                    return m.group(0)
            return found
        line = proc.stdout.readline() if proc.stdout else ""
        if not line:
            time.sleep(0.2)
            continue
        buf += line
        LOG.info("cloudflared: %s", line.strip())
        if url_from_logs:
            m = URL_RE.search(line) or URL_RE.search(buf)
            if m:
                found = m.group(0)
                return found
    if url_from_logs:
        LOG.warning("no apareció URL trycloudflare en el tiempo de espera")
    return found


def _warn_origin_mismatch(spec: NamedTunnelSpec, port: int) -> None:
    if spec.kind != "config" or spec.config_path is None or not spec.config_path.is_file():
        return
    text = spec.config_path.read_text(encoding="utf-8", errors="replace")
    service = parse_local_service(text)
    if not service:
        return
    m = LOCAL_ORIGIN_RE.search(service)
    if m and int(m.group(1)) != port:
        LOG.warning(
            "config.yml service=%s no coincide con --port %s. "
            "El túnel va a pegarle al origen del YAML, no al puerto del worker.",
            service,
            port,
        )


def _start_named_tunnel(
    exe: str,
    spec: NamedTunnelSpec,
    *,
    port: int,
    token: str,
    startup_wait: float,
) -> subprocess.Popen[str] | None:
    if spec.error and spec.kind == "config" and (spec.config_path is None or not spec.config_path.is_file()):
        LOG.error("%s", spec.error)
        write_bridge_info(
            url=spec.public_url or "(túnel con nombre: config ausente)",
            token=token,
            port=port,
            stable=True,
        )
        return None
    if spec.error:
        LOG.warning("%s", spec.error)
    if spec.kind == "config" and spec.config_path is None:
        LOG.error("túnel con nombre sin config.yml")
        return None
    if spec.kind == "token" and not spec.token:
        LOG.error("túnel con nombre sin token")
        return None

    _warn_origin_mismatch(spec, port)
    cmd = named_tunnel_command(exe, spec)
    LOG.info("starting named tunnel: %s", cmd_for_log(cmd))
    proc = _spawn(cmd)
    _read_until(proc=proc, deadline=time.time() + max(startup_wait, 0), url_from_logs=False)
    url = spec.public_url or "(túnel con nombre: definí AFIP_TUNNEL_HOSTNAME)"
    write_bridge_info(url=url, token=token, port=port, stable=True)
    if proc.poll() is not None:
        return None
    LOG.info("TUNNEL_URL %s (estable)", url)
    _drain_stdout(proc)
    return proc


def _start_quick_tunnel(
    exe: str,
    *,
    port: int,
    token: str,
    url_wait: float,
) -> subprocess.Popen[str] | None:
    cmd = quick_tunnel_command(exe, port)
    LOG.info("starting tunnel: %s", cmd_for_log(cmd))
    proc = _spawn(cmd)
    url = _read_until(proc=proc, deadline=time.time() + url_wait, url_from_logs=True)
    if proc.poll() is not None and url is None:
        return None
    if url:
        write_bridge_info(url=url, token=token, port=port, stable=False)
        LOG.info("TUNNEL_URL %s", url)
    else:
        write_bridge_info(url="(túnel sin URL todavía)", token=token, port=port, stable=False)
    if proc.poll() is None:
        _drain_stdout(proc)
        return proc
    return None


def start_cloudflared_tunnel(
    *,
    port: int,
    token: str,
    named_startup_wait: float = 2.0,
    quick_url_wait: float = 45.0,
) -> subprocess.Popen[str] | None:
    exe = find_cloudflared()
    named = detect_named_tunnel()
    if not exe:
        LOG.warning(
            "cloudflared no está instalado. La web en la nube no va a llegar a esta PC. "
            "Instalá: winget install --id Cloudflare.cloudflared"
        )
        if named and named.public_url:
            url = named.public_url
            stable = True
        else:
            url = "http://127.0.0.1:%s (solo local)" % port
            stable = False
        write_bridge_info(url=url, token=token, port=port, stable=stable)
        return None
    if named is not None:
        return _start_named_tunnel(
            exe,
            named,
            port=port,
            token=token,
            startup_wait=named_startup_wait,
        )
    return _start_quick_tunnel(exe, port=port, token=token, url_wait=quick_url_wait)
