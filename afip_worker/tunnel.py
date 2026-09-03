# -*- coding: utf-8 -*-
"""Túnel Cloudflare (trycloudflare) para exponer la API local a Streamlit Cloud."""
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import time
from pathlib import Path

from .jobs import _repo_root
from .token import bridge_info_path

LOG = logging.getLogger("afip_worker.tunnel")
URL_RE = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com", re.I)


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


def write_bridge_info(*, url: str, token: str, port: int) -> Path:
    path = bridge_info_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "\n".join(
            [
                "Pegá esto en Streamlit Cloud → Manage app → Settings → Secrets:",
                "",
                f'AFIP_WORKER_URL = "{url}"',
                f'AFIP_WORKER_TOKEN = "{token}"',
                "",
                f"API local: http://127.0.0.1:{port}",
                "El túnel rápido cambia de URL si reiniciás el worker: actualizá Secrets.",
                "",
            ]
        ),
        encoding="utf-8",
    )
    return path


def start_cloudflared_tunnel(*, port: int, token: str) -> subprocess.Popen[str] | None:
    exe = find_cloudflared()
    if not exe:
        LOG.warning(
            "cloudflared no está instalado. La web en la nube no va a llegar a esta PC. "
            "Instalá: winget install --id Cloudflare.cloudflared"
        )
        write_bridge_info(url="http://127.0.0.1:%s (solo local)" % port, token=token, port=port)
        return None
    cmd = [exe, "tunnel", "--no-autoupdate", "--url", f"http://127.0.0.1:{port}"]
    LOG.info("starting tunnel: %s", " ".join(cmd))
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    deadline = time.time() + 45
    buf = ""
    while time.time() < deadline:
        if proc.poll() is not None:
            rest = proc.stdout.read() if proc.stdout else ""
            LOG.error("cloudflared salió: %s%s", buf, rest)
            return None
        line = proc.stdout.readline() if proc.stdout else ""
        if not line:
            time.sleep(0.2)
            continue
        buf += line
        LOG.info("cloudflared: %s", line.strip())
        m = URL_RE.search(line) or URL_RE.search(buf)
        if m:
            url = m.group(0)
            write_bridge_info(url=url, token=token, port=port)
            LOG.info("TUNNEL_URL %s", url)
            return proc
    LOG.warning("no apareció URL trycloudflare en 45s")
    write_bridge_info(url="(túnel sin URL todavía)", token=token, port=port)
    return proc
