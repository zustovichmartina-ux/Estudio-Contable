# -*- coding: utf-8 -*-
"""Túnel Cloudflare (trycloudflare) para exponer la API local a Streamlit Cloud."""
from __future__ import annotations

import base64
import json
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
TUNNEL_URL_REL = "afip_worker/current_tunnel_url.txt"
GITHUB_REPO = "zustovichmartina-ux/Estudio-Contable"
GITHUB_RAW_TUNNEL_URL = (
    f"https://raw.githubusercontent.com/{GITHUB_REPO}/master/{TUNNEL_URL_REL}"
)


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
                "La web busca la URL sola (afip_worker/current_tunnel_url.txt).",
                "En Secrets alcanza con AFIP_WORKER_TOKEN; la URL se actualiza al arrancar el worker.",
                "",
            ]
        ),
        encoding="utf-8",
    )
    return path


def public_url_path() -> Path:
    return _repo_root() / TUNNEL_URL_REL


def save_public_url(url: str) -> Path:
    path = public_url_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(url.strip() + "\n", encoding="utf-8")
    return path


def publish_public_url(url: str) -> None:
    """Deja la URL del túnel en git para que Streamlit Cloud la lea sin tocar Secrets."""
    url = (url or "").strip()
    if not url.startswith("https://"):
        return
    save_public_url(url)
    gh = shutil.which("gh")
    if not gh:
        LOG.warning("gh no está: la URL queda local hasta el próximo push")
        return
    content_b64 = base64.b64encode((url + "\n").encode("utf-8")).decode("ascii")
    api = f"repos/{GITHUB_REPO}/contents/{TUNNEL_URL_REL}"
    sha = ""
    try:
        got = subprocess.run(
            [gh, "api", api],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
        if got.returncode == 0:
            sha = str((json.loads(got.stdout) or {}).get("sha") or "")
    except Exception as exc:
        LOG.warning("no pude leer URL publicada: %s", type(exc).__name__)
        return
    payload: dict[str, str] = {
        "message": "Refresh ARCA tunnel URL",
        "content": content_b64,
    }
    if sha:
        payload["sha"] = sha
    try:
        put = subprocess.run(
            [gh, "api", "-X", "PUT", api, "--input", "-"],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
        if put.returncode == 0:
            LOG.info("URL de túnel publicada para la web")
        else:
            LOG.warning("no pude publicar URL: %s", (put.stderr or put.stdout)[:180])
    except Exception as exc:
        LOG.warning("no pude publicar URL: %s", type(exc).__name__)


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
            publish_public_url(url)
            LOG.info("TUNNEL_URL %s", url)
            return proc
    LOG.warning("no apareció URL trycloudflare en 45s")
    write_bridge_info(url="(túnel sin URL todavía)", token=token, port=port)
    return proc
