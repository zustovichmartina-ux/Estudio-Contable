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
import threading
import time
import urllib.request
from dataclasses import dataclass
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
DISCOVERY_URLS = (
    GITHUB_RAW_TUNNEL_URL,
    f"https://cdn.jsdelivr.net/gh/{GITHUB_REPO}@master/{TUNNEL_URL_REL}",
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


def public_health(url: str) -> bool:
    text = (url or "").strip().rstrip("/")
    if not text.startswith("https://"):
        return False
    req = urllib.request.Request(
        f"{text}/health",
        headers={"User-Agent": "EstudioContable-ARCA/1.0", "Accept": "application/json"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            return int(getattr(resp, "status", 200) or 200) == 200
    except Exception:
        return False


def _publish_via_git(url: str) -> None:
    git = shutil.which("git")
    if not git:
        LOG.warning("git no está: la web no ve la URL hasta un push")
        return
    root = _repo_root()
    rel = TUNNEL_URL_REL.replace("\\", "/")
    try:
        subprocess.run(
            [git, "add", "--", rel],
            cwd=root,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
        diff = subprocess.run(
            [git, "diff", "--cached", "--quiet", "--", rel],
            cwd=root,
            capture_output=True,
            timeout=20,
        )
        if diff.returncode == 0:
            subprocess.run(
                [git, "push", "origin", "HEAD"],
                cwd=root,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=40,
            )
            return
        commit = subprocess.run(
            [git, "commit", "-m", "Refresh ARCA tunnel URL", "--", rel],
            cwd=root,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
        if commit.returncode != 0:
            LOG.warning("no pude commitear URL: %s", (commit.stderr or commit.stdout)[:180])
            return
        push = subprocess.run(
            [git, "push", "origin", "HEAD"],
            cwd=root,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=40,
        )
        if push.returncode == 0:
            LOG.info("URL de túnel pusheada para la web")
        else:
            LOG.warning("no pude pushear URL: %s", (push.stderr or push.stdout)[:180])
    except Exception as exc:
        LOG.warning("no pude publicar URL por git: %s", type(exc).__name__)


def publish_public_url(url: str) -> None:
    """Deja la URL del túnel en git para que Streamlit Cloud la lea sin tocar Secrets."""
    url = (url or "").strip()
    if not url.startswith("https://"):
        return
    save_public_url(url)
    gh = shutil.which("gh")
    if not gh:
        _publish_via_git(url)
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
        _publish_via_git(url)
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
            return
        LOG.warning("no pude publicar URL: %s", (put.stderr or put.stdout)[:180])
    except Exception as exc:
        LOG.warning("no pude publicar URL: %s", type(exc).__name__)
    _publish_via_git(url)


def _drain_stdout(proc: subprocess.Popen[str]) -> None:
    try:
        if not proc.stdout:
            return
        for line in proc.stdout:
            text = (line or "").strip()
            if text:
                LOG.info("cloudflared: %s", text)
    except Exception:
        return


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
            threading.Thread(target=_drain_stdout, args=(proc,), daemon=True).start()
            return proc
    LOG.warning("no apareció URL trycloudflare en 45s")
    write_bridge_info(url="(túnel sin URL todavía)", token=token, port=port)
    threading.Thread(target=_drain_stdout, args=(proc,), daemon=True).start()
    return proc


@dataclass
class TunnelState:
    proc: subprocess.Popen[str] | None = None
    url: str = ""
    port: int = 8765
    token: str = ""
    next_check: float = 0.0
    failures: int = 0


def _read_saved_url() -> str:
    try:
        raw = public_url_path().read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    parts = raw.strip().split()
    if not parts:
        return ""
    text = parts[0].strip().strip('"').strip("'")
    if text.startswith("https://"):
        return text.rstrip("/")
    return ""


def maintain_tunnel(state: TunnelState) -> None:
    """Si cloudflared se colgó o el hostname de trycloudflare murió, relanza el túnel."""
    now = time.time()
    if now < state.next_check:
        return
    state.next_check = now + 45.0
    proc_dead = state.proc is None or state.proc.poll() is not None
    url = state.url or _read_saved_url()
    url_ok = (not proc_dead) and public_health(url)
    if url_ok:
        state.url = url
        state.failures = 0
        return
    state.failures += 1
    LOG.warning(
        "túnel ARCA caído (proc_dead=%s url_ok=%s fails=%s); relanzo cloudflared",
        proc_dead,
        url_ok,
        state.failures,
    )
    if state.proc and state.proc.poll() is None:
        try:
            state.proc.kill()
        except Exception:
            pass
        try:
            state.proc.wait(timeout=5)
        except Exception:
            pass
    state.proc = start_cloudflared_tunnel(port=state.port, token=state.token)
    state.url = _read_saved_url()
    state.next_check = time.time() + 60.0
