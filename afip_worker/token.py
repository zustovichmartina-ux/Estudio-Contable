# -*- coding: utf-8 -*-
"""Token del puente nube → PC. Nunca va a git ni a Excel."""
from __future__ import annotations

import os
import secrets
from pathlib import Path

from .jobs import _repo_root, jobs_root


def token_path() -> Path:
    raw = (os.environ.get("AFIP_WORKER_TOKEN_FILE") or "").strip()
    if raw:
        return Path(raw)
    return jobs_root() / ".worker_token"


def load_or_create_token() -> str:
    env = (os.environ.get("AFIP_WORKER_TOKEN") or "").strip()
    if env:
        return env
    path = token_path()
    if path.exists():
        val = path.read_text(encoding="utf-8").strip()
        if val:
            return val
    token = secrets.token_urlsafe(32)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(token + "\n", encoding="utf-8")
    return token


def bridge_info_path() -> Path:
    return _repo_root() / "jobs" / "cloud_bridge.txt"
