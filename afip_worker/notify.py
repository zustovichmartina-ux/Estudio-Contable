# -*- coding: utf-8 -*-
"""Aviso en RECEPCION cuando una tarea ARCA termina. Sin claves."""
from __future__ import annotations

import logging
import subprocess

LOG = logging.getLogger("afip_worker.notify")


def aviso_tarea(titulo: str, cuerpo: str) -> None:
    """Globo de Windows. No bloquea la cola."""
    t = _limpio(titulo, 70)
    c = _limpio(cuerpo, 180)
    if not t:
        return
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "Add-Type -AssemblyName System.Drawing; "
        "$n = New-Object System.Windows.Forms.NotifyIcon; "
        "$n.Icon = [System.Drawing.SystemIcons]::Information; "
        "$n.Visible = $true; "
        f"$n.ShowBalloonTip(10000, '{t}', '{c}', "
        "[System.Windows.Forms.ToolTipIcon]::Info); "
        "Start-Sleep -Seconds 11; $n.Dispose()"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        LOG.warning("no pude avisar en Windows: %s", type(exc).__name__)


def _limpio(texto: str, max_len: int) -> str:
    t = (texto or "").replace("'", " ").replace('"', " ").replace("\n", " ")
    t = t.replace("$", " ").replace("`", " ").replace(";", " ")
    t = " ".join(t.split())
    return t[:max_len]
