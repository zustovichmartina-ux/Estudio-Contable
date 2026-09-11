# -*- coding: utf-8 -*-
"""Chrome oculto con autofill. Si vence la sesión, reingresa solo (clave ya guardada)."""
from __future__ import annotations

import atexit
import ctypes
import logging
import os
import subprocess
import time
from collections.abc import Iterator
from contextlib import contextmanager
from ctypes import wintypes
from pathlib import Path

from playwright.sync_api import BrowserContext, Page, Playwright, sync_playwright

from .jobs import jobs_root

LOG = logging.getLogger("afip_worker.browser")

_CHROME_ARGS = [
    "--disable-blink-features=AutomationControlled",
    "--no-first-run",
    "--no-default-browser-check",
]

_pw: Playwright | None = None
_context: BrowserContext | None = None
_attached = False
_enum_cb = None

SW_HIDE = 0
SW_SHOW = 5
SW_RESTORE = 9
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000


def _env_flag(name: str, default: bool) -> bool:
    raw = (os.environ.get(name) or "").strip().lower()
    if not raw:
        return default
    return raw in {"1", "true", "yes", "si", "sí"}


def _visible_login() -> bool:
    return _env_flag("AFIP_CHROME_HEADED", False)


def _cdp_url() -> str:
    return (os.environ.get("AFIP_CHROME_CDP") or "").strip()


def _profile_dir() -> Path:
    raw = (os.environ.get("AFIP_CHROME_USER_DATA") or "").strip()
    if raw:
        return Path(raw)
    return jobs_root() / "chrome_afip"


def _downloads_dir() -> Path:
    d = _profile_dir() / "downloads"
    d.mkdir(parents=True, exist_ok=True)
    return d


def login_timeout_ms() -> int:
    raw = (os.environ.get("AFIP_LOGIN_TIMEOUT_S") or "").strip()
    if raw:
        try:
            return max(10, int(raw)) * 1000
        except ValueError:
            pass
    # Login visible (2FA): 3 min. Autofill oculto: 45 s.
    return 180_000 if _visible_login() else 45_000


def _context_alive(ctx: BrowserContext) -> bool:
    try:
        _ = ctx.pages
        return True
    except Exception:
        return False


def _worker_chrome_pids() -> set[int]:
    marker = str(_profile_dir()).replace("/", "\\").lower()
    cmd = (
        "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
        "ForEach-Object { '{0}\t{1}' -f $_.ProcessId, $_.CommandLine }"
    )
    try:
        raw = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", cmd],
            text=True,
            timeout=12,
            errors="replace",
        )
    except Exception:
        return set()
    pids: set[int] = set()
    for line in raw.splitlines():
        if "\t" not in line:
            continue
        pid_s, cmdl = line.split("\t", 1)
        blob = (cmdl or "").lower()
        if marker.lower() not in blob and "chrome_afip" not in blob:
            continue
        try:
            pids.add(int(pid_s.strip()))
        except ValueError:
            continue
    return pids


def hide_worker_chrome() -> int:
    """Oculta el Chrome del worker (sin ícono Test en la barra). Autofill sigue andando."""
    if _visible_login():
        return 0
    pids = _worker_chrome_pids()
    if not pids:
        return 0
    user32 = ctypes.windll.user32
    hidden = 0

    def _enum(hwnd: int, _lp: int) -> bool:
        nonlocal hidden
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        try:
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE, (style | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW
            )
            user32.ShowWindow(hwnd, SW_HIDE)
            hidden += 1
        except Exception:
            pass
        return True

    wnd_enum = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    global _enum_cb
    _enum_cb = wnd_enum(_enum)
    user32.EnumWindows(_enum_cb, 0)
    return hidden


def show_worker_chrome() -> int:
    """Trae al frente el Chrome del worker (login visible)."""
    pids = _worker_chrome_pids()
    if not pids:
        return 0
    user32 = ctypes.windll.user32
    shown = 0

    def _enum(hwnd: int, _lp: int) -> bool:
        nonlocal shown
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        try:
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(
                hwnd, GWL_EXSTYLE, (style | WS_EX_APPWINDOW) & ~WS_EX_TOOLWINDOW
            )
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            shown += 1
        except Exception:
            pass
        return True

    wnd_enum = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    global _enum_cb
    _enum_cb = wnd_enum(_enum)
    user32.EnumWindows(_enum_cb, 0)
    return shown


def kill_profile_chrome() -> int:
    """Cierra solo el Chrome de jobs/chrome_afip (no el Chrome de todos los días)."""
    pids = _worker_chrome_pids()
    killed = 0
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F", "/T"],
                capture_output=True,
                timeout=8,
                check=False,
            )
            killed += 1
        except Exception:
            continue
    if pids:
        time.sleep(1.2)
    profile = _profile_dir()
    for name in ("SingletonLock", "SingletonCookie", "SingletonSocket", "lockfile"):
        lock = profile / name
        try:
            if lock.exists() or lock.is_symlink():
                lock.unlink(missing_ok=True)
        except OSError:
            pass
    if killed:
        LOG.info("cerré Chrome oculto del worker (%s procesos) para poder abrir uno visible", killed)
    return killed


def _hide_for_a_bit() -> None:
    for _ in range(12):
        hide_worker_chrome()
        time.sleep(0.25)


def _launch_persistent(pw: Playwright, profile: Path, download_dir: Path) -> BrowserContext:
    args = list(_CHROME_ARGS)
    kwargs: dict = {
        "user_data_dir": str(profile),
        "headless": False,
        "accept_downloads": True,
        "downloads_path": str(download_dir),
        "locale": "es-AR",
        "timezone_id": "America/Argentina/Buenos_Aires",
        "ignore_default_args": ["--enable-automation"],
        "channel": "chrome",
        "args": args,
        "viewport": {"width": 1280, "height": 800},
    }
    if _visible_login():
        kill_profile_chrome()
        args.extend(["--window-position=80,80", "--window-size=1280,800"])
        kwargs["args"] = args
        LOG.info("Chrome visible (login/2FA una vez) profile=%s", profile)
    else:
        args.extend(["--start-minimized", "--window-position=-32000,-32000"])
        kwargs["args"] = args
        LOG.info("Chrome oculto + autofill profile=%s", profile)
    try:
        ctx = pw.chromium.launch_persistent_context(**kwargs)
    except Exception as exc:
        LOG.warning("channel=chrome falló (%s); Chromium", exc)
        kwargs.pop("channel", None)
        ctx = pw.chromium.launch_persistent_context(**kwargs)
    if _visible_login():
        time.sleep(0.4)
        show_worker_chrome()
    else:
        _hide_for_a_bit()
    return ctx


def ensure_browser(*, download_dir: Path | None = None) -> tuple[BrowserContext, bool]:
    global _pw, _context, _attached
    dest = download_dir or _downloads_dir()
    dest.mkdir(parents=True, exist_ok=True)

    if _context is not None and _context_alive(_context):
        hide_worker_chrome()
        return _context, _attached

    close_browser()
    _pw = sync_playwright().start()
    _attached = False

    cdp = _cdp_url()
    if cdp:
        try:
            browser = _pw.chromium.connect_over_cdp(cdp)
            _context = browser.contexts[0] if browser.contexts else browser.new_context(
                accept_downloads=True
            )
            _attached = True
            LOG.info("Chrome via CDP %s (no nueva ventana)", cdp)
            return _context, _attached
        except Exception as exc:
            LOG.warning("CDP no disponible (%s); perfil persistente oculto", exc)

    profile = _profile_dir()
    profile.mkdir(parents=True, exist_ok=True)
    _context = _launch_persistent(_pw, profile, dest)
    _attached = False
    return _context, _attached


def close_browser() -> None:
    global _pw, _context, _attached
    ctx, pw, attached = _context, _pw, _attached
    _context = None
    _pw = None
    _attached = False
    if ctx is not None and not attached:
        try:
            ctx.close()
        except Exception:
            pass
    if pw is not None:
        try:
            pw.stop()
        except Exception:
            pass


atexit.register(close_browser)


@contextmanager
def chrome_context(*, download_dir: Path) -> Iterator[tuple[BrowserContext, Page, bool]]:
    """Mismo Chrome oculto para todos los jobs: cookies + autofill de clave."""
    context, attached = ensure_browser(download_dir=download_dir)
    page = context.new_page()
    page.set_default_timeout(30_000)
    if _visible_login():
        show_worker_chrome()
    else:
        hide_worker_chrome()
    try:
        yield context, page, attached
    finally:
        try:
            page.close()
        except Exception:
            pass
        if _visible_login():
            show_worker_chrome()
        else:
            hide_worker_chrome()
