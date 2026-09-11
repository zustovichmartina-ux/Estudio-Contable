# -*- coding: utf-8 -*-
"""Login AFIP visible. Espera hasta que el usuario entre al portal. Sin apuro."""
from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

os.environ["AFIP_CHROME_HEADED"] = "1"

from afip_worker.actions.comprobantes import LOGIN_URL, _is_login, _is_portal
from afip_worker.browser import chrome_context, close_browser, kill_profile_chrome, show_worker_chrome
from afip_worker.jobs import acquire_login_lock, release_login_lock

LOG = logging.getLogger("afip_worker.login_once")

CUIT_LOGIN = "23422843434"


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    dest = Path.cwd() / "jobs" / "chrome_afip" / "downloads"
    dest.mkdir(parents=True, exist_ok=True)
    acquire_login_lock()
    try:
        n = kill_profile_chrome()
        print()
        print("========================================")
        print("  PASO 1 — Entrar a ARCA (una vez)")
        print("========================================")
        if n:
            print(f"Cerré el Chrome oculto del worker ({n}).")
        print("Se abre Chrome ADELANTE. No lo minimices.")
        print("1) Revisá que el CUIT sea el de Martina")
        print("2) Escribí la clave fiscal")
        print("3) Si Chrome pregunta, tocá Guardar contraseña")
        print("4) Completá el 2FA si aparece")
        print("Cuando veas el portal (Mis servicios), esta ventana dice OK sola.")
        print("No hay tiempo límite. Ctrl+C si querés cancelar.")
        print()

        with chrome_context(download_dir=dest) as (_ctx, page, _att):
            show_worker_chrome()
            page.goto(LOGIN_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(800)
            user = page.locator("#F1\\:username")
            if user.count() == 0:
                user = page.locator("input[name='F1:username'], input[type='text']").first
            try:
                if user.count():
                    user.first.click(force=True)
                    user.first.fill(CUIT_LOGIN)
            except Exception:
                print("No pude completar el CUIT: escribilo vos en Chrome.")
            show_worker_chrome()

            waited = 0
            while True:
                if _is_portal(page) and not _is_login(page):
                    print()
                    print("OK — sesión lista. Ya podés cerrar esta ventana.")
                    print("Después: pedí la descarga en la web ARCA (o avisame).")
                    close_browser()
                    return 0
                time.sleep(2.0)
                waited += 2
                if waited % 20 == 0:
                    print(f"Sigo esperando el portal... ({waited}s) — entrá en Chrome.")
    except KeyboardInterrupt:
        print("\nCancelado.")
        return 1
    finally:
        release_login_lock()


if __name__ == "__main__":
    raise SystemExit(main())
