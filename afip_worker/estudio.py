# -*- coding: utf-8 -*-
"""CUIT con el que RECEPCION entra a AFIP (clave del estudio, no la del cliente)."""
from __future__ import annotations

import os
import re


def cuit_login_recepcion() -> str:
    """11 dígitos. Override: AFIP_LOGIN_CUIT. El job.cuit es el representado, no el login."""
    raw = (os.environ.get("AFIP_LOGIN_CUIT") or "23422843434").strip()
    digits = re.sub(r"\D", "", raw)
    return digits if len(digits) == 11 else "23422843434"
