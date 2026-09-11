# -*- coding: utf-8 -*-
"""Resultado de una acción del worker (dry-run o live)."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ActionResult:
    ok: bool = False
    message: str = ""
    files: list[str] = field(default_factory=list)
    needs_auth: bool = False
    extra: dict = field(default_factory=dict)
