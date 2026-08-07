"""Reporte de errores hacia Cursor (archivo + UI Streamlit)."""

from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
ERROR_FILE = LOGS_DIR / "ultimo_error_cursor.txt"
FLAG_FILE = LOGS_DIR / "cursor_error_pendiente.flag"


def guardar_error_para_cursor(
    exc: BaseException | None = None,
    *,
    contexto: str = "",
    tb_text: str | None = None,
) -> Path:
    """Persiste el traceback para que el agente de Cursor lo lea."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    if tb_text is None:
        if exc is not None:
            tb_text = "".join(
                traceback.format_exception(type(exc), exc, exc.__traceback__)
            )
        else:
            tb_text = traceback.format_exc()

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cuerpo = (
        f"# Error para Cursor\n"
        f"Fecha: {stamp}\n"
        f"Contexto: {contexto or '(sin contexto)'}\n\n"
        f"```\n{tb_text.strip()}\n```\n"
    )
    ERROR_FILE.write_text(cuerpo, encoding="utf-8")
    FLAG_FILE.write_text(stamp, encoding="utf-8")
    return ERROR_FILE


def render_boton_enviar_error_cursor(
    exc: BaseException,
    *,
    contexto: str = "",
    key: str = "btn_enviar_error_cursor",
) -> None:
    """UI: en lugar de 'Preguntale a ChatGPT', envía el error al agente de Cursor."""
    st.error(f"**{type(exc).__name__}:** {exc}")
    with st.expander("Ver detalle técnico", expanded=False):
        st.code(
            "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)),
            language="text",
        )

    c1, c2 = st.columns([2, 1])
    with c1:
        if st.button(
            "📨 Enviarme error a Cursor",
            type="primary",
            key=key,
            help="Guarda el error en logs/ultimo_error_cursor.txt para que el agente de Cursor lo lea.",
        ):
            path = guardar_error_para_cursor(exc, contexto=contexto)
            st.session_state["_error_cursor_enviado"] = str(path)
            st.success(
                "Error enviado a Cursor. En el chat pedime: "
                "**leé el último error** (o abrí `logs/ultimo_error_cursor.txt`)."
            )
    with c2:
        if st.button("Reintentar", key=f"{key}_retry"):
            st.rerun()

    if st.session_state.get("_error_cursor_enviado"):
        st.caption(f"Último reporte: `{st.session_state['_error_cursor_enviado']}`")


def consumir_flag_error_cursor() -> str | None:
    """Si hay un error pendiente, lo lee y limpia el flag. Para uso del agente."""
    if not FLAG_FILE.exists() or not ERROR_FILE.exists():
        return None
    texto = ERROR_FILE.read_text(encoding="utf-8")
    try:
        FLAG_FILE.unlink(missing_ok=True)
    except OSError:
        pass
    return texto
