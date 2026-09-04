# -*- coding: utf-8 -*-
"""Chat Tango tipo Grok dentro de la app Streamlit."""
from __future__ import annotations

import streamlit as st

from tango_bot import INDEX_PATH, reconstruir_indice, responder

_SUGERIDAS = [
    ("Asiento IVA al cargar una factura", "¿Cómo se genera el asiento automático de IVA al cargar una factura?"),
    ("Determinación mensual DETIVA", "Diferencia entre asiento por comprobante y determinación mensual DETIVA"),
    ("Por qué IVA se exporta como VARIOS", "Por qué el estudio exporta IVA como VARIOS y no como tipo IVA"),
    ("Fórmula de sueldo básico", "Fórmula de sueldo básico concepto 1 en Tango Sueldos"),
    ("Cuentas con auxiliares", "Qué pasa si una cuenta usa auxiliares al importar Excel a Tango"),
]

_CSS = """
<style>
div[data-testid="stChatMessage"] {
    background: #fff;
    border: 1px solid #E6E4EA;
    border-radius: 16px;
    padding: 0.35rem 0.2rem;
    margin-bottom: 0.55rem;
}
.tango-hero {
    text-align: center;
    padding: 3.2rem 0.75rem 1.4rem;
}
.tango-hero h1 {
    font-size: 1.85rem;
    font-weight: 700;
    color: #1F4E79;
    letter-spacing: -0.03em;
    margin: 0 0 0.45rem 0;
}
.tango-hero p {
    color: #6B6B75;
    font-size: 1.02rem;
    margin: 0;
}
</style>
"""


def _secret(nombre: str) -> str:
    try:
        return str(st.secrets.get(nombre) or "").strip().strip('"').strip("'")
    except Exception:
        return ""


def render_tango_bot() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)

    if "tango_chat" not in st.session_state:
        st.session_state.tango_chat = []

    api_key = _secret("XAI_API_KEY")
    model = _secret("XAI_MODEL") or "grok-4"
    chat = st.session_state.tango_chat

    if chat:
        top_l, top_r = st.columns([4, 1])
        with top_l:
            with st.expander("Ajustes", expanded=False):
                st.caption("Responde con las ayudas Tango del escritorio y normativas, más las reglas de esta app.")
                if not api_key:
                    st.caption("Opcional: `XAI_API_KEY` en Secrets para redacción tipo Grok.")
                else:
                    st.caption(f"Grok activo (`{model}`).")
                if st.button("Reindexar ayudas Tango", key="tango_reindex"):
                    with st.spinner("Leyendo Desktop\\Tango y normativas…"):
                        docs = reconstruir_indice(guardar=True)
                    st.success(f"Listo: {len(docs)} fragmentos")
        with top_r:
            if st.button("Nueva charla", key="tango_clear", use_container_width=True):
                st.session_state.tango_chat = []
                st.rerun()

    if not chat:
        st.markdown(
            '<div class="tango-hero"><h1>¿En qué te ayudo con Tango?</h1>'
            "<p>Preguntame por menús, asientos, IVA, sueldos o la exportación del estudio.</p></div>",
            unsafe_allow_html=True,
        )
        for i, (label, pregunta) in enumerate(_SUGERIDAS):
            if st.button(label, key=f"tango_sug_{i}", use_container_width=True):
                with st.spinner("Pensando…"):
                    _enviar(pregunta, api_key, model)
                st.rerun()
    else:
        for msg in chat:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    prompt = st.chat_input("Escribí tu duda de Tango…")
    if prompt:
        with st.spinner("Pensando…"):
            _enviar(prompt, api_key, model)
        st.rerun()


def _enviar(prompt: str, api_key: str, model: str) -> None:
    st.session_state.tango_chat.append({"role": "user", "content": prompt})
    historial = [
        {"role": m["role"], "content": m["content"]}
        for m in st.session_state.tango_chat
        if m["role"] in {"user", "assistant"}
    ][:-1]
    if not INDEX_PATH.exists():
        reconstruir_indice(guardar=True)
    out = responder(prompt, historial, api_key=api_key, model=model)
    st.session_state.tango_chat.append({
        "role": "assistant",
        "content": out["texto"],
        "fuentes": out.get("fuentes") or [],
    })
