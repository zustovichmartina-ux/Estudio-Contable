# -*- coding: utf-8 -*-
"""Chat del agente Tango: responde, formula y lee capturas."""
from __future__ import annotations

import os

import streamlit as st

from tango_bot import INDEX_PATH, preparar_imagen, reconstruir_indice, responder

_SUGERIDAS = [
    ("Fórmula sueldo básico", "Armá y explicá la fórmula del concepto 1 Sueldo básico en Tango Sueldos"),
    ("Asiento IVA al cargar factura", "¿Cómo se genera el asiento automático de IVA al cargar una factura?"),
    ("DETIVA mensual", "Diferencia entre asiento por comprobante y determinación mensual DETIVA"),
    ("IVA como VARIOS", "Por qué el estudio exporta IVA como VARIOS y no como tipo IVA"),
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
        val = st.secrets.get(nombre)
        if val:
            texto = str(val).strip().strip('"').strip("'")
            if texto and texto not in {"xai-...", "xai-", "..."}:
                return texto
        ofi = st.secrets.get("oficina_usuarios")
        if ofi:
            bloques = ofi.values() if hasattr(ofi, "values") else []
            for bloque in bloques:
                try:
                    extra = bloque.get(nombre) if bloque is not None else None
                except Exception:
                    extra = None
                if extra:
                    texto = str(extra).strip().strip('"').strip("'")
                    if texto and texto not in {"xai-...", "xai-", "..."}:
                        return texto
    except Exception:
        return ""
    return ""


def _leer_chat_input(raw: object) -> tuple[str, list]:
    if raw is None:
        return "", []
    if isinstance(raw, str):
        return raw.strip(), []
    texto = str(getattr(raw, "text", "") or "").strip()
    archivos = list(getattr(raw, "files", None) or [])
    return texto, archivos


def render_tango_bot() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)

    if "tango_chat" not in st.session_state:
        st.session_state.tango_chat = []

    if "tango_xai_key" not in st.session_state:
        st.session_state.tango_xai_key = ""
    api_key = _secret("XAI_API_KEY") or str(st.session_state.get("tango_xai_key") or "").strip()
    model = _secret("XAI_MODEL") or "grok-4.6"
    for nombre in ("OPENAI_API_KEY", "OPENAI_MODEL", "GROQ_API_KEY", "GROQ_MODEL", "XAI_API_KEY", "XAI_MODEL"):
        val = _secret(nombre)
        if val:
            os.environ[nombre] = val
    if api_key:
        os.environ["XAI_API_KEY"] = api_key
        os.environ.setdefault("XAI_MODEL", model)
    hay_ia = bool(api_key)
    chat = st.session_state.tango_chat

    def _panel_grok() -> None:
        if hay_ia:
            st.caption(f"Grok activo (`{model}`). Razona, formula y lee capturas.")
            return
        st.caption("Sin Grok el chat solo pega ayudas. Pegá la clave xAI para activarlo.")
        clave = st.text_input(
            "Clave Grok (xAI)",
            type="password",
            key="tango_xai_input",
            placeholder="xai-...",
            help="Creala en https://console.x.ai → API keys. Para que quede fija: Manage app → Secrets → XAI_API_KEY.",
        )
        if st.button("Activar Grok", key="tango_activar_grok"):
            if clave.strip().startswith("xai-"):
                st.session_state.tango_xai_key = clave.strip()
                st.rerun()
            else:
                st.warning("La clave tiene que empezar con xai- (console.x.ai).")

    if chat:
        top_l, top_r = st.columns([4, 1])
        with top_l:
            with st.expander("Ajustes", expanded=not hay_ia):
                st.caption("Agente Tango con Grok. Usa las ayudas y el export de sueldos del estudio.")
                _panel_grok()
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
            '<div class="tango-hero"><h1>Agente Tango</h1>'
            "<p>Grok responde, formula y lee capturas de Tango.</p></div>",
            unsafe_allow_html=True,
        )
        if not hay_ia:
            _panel_grok()
        for i, (label, pregunta) in enumerate(_SUGERIDAS):
            if st.button(label, key=f"tango_sug_{i}", use_container_width=True):
                with st.spinner("El agente está pensando…"):
                    _enviar(pregunta, api_key, model, [])
                st.rerun()
    else:
        for msg in chat:
            with st.chat_message(msg["role"]):
                for data_url in msg.get("imagenes") or []:
                    st.image(data_url, width=420)
                st.markdown(msg["content"])

    try:
        raw = st.chat_input(
            "Escribí o adjuntá una captura de Tango…",
            accept_file=True,
            file_type=["png", "jpg", "jpeg", "webp"],
        )
    except TypeError:
        raw = st.chat_input("Escribí tu consulta de Tango…")
    texto, archivos = _leer_chat_input(raw)
    if texto or archivos:
        imagenes = []
        for f in archivos[:4]:
            imagenes.append(preparar_imagen(f.getvalue(), getattr(f, "name", "captura.png")))
        with st.spinner("El agente está leyendo y formulando…"):
            _enviar(texto, api_key, model, imagenes)
        st.rerun()


def _enviar(prompt: str, api_key: str, model: str, imagenes: list[dict[str, str]]) -> None:
    st.session_state.tango_chat.append({
        "role": "user",
        "content": prompt or "(captura de Tango)",
        "imagenes": [img["data_url"] for img in imagenes],
    })
    historial = [
        {"role": m["role"], "content": m["content"]}
        for m in st.session_state.tango_chat
        if m["role"] in {"user", "assistant"}
    ][:-1]
    if not INDEX_PATH.exists():
        reconstruir_indice(guardar=True)
    out = responder(
        prompt,
        historial,
        api_key=api_key,
        model=model,
        imagenes=imagenes,
    )
    st.session_state.tango_chat.append({
        "role": "assistant",
        "content": out["texto"],
        "fuentes": out.get("fuentes") or [],
    })
