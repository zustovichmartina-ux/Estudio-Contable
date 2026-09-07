# -*- coding: utf-8 -*-
"""Chat del agente Tango: responde, formula y lee capturas."""
from __future__ import annotations

import os

import streamlit as st

from tango_bot import INDEX_PATH, preparar_imagen, reconstruir_indice, responder

_SUGERIDAS = [
    ("Fórmula sueldo básico", "Formulá el concepto 1 Sueldo básico y dame FormulaImporte y FormulaCantidad listas para copiar y pegar en Tango"),
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


_PLACEHOLDERS = {
    "xai-...", "xai-", "...", "sk-ant-...", "sk-ant-", "sk-...", "PEGAR_CLAVE",
}


def _secret(nombre: str) -> str:
    try:
        val = st.secrets.get(nombre)
        if val:
            texto = str(val).strip().strip('"').strip("'")
            if texto and texto not in _PLACEHOLDERS:
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
                    if texto and texto not in _PLACEHOLDERS:
                        return texto
    except Exception:
        return ""
    return ""


def _modelo_para_clave(api_key: str) -> str:
    if api_key.startswith("sk-ant-"):
        return _secret("ANTHROPIC_MODEL") or "claude-sonnet-5"
    if api_key.startswith("xai-"):
        return _secret("XAI_MODEL") or "grok-4.6"
    if api_key.startswith("gsk_"):
        return _secret("GROQ_MODEL") or "llama-3.3-70b-versatile"
    if api_key.startswith("sk-"):
        return _secret("OPENAI_MODEL") or "gpt-4o-mini"
    return _secret("ANTHROPIC_MODEL") or "claude-sonnet-5"


def _etiqueta_ia(api_key: str, model: str) -> str:
    if api_key.startswith("sk-ant-") or "claude" in model.lower():
        return "Claude"
    if api_key.startswith("xai-") or model.startswith("grok"):
        return "Grok"
    if api_key.startswith("gsk_"):
        return "Groq"
    if api_key.startswith("sk-"):
        return "OpenAI"
    return "IA"


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

    if "tango_api_key" not in st.session_state:
        st.session_state.tango_api_key = str(st.session_state.get("tango_xai_key") or "")
    for nombre in (
        "ANTHROPIC_API_KEY",
        "ANTHROPIC_MODEL",
        "OPENAI_API_KEY",
        "OPENAI_MODEL",
        "GROQ_API_KEY",
        "GROQ_MODEL",
        "XAI_API_KEY",
        "XAI_MODEL",
    ):
        val = _secret(nombre)
        if val:
            os.environ[nombre] = val
    api_key = (
        _secret("ANTHROPIC_API_KEY")
        or _secret("XAI_API_KEY")
        or _secret("OPENAI_API_KEY")
        or _secret("GROQ_API_KEY")
        or str(st.session_state.get("tango_api_key") or "").strip()
    )
    model = _modelo_para_clave(api_key)
    hay_ia = bool(api_key)
    etiqueta = _etiqueta_ia(api_key, model)
    chat = st.session_state.tango_chat

    def _panel_ia() -> None:
        if hay_ia:
            st.success(f"{etiqueta} activo (`{model}`). Ya podés preguntar abajo.")
        st.info("Pegá acá la clave de Grok. No la pongas en el chat de abajo.")
        clave = st.text_input(
            "Clave Grok (xAI)",
            type="password",
            key="tango_api_input",
            placeholder="xai-...",
            help="Creala en https://console.x.ai → API keys. Empieza con xai-.",
        )
        if st.button("Activar Grok", type="primary", key="tango_activar_ia"):
            texto = clave.strip()
            if texto.startswith("xai-") or texto.startswith("sk-ant-") or texto.startswith("sk-"):
                st.session_state.tango_api_key = texto
                st.rerun()
            else:
                st.warning("La clave de Grok empieza con xai- (console.x.ai).")

    if chat:
        top_l, top_r = st.columns([4, 1])
        with top_l:
            with st.expander("Ajustes", expanded=not hay_ia):
                st.caption("Agente Tango con Grok o Claude. Usa las ayudas y el export de sueldos del estudio.")
                _panel_ia()
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
        _panel_ia()
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
