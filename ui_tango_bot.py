# -*- coding: utf-8 -*-
"""Chat del agente Tango: responde, formula y lee capturas."""
from __future__ import annotations

import os
from pathlib import Path

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
    border: 1px solid #B8C7D9;
    border-radius: 16px;
    padding: 0.45rem 0.35rem;
    margin-bottom: 0.55rem;
    box-shadow: 0 2px 8px rgba(31, 78, 121, 0.06);
}
.tango-hero {
    text-align: center;
    padding: 2.4rem 0.75rem 1.1rem;
}
.tango-hero h1 {
    font-size: 1.85rem;
    font-weight: 700;
    color: #1F4E79;
    letter-spacing: -0.03em;
    margin: 0 0 0.45rem 0;
}
.tango-hero p {
    color: #3D4F63;
    font-size: 1.05rem;
    margin: 0;
}
.tango-sug-label {
    color: #1F4E79;
    font-weight: 700;
    font-size: 0.95rem;
    letter-spacing: 0.02em;
    margin: 0.35rem 0 0.65rem 0;
    text-align: center;
}
[class*="st-key-tango_sug_box"],
div.stMarkdown:has(.tango-sug-label) + div [data-testid="stVerticalBlockBorderWrapper"] {
    background: #FFFFFF !important;
    border: 2px solid #1F4E79 !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 16px rgba(31, 78, 121, 0.12) !important;
}
[class*="st-key-tango_sug_box"] .stButton > button,
[class*="st-key-tango_sug_box"] button[kind="secondary"],
[class*="st-key-tango_sug_"] .stButton > button,
[class*="st-key-tango_sug_"] button[kind="secondary"] {
    background: #F4F8FC !important;
    color: #1F4E79 !important;
    border: 2px solid #1F4E79 !important;
    font-weight: 600 !important;
    font-size: 0.98rem !important;
    min-height: 3.05rem !important;
    box-shadow: 0 2px 10px rgba(31, 78, 121, 0.12) !important;
}
[class*="st-key-tango_sug_box"] .stButton > button:hover,
[class*="st-key-tango_sug_box"] button[kind="secondary"]:hover,
[class*="st-key-tango_sug_"] .stButton > button:hover,
[class*="st-key-tango_sug_"] button[kind="secondary"]:hover {
    background: #1F4E79 !important;
    color: #FFFFFF !important;
    opacity: 1 !important;
}
[data-testid="stBottomBlockContainer"] {
    background: linear-gradient(to top, #E8EEF5 70%, rgba(247, 249, 252, 0)) !important;
    padding-bottom: 1.15rem !important;
}
[data-testid="stChatInput"] {
    background: #FFFFFF !important;
    border: 2px solid #1F4E79 !important;
    border-radius: 18px !important;
    box-shadow: 0 6px 22px rgba(31, 78, 121, 0.18) !important;
}
[data-testid="stChatInput"] > div,
[data-testid="stChatInput"] .stChatInputContainer {
    background: #FFFFFF !important;
    border: none !important;
    box-shadow: none !important;
    border-radius: 16px !important;
}
[data-testid="stChatInput"] textarea,
[data-testid="stChatInput"] [data-testid="stChatInputTextArea"],
[data-testid="stChatInput"] input,
[data-testid="stChatInput"] [data-baseweb="textarea"] textarea {
    color: #1A1A1A !important;
    font-size: 1.05rem !important;
    caret-color: #1F4E79 !important;
}
[data-testid="stChatInput"] textarea::placeholder,
[data-testid="stChatInput"] [data-baseweb="textarea"] textarea::placeholder {
    color: #4A5D73 !important;
    opacity: 1 !important;
}
[data-testid="stChatInput"] button {
    background: #1F4E79 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 10px !important;
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
    return _secret("XAI_MODEL") or "grok-4.6"


def _etiqueta_ia(api_key: str, model: str) -> str:
    if api_key.startswith("xai-") or model.startswith("grok"):
        return "Grok"
    if api_key.startswith("sk-ant-") or "claude" in model.lower():
        return "Claude"
    if api_key.startswith("gsk_"):
        return "Groq"
    if api_key.startswith("sk-"):
        return "OpenAI"
    return "IA"


def _es_web_publica() -> bool:
    """True en Streamlit Cloud. Ajustes/clave solo se muestran en la PC local."""
    flags = (
        os.environ.get("STREAMLIT_SHARING_MODE"),
        os.environ.get("STREAMLIT_CLOUD"),
        os.environ.get("IS_STREAMLIT_CLOUD"),
    )
    if any(str(f).strip().lower() in {"1", "true", "yes"} for f in flags if f):
        return True
    if Path("/mount/src").is_dir() or Path("/home/appuser").is_dir():
        return True
    host = str(os.environ.get("HOSTNAME") or os.environ.get("COMPUTERNAME") or "").lower()
    return host.endswith(".streamlit.app") or "streamlit" in host


def _leer_chat_input(raw: object) -> tuple[str, list]:
    if raw is None:
        return "", []
    if isinstance(raw, str):
        return raw.strip(), []
    texto = str(getattr(raw, "text", "") or "").strip()
    archivos = getattr(raw, "files", None) or []
    if archivos and not isinstance(archivos, (list, tuple)):
        archivos = [archivos]
    else:
        archivos = list(archivos)
    return texto, archivos


def render_tango_bot() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)

    if "tango_chat" not in st.session_state:
        st.session_state.tango_chat = []

    if "tango_api_key" not in st.session_state:
        st.session_state.tango_api_key = str(st.session_state.get("tango_xai_key") or "")
    web = _es_web_publica()
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
    # Clave pegada en Tango (esta sesión) pisa Secrets, también en la web.
    pasted = str(st.session_state.get("tango_api_key") or "").strip()
    if pasted in _PLACEHOLDERS:
        pasted = ""
    api_key = (
        pasted
        or _secret("XAI_API_KEY")
        or _secret("ANTHROPIC_API_KEY")
        or _secret("OPENAI_API_KEY")
        or _secret("GROQ_API_KEY")
    )
    if api_key.startswith("xai-"):
        os.environ["XAI_API_KEY"] = api_key
    elif api_key.startswith("sk-ant-"):
        os.environ["ANTHROPIC_API_KEY"] = api_key
    model = _modelo_para_clave(api_key)
    hay_ia = bool(api_key)
    etiqueta = _etiqueta_ia(api_key, model)
    chat = st.session_state.tango_chat

    def _panel_ia() -> None:
        if hay_ia:
            st.success(f"**Grok activo** (`{model}`). Ya podés preguntar como en el chat de Grok.")
        else:
            st.error("**Grok está apagado.** Sin clave de xAI este chat no tiene IA: solo arma respuestas de archivo.")
        st.caption("Pegá acá la clave de Grok (console.x.ai). No la escribas en el chat de abajo. Queda en tu sesión.")
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
            with st.expander("Grok", expanded=not hay_ia):
                _panel_ia()
                if not web:
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
            "<p>Grok formula la respuesta con lo del estudio.</p></div>",
            unsafe_allow_html=True,
        )
        _panel_ia()
        st.markdown(
            '<p class="tango-sug-label">Consultas frecuentes</p>',
            unsafe_allow_html=True,
        )
        try:
            sug_box = st.container(border=True, key="tango_sug_box")
        except TypeError:
            sug_box = st.container(border=True)
        with sug_box:
            for i, (label, pregunta) in enumerate(_SUGERIDAS):
                if st.button(label, key=f"tango_sug_{i}", use_container_width=True):
                    with st.spinner("El agente está pensando…"):
                        _enviar(pregunta, api_key, model, [])
                    st.rerun()
    else:
        for msg in chat:
            with st.chat_message(msg["role"]):
                for data_url in msg.get("imagenes") or []:
                    st.image(data_url, width=480, caption="Captura de Tango")
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
