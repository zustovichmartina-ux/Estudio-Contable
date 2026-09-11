# -*- coding: utf-8 -*-
"""Chat del agente Tango: responde, formula y lee capturas."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import streamlit as st

from tango_bot import cargar_indice, preparar_imagen, reconstruir_indice, responder

_SUGERIDAS = [
    ("Fórmula sueldo básico", "Formulá el concepto 1 Sueldo básico y dame FormulaImporte y FormulaCantidad listas para copiar y pegar en Tango"),
    ("Asiento IVA al cargar factura", "¿Cómo se genera el asiento automático de IVA al cargar una factura?"),
    ("DETIVA mensual", "Diferencia entre asiento por comprobante y determinación mensual DETIVA"),
    ("IVA como VARIOS", "Por qué el estudio exporta IVA como VARIOS y no como tipo IVA"),
    ("Mis Retenciones", "En la DDJJ de IVA, ¿qué fecha e importe se toman de Mis Retenciones?"),
]

_CSS = """
<style>
div[data-testid="stChatMessage"] {
    background: var(--ec-card, #FFFFFF);
    border: none;
    border-radius: 16px;
    padding: 0.55rem 0.7rem;
    margin-bottom: 0.55rem;
    box-shadow: none;
    color: var(--ec-night, #0B0D10);
}
div[data-testid="stChatMessage"] p,
div[data-testid="stChatMessage"] li,
div[data-testid="stChatMessage"] span {
    color: var(--ec-night, #0B0D10) !important;
}
.tango-hero {
    text-align: left;
    padding: 0.15rem 0 0.55rem;
}
.tango-hero h1 {
    font-family: var(--ec-display, Outfit, sans-serif);
    font-size: 1.45rem;
    font-weight: 650;
    color: var(--ec-ink, #0F172A);
    letter-spacing: -0.03em;
    margin: 0 0 0.3rem 0;
    text-transform: none;
}
.tango-hero p {
    color: var(--ec-muted, #64748B);
    font-size: 0.95rem;
    margin: 0;
}
.tango-sug-label {
    color: var(--ec-muted, #64748B);
    font-weight: 600;
    font-size: 0.7rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    margin: 0.2rem 0 0.45rem 0;
    text-align: left;
}
[class*="st-key-tango_sug_box"],
div.stMarkdown:has(.tango-sug-label) + div [data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--ec-card, #FFFFFF) !important;
    border: none !important;
    border-radius: 16px !important;
    box-shadow: none !important;
}
[class*="st-key-tango_sug_box"] .stButton > button,
[class*="st-key-tango_sug_box"] button[kind="secondary"],
[class*="st-key-tango_sug_"] .stButton > button,
[class*="st-key-tango_sug_"] button[kind="secondary"] {
    background: transparent !important;
    color: var(--ec-night, #0B0D10) !important;
    border: 1px solid rgba(11, 13, 16, 0.16) !important;
    font-weight: 500 !important;
    font-size: 0.85rem !important;
    min-height: 2.35rem !important;
    border-radius: 10px !important;
    box-shadow: none !important;
}
[class*="st-key-tango_sug_box"] .stButton > button:hover,
[class*="st-key-tango_sug_box"] button[kind="secondary"]:hover,
[class*="st-key-tango_sug_"] .stButton > button:hover,
[class*="st-key-tango_sug_"] button[kind="secondary"]:hover {
    background: var(--ec-lagoon, #2563EB) !important;
    color: #FFFFFF !important;
    border-color: var(--ec-lagoon, #2563EB) !important;
    opacity: 1 !important;
}
[data-testid="stBottomBlockContainer"] {
    background: #F4F6FA !important;
    padding-bottom: 0.85rem !important;
}
[data-testid="stChatInput"] {
    background: #E8EAED !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 16px !important;
    box-shadow: none !important;
    color: #0F172A !important;
}
[data-testid="stChatInput"] > div,
[data-testid="stChatInput"] .stChatInputContainer,
[data-testid="stChatInput"] [data-baseweb="base-input"],
[data-testid="stChatInput"] [data-baseweb="textarea"] {
    background: #E8EAED !important;
    border: none !important;
    box-shadow: none !important;
    border-radius: 16px !important;
    color: #0F172A !important;
}
[data-testid="stChatInput"] textarea,
[data-testid="stChatInput"] [data-testid="stChatInputTextArea"],
[data-testid="stChatInput"] input,
[data-testid="stChatInput"] [data-baseweb="textarea"] textarea,
[data-testid="stChatInput"] [data-baseweb="input"] input {
    color: #0F172A !important;
    background: transparent !important;
    font-size: 0.95rem !important;
    caret-color: #2563EB !important;
    -webkit-text-fill-color: #0F172A !important;
}
[data-testid="stChatInput"] textarea::placeholder,
[data-testid="stChatInput"] [data-baseweb="textarea"] textarea::placeholder {
    color: #475569 !important;
    opacity: 1 !important;
    -webkit-text-fill-color: #475569 !important;
}
[data-testid="stChatInput"] button {
    background: #2563EB !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 10px !important;
}
[data-testid="stChatInput"] [data-testid="stChatInputFileUploadMessage"],
[data-testid="stChatInput"] [data-testid="stFileUploadDropzone"],
[data-testid="stChatInput"] [class*="uploadedFile"] {
    color: #0F172A !important;
    background: #F8FAFC !important;
}
</style>
"""


_PLACEHOLDERS = {
    "xai-...", "xai-", "...", "sk-ant-...", "sk-ant-", "sk-...", "PEGAR_CLAVE",
}
_SECRETS_PATH = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"


def _guardar_clave_xai(clave: str) -> None:
    """Guarda la clave para TODO el estudio (cualquier usuario). No va a GitHub."""
    texto = (clave or "").strip().strip('"').strip("'")
    if not texto.startswith("xai-") or texto in _PLACEHOLDERS or len(texto) < 20:
        return
    _SECRETS_PATH.parent.mkdir(parents=True, exist_ok=True)
    actual = ""
    if _SECRETS_PATH.exists():
        actual = _SECRETS_PATH.read_text(encoding="utf-8")
    linea = f"XAI_API_KEY = {json.dumps(texto)}\n"
    modelo = 'XAI_MODEL = "grok-4.6"\n'
    if re.search(r"^XAI_API_KEY\s*=", actual, re.M):
        actual = re.sub(r"^XAI_API_KEY\s*=.*$", linea.rstrip(), actual, count=1, flags=re.M)
    else:
        actual = linea + actual
    if not re.search(r"^XAI_MODEL\s*=", actual, re.M):
        actual = modelo + actual
    _SECRETS_PATH.write_text(actual.lstrip("\n") if actual.startswith("\n") else actual, encoding="utf-8")
    os.environ["XAI_API_KEY"] = texto
    os.environ.setdefault("XAI_MODEL", "grok-4.6")


def _secrets_disco() -> dict:
    if not _SECRETS_PATH.exists():
        return {}
    try:
        import tomllib
        data = tomllib.loads(_SECRETS_PATH.read_bytes())
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _secret(nombre: str) -> str:
    disco = _secrets_disco().get(nombre)
    if disco:
        texto = str(disco).strip().strip('"').strip("'")
        if texto and texto not in _PLACEHOLDERS:
            return texto
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


@st.fragment
def render_tango_bot() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)
    cargar_indice()

    if "tango_chat" not in st.session_state:
        st.session_state.tango_chat = []

    if "tango_api_key" not in st.session_state:
        st.session_state.tango_api_key = str(st.session_state.get("tango_xai_key") or "")
    if "tango_clave_invalida" not in st.session_state:
        st.session_state.tango_clave_invalida = False
    web = _es_web_publica()
    pasted = str(st.session_state.get("tango_api_key") or "").strip()
    if pasted in _PLACEHOLDERS:
        pasted = ""
    if pasted.startswith("xai-"):
        _guardar_clave_xai(pasted)
    secret_muerto = bool(st.session_state.get("tango_clave_invalida")) and not pasted
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
        if nombre == "XAI_API_KEY" and secret_muerto:
            os.environ.pop("XAI_API_KEY", None)
            continue
        val = _secret(nombre)
        if val:
            os.environ[nombre] = val
    # Clave pegada en Tango (esta sesión) pisa Secrets, también en la web.
    api_key = pasted
    if not api_key and not secret_muerto:
        api_key = _secret("XAI_API_KEY")
    if not api_key:
        api_key = (
            _secret("ANTHROPIC_API_KEY")
            or _secret("OPENAI_API_KEY")
            or _secret("GROQ_API_KEY")
        )
    if api_key.startswith("xai-"):
        os.environ["XAI_API_KEY"] = api_key
    elif api_key.startswith("sk-ant-"):
        os.environ["ANTHROPIC_API_KEY"] = api_key
    model = _modelo_para_clave(api_key)
    hay_ia = bool(api_key) and not st.session_state.get("tango_clave_invalida")
    etiqueta = _etiqueta_ia(api_key, model)
    chat = st.session_state.tango_chat

    def _panel_ia() -> None:
        if hay_ia and not st.session_state.get("tango_clave_invalida"):
            st.caption("Grok está activo para todo el estudio.")
            return
        if st.session_state.get("tango_clave_invalida"):
            st.error(
                "**Grok está apagado.** La clave no vale. "
                "Pegá una nueva abajo y tocá Activar Grok."
            )
        else:
            st.error("**Grok está apagado.** Pegá la clave de xAI para activarlo.")
        st.caption("Pegá acá la clave de Grok. Queda guardada en esta PC, no en GitHub.")
        clave = st.text_input(
            "Clave Grok (xAI)",
            type="password",
            key="tango_api_input",
            placeholder="xai-...",
            help="Empieza con xai-. No la pegues en el chat.",
        )
        if st.button("Activar Grok", type="primary", key="tango_activar_ia"):
            texto = clave.strip()
            if texto.startswith("xai-") or texto.startswith("sk-ant-") or texto.startswith("sk-"):
                st.session_state.tango_api_key = texto
                st.session_state.tango_clave_invalida = False
                if texto.startswith("xai-"):
                    _guardar_clave_xai(texto)
                st.rerun()
            else:
                st.warning("La clave de Grok empieza con xai-.")

    if chat:
        top_l, top_r = st.columns([4, 1])
        with top_l:
            with st.expander("Grok", expanded=not hay_ia):
                _panel_ia()
                if not web:
                    if st.button("Reindexar ayudas Tango", key="tango_reindex"):
                        with st.spinner("Leyendo el manual HTML de Tango…"):
                            docs = reconstruir_indice(guardar=True)
                        st.success(f"Listo: {len(docs)} fragmentos HTML (sin PDF)")
        with top_r:
            if st.button("Nueva charla", key="tango_clear", use_container_width=True):
                st.session_state.tango_chat = []
                _rerun_chat()

    if not chat:
        st.markdown(
            '<div class="tango-hero"><h1>Agente Tango</h1>'
            "<p>Preguntá por una fórmula o adjuntá una captura.</p></div>",
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
                    _rerun_chat()
    else:
        for msg in chat:
            with st.chat_message(msg["role"]):
                for data_url in msg.get("imagenes") or []:
                    st.image(data_url, width=480, caption="Captura de Tango")
                st.markdown(msg["content"])
                if msg.get("role") == "assistant":
                    if msg.get("pide_formula") and not msg.get("formula_encontrada"):
                        st.warning(
                            "No encontré esa fórmula en el export de Tango Sueldos. "
                            "Verificá en Tango antes de pegar nada."
                        )
                    fuentes = msg.get("fuentes") or []
                    if fuentes:
                        with st.expander("De dónde salió"):
                            for fte in fuentes:
                                titulo = str(fte.get("title") or "").strip() or "(sin título)"
                                origen = str(fte.get("source") or "").strip()
                                st.caption(f"• {titulo}" + (f" — {origen}" if origen else ""))
                    if msg.get("uso_ia") is False:
                        st.caption("Sin Grok en esta respuesta: solo material del estudio.")

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
        _rerun_chat()


def _rerun_chat() -> None:
    try:
        st.rerun(scope="fragment")
    except TypeError:
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
    out = responder(
        prompt,
        historial,
        api_key=api_key,
        model=model,
        imagenes=imagenes,
    )
    if out.get("clave_invalida"):
        st.session_state.tango_clave_invalida = True
        st.session_state.tango_api_key = ""
    st.session_state.tango_chat.append({
        "role": "assistant",
        "content": out["texto"],
        "fuentes": out.get("fuentes") or [],
        "pide_formula": bool(out.get("pide_formula")),
        "formula_encontrada": bool(out.get("formula_encontrada")),
        "uso_ia": bool(out.get("uso_ia")),
    })
