# -*- coding: utf-8 -*-
"""Aviso de versión nueva + listado de cambios al actualizar la web."""
from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

import streamlit as st
import streamlit.components.v1 as components

BASE_DIR = Path(__file__).resolve().parent
CAMBIOS_PATH = BASE_DIR / "data" / "cambios_web.json"
GITHUB_RAW = (
    "https://raw.githubusercontent.com/zustovichmartina-ux/"
    "Estudio-Contable/master/data/cambios_web.json"
)
STORAGE_KEY = "estudio_web_v"
_POLL_SEG = 45.0


def cargar_cambios() -> dict[str, Any]:
    if not CAMBIOS_PATH.is_file():
        return {"version": "", "cambios": []}
    try:
        data = json.loads(CAMBIOS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"version": "", "cambios": []}
    if not isinstance(data, dict):
        return {"version": "", "cambios": []}
    items = data.get("cambios")
    if not isinstance(items, list):
        items = []
    return {
        "version": str(data.get("version") or ""),
        "cambios": [c for c in items if isinstance(c, dict) and str(c.get("id") or "")],
    }


def version_local(data: dict[str, Any] | None = None) -> str:
    payload = data if data is not None else cargar_cambios()
    ver = str(payload.get("version") or "").strip()
    if ver:
        return ver
    cambios = payload.get("cambios") or []
    if cambios:
        return str(cambios[0].get("id") or "").strip()
    return ""


def cambios_desde(data: dict[str, Any], desde_id: str) -> list[dict[str, Any]]:
    """Entradas más nuevas que `desde_id` (el archivo va de lo último a lo viejo)."""
    filas = list(data.get("cambios") or [])
    if not desde_id:
        return filas[:1]
    out: list[dict[str, Any]] = []
    for fila in filas:
        cid = str(fila.get("id") or "")
        if cid == desde_id:
            break
        out.append(fila)
    return out or filas[:1]


def _fetch_version_remota() -> str:
    req = urllib.request.Request(
        GITHUB_RAW,
        headers={
            "User-Agent": "EstudioContable-Web/1.0",
            "Cache-Control": "no-cache",
            "Accept": "application/json",
        },
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=4) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
    parsed = json.loads(raw)
    if not isinstance(parsed, dict):
        return ""
    ver = str(parsed.get("version") or "").strip()
    if ver:
        return ver
    cambios = parsed.get("cambios") or []
    if isinstance(cambios, list) and cambios and isinstance(cambios[0], dict):
        return str(cambios[0].get("id") or "").strip()
    return ""


def _version_github_cached() -> str:
    now = time.monotonic()
    ts = float(st.session_state.get("_ec_remoto_ts") or 0)
    cached = str(st.session_state.get("_ec_remoto_id") or "")
    if cached and (now - ts) < _POLL_SEG:
        return cached
    try:
        remoto = _fetch_version_remota()
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, ValueError):
        remoto = cached
    st.session_state["_ec_remoto_id"] = remoto
    st.session_state["_ec_remoto_ts"] = now
    return remoto


def _inyectar_bridge(version: str, marcar_visto: bool) -> None:
    visto_js = "true" if marcar_visto else "false"
    ver = json.dumps(version)
    key = json.dumps(STORAGE_KEY)
    components.html(
        f"""
<script>
(function () {{
  var CUR = {ver};
  var KEY = {key};
  var marcar = {visto_js};
  try {{
    var store = window.parent.localStorage;
    if (marcar && CUR) {{
      store.setItem(KEY, CUR);
      return;
    }}
    var prev = store.getItem(KEY) || "";
    var loc = window.parent.location;
    var u = new URL(loc.href);
    if (prev && CUR && prev !== CUR && u.searchParams.get("ec_novedades") !== "1") {{
      u.searchParams.set("ec_novedades", "1");
      u.searchParams.set("ec_desde", prev);
      loc.replace(u.toString());
      return;
    }}
    if (!prev && CUR) {{
      store.setItem(KEY, CUR);
    }}
  }} catch (e) {{}}
}})();
</script>
""",
        height=1,
    )


def _render_novedades(data: dict[str, Any], actual: str) -> None:
    qp = st.query_params
    if str(qp.get("ec_novedades") or "") != "1":
        return
    desde = str(qp.get("ec_desde") or "")
    nuevos = cambios_desde(data, desde)
    st.success("La web se actualizó a la última versión.")
    for fila in nuevos:
        titulo = str(fila.get("titulo") or "Cambios")
        fecha = str(fila.get("fecha") or "")
        cabe = f"**{fecha} — {titulo}**" if fecha else f"**{titulo}**"
        st.markdown(cabe)
        for item in fila.get("items") or []:
            texto = str(item).strip()
            if texto:
                st.markdown(f"- {texto}")
    if st.button("Entendido", type="primary", key="ec_novedades_ok"):
        st.session_state["_ec_marcar_visto"] = actual
        try:
            del st.query_params["ec_novedades"]
        except KeyError:
            pass
        try:
            del st.query_params["ec_desde"]
        except KeyError:
            pass
        st.rerun()


@st.fragment(run_every=60)
def _chequeo_remoto(actual: str) -> None:
    if not actual:
        return
    remoto = _version_github_cached()
    if not remoto or remoto == actual:
        return
    st.warning("Hay una versión nueva de la web. Actualizá para ver los cambios.")
    if os.name == "nt" and not Path("/mount/src").is_dir():
        st.caption(
            "Si estás en una PC del estudio (no streamlit.app), primero hay que "
            "bajar el código nuevo y reabrir la web."
        )
    if st.button("Actualizar a la última versión", type="primary", key="ec_web_actualizar"):
        components.html(
            "<script>window.parent.location.reload();</script>",
            height=1,
        )


def render_aviso_version() -> None:
    """Cartel de ‘actualizá’ si GitHub está más nuevo; al recargar, lista de cambios."""
    data = cargar_cambios()
    actual = version_local(data)
    if not actual:
        return
    marcar = str(st.session_state.get("_ec_marcar_visto") or "") == actual
    _inyectar_bridge(actual, marcar)
    if marcar:
        st.session_state.pop("_ec_marcar_visto", None)
    _render_novedades(data, actual)
    _chequeo_remoto(actual)
