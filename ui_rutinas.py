"""Solapa Rutinas: la oficina encola tareas. No se ejecutan en esta web."""
from __future__ import annotations

from datetime import timedelta

import streamlit as st

from rutinas import (
    RUTINAS,
    ErrorRutina,
    PedidoDuplicado,
    cancelar_pedido,
    crear_pedido,
    filas_tabla_preview,
    formatear_fecha_ar,
    listar_pedidos,
    pedido_abierto,
    ultimo_pedido,
)

_HISTORIAL = 40
_REFRESCO = timedelta(seconds=5)
_COLUMNAS = (
    "Nro",
    "Rutina",
    "Parámetros",
    "Pidió",
    "Estado",
    "Creado",
    "Tomado",
    "Terminado",
    "Resultado",
    "Archivos",
)
_ESTADOS_CERRADOS = ("OK", "ERROR", "CANCELADO")


def render_rutinas() -> None:
    st.markdown("##### Rutinas de oficina")
    st.caption(
        "Elegí la rutina, tildá los requisitos y apretá Ejecutar. "
        "Eso la deja en la cola: la corre el asistente (archivos y ARCA), no esta web."
    )
    aviso = st.session_state.pop("rutina_flash", None)
    if aviso:
        st.success(aviso)

    nombres = [item["nombre"] for item in RUTINAS]
    nombre = st.selectbox("Rutina", nombres, key="rutina_elegida")
    item = next(r for r in RUTINAS if r["nombre"] == nombre)
    codigo = item["codigo"]
    st.write(item["descripcion"])

    st.markdown("##### Requisitos")
    st.caption("Tildá cada uno. Ejecutar se habilita cuando están todos. Los que dicen A CONFIRMAR se van a ajustar.")
    tildes = []
    for indice, texto in enumerate(item["requisitos"]):
        tildes.append(st.checkbox(texto, key=f"rutina_req_{codigo}_{indice}"))
    requisitos_ok = all(tildes)

    if "rutinas_solicitado_por" not in st.session_state:
        st.session_state["rutinas_solicitado_por"] = _nombre_sesion()
    quien = st.text_input(
        "Quién pide",
        key="rutinas_solicitado_por",
        placeholder="Nombre de quien pide la rutina",
        help="Queda asentado en el pedido. Es obligatorio.",
    )
    parametros = st.text_input(
        "Parámetros (opcional)",
        key=f"rutina_param_{codigo}",
        placeholder="Período o cliente, si hace falta",
    )

    abierto = pedido_abierto(codigo)
    ultimo = ultimo_pedido(codigo)
    puede_listo = _puede_marcar_listo(ultimo)
    col_ejecutar, col_listo = st.columns(2)
    with col_ejecutar:
        puede = requisitos_ok and bool(str(quien or "").strip()) and abierto is None
        if st.button("Ejecutar", type="primary", disabled=not puede, key="rutina_ejecutar"):
            _encolar(codigo, quien, parametros)
    with col_listo:
        if st.button("Listo", disabled=not puede_listo, key="rutina_listo"):
            _marcar_visto(int(ultimo["id"]))
            st.rerun()

    if not requisitos_ok:
        st.caption("Tildá todos los requisitos para poder ejecutar.")
    elif not str(quien or "").strip():
        st.caption("Escribí quién pide la rutina.")
    if abierto and abierto["estado"] == "PENDIENTE":
        st.warning(
            f"Ya hay un pedido pendiente (n.º {abierto['id']}, "
            f"lo pidió {abierto['solicitado_por']}). No se crea otro."
        )
        if st.button("Cancelar pedido", key="rutina_cancelar"):
            _cancelar(int(abierto["id"]))
    elif abierto:
        st.info(
            f"El asistente ya la está ejecutando (pedido n.º {abierto['id']}, "
            f"lo pidió {abierto['solicitado_por']})."
        )

    if abierto:
        _zona_viva(codigo, str(abierto["estado"]))
    else:
        _zona(codigo)


def _nombre_sesion() -> str:
    return str(
        st.session_state.get("usuario_oficina_nombre")
        or st.session_state.get("usuario_oficina")
        or ""
    ).strip()


def _vistos() -> list:
    vistos = st.session_state.get("rutinas_vistos")
    if not isinstance(vistos, list):
        vistos = []
        st.session_state["rutinas_vistos"] = vistos
    return vistos


def _marcar_visto(pedido_id: int) -> None:
    vistos = _vistos()
    if pedido_id not in vistos:
        vistos.append(pedido_id)


def _puede_marcar_listo(pedido: dict | None) -> bool:
    if not pedido or pedido.get("estado") not in _ESTADOS_CERRADOS:
        return False
    return int(pedido["id"]) not in _vistos()


def _encolar(codigo: str, quien: str, parametros: str) -> None:
    try:
        pedido = crear_pedido(codigo, str(quien or "").strip(), parametros or None)
    except PedidoDuplicado as exc:
        st.warning(str(exc))
        return
    except ErrorRutina as exc:
        st.error(str(exc))
        return
    st.session_state["rutina_flash"] = f"Pedido n.º {pedido['id']} en cola."
    st.rerun()


def _cancelar(pedido_id: int) -> None:
    try:
        cancelar_pedido(pedido_id)
    except ErrorRutina as exc:
        st.error(str(exc))
        return
    st.session_state["rutina_flash"] = f"Pedido n.º {pedido_id} cancelado."
    st.rerun()


@st.fragment(run_every=_REFRESCO)
def _zona_viva(codigo: str, estado_al_abrir: str) -> None:
    """Se actualiza sola mientras el pedido de esta rutina sigue abierto."""
    _zona(codigo, en_vivo=True)
    vivo = pedido_abierto(codigo)
    if vivo is None or str(vivo["estado"]) != estado_al_abrir:
        st.rerun(scope="app")


def _zona(codigo: str, *, en_vivo: bool = False) -> None:
    st.divider()
    _vista_previa(codigo, en_vivo=en_vivo)
    st.divider()
    _historial()


def _vista_previa(codigo: str, *, en_vivo: bool) -> None:
    st.markdown("##### Vista previa")
    if en_vivo:
        st.caption("Se actualiza sola cada 5 segundos mientras esta rutina está pendiente o en curso.")
    if st.button("Refrescar", key="rutina_refrescar"):
        st.rerun()

    pedido = ultimo_pedido(codigo)
    if not pedido:
        st.caption("Todavía no hay una ejecución de esta rutina.")
        return
    if pedido.get("estado") in _ESTADOS_CERRADOS and int(pedido["id"]) in _vistos():
        st.caption(f"Resultado del pedido n.º {pedido['id']} marcado como visto.")
        return

    st.markdown(
        f"**Pedido n.º {pedido['id']}** · estado **{pedido['estado']}** · "
        f"pidió {pedido.get('solicitado_por') or '—'}"
    )
    fechas = []
    if pedido.get("creado_en"):
        fechas.append(f"creado {formatear_fecha_ar(pedido.get('creado_en'))}")
    if pedido.get("tomado_en"):
        fechas.append(f"tomado {formatear_fecha_ar(pedido.get('tomado_en'))}")
    if pedido.get("terminado_en"):
        fechas.append(f"terminado {formatear_fecha_ar(pedido.get('terminado_en'))}")
    if fechas:
        st.caption(" · ".join(fechas) + " (hora Argentina)")

    if pedido.get("estado") in ("PENDIENTE", "EN_CURSO"):
        st.info("El asistente todavía no terminó. La vista previa aparece cuando cierra el pedido.")
        return

    preview = pedido.get("preview") if isinstance(pedido.get("preview"), dict) else {}
    resumen = str(preview.get("resumen_md") or pedido.get("resultado") or "").strip()
    if resumen:
        st.markdown(resumen)
    for tabla in preview.get("tablas") or []:
        titulo = str(tabla.get("titulo") or "Cambios")
        st.markdown(f"**{titulo}**")
        filas = filas_tabla_preview(tabla)
        if filas:
            st.dataframe(filas, use_container_width=True, hide_index=True)
        else:
            st.caption("Sin filas.")
    _rutas_servidor(preview, pedido)


def _rutas_servidor(preview: dict, pedido: dict) -> None:
    rutas = [str(ruta) for ruta in (preview.get("archivos") or []) if str(ruta).strip()]
    if not rutas and pedido.get("archivos"):
        rutas = [linea.strip() for linea in str(pedido["archivos"]).splitlines() if linea.strip()]
    if not rutas:
        return
    st.markdown("**Archivos en el servidor**")
    st.caption("Rutas del servidor de archivos. No se descarga nada desde acá.")
    for ruta in rutas:
        st.text(ruta)


def _historial() -> None:
    st.markdown("##### Historial")
    st.caption("Horario de Argentina (America/Buenos_Aires).")
    try:
        pedidos = listar_pedidos(limite=_HISTORIAL)
    except ErrorRutina as exc:
        st.error(str(exc))
        return
    if not pedidos:
        st.caption("Todavía no hay pedidos.")
        return
    filas = [
        {
            "Nro": pedido["id"],
            "Rutina": pedido.get("nombre") or pedido["rutina"],
            "Parámetros": pedido.get("parametros") or "",
            "Pidió": pedido.get("solicitado_por") or "",
            "Estado": pedido.get("estado") or "",
            "Creado": formatear_fecha_ar(pedido.get("creado_en")),
            "Tomado": formatear_fecha_ar(pedido.get("tomado_en")),
            "Terminado": formatear_fecha_ar(pedido.get("terminado_en")),
            "Resultado": pedido.get("resultado") or "",
            "Archivos": pedido.get("archivos") or "",
        }
        for pedido in pedidos
    ]
    st.dataframe(filas, column_order=list(_COLUMNAS), use_container_width=True, hide_index=True)
