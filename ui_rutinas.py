"""Solapa Rutinas: la oficina encola tareas. No se ejecutan en esta web."""
from __future__ import annotations

import streamlit as st

from rutinas import (
    RUTINAS,
    ErrorRutina,
    PedidoDuplicado,
    cancelar_pedido,
    crear_pedido,
    formatear_fecha_ar,
    listar_pedidos,
    pedido_abierto,
)

_HISTORIAL = 40
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


def render_rutinas() -> None:
    st.markdown("##### Rutinas de oficina")
    st.caption(
        "Apretar Ejecutar deja el pedido en la cola. "
        "Lo corre el asistente de la oficina (tiene el servidor de archivos y ARCA). "
        "Esta web no ejecuta la tarea."
    )
    aviso = st.session_state.pop("rutina_flash", None)
    if aviso:
        st.success(aviso)

    if "rutinas_solicitado_por" not in st.session_state:
        st.session_state["rutinas_solicitado_por"] = _nombre_sesion()
    st.text_input(
        "Quién pide",
        key="rutinas_solicitado_por",
        placeholder="Nombre de quien pide la rutina",
        help="Queda asentado en el pedido. Es obligatorio.",
    )

    for item in RUTINAS:
        _tarjeta(item)

    codigo = st.session_state.get("rutina_confirmar")
    if codigo:
        _dialogo_ejecutar(str(codigo))

    st.divider()
    _historial()


def _nombre_sesion() -> str:
    return str(
        st.session_state.get("usuario_oficina_nombre")
        or st.session_state.get("usuario_oficina")
        or ""
    ).strip()


def _tarjeta(item: dict[str, str]) -> None:
    codigo = item["codigo"]
    abierto = pedido_abierto(codigo)
    with st.container(border=True):
        st.markdown(f"**{item['nombre']}**")
        st.caption(item["descripcion"])
        st.text_input(
            "Parámetros (opcional)",
            key=f"rutina_param_{codigo}",
            placeholder="Período o cliente, si hace falta",
            disabled=abierto is not None,
        )
        if abierto and abierto["estado"] == "PENDIENTE":
            st.warning(
                f"Ya hay un pedido pendiente (n.º {abierto['id']}, "
                f"lo pidió {abierto['solicitado_por']}). No se crea otro."
            )
            if st.button("Cancelar pedido", key=f"rutina_cancel_{codigo}"):
                try:
                    cancelar_pedido(int(abierto["id"]))
                except ErrorRutina as exc:
                    st.error(str(exc))
                else:
                    st.session_state["rutina_flash"] = f"Pedido n.º {abierto['id']} cancelado."
                    st.rerun()
        elif abierto:
            st.info(
                f"El asistente ya la está ejecutando (pedido n.º {abierto['id']}, "
                f"lo pidió {abierto['solicitado_por']})."
            )
        elif st.button("Ejecutar", type="primary", key=f"rutina_ejecutar_{codigo}"):
            quien = str(st.session_state.get("rutinas_solicitado_por") or "").strip()
            if not quien:
                st.error("Escribí quién pide la rutina.")
            else:
                st.session_state["rutina_confirmar"] = codigo


@st.dialog("Ejecutar rutina")
def _dialogo_ejecutar(codigo: str) -> None:
    item = next((r for r in RUTINAS if r["codigo"] == codigo), None)
    if item is None:
        st.session_state.pop("rutina_confirmar", None)
        st.warning("Esa rutina ya no está en el listado.")
        return
    params = str(st.session_state.get(f"rutina_param_{codigo}") or "").strip()
    quien = str(st.session_state.get("rutinas_solicitado_por") or "").strip()
    st.markdown(f"**{item['nombre']}**")
    st.caption(item["descripcion"])
    if params:
        st.write(f"Parámetros: {params}")
    else:
        st.caption("Sin parámetros.")
    st.write(f"Pide: {quien}")
    st.caption("Confirmar la pone en cola. No hace falta otra aprobación.")
    col_ok, col_no = st.columns(2)
    with col_ok:
        if st.button("Confirmar", type="primary", key="rutina_dlg_ok"):
            if not quien:
                st.error("Escribí quién pide la rutina.")
                return
            try:
                pedido = crear_pedido(codigo, quien, params or None)
            except PedidoDuplicado as exc:
                st.warning(str(exc))
                return
            except ErrorRutina as exc:
                st.error(str(exc))
                return
            st.session_state.pop("rutina_confirmar", None)
            st.session_state["rutina_flash"] = (
                f"Pedido n.º {pedido['id']} en cola ({item['nombre']})."
            )
            st.rerun()
    with col_no:
        if st.button("Volver", key="rutina_dlg_volver"):
            st.session_state.pop("rutina_confirmar", None)
            st.rerun()


def _historial() -> None:
    st.markdown("##### Historial")
    st.caption("Horario de Argentina (America/Buenos_Aires).")
    if st.button("Refrescar", key="rutina_refrescar"):
        st.rerun()
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
