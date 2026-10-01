"""Solapa Rutinas: la oficina encola tareas. No se ejecutan en esta web."""
from __future__ import annotations

import hashlib
import time
from datetime import timedelta

import streamlit as st
from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx

import cache_lecturas
import database
from rutinas import (
    AVISO_EJERCICIO,
    OPCION_TODOS,
    RUTINAS,
    ErrorRutina,
    PedidoDuplicado,
    cancelar_pedido,
    crear_pedido,
    es_proyeccion,
    etiquetas_clientes,
    exportar_ficha,
    filas_tabla_preview,
    formatear_fecha_ar,
    importar_ficha_xlsx,
    listar_pedidos,
    parametros_proyeccion,
    pedido_abierto,
    periodo_mmaaaa,
    periodo_sugerido,
    periodo_valido,
    resolver_clientes,
    resumen_parametros,
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
        "Elegí la rutina y apretá Ejecutar. "
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

    if es_proyeccion(codigo):
        parametros, requisitos_ok = _formulario_proyeccion(item)
    else:
        parametros, requisitos_ok = _formulario_checklist(item)

    if "rutinas_solicitado_por" not in st.session_state:
        st.session_state["rutinas_solicitado_por"] = _nombre_sesion()
    quien = st.text_input(
        "Quién pide",
        key="rutinas_solicitado_por",
        placeholder="Nombre de quien pide la rutina",
        help="Queda asentado en el pedido. Es obligatorio.",
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

    if not requisitos_ok and es_proyeccion(codigo):
        st.caption("Tildá los 4 requisitos y elegí el período y al menos un cliente.")
    elif not requisitos_ok:
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


def _formulario_checklist(item: dict) -> tuple[str, bool]:
    st.markdown("##### Requisitos")
    st.caption("Tildá cada uno. Ejecutar se habilita cuando están todos. Los que dicen A CONFIRMAR se van a ajustar.")
    tildes = []
    for indice, texto in enumerate(item["requisitos"]):
        tildes.append(st.checkbox(texto, key=f"rutina_req_{item['codigo']}_{indice}"))
    parametros = st.text_input(
        "Parámetros (opcional)",
        key=f"rutina_param_{item['codigo']}",
        placeholder="Período o cliente, si hace falta",
    )
    return str(parametros or "").strip(), all(tildes)


def _formulario_proyeccion(item: dict) -> tuple[str, bool]:
    st.markdown("##### Requisitos")
    st.caption(" ".join(item.get("ayuda") or ()))
    tildes = []
    for indice, texto in enumerate(item["requisitos"]):
        tildes.append(st.checkbox(texto, key=f"rutina_req_{item['codigo']}_{indice}"))

    st.markdown("##### Ficha de clientes")
    archivo = st.file_uploader(
        "Actualizar ficha de clientes (xlsx)",
        type=["xlsx"],
        key="rutina_ficha_uploader",
        help="Columnas: Sociedad, CUIT, Activa, TISH, Mes inicio, Mes cierre, Día revisión, Proyección vigente, Carpeta Proyecciones, Carpeta IIBB, Carpeta F931, Último mes cargado, Notas.",
    )
    if archivo is not None:
        datos = archivo.getvalue()
        firma = hashlib.sha256(datos).hexdigest()
        if st.session_state.get("_ficha_hash") != firma:
            try:
                cantidad = importar_ficha_xlsx(datos)
            except ErrorRutina as exc:
                st.error(str(exc))
            else:
                st.session_state["_ficha_hash"] = firma
                st.session_state["rutina_flash"] = f"Ficha actualizada: {cantidad} cliente(s)."
                st.rerun()

    ficha = exportar_ficha()
    activos = etiquetas_clientes()
    st.caption(f"{len(ficha)} cliente(s) en la ficha · {len(activos)} activo(s).")
    if ficha:
        with st.expander("Ver ficha cargada", expanded=False):
            st.dataframe(ficha, use_container_width=True, hide_index=True)
    else:
        st.info("Todavía no hay ficha. Subí el Excel para poder elegir clientes.")

    sugerido = periodo_sugerido()
    if "proy_mes" not in st.session_state:
        st.session_state["proy_mes"] = sugerido[:2]
    if "proy_anio" not in st.session_state:
        st.session_state["proy_anio"] = int(sugerido[3:])
    col_mes, col_anio = st.columns(2)
    with col_mes:
        mes = st.selectbox("Mes", [f"{numero:02d}" for numero in range(1, 13)], key="proy_mes")
    with col_anio:
        anio = st.number_input("Año", min_value=2000, max_value=2100, step=1, key="proy_anio")
    periodo = f"{mes}-{int(anio)}"
    if periodo_valido(periodo):
        st.caption(f"Período: {periodo} · carpeta {periodo_mmaaaa(periodo)}")
    else:
        st.caption("El período tiene que ser MM-AAAA.")

    opciones = [OPCION_TODOS] + [etiqueta for etiqueta, _fila in activos]
    elegidos = st.multiselect(
        "Clientes",
        opciones,
        key="proy_clientes",
        placeholder="Elegí clientes o Todos",
        disabled=not activos,
    )
    clientes, todos = resolver_clientes(list(elegidos))
    if todos:
        st.caption(f"Todos: {len(clientes)} cliente(s) activo(s).")
    if not all(tildes) or not clientes or not periodo_valido(periodo):
        return "", False
    try:
        return parametros_proyeccion(periodo, clientes, todos=todos), True
    except ErrorRutina as exc:
        st.error(str(exc))
        return "", False


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
    fragmento = _rerun_solo_fragmento()
    t0 = time.perf_counter() if fragmento else 0.0
    antes = database.metricas_rerun() if fragmento else None
    try:
        if fragmento:
            _refrescar_cola()
        _zona(codigo, en_vivo=True)
        vivo = pedido_abierto(codigo)
        if vivo is None or str(vivo["estado"]) != estado_al_abrir:
            st.rerun(scope="app")
    finally:
        if fragmento and antes is not None:
            despues = database.metricas_rerun()
            database.imprimir_linea_rerun(
                total=time.perf_counter() - t0,
                sync_s=float(despues["sync_s"]) - float(antes["sync_s"]),
                consultas=int(despues["consultas"]) - int(antes["consultas"]),
                syncs=int(despues["syncs"]) - int(antes["syncs"]),
                pagina="Rutinas",
            )


def _zona(codigo: str, *, en_vivo: bool = False) -> None:
    st.divider()
    _vista_previa(codigo, en_vivo=en_vivo)
    st.divider()
    _historial()


def _rerun_solo_fragmento() -> bool:
    """True cuando Streamlit reejecuta solo el fragmento (auto-refresh o el botón)."""
    try:
        ctx = get_script_run_ctx(suppress_warning=True)
    except TypeError:
        ctx = get_script_run_ctx()
    except Exception:
        return False
    if ctx is None:
        return False
    return bool(getattr(ctx, "fragment_ids_this_run", None))


def _refrescar_cola() -> None:
    """Sync de rutina_pedidos y se tira el cache corto para leer lo que acaba de llegar."""
    database.sincronizar_cola_rutinas()
    cache_lecturas.pedido_abierto.clear()
    cache_lecturas.ultimo_pedido.clear()
    cache_lecturas.pedidos.clear()


def _vista_previa(codigo: str, *, en_vivo: bool) -> None:
    st.markdown("##### Vista previa")
    if es_proyeccion(codigo):
        st.info(AVISO_EJERCICIO)
    if en_vivo:
        st.caption("Se actualiza sola cada 5 segundos mientras esta rutina está pendiente o en curso.")
    if st.button("Refrescar", key="rutina_refrescar"):
        # Dentro del fragmento el sync lo hace el propio rerun del fragmento.
        if not _rerun_solo_fragmento():
            _refrescar_cola()
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
            "Parámetros": resumen_parametros(pedido.get("parametros")),
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
