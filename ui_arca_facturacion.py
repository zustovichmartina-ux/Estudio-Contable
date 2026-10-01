"""Solapa Facturación dentro de ARCA: emitir, consultar y datos del emisor."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st

from arca.codigos import ALICUOTAS, COND_IVA_RECEPTOR, CONCEPTOS, DOC_TIPOS, TIPOS_CBTE
from arca.comprobante import COND_IVA_EMISOR, COND_VENTA, a_fecha, cuit_valido, solo_digitos
from arca.formulario import (
    campos_desde_asociado,
    campos_desde_emisor,
    candidatos_asociados,
    etiqueta_asociado,
    mensaje_cuit,
    tipos_para_condicion,
)
from arca.persistencia import cargar_emisor, guardar_emisor, listar_emisiones, listar_emisores
from arca.secretos import CUIT_ESTUDIO, CredencialesError
from arca.servicio import (
    EmisionAbortada,
    WSFEError,
    categoria_monotributo,
    consultar_y_pdf,
    emitir_excel,
    excel_desde_filas,
    listar_puntos_venta,
    pdf_de_emision,
    ultimos_numeros,
    validar_excel,
    zip_pdfs,
)

_RUTA_MODELO = Path(__file__).resolve().parent / "plantillas" / "Facturas_a_emitir_MODELO.xlsx"
_COLUMNAS_VISTA = [
    "Filas Facturas", "ID factura", "CUIT emisor", "Tipo", "Número", "CAE",
    "Vto. CAE", "Estado", "Total", "Errores / observaciones",
]


def render_facturacion_arca() -> None:
    st.markdown("##### Facturación electrónica (WSFEv1)")
    st.caption(
        f"El certificado es del estudio (CUIT {CUIT_ESTUDIO[:2]}-{CUIT_ESTUDIO[2:10]}-{CUIT_ESTUDIO[10:]}). "
        "Cada cliente tiene que haber delegado Facturación Electrónica a esa CUIT. "
        "El comprobante sale con la CUIT del cliente, en un punto de venta de Web Services."
    )
    env, prod_ok = _ambiente()
    tab_emitir, tab_consulta, tab_emisor = st.tabs(["Emitir", "Consulta", "Datos del emisor"])
    with tab_emitir:
        _tab_emitir(env, prod_ok)
    with tab_consulta:
        _tab_consulta(env, prod_ok)
    with tab_emisor:
        _tab_emisor()


def _ambiente() -> tuple[str, bool]:
    etiqueta = st.radio(
        "Ambiente",
        ["Homologación", "Producción"],
        horizontal=True,
        key="fe_ambiente",
        help="Homologación no tiene validez fiscal. Producción pide una confirmación extra.",
    )
    env = "homo" if etiqueta == "Homologación" else "prod"
    if env == "homo":
        st.caption("Homologación: el CAE no tiene validez y el PDF sale marcado.")
        return env, True
    st.warning("Producción emite comprobantes con validez fiscal.")
    ok_check = st.checkbox(
        "Confirmo que voy a operar en PRODUCCIÓN",
        key="fe_prod_check",
    )
    frase = st.text_input(
        "Para habilitar producción escribí EMITIR",
        key="fe_prod_frase",
    )
    prod_ok = bool(ok_check) and frase.strip().upper() == "EMITIR"
    if not prod_ok:
        st.caption("Hasta que no confirmes, no se consulta ni se emite en producción.")
    return env, prod_ok


def _tab_emitir(env: str, prod_ok: bool) -> None:
    if st.session_state.pop("_fe_reset_revise", None):
        st.session_state.pop("fe_revise", None)
    _descarga_modelo()
    _panel_resultado_previo()
    borrador = st.session_state.get("fe_borrador")
    if borrador and borrador.get("ambiente") == env:
        _panel_borrador(borrador, env, prod_ok)
    modo = st.radio("Carga", ["Un comprobante", "Lote desde Excel"], horizontal=True, key="fe_modo_carga")
    if modo == "Un comprobante":
        _formulario_manual(env)
    else:
        _carga_excel(env)


def _descarga_modelo() -> None:
    if not _RUTA_MODELO.is_file():
        st.error("No está el Excel modelo en el servidor.")
        return
    st.download_button(
        "Descargar Excel modelo",
        data=_RUTA_MODELO.read_bytes(),
        file_name="Facturas_a_emitir_MODELO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="fe_dl_modelo",
        help="Una fila por comprobante. Mismo ID factura = varios ítems. EJEMPLO no se emite en producción.",
    )


def _panel_resultado_previo() -> None:
    ultimo = st.session_state.get("fe_ultimo")
    if not ultimo:
        return
    st.success("Emisión terminada. Los CAE quedaron guardados en la base del estudio.")
    _tabla(ultimo.get("filas") or [])
    _descargas(ultimo.get("pdfs") or [], ultimo.get("excel"), "emitido", "Comprobantes emitidos")
    if st.button("Ocultar este resultado", key="fe_ocultar_ultimo"):
        st.session_state.pop("fe_ultimo", None)
        st.rerun()


def _panel_borrador(borrador: dict, env: str, prod_ok: bool) -> None:
    st.markdown("###### Borrador")
    st.caption(
        "Todavía no se envió nada a ARCA. Revisá la tabla y el PDF. "
        "Confirmar emite este borrador: si cambiás los datos de abajo, volvé a validar."
    )
    filas = borrador.get("filas") or []
    _tabla(filas)
    _descargas(borrador.get("pdfs") or [], borrador.get("salida"), "borrador", "PDFs borrador")
    listos = [f for f in filas if str(f.get("Estado") or "").startswith("OK-DRYRUN")]
    rechazados = [f for f in filas if "RECHAZADO" in str(f.get("Estado") or "") or str(f.get("Estado") or "").startswith("ERROR")]
    if rechazados:
        st.warning(f"{len(rechazados)} comprobante(s) con error: esos no se envían.")
    if not listos:
        st.error("No hay comprobantes listos para emitir.")
        return
    st.info(f"{len(listos)} comprobante(s) listos. Los que ya figuran como YA EMITIDO no se reenvían.")
    revise = st.checkbox("Revisé el borrador y quiero emitirlo", key="fe_revise")
    if env == "prod" and not prod_ok:
        st.caption("Falta la confirmación de producción (casilla + escribir EMITIR).")
    puede = revise and (env == "homo" or prod_ok)
    etiqueta = "Confirmar emisión en PRODUCCIÓN" if env == "prod" else "Confirmar emisión en homologación"
    if st.button(etiqueta, type="primary", disabled=not puede, key="fe_confirmar"):
        _confirmar(borrador, env)


def _confirmar(borrador: dict, env: str) -> None:
    try:
        with st.spinner("Emitiendo en ARCA…"):
            resultado = emitir_excel(
                borrador["entrada"],
                env,
                confirmar_produccion=(env == "prod"),
                incluir_ejemplos=bool(borrador.get("incluir_ejemplos")),
            )
    except (CredencialesError, EmisionAbortada, WSFEError) as ex:
        st.error(str(ex))
        return
    except Exception as ex:
        st.error(f"No se pudo emitir: {ex}")
        return
    st.session_state["fe_ultimo"] = resultado
    st.session_state.pop("fe_borrador", None)
    st.session_state["_fe_reset_revise"] = True
    st.rerun()


def _mem() -> dict:
    mem = st.session_state.get("_fe_mem")
    if not isinstance(mem, dict):
        mem = {
            "razon": "", "domicilio": "", "iibb": "", "inicio": "", "cond": "",
            "pto": 1, "tipo": "A",
            "doc_tipo": "CUIT", "doc_nro": "", "nombre_rec": "", "dom_rec": "",
            "cond_rec": "", "detalle": "", "neto": 0.0, "alicuota": "21%",
            "asoc_tipo": "", "asoc_pto": 0, "asoc_nro": 0, "asoc_fecha": "", "asoc_id": "",
        }
        st.session_state["_fe_mem"] = mem
    return mem


def _gen(nombre: str) -> int:
    return int(st.session_state.get(nombre) or 0)


def _texto(etiqueta: str, mem_key: str, gen: int, clave: str, **kwargs) -> str:
    mem = _mem()
    wkey = f"{clave}_{gen}"
    extra = {}
    if wkey not in st.session_state:
        extra["value"] = str(mem.get(mem_key) or "")
    valor = st.text_input(etiqueta, key=wkey, **extra, **kwargs)
    mem[mem_key] = valor
    return valor


def _opcion(etiqueta: str, mem_key: str, gen: int, clave: str, opciones: list, **kwargs):
    mem = _mem()
    wkey = f"{clave}_{gen}"
    actual = mem.get(mem_key)
    if actual not in opciones:
        actual = opciones[0] if opciones else ""
        mem[mem_key] = actual
        st.session_state.pop(wkey, None)
    extra = {}
    if wkey not in st.session_state:
        extra["index"] = opciones.index(actual) if actual in opciones else 0
    valor = st.selectbox(etiqueta, opciones, key=wkey, **extra, **kwargs)
    mem[mem_key] = valor
    return valor


def _numero(etiqueta: str, mem_key: str, gen: int, clave: str, minimo, maximo, paso, flotante: bool = False):
    mem = _mem()
    wkey = f"{clave}_{gen}"
    extra = {}
    if wkey not in st.session_state:
        bruto = mem.get(mem_key, minimo)
        try:
            extra["value"] = float(bruto) if flotante else int(bruto)
        except (TypeError, ValueError):
            extra["value"] = float(minimo) if flotante else int(minimo)
    kwargs = {"min_value": minimo, "step": paso, "key": wkey}
    if maximo is not None:
        kwargs["max_value"] = maximo
    if flotante:
        kwargs["format"] = "%.2f"
    valor = st.number_input(etiqueta, **kwargs, **extra)
    mem[mem_key] = valor
    return valor


def _aplicar_emisor_guardado(cuit: str) -> None:
    """Si la CUIT es válida y la ficha cambió, completa el formulario una sola vez."""
    if not cuit_valido(cuit):
        return
    emisor = cargar_emisor(cuit) or {}
    firma = f"{cuit}|{emisor.get('actualizado_en') or ''}"
    if st.session_state.get("_fe_aplicado") == firma:
        return
    st.session_state["_fe_aplicado"] = firma
    mem = _mem()
    campos = campos_desde_emisor(emisor)
    if not campos:
        if not st.session_state.get("_fe_cargo"):
            return
        mem["razon"] = ""
        mem["domicilio"] = ""
        mem["iibb"] = ""
        mem["inicio"] = ""
        mem["cond"] = ""
        mem["pto"] = 1
        mem["tipo"] = "A"
        st.session_state["_fe_cargo"] = ""
    else:
        mem["razon"] = campos.get("razon") or ""
        mem["domicilio"] = campos.get("domicilio") or ""
        mem["iibb"] = campos.get("iibb") or ""
        mem["inicio"] = campos.get("inicio") or ""
        mem["cond"] = campos.get("cond") or ""
        if "pto" in campos:
            mem["pto"] = campos["pto"]
        if campos.get("tipo"):
            mem["tipo"] = campos["tipo"]
        st.session_state["_fe_cargo"] = cuit
    st.session_state["_fe_gen_emisor"] = _gen("_fe_gen_emisor") + 1


def _on_elegir_asociado() -> None:
    etiq = str(st.session_state.get("fe_asoc_lista") or "")
    row = (st.session_state.get("_fe_asoc_map") or {}).get(etiq)
    if not row:
        return
    mem = _mem()
    for clave, valor in campos_desde_asociado(row).items():
        if valor is None:
            continue
        mem[clave] = valor
    st.session_state["_fe_gen_rec"] = _gen("_fe_gen_rec") + 1


def _formulario_manual(env: str) -> None:
    st.markdown("###### Comprobante")
    cuit_txt = st.text_input("CUIT emisor (cliente)", key="fe_cuit", placeholder="11 dígitos, sin guiones")
    aviso = mensaje_cuit(cuit_txt)
    if aviso:
        st.warning(aviso)
    cuit = solo_digitos(cuit_txt)
    _aplicar_emisor_guardado(cuit)
    if st.session_state.get("_fe_cargo") == cuit and cuit_valido(cuit):
        st.caption("Se cargaron los datos guardados de esta CUIT.")
    gen_e = _gen("_fe_gen_emisor")
    gen_r = _gen("_fe_gen_rec")
    c1, c2 = st.columns(2)
    with c1:
        razon = _texto("Razón social", "razon", gen_e, "fe_razon")
        domicilio = _texto("Domicilio comercial", "domicilio", gen_e, "fe_dom")
        cond_emisor = _opcion(
            "Condición IVA emisor", "cond", gen_e, "fe_cond",
            [""] + list(COND_IVA_EMISOR),
            format_func=lambda v: v or "Elegí la condición",
        )
    with c2:
        iibb = _texto("Ingresos brutos", "iibb", gen_e, "fe_iibb")
        inicio = _texto("Inicio de actividades (dd/mm/aaaa)", "inicio", gen_e, "fe_inicio")
        pto = _numero("Punto de venta (WS)", "pto", gen_e, "fe_pto", 1, 99999, 1)
    if cond_emisor == "Responsable Monotributo":
        st.caption("Monotributo: solo se puede emitir C, nota de crédito C o nota de débito C.")
    tipos = tipos_para_condicion(cond_emisor)
    if _mem().get("tipo") not in tipos:
        _mem()["tipo"] = "C" if cond_emisor == "Responsable Monotributo" else tipos[0]
        st.session_state.pop(f"fe_tipo_{gen_e}", None)
    c3, c4, c5 = st.columns(3)
    with c3:
        tipo = _opcion("Tipo de comprobante", "tipo", gen_e, "fe_tipo", tipos)
    with c4:
        concepto = st.selectbox("Concepto", list(CONCEPTOS), key="fe_concepto")
    with c5:
        fecha = st.date_input("Fecha del comprobante", key="fe_fecha")
    cond_venta = st.selectbox("Condición de venta", COND_VENTA, key="fe_cond_venta")
    id_factura = st.text_input("ID factura (opcional, para agrupar o para que una NC lo referencie)", key="fe_id")

    st.markdown("###### Receptor")
    letra = str(tipo)[-1:] if tipo else "C"
    conds = [desc for _cod, (desc, clases) in COND_IVA_RECEPTOR.items() if letra in clases or letra not in "ABC"]
    if _mem().get("cond_rec") not in conds:
        _mem()["cond_rec"] = conds[0]
        st.session_state.pop(f"fe_cond_rec_{gen_r}", None)
    if str(tipo).startswith(("NC", "ND")):
        _elegir_asociado(cuit, env, str(tipo))
    r1, r2 = st.columns(2)
    with r1:
        doc_tipo = _opcion("Documento", "doc_tipo", gen_r, "fe_doc_tipo", list(DOC_TIPOS))
        doc_nro = _texto("Número de documento", "doc_nro", gen_r, "fe_doc_nro")
        if doc_tipo in ("CUIT", "CUIL"):
            aviso_doc = mensaje_cuit(doc_nro)
            if aviso_doc:
                st.warning(aviso_doc)
        cond_rec = _opcion("Condición IVA receptor (RG 5616)", "cond_rec", gen_r, "fe_cond_rec", conds)
    with r2:
        nombre_rec = _texto("Nombre / razón social", "nombre_rec", gen_r, "fe_nombre_rec")
        dom_rec = _texto("Domicilio", "dom_rec", gen_r, "fe_dom_rec")

    st.markdown("###### Ítem")
    detalle = _texto("Detalle", "detalle", gen_r, "fe_detalle")
    if str(tipo).endswith("C") and _mem().get("alicuota") != "No corresponde (C)":
        _mem()["alicuota"] = "No corresponde (C)"
        st.session_state.pop(f"fe_alic_{gen_r}", None)
    i1, i2 = st.columns(2)
    with i1:
        neto = _numero("Neto (A/B) o importe (C)", "neto", gen_r, "fe_neto", 0.0, None, 0.01, flotante=True)
    with i2:
        alicuota = _opcion("Alícuota IVA", "alicuota", gen_r, "fe_alic", list(ALICUOTAS))
    if str(tipo).endswith("C"):
        st.caption("En C la alícuota tiene que ser «No corresponde (C)». El importe es el total.")

    fsd = fsh = fvto = None
    if concepto != "Productos":
        st.markdown("###### Servicio")
        s1, s2, s3 = st.columns(3)
        with s1:
            fsd = st.date_input("Desde", key="fe_fsd")
        with s2:
            fsh = st.date_input("Hasta", key="fe_fsh")
        with s3:
            fvto = st.date_input("Vencimiento de pago", key="fe_fvto")
    elif str(tipo).startswith("FCE"):
        fvto = st.date_input("Vencimiento de pago (FCE)", key="fe_fvto_fce")

    cbu = fce_transf = None
    if str(tipo).startswith("FCE"):
        st.markdown("###### Factura de crédito")
        cbu = st.text_input("CBU del emisor (22 dígitos)", key="fe_cbu")
        fce_transf = st.selectbox(
            "Transferencia",
            ["SCA - Transferencia al Sistema de Circulación Abierta", "ADC - Agente de Depósito Colectivo"],
            key="fe_fce_transf",
        )

    asoc_id = asoc_tipo = asoc_pto = asoc_nro = asoc_fecha = None
    if str(tipo).startswith(("NC", "ND")):
        st.markdown("###### Comprobante asociado")
        st.caption("Elegí uno ya emitido o completá tipo, punto y número. No uses las dos cosas a la vez.")
        asoc_id = _texto("ID factura asociado (si está en este lote)", "asoc_id", gen_r, "fe_asoc_id")
        a1, a2, a3, a4 = st.columns(4)
        with a1:
            asoc_tipo = _opcion("Tipo asociado", "asoc_tipo", gen_r, "fe_asoc_tipo", [""] + list(TIPOS_CBTE))
        with a2:
            asoc_pto = _numero("Punto de venta", "asoc_pto", gen_r, "fe_asoc_pto", 0, 99998, 1)
        with a3:
            asoc_nro = _numero("Número", "asoc_nro", gen_r, "fe_asoc_nro", 0, 99999999, 1)
        with a4:
            asoc_fecha = _texto("Fecha (dd/mm/aaaa)", "asoc_fecha", gen_r, "fe_asoc_fecha")

    if "fe_recordar" not in st.session_state:
        st.session_state["fe_recordar"] = True
    recordar = st.checkbox("Guardar estos datos de emisor para la CUIT", key="fe_recordar")
    if st.button("Validar y armar borrador", type="primary", key="fe_validar_manual"):
        aviso_emit = mensaje_cuit(cuit_txt)
        if aviso_emit:
            st.error(aviso_emit)
            return
        fila = {
            "id": (id_factura or "").strip() or None,
            "cuit_emisor": cuit,
            "razon_emisor": razon,
            "dom_emisor": domicilio,
            "iibb_emisor": iibb,
            "inicio_act": (inicio or "").strip() or None,
            "cond_iva_emisor": cond_emisor,
            "pto_vta": int(pto),
            "tipo": tipo,
            "concepto": concepto,
            "fecha": fecha,
            "cond_venta": cond_venta,
            "doc_tipo": doc_tipo,
            "doc_nro": doc_nro,
            "nombre_rec": nombre_rec,
            "dom_rec": dom_rec,
            "cond_iva": cond_rec,
            "detalle": detalle,
            "neto": neto,
            "alicuota": alicuota,
            "fsd": fsd,
            "fsh": fsh,
            "fvto": fvto,
            "cbu": cbu,
            "fce_transf": fce_transf,
            "asoc_id": (asoc_id or "").strip() or None,
            "asoc_tipo": asoc_tipo or None,
            "asoc_pto": int(asoc_pto) if asoc_pto else None,
            "asoc_nro": int(asoc_nro) if asoc_nro else None,
            "asoc_fecha": (asoc_fecha or "").strip() or None,
        }
        if recordar and cuit_valido(fila["cuit_emisor"]):
            _guardar_emisor_form(fila)
        _validar_y_guardar(excel_desde_filas([fila]), env, incluir_ejemplos=False)


def _elegir_asociado(cuit: str, env: str, tipo: str) -> None:
    filas = candidatos_asociados(listar_emisiones(cuit, env, limite=500), tipo) if cuit_valido(cuit) else []
    if not filas:
        st.caption("No hay comprobantes aprobados de esta CUIT en este ambiente para asociar.")
        st.session_state["_fe_asoc_map"] = {}
        return
    mapa = {etiqueta_asociado(row): row for row in filas}
    st.session_state["_fe_asoc_map"] = mapa
    st.selectbox(
        "Usar un comprobante ya emitido",
        [""] + list(mapa),
        key="fe_asoc_lista",
        on_change=_on_elegir_asociado,
        format_func=lambda v: v or "Elegí una factura",
    )


def _carga_excel(env: str) -> None:
    archivo = st.file_uploader("Excel con la hoja Facturas", type=["xlsx"], key="fe_upload")
    incluir = False
    if env == "homo":
        incluir = st.checkbox(
            "Incluir filas marcadas EJEMPLO (solo homologación; nunca van a producción)",
            key="fe_incluir_ejemplos",
        )
    else:
        st.caption("En producción las filas EJEMPLO se ignoran siempre.")
    if st.button("Validar Excel y armar borrador", type="primary", key="fe_validar_excel"):
        if archivo is None:
            st.error("Subí el Excel.")
            return
        _validar_y_guardar(archivo.getvalue(), env, incluir_ejemplos=incluir)


def _validar_y_guardar(excel_bytes: bytes, env: str, incluir_ejemplos: bool) -> None:
    try:
        with st.spinner("Validando…"):
            resultado = validar_excel(excel_bytes, env, incluir_ejemplos=incluir_ejemplos)
    except EmisionAbortada as ex:
        st.error(str(ex))
        return
    except Exception as ex:
        st.error(f"No se pudo leer o validar: {ex}")
        return
    if not resultado.get("filas"):
        st.warning(
            "No había comprobantes para procesar. "
            "Si el Excel tiene filas EJEMPLO, en homologación marcá la casilla para incluirlas."
        )
        return
    st.session_state["fe_borrador"] = {
        "entrada": excel_bytes,
        "salida": resultado["excel"],
        "filas": resultado["filas"],
        "pdfs": resultado["pdfs"],
        "ambiente": env,
        "incluir_ejemplos": incluir_ejemplos and env == "homo",
        "resumen": resultado["resumen"],
    }
    st.session_state["_fe_reset_revise"] = True
    st.rerun()


def _prefill_emisor(cuit: str) -> None:
    if len(cuit) != 11 or st.session_state.get("_fe_prefijado") == cuit:
        return
    st.session_state["_fe_prefijado"] = cuit
    emisor = cargar_emisor(cuit)
    if not emisor:
        return
    st.session_state["fe_razon"] = emisor.get("razon_social") or ""
    st.session_state["fe_dom"] = emisor.get("domicilio") or ""
    st.session_state["fe_iibb"] = emisor.get("iibb") or ""
    st.session_state["fe_inicio"] = emisor.get("inicio_actividades") or ""
    cond = emisor.get("condicion_iva") or ""
    if cond in COND_IVA_EMISOR:
        st.session_state["fe_cond_emisor"] = cond


def _guardar_emisor_form(fila: dict) -> None:
    inicio = fila.get("inicio_act")
    if isinstance(inicio, (dt.date, dt.datetime)):
        inicio = inicio.strftime("%d/%m/%Y") if not isinstance(inicio, dt.datetime) else inicio.date().strftime("%d/%m/%Y")
    try:
        if inicio:
            inicio = a_fecha(inicio).strftime("%d/%m/%Y")
    except ValueError:
        pass
    guardar_emisor(
        fila["cuit_emisor"],
        razon_social=str(fila.get("razon_emisor") or ""),
        domicilio=str(fila.get("dom_emisor") or ""),
        condicion_iva=str(fila.get("cond_iva_emisor") or ""),
        iibb=str(fila.get("iibb_emisor") or ""),
        inicio_actividades=str(inicio or ""),
        pto_vta=str(int(fila.get("pto_vta") or 0) or ""),
        tipo=str(fila.get("tipo") or ""),
    )


def _tab_consulta(env: str, prod_ok: bool) -> None:
    st.markdown("###### Puntos de venta y último número")
    cuit = st.text_input("CUIT", key="fe_cons_cuit", placeholder="CUIT representada")
    if st.button("Listar puntos de venta WS", key="fe_btn_puntos"):
        if not _puede_operar(env, prod_ok) or not _cuit_listo(cuit):
            return
        try:
            with st.spinner("Consultando ARCA…"):
                data = listar_puntos_venta(cuit, env)
        except (CredencialesError, WSFEError) as ex:
            st.error(str(ex))
        except Exception as ex:
            st.error(f"No se pudo consultar: {ex}")
        else:
            st.session_state["fe_puntos"] = data
            st.session_state["fe_puntos_cuit"] = solo_digitos(cuit)
            st.session_state["fe_puntos_env"] = env
    data = st.session_state.get("fe_puntos")
    if data and st.session_state.get("fe_puntos_env") == env:
        puntos = data.get("puntos") or []
        if data.get("omitidos"):
            st.caption(f"Se omitieron {data['omitidos']} punto(s) de baja o bloqueados.")
        if not puntos:
            st.info("ARCA no devolvió puntos de venta habilitados para Web Services.")
        else:
            st.dataframe(pd.DataFrame(puntos), hide_index=True, use_container_width=True)
            nros = [p["Punto de venta"] for p in puntos]
            pto = st.selectbox("Punto de venta", nros, key="fe_cons_pto")
            incluir_fce = st.checkbox("Incluir FCE", key="fe_incluir_fce")
            if st.button("Último número por tipo", key="fe_btn_ultimos"):
                if _puede_operar(env, prod_ok) and _cuit_listo(cuit):
                    try:
                        with st.spinner("Consultando numeración…"):
                            tabla = ultimos_numeros(cuit, env, int(pto), incluir_fce=incluir_fce)
                        st.dataframe(pd.DataFrame(tabla), hide_index=True, use_container_width=True)
                    except (CredencialesError, WSFEError) as ex:
                        st.error(str(ex))
                    except Exception as ex:
                        st.error(str(ex))

    st.divider()
    st.markdown("###### Comprobante ya emitido")
    q1, q2, q3 = st.columns(3)
    with q1:
        tipo = st.selectbox("Tipo", list(TIPOS_CBTE), key="fe_q_tipo")
    with q2:
        pto_q = st.number_input("Punto de venta", min_value=1, max_value=99999, step=1, key="fe_q_pto")
    with q3:
        nro_q = st.number_input("Número", min_value=1, step=1, key="fe_q_nro")
    if st.button("Consultar en ARCA y armar PDF", key="fe_btn_consultar"):
        if _puede_operar(env, prod_ok) and _cuit_listo(cuit):
            try:
                with st.spinner("Consultando el comprobante…"):
                    r = consultar_y_pdf(cuit, env, int(pto_q), tipo, int(nro_q))
            except (CredencialesError, WSFEError, ValueError) as ex:
                st.error(str(ex))
            except Exception as ex:
                st.error(f"No se pudo consultar: {ex}")
            else:
                if not r.get("ok"):
                    st.error(" ".join(r.get("mensajes") or []))
                else:
                    st.dataframe(pd.DataFrame([r["resumen"]]), hide_index=True, use_container_width=True)
                    if r.get("mensajes"):
                        st.caption(" | ".join(r["mensajes"]))
                    st.caption(f"PDF armado con {r.get('origen_pdf')}.")
                    nombre = f"{solo_digitos(cuit)}_{tipo.replace(' ', '')}_{int(pto_q):05d}-{int(nro_q):08d}.pdf"
                    st.download_button(
                        "Descargar PDF",
                        data=r["pdf"],
                        file_name=nombre,
                        mime="application/pdf",
                        key="fe_dl_consulta",
                    )

    st.divider()
    st.markdown("###### Categoría de monotributo")
    st.caption("Usa el padrón de constancia de inscripción. Si el servicio no está autorizado al certificado, se avisa y la facturación sigue funcionando.")
    if st.button("Consultar categoría", key="fe_btn_mono"):
        if _puede_operar(env, prod_ok) and _cuit_listo(cuit):
            try:
                with st.spinner("Consultando padrón…"):
                    cat = categoria_monotributo(env, cuit)
            except CredencialesError as ex:
                st.error(str(ex))
            except Exception as ex:
                st.error(str(ex))
            else:
                if cat.get("ok") and cat.get("categoria"):
                    st.success(cat.get("mensaje") or "")
                    if cat.get("razon"):
                        st.caption(cat["razon"])
                elif cat.get("no_autorizado"):
                    st.warning(cat.get("mensaje") or "")
                elif cat.get("ok"):
                    st.info(cat.get("mensaje") or "")
                else:
                    st.error(cat.get("mensaje") or "Sin respuesta del padrón.")

    st.divider()
    _historial(env)


def _historial(env: str) -> None:
    st.markdown("###### Comprobantes guardados")
    st.caption("Quedan en la base del estudio (no se pierden si Streamlit se reinicia).")
    cuit_f = solo_digitos(st.text_input("Filtrar por CUIT emisor", key="fe_hist_cuit"))
    filas = listar_emisiones(cuit_f, env, limite=100)
    if not filas:
        st.caption("Todavía no hay comprobantes guardados en este ambiente.")
        return
    vista = []
    for row in filas:
        vista.append({
            "Id": row["id"],
            "Fecha proceso": row.get("fecha_hora") or "",
            "CUIT": row.get("cuit_emisor") or "",
            "Tipo": row.get("tipo") or "",
            "Pto": row.get("pto_vta") or "",
            "Número": row.get("numero") or "",
            "CAE": row.get("cae") or "",
            "Vto. CAE": row.get("vto_cae") or "",
            "Estado": row.get("estado") or "",
            "Total": row.get("total") if row.get("total") is not None else "",
        })
    st.dataframe(pd.DataFrame(vista), hide_index=True, use_container_width=True)
    ids = [int(r["id"]) for r in filas if r.get("payload_json")]
    if not ids:
        return
    elegido = st.selectbox("Regenerar PDF de un guardado", ids, key="fe_hist_id")
    pdf = pdf_de_emision(int(elegido))
    if pdf:
        st.download_button(
            "Descargar PDF guardado",
            data=pdf,
            file_name=f"comprobante_{elegido}.pdf",
            mime="application/pdf",
            key="fe_dl_hist",
        )


def _tab_emisor() -> None:
    st.markdown("###### Datos que salen en el PDF")
    st.caption("Se guardan por CUIT en la base del estudio. Si el Excel trae la celda vacía, se completan solos.")
    guardados = listar_emisores()
    opciones = [""] + [f"{g['cuit']} — {g.get('razon_social') or ''}".strip() for g in guardados]
    elegido = st.selectbox("Cargar uno guardado", opciones, key="fe_emisor_lista")
    if elegido and st.session_state.get("_fe_emisor_lista_aplicada") != elegido:
        st.session_state["_fe_emisor_lista_aplicada"] = elegido
        cuit_sel = solo_digitos(elegido.split("—")[0])
        em = cargar_emisor(cuit_sel) or {}
        st.session_state["fe_em_cuit"] = cuit_sel
        st.session_state["fe_em_razon"] = em.get("razon_social") or ""
        st.session_state["fe_em_dom"] = em.get("domicilio") or ""
        st.session_state["fe_em_cond"] = em.get("condicion_iva") or ""
        st.session_state["fe_em_iibb"] = em.get("iibb") or ""
        st.session_state["fe_em_inicio"] = em.get("inicio_actividades") or ""
        st.rerun()
    cuit = st.text_input("CUIT", key="fe_em_cuit")
    aviso_cuit = mensaje_cuit(cuit)
    if aviso_cuit:
        st.warning(aviso_cuit)
    razon = st.text_input("Razón social", key="fe_em_razon")
    domicilio = st.text_input("Domicilio comercial", key="fe_em_dom")
    if "fe_em_cond" not in st.session_state:
        st.session_state["fe_em_cond"] = ""
    cond = st.selectbox(
        "Condición frente al IVA",
        [""] + list(COND_IVA_EMISOR),
        format_func=lambda v: v or "Elegí la condición",
        key="fe_em_cond",
    )
    iibb = st.text_input("Ingresos brutos", key="fe_em_iibb")
    inicio = st.text_input("Inicio de actividades (dd/mm/aaaa)", key="fe_em_inicio")
    if st.button("Guardar datos del emisor", type="primary", key="fe_em_guardar"):
        digits = solo_digitos(cuit)
        if not digits:
            st.error("Ingresá la CUIT.")
            return
        aviso_guardar = mensaje_cuit(digits)
        if aviso_guardar:
            st.error(aviso_guardar)
            return
        texto_inicio = (inicio or "").strip()
        if texto_inicio:
            try:
                texto_inicio = a_fecha(texto_inicio).strftime("%d/%m/%Y")
            except ValueError:
                st.error("La fecha de inicio no se entiende. Usá dd/mm/aaaa.")
                return
        guardar_emisor(digits, razon, domicilio, cond, iibb, texto_inicio)
        st.success(f"Guardado para el CUIT {digits}.")


def _cuit_listo(cuit: str) -> bool:
    if len(solo_digitos(cuit)) != 11:
        st.error("Ingresá la CUIT con 11 dígitos.")
        return False
    return True


def _puede_operar(env: str, prod_ok: bool) -> bool:
    if env == "prod" and not prod_ok:
        st.error("Producción está bloqueada hasta marcar la casilla y escribir EMITIR.")
        return False
    return True


def _tabla(filas: list[dict]) -> None:
    if not filas:
        return
    cols = [c for c in _COLUMNAS_VISTA if any(c in f for f in filas)]
    df = pd.DataFrame(filas)
    st.dataframe(df[cols] if cols else df, hide_index=True, use_container_width=True)


def _descargas(pdfs: list, excel: bytes | None, sufijo: str, titulo_pdf: str) -> None:
    if excel:
        st.download_button(
            "Descargar Excel con hoja Resultado",
            data=excel,
            file_name=f"Facturas_resultado_{sufijo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"fe_dl_xlsx_{sufijo}",
        )
    if not pdfs:
        return
    if len(pdfs) == 1:
        nombre, data = pdfs[0]
        st.download_button(f"Descargar PDF ({titulo_pdf})", data=data, file_name=nombre, mime="application/pdf", key=f"fe_dl_pdf_{sufijo}")
        return
    st.download_button(
        f"Descargar {len(pdfs)} PDFs (zip)",
        data=zip_pdfs(pdfs),
        file_name=f"comprobantes_{sufijo}.zip",
        mime="application/zip",
        key=f"fe_dl_zip_{sufijo}",
    )
    for i, (nombre, data) in enumerate(pdfs[:8]):
        st.download_button(nombre, data=data, file_name=nombre, mime="application/pdf", key=f"fe_dl_pdf_{sufijo}_{i}")
