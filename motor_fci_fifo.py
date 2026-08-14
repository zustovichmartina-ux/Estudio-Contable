# -*- coding: utf-8 -*-
"""
Motor FIFO genérico para Fondos Comunes de Inversión (cualquier cliente / banco).

Entrada obligatoria:
  - Extracto del **último mes del ejercicio anterior** → saldo inicial
  - Extractos de los **12 meses del ejercicio** → movimientos + VC de cierre

Salida:
  - Excel de análisis FCI (intereses por lote, tenencia, fórmula de control)
  - El mayor cierra contra: cuotas finales × VC del último día del ejercicio
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Any, Iterable

import pandas as pd

from extraccion_fci import _leer_entrada, parsear_pdf_fci
from inversiones import aplicar_fifo, identificar_movimientos
from inversiones_catalogo import TIPO_FCI
from inversiones_fci_formato import (
    consolidar_lotes_abiertos,
    exportar_analisis_fci_excel,
    resultado_por_tenencia,
)

_MESES_ES = (
    "",
    "Ene",
    "Feb",
    "Mar",
    "Abr",
    "May",
    "Jun",
    "Jul",
    "Ago",
    "Sep",
    "Oct",
    "Nov",
    "Dic",
)


@dataclass
class ExtractoFciMes:
    anio: int
    mes: int
    periodo: str
    fecha_desde: date | None
    fecha_hasta: date | None
    fondo: str
    banco: str
    archivo: str
    movimientos: list[dict] = field(default_factory=list)
    cuotas_cierre: float | None = None
    vc_cierre: float | None = None
    pesos_cierre: float | None = None
    fuente_posicion: str = ""
    ok: bool = True
    error: str | None = None


@dataclass
class ResultadoEjercicioFci:
    ok: bool
    avisos: list[str] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)
    extractos: list[ExtractoFciMes] = field(default_factory=list)
    meses_faltantes: list[str] = field(default_factory=list)
    saldo_inicial: dict[str, Any] = field(default_factory=dict)
    vc_cierre: float | None = None
    fecha_cierre: date | None = None
    cuotas_finales: float | None = None
    valuacion_extracto: float | None = None
    df_mov: pd.DataFrame | None = None
    df_inicial: pd.DataFrame | None = None
    resultado_fifo: Any = None
    tenencia: dict[str, float] = field(default_factory=dict)
    control: dict[str, float] = field(default_factory=dict)
    excel_bytes: bytes | None = None
    meta: dict[str, Any] = field(default_factory=dict)


def _parse_fecha_txt(raw: str) -> date | None:
    t = (raw or "").strip()
    if not t:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    return None


def _parse_periodo(periodo: str) -> tuple[date | None, date | None]:
    if not periodo:
        return None, None
    partes = re_split_periodo(periodo)
    if len(partes) >= 2:
        return _parse_fecha_txt(partes[0]), _parse_fecha_txt(partes[1])
    if len(partes) == 1:
        d = _parse_fecha_txt(partes[0])
        return d, d
    return None, None


def re_split_periodo(periodo: str) -> list[str]:
    return [p.strip() for p in re.split(r"\s+al\s+", str(periodo), flags=re.I) if p.strip()]


def _clave_mes(anio: int, mes: int) -> tuple[int, int]:
    return (int(anio), int(mes))


def label_mes(anio: int, mes: int) -> str:
    return f"{_MESES_ES[mes]}-{str(anio)[2:]}"


def _label_mes(anio: int, mes: int) -> str:
    return label_mes(anio, mes)


def meses_del_ejercicio(fecha_inicio: date, fecha_cierre: date) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    y, m = fecha_inicio.year, fecha_inicio.month
    fin = (fecha_cierre.year, fecha_cierre.month)
    while (y, m) <= fin:
        out.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def mes_saldo_inicial(fecha_inicio: date) -> tuple[int, int]:
    """Último mes del ejercicio anterior (día anterior al inicio)."""
    d = fecha_inicio - timedelta(days=1)
    return (d.year, d.month)


def _num(val: object, default: float | None = None) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _clasificar_pdf(data: bytes, nombre: str) -> ExtractoFciMes:
    res = parsear_pdf_fci(data, nombre)
    periodo = str(res.get("periodo") or "")
    desde, hasta = _parse_periodo(periodo)
    filas = list(res.get("filas") or [])
    if desde is None and filas:
        fechas = [f.get("Fecha") for f in filas if isinstance(f.get("Fecha"), date)]
        if fechas:
            desde, hasta = min(fechas), max(fechas)
    if hasta is None and desde is not None:
        hasta = desde
    if hasta is None:
        # Intentar mes desde el nombre del archivo (ej. 06-2026, junio2026)
        m = re.search(r"(20\d{2})[-_/]?(\d{2})", nombre)
        if not m:
            m = re.search(r"(\d{2})[-_/](20\d{2})", nombre)
            if m:
                mes, anio = int(m.group(1)), int(m.group(2))
                if 1 <= mes <= 12:
                    hasta = date(anio, mes, 1)
        else:
            anio, mes = int(m.group(1)), int(m.group(2))
            if 1 <= mes <= 12:
                hasta = date(anio, mes, 1)

    anio = hasta.year if hasta else 0
    mes = hasta.month if hasta else 0
    pos = res.get("posicion") or {}
    return ExtractoFciMes(
        anio=anio,
        mes=mes,
        periodo=periodo,
        fecha_desde=desde,
        fecha_hasta=hasta,
        fondo=str(res.get("fondo") or ""),
        banco=str(res.get("banco") or ""),
        archivo=nombre,
        movimientos=filas,
        cuotas_cierre=_num(pos.get("cuotas")),
        vc_cierre=_num(pos.get("valor_cuota")),
        pesos_cierre=_num(pos.get("importe")),
        fuente_posicion=str(pos.get("fuente") or ""),
        ok=bool(res.get("ok")) or bool(filas) or bool(pos.get("cuotas")),
        error=res.get("error"),
    )


def _movimientos_a_df(movs: list[dict], fondo_default: str = "") -> pd.DataFrame:
    filas: list[dict] = []
    for m in movs:
        fecha = m.get("Fecha")
        if not isinstance(fecha, date):
            fecha = _parse_fecha_txt(str(fecha or ""))
        if fecha is None:
            continue
        desc = str(m.get("Descripcion") or "").upper().strip()
        if "SUSCRIP" in desc or desc == "COMPRA":
            tipo = "Suscripcion"
        elif "RESCATE" in desc or desc == "VENTA":
            tipo = "Rescate"
        else:
            continue
        cant = abs(_num(m.get("Cantidad de Cuotas"), 0.0) or 0.0)
        precio = _num(m.get("Valor CC"), 0.0) or 0.0
        monto = abs(_num(m.get("Total"), 0.0) or 0.0)
        if cant <= 0 and precio > 0 and monto > 0:
            cant = round(monto / precio, 4)
        if cant <= 0:
            continue
        if monto <= 0 and precio > 0:
            monto = round(cant * precio, 2)
        fondo = str(m.get("Fondo") or fondo_default or "FCI").strip() or "FCI"
        filas.append(
            {
                "Fecha": fecha.strftime("%d/%m/%Y"),
                "_fecha": fecha,
                "Especie": fondo,
                "Descripcion": desc,
                "Tipo_Operacion": tipo,
                "Cantidad": cant,
                "Precio": precio,
                "Monto_Total": monto,
                "Moneda": "ARS",
                "Tipo_inversion": TIPO_FCI,
                "Archivo": str(m.get("Archivo") or ""),
            }
        )
    if not filas:
        return pd.DataFrame()
    df = pd.DataFrame(filas)
    return identificar_movimientos(df)


def _df_saldo_inicial(
    *,
    fondo: str,
    cuotas: float,
    vc: float,
    pesos: float | None = None,
    origen: str = "Extracto mes anterior",
) -> pd.DataFrame:
    total = round(float(pesos), 2) if pesos not in (None, 0) else round(float(cuotas) * float(vc), 2)
    return pd.DataFrame(
        [
            {
                "Especie": fondo or "FCI",
                "Especie_canonica": fondo or "FCI",
                "Tipo_inversion": TIPO_FCI,
                "Grupo": "FCI",
                "Cantidad": float(cuotas),
                "Costo_Unitario": float(vc),
                "Costo_Total": total,
                "Moneda": "ARS",
                "Origen": origen,
                "Fecha": "01/01/1900",
                "Revisar": "",
                "Confianza": "alta",
            }
        ]
    )


def _control_mayor(
    *,
    saldo_inicial_pesos: float,
    suscripciones: float,
    rescates: float,
    intereses: float,
    tenencia_contabilizar: float,
    valuacion_extracto: float,
) -> dict[str, float]:
    izq = round(
        saldo_inicial_pesos + suscripciones + intereses + tenencia_contabilizar - rescates,
        2,
    )
    return {
        "lado_izquierdo": izq,
        "saldo_final_extracto": round(valuacion_extracto, 2),
        "diferencia": round(izq - valuacion_extracto, 2),
    }


def procesar_ejercicio_fci(
    entradas: Iterable[Any],
    *,
    fecha_inicio: date,
    fecha_cierre: date,
    si_cuotas: float | None = None,
    si_vc: float | None = None,
    si_pesos: float | None = None,
    vc_cierre: float | None = None,
    fondo_forzado: str | None = None,
    cliente: str = "",
) -> ResultadoEjercicioFci:
    """
    Procesa el ejercicio FCI completo.

    Subí sí o sí: mes anterior al inicio (SI) + 12 meses del ejercicio.
    """
    out = ResultadoEjercicioFci(ok=False, fecha_cierre=fecha_cierre)
    if fecha_cierre < fecha_inicio:
        out.errores.append("La fecha de cierre no puede ser anterior al inicio del ejercicio.")
        return out

    extractos: list[ExtractoFciMes] = []
    for item in entradas or []:
        nombre, data = _leer_entrada(item)
        if not str(nombre).lower().endswith(".pdf"):
            out.avisos.append(f"{nombre}: se ignoró (solo PDF de extracto FCI en este motor).")
            continue
        extractos.append(_clasificar_pdf(data, nombre))

    out.extractos = extractos
    if not extractos:
        out.errores.append("No se recibieron extractos PDF.")
        return out

    # Índice por mes (si hay varios del mismo mes, se mergean movimientos; posición = última)
    por_mes: dict[tuple[int, int], ExtractoFciMes] = {}
    for ex in extractos:
        if not ex.anio or not ex.mes:
            out.avisos.append(
                f"{ex.archivo}: no se pudo determinar el mes del extracto "
                f"(periodo='{ex.periodo}'). Nombralo MM-AAAA o revisá el PDF."
            )
            continue
        key = _clave_mes(ex.anio, ex.mes)
        if key not in por_mes:
            por_mes[key] = ex
        else:
            base = por_mes[key]
            base.movimientos.extend(ex.movimientos)
            if ex.cuotas_cierre and ex.vc_cierre:
                base.cuotas_cierre = ex.cuotas_cierre
                base.vc_cierre = ex.vc_cierre
                base.pesos_cierre = ex.pesos_cierre
                base.fuente_posicion = ex.fuente_posicion
            if not base.fondo and ex.fondo:
                base.fondo = ex.fondo
            if not base.banco and ex.banco:
                base.banco = ex.banco

    mes_si = mes_saldo_inicial(fecha_inicio)
    meses_ej = meses_del_ejercicio(fecha_inicio, fecha_cierre)
    faltan = [m for m in meses_ej if m not in por_mes]
    out.meses_faltantes = [_label_mes(a, m) for a, m in faltan]
    if mes_si not in por_mes and not (si_cuotas and si_vc):
        out.errores.append(
            f"Falta el extracto del último mes del ejercicio anterior "
            f"({_label_mes(*mes_si)}) para determinar el saldo inicial, "
            "o cargá cuotas/VC de SI manualmente."
        )
    if faltan:
        out.avisos.append(
            "Meses del ejercicio sin extracto: " + ", ".join(out.meses_faltantes)
        )

    # Fondo / banco
    fondos = [ex.fondo for ex in extractos if ex.fondo]
    bancos = [ex.banco for ex in extractos if ex.banco and ex.banco != "Desconocido"]
    fondo = (fondo_forzado or (max(set(fondos), key=fondos.count) if fondos else "FCI")).strip()
    banco = max(set(bancos), key=bancos.count) if bancos else ""

    # Saldo inicial
    ex_si = por_mes.get(mes_si)
    cuotas_si = si_cuotas
    vc_si = si_vc
    pesos_si = si_pesos
    origen_si = "Manual"
    if ex_si is not None:
        if cuotas_si is None:
            cuotas_si = ex_si.cuotas_cierre
        if vc_si is None:
            vc_si = ex_si.vc_cierre
        if pesos_si is None:
            pesos_si = ex_si.pesos_cierre
        origen_si = f"Extracto {_label_mes(*mes_si)} ({ex_si.archivo})"
        if not fondo_forzado and ex_si.fondo:
            fondo = ex_si.fondo

    if not cuotas_si or not vc_si:
        out.errores.append(
            "No se pudo armar el saldo inicial (faltan cuotas y/o valor cuota). "
            "Completá los campos manuales o revisá el extracto del mes anterior."
        )
        return out

    df_ini = _df_saldo_inicial(
        fondo=fondo,
        cuotas=float(cuotas_si),
        vc=float(vc_si),
        pesos=pesos_si,
        origen=origen_si,
    )
    out.saldo_inicial = {
        "Fondo": fondo,
        "Mes": _label_mes(*mes_si),
        "Cuotas": float(cuotas_si),
        "Valor_cuota": float(vc_si),
        "Pesos": float(df_ini.iloc[0]["Costo_Total"]),
        "Origen": origen_si,
    }
    out.df_inicial = df_ini

    # Movimientos del ejercicio (no incluir mes SI)
    movs: list[dict] = []
    for key in meses_ej:
        ex = por_mes.get(key)
        if not ex:
            continue
        for m in ex.movimientos:
            fecha = m.get("Fecha")
            if isinstance(fecha, date) and (fecha < fecha_inicio or fecha > fecha_cierre):
                continue
            mm = dict(m)
            if not mm.get("Fondo"):
                mm["Fondo"] = fondo
            movs.append(mm)

    df_mov = _movimientos_a_df(movs, fondo_default=fondo)
    out.df_mov = df_mov

    # VC / cuotas de cierre: último mes del ejercicio
    mes_cierre = (fecha_cierre.year, fecha_cierre.month)
    ex_cierre = por_mes.get(mes_cierre)
    vc_fin = vc_cierre
    cuotas_fin = None
    pesos_fin = None
    if ex_cierre is not None:
        if vc_fin is None:
            vc_fin = ex_cierre.vc_cierre
        cuotas_fin = ex_cierre.cuotas_cierre
        pesos_fin = ex_cierre.pesos_cierre
    if vc_fin is None or vc_fin <= 0:
        out.errores.append(
            "Falta el valor cuota de cierre (último extracto del ejercicio). "
            "Cargalo manualmente si el PDF no lo trae."
        )
        return out

    res = aplicar_fifo(df_mov if df_mov is not None else pd.DataFrame(), df_ini)
    out.resultado_fifo = res
    out.avisos.extend(list(res.avisos or []))

    # Cuotas finales: preferir posición del extracto; si no, suma lotes FIFO
    if cuotas_fin is None and res.lotes_abiertos:
        cuotas_fin = round(sum(float(l["Cantidad"]) for l in res.lotes_abiertos), 4)
    if pesos_fin is None and cuotas_fin is not None:
        pesos_fin = round(float(cuotas_fin) * float(vc_fin), 2)

    out.vc_cierre = float(vc_fin)
    out.cuotas_finales = float(cuotas_fin) if cuotas_fin is not None else None
    out.valuacion_extracto = float(pesos_fin) if pesos_fin is not None else None

    # Tenencia analítica (por lotes)
    lotes = consolidar_lotes_abiertos(res.lotes_abiertos, fecha_inicio - timedelta(days=1))
    ten = resultado_por_tenencia(lotes, float(vc_fin)) if lotes else {
        "cuotas": float(cuotas_fin or 0),
        "vc_cierre": float(vc_fin),
        "valuacion": float(pesos_fin or 0),
        "costo_origen": float(df_ini.iloc[0]["Costo_Total"]),
        "resultado_tenencia": 0.0,
    }
    out.tenencia = ten

    # Totales para control (Totales del banco)
    if df_mov is not None and not df_mov.empty:
        sus = float(df_mov.loc[df_mov["Tipo_Operacion"].str.upper().str.contains("SUSCRIP"), "Monto_Total"].sum())
        resc = float(df_mov.loc[df_mov["Tipo_Operacion"].str.upper().str.contains("RESCATE"), "Monto_Total"].sum())
    else:
        sus = resc = 0.0
    intereses = 0.0
    if res.movimientos:
        for m in res.movimientos:
            if str(m.get("Estado")) in {"INTERES", "SIN_STOCK"} and str(m.get("Tipo_Operacion", "")).lower().startswith("resc"):
                intereses += float(m.get("Resultado") or 0)
    intereses = round(intereses, 2)

    si_pesos_val = float(out.saldo_inicial["Pesos"])
    valuacion = float(out.valuacion_extracto or ten.get("valuacion") or 0)
    ten_lotes = float(ten.get("resultado_tenencia") or 0)
    # Ajuste para que la fórmula de control cierre contra el extracto
    subtotal = round(si_pesos_val + sus + intereses + ten_lotes - resc, 2)
    ajuste = round(valuacion - subtotal, 2)
    ten_contab = round(ten_lotes + ajuste, 2)
    out.control = {
        "saldo_inicial": si_pesos_val,
        "suscripciones": round(sus, 2),
        "intereses_ganados": intereses,
        "tenencia_lotes": ten_lotes,
        "ajuste_redondeo": ajuste,
        "tenencia_a_contabilizar": ten_contab,
        "rescates": round(resc, 2),
        **_control_mayor(
            saldo_inicial_pesos=si_pesos_val,
            suscripciones=sus,
            rescates=resc,
            intereses=intereses,
            tenencia_contabilizar=ten_contab,
            valuacion_extracto=valuacion,
        ),
    }

    meta = {
        "cliente": cliente or "",
        "fondo": fondo,
        "banco": banco,
        "ejercicio": f"{fecha_inicio.strftime('%d/%m/%Y')} al {fecha_cierre.strftime('%d/%m/%Y')}",
        "nota": (
            f"FCI FIFO · {fondo}"
            + (f" · {banco}" if banco else "")
            + (f" · {cliente}" if cliente else "")
        ),
    }
    out.meta = meta
    out.excel_bytes = exportar_analisis_fci_excel(
        df_mov if df_mov is not None else pd.DataFrame(),
        res.lotes_abiertos,
        df_ini,
        vc_cierre=float(vc_fin),
        fecha_cierre=fecha_cierre,
        meta=meta,
    )
    out.ok = True
    if out.errores:
        # Errores duros ya retornaron; acá solo avisos
        pass
    return out


def cuadro_cobertura_meses(
    fecha_inicio: date,
    fecha_cierre: date,
    extractos: list[ExtractoFciMes],
) -> pd.DataFrame:
    """Tabla para la UI: qué meses están y cuáles faltan."""
    mes_si = mes_saldo_inicial(fecha_inicio)
    keys = [mes_si] + meses_del_ejercicio(fecha_inicio, fecha_cierre)
    por = {(e.anio, e.mes): e for e in extractos if e.anio and e.mes}
    filas = []
    for i, key in enumerate(keys):
        ex = por.get(key)
        rol = "Saldo inicial (ejercicio anterior)" if i == 0 else "Ejercicio"
        filas.append(
            {
                "Mes": _label_mes(*key),
                "Rol": rol,
                "Estado": "OK" if ex else "FALTA",
                "Archivo": ex.archivo if ex else "",
                "Movimientos": len(ex.movimientos) if ex else 0,
                "Cuotas cierre": ex.cuotas_cierre if ex else None,
                "VC cierre": ex.vc_cierre if ex else None,
                "Banco": ex.banco if ex else "",
                "Fondo": ex.fondo if ex else "",
            }
        )
    return pd.DataFrame(filas)
