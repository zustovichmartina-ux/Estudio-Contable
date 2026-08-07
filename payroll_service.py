"""Motor de liquidación de sueldos."""

from __future__ import annotations

from datetime import date
from typing import Any

from cct_escalas import resolver_basico_cct

DEFAULTS = {
    "antiguedadPorAnioPct": 0.01,  # Comercio CCT: 1% por año
    "presentismoDivisor": 12,
    "horasMensuales": 200,
    "horasExtras50Multiplicador": 1.5,
    "diasMes": 30,
    "jubilacionPct": 0.11,
    "pamiPct": 0.03,
    "obraSocialPct": 0.03,
    "aporteSindicalPct": 0.02,
}


def _round2(n: float) -> float:
    return round(float(n) + 1e-12, 2)


def _pct(reglas: dict[str, Any], key: str) -> float:
    v = reglas.get(key, DEFAULTS[key])
    try:
        n = float(v)
    except (TypeError, ValueError):
        return float(DEFAULTS[key])
    if n > 1:
        return n / 100.0
    return n


def _num(reglas: dict[str, Any], key: str) -> float:
    v = reglas.get(key, DEFAULTS[key])
    try:
        n = float(v)
        return n if n > 0 else float(DEFAULTS[key])
    except (TypeError, ValueError):
        return float(DEFAULTS[key])


def periodo_actual(hoy: date | None = None) -> str:
    d = hoy or date.today()
    return f"{d.year:04d}-{d.month:02d}"


def calcular_liquidacion(
    empleado: dict[str, Any],
    novedad: dict[str, Any],
    cct_reglas: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Calcula liquidación. El básico sale de la escala del CCT por categoría
    (salvo que el legajo tenga básico > 0 y usarEscalaCct=False).
    """
    reglas = dict(cct_reglas or {})
    categoria = str(empleado.get("categoria") or "")
    basico_legajo = float(empleado.get("sueldo_basico") or empleado.get("sueldoBasico") or 0)
    basico_cct, cat_canon = resolver_basico_cct(categoria, reglas)
    forzar_escala = bool(reglas.get("usarEscalaCct", True))

    if forzar_escala and basico_cct is not None:
        basico = float(basico_cct)
        fuente_basico = f"cct:{cat_canon}"
    elif basico_legajo > 0:
        basico = basico_legajo
        fuente_basico = "legajo"
    elif basico_cct is not None:
        basico = float(basico_cct)
        fuente_basico = f"cct:{cat_canon}"
    else:
        basico = 0.0
        fuente_basico = "sin_dato"

    anios = int(empleado.get("antiguedad_anios") or empleado.get("antiguedadAnios") or 0)
    dias_ausencia = max(0, int(novedad.get("dias_ausencia") or novedad.get("diasAusencia") or 0))
    he50 = max(0.0, float(novedad.get("horas_extras_50") or novedad.get("horasExtras50") or 0))
    no_rem = max(
        0.0,
        float(novedad.get("no_remunerativo_extra") or novedad.get("noRemunerativoExtra") or 0),
    )

    ant_pct = _pct(reglas, "antiguedadPorAnioPct")
    divisor_presentismo = _num(reglas, "presentismoDivisor")
    horas_mes = _num(reglas, "horasMensuales")
    he_mult = _num(reglas, "horasExtras50Multiplicador")
    dias_mes = _num(reglas, "diasMes")

    antiguedad = _round2(basico * (anios * ant_pct))
    base = basico + antiguedad
    presentismo = _round2(base / divisor_presentismo) if dias_ausencia == 0 else 0.0
    valor_hora = _round2(base / horas_mes)
    horas_extras_50 = _round2(valor_hora * he_mult * he50)
    descuento_ausencia = _round2((base / dias_mes) * dias_ausencia)

    total_remunerativo = _round2(
        basico + antiguedad + presentismo + horas_extras_50 - descuento_ausencia
    )
    total_no_remunerativo = _round2(no_rem)

    jubilacion = _round2(total_remunerativo * _pct(reglas, "jubilacionPct"))
    pami = _round2(total_remunerativo * _pct(reglas, "pamiPct"))
    obra_social = _round2(total_remunerativo * _pct(reglas, "obraSocialPct"))
    aporte_sindical = _round2(total_remunerativo * _pct(reglas, "aporteSindicalPct"))
    total_descuentos = _round2(jubilacion + pami + obra_social + aporte_sindical)
    neto = _round2(total_remunerativo + total_no_remunerativo - total_descuentos)

    desc_basico = (
        f"Sueldo básico CCT ({cat_canon or categoria})"
        if fuente_basico.startswith("cct:")
        else "Sueldo básico"
    )

    conceptos: list[dict[str, Any]] = [
        {
            "codigo": "1",
            "descripcion": desc_basico,
            "tipo": "remunerativo",
            "importe": _round2(basico),
        },
        {
            "codigo": "4",
            "descripcion": f"Antigüedad ({anios} años × {ant_pct * 100:.0f}%)",
            "tipo": "remunerativo",
            "importe": antiguedad,
        },
    ]
    if presentismo > 0:
        conceptos.append(
            {
                "codigo": "9",
                "descripcion": "Presentismo",
                "tipo": "remunerativo",
                "importe": presentismo,
            }
        )
    if horas_extras_50 > 0:
        conceptos.append(
            {
                "codigo": "6",
                "descripcion": f"Horas extras 50% ({he50} hs)",
                "tipo": "remunerativo",
                "importe": horas_extras_50,
            }
        )
    if descuento_ausencia > 0:
        conceptos.append(
            {
                "codigo": "20",
                "descripcion": f"Inasistencias ({dias_ausencia} días)",
                "tipo": "descuento",
                "importe": descuento_ausencia,
            }
        )
    if total_no_remunerativo > 0:
        conceptos.append(
            {
                "codigo": "50018",
                "descripcion": "No remunerativo extra",
                "tipo": "no_remunerativo",
                "importe": total_no_remunerativo,
            }
        )
    conceptos.extend(
        [
            {
                "codigo": "20000",
                "descripcion": "Jubilación (11%)",
                "tipo": "descuento",
                "importe": jubilacion,
            },
            {
                "codigo": "20001",
                "descripcion": "Ley 19032 / PAMI (3%)",
                "tipo": "descuento",
                "importe": pami,
            },
            {
                "codigo": "20002",
                "descripcion": "Obra social (3%)",
                "tipo": "descuento",
                "importe": obra_social,
            },
            {
                "codigo": "20004",
                "descripcion": "Sindicato (2%)",
                "tipo": "descuento",
                "importe": aporte_sindical,
            },
        ]
    )

    return {
        "antiguedad": antiguedad,
        "presentismo": presentismo,
        "valor_hora": valor_hora,
        "horas_extras_50": horas_extras_50,
        "descuento_ausencia": descuento_ausencia,
        "total_remunerativo": total_remunerativo,
        "total_no_remunerativo": total_no_remunerativo,
        "descuentos": {
            "jubilacion": jubilacion,
            "pami": pami,
            "obra_social": obra_social,
            "aporte_sindical": aporte_sindical,
        },
        "total_descuentos": total_descuentos,
        "neto_a_percibir": neto,
        "conceptos": conceptos,
        "sueldo_basico_usado": _round2(basico),
        "fuente_basico": fuente_basico,
        "categoria_cct": cat_canon,
    }
