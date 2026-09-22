# -*- coding: utf-8 -*-
"""Catálogo estudio: clasificación del extracto → cuenta Tango (igual en todas las sociedades)."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from capa_revision import resolver_codigo_plan
from excel_formato_estudio import guardar_informe_excel

# Códigos del plan típico del estudio. En el extracto se pueden cambiar a mano
# si esa sociedad usa otra cuenta.
MAPA_CLASIF_ESTUDIO: dict[str, str] = {
    "Pago ARBA": "21404",
    "SIRCREB": "11421",
    "IIBB": "11418",
    "Retención bancaria": "11419",
    "Retenciones IIBB bancos": "11419",
    "Ingresos brutos Tucuman": "11404",
    "Impuestos a los débitos y créditos": "42506",
    "Impuesto a los sellos": "42512",
    "Percepción IVA": "11403",
    "IVA": "11410",
    "Rescate FIMA": "11203",
    "Suscripción FCI": "11203",
    "Gastos Bancarios": "42501",
    "Inversiones": "11203",
    "Intereses": "42508",
    "Depositos en efvo": "11101",
    "Cheques recibidos": "11107",
    "Cheques emitidos": "21705",
    "Pago de haberes": "21309",
    "Pagos AFIP": "21405",
    "Pagos tarjeta corporativa": "21202",
    "Tarjetas de crédito a pagar": "21202",
    "Compras": "21101",
    "Pago de Servicios": "21101",
    "Transferencias recibidas": "11301",
    "Transferencias emitidas": "21101",
    "Pagos recibidos": "11301",
    "Acreditaciones comercios": "11302",
}

CUENTA_TIPICA_ESTUDIO: dict[str, str] = {
    "11101": "Caja en $",
    "11107": "Valores a Depositar",
    "11203": "Fondo Fima",
    "11301": "Deudores por Ventas",
    "11302": "Deudores por Tarjeta",
    "11403": "IVA Percepción 3337",
    "11404": "Percepción Ingresos Brutos PBA",
    "11410": "Retenciones IVA",
    "11418": "Retenciones Imp. sobre los Ing Brutos",
    "11419": "Retenciones ingresos brutos banco",
    "11421": "Reten. Bancarias Sircreb Ing Brutos",
    "21101": "Proveedores",
    "21202": "Tarjeta a pagar",
    "21309": "Sueldos y Jornales a Pagar",
    "21404": "Impuesto a los Ingresos Brutos a Pagar",
    "21405": "AFIP - IVA a pagar",
    "21705": "Cheques Emitidos Pendientes de Debito",
    "42501": "Gastos y Comisiones Bancarias",
    "42506": "Impuesto a los débitos y créditos a cuenta de ganancias",
    "42508": "Intereses Bancarios",
    "42512": "Gastos de Sellado",
}

NOTA_MANUAL = "Si el plan de esa sociedad usa otra, cambiala en el extracto"

# Busca estas leyendas en el plan de ESA sociedad (no inventar código).
TEXTOS_CUENTA_CLASIF: dict[str, tuple[str, ...]] = {
    "Impuestos a los débitos y créditos": (
        "Impuesto a los débitos y créditos a cuenta de ganancias",
        "Impuesto al débito a cuenta de ganancias",
        "débitos y créditos a cuenta de ganancias",
        "debito a cuenta de ganancias",
    ),
    "Retención bancaria": (
        "Retenciones ingresos brutos banco",
        "Reten. Bancarias Imp. sobre los Ing Brutos",
        "Reten. Bancarias Imp. sobre los IIBB",
        "Retenciones Imp. sobre los Ing Brutos",
    ),
    "Retenciones IIBB bancos": (
        "Retenciones ingresos brutos banco",
        "Reten. Bancarias Imp. sobre los Ing Brutos",
        "Reten. Bancarias Imp. sobre los IIBB",
    ),
}


def resolver_cuenta_clasif_estudio(
    nombre: str,
    plan_df: pd.DataFrame | None,
) -> tuple[str, str]:
    """Cuenta del plan del cliente para una clasificación fija del estudio."""
    nom = str(nombre or "").strip()
    if not nom:
        return "99999", nom
    hint = MAPA_CLASIF_ESTUDIO.get(nom, "")
    textos = (nom,) + TEXTOS_CUENTA_CLASIF.get(nom, ())
    mejor_cod = "99999"
    mejor_desc = nom
    mejor = 0.0
    for texto in textos:
        cod, desc, score = resolver_codigo_plan(
            texto, plan_df, hints={texto: hint} if hint else None, score_min=70.0
        )
        if score > mejor and str(cod).strip() not in {"", "99999"}:
            mejor = score
            mejor_cod = str(cod).strip()
            mejor_desc = desc or nom
            if score >= 99.0:
                return mejor_cod, mejor_desc
    if mejor_cod != "99999":
        return mejor_cod, mejor_desc
    if hint:
        return hint, CUENTA_TIPICA_ESTUDIO.get(hint, nom)
    return "99999", nom


def df_catalogo_clasif() -> pd.DataFrame:
    filas = []
    for nombre, codigo in MAPA_CLASIF_ESTUDIO.items():
        filas.append(
            {
                "Clasificación": nombre,
                "Código Tango": codigo,
                "Cuenta (plan típico)": CUENTA_TIPICA_ESTUDIO.get(codigo, ""),
                "Cambio manual": NOTA_MANUAL,
            }
        )
    return pd.DataFrame(filas)


def guardar_excel_catalogo_clasif(ruta: str | Path) -> Path:
    df = df_catalogo_clasif()
    return guardar_informe_excel(
        ruta,
        titulo="Clasificaciones del extracto → cuenta Tango",
        subtitulo="Misma asociación en todas las sociedades. En el extracto se puede cambiar la cuenta a mano.",
        periodo="Plan típico del estudio",
        kpis=[
            ("Clasificaciones", len(df)),
            ("Cuentas distintas", df["Código Tango"].nunique()),
        ],
        detalle=df,
        hoja_detalle="Clasificaciones",
        col_texto=["Clasificación", "Código Tango", "Cuenta (plan típico)", "Cambio manual"],
    )
