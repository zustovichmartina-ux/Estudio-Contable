"""Verifica lógica ASIENTO DE REFUNDICIÓN / POSICIÓN MENSUAL DE IVA."""
import importlib.util
import sys
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("app_mod", ROOT / "app.py")
app_mod = importlib.util.module_from_spec(spec)
sys.modules["app_mod"] = app_mod
spec.loader.exec_module(app_mod)

_generar = app_mod._generar_asiento_cierre_iva_segregado
_generar_df = app_mod._generar_asiento_cierre_iva_desde_df
_buscar = app_mod._buscar_cuenta_iva_alias
_guardar = app_mod._guardar_plan_cliente_en_disco
_cargar_plan = app_mod.cargar_plan_cuentas
_DATA_DIR = app_mod.DATA_PLANES_DIR
_ruta_csv = app_mod._ruta_plan_csv

PLAN = pd.DataFrame([
    {"codigo": "2140101", "descripcion": "IVA Ventas 21%", "imputable": True},
    {"codigo": "2140102", "descripcion": "IVA Ventas 10,5%", "imputable": True},
    {"codigo": "1140201", "descripcion": "IVA Credito Fiscal 21", "imputable": True},
    {"codigo": "1140202", "descripcion": "IVA Credito Fiscal 10.5", "imputable": True},
    {"codigo": "1140299", "descripcion": "IVA Credito Fiscal general", "imputable": True},
    {"codigo": "2140301", "descripcion": "Percepciones IVA sufridas", "imputable": True},
    {"codigo": "2140401", "descripcion": "IVA a Pagar ARCA", "imputable": True},
])


def _guardar_balance_excel(path: str, df: pd.DataFrame, sheet: str = "IVA") -> None:
    """Simula un balance real con solapa nombrada según el impuesto activo."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False, header=False)


TEXTO_BASE = """
Debito Fiscal Actividades
21,00 % ventas gravadas 1.000.000,00 210.000,00
Crédito Fiscal Actividades
21,00 % compras gravadas 500.000,00 105.000,00
Total del crédito fiscal 105.000,00
Liquidación
Percepciones Impositivas sufridas 5.000,00
Saldo del Impuesto a Favor de ARCA 110.000,00
"""

TEXTO_DUAL = """
Debito Fiscal Actividades
10,50 % ventas gravadas 100.000,00 10.500,00
21,00 % ventas gravadas 200.000,00 42.000,00
Crédito Fiscal Actividades
10,50 % compras gravadas 50.000,00 5.250,00
21,00 % compras gravadas 80.000,00 16.800,00
Liquidación
"""


def test_ventas_21_debe():
    lineas, _, _ = _generar(TEXTO_BASE, PLAN)
    ventas = [l for l in lineas if l["Debe"] > 0 and "21" in l["Detalle"]]
    assert ventas, "Debe existir línea de ventas 21%"
    v = ventas[0]
    assert v["Debe"] == 210000.0, f"Ventas 21% debe ir a DEBE, got Debe={v['Debe']}"
    assert v["Haber"] == 0.0
    print("OK test_ventas_21_debe")


def test_compras_21_haber():
    lineas, _, _ = _generar(TEXTO_BASE, PLAN)
    compras = [l for l in lineas if l["Haber"] > 0 and "21" in l["Detalle"] and l["Cuenta"] == "1140201"]
    assert compras, "Debe existir línea de compras/crédito 21%"
    assert compras[0]["Haber"] == 105000.0
    print("OK test_compras_21_haber")


def test_loop_review_ajusta_ventas():
    texto = """
Debito Fiscal Actividades
21,00 % ventas 100,03
Crédito Fiscal Actividades
21,00 % compras 100,00
Liquidación
"""
    lineas, dif_antes, ajuste = _generar(texto, PLAN)
    assert abs(dif_antes - 0.03) < 0.001
    assert ajuste
    t_debe = sum(l["Debe"] for l in lineas)
    t_haber = sum(l["Haber"] for l in lineas)
    assert round(t_debe - t_haber, 2) == 0.0
    print("OK test_loop_review_ajusta_ventas")


def test_columnas_plan_mayusculas():
    plan_alt = pd.DataFrame([
        {"Código": "99901", "Descripción": "IVA Ventas 21%"},
        {"Código": "99902", "Descripción": "IVA Credito Fiscal 21"},
    ])
    lineas, _, _ = _generar(TEXTO_BASE, plan_alt)
    assert any(l["Cuenta"] == "99901" for l in lineas if l["Debe"] > 0)
    assert any(l["Cuenta"] == "99902" for l in lineas if l["Haber"] > 0)
    print("OK test_columnas_plan_mayusculas")


def test_segregacion_dual_alicuotas():
    lineas, _, _ = _generar(TEXTO_DUAL, PLAN)
    ventas_debe = [l for l in lineas if l["Debe"] > 0 and l["Cuenta"] in ("2140101", "2140102")]
    compras_haber = [l for l in lineas if l["Haber"] > 0 and l["Cuenta"] in ("1140201", "1140202")]
    assert len(ventas_debe) == 2
    assert len(compras_haber) == 2
    assert sorted(l["Debe"] for l in ventas_debe) == [10500.0, 42000.0]
    assert sorted(l["Haber"] for l in compras_haber) == [5250.0, 16800.0]
    print("OK test_segregacion_dual_alicuotas")


def test_no_cuenta_generica_si_hay_especifica():
    plan_mix = pd.DataFrame([
        {"codigo": "1140000", "descripcion": "IVA Credito Fiscal Compras", "imputable": True},
        {"codigo": "1140021", "descripcion": "IVA Credito Fiscal Compras 21%", "imputable": True},
        {"codigo": "2140021", "descripcion": "IVA Ventas 21%", "imputable": True},
    ])
    lineas, _, _ = _generar(TEXTO_BASE, plan_mix)
    compras_21 = [l for l in lineas if l["Haber"] > 0 and l["Cuenta"] == "1140021"]
    assert compras_21
    assert compras_21[0]["Haber"] == 105000.0
    print("OK test_no_cuenta_generica_si_hay_especifica")


def test_alias_debito_sin_palabra_ventas():
    plan_debito = pd.DataFrame([
        {"codigo": "2140999", "descripcion": "IVA Debito Fiscal 21%", "imputable": True},
        {"codigo": "1140999", "descripcion": "IVA Credito Fiscal Compras 21%", "imputable": True},
        {"codigo": "2140401", "descripcion": "IVA a Pagar", "imputable": True},
    ])
    lineas, _, _ = _generar(TEXTO_BASE, plan_debito)
    ventas = [l for l in lineas if l["Cuenta"] == "2140999" and l["Debe"] > 0]
    assert ventas, "Debe matchear IVA Debito Fiscal sin palabra 'ventas'"
    assert ventas[0]["Debe"] == 210000.0
    print("OK test_alias_debito_sin_palabra_ventas")


def test_carga_desde_csv_en_disco():
    cuit = "TEST99999"
    csv_path = _ruta_csv(cuit)
    _DATA_DIR.mkdir(parents=True, exist_ok=True)
    PLAN.to_csv(csv_path, index=False)
    try:
        df = _cargar_plan(csv_path)
        assert len(df) == len(PLAN)
        assert "codigo" in df.columns
        cod, _ = _buscar(df, "descripcion", "codigo", "ventas", "21")
        assert cod == "2140101"
    finally:
        if csv_path.exists():
            csv_path.unlink()
    print("OK test_carga_desde_csv_en_disco")


def test_liquidacion_desde_dataframe_arca():
    """Simula export ARCA en Excel con columnas separadas."""
    df = pd.DataFrame([
        ["Debito Fiscal Actividades", "", ""],
        ["21,00 % ventas gravadas", "1.000.000,00", "210.000,00"],
        ["Crédito Fiscal Actividades", "", ""],
        ["21,00 % compras gravadas", "500.000,00", "105.000,00"],
        ["Total del crédito fiscal", "", "105.000,00"],
        ["Liquidación", "", ""],
        ["Percepciones Impositivas sufridas", "", "5.000,00"],
        ["Saldo del Impuesto a Favor de ARCA", "", "110.000,00"],
    ])
    lineas, _, _ = _generar_df(df, PLAN)
    ventas = [l for l in lineas if l["Debe"] > 0 and l["Cuenta"] == "2140101"]
    compras = [l for l in lineas if l["Haber"] > 0 and l["Cuenta"] == "1140201"]
    assert ventas and ventas[0]["Debe"] == 210000.0
    assert compras and compras[0]["Haber"] == 105000.0
    assert len(lineas) >= 4
    assert all(l["Estado"] == "Ingresado" for l in lineas)
    print("OK test_liquidacion_desde_dataframe_arca")


def test_no_falso_positivo_caja_en_debito_21():
    """Caja en $ no debe mapearse como IVA Débito 21%."""
    _obtener = app_mod.obtener_cuenta_tango
    plan = pd.DataFrame([
        {"codigo": "1110101", "descripcion": "Caja en $", "imputable": True},
        {"codigo": "2140101", "descripcion": "IVA Ventas 21%", "imputable": True},
    ])
    cod, _ = _obtener(plan, "debito_21")
    assert cod == "2140101", f"Esperaba IVA Ventas 21%, obtuvo {cod}"
    print("OK test_no_falso_positivo_caja_en_debito_21")


def test_no_falso_positivo_retencion_iibb():
    """Retención IIBB no debe mapearse como Retenciones IVA."""
    _obtener = app_mod.obtener_cuenta_tango
    plan = pd.DataFrame([
        {"codigo": "2140501", "descripcion": "Retencion IIBB sufrida", "imputable": True},
        {"codigo": "2140502", "descripcion": "Retencion IVA sufrida", "imputable": True},
    ])
    cod, _ = _obtener(plan, "retenciones")
    assert cod == "2140502", f"Esperaba Retencion IVA, obtuvo {cod}"
    print("OK test_no_falso_positivo_retencion_iibb")


def test_map_iva_debito_fiscal_generico_tango():
    """Plan Tango real: 'IVA Débito Fiscal' sin % en el nombre."""
    _obtener = app_mod.obtener_cuenta_tango
    plan = pd.DataFrame([
        {"codigo": "21401", "descripcion": "IVA Débito Fiscal", "imputable": True},
        {"codigo": "21402", "descripcion": "Iva Ventas 10,5 %", "imputable": True},
        {"codigo": "11301", "descripcion": "Deudores por Ventas", "imputable": True},
        {"codigo": "41100", "descripcion": "Ventas de Servicios", "imputable": True},
    ])
    cod, desc = _obtener(plan, "debito_21")
    assert cod == "21401", f"Esperaba 21401, obtuvo {cod} ({desc})"
    cod_cf, _ = _obtener(
        pd.DataFrame([
            {"codigo": "11401", "descripcion": "IVA Crédito Fiscal", "imputable": True},
            {"codigo": "11402", "descripcion": "IVA Credito Fiscal 10.5", "imputable": True},
        ]),
        "credito_21",
    )
    assert cod_cf == "11401"
    print("OK test_map_iva_debito_fiscal_generico_tango")


def test_rescate_filtra_por_concepto():
    """Opciones de rescate no deben incluir cuentas ajenas al concepto IVA."""
    _rescate = app_mod._opciones_cuentas_rescate
    plan = pd.DataFrame([
        {"codigo": "1110101", "descripcion": "Caja en $", "imputable": True},
        {"codigo": "21401", "descripcion": "IVA Débito Fiscal", "imputable": True},
        {"codigo": "21402", "descripcion": "Iva Ventas 10,5 %", "imputable": True},
        {"codigo": "11301", "descripcion": "Deudores por Ventas", "imputable": True},
        {"codigo": "3100101", "descripcion": "Capital Social", "imputable": True},
    ])
    opts = _rescate(plan, "debito_21")
    cods = {c for c, _ in opts}
    assert "21401" in cods
    assert "1110101" not in cods
    assert "11301" not in cods
    assert "3100101" not in cods
    print("OK test_rescate_filtra_por_concepto")


def test_periodo_planilla_fecha_fin_mes():
    """B7 '04-2025' debe producir fecha de asiento 30/04/2025."""
    from datetime import date

    _parsear = app_mod._parsear_periodo_texto
    _parsear_celda = app_mod._parsear_celda_periodo_planilla
    _fecha = app_mod._fecha_asiento_iva_tango
    _fmt = app_mod._formatear_fecha_tango

    assert _parsear("04-2025") == (4, 2025)
    assert _parsear("04/2025") == (4, 2025)
    assert _parsear_celda("04-2025") == (4, 2025)
    assert _parsear_celda(pd.Timestamp("2025-04-01")) == (4, 2025)

    fecha = _fecha(4, 2025)
    assert fecha == date(2025, 4, 30)
    assert _fmt(fecha) == "30/04/2025"

    df = pd.DataFrame([
        ["Débito 21", 100.0],
        ["Débito 10.5", 0.0],
        ["Crédito 21", 50.0],
        ["Crédito 10.5", 0.0],
        ["NC ventas", 0.0],
        ["NC compras", 0.0],
        ["Período", "04-2025"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df)
        buf.close()
        asientos, resumen = app_mod.procesar_planilla_iva(buf.name, PLAN, 0.0, 0.0, 0.0, 0.0)
        assert len(asientos) == 1
        assert asientos[0].fecha == date(2025, 4, 30)
        assert asientos[0].fecha.strftime("%d/%m/%Y") == "30/04/2025"
        assert asientos[0].periodo == "04/2025"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_periodo_planilla_fecha_fin_mes")


def test_planilla_iva_posicional():
    """Planilla fija 7 filas columna B + procesar_planilla_iva."""
    _leer = app_mod.leer_planilla_iva_posicional
    _procesar = app_mod.procesar_planilla_iva
    _generar_pos = app_mod._generar_asiento_iva_planilla_posicional

    df = pd.DataFrame([
        ["Débito 21", 210000.0],
        ["Débito 10.5", 0.0],
        ["Crédito 21", 105000.0],
        ["Crédito 10.5", 0.0],
        ["NC ventas", 0.0],
        ["NC compras", 0.0],
        ["Período", "04-2026"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df)
        buf.close()
        datos = _leer(buf.name)
        assert datos["df_debito_21"] == 210000.0
        assert datos["cf_credito_21"] == 105000.0
        assert datos["periodo_texto"] == "04-2026"

        lineas, resumen, _, _ = _generar_pos(
            PLAN, datos, saldo_tecnico=0.0, saldo_libre=0.0,
            retenciones=0.0, percepciones=5000.0,
        )
        assert resumen["diferencia_previa"] == 100000.0
        assert resumen["resultado_tipo"] == "IVA a Pagar"
        ventas = [l for l in lineas if l["Debe"] > 0 and l["Cuenta"] == "2140101"]
        compras = [l for l in lineas if l["Haber"] > 0 and l["Cuenta"] == "1140201"]
        assert ventas and ventas[0]["Debe"] == 210000.0
        assert compras and compras[0]["Haber"] == 105000.0
        assert all(l["Estado"] == "Ingresado" for l in lineas)
        assert all(l["Haber"] == 0.0 or isinstance(l["Haber"], float) for l in lineas)

        asientos, _ = _procesar(buf.name, PLAN, 0.0, 0.0, 0.0, 5000.0)
        assert len(asientos) == 1
        assert asientos[0].fecha.month == 4
        assert asientos[0].fecha.year == 2026
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_planilla_iva_posicional")


def test_posicion_multi_saldo_saldo_favor():
    """Multi-saldos ➕ con Haber > Debe debe imputar Saldo a Favor al DEBE (no IVA a Pagar)."""
    _generar_pos = app_mod._generar_asiento_iva_planilla_posicional
    datos = {
        "df_debito_21": 100000.0,
        "df_debito_105": 0.0,
        "nc_compras": 0.0,
        "cf_credito_21": 80000.0,
        "cf_credito_105": 0.0,
        "nc_ventas": 0.0,
    }
    saldos_tec = [12000.0, 8000.0]
    saldos_lib = [5000.0]
    lineas, resumen, _, _ = _generar_pos(
        PLAN, datos, saldos_tec, saldos_lib, 0.0, 0.0,
    )
    assert resumen["total_saldos_anteriores"] == 25000.0
    assert resumen["diferencia_previa"] == -5000.0
    assert resumen["resultado_tipo"] == "Saldo a Favor IVA Nuevo Período"
    assert resumen["resultado_lado"] == "Debe"
    cierre = [l for l in lineas if l.get("_rol") == "saldo_favor"]
    pagar = [l for l in lineas if l.get("_rol") == "pagar"]
    assert len(pagar) == 0
    assert len(cierre) == 1
    assert cierre[0]["Debe"] == 5000.0
    assert cierre[0]["Haber"] == 0.0
    total_debe = sum(l["Debe"] for l in lineas)
    total_haber = sum(l["Haber"] for l in lineas)
    assert round(total_debe - total_haber, 2) == 0.0
    print("OK test_posicion_multi_saldo_saldo_favor")


def test_normalizar_plan_columnas_duplicadas():
    """Plan con columnas 'codigo' repetidas no debe crashear con AttributeError."""
    import procesador as proc
    df = pd.DataFrame(
        [["11401", "IVA CF", "11401", "IVA CF dup"]],
        columns=["codigo", "descripcion", "codigo", "descripcion"],
    )
    out = proc._normalizar_plan_cuentas_df(df)
    assert isinstance(out["codigo"], pd.Series)
    assert out.iloc[0]["codigo"] == "11401"
    assert out.iloc[0]["descripcion"] == "IVA CF"
    print("OK test_normalizar_plan_columnas_duplicadas")


def test_planilla_iva_etiquetas_con_27():
    """Planilla expandida con filas 27%: lectura por etiquetas, no por índice fijo."""
    _leer = app_mod.leer_planilla_iva_por_etiquetas
    df = pd.DataFrame([
        ["Débito Fiscal 21%", 100000.0],
        ["Débito Fiscal 10,5%", 0.0],
        ["Débito Fiscal 27%", 27000.0],
        ["Crédito Fiscal 21%", 50000.0],
        ["Crédito Fiscal 10,5%", 0.0],
        ["Crédito Fiscal 27%", 5400.0],
        ["NC ventas", 0.0],
        ["NC compras", 0.0],
        ["NC ventas 27%", 1000.0],
        ["Nota de crédito compras 27%", 500.0],
        ["Período", "05-2026"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df)
        buf.close()
        datos = _leer(buf.name)
        assert datos["df_debito_21"] == 100000.0
        assert datos["df_debito_27"] == 27000.0
        assert datos["cf_credito_27"] == 5400.0
        assert datos["nc_ventas_27"] == 1000.0
        assert datos["nc_compras_27"] == 500.0
        assert datos["periodo_texto"] == "05-2026"
        lineas, resumen, _, _ = app_mod._generar_asiento_iva_planilla_posicional(
            PLAN, datos, 0.0, 0.0, 0.0, 0.0,
        )
        roles = {l.get("_rol") for l in lineas}
        assert "ventas_27" in roles
        assert "compras_27" in roles
        assert "nc_ventas_27" in roles
        assert "nc_compras_27" in roles
        assert round(sum(l["Debe"] for l in lineas) - sum(l["Haber"] for l in lineas), 2) == 0.0
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_planilla_iva_etiquetas_con_27")


def test_balance_solapa_dinamica():
    """Mapeo por nombre de solapa: tolerante a espacios y mayúsculas; ignora otras solapas."""
    import procesador as proc

    df_iva = pd.DataFrame([
        ["Débito 21", 150000.0],
        ["Crédito 21", 80000.0],
        ["Período", "06-2025"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df_iva.to_excel(writer, sheet_name=" iva ", index=False, header=False)
            pd.DataFrame([["Total", 999999.0]]).to_excel(
                writer, sheet_name="Resumen", index=False, header=False,
            )
        buf.close()
        hoja = proc.resolver_solapa_por_impuesto(buf.name, "IVA")
        assert proc._normalizar_nombre_solapa_balance(hoja) == "iva"
        datos = app_mod.leer_planilla_iva_por_etiquetas(
            buf.name, nombre_solapa_impuesto="IVA",
        )
        assert datos["df_debito_21"] == 150000.0
        assert datos["cf_credito_21"] == 80000.0
        assert datos["periodo_texto"] == "06-2025"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_balance_solapa_dinamica")


def test_planilla_iibb_posicion():
    """Planilla IIBB por etiquetas + ecuación de posición mensual."""
    _leer = app_mod.leer_planilla_iibb_por_etiquetas
    _procesar = app_mod.procesar_planilla_iibb
    _generar_pos = app_mod._generar_asiento_iibb_planilla_posicional

    plan_iibb = pd.DataFrame([
        {"codigo": "2150101", "descripcion": "IIBB Devengado / Impuesto Determinado", "imputable": True},
        {"codigo": "2150201", "descripcion": "Retencion IIBB sufrida", "imputable": True},
        {"codigo": "2150301", "descripcion": "Percepciones IIBB sufridas", "imputable": True},
        {"codigo": "2150401", "descripcion": "Retenciones Bancarias IIBB Sircreb", "imputable": True},
        {"codigo": "2150501", "descripcion": "IIBB a Pagar", "imputable": True},
        {"codigo": "1150601", "descripcion": "Saldo a Favor IIBB Nuevo Periodo", "imputable": True},
    ])

    df = pd.DataFrame([
        ["Impuesto Determinado", 500000.0],
        ["Retenciones IIBB", 0.0],
        ["Percepciones IIBB", 0.0],
        ["Retenciones Bancarias", 0.0],
        ["Período", "07-2026"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df, sheet="IIBB")
        buf.close()
        datos = _leer(buf.name, nombre_solapa_impuesto="Ingresos Brutos")
        assert datos["impuesto_determinado"] == 500000.0
        assert datos["periodo_texto"] == "07-2026"

        lineas, resumen, _, _ = _generar_pos(
            plan_iibb, datos,
            saldos_favor=[8000.0],
            retenciones=10000.0,
            percepciones=5000.0,
            retenciones_bancarias=2000.0,
        )
        assert resumen["diferencia_previa"] == 475000.0
        assert resumen["resultado_tipo"] == "IIBB a Pagar"
        assert resumen["resultado_lado"] == "Haber"
        imp = [l for l in lineas if l.get("_rol") == "impuesto_determinado"]
        pagar = [l for l in lineas if l.get("_rol") == "iibb_pagar"]
        assert imp and imp[0]["Debe"] == 500000.0
        assert pagar and pagar[0]["Haber"] == 475000.0
        assert round(sum(l["Debe"] for l in lineas) - sum(l["Haber"] for l in lineas), 2) == 0.0

        asientos, _ = _procesar(
            buf.name, plan_iibb, [8000.0], 10000.0, 5000.0, 2000.0,
            nombre_solapa_impuesto="Ingresos Brutos",
        )
        assert len(asientos) == 1
        assert asientos[0].tipo == "IIBB"
        assert asientos[0].periodo == "07/2026"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_planilla_iibb_posicion")


def test_balance_solapa_iibb():
    """Mapeo por nombre de solapa IIBB vía aliases de procesador."""
    import procesador as proc

    df_iibb = pd.DataFrame([
        ["IIBB Devengado", 120000.0],
        ["Retenciones", 3000.0],
        ["Percepciones", 1500.0],
        ["Período", "08-2025"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df_iibb.to_excel(writer, sheet_name="IIBB", index=False, header=False)
            pd.DataFrame([["Total", 999999.0]]).to_excel(
                writer, sheet_name="Resumen", index=False, header=False,
            )
        buf.close()
        hoja = proc.resolver_solapa_por_impuesto(buf.name, "Ingresos Brutos")
        assert proc._normalizar_nombre_solapa_balance(hoja) == "iibb"
        datos = app_mod.leer_planilla_iibb_por_etiquetas(
            buf.name, nombre_solapa_impuesto="Ingresos Brutos",
        )
        assert datos["impuesto_determinado"] == 120000.0
        assert datos["retenciones_planilla"] == 3000.0
        assert datos["percepciones_planilla"] == 1500.0
        assert datos["periodo_texto"] == "08-2025"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_balance_solapa_iibb")


def test_balance_unc_path():
    """cargar_balance_desde_ruta_unc lee Excel local sin validación http."""
    import procesador as proc

    df = pd.DataFrame([
        ["Débito 21", 100000.0],
        ["Crédito 21", 50000.0],
        ["Período", "01-2026"],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df, sheet="IVA")
        buf.close()
        loaded = proc.cargar_balance_desde_ruta_unc(buf.name)
        assert loaded.getvalue()
        hoja = proc.resolver_solapa_por_impuesto(loaded, "IVA")
        assert proc._normalizar_nombre_solapa_balance(hoja) == "iva"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_balance_unc_path")


def test_sanitizar_ruta_unc_comillas():
    """Quita comillas de 'Copiar como ruta' en Windows antes de abrir el archivo."""
    import procesador as proc

    df = pd.DataFrame([["Débito 21", 100.0], ["Período", "02-2026"]])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        _guardar_balance_excel(buf.name, df, sheet="IVA")
        buf.close()
        ruta_con_comillas = f'"{buf.name}"'
        assert proc.sanitizar_ruta_unc(ruta_con_comillas) == buf.name
        loaded = proc.cargar_balance_desde_ruta_unc(ruta_con_comillas)
        assert loaded.getvalue()
        ruta_simple = f"'{buf.name}'"
        loaded2 = proc.cargar_balance_desde_ruta_unc(ruta_simple)
        assert loaded2.getvalue()
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_sanitizar_ruta_unc_comillas")


def test_tax_registry_solapas_sueldos_tish():
    import procesador as proc

    sueldos = proc.solapas_impuesto("Sueldos")
    tish = proc.solapas_impuesto("TISH")
    assert "Sueldos" in sueldos
    assert "F931" in sueldos
    assert "TISH" in tish
    assert "Seguridad e Higiene" in tish
    ficha_iva = proc.obtener_ficha_impuesto("IVA")
    assert ficha_iva["motor"] == "iva"
    assert ficha_iva["cuenta_ajuste_centavos_rol"] == "ventas_21"
    print("OK test_tax_registry_solapas_sueldos_tish")


def test_detectar_periodos_en_balance_df():
    import procesador as proc

    df = pd.DataFrame([
        ["Concepto", "01/2026", "02/2026", "03/2026"],
        ["Débito 21", 100000.0, 120000.0, 130000.0],
        ["Crédito 21", 50000.0, 60000.0, 65000.0],
    ])
    periodos = proc.detectar_periodos_en_balance_df(df)
    assert periodos == ["01/2026", "02/2026", "03/2026"]
    print("OK test_detectar_periodos_en_balance_df")


def test_resolver_indice_columna_periodo():
    import procesador as proc

    df = pd.DataFrame([
        ["Concepto", "01/2026", "02/2026", "03/2026"],
        ["Débito 21", 100000.0, 120000.0, 130000.0],
    ])
    assert proc.resolver_indice_columna_periodo(df, "02/2026") == 2
    assert proc.resolver_indice_columna_periodo(df, "01-2026") == 1
    assert proc.resolver_indice_columna_periodo(df, "12/2099") is None
    print("OK test_resolver_indice_columna_periodo")


def test_planilla_iva_etiquetas_periodo_mensual():
    """Multi-columna: lee montos de la columna del período seleccionado."""
    _leer = app_mod.leer_planilla_iva_por_etiquetas
    df_alt = pd.DataFrame([
        ["Concepto", "01/2026", "02/2026"],
        ["Débito Fiscal 21%", 100000.0, 210000.0],
        ["Crédito Fiscal 21%", 50000.0, 105000.0],
        ["Retenciones", 1000.0, 2000.0],
        ["Percepciones", 500.0, 5000.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df_alt.to_excel(writer, sheet_name="IVA", index=False, header=False)
        buf.close()
        datos_ene = _leer(buf.name, nombre_solapa_impuesto="IVA", periodo_mensual="01/2026")
        datos_feb = _leer(buf.name, nombre_solapa_impuesto="IVA", periodo_mensual="02/2026")
        assert datos_ene["df_debito_21"] == 100000.0
        assert datos_ene["retenciones_planilla"] == 1000.0
        assert datos_feb["df_debito_21"] == 210000.0
        assert datos_feb["retenciones_planilla"] == 2000.0
        assert datos_feb["percepciones_planilla"] == 5000.0
        assert datos_feb["periodo_mes"] == 2
        assert datos_feb["periodo_anio"] == 2026
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_planilla_iva_etiquetas_periodo_mensual")


def test_periodo_siguiente():
    _sig = app_mod._periodo_siguiente
    assert _sig("01/2026") == "02/2026"
    assert _sig("12/2026") == "01/2027"
    assert _sig("06-2025") == "07/2025"
    assert _sig("invalid") is None
    print("OK test_periodo_siguiente")


def test_match_fechas_completas_encabezados():
    import procesador as proc

    assert proc.etiqueta_periodo_desde_celda("31/01/2025") == "01/2025"
    assert proc.etiqueta_periodo_desde_celda("28/02/2025") == "02/2025"
    assert proc.etiqueta_periodo_desde_celda("30/04/2025") == "04/2025"
    ts = pd.Timestamp("2025-04-30")
    assert proc.etiqueta_periodo_desde_celda(ts) == "04/2025"
    df = pd.DataFrame([
        ["Balance IIBB — Cliente XYZ", None, None, None],
        ["Control interno", None, None, None],
        ["Concepto", "31/01/2025", "28/02/2025", "30/04/2025"],
        ["42405 Impuesto determinado", 100.0, 200.0, 300.0],
    ])
    assert proc.resolver_indice_columna_periodo(df, "04/2025") == 3
    assert proc.resolver_indice_columna_periodo(df, "01/2025") == 1
    periodos = proc.detectar_periodos_en_balance_df(df)
    assert periodos == ["01/2025", "02/2025", "04/2025"]
    print("OK test_match_fechas_completas_encabezados")


def test_filtro_ruido_totales_balance():
    import procesador as proc

    assert proc._es_fila_ruido_balance("TOTAL GENERAL IIBB")
    assert proc._es_fila_ruido_balance("Subtotal Retenciones")
    assert proc._es_fila_ruido_balance("Control de cuadro")
    assert not proc._es_fila_ruido_balance("42405 Impuesto determinado IIBB")
    assert not proc._es_fila_ruido_balance("11418 a Retenciones Sircreb")

    df = pd.DataFrame([
        ["Concepto", "30/04/2025"],
        ["42405 Impuesto determinado IIBB", 150000.0],
        ["TOTAL GENERAL", 150000.0],
        ["11418 a Retenciones Sircreb", 5000.0],
        ["Control interno", 0.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="IIBB", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            filas = proc.extraer_filas_universales_balance_por_periodo(f, "IIBB", "04/2025")
        cods = {x["codigo"] for x in filas}
        assert "42405" in cods
        assert "11418" in cods
        assert len(filas) == 2
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_filtro_ruido_totales_balance")


def test_extractor_universal_codigo_y_tipo():
    import procesador as proc

    cod, desc = proc.extraer_codigo_cuenta_tango_desde_concepto("42405 Impuesto sobre los Ingresos Brutos")
    assert cod == "42405"
    assert "Impuesto" in desc

    cod2, _ = proc.extraer_codigo_cuenta_tango_desde_concepto("11418 a Retenciones Sircreb")
    assert cod2 == "11418"
    assert proc.inferir_tipo_movimiento_desde_concepto("11418 a Retenciones Sircreb") == "Haber"
    assert proc.inferir_tipo_movimiento_desde_concepto("42405 Impuesto determinado") == "Debe"
    assert proc.inferir_tipo_movimiento_desde_concepto("21105 IIBB a pagar") == "Haber"
    assert proc.inferir_tipo_movimiento_desde_concepto(
        "21309 a Sueldos a Pagar", codigo="21309", impuesto="Sueldos",
    ) == "Haber"
    assert proc.inferir_tipo_movimiento_desde_concepto(
        "42203 Cargas Sociales", codigo="42203", impuesto="Sueldos",
    ) == "Debe"
    assert proc.inferir_tipo_movimiento_desde_concepto(
        "Retención ganancias", monto_original=-142788.68, impuesto="Sueldos",
    ) == "Haber"
    assert proc._celda_a_float_balance("(1.234,56)") == -1234.56
    assert proc._celda_a_float_balance(-5000) == -5000.0
    print("OK test_extractor_universal_codigo_y_tipo")


def test_extractor_universal_filas_balance():
    import procesador as proc

    df = pd.DataFrame([
        ["Concepto", "30/04/2025", "31/05/2025"],
        ["42405 Impuesto determinado IIBB", 150000.0, 160000.0],
        ["11418 a Retenciones bancarias Sircreb", 5000.0, 6000.0],
        ["11502 a Percepciones IIBB", 3000.0, 3500.0],
        ["Sin monto", 0.0, 100.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="IIBB", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            filas = proc.extraer_filas_universales_balance_por_periodo(
                f, "IIBB", "04/2025",
            )
        assert len(filas) == 3
        by_cod = {f["codigo"]: f for f in filas}
        assert by_cod["42405"]["monto"] == 150000.0
        assert by_cod["42405"]["tipo"] == "Debe"
        assert by_cod["11418"]["monto"] == 5000.0
        assert by_cod["11418"]["tipo"] == "Haber"
        assert by_cod["11502"]["tipo"] == "Haber"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_extractor_universal_filas_balance")


def test_extractor_banco_sincronico_fila_por_fila():
    """Bancos: monto y código en la misma fila; filas con saldo cero no desplazan montos."""
    import procesador as proc

    df = pd.DataFrame([
        ["Concepto", "may-25", "jun-25"],
        ["42405 FATSA retencion", None, 0],
        ["11418 SICORE", None, 0],
        ["Cheques", None, 10317775.01],
        ["11101 Valores a depositar", None, None],
        ["11301 PAMI", None, 500000.0],
        ["11104 Cta cte Santander", None, 1200000.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SANTANDER", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            resultado = proc.extraer_filas_universales_balance_por_banco_con_errores(
                f, "Santander", "06/2025",
            )
        assert not resultado.error, resultado.error
        assert len(resultado.filas) == 3
        assert [f["fila_idx"] for f in resultado.filas] == [3, 5, 6]
        cheques = [x for x in resultado.filas if "Cheques" in x["concepto_raw"]]
        assert len(cheques) == 1
        assert cheques[0]["codigo"] == ""
        assert cheques[0]["monto"] == 10317775.01
        assert cheques[0]["tipo"] == "Debe"
        by_cod = {item["codigo"]: item for item in resultado.filas if item["codigo"]}
        assert "42405" not in by_cod
        assert "11418" not in by_cod
        assert by_cod["11301"]["monto"] == 500000.0
        assert by_cod["11301"]["tipo"] == "Debe"
        assert by_cod["11104"]["monto"] == 1200000.0
        assert by_cod["11104"]["tipo"] == "Debe"
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_extractor_banco_sincronico_fila_por_fila")


def test_extractor_banco_partida_doble_columnas():
    """Bancos: idx_debe → Debe; idx_haber → Haber (inyección directa por columna)."""
    import procesador as proc

    df = pd.DataFrame([
        ["Concepto", "debe", "haber", "debe", "haber"],
        ["", "may-25", "", "jun-25", ""],
        ["11102 Banco Santander", 1000, 500, 2000, 800],
        ["11301 PAMI", 0, 300000, 0, 500000],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SANTANDER", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            resultado = proc.extraer_filas_universales_balance_por_banco_con_errores(
                f, "Santander", "06/2025",
            )
        assert not resultado.error, resultado.error
        pami = [x for x in resultado.filas if x["codigo"] == "11301"]
        assert len(pami) == 1
        assert pami[0]["tipo"] == "Haber"
        assert pami[0]["haber"] == 500000.0
        assert pami[0]["debe"] == 0.0
        sant = [x for x in resultado.filas if x["codigo"] == "11102"]
        assert len(sant) == 2
        tipos = {x["tipo"] for x in sant}
        assert tipos == {"Debe", "Haber"}
        debe_s = next(x for x in sant if x["tipo"] == "Debe")
        haber_s = next(x for x in sant if x["tipo"] == "Haber")
        assert debe_s["debe"] == 2000.0
        assert haber_s["haber"] == 800.0
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_extractor_banco_partida_doble_columnas")


def test_mercado_pago_disponible_y_resuelve_solapa():
    """Mercado Pago participa del mismo motor bancario y acepta aliases de solapa."""
    import procesador as proc

    assert "Mercado Pago" in proc.listar_bancos_conciliacion()
    ficha = proc.obtener_ficha_banco("mercadopago")
    assert ficha["motor"] == "banco"
    assert ficha["codigo_tango"] == "MERCADOPAGO"

    df = pd.DataFrame([
        ["Concepto", "debe", "haber"],
        ["", "jul-26", ""],
        ["11107 Mercado Pago", 125000.0, 25000.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="MercadoPago", index=False, header=False)
        buf.close()

        assert proc.resolver_solapa_por_banco(buf.name, "Mercado Pago") == "MercadoPago"
        with open(buf.name, "rb") as f:
            resultado = proc.extraer_filas_universales_balance_por_banco_con_errores(
                f, "Mercado Pago", "07/2026",
            )
        assert not resultado.error, resultado.error
        filas_mp = [x for x in resultado.filas if x["codigo"] == "11107"]
        assert {x["tipo"] for x in filas_mp} == {"Debe", "Haber"}
        assert {x["monto"] for x in filas_mp} == {125000.0, 25000.0}
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_mercado_pago_disponible_y_resuelve_solapa")


def test_resolver_cuenta_banco_hibrida():
    """Intérprete híbrido: código, caja por texto, banco+nro, sin nro → caja."""
    import procesador as proc

    plan = [
        ("11102", "BANCO SANTANDER"),
        ("11105", "Banco Galicia Cuenta Corriente"),
        ("11100", "Banco Macro Cuenta Corriente 3-684-0942508301-8"),
        ("11301", "PAMI"),
        ("11101", "CAJA"),
        ("42405", "IVA DEBITO FISCAL"),
        ("42203", "GASTOS BANCARIOS"),
    ]
    c, d = proc.resolver_cuenta_banco_hibrida("11301 PAMI", plan)
    assert c == "11301"
    assert "PAMI" in d.upper()
    c, d = proc.resolver_cuenta_banco_hibrida("rendicion caja chica", plan)
    assert c == "11101"
    c, d = proc.resolver_cuenta_banco_hibrida("Cheques", plan)
    assert c == ""
    assert d == "Cheques"
    c, d = proc.resolver_cuenta_banco_hibrida("---", plan)
    assert c == ""
    assert d == ""
    c, d = proc.resolver_cuenta_banco_hibrida("comision mantenimiento", plan)
    assert c in ("42203", "")
    # Con número de cuenta → busca en el plan
    c, d = proc.resolver_cuenta_banco_hibrida(
        "Banco Macro Cuenta Corriente 3-684-0942508301-8", plan,
    )
    assert c == "11100"
    # Banco sin número y una sola cuenta de ese banco → esa cuenta
    c, d = proc.resolver_cuenta_banco_hibrida("Banco Galicia Cuenta Corriente", plan)
    assert c == "11105"
    # Banco genérico / varias sin número claro → Caja
    plan_multi_macro = plan + [("11109", "Banco Macro Caja de Ahorro")]
    c, d = proc.resolver_cuenta_banco_hibrida("Transferencia Banco Macro", plan_multi_macro)
    assert c == "11101"
    print("OK test_resolver_cuenta_banco_hibrida")


def test_conciliar_banco_con_tango_match_monto_y_comision():
    """Conciliación Galicia: match por monto en subdiario y reglas fijas de gastos."""
    import procesador as proc

    df_galicia = pd.DataFrame([
        ["31/05/2025", "TRANSFERENCIA PAMI SA", "", "500000,00", "", "1500000,00"],
        ["01/06/2025", "COMISION MANTENIMIENTO CTA", "", "", "-1.500,00", "1498500,00"],
    ], columns=["Fecha", "Descripción", "Origen", "Crédito", "Débito", "Saldo"])
    df_tango = pd.DataFrame([
        {"Cuenta": "11301", "Razón social": "PAMI SA", "Debe": 500000.0, "Haber": 0.0},
        {"Cuenta": "21101", "Razón social": "TELECOM", "Debe": 0.0, "Haber": 12000.0},
    ])
    plan = [
        ("11301", "Deudores por Ventas"),
        ("21101", "Proveedores"),
        ("42203", "Gastos Bancarios"),
    ]
    lineas = proc.conciliar_banco_con_tango(df_galicia, df_tango, plan)
    assert len(lineas) == 2
    assert lineas[0]["tipo"] == "Haber"
    assert lineas[0]["monto"] == 500000.0
    assert lineas[0]["cuenta"].startswith("11301")
    assert lineas[1]["tipo"] == "Debe"
    assert lineas[1]["monto"] == 1500.0
    assert "42203" in lineas[1]["cuenta"] or "42400" in lineas[1]["cuenta"] or "GASTOS" in lineas[1]["cuenta"].upper()

    c, _ = proc.sugerir_cuenta_conciliacion_tango(
        "LEY 25413 DEBITO", 900.0, False, None, plan,
    )
    assert c in ("42405", "")
    print("OK test_conciliar_banco_con_tango_match_monto_y_comision")


def test_formateo_periodo_y_fecha_grilla():
    """Período MM/YYYY y Fecha DD/MM/YYYY sin timestamps ni hora."""
    import procesador as proc

    assert proc.formatear_periodo_mm_yyyy("06/2025") == "06/2025"
    assert proc.formatear_periodo_mm_yyyy("06-2025") == "06/2025"
    assert proc.formatear_periodo_mm_yyyy(pd.Timestamp("2025-06-30")) == "06/2025"
    assert proc.formatear_fecha_dd_mm_yyyy(date(2025, 6, 30)) == "30/06/2025"
    assert proc.formatear_fecha_dd_mm_yyyy(pd.Timestamp("2025-06-30 00:00:00")) == "30/06/2025"
    assert proc.formatear_fecha_dd_mm_yyyy("2025-06-30 00:00:00") == "30/06/2025"
    assert proc.formatear_fecha_dd_mm_yyyy("30/06/2025") == "30/06/2025"
    assert "T" not in proc.formatear_fecha_dd_mm_yyyy("2025-06-30T00:00:00")
    assert "00:00:00" not in proc.formatear_fecha_dd_mm_yyyy("2025-06-30 00:00:00")
    print("OK test_formateo_periodo_y_fecha_grilla")


def test_balance_ruta_relativa_proyecto():
    """Rutas ./ relativas se resuelven contra BASE_DIR del proyecto."""
    import procesador as proc

    origen = proc.BASE_DIR / "Copia de OFTALMOLOGIA RELE Balance 2026.xlsx"
    if not origen.is_file():
        print("SKIP test_balance_ruta_relativa_proyecto (sin Excel real en raíz)")
        return

    path = proc.resolver_ruta_balance_archivo(r"./Copia de OFTALMOLOGIA RELE Balance 2026.xlsx")
    assert path.is_file()
    loaded = proc.cargar_balance_desde_ruta(r"./Copia de OFTALMOLOGIA RELE Balance 2026.xlsx")
    assert loaded.getvalue()
    assert proc.ruta_balance_local_por_sociedad(
        nombre="OFTALMOLOGIA RELE MAR DEL PLATA S.R.L.",
        cuit="30718022742",
        sociedad_id=177,
    ).endswith("Copia de OFTALMOLOGIA RELE Balance 2026.xlsx")
    print("OK test_balance_ruta_relativa_proyecto")


def test_filtro_ruido_posicion_ejercicio():
    import procesador as proc

    assert proc._es_fila_ruido_balance("POSICION IVA ABRIL")
    assert proc._es_fila_ruido_balance("Ejercicio cerrado al 31/03/2026")
    assert proc._es_fila_ruido_balance("CHECK cuadro")
    assert proc._es_fila_ruido_balance("DIFERENCIA DE CONTROL")
    assert not proc._es_fila_ruido_balance("42405 Impuesto determinado IIBB")
    # NC: proyectar en asiento (no tratar "Nota de crédito…" como nota de pie)
    assert not proc._es_fila_ruido_balance("Nota de crédito compras")
    assert not proc._es_fila_ruido_balance("Nota de credito ventas 21%")
    assert not proc._es_fila_ruido_balance("NC compras")
    assert proc._es_fila_ruido_balance("Nota: ver detalle")
    assert proc._es_fila_ruido_balance("PAYWAY 06/2024")
    assert proc.inferir_tipo_movimiento_desde_concepto(
        "Nota de crédito compras", monto_original=-1500.0, impuesto="IVA",
    ) == "Debe"
    assert proc.inferir_tipo_movimiento_desde_concepto(
        "NC ventas", monto_original=-800.0, impuesto="IVA",
    ) == "Haber"
    print("OK test_filtro_ruido_posicion_ejercicio")


def test_extractor_oftalmologia_iva_real():
    """Extracción matricial sobre el balance real en raíz (IVA 04/2025)."""
    import procesador as proc

    origen = proc.BASE_DIR / "Copia de OFTALMOLOGIA RELE Balance 2026.xlsx"
    if not origen.is_file():
        print("SKIP test_extractor_oftalmologia_iva_real (sin Excel real en raíz)")
        return

    resultado = proc.extraer_filas_universales_balance_por_periodo_con_errores(
        r"./Copia de OFTALMOLOGIA RELE Balance 2026.xlsx",
        "IVA",
        "04/2025",
    )
    assert resultado.error is None
    assert resultado.solapa_resuelta == "IVA"
    assert len(resultado.filas) >= 3
    codigos = {f["codigo"] for f in resultado.filas}
    assert "21401" in codigos
    print("OK test_extractor_oftalmologia_iva_real")


def test_match_meses_texto_cabeceras_sueldos():
    """Traductor de Períodos Textuales: Abril, ABRIL, Sueldos Abril, Abr → 04/YYYY."""
    import procesador as proc

    assert proc.etiqueta_periodo_desde_celda("Abril 2025") == "04/2025"
    assert proc.etiqueta_periodo_desde_celda("ABRIL") is None
    assert proc.celda_coincide_periodo_seleccionado("ABRIL", 4, 2025)
    assert proc.celda_coincide_periodo_seleccionado("Sueldos Abril", 4, 2025)
    assert proc.celda_coincide_periodo_seleccionado("Abr", 4, 2025)
    assert proc.celda_coincide_periodo_seleccionado("Mayo 2025", 5, 2025)
    assert not proc.celda_coincide_periodo_seleccionado("Marzo", 4, 2025)

    df = pd.DataFrame([
        ["", "Ejercicio cerrado al 31/03/2026", "", "", ""],
        ["", "Asiento de Sueldos", "", "Sueldos Abril", "MAYO"],
        ["42201", "Sueldos y Jornales", "", 9583921.72, 9914286.67],
        ["42203", "Cargas Sociales", "", 901348.36, 985974.58],
        ["", "Total Debe", "", 10485270.75, 10900261.25],
    ])
    assert proc.resolver_indice_columna_periodo(df, "04/2025") == 3
    assert proc.resolver_indice_columna_periodo(df, "05/2025") == 4

    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SUELDOS", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            filas = proc.extraer_filas_universales_balance_por_periodo(f, "Sueldos", "04/2025")
        cods = {x["codigo"] for x in filas}
        assert "42201" in cods
        assert "42203" in cods
        assert "TOTAL" not in {x["concepto_raw"].upper() for x in filas}
        assert len(filas) == 2
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_match_meses_texto_cabeceras_sueldos")


def test_extractor_sueldos_pasivos_negativos():
    """Pasivos e impuestos con signo negativo o paréntesis deben succionarse al Haber."""
    import procesador as proc

    df = pd.DataFrame([
        ["", "Ejercicio cerrado al 31/03/2026", "", "", ""],
        ["", "Asiento de Sueldos", "", "Abril", "Mayo"],
        ["42201", "Sueldos y Jornales", "", 1000000.0, 1100000.0],
        ["21309", "a Sueldos y Jornales a Pagar", "", -832483.48, "(900000,00)"],
        ["21302", "a Cargas Sociales a pagar", "", -229642.34, -240000.0],
        ["", "Retención ganancias", "", -142788.68, 0.0],
        ["", "Total Haber", "", 1200000.0, 0.0],
    ])
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SUELDOS", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            filas = proc.extraer_filas_universales_balance_por_periodo(f, "Sueldos", "04/2025")
        by_cod = {f["codigo"]: f for f in filas}
        assert by_cod["42201"]["tipo"] == "Debe"
        assert by_cod["42201"]["monto"] == 1000000.0
        assert by_cod["21309"]["tipo"] == "Haber"
        assert by_cod["21309"]["monto"] == 832483.48
        assert by_cod["21302"]["tipo"] == "Haber"
        assert by_cod["21302"]["monto"] == 229642.34
        assert "99999" in by_cod
        assert by_cod["99999"]["tipo"] == "Haber"
        assert by_cod["99999"]["monto"] == 142788.68
        assert len(filas) == 4
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_extractor_sueldos_pasivos_negativos")


def test_columna_periodo_estricta_sin_desfase():
    """Los montos extraídos deben coincidir con df.iloc[fila, col_idx] exacto."""
    import procesador as proc

    df = pd.DataFrame([
        ["", "Concepto", "30/04/2025", "31/05/2025"],
        ["42201", "Sueldos y Jornales", 1000000.0, 1100000.0],
        ["21309", "a Sueldos a Pagar", 832483.48, 900000.0],
    ])
    fila, col, cab, _ = proc.localizar_columna_periodo_estricto(df, "04/2025")
    assert col == 2
    assert cab == "30/04/2025"
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SUELDOS", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            r = proc.extraer_filas_universales_balance_por_periodo_con_errores(
                f, "Sueldos", "04/2025",
            )
        assert r.columna_indice == 2
        assert r.columna_cabecera_texto == "30/04/2025"
        by_cod = {x["codigo"]: x for x in r.filas}
        assert by_cod["42201"]["monto"] == 1000000.0
        assert by_cod["21309"]["monto"] == 832483.48
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_columna_periodo_estricta_sin_desfase")


def test_sueldos_ofthalmologia_sin_desfase_primeras_filas():
    """Primeras filas SUELDOS 04/2025 no deben tomar importes de otra columna (ej. IVA 304711.44)."""
    import procesador as proc

    origen = proc.BASE_DIR / "Copia de OFTALMOLOGIA RELE Balance 2026.xlsx"
    usa_real = origen.is_file()
    if usa_real:
        resultado = proc.extraer_filas_universales_balance_por_periodo_con_errores(
            r"./Copia de OFTALMOLOGIA RELE Balance 2026.xlsx",
            "Sueldos",
            "04/2025",
        )
        col_esperada = 3
    else:
        df = pd.DataFrame([
            [None, "Ejercicio cerrado al 31/03/2026", None, None, None, None],
            [None, "Asiento de Sueldos", None, "28/02/2025", "30/04/2025", "31/05/2025"],
            [None, None, None, None, None, None],
            [None, "Devengamiento de Sueldos", None, None, None, None],
            [None, 42201, "Sueldos y Jornales (R+NR) REM 9", 304711.44, 9583921.72, 9914286.67],
            [None, 42203, "Cargas Sociales", 156913.53, 901348.36, 985974.58],
            [None, 21309, "a Sueldos y Jornales a Pagar", 211631.58, 8324831.48, 8624455.34],
            [None, 21302, "a Contribuciones Seg Social a pagar", 653.92, 229642.34, 287259.72],
        ])
        buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        try:
            with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="SUELDOS", index=False, header=False)
            buf.close()
            with open(buf.name, "rb") as f:
                resultado = proc.extraer_filas_universales_balance_por_periodo_con_errores(
                    f, "Sueldos", "04/2025",
                )
        finally:
            Path(buf.name).unlink(missing_ok=True)
        col_esperada = 4

    assert resultado.error is None, resultado.error
    assert resultado.columna_indice == col_esperada
    assert resultado.coordenadas is not None
    assert resultado.idx_debe == col_esperada
    assert resultado.idx_haber == col_esperada + 1
    if resultado.coordenadas is not None:
        assert not resultado.coordenadas.mellizas_debe_haber

    by_cod = {x["codigo"]: x for x in resultado.filas}
    assert "42201" in by_cod
    assert abs(by_cod["42201"]["monto"] - 9583921.72) < 0.02
    if "42203" in by_cod:
        assert abs(by_cod["42203"]["monto"] - 901348.36) < 0.02
    if "21309" in by_cod:
        assert abs(by_cod["21309"]["monto"] - 8324831.48) < 0.02

    trap = [f for f in resultado.filas if abs(f["monto"] - 304711.44) < 0.01]
    assert not trap, f"Desfase en primeras filas: {trap}"
    print("OK test_sueldos_ofthalmologia_sin_desfase_primeras_filas")


def test_escaneo_cabecera_seguro_texto_institucional():
    """El escáner no debe crashear con razón social ni textos largos en cabecera."""
    import procesador as proc

    df = pd.DataFrame([
        [None, "OFTALMOLOGIA RELE MAR DEL PLATA S.R.L.", None, None, None],
        [None, "Ejercicio cerrado al 31/03/2026", None, None, None],
        [None, "Asiento de Sueldos", None, "30/04/2025", "31/05/2025"],
        [None, 42201, "Sueldos y Jornales", 9583921.72, 9914286.67],
    ])
    coords = proc.congelar_coordenadas_balance(df, "04/2025")
    assert coords is not None
    assert coords.idx_debe == 3
    assert coords.idx_haber == 4
    buf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        with pd.ExcelWriter(buf.name, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="SUELDOS", index=False, header=False)
        buf.close()
        with open(buf.name, "rb") as f:
            r = proc.extraer_filas_universales_balance_por_periodo_con_errores(
                f, "Sueldos", "04/2025",
            )
        assert r.error is None
        assert r.filas
    finally:
        Path(buf.name).unlink(missing_ok=True)
    print("OK test_escaneo_cabecera_seguro_texto_institucional")


def test_matching_flexible_marzo_cierre():
    """Marzo (mes de cierre) debe resolverse con texto, 03 o fecha 31/03."""
    import procesador as proc

    assert proc.celda_coincide_periodo_flexible("Marzo 2026", 3, 2026)
    assert proc.celda_coincide_periodo_flexible("mar 2026", 3, 2026)
    assert proc.celda_coincide_periodo_flexible("31/03/2026", 3, 2026)
    assert proc.celda_coincide_periodo_flexible("03/2026", 3, 2026)
    df = pd.DataFrame([
        [None, "Ejercicio cerrado al 31/03/2026", None, "Marzo 2026", "TOTAL"],
        [None, 42201, "Sueldos", 100.0, None],
    ])
    coords = proc.congelar_coordenadas_balance(df, "03/2026")
    assert coords is not None
    assert coords.idx_debe == 3
    assert coords.idx_haber == 3
    print("OK test_matching_flexible_marzo_cierre")


def test_export_tango_fecha_txt_aaaammdd():
  """Fechas de export TXT deben ir en AAAAMMDD, no DD/MM/YYYY."""
  import procesador as proc
  from datetime import date

  assert proc.formatear_fecha_tango_export_txt(date(2026, 3, 31)) == "20260331"
  assert proc.formatear_fecha_tango_export_txt("31/03/2026") == "20260331"
  assert proc.formatear_fecha_tango_export_txt("20260331") == "20260331"
  print("OK test_export_tango_fecha_txt_aaaammdd")


def test_export_tango_balancea_centavos():
  """Diferencias de redondeo ≤ $5 se ajustan automáticamente al exportar."""
  import procesador as proc
  from datetime import date

  asiento = proc.AsientoDevengamiento(
      identificador=1,
      concepto="IVA 03/2026",
      fecha=date(2026, 3, 31),
      renglones=[
          proc.RenglonAsiento("2140101", "IVA Ventas", debe=1000.03, haber=0),
          proc.RenglonAsiento("2140401", "IVA a Pagar", debe=0, haber=1000.00),
      ],
  )
  assert not asiento.balanceado
  proc.balancear_asiento_export_tango(asiento, tolerancia_centavos=5.0)
  assert asiento.balanceado
  assert asiento.total_debe == asiento.total_haber == 1000.00
  print("OK test_export_tango_balancea_centavos")


def test_export_tango_rechaza_cuenta_madre():
  """Cuentas Rubro/Madre no imputables deben bloquear la exportación."""
  import procesador as proc
  from datetime import date

  plan = pd.DataFrame([
      {"codigo": "21401", "descripcion": "IVA Débito Fiscal", "imputable": False},
      {"codigo": "2140101", "descripcion": "IVA Ventas 21%", "imputable": True},
  ])
  asiento = proc.AsientoDevengamiento(
      identificador=1,
      concepto="IVA 03/2026",
      fecha=date(2026, 3, 31),
      renglones=[
          proc.RenglonAsiento("21401", "IVA Débito Fiscal", debe=100.0, haber=0),
          proc.RenglonAsiento("2140401", "IVA a Pagar", debe=0, haber=100.0),
      ],
  )
  asiento.periodo = "03/2026"
  errores = proc.auditar_renglones_imputables_tango([asiento], plan)
  assert len(errores) == 1
  assert errores[0]["codigo"] == "21401"
  motivo = errores[0]["motivo"].lower()
  assert "madre" in motivo or "rubro" in motivo or "no imputable" in motivo
  print("OK test_export_tango_rechaza_cuenta_madre")


def test_codigo_tipo_asiento_tango_mapping():
  """El tipo interno del asiento debe mapear al código válido del template Tango."""
  import procesador as proc

  assert proc._codigo_tipo_asiento_tango("SUELDOS") == "SUELDOS"
  assert proc._codigo_tipo_asiento_tango("IVA") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("IIBB") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("IIBB_ARBA") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("IIBB_CM03") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("BANCO") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("CM") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("") == "VARIOS"
  assert proc._codigo_tipo_asiento_tango("SUELDOSRES") == "SUELDOSRES"
  print("OK test_codigo_tipo_asiento_tango_mapping")


def test_generar_excel_tango_nativo_tipo_sueldos():
  """El Excel nativo debe exportar SUELDOS, no VARIOS, cuando el asiento es de sueldos."""
  import procesador as proc
  from datetime import date
  import openpyxl
  from pathlib import Path

  asiento = proc.AsientoDevengamiento(
      identificador=1001,
      concepto="Devengamiento SUELDOS",
      fecha=date(2026, 5, 31),
      renglones=[
          proc.RenglonAsiento("42201", "Sueldos", debe=100.0, haber=0),
          proc.RenglonAsiento("21312", "Sueldos a Pagar", debe=0, haber=100.0),
      ],
  )
  asiento.tipo = "SUELDOS"  # type: ignore[attr-defined]
  asiento.periodo = "05/2026"  # type: ignore[attr-defined]
  plan = pd.DataFrame([
      {"codigo": "42201", "descripcion": "Sueldos", "imputable": True, "usa_auxiliares": True},
      {"codigo": "21312", "descripcion": "Pasivo", "imputable": True, "usa_auxiliares": True},
  ])
  nombre = "_test_tipo_sueldos_tmp.xlsx"
  ruta = proc.generar_excel_tango_nativo(
      [asiento], "Test", "30714058386", 5, 2026,
      nombre_archivo=nombre,
      plan_cuentas=plan,
  )
  try:
    wb = openpyxl.load_workbook(str(ruta), read_only=True)
    ws = wb["Asientos contables"]
    fila = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))[0]
    wb.close()
    assert fila[3] == "SUELDOS", f"tipo asiento esperado SUELDOS, obtuvo {fila[3]!r}"
  finally:
    Path(ruta).unlink(missing_ok=True)
  print("OK test_generar_excel_tango_nativo_tipo_sueldos")


def test_export_tango_advierte_cuenta_auxiliar_42201():
  """Cuentas con auxiliares contables generan advertencia, no bloqueo."""
  import procesador as proc
  from datetime import date

  plan = pd.DataFrame([
      {
          "codigo": "42201",
          "descripcion": "Sueldos y Jornales",
          "imputable": True,
          "usa_auxiliares": True,
      },
      {
          "codigo": "21309",
          "descripcion": "Sueldos y Jornales a Pagar",
          "imputable": True,
          "usa_auxiliares": True,
      },
  ])
  asiento = proc.AsientoDevengamiento(
      identificador=1,
      concepto="Devengamiento SUELDOS",
      fecha=date(2026, 5, 31),
      renglones=[
          proc.RenglonAsiento("42201", "Sueldos y Jornales", debe=1000000.0, haber=0),
          proc.RenglonAsiento("21309", "Sueldos a Pagar", debe=0, haber=1000000.0),
      ],
  )
  informe = proc.auditar_exportacion_tango([asiento], plan)
  assert not informe["bloqueantes"]
  assert len(informe["advertencias"]) == 2
  assert proc.auditar_renglones_imputables_tango([asiento], plan) == []
  print("OK test_export_tango_advierte_cuenta_auxiliar_42201")


def test_plan_tango_habilitado_marca_imputable_hoja():
  """Plan exportado de Tango: Habilitado=Si sin hijas → imputable."""
  import procesador as proc
  from pathlib import Path

  ruta = Path("data/planes_cuentas/plan_30714058386.xlsx")
  if not ruta.is_file():
    print("SKIP test_plan_tango_habilitado_marca_imputable_hoja")
    return
  plan = proc.cargar_plan_cuentas(ruta)
  for cod in ("42201", "21316", "21317", "21307"):
    fila = plan[plan["codigo"].astype(str) == cod]
    assert not fila.empty
    assert bool(fila.iloc[0]["imputable"]), f"{cod} debería ser imputable"
  print("OK test_plan_tango_habilitado_marca_imputable_hoja")


def test_preparar_asientos_export_tango_pipeline():
  """Pipeline pre-export normaliza códigos, fechas y balancea al centavo."""
  import procesador as proc
  from datetime import date

  asiento = proc.AsientoDevengamiento(
      identificador=1,
      concepto="IVA 03/2026",
      fecha="31/03/2026",
      renglones=[
          proc.RenglonAsiento("2140101.0", "IVA Ventas", debe=500.03, haber=0),
          proc.RenglonAsiento("2140401", "IVA a Pagar", debe=0, haber=500.00),
      ],
  )
  plan = pd.DataFrame([
      {"codigo": "2140101", "descripcion": "IVA Ventas 21%", "imputable": True},
      {"codigo": "2140401", "descripcion": "IVA a Pagar", "imputable": True},
  ])
  preparados = proc.preparar_asientos_export_tango([asiento], plan)
  assert len(preparados) == 1
  p = preparados[0]
  assert p.balanceado
  assert p.renglones[0].codigo_cuenta == "2140101"
  assert isinstance(p.fecha, date)
  print("OK test_preparar_asientos_export_tango_pipeline")


def test_formatear_comprobante_tango_monotributo():
  import procesador as proc

  assert proc.formatear_comprobante_tango(2, 142) == "00002-00000142"
  assert proc.formatear_comprobante_tango("10", "94") == "00010-00000094"
  print("OK test_formatear_comprobante_tango_monotributo")


def test_parsear_factura_afip_texto_servicios():
  import procesador as proc

  texto = """
  FACTURA C
  Codigo: 011
  Fecha de Emisión: 15/05/2025
  Punto de Venta: 2
  Comp. Nro: 142
  Concepto: 2 - Servicios
  Período Facturado Desde: 01/05/2025 Hasta: 31/05/2025
  Importe Total: $ 150.000,50
  CAE: 12345678901234
  """
  parsed = proc.parsear_factura_afip_texto(texto, "factura_test.pdf")
  assert parsed is not None
  assert parsed["Período Desde"] == "01/05/2025"
  assert parsed["Período Hasta"] == "31/05/2025"
  assert parsed["Concepto"] == "2 - Servicios"
  assert abs(parsed["Importe Total"] - 150000.50) < 0.02
  assert parsed["Comprobante"] == "00002-00000142"
  assert parsed["Tipo"].startswith("Factura")
  assert parsed["Importe Total"] > 0
  print("OK test_parsear_factura_afip_texto_servicios")


def test_parsear_nota_credito_afip_importe_negativo():
  import procesador as proc

  texto = """
  NOTA DE CREDITO C
  Codigo: 013
  Fecha de Emisión: 20/05/2025
  Punto de Venta: 2
  Comp. Nro: 15
  Concepto: 2 - Servicios
  Período Facturado Desde: 01/05/2025 Hasta: 31/05/2025
  Importe Total: $ 25.000,00
  CAE: 99998888777766
  """
  parsed = proc.parsear_factura_afip_texto(texto, "nc_test.pdf")
  assert parsed is not None
  assert "Crédito" in parsed["Tipo"] or "Credito" in parsed["Tipo"]
  assert parsed["Código AFIP"] == "013"
  assert abs(parsed["Importe Total"] + 25000.0) < 0.02
  print("OK test_parsear_nota_credito_afip_importe_negativo")


def test_procesar_facturas_monotributo_orden_cronologico():
  import io
  import zipfile
  import unittest.mock as mock
  import procesador as proc

  class _FakeUpload:
    def __init__(self, name: str, data: bytes):
      self.name = name
      self._data = data

    def getvalue(self):
      return self._data

  t1 = """
  FACTURA C
  Codigo: 011
  Fecha de Emisión: 01/06/2025
  Punto de Venta: 1
  Comp. Nro: 10
  Concepto: 2 - Servicios
  Período Facturado Desde: 01/06/2025 al 30/06/2025
  Importe Total: 10.000,00
  CAE: 11112222333344
  """
  t2 = """
  FACTURA C
  Codigo: 011
  Fecha de Emisión: 01/04/2025
  Punto de Venta: 1
  Comp. Nro: 8
  Concepto: 2 - Servicios
  Período Facturado Desde: 01/04/2025 al 30/04/2025
  Importe Total: 8.000,00
  CAE: 55556666777788
  """
  t_nc = """
  NOTA DE CREDITO C
  Codigo: 013
  Fecha de Emisión: 10/06/2025
  Punto de Venta: 1
  Comp. Nro: 2
  Concepto: 2 - Servicios
  Período Facturado Desde: 01/06/2025 al 30/06/2025
  Importe Total: 2.000,00
  CAE: 12121212121212
  """
  # Mismo CAE que t1 pero bytes distintos — debe descartarse por dedupe de CAE
  t1_dup = t1 + "\n(reimpresion)\n"
  buf = io.BytesIO()
  with zipfile.ZipFile(buf, "w") as zf:
    zf.writestr("a.pdf", t1.encode("utf-8"))
    zf.writestr("b.pdf", t2.encode("utf-8"))
    zf.writestr("nc.pdf", t_nc.encode("utf-8"))
    zf.writestr("a_dup.pdf", t1_dup.encode("utf-8"))
  uploads = [_FakeUpload("lote.zip", buf.getvalue())]

  def _fake_extraer(pdf_bytes: bytes) -> str:
    return pdf_bytes.decode("utf-8")

  with mock.patch.object(proc, "extraer_texto_factura_afip", side_effect=_fake_extraer):
    df, errores = proc.procesar_facturas_monotributo(uploads)
  assert any("duplicado" in str(e.get("motivo", "")).lower() for e in errores)
  assert len(df) == 3
  assert df.iloc[0]["Período Desde"] == "01/04/2025"
  assert round(float(df["Importe Total"].sum()), 2) == 16000.0  # 8k + 10k - 2k
  assert (df["Importe Total"] < 0).sum() == 1
  xlsx = proc.exportar_monotributo_excel(df)
  assert len(xlsx) > 100
  print("OK test_procesar_facturas_monotributo_orden_cronologico")


if __name__ == "__main__":
    test_ventas_21_debe()
    test_compras_21_haber()
    test_loop_review_ajusta_ventas()
    test_columnas_plan_mayusculas()
    test_segregacion_dual_alicuotas()
    test_no_cuenta_generica_si_hay_especifica()
    test_alias_debito_sin_palabra_ventas()
    test_carga_desde_csv_en_disco()
    test_liquidacion_desde_dataframe_arca()
    test_no_falso_positivo_caja_en_debito_21()
    test_no_falso_positivo_retencion_iibb()
    test_map_iva_debito_fiscal_generico_tango()
    test_rescate_filtra_por_concepto()
    test_periodo_planilla_fecha_fin_mes()
    test_planilla_iva_posicional()
    test_posicion_multi_saldo_saldo_favor()
    test_normalizar_plan_columnas_duplicadas()
    test_planilla_iva_etiquetas_con_27()
    test_balance_solapa_dinamica()
    test_planilla_iibb_posicion()
    test_balance_solapa_iibb()
    test_balance_unc_path()
    test_sanitizar_ruta_unc_comillas()
    test_tax_registry_solapas_sueldos_tish()
    test_detectar_periodos_en_balance_df()
    test_resolver_indice_columna_periodo()
    test_planilla_iva_etiquetas_periodo_mensual()
    test_periodo_siguiente()
    test_match_fechas_completas_encabezados()
    test_filtro_ruido_totales_balance()
    test_extractor_universal_codigo_y_tipo()
    test_extractor_universal_filas_balance()
    test_extractor_banco_sincronico_fila_por_fila()
    test_extractor_banco_partida_doble_columnas()
    test_resolver_cuenta_banco_hibrida()
    test_conciliar_banco_con_tango_match_monto_y_comision()
    test_formateo_periodo_y_fecha_grilla()
    test_balance_ruta_relativa_proyecto()
    test_filtro_ruido_posicion_ejercicio()
    test_extractor_oftalmologia_iva_real()
    test_match_meses_texto_cabeceras_sueldos()
    test_extractor_sueldos_pasivos_negativos()
    test_columna_periodo_estricta_sin_desfase()
    test_sueldos_ofthalmologia_sin_desfase_primeras_filas()
    test_escaneo_cabecera_seguro_texto_institucional()
    test_matching_flexible_marzo_cierre()
    test_export_tango_fecha_txt_aaaammdd()
    test_export_tango_balancea_centavos()
    test_export_tango_rechaza_cuenta_madre()
    test_codigo_tipo_asiento_tango_mapping()
    test_generar_excel_tango_nativo_tipo_sueldos()
    test_export_tango_advierte_cuenta_auxiliar_42201()
    test_plan_tango_habilitado_marca_imputable_hoja()
    test_preparar_asientos_export_tango_pipeline()
    test_formatear_comprobante_tango_monotributo()
    test_parsear_factura_afip_texto_servicios()
    test_parsear_nota_credito_afip_importe_negativo()
    test_procesar_facturas_monotributo_orden_cronologico()
    print("\nTodos los tests pasaron.")
