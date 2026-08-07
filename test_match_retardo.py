"""Tests del match proveedores: nombre → monto±3 → retardo → sumas."""
from datetime import date

import pandas as pd

from procesador import (
    _score_match_debito_factura,
    matchear_debitos_con_facturas,
    _normalizar_texto,
)


def _debito(did, fecha, importe, desc):
    return {
        "debito_id": did,
        "fecha": fecha,
        "importe": importe,
        "descripcion": desc,
        "descripcion_norm": _normalizar_texto(desc),
        "comprobante": "",
    }


def _factura(fid, fecha, importe, prov):
    return {
        "factura_id": fid,
        "fecha": fecha,
        "proveedor": prov,
        "proveedor_norm": _normalizar_texto(prov),
        "importe": importe,
        "comprobante": f"FCC-{fid}",
        "tipo": "FCC",
        "cuenta": "21101",
        "descripcion": "",
        "codigo_prov": "",
        "cuit": "",
        "debe": 0.0,
        "haber": importe,
    }


def test_pase1_nombre_rechaza_fecha_lejos():
    score, motivo = _score_match_debito_factura(
        date(2026, 4, 15),
        _normalizar_texto("Transferencia realizada a OPTICA SUR SA"),
        125000.0,
        date(2026, 1, 10),
        _normalizar_texto("OPTICA SUR SA"),
        125000.0,
        modo="nombre",
    )
    assert score == 0.0
    assert motivo == "fecha_lejos"


def test_retardo_acepta_meses():
    score, motivo = _score_match_debito_factura(
        date(2026, 4, 15),
        _normalizar_texto("Transferencia realizada a OPTICA SUR SA"),
        125000.0,
        date(2026, 1, 10),
        _normalizar_texto("OPTICA SUR SA"),
        125000.0,
        modo="retardo",
    )
    assert score >= 62.0
    assert "retardo_nombre" in motivo


def test_matchear_nombre_y_retardo():
    debitos = pd.DataFrame([
        # Cerca en fecha → pase 1 nombre
        _debito(0, date(2026, 3, 5), 10000.0, "Transferencia realizada LABORATORIO RELE"),
        # 90 días después → pase 3 retardo
        _debito(1, date(2026, 4, 20), 55555.55, "Transferencia realizada OPTICA SUR SA"),
    ])
    facturas = pd.DataFrame([
        _factura(0, date(2026, 3, 3), 10000.0, "LABORATORIO RELE SRL"),
        _factura(1, date(2026, 1, 15), 55555.55, "OPTICA SUR SA"),
    ])
    res = matchear_debitos_con_facturas(debitos, facturas)
    calz = res["calzados"]
    assert len(calz) == 2
    criterios = calz["Criterio"].astype(str).tolist()
    assert any("retardo_nombre" in c for c in criterios)
    assert any("retardo_nombre" not in c and "+nombre" in c for c in criterios)
    retardo = calz[calz["Criterio"].astype(str).str.contains("retardo_nombre")].iloc[0]
    assert retardo["Dias retardo"] == (date(2026, 4, 20) - date(2026, 1, 15)).days
    resumen = res["resumen"].set_index("Concepto")["Cantidad"]
    assert int(resumen["Calzados pase 1 (nombre + monto)"]) == 1
    assert int(resumen["Calzados pase 3 (retardo, sin fecha)"]) == 1


def test_retardo_no_calza_sin_nombre():
    score, _ = _score_match_debito_factura(
        date(2026, 4, 15),
        _normalizar_texto("Transferencia realizada VARIOS"),
        125000.0,
        date(2026, 1, 10),
        _normalizar_texto("OPTICA SUR SA"),
        125000.0,
        modo="retardo",
    )
    assert score == 0.0


def test_trf_inmed_proveed_es_matcheable():
    from procesador import es_debito_pago_o_transferencia
    assert es_debito_pago_o_transferencia("TRF INMED PROVEED") is True
    assert es_debito_pago_o_transferencia("TRF INMED PROVEEDOR") is True
    assert es_debito_pago_o_transferencia("IVA 21") is False


def test_monto_fecha_2_dias_calza():
    score, motivo = _score_match_debito_factura(
        date(2026, 3, 12),
        _normalizar_texto("Transferencia realizada VARIOS"),
        88000.0,
        date(2026, 3, 10),
        _normalizar_texto("OPTICA SUR SA"),
        88000.0,
        modo="monto_fecha",
    )
    assert score >= 70.0
    assert "monto+fecha3" in motivo


def test_monto_fecha_10_dias_no_calza():
    score, motivo = _score_match_debito_factura(
        date(2026, 3, 20),
        _normalizar_texto("Transferencia realizada VARIOS"),
        88000.0,
        date(2026, 3, 10),
        _normalizar_texto("OPTICA SUR SA"),
        88000.0,
        modo="monto_fecha",
    )
    assert score == 0.0
    assert motivo == "fecha_lejos"


def test_matchear_pase2_monto_fecha():
    debitos = pd.DataFrame([
        _debito(0, date(2026, 5, 4), 12345.67, "Transferencia realizada"),
    ])
    facturas = pd.DataFrame([
        _factura(0, date(2026, 5, 2), 12345.67, "PROVEEDOR XYZ COMERCIAL SRL"),
    ])
    res = matchear_debitos_con_facturas(debitos, facturas)
    calz = res["calzados"]
    assert len(calz) == 1
    assert "monto+fecha3" in str(calz.iloc[0]["Criterio"])
    assert int(res["resumen"].set_index("Concepto")["Cantidad"]["Calzados pase 2 (monto ±3 / único ±15)"]) == 1


def test_monto_unico_10_dias():
    """Monto exacto único calza aunque esté fuera de ±3 (hasta ±15)."""
    debitos = pd.DataFrame([
        _debito(0, date(2026, 5, 15), 99999.99, "N/D Transf. MacrOnline E-set D/T"),
    ])
    facturas = pd.DataFrame([
        _factura(0, date(2026, 5, 5), 99999.99, "PROVEEDOR UNICO SA"),
        _factura(1, date(2026, 5, 5), 11111.11, "OTRO PROVEEDOR SA"),
    ])
    res = matchear_debitos_con_facturas(debitos, facturas)
    calz = res["calzados"]
    assert len(calz) == 1
    assert "monto+unico" in str(calz.iloc[0]["Criterio"])
    assert float(calz.iloc[0]["Debito banco"]) == 99999.99



def test_pase3_dos_transf_una_factura():
    debitos = pd.DataFrame([
        _debito(0, date(2026, 3, 10), 40000.0, "Transferencia realizada OPTICA SUR SA"),
        _debito(1, date(2026, 3, 12), 60000.0, "Transferencia realizada OPTICA SUR SA"),
        # distractor otro proveedor
        _debito(2, date(2026, 3, 11), 100000.0, "Transferencia realizada OTRO PROVEEDOR SA"),
    ])
    facturas = pd.DataFrame([
        _factura(0, date(2026, 2, 1), 100000.0, "OPTICA SUR SA"),
    ])
    res = matchear_debitos_con_facturas(debitos, facturas)
    calz = res["calzados"]
    assert len(calz) == 1
    assert "pase3_2transf_1fct" in str(calz.iloc[0]["Criterio"])
    assert float(calz.iloc[0]["Debito banco"]) == 100000.0
    assert float(calz.iloc[0]["Importe factura"]) == 100000.0
    assert int(res["resumen"].set_index("Concepto")["Cantidad"]["Calzados pase 4 (sumas 2↔1)"]) == 1


def test_pase3_una_transf_dos_facturas():
    debitos = pd.DataFrame([
        _debito(0, date(2026, 4, 1), 150000.0, "Transferencia realizada BRAVO EVA MARIANA"),
    ])
    facturas = pd.DataFrame([
        _factura(0, date(2026, 1, 10), 90000.0, "BRAVO EVA MARIANA"),
        _factura(1, date(2026, 2, 15), 60000.0, "BRAVO EVA MARIANA"),
    ])
    res = matchear_debitos_con_facturas(debitos, facturas)
    calz = res["calzados"]
    assert len(calz) == 1
    assert "pase3_1transf_2fct" in str(calz.iloc[0]["Criterio"])
    assert float(calz.iloc[0]["Debito banco"]) == 150000.0
    assert float(calz.iloc[0]["Importe factura"]) == 150000.0
    assert "+" in str(calz.iloc[0]["Comprobante factura"])


if __name__ == "__main__":
    test_pase1_nombre_rechaza_fecha_lejos()
    test_retardo_acepta_meses()
    test_matchear_nombre_y_retardo()
    test_retardo_no_calza_sin_nombre()
    test_trf_inmed_proveed_es_matcheable()
    test_monto_fecha_2_dias_calza()
    test_monto_fecha_10_dias_no_calza()
    test_matchear_pase2_monto_fecha()
    test_monto_unico_10_dias()
    test_pase3_dos_transf_una_factura()
    test_pase3_una_transf_dos_facturas()
    print("OK test_match_retardo")
