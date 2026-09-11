"""OCR 11.xxx leído como 71.xxx no debe fabricar débitos de 60 millones."""
from __future__ import annotations

import unittest

from procesador import (
    _corregir_filas_extracto_por_saldos,
    _elegir_monto_y_saldo_extracto,
    _resolver_saldo_ocr,
)


class TestOcrSaldo71Vs11(unittest.TestCase):
    def test_resuelve_71_a_11_si_cierra_el_debito(self):
        # D Amico 30/10: movimiento 272.250, saldo OCR 71.201.888,77, previo 11.474.138,77
        saldo = _resolver_saldo_ocr(71_201_888.77, 11_474_138.77, 272_250.0)
        self.assertEqual(saldo, 11_201_888.77)

    def test_elige_272250_y_no_60_millones(self):
        mov, saldo = _elegir_monto_y_saldo_extracto(
            [272_250.0, 71_201_888.77],
            saldo_prev=11_474_138.77,
        )
        self.assertEqual(mov, -272_250.0)
        self.assertEqual(saldo, 11_201_888.77)

    def test_cadena_corrige_debito_60m_pegado(self):
        # Cache Gastro: crédito Lab Dominguez + débito D Amico
        movs = [
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "IVA percepcion",
                "Importe": -1599.0,
                "Debito": 1599.0,
                "Credito": None,
                "Saldo": 10_779_874.30,
            },
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "Pago a proveedores recibido",
                "Importe": 60_694_264.47,
                "Debito": None,
                "Credito": 60_694_264.47,
                "Saldo": 71_474_138.77,
            },
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "TRANSFERENCIA REALIZADA",
                "Importe": -60_272_250.0,
                "Debito": 60_272_250.0,
                "Credito": None,
                "Saldo": 11_201_888.77,
            },
        ]
        out = _corregir_filas_extracto_por_saldos(movs)
        self.assertAlmostEqual(float(out[1]["Credito"]), 694_264.47, places=2)
        self.assertAlmostEqual(float(out[1]["Saldo"]), 11_474_138.77, places=2)
        self.assertAlmostEqual(float(out[2]["Debito"]), 272_250.0, places=2)
        self.assertAlmostEqual(float(out[2]["Saldo"]), 11_201_888.77, places=2)

    def test_alquiler_royo_960000(self):
        mov, saldo = _elegir_monto_y_saldo_extracto(
            [960_000.0, 10_854_201.36],
            saldo_prev=11_814_201.36,
        )
        self.assertEqual(mov, -960_000.0)
        self.assertEqual(saldo, 10_854_201.36)

    def test_cadena_arba_71m_luego_royo(self):
        movs = [
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "Retencion arba",
                "Importe": 59_942_823.90,
                "Debito": None,
                "Credito": 59_942_823.90,
                "Saldo": 71_814_201.36,
            },
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "Transferencia inmediata",
                "Importe": -60_960_000.0,
                "Debito": 60_960_000.0,
                "Credito": None,
                "Saldo": 10_854_201.36,
            },
        ]
        out = _corregir_filas_extracto_por_saldos(movs)
        self.assertAlmostEqual(float(out[0]["Saldo"]), 11_814_201.36, places=2)
        self.assertAlmostEqual(float(out[1]["Debito"]), 960_000.0, places=2)

    def test_usa_importe_chico_de_la_leyenda(self):
        movs = [
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "IVA",
                "Importe": -100.0,
                "Debito": 100.0,
                "Credito": None,
                "Saldo": 17_589_304.15,
            },
            {
                "Tipo fila": "Movimiento",
                "Descripcion": "TRANSFERENCIA REALIZADA",
                "Detalle": "2.853,47 pensar digital srl varios var 30715823663",
                "Importe": -17_589_204.15,
                "Debito": 17_589_204.15,
                "Credito": None,
                "Saldo": 17.44,
            },
        ]
        out = _corregir_filas_extracto_por_saldos(movs)
        self.assertAlmostEqual(float(out[1]["Debito"]), 2853.47, places=2)


if __name__ == "__main__":
    unittest.main()
