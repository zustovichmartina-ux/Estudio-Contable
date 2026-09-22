"""Parseo por Y: últimos 2 números = importe + saldo."""
from __future__ import annotations

import unittest

from extracto_layout import parse_lineas, parsear_lineas, score_cadena
from imputacion_bancaria import imputar_extracto


class TestExtractoLayout(unittest.TestCase):
    def test_linea_galicia_credito_y_saldo(self):
        lineas = [
            "01/04/26 TRANSFERENCIA DE TERCEROS VENOM S.A. 51.053.280,00 88.597.393,35",
            "01/04/26 ING. BRUTOS S/ CRED LEY 12837-BS.AS. 1.276.332,00- 87.321.061,35",
        ]
        movs = parse_lineas(lineas)
        self.assertEqual(len(movs), 2)
        self.assertAlmostEqual(movs[0]["credito"], 51053280.00, places=2)
        self.assertAlmostEqual(movs[0]["saldo"], 88597393.35, places=2)
        self.assertAlmostEqual(movs[1]["debito"], 1276332.00, places=2)
        self.assertEqual(score_cadena(movs), 1)

    def test_elige_mejor_cadena(self):
        lineas = [
            "01/04/26 ANULACION DE TRANSFERENCIA 53.073,71 37.544.113,35",
            "01/04/26 TRANSFERENCIA DE TERCEROS 51.053.280,00 88.597.393,35",
        ]
        movs = parsear_lineas(lineas)
        self.assertGreaterEqual(len(movs), 2)

    def test_imputacion_fija_ley_25413(self):
        out = imputar_extracto(
            [
                {
                    "fecha": "01/04/26",
                    "descripcion": "IMP. DEB. LEY 25413 GRAL.",
                    "detalle": "",
                    "credito": 0,
                    "debito": 7657.99,
                }
            ]
        )
        self.assertEqual(out["movimientos"][0]["origen_imputacion"], "fija")
        self.assertFalse(out["movimientos"][0]["editable"])

    def test_omite_saldo_anterior_y_une_debin_con_detalle(self):
        lineas = [
            "01/08/26 SALDO ANTERIOR 34.154.678,27 34.154.678,27",
            "02/08/26 CR.DEBIN 15.000,00 34.169.678,27",
            "CLINICA DEL SOL 30717847810",
        ]
        movs = parse_lineas(lineas)
        self.assertEqual(len(movs), 1)
        self.assertIn("CR.DEBIN", movs[0]["descripcion"].upper())
        self.assertIn("CLINICA", movs[0]["descripcion"].upper())
        self.assertIn("30717847810", movs[0]["descripcion"])


if __name__ == "__main__":
    unittest.main()

