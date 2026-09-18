# -*- coding: utf-8 -*-
"""Vista AE-Studio sobre las reglas que ya tiene la web."""
from __future__ import annotations

import unittest

from motor_conciliacion import (
    bucket_ae,
    clasificar,
    correr_motor,
    origen_linea_extracto,
    renglones_asiento_banco_mes,
)


class TestBucketAeUsaReglasWeb(unittest.TestCase):
    def test_ley_25413_es_retencion(self):
        clf = clasificar("IMP. DEB. LEY 25413 GRAL.", banco="", debito=100, credito=0)
        vista = bucket_ae(
            {
                "tipo": clf["tipo"],
                "categoria": clf["categoria"],
                "descripcion": "IMP. DEB. LEY 25413 GRAL.",
                "debito": 100,
                "credito": 0,
            },
            extracto_label="Impuestos a los débitos y créditos",
        )
        self.assertEqual(vista, "retencion")
        self.assertEqual(clf["fuente"], "regla_local")

    def test_transferencia_terceros_es_ingreso(self):
        clf = clasificar("TRANSFERENCIA DE TERCEROS VENOM S.A.", banco="", debito=0, credito=51000)
        vista = bucket_ae(
            {
                "tipo": clf["tipo"],
                "categoria": clf["categoria"],
                "descripcion": "TRANSFERENCIA DE TERCEROS VENOM S.A.",
                "debito": 0,
                "credito": 51000,
            },
            extracto_label="Transferencias recibidas",
        )
        self.assertEqual(clf["tipo"], "INGRESO")
        self.assertEqual(vista, "ingreso")

    def test_fci_es_intercuenta_aunque_el_seed_diga_ingreso(self):
        clf = clasificar("RESCATE FIMA PREMIUM", banco="", debito=0, credito=1000)
        vista = bucket_ae(
            {
                "tipo": clf["tipo"],
                "categoria": clf["categoria"],
                "descripcion": "RESCATE FIMA PREMIUM",
                "debito": 0,
                "credito": 1000,
            },
            extracto_label="Rescate FIMA",
        )
        self.assertEqual(vista, "inter-cta")

    def test_motor_sigue_pidiendo_confirmacion_en_regla_local(self):
        out = correr_motor(
            [
                {
                    "fecha": "10/09/2026",
                    "descripcion": "IMP. DEB. LEY 25413 $ 10",
                    "debito": 10,
                    "credito": 0,
                    "saldo": 0,
                }
            ],
            None,
            [],
            [],
            cliente_id=1,
            banco="",
            periodo=None,
            saldo_ok=True,
        )
        self.assertEqual(out[0]["estado"], "PENDIENTE")
        self.assertEqual(bucket_ae(out[0]), "retencion")


class TestAsientoExtractoAgrupaCuentas(unittest.TestCase):
    def test_partida_doble_contra_banco(self):
        rows = renglones_asiento_banco_mes(
            [
                {
                    "debito": 100,
                    "credito": 0,
                    "cuenta_codigo": "52201",
                    "cuenta_plan": "Gastos bancarios",
                },
                {
                    "debito": 0,
                    "credito": 50,
                    "cuenta_codigo": "11301",
                    "cuenta_plan": "Deudores",
                },
            ],
            codigo_banco="11104",
            descripcion_banco="Banco Galicia",
            periodo="09/2026",
            fecha_str="30/09/2026",
        )
        by_cod = {r["Código"]: r for r in rows}
        self.assertEqual(by_cod["52201"]["Debe"], 100)
        self.assertEqual(by_cod["11301"]["Haber"], 50)
        self.assertEqual(by_cod["11104"]["Haber"], 50)
        debe = sum(r["Debe"] for r in rows)
        haber = sum(r["Haber"] for r in rows)
        self.assertAlmostEqual(debe, haber, places=2)

    def test_origen_sin_regla_es_a_clasificar(self):
        self.assertEqual(
            origen_linea_extracto(
                {"fuente": "", "categoria": "Movimientos a identificar", "cuenta_codigo": "99999"}
            ),
            "a_clasificar",
        )
        self.assertEqual(
            origen_linea_extracto(
                {"fuente": "regla_local", "categoria": "Comis. y Gtos Bcarios.", "cuenta_codigo": "52201"}
            ),
            "regla",
        )


if __name__ == "__main__":
    unittest.main()
