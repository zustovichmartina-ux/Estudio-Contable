# -*- coding: utf-8 -*-
"""Vista AE-Studio sobre las reglas que ya tiene la web."""
from __future__ import annotations

import unittest

from motor_conciliacion import bucket_ae, clasificar, correr_motor


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


if __name__ == "__main__":
    unittest.main()
