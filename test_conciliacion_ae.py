# -*- coding: utf-8 -*-
"""Conciliación: reglas de la web, papeles mensuales, match 1:N, asiento 99999."""
from __future__ import annotations

import unittest
from datetime import date
from decimal import Decimal

import pandas as pd

from motor_conciliacion import (
    armar_papel_mes,
    bucket_ae,
    clasificar,
    correr_motor,
    df_extracto_a_filas,
    match_proveedor,
    movimientos_a_filas_grilla_tango,
    origen_linea_extracto,
    papeles_por_mes,
    renglones_asiento_banco_mes,
    validar_saldos_corridos,
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


class TestSaldosYSignos(unittest.TestCase):
    def test_excel_sin_saldo_no_se_marca_inconsistente(self):
        filas = [
            {"credito": 100, "debito": 0, "saldo": 0},
            {"credito": 0, "debito": 40, "saldo": 0},
        ]
        ok, msg = validar_saldos_corridos(filas)
        self.assertTrue(ok)
        self.assertEqual(msg, "")

    def test_cadena_de_saldos_detecta_salto(self):
        filas = [
            {"credito": 100, "debito": 0, "saldo": 1100},
            {"credito": 0, "debito": 40, "saldo": 2000},
        ]
        ok, msg = validar_saldos_corridos(filas)
        self.assertFalse(ok)
        self.assertIn("inconsistente", msg.lower())

    def test_ncc_sin_dc_va_a_credito(self):
        df = pd.DataFrame(
            [
                {
                    "Fecha": "02/01/2026",
                    "Descripcion": "NCC REINTEGRO COMISION",
                    "Debito": 0,
                    "Credito": 0,
                    "Importe": 250.5,
                    "Saldo": 0,
                    "Tipo Movimiento": "",
                    "Banco": "Galicia",
                }
            ]
        )
        filas = df_extracto_a_filas(df)
        self.assertEqual(len(filas), 1)
        self.assertEqual(filas[0]["credito"], Decimal("250.50"))
        self.assertEqual(filas[0]["debito"], Decimal("0.00"))

    def test_importe_negativo_va_a_debito(self):
        df = pd.DataFrame(
            [
                {
                    "Fecha": "02/01/2026",
                    "Descripcion": "PAGO SERVICIO",
                    "Debito": 0,
                    "Credito": 0,
                    "Importe": -80,
                    "Saldo": 0,
                    "Banco": "Galicia",
                }
            ]
        )
        filas = df_extracto_a_filas(df)
        self.assertEqual(filas[0]["debito"], Decimal("80.00"))
        self.assertEqual(filas[0]["credito"], Decimal("0.00"))

    def test_fila_sin_importe_se_omite(self):
        df = pd.DataFrame(
            [
                {
                    "Fecha": "02/01/2026",
                    "Descripcion": "ENCABEZADO",
                    "Debito": 0,
                    "Credito": 0,
                    "Importe": 0,
                    "Saldo": 0,
                }
            ]
        )
        self.assertEqual(df_extracto_a_filas(df), [])


class TestPapelesMensuales(unittest.TestCase):
    def test_primer_mes_abre_con_extracto_y_cierre_formula(self):
        movs = [
            {
                "fecha": date(2026, 1, 2),
                "credito": 100,
                "debito": 0,
                "saldo": 1100,
            },
            {
                "fecha": date(2026, 1, 5),
                "credito": 0,
                "debito": 40,
                "saldo": 1060,
            },
        ]
        papel = armar_papel_mes(movs)
        self.assertEqual(papel["origen_apertura"], "extracto")
        self.assertEqual(papel["apertura"], 1000.0)
        self.assertEqual(papel["creditos"], 100.0)
        self.assertEqual(papel["debitos"], 40.0)
        self.assertEqual(papel["cierre"], 1060.0)
        self.assertEqual(papel["segun_resumen"], 1060.0)
        self.assertEqual(papel["diferencia"], 0.0)

    def test_mes_siguiente_arrastra_y_no_fuerza_diferencia_a_cero(self):
        movs = [
            {
                "fecha": date(2026, 1, 2),
                "credito": 100,
                "debito": 0,
                "saldo": 1100,
            },
            {
                "fecha": date(2026, 1, 31),
                "credito": 0,
                "debito": 40,
                "saldo": 1060,
            },
            {
                "fecha": date(2026, 2, 3),
                "credito": 200,
                "debito": 0,
                "saldo": 1260,
            },
            {
                "fecha": date(2026, 2, 28),
                "credito": 0,
                "debito": 30,
                "saldo": 1225,
            },
        ]
        cadena = papeles_por_mes(movs)
        self.assertEqual(len(cadena), 2)
        ene, feb = cadena
        self.assertEqual(ene["origen_apertura"], "extracto")
        self.assertEqual(ene["cierre"], 1060.0)
        self.assertEqual(feb["origen_apertura"], "arrastre")
        self.assertEqual(feb["apertura"], 1060.0)
        self.assertEqual(feb["cierre"], 1230.0)
        self.assertEqual(feb["segun_resumen"], 1225.0)
        self.assertEqual(feb["diferencia"], -5.0)
        self.assertNotEqual(feb["diferencia"], 0.0)

    def test_no_inventa_mes_sin_movimientos(self):
        movs = [
            {"fecha": date(2026, 1, 2), "credito": 10, "debito": 0, "saldo": 10},
            {"fecha": date(2026, 3, 2), "credito": 5, "debito": 0, "saldo": 15},
        ]
        cadena = papeles_por_mes(movs)
        self.assertEqual([p["mes"] for p in cadena], [1, 3])
        self.assertEqual(cadena[1]["origen_apertura"], "extracto")


class TestMatchProveedor1N(unittest.TestCase):
    def test_un_pago_cubre_dos_facturas(self):
        mov = {
            "fecha": date(2026, 3, 10),
            "descripcion": "TRF INMED PROVEED / ACME SA / 30708982497",
            "debito": 150,
        }
        pendientes = [
            {
                "id": 1,
                "razon_social": "ACME SA",
                "importe": 100,
                "fecha": date(2026, 3, 1),
                "tipo_comp": "FC",
                "num_comp": "A-1",
            },
            {
                "id": 2,
                "razon_social": "ACME SA",
                "importe": 50,
                "fecha": date(2026, 3, 2),
                "tipo_comp": "FC",
                "num_comp": "A-2",
            },
            {
                "id": 3,
                "razon_social": "OTRA SRL",
                "importe": 150,
                "fecha": date(2026, 3, 1),
                "tipo_comp": "FC",
                "num_comp": "B-1",
            },
        ]
        hit = match_proveedor(mov, pendientes)
        self.assertIsNotNone(hit)
        ids = {f["id"] for f in hit["facturas"]}
        self.assertEqual(ids, {1, 2})
        self.assertIn("1:N", hit["detalle"])

    def test_motor_marca_ambas_facturas_usadas(self):
        filas = [
            {
                "fecha": date(2026, 3, 10),
                "descripcion": "TRF INMED PROVEED / ACME SA / 30708982497",
                "debito": 150,
                "credito": 0,
                "saldo": 0,
            }
        ]
        prov = [
            {
                "id": 1,
                "razon_social": "ACME SA",
                "importe": 90,
                "fecha": date(2026, 3, 1),
                "tipo_comp": "FC",
                "num_comp": "1",
                "usado": False,
            },
            {
                "id": 2,
                "razon_social": "ACME SA",
                "importe": 60,
                "fecha": date(2026, 3, 2),
                "tipo_comp": "FC",
                "num_comp": "2",
                "usado": False,
            },
        ]
        reglas = [
            {
                "patron": "TRF INMED PROVEED",
                "categoria": "Proveedores varios (a conciliar por importe)",
                "tipo": "DEBITO_PROVEEDOR",
                "orden": 0,
                "activo": 1,
            }
        ]
        out = correr_motor(
            filas, reglas, prov, [], cliente_id=1, banco="", periodo=None, saldo_ok=True
        )
        self.assertEqual(out[0]["estado"], "CONCILIADO")
        self.assertTrue(prov[0]["usado"])
        self.assertTrue(prov[1]["usado"])


class TestAsientoTango(unittest.TestCase):
    def test_pendiente_sin_match_aparece_como_99999(self):
        movs = [
            {
                "fecha": date(2026, 3, 10),
                "descripcion": "MOVIMIENTO RARO XYZ",
                "categoria": "Movimientos a identificar",
                "tipo": "DEBITO_REVISAR",
                "estado": "PENDIENTE",
                "credito": 0,
                "debito": 25,
            }
        ]
        filas = movimientos_a_filas_grilla_tango(movs)
        self.assertEqual(len(filas), 1)
        self.assertEqual(filas[0]["cuenta_sugerida"], "99999")
        self.assertEqual(filas[0]["debe"], 25.0)

    def test_regla_local_confirmable_entra_al_asiento(self):
        plan = pd.DataFrame(
            [
                {"codigo": "51201", "descripcion": "Imp. Déb/Cred. Ley 25413 (sobre débitos)"},
            ]
        )
        movs = [
            {
                "fecha": date(2026, 3, 10),
                "descripcion": "IMP. DEB. LEY 25413",
                "categoria": "Imp. Déb/Cred. Ley 25413 (sobre débitos)",
                "tipo": "DEBITO_IMPUESTO",
                "estado": "PENDIENTE",
                "credito": 0,
                "debito": 10,
            }
        ]
        filas = movimientos_a_filas_grilla_tango(movs, plan_cuentas=plan)
        self.assertEqual(len(filas), 1)
        self.assertEqual(filas[0]["cuenta_sugerida"], "51201")
        self.assertEqual(filas[0]["estado"], "PENDIENTE")


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


class TestExtractoComponente(unittest.TestCase):
    def test_aplica_cuenta_manual_antes_del_asiento(self):
        from ui_conciliacion_ae import _aplicar_filas_componente

        movs = [
            {
                "_idx": 0,
                "cuenta_codigo": "99999",
                "categoria": "Identificar",
                "origen": "a_clasificar",
            }
        ]
        out = _aplicar_filas_componente(
            movs,
            [{"i": 0, "codigo": "54101", "clasif": "Honorarios", "origen": "sugerido"}],
            None,
        )
        self.assertEqual(out[0]["cuenta_codigo"], "54101")
        self.assertEqual(out[0]["origen"], "sugerido")
        self.assertEqual(out[0]["categoria"], "Honorarios")

    def test_plan_lista_todas_las_cuentas(self):
        import pandas as pd
        from ui_conciliacion_ae import _opciones_plan

        plan = pd.DataFrame(
            {"codigo": ["11104", "52101", "53305"], "descripcion": ["Banco", "Gastos", "IDC"]}
        )
        opts = _opciones_plan(plan)
        self.assertTrue(any(o.startswith("11104") for o in opts))
        self.assertTrue(any(o.startswith("52101") for o in opts))
        self.assertGreaterEqual(len(opts), 4)

    def test_movimientos_iguales_quedan_en_la_misma_cuenta(self):
        from ui_conciliacion_ae import _propagar_cuentas_repetidas

        movs = [
            {"descripcion": "Pago haberes", "categoria": "Pago haberes", "cuenta_codigo": "52120", "origen": "regla"},
            {"descripcion": "Pago haberes", "categoria": "Pago haberes", "cuenta_codigo": "99999", "origen": "a_clasificar"},
            {"descripcion": "Pago haberes 2509025072", "categoria": "Pago haberes", "cuenta_codigo": "99999", "origen": "a_clasificar"},
        ]
        out = _propagar_cuentas_repetidas(movs)
        self.assertEqual(out[1]["cuenta_codigo"], "52120")
        self.assertEqual(out[2]["cuenta_codigo"], "52120")


if __name__ == "__main__":
    unittest.main()
