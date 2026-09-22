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

    def test_saldo_anterior_y_final_se_omiten(self):
        df = pd.DataFrame(
            [
                {
                    "Fecha": "30/12/2025",
                    "Descripcion": "SALDO ANTERIOR",
                    "Detalle": "nan",
                    "Debito": 0,
                    "Credito": 1327591.30,
                    "Saldo": 1327591.30,
                },
                {
                    "Fecha": "02/01/2026",
                    "Descripcion": "CR.DEBIN 31/12",
                    "Detalle": "nan",
                    "Debito": 0,
                    "Credito": 5941.92,
                    "Saldo": 1333533.22,
                },
                {
                    "Fecha": "31/01/2026",
                    "Descripcion": "SALDO FINAL",
                    "Debito": 0,
                    "Credito": 1333533.22,
                    "Saldo": 1333533.22,
                },
            ]
        )
        filas = df_extracto_a_filas(df)
        self.assertEqual(len(filas), 1)
        self.assertIn("CR.DEBIN", filas[0]["descripcion"])
        self.assertNotIn("nan", filas[0]["descripcion"].lower())
        self.assertNotIn("Saldo", filas[0]["descripcion"])

    def test_sin_descripcion_se_omite(self):
        df = pd.DataFrame(
            [
                {
                    "Fecha": "01/08/2026",
                    "Descripcion": "Sin descripción",
                    "Debito": 34154678.27,
                    "Credito": 0,
                    "Saldo": 34154678.27,
                },
                {
                    "Fecha": "02/08/2026",
                    "Descripcion": "CR.DEBIN CLINICA DEL SOL",
                    "Debito": 0,
                    "Credito": 15000,
                    "Saldo": 34169678.27,
                },
            ]
        )
        filas = df_extracto_a_filas(df)
        self.assertEqual(len(filas), 1)
        self.assertIn("CR.DEBIN", filas[0]["descripcion"])


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

    def test_engloba_por_clasificacion_y_toma_la_cuenta_del_grupo(self):
        rows = renglones_asiento_banco_mes(
            [
                {
                    "debito": 10,
                    "credito": 0,
                    "cuenta_codigo": "11408",
                    "categoria": "Impuesto a los débitos",
                },
                {
                    "debito": 20,
                    "credito": 0,
                    "cuenta_codigo": "99999",
                    "categoria": "Impuesto a los débitos",
                },
                {
                    "debito": 0,
                    "credito": 50,
                    "cuenta_codigo": "99999",
                    "categoria": "Transferencias recibidas",
                },
            ],
            codigo_banco="11104",
            descripcion_banco="Banco",
            periodo="01/2026",
            fecha_str="31/01/2026",
        )
        by_desc = {r["Descripción"]: r for r in rows}
        self.assertEqual(by_desc["Impuesto a los débitos"]["Debe"], 30)
        self.assertEqual(by_desc["Impuesto a los débitos"]["Código"], "11408")
        self.assertEqual(by_desc["Transferencias recibidas"]["Haber"], 50)

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

    def test_plan_vacio_solo_a_clasificar(self):
        import pandas as pd
        from ui_conciliacion_ae import _cuentas_componente, _opciones_plan

        opts = _opciones_plan(pd.DataFrame({"codigo": [], "descripcion": []}))
        self.assertEqual(opts, ["99999 — A clasificar"])
        self.assertEqual(_cuentas_componente(opts), [{"codigo": "99999", "label": "99999 — A clasificar"}])

    def test_cuentas_componente_trae_todo_el_plan(self):
        import pandas as pd
        from ui_conciliacion_ae import _cuentas_componente, _opciones_plan

        plan = pd.DataFrame(
            {
                "codigo": [f"{11100 + i}" for i in range(25)],
                "descripcion": [f"Cuenta {i}" for i in range(25)],
            }
        )
        cuentas = _cuentas_componente(_opciones_plan(plan))
        self.assertEqual(len(cuentas), 26)
        self.assertTrue(any(c["codigo"] == "11104" for c in cuentas))
        self.assertTrue(any(c["codigo"] == "11124" for c in cuentas))

    def test_cargar_plan_excel_sin_solapa_tango(self):
        import tempfile
        from pathlib import Path

        import pandas as pd
        import procesador as proc

        df = pd.DataFrame(
            {"Código": ["11101", "52101"], "Descripción": ["Caja", "Gastos"]}
        )
        with tempfile.TemporaryDirectory() as td:
            ruta = Path(td) / "plan.xlsx"
            df.to_excel(ruta, sheet_name="Hoja1", index=False)
            out = proc.cargar_plan_cuentas(ruta)
            self.assertTrue(proc.plan_cuentas_tiene_filas(out))
            self.assertEqual(list(out["codigo"]), ["11101", "52101"])

    def test_plantilla_tango_vacia_no_tiene_cuentas(self):
        from pathlib import Path

        import procesador as proc

        ruta = Path("data/planes_cuentas/plan_99000000015.xlsx")
        if not ruta.is_file():
            self.skipTest("sin plantilla gastro")
        df = proc.cargar_plan_cuentas(ruta)
        self.assertFalse(proc.plan_cuentas_tiene_filas(df))

    def test_plan_en_sqlite_se_recupera_sin_excel(self):
        import tempfile
        from pathlib import Path

        import pandas as pd

        import database as db
        from procesador import plan_cuentas_desde_csv, plan_cuentas_tiene_filas, serializar_plan_cuentas

        orig = db.DB_PATH
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
            db.DB_PATH = Path(td) / "plan.db"
            try:
                db.inicializar_bd()
                cid = db.crear_cliente(
                    "TEST PLAN CSV", "30999999991", "Persona Jurídica", mes_cierre_balance=12
                )
                df = pd.DataFrame(
                    {"codigo": ["11101"], "descripcion": ["Caja"], "imputable": ["S"]}
                )
                db.guardar_plan_cuentas_csv(cid, serializar_plan_cuentas(df))
                raw = db.plan_cuentas_csv_cliente(cid)
                out = plan_cuentas_desde_csv(raw)
                self.assertTrue(plan_cuentas_tiene_filas(out))
                self.assertEqual(str(out.iloc[0]["codigo"]), "11101")
            finally:
                db.DB_PATH = orig

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

    def test_clasificacion_trae_la_cuenta_asociada(self):
        from ui_conciliacion_ae import _aplicar_filas_componente

        movs = [
            {
                "_idx": 0,
                "cuenta_codigo": "99999",
                "categoria": "Movimientos a identificar",
                "origen": "a_clasificar",
                "descripcion": "RET IIBB 1",
            },
            {
                "_idx": 1,
                "cuenta_codigo": "99999",
                "categoria": "Movimientos a identificar",
                "origen": "a_clasificar",
                "descripcion": "RET IIBB 2",
            },
        ]
        out = _aplicar_filas_componente(
            movs,
            [
                {"i": 0, "codigo": "11402", "clasif": "IIBB", "origen": "sugerido"},
                {"i": 1, "codigo": "99999", "clasif": "IIBB", "origen": "a_clasificar"},
            ],
            None,
        )
        self.assertEqual(out[0]["categoria"], "IIBB")
        self.assertEqual(out[0]["cuenta_codigo"], "11402")
        self.assertEqual(out[1]["cuenta_codigo"], "11402")
        self.assertEqual(out[1]["categoria"], "IIBB")
        self.assertEqual(out[1]["origen"], "sugerido")

    def test_clasificacion_iibb_bancos_usa_la_cuenta_del_catalogo(self):
        from ui_conciliacion_ae import _aplicar_filas_componente

        movs = [
            {
                "_idx": 0,
                "cuenta_codigo": "99999",
                "categoria": "Movimientos a identificar",
                "origen": "a_clasificar",
            },
            {
                "_idx": 1,
                "cuenta_codigo": "99999",
                "categoria": "Movimientos a identificar",
                "origen": "a_clasificar",
            },
        ]
        out = _aplicar_filas_componente(
            movs,
            [
                {"i": 0, "codigo": "99999", "clasif": "Retenciones IIBB bancos", "origen": "a_clasificar"},
                {"i": 1, "codigo": "99999", "clasif": "Retenciones IIBB bancos", "origen": "a_clasificar"},
            ],
            None,
            mapa_clasif={"Retenciones IIBB bancos": "11419"},
        )
        self.assertEqual(out[0]["cuenta_codigo"], "11419")
        self.assertEqual(out[1]["cuenta_codigo"], "11419")

    def test_mapa_clasif_completa_99999(self):
        from ui_conciliacion_ae import _aplicar_mapa_clasif

        movs = [
            {"categoria": "IIBB", "cuenta_codigo": "99999", "origen": "a_clasificar"},
            {"categoria": "Gastos Bancarios", "cuenta_codigo": "42501", "origen": "regla"},
        ]
        out = _aplicar_mapa_clasif(movs, {"IIBB": "11402"})
        self.assertEqual(out[0]["cuenta_codigo"], "11402")
        self.assertEqual(out[0]["origen"], "sugerido")
        self.assertEqual(out[1]["cuenta_codigo"], "42501")

    def test_catalogo_estudio_asocia_iibb_bancos_a_11419(self):
        from clasif_cuentas_extracto import MAPA_CLASIF_ESTUDIO
        from ui_conciliacion_ae import _combinar_mapa_clasif, _overlay_mapa_clasif

        mapa = _combinar_mapa_clasif({})
        self.assertEqual(mapa["Retenciones IIBB bancos"], "11419")
        self.assertEqual(mapa["Gastos Bancarios"], MAPA_CLASIF_ESTUDIO["Gastos Bancarios"])
        overlay = _overlay_mapa_clasif({**mapa, "Gastos Bancarios": "52201"})
        self.assertEqual(overlay, {"Gastos Bancarios": "52201"})
        self.assertNotIn("Retenciones IIBB bancos", overlay)

    def test_saca_saldo_anterior_de_la_grilla(self):
        from ui_conciliacion_ae import _filas_componente, _sin_filas_saldo

        movs = [
            {"descripcion": "SALDO ANTERIOR nan", "credito": 100, "debito": 0, "saldo": 100, "_idx": 0},
            {"descripcion": "CR.DEBIN nan", "credito": 5, "debito": 0, "saldo": 105, "_idx": 1},
            {"descripcion": "SALDO FINAL", "credito": 0, "debito": 0, "saldo": 105, "_idx": 2},
        ]
        limpio = _sin_filas_saldo(movs)
        self.assertEqual(len(limpio), 1)
        self.assertEqual(limpio[0]["descripcion"], "CR.DEBIN")
        filas = _filas_componente(movs)
        self.assertEqual(len(filas), 1)
        self.assertNotIn("saldo", filas[0])


if __name__ == "__main__":
    unittest.main()
