"""Cola de rutinas, sin Turso ni red. SQLite temporal."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import pytest

import database
import rutinas

ROOT = Path(__file__).resolve().parents[1]


def _db(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", tmp_path / "estudio.db")
    monkeypatch.setattr(database, "_turso_conn_obj", None)
    monkeypatch.setattr(database, "_credenciales_turso", lambda: ("", ""))


def _cli():
    ruta = ROOT / "scripts" / "cola_rutinas.py"
    spec = importlib.util.spec_from_file_location("cola_rutinas_cli", ruta)
    assert spec and spec.loader
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_tablas_idempotentes(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    rutinas.asegurar_tablas()
    pedido = rutinas.crear_pedido("fcc_monotributistas", "Marti", "2026-09")
    rutinas.asegurar_tablas()
    otra = rutinas.obtener_pedido(pedido["id"])
    assert otra is not None
    assert otra["parametros"] == "2026-09"
    with database.obtener_conexion() as conn:
        idx = conn.execute(
            "SELECT name FROM sqlite_master WHERE type = 'index' AND name = 'idx_rutina_pedidos_abierto'"
        ).fetchone()
        cols = {fila[1] for fila in conn.execute("PRAGMA table_info(rutina_pedidos)")}
    assert idx is not None
    assert cols == {
        "id",
        "rutina",
        "parametros",
        "solicitado_por",
        "creado_en",
        "estado",
        "tomado_en",
        "terminado_en",
        "resultado",
        "archivos",
        "preview",
    }


def test_agrega_preview_si_la_tabla_era_vieja(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    with database.obtener_conexion() as conn:
        conn.execute(
            """
            CREATE TABLE rutina_pedidos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rutina TEXT NOT NULL,
                parametros TEXT,
                solicitado_por TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                estado TEXT NOT NULL,
                tomado_en TEXT,
                terminado_en TEXT,
                resultado TEXT,
                archivos TEXT
            )
            """
        )
        conn.execute(
            """
            INSERT INTO rutina_pedidos (rutina, solicitado_por, creado_en, estado)
            VALUES ('fcc_monotributistas', 'Marti', '2026-10-01T15:00:00+00:00', 'OK')
            """
        )
        conn.commit()
    rutinas.asegurar_tablas()
    rutinas.asegurar_tablas()
    pedido = rutinas.obtener_pedido(1)
    assert pedido is not None
    assert pedido["preview"] is None
    assert pedido["estado"] == "OK"


def test_catalogo_requisitos_a_confirmar():
    codigos = [item["codigo"] for item in rutinas.RUTINAS]
    assert "proyecciones_ganancias_iva" in codigos
    proy = rutinas.rutina_por_codigo("proyecciones_ganancias_iva")
    assert proy is not None
    assert any("compras y ventas" in texto for texto in proy["requisitos"])
    for item in rutinas.RUTINAS:
        assert item["nombre"] and item["descripcion"] and item["requisitos"]
        for texto in item["requisitos"]:
            assert str(texto).startswith("A CONFIRMAR")


def test_inicializar_bd_engancha_la_tabla():
    texto = (ROOT / "database.py").read_text(encoding="utf-8")
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "inicializar_tablas_rutinas" in texto
    assert '"Rutinas"' in app
    assert "render_rutinas" in app


def test_no_duplica_pendiente_ni_en_curso(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    primero = rutinas.crear_pedido("control_fcc_portal_iva", "Sol")
    with pytest.raises(rutinas.PedidoDuplicado):
        rutinas.crear_pedido("control_fcc_portal_iva", "Guadi", "otro")
    rutinas.tomar_pedido(primero["id"])
    with pytest.raises(rutinas.PedidoDuplicado):
        rutinas.crear_pedido("control_fcc_portal_iva", "Sol")
    rutinas.terminar_pedido(primero["id"], "OK", "Revisado")
    segundo = rutinas.crear_pedido("control_fcc_portal_iva", "Sol", "  ")
    assert segundo["id"] != primero["id"]
    assert segundo["parametros"] is None
    assert segundo["estado"] == "PENDIENTE"


def test_dos_rutinas_abiertas_a_la_vez(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    a = rutinas.crear_pedido("aviso_bazan_bajar_archivos", "Lu")
    b = rutinas.crear_pedido("bazan_detalle_items", "Lu", "marzo")
    assert a["rutina"] != b["rutina"]
    assert rutinas.pedido_abierto(a["rutina"])["id"] == a["id"]


def test_cancelar_solo_pendiente(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    pedido = rutinas.crear_pedido("seguimiento_balances_urgencia", "Tobi")
    cancelado = rutinas.cancelar_pedido(pedido["id"])
    assert cancelado["estado"] == "CANCELADO"
    assert cancelado["terminado_en"]
    assert "Cancelado" in (cancelado["resultado"] or "")
    with pytest.raises(rutinas.EstadoInvalido):
        rutinas.cancelar_pedido(pedido["id"])
    nuevo = rutinas.crear_pedido("seguimiento_balances_urgencia", "Tobi")
    rutinas.tomar_pedido(nuevo["id"])
    with pytest.raises(rutinas.EstadoInvalido):
        rutinas.cancelar_pedido(nuevo["id"])


def test_tomar_es_una_sola_vez(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    pedido = rutinas.crear_pedido("bazan_detalle_items", "Cami")
    tomado = rutinas.tomar_pedido(pedido["id"])
    assert tomado["estado"] == "EN_CURSO"
    assert tomado["tomado_en"]
    with pytest.raises(rutinas.EstadoInvalido):
        rutinas.tomar_pedido(pedido["id"])
    with pytest.raises(rutinas.PedidoNoEncontrado):
        rutinas.tomar_pedido(99999)


def test_terminar_ok_y_error(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    pedido = rutinas.crear_pedido("fcc_monotributistas", "Agus", "2026-08")
    with pytest.raises(rutinas.EstadoInvalido):
        rutinas.terminar_pedido(pedido["id"], "OK", "todavía no")
    rutinas.tomar_pedido(pedido["id"])
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.terminar_pedido(pedido["id"], "CANCELADO", "no")
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.terminar_pedido(pedido["id"], "OK", "   ")
    cerrado = rutinas.terminar_pedido(
        pedido["id"],
        "error",
        "Portal caído",
        archivos=r"\\servidor\FCC\salida.xlsx",
    )
    assert cerrado["estado"] == "ERROR"
    assert cerrado["resultado"] == "Portal caído"
    assert cerrado["archivos"] == r"\\servidor\FCC\salida.xlsx"
    assert cerrado["terminado_en"]
    with pytest.raises(rutinas.EstadoInvalido):
        rutinas.terminar_pedido(pedido["id"], "OK", "de nuevo")


def test_nombre_vacio_y_rutina_desconocida(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.crear_pedido("fcc_monotributistas", "  ")
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.crear_pedido("no_existe", "Marti")


def test_fecha_argentina():
    assert rutinas.formatear_fecha_ar("2026-10-01T15:30:00+00:00") == "01/10/2026 12:30"
    assert rutinas.formatear_fecha_ar("") == ""
    assert rutinas.formatear_fecha_ar(None) == ""


def test_listar_por_estado(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    rutinas.crear_pedido("fcc_monotributistas", "Marti")
    otro = rutinas.crear_pedido("bazan_detalle_items", "Marti")
    rutinas.tomar_pedido(otro["id"])
    pendientes = rutinas.listar_pedidos(estado="PENDIENTE")
    assert len(pendientes) == 1
    assert pendientes[0]["rutina"] == "fcc_monotributistas"
    assert len(rutinas.listar_pedidos(estado="EN_CURSO")) == 1
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.listar_pedidos(estado="RARO")


def test_terminar_guarda_preview(tmp_path, monkeypatch):
    _db(tmp_path, monkeypatch)
    pedido = rutinas.crear_pedido("proyecciones_ganancias_iva", "Sol", "2026-09")
    rutinas.tomar_pedido(pedido["id"])
    with pytest.raises(rutinas.ErrorRutina):
        rutinas.normalizar_preview(["no"])
    ruta = r"\\TANGOSRV\Compartido\CLIENTES\ACME\Ganancias.xlsx"
    cerrado = rutinas.terminar_pedido(
        pedido["id"],
        "OK",
        "Proyección actualizada",
        preview={
            "resumen_md": "Quedó la proyección de **Ganancias**.",
            "tablas": [
                {"titulo": "Cambios", "columnas": ["Cliente", "Celda"], "filas": [["ACME", "C12"]]}
            ],
            "archivos": [ruta],
        },
    )
    assert cerrado["preview"]["resumen_md"].startswith("Quedó")
    assert cerrado["preview"]["tablas"][0]["filas"] == [["ACME", "C12"]]
    assert cerrado["archivos"] == ruta
    assert rutinas.filas_tabla_preview(cerrado["preview"]["tablas"][0]) == [
        {"Cliente": "ACME", "Celda": "C12"}
    ]
    assert rutinas.ultimo_pedido("proyecciones_ganancias_iva")["id"] == pedido["id"]


def test_cli_preview_json(tmp_path, monkeypatch, capsys):
    _db(tmp_path, monkeypatch)
    cli = _cli()
    pedido = rutinas.crear_pedido("bazan_detalle_items", "Lu")
    rutinas.tomar_pedido(pedido["id"])
    path = tmp_path / "preview.json"
    path.write_text(
        json.dumps(
            {
                "resumen_md": "Detalle **armado**.",
                "tablas": [
                    {
                        "titulo": "Ítems",
                        "columnas": ["Archivo", "Fila"],
                        "filas": [["items.xlsx", 4]],
                    }
                ],
                "archivos": [r"D:\Bazan\items.xlsx"],
            }
        ),
        encoding="utf-8",
    )
    assert (
        cli.main(
            [
                "terminar",
                str(pedido["id"]),
                "--estado",
                "OK",
                "--resultado",
                "Detalle listo",
                "--preview-json",
                str(path),
            ]
        )
        == 0
    )
    data = json.loads(capsys.readouterr().out)
    assert data["pedido"]["preview"]["tablas"][0]["titulo"] == "Ítems"
    assert data["pedido"]["archivos"] == r"D:\Bazan\items.xlsx"

    otro = rutinas.crear_pedido("bazan_detalle_items", "Lu")
    rutinas.tomar_pedido(otro["id"])
    malo = tmp_path / "malo.json"
    malo.write_text("{", encoding="utf-8")
    assert (
        cli.main(
            [
                "terminar",
                str(otro["id"]),
                "--estado",
                "OK",
                "--resultado",
                "no",
                "--preview-json",
                str(malo),
            ]
        )
        == 1
    )
    fallo = json.loads(capsys.readouterr().out)
    assert fallo["ok"] is False
    assert rutinas.obtener_pedido(otro["id"])["estado"] == "EN_CURSO"
    assert (
        cli.main(
            [
                "terminar",
                str(otro["id"]),
                "--estado",
                "ERROR",
                "--resultado",
                "sin archivo",
                "--preview-json",
                str(tmp_path / "no-esta.json"),
            ]
        )
        == 1
    )


def test_cli_listar_tomar_terminar(tmp_path, monkeypatch, capsys):
    _db(tmp_path, monkeypatch)
    cli = _cli()
    assert cli.main(["listar", "--estado", "PENDIENTE"]) == 0
    assert json.loads(capsys.readouterr().out) == []

    pedido = rutinas.crear_pedido("aviso_bazan_bajar_archivos", "Hernan", "hoy")
    assert cli.main(["listar", "--estado", "PENDIENTE"]) == 0
    lista = json.loads(capsys.readouterr().out)
    assert lista[0]["id"] == pedido["id"]
    assert lista[0]["parametros"] == "hoy"
    assert lista[0]["nombre"] == "Aviso Bazan bajar archivos"

    assert cli.main(["tomar", str(pedido["id"])]) == 0
    tomado = json.loads(capsys.readouterr().out)
    assert tomado["ok"] is True
    assert tomado["pedido"]["estado"] == "EN_CURSO"

    assert cli.main(["tomar", str(pedido["id"])]) == 1
    fallo = json.loads(capsys.readouterr().out)
    assert fallo["ok"] is False
    assert "PENDIENTE" in fallo["error"]

    assert (
        cli.main(
            [
                "terminar",
                str(pedido["id"]),
                "--estado",
                "OK",
                "--resultado",
                "Archivos bajados",
                "--archivos",
                r"D:\Bazan\octubre",
            ]
        )
        == 0
    )
    fin = json.loads(capsys.readouterr().out)
    assert fin["pedido"]["estado"] == "OK"
    assert fin["pedido"]["resultado"] == "Archivos bajados"
    assert fin["pedido"]["archivos"] == r"D:\Bazan\octubre"

    assert cli.main(["listar", "--estado", "PENDIENTE"]) == 0
    assert json.loads(capsys.readouterr().out) == []
