"""Aviso de versión nueva y recorte del changelog."""
from __future__ import annotations

import unittest
from pathlib import Path

import ui_version_web as uv


class TestCambiosWeb(unittest.TestCase):
    def test_archivo_y_version(self) -> None:
        data = uv.cargar_cambios()
        self.assertTrue(data["version"])
        self.assertTrue(data["cambios"])
        self.assertEqual(uv.version_local(data), data["version"])
        self.assertTrue((Path(__file__).resolve().parent / "data" / "cambios_web.json").is_file())

    def test_cambios_desde_corta_en_el_id_viejo(self) -> None:
        data = {
            "version": "c",
            "cambios": [
                {"id": "c", "titulo": "Nuevo", "items": ["x"]},
                {"id": "b", "titulo": "Medio", "items": ["y"]},
                {"id": "a", "titulo": "Viejo", "items": ["z"]},
            ],
        }
        nuevos = uv.cambios_desde(data, "b")
        self.assertEqual([c["id"] for c in nuevos], ["c"])

    def test_sin_visto_previo_solo_el_ultimo(self) -> None:
        data = {
            "version": "c",
            "cambios": [
                {"id": "c", "titulo": "Nuevo", "items": ["x"]},
                {"id": "b", "titulo": "Medio", "items": ["y"]},
            ],
        }
        self.assertEqual([c["id"] for c in uv.cambios_desde(data, "")], ["c"])


if __name__ == "__main__":
    unittest.main()
