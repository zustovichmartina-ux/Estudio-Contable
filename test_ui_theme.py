"""Skin CSS: paleta/logo intactos y módulos de negocio sin tocar."""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent


class TestSkinSoloCosmetico(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.css = (ROOT / "assets" / "estudio.css").read_text(encoding="utf-8")
        cls.app = (ROOT / "app.py").read_text(encoding="utf-8")
        cls.config = (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")

    def test_paleta_y_logo_intactos(self) -> None:
        for color in ("#2563EB", "#0B1C33", "#F4F6FA", "#0F172A", "#3B82F6"):
            self.assertIn(color, self.css)
        self.assertIn('primaryColor = "#2563EB"', self.config)
        self.assertIn("estudio-zona-guemes-wordmark-oscuro.png", self.app)
        self.assertTrue((ROOT / "assets" / "estudio-zona-guemes-wordmark-oscuro.png").is_file())

    def test_app_solo_inyecta_css(self) -> None:
        self.assertIn('BASE_DIR / "assets" / "estudio.css"', self.app)
        self.assertNotIn("_formulario_login_oficina", self.app)
        self.assertNotIn("col_login", self.app)

    def test_modulos_intactos(self) -> None:
        for nombre in (
            "Devengamiento de Impuestos",
            "Conciliación Bancaria",
            "Préstamos Financieros",
            "Inversiones",
            "Herramientas",
            "Tango",
            "ARCA",
        ):
            self.assertIn(nombre, self.app)
        self.assertIn("ventana_principal_v4", self.app)
        self.assertIn("_pantalla_login_oficina", self.app)
        self.assertIn("render_arca_module", self.app)

    def test_radios_suaves_sin_teal(self) -> None:
        self.assertIn("--ec-radius: 16px", self.css)
        self.assertIn("--ec-radius-sm: 12px", self.css)
        bajo = self.css.lower()
        for token in ("#0d9488", "#14b8a6", "#d4af37", "#c9a227"):
            self.assertNotIn(token, bajo)


if __name__ == "__main__":
    unittest.main()
