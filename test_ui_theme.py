"""Tema visual: paleta del estudio, logo, sin marca Samsara."""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent


class TestUiThemeEstudio(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.css = (ROOT / "assets" / "estudio.css").read_text(encoding="utf-8")
        cls.app = (ROOT / "app.py").read_text(encoding="utf-8")
        cls.theme = (ROOT / "ui_theme.py").read_text(encoding="utf-8")
        cls.config = (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")

    def test_css_compartido_existe(self) -> None:
        self.assertTrue((ROOT / "assets" / "estudio.css").is_file())
        self.assertIn("--ec-lagoon: #2563EB", self.css)
        self.assertIn("--ec-radius: 22px", self.css)
        self.assertIn("--ec-radius-lg: 28px", self.css)
        self.assertIn("cubic-bezier", self.css)

    def test_paleta_estudio_intacta(self) -> None:
        for color in ("#0B1C33", "#2563EB", "#3B82F6", "#F4F6FA", "#0F172A", "#64748B"):
            self.assertIn(color, self.css)
        self.assertIn('primaryColor = "#2563EB"', self.config)
        self.assertIn('backgroundColor = "#F4F6FA"', self.config)
        self.assertIn('textColor = "#0F172A"', self.config)

    def test_logo_estudio_sin_cambiar(self) -> None:
        logo = ROOT / "assets" / "estudio-zona-guemes-wordmark-oscuro.png"
        self.assertTrue(logo.is_file())
        self.assertIn("estudio-zona-guemes-wordmark-oscuro.png", self.app)
        self.assertIn("estudio-zona-guemes-wordmark-oscuro.png", self.theme)

    def test_sin_teal_ni_gold_samsara(self) -> None:
        bajo = self.css.lower()
        prohibidos = (
            "#0d9488",
            "#14b8a6",
            "#2dd4bf",
            "#0f766e",
            "#134e4a",
            "#d4af37",
            "#c9a227",
            "#d4a017",
            "#c9a84c",
        )
        for token in prohibidos:
            self.assertNotIn(token, bajo)

    def test_app_inyecta_tema_compartido(self) -> None:
        self.assertIn("from ui_theme import inyectar_tema", self.app)
        self.assertIn("inyectar_tema()", self.app)
        self.assertIn("def inyectar_tema", self.theme)


if __name__ == "__main__":
    unittest.main()
