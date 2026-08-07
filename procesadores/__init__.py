"""Parsers especializados (liquidaciones de tarjeta, etc.)."""

from .tarjeta_parser import (
    PLANTILLAS_TARJETAS,
    detectar_entidad_por_texto,
    extraer_con_plantilla,
    extraer_texto_liquidacion_pdf,
)

__all__ = [
    "PLANTILLAS_TARJETAS",
    "detectar_entidad_por_texto",
    "extraer_con_plantilla",
    "extraer_texto_liquidacion_pdf",
]
