"""
Catálogo e identificación de especies de inversión.

Paso 1 del analizador: qué es (FCI / bono / acción / USD-MEP / otro).
El FIFO y las reglas de evento se aplican recién después.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

TIPO_FCI = "fci"
TIPO_BONO = "bono"
TIPO_ACCION = "accion"
TIPO_USD_MEP = "usd_mep"
TIPO_OTRO = "otro"

TIPOS_INVERSION = (TIPO_FCI, TIPO_BONO, TIPO_ACCION, TIPO_USD_MEP, TIPO_OTRO)

TIPO_A_GRUPO = {
    TIPO_FCI: "FCI",
    TIPO_BONO: "Bonos / Títulos públicos",
    TIPO_ACCION: "Acciones / Cedears",
    TIPO_USD_MEP: "Dólar / MEP",
    TIPO_OTRO: "Otros",
}

GRUPO_A_TIPO = {v: k for k, v in TIPO_A_GRUPO.items()}

GRUPOS_ESPECIE = tuple(TIPO_A_GRUPO[t] for t in TIPOS_INVERSION)

TIPO_LABEL = {
    TIPO_FCI: "FCI",
    TIPO_BONO: "Bono / Título público",
    TIPO_ACCION: "Acción / Cedear",
    TIPO_USD_MEP: "Dólar / MEP (Caja USD)",
    TIPO_OTRO: "Otro / revisar",
}

CONF_ALTA = "alta"
CONF_REVISAR = "revisar"

EVENTO_ENTRADA = "entrada"
EVENTO_SALIDA = "salida"
EVENTO_INGRESO = "ingreso"
EVENTO_AMORT = "amortizacion"
EVENTO_OMITIR_USD = "omitir_usd"
EVENTO_REVISAR = "revisar"
EVENTO_OMITIR = "omitir"

TITULOS_PUBLICOS = {
    "AL30", "AL30D", "GD30", "GD30D", "AL29", "GD29", "AL35", "GD35",
    "AL41", "GD41", "AE38", "GE38", "TX26", "TX28", "T2X4", "T4X4",
    "BONAR", "BOPREAL", "LECAP", "LETRA", "S31L", "S30O",
}

# Alias de trabajo (Seco / Galicia FIMA). Orden: más específico primero.
_FIMA_CANON = (
    (("renta fija", "dolar"), "FIMA RENTA FIJA DOLARES CLASE A"),
    (("renta fija", "dolares"), "FIMA RENTA FIJA DOLARES CLASE A"),
    (("ahorro", "peso"), "FIMA AHORRO PESOS CLASE A"),
    (("premium",), "FIMA PREMIUM CLASE A"),
)

_USD_TOKENS = {"USD", "US", "USS", "DOLAR", "DOLARES", "U$S"}


def _norm(s: object) -> str:
    t = str(s or "").replace("\n", " ").strip().lower()
    t = (
        t.replace("á", "a").replace("é", "e").replace("í", "i")
        .replace("ó", "o").replace("ú", "u").replace("ñ", "n")
    )
    return re.sub(r"\s+", " ", t)


def _token_especie(especie: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", str(especie or "")).upper()


def _limpiar_nombre_fondo(nombre: str) -> str:
    t = re.sub(r"\s+", " ", str(nombre or "")).strip()
    t = re.sub(r"(?i)^fondo\s*[-–:]\s*", "", t).strip()
    return t or "FCI"


def canon_fima(texto: str) -> str | None:
    """Devuelve el nombre canónico FIMA si el texto matchea un alias conocido."""
    n = _norm(texto)
    if "fima" not in n and "fci" not in n and "fondo" not in n:
        return None
    for keys, canon in _FIMA_CANON:
        if all(k in n for k in keys):
            return canon
    if "fima" in n:
        return _limpiar_nombre_fondo(texto).upper()
    return None


@dataclass(frozen=True)
class IdentificacionEspecie:
    especie_original: str
    especie_canonica: str
    tipo_inversion: str
    grupo: str
    confianza: str
    motivo: str

    @property
    def tipo_label(self) -> str:
        return TIPO_LABEL.get(self.tipo_inversion, self.tipo_inversion)


def identificar_especie(
    especie: str,
    descripcion: str = "",
    tipo_op: str = "",
    moneda: str = "",
) -> IdentificacionEspecie:
    """Identifica tipo y nombre canónico. No aplica FIFO."""
    orig = str(especie or "").strip()
    blob = _norm(f"{especie} {descripcion} {tipo_op} {moneda}")
    token = _token_especie(orig)

    # FCI / FIMA primero (un "renta fija dólares" no es caja USD)
    fima = canon_fima(f"{orig} {descripcion}")
    if fima:
        return IdentificacionEspecie(
            orig, fima, TIPO_FCI, TIPO_A_GRUPO[TIPO_FCI], CONF_ALTA, "Alias FIMA / FCI",
        )
    if any(k in blob for k in ("fci", "fondo comun", "cuotaparte", "money market", "fima")):
        canon = _limpiar_nombre_fondo(orig)
        if "fima" in _norm(canon):
            canon = canon.upper()
        return IdentificacionEspecie(
            orig, canon or "FCI", TIPO_FCI, TIPO_A_GRUPO[TIPO_FCI], CONF_ALTA, "Fondo común",
        )

    # USD caja / billete (no título, no fondo)
    if token in _USD_TOKENS or orig.upper() in {"USD", "U$S", "U$D"}:
        return IdentificacionEspecie(
            orig, "USD", TIPO_USD_MEP, TIPO_A_GRUPO[TIPO_USD_MEP], CONF_ALTA, "Moneda / caja USD",
        )
    if any(k in blob for k in ("caja de ahorro en dolar", "caja de ahorro en dolares")):
        return IdentificacionEspecie(
            orig or "USD", "USD", TIPO_USD_MEP, TIPO_A_GRUPO[TIPO_USD_MEP],
            CONF_ALTA, "Extracto CA en dólares",
        )

    # Títulos públicos (AL30 sigue siendo bono aunque el texto diga MEP)
    for t in sorted(TITULOS_PUBLICOS, key=len, reverse=True):
        if token == t or token.startswith(t) or re.search(rf"\b{re.escape(t)}\b", orig.upper()):
            return IdentificacionEspecie(
                orig, t, TIPO_BONO, TIPO_A_GRUPO[TIPO_BONO], CONF_ALTA, f"Título público {t}",
            )
    if any(k in blob for k in ("titulo publico", "lecap", "letra del tesoro", "bopreal", "bonar")):
        return IdentificacionEspecie(
            orig, (orig or "BONO").upper(), TIPO_BONO, TIPO_A_GRUPO[TIPO_BONO],
            CONF_ALTA, "Texto de título público",
        )
    if re.search(r"\bbono\b", blob) and "carbono" not in blob:
        return IdentificacionEspecie(
            orig, (orig or "BONO").upper(), TIPO_BONO, TIPO_A_GRUPO[TIPO_BONO],
            CONF_ALTA, "Texto de bono",
        )

    # Operación MEP / dólar sin ticker de título
    if any(k in blob for k in ("dolar mep", "compra mep", "venta mep")):
        return IdentificacionEspecie(
            orig or "USD", "USD", TIPO_USD_MEP, TIPO_A_GRUPO[TIPO_USD_MEP],
            CONF_ALTA, "Operación MEP — usar Caja USD",
        )

    # Acciones / Cedears
    if "cedear" in blob:
        canon = re.sub(r"(?i)^cedear\s+", "CEDEAR ", orig).strip() or orig.upper()
        return IdentificacionEspecie(
            orig, canon, TIPO_ACCION, TIPO_A_GRUPO[TIPO_ACCION], CONF_ALTA, "Cedear",
        )
    if any(k in blob for k in ("accion", "equity", "adr")):
        return IdentificacionEspecie(
            orig, (orig or token or "ACCION").upper(), TIPO_ACCION, TIPO_A_GRUPO[TIPO_ACCION],
            CONF_ALTA, "Acción",
        )
    if re.fullmatch(r"[A-Z]{3,5}", token) and token not in TITULOS_PUBLICOS:
        return IdentificacionEspecie(
            orig, token, TIPO_ACCION, TIPO_A_GRUPO[TIPO_ACCION], CONF_ALTA, "Ticker corto",
        )

    if "obligacion negociable" in blob or re.search(r"\bon\b", blob):
        return IdentificacionEspecie(
            orig, orig or "ON", TIPO_OTRO, TIPO_A_GRUPO[TIPO_OTRO],
            CONF_REVISAR, "ON — confirmar tipo",
        )

    return IdentificacionEspecie(
        orig, orig or "Sin especie", TIPO_OTRO, TIPO_A_GRUPO[TIPO_OTRO],
        CONF_REVISAR, "No identificado",
    )


def tipo_desde_grupo(grupo: str) -> str | None:
    g = str(grupo or "").strip()
    if g in GRUPO_A_TIPO:
        return GRUPO_A_TIPO[g]
    n = _norm(g)
    for tipo, label in TIPO_LABEL.items():
        if n == _norm(label) or n == tipo:
            return tipo
    return None


def evento_para(tipo_inversion: str, tipo_operacion: str) -> str:
    """Qué hace este movimiento según el tipo de inversión."""
    tipo = (tipo_inversion or TIPO_OTRO).strip().lower()
    op = _norm(tipo_operacion)

    if tipo == TIPO_USD_MEP:
        return EVENTO_OMITIR_USD
    if tipo == TIPO_OTRO:
        return EVENTO_REVISAR

    if tipo == TIPO_ACCION and any(k in op for k in ("dividendo", "renta")):
        return EVENTO_INGRESO
    if tipo == TIPO_BONO and any(k in op for k in ("renta", "cupon")):
        return EVENTO_INGRESO
    if tipo == TIPO_BONO and "amortiz" in op:
        return EVENTO_AMORT

    entradas = {
        TIPO_FCI: ("suscrip", "aporte", "compra"),
        TIPO_BONO: ("compra",),
        TIPO_ACCION: ("compra",),
    }
    salidas = {
        TIPO_FCI: ("rescate", "venta"),
        TIPO_BONO: ("venta",),
        TIPO_ACCION: ("venta",),
    }
    for key in entradas.get(tipo, ()):
        if key in op:
            return EVENTO_ENTRADA
    for key in salidas.get(tipo, ()):
        if key in op:
            return EVENTO_SALIDA
    return EVENTO_OMITIR


def orden_evento(evento: str) -> int:
    """Mismo día: entradas antes que salidas."""
    return {
        EVENTO_ENTRADA: 0,
        EVENTO_INGRESO: 1,
        EVENTO_AMORT: 2,
        EVENTO_SALIDA: 3,
        EVENTO_OMITIR: 4,
        EVENTO_OMITIR_USD: 5,
        EVENTO_REVISAR: 6,
    }.get(evento, 9)
