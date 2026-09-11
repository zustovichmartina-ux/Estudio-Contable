"""Instructivo Conceptos Bancos 2.xlsx: lista cerrada de cuentas + conceptos por banco.

La cuenta sale de la hoja CUENTAS. Si el extracto no alcanza, se usa la
cuenta «a identificar» que corresponda y se marca para revisar.
"""
from __future__ import annotations

import json
import re
import unicodedata
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd

try:
    from rapidfuzz import fuzz as _fuzz
except Exception:
    _fuzz = None

ROOT = Path(__file__).resolve().parent
CACHE_PATH = ROOT / "data" / "conceptos_bancos_cache.json"

RUTAS_INSTRUCTIVO = (
    Path(r"\\TANGOSRV\Compartido\CLIENTES\zzInstrucciones\Conceptos Bancos 2.xlsx"),
    Path(r"C:\Users\recep\Desktop\Nueva carpeta\zzInstrucciones\Conceptos Bancos 2.xlsx"),
)

_HOJAS_META = {
    "cuentas",
    "informacion general",
    "asientos contables",
    "revision",
    "revisión",
}

_ALIAS_BANCO: dict[str, tuple[str, ...]] = {
    "galicia": ("galicia",),
    "bbva": ("bbva", "frances", "francés", "banco frances"),
    "provincia": ("provincia", "bapro"),
    "nacion": ("nacion", "nación", "banco nacion"),
    "credicoop": ("credicoop",),
    "macro": ("macro",),
    "santander": ("santander", "rio", "río"),
    "supervielle": ("supervielle", "superville"),
    "icbc": ("icbc",),
    "mercado pago": ("mercado pago", "mercadopago", "coelsa", "mp "),
}

_CUENTA_DEFAULT = "Movimientos a identificar"
_SCORE_MIN = 68.0


def _norm(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _tokens(texto: str) -> set[str]:
    return {t for t in _norm(texto).split() if len(t) >= 3}


def _es_hoja_banco(nombre: str) -> bool:
    n = _norm(nombre)
    if n in _HOJAS_META:
        return False
    return n.startswith("banco") or n.startswith("mercado pago")


def _clave_banco(nombre_hoja: str) -> str:
    n = _norm(nombre_hoja)
    n = re.sub(r"^banco\s+", "", n)
    if "mercado pago" in n:
        return "mercado pago"
    if "superville" in n:
        return "supervielle"
    return n.strip()


def _banco_desde_texto(banco: str) -> str:
    n = _norm(banco)
    for clave, alias in _ALIAS_BANCO.items():
        if any(a in n for a in alias) or clave in n:
            return clave
    n = re.sub(r"^banco\s+", "", n)
    return n.strip()


def _lado_movimiento(concepto: str, debito: float, credito: float) -> bool:
    """False si el concepto es solo débito/crédito y el movimiento es el otro lado."""
    toks = _tokens(concepto)
    es_deb = debito > 0.005 and credito <= 0.005
    es_cre = credito > 0.005 and debito <= 0.005
    if "debito" in toks and "credito" not in toks and es_cre:
        return False
    if "credito" in toks and "debito" not in toks and es_deb:
        return False
    return True


def _mtime(path: Path) -> float:
    try:
        return path.stat().st_mtime
    except OSError:
        return 0.0


def _origen_fuente(fuente) -> str:
    if isinstance(fuente, (bytes, bytearray)):
        return "upload"
    if isinstance(fuente, (str, Path)):
        return str(fuente)
    return str(getattr(fuente, "name", "") or "upload")


def _score_concepto(concepto: str, descripcion: str) -> float:
    c = _norm(concepto)
    d = _norm(descripcion)
    if not c or not d:
        return 0.0
    if c in d or d in c:
        return 100.0
    tok_c = _tokens(concepto)
    tok_d = _tokens(descripcion)
    if not tok_c:
        return 0.0
    overlap = len(tok_c & tok_d) / len(tok_c)
    score = overlap * 90.0
    if _fuzz is not None:
        score = max(score, float(_fuzz.partial_ratio(c, d)))
        score = max(score, float(_fuzz.token_set_ratio(c, d)))
    if "25413" in c and "25413" in d:
        score = max(score, 96.0)
    if "fima" in c and "fima" in d:
        score = max(score, 94.0)
    if "coelsa" in c and "coelsa" in d:
        score = max(score, 94.0)
    if "prisma" in c and "prisma" in d:
        score = max(score, 92.0)
    return score


def _leer_workbook(fuente) -> dict[str, pd.DataFrame]:
    if isinstance(fuente, (bytes, bytearray)):
        fuente = BytesIO(fuente)
    xl = pd.ExcelFile(fuente)
    hojas: dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        hojas[str(name)] = pd.read_excel(xl, sheet_name=name, header=None, dtype=object)
    return hojas


def _cuenta_valida(nombre: str, cuentas: list[str]) -> str:
    n = str(nombre or "").strip()
    if not n:
        return _CUENTA_DEFAULT
    cerradas = { _norm(c): c for c in cuentas }
    if _norm(n) in cerradas:
        return cerradas[_norm(n)]
    # Plantilla "Transferencias entre cuentas propias - [Banco X]"
    if "transferencias entre cuentas propias" in _norm(n):
        for c in cuentas:
            if "transferencias entre cuentas propias" in _norm(c):
                return c
    for c in cuentas:
        if "identificar" in _norm(c) and "movimiento" in _norm(c):
            return c
    return _CUENTA_DEFAULT


def parsear_instructivo(fuente) -> dict[str, Any]:
    """Parsea el xlsx. fuente: path, bytes o file-like."""
    hojas = _leer_workbook(fuente)
    cuentas: list[str] = []
    bancos: dict[str, list[dict[str, str]]] = {}
    clientes: dict[str, str] = {}
    origen = ""

    for nombre, df in hojas.items():
        if df is None or df.empty:
            continue
        nkey = _norm(nombre)
        if nkey == "cuentas":
            col0 = df.iloc[:, 0]
            for val in col0.tolist()[1:]:
                txt = str(val or "").strip()
                if txt and txt.lower() not in {"nan", "cuenta contable", "none"}:
                    cuentas.append(txt)
            continue
        if _es_hoja_banco(nombre):
            header = None
            for i, row in df.iterrows():
                vals = [str(v or "").strip().lower() for v in row.tolist()[:3]]
                joined = " ".join(vals)
                if "concepto" in joined and "cuenta" in joined:
                    header = int(i)
                    break
            if header is None:
                continue
            reglas: list[dict[str, str]] = []
            for _, row in df.iloc[header + 1 :].iterrows():
                concepto = str(row.iloc[0] or "").strip() if len(row) > 0 else ""
                cuenta = str(row.iloc[1] or "").strip() if len(row) > 1 else ""
                trat = str(row.iloc[2] or "").strip() if len(row) > 2 else ""
                if not concepto or concepto.lower() in {"nan", "none", "notas del banco", "regla de la hoja"}:
                    continue
                if concepto.lower().startswith("notas"):
                    break
                reglas.append(
                    {
                        "concepto": concepto,
                        "cuenta": cuenta,
                        "tratamiento": trat if trat.lower() not in {"nan", "none"} else "",
                    }
                )
            bancos[_clave_banco(nombre)] = reglas
            continue
        if nkey not in _HOJAS_META:
            texto = " ".join(
                str(v) for v in df.iloc[:8, 0].tolist() if str(v or "").strip() not in {"", "nan"}
            )
            clientes[str(nombre).strip()] = texto[:500]

    if not cuentas:
        raise ValueError("El instructivo no tiene hoja CUENTAS con cuentas cargadas.")

    return {
        "cuentas": cuentas,
        "bancos": bancos,
        "clientes": clientes,
        "origen": origen,
    }


def guardar_cache(data: dict[str, Any], origen: str = "") -> Path:
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(data)
    payload["origen"] = origen
    CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return CACHE_PATH


def cargar_cache() -> dict[str, Any] | None:
    if not CACHE_PATH.exists():
        return None
    try:
        data = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict) or not data.get("bancos"):
        return None
    return data


def cargar_instructivo(fuente=None, *, forzar: bool = False) -> dict[str, Any]:
    """Carga el xlsx (path/bytes) o el cache. Prefiere Conceptos Bancos 2.

    Si hay cache y el xlsx de red/disco no es más nuevo, no se reparsea.
    """
    if fuente is not None:
        data = parsear_instructivo(fuente)
        origen = _origen_fuente(fuente)
        guardar_cache(data, origen=origen)
        data["origen"] = origen
        return data

    cached = None if forzar else cargar_cache()
    cache_mtime = _mtime(CACHE_PATH) if cached else 0.0

    for ruta in RUTAS_INSTRUCTIVO:
        try:
            if not ruta.is_file():
                continue
            if cached and not forzar and _mtime(ruta) <= cache_mtime + 0.5:
                if not cached.get("origen"):
                    cached["origen"] = str(ruta)
                return cached
            data = parsear_instructivo(ruta)
            guardar_cache(data, origen=str(ruta))
            data["origen"] = str(ruta)
            return data
        except OSError:
            continue

    if cached:
        return cached
    raise FileNotFoundError(
        "No se encontró Conceptos Bancos 2.xlsx. "
        "Subilo en Conciliación o copiálo a zzInstrucciones."
    )


def clasificar_movimiento_instructivo(
    descripcion: str,
    banco: str,
    *,
    debito: float = 0.0,
    credito: float = 0.0,
    instructivo: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Devuelve cuenta del instructivo. Si no hay match claro → a identificar."""
    data = instructivo or cargar_cache() or {}
    cuentas = list(data.get("cuentas") or [])
    clave = _banco_desde_texto(banco)
    reglas = list((data.get("bancos") or {}).get(clave) or [])
    if not reglas:
        # banco no mapeado: no inventar, a identificar
        cuenta = _cuenta_valida(_CUENTA_DEFAULT, cuentas) if cuentas else _CUENTA_DEFAULT
        return {
            "cuenta": cuenta,
            "concepto": "",
            "tratamiento": "",
            "banco_hoja": clave,
            "score": 0.0,
            "fuente": "Conceptos Bancos",
            "citar": f"Sin hoja de banco para «{banco or '?'}» → {cuenta}",
            "identificado": False,
        }

    mejor: dict[str, Any] | None = None
    mejor_score = 0.0
    for regla in reglas:
        if not _lado_movimiento(str(regla.get("concepto") or ""), float(debito or 0), float(credito or 0)):
            continue
        score = _score_concepto(str(regla.get("concepto") or ""), descripcion)
        if score > mejor_score:
            mejor_score = score
            mejor = regla

    if not mejor or mejor_score < _SCORE_MIN:
        cuenta = _cuenta_valida(_CUENTA_DEFAULT, cuentas)
        return {
            "cuenta": cuenta,
            "concepto": (mejor or {}).get("concepto") or "",
            "tratamiento": "El extracto no alcanza para definir la cuenta.",
            "banco_hoja": clave,
            "score": round(mejor_score, 1),
            "fuente": "Conceptos Bancos",
            "citar": f"Hoja {clave}: sin match ≥{_SCORE_MIN:.0f} → {cuenta}",
            "identificado": False,
        }

    cuenta = _cuenta_valida(str(mejor.get("cuenta") or ""), cuentas)
    identificado = "identificar" not in _norm(cuenta)
    return {
        "cuenta": cuenta,
        "concepto": str(mejor.get("concepto") or ""),
        "tratamiento": str(mejor.get("tratamiento") or ""),
        "banco_hoja": clave,
        "score": round(mejor_score, 1),
        "fuente": "Conceptos Bancos",
        "citar": (
            f"Hoja {clave} · «{mejor.get('concepto')}» → {cuenta} "
            f"(score {mejor_score:.0f})"
        ),
        "identificado": identificado,
    }


def tipo_desde_cuenta(cuenta: str, debito: float, credito: float) -> str:
    n = _norm(cuenta)
    if "identificar" in n:
        return "DEBITO_REVISAR"
    if "proveedor" in n:
        return "DEBITO_PROVEEDOR"
    if "deudor" in n or "venta" in n:
        return "INGRESO"
    if any(k in n for k in ("iva", "iibb", "debito", "credito", "25413", "percepcion", "retencion", "impuesto")):
        return "DEBITO_IMPUESTO"
    if "transferencias entre cuentas propias" in n:
        return "INGRESO_O_DEBITO_PROPIO"
    if credito > debito:
        return "INGRESO"
    return "DEBITO_FIJO"
