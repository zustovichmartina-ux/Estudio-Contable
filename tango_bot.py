# -*- coding: utf-8 -*-
"""Bot Tango: índice de ayudas Axoft + reglas del estudio. Nunca guarda claves."""
from __future__ import annotations

import base64
import csv
import json
import os
import re
import unicodedata
import urllib.error
import urllib.request
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path
from typing import Any

try:
    from PIL import Image
except ImportError:
    Image = None  # type: ignore[misc, assignment]

ROOT = Path(__file__).resolve().parent
CONOCIMIENTO_DIR = ROOT / "tango_conocimiento"
INDEX_PATH = CONOCIMIENTO_DIR / "index.jsonl"

_SKIP_NAME = re.compile(
    r"thumbs\.db|tango_db_avance|tango\.ini|password|clave",
    re.I,
)
_TOKEN = re.compile(r"[a-záéíóúñü0-9]{3,}", re.I)


class _HtmlText(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._skip = 0
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style", "noscript"}:
            self._skip += 1
        if tag in {"p", "div", "li", "tr", "br", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style", "noscript"} and self._skip:
            self._skip -= 1

    def handle_data(self, data: str) -> None:
        if not self._skip:
            self.parts.append(data)


def _html_a_texto(raw: str) -> str:
    parser = _HtmlText()
    try:
        parser.feed(raw)
        parser.close()
    except Exception:
        return re.sub(r"<[^>]+>", " ", raw)
    text = re.sub(r"[ \t]+", " ", "".join(parser.parts))
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _fuentes_dir() -> list[Path]:
    home = Path.home() / "Desktop"
    return [
        CONOCIMIENTO_DIR,
        home / "Tango",
        Path(r"\\TANGOSRV\Compartido\CLIENTES\1 - Normativa - Vencimientos\Tango"),
        home,
    ]


def _omitir(path: Path) -> bool:
    name = path.name
    if _SKIP_NAME.search(name):
        return True
    if name.startswith("~$"):
        return True
    return False


def _trozar(texto: str, size: int = 1400, overlap: int = 180) -> list[str]:
    texto = (texto or "").strip()
    if not texto:
        return []
    if len(texto) <= size:
        return [texto]
    out: list[str] = []
    i = 0
    while i < len(texto):
        out.append(texto[i : i + size].strip())
        i += size - overlap
    return [c for c in out if len(c) > 40]


def _leer_archivo(path: Path) -> list[dict[str, str]]:
    suf = path.suffix.lower()
    try:
        if suf in {".html", ".htm"}:
            raw = path.read_text(encoding="utf-8", errors="replace")
            title_m = re.search(r"<title>(.*?)</title>", raw, re.I | re.S)
            title = re.sub(r"\s+", " ", title_m.group(1)).strip() if title_m else path.stem
            chunks = _trozar(_html_a_texto(raw))
            return [{"source": str(path), "title": title, "text": c} for c in chunks]
        if suf in {".md", ".txt"}:
            raw = path.read_text(encoding="utf-8", errors="replace")
            chunks = _trozar(raw)
            return [{"source": str(path), "title": path.stem, "text": c} for c in chunks]
        if suf == ".csv" and "tango_sueldos" in path.name.lower():
            return _csv_sueldos(path)
    except OSError:
        return []
    return []


def _csv_sueldos(path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with path.open(encoding="utf-8-sig", errors="replace", newline="") as fh:
        reader = csv.DictReader(fh)
        for i, row in enumerate(reader):
            if i > 400:
                break
            partes = [f"{k}: {v}" for k, v in row.items() if str(v or "").strip()]
            if not partes:
                continue
            codigo = str(row.get("Codigo") or row.get("TIPO_CONCEPTO") or "").strip() or str(i + 1)
            desc = str(row.get("Descripcion") or row.get("DESC_TIPO_CONCEPTO") or path.stem)
            rows.append({
                "source": str(path),
                "title": f"Sueldos {codigo} {desc}".strip(),
                "text": " | ".join(partes)[:2500],
            })
    return rows


def reconstruir_indice(*, guardar: bool = True) -> list[dict[str, str]]:
    """Recorre Desktop/Tango, normativas UNC y tango_conocimiento/."""
    CONOCIMIENTO_DIR.mkdir(parents=True, exist_ok=True)
    docs: list[dict[str, str]] = []
    vistos: set[str] = set()

    def _clave(path: Path, raiz: Path) -> str:
        try:
            return path.relative_to(raiz).as_posix().lower()
        except ValueError:
            return path.name.lower()

    for raiz in _fuentes_dir():
        if not raiz.exists():
            continue
        if raiz.name.lower() == "desktop":
            for path in raiz.glob("Tango_Sueldos_*.csv"):
                if _omitir(path) or path.name.lower() in vistos:
                    continue
                vistos.add(path.name.lower())
                docs.extend(_leer_archivo(path))
            for path in raiz.glob("Tango_Sueldos_arbol_*.txt"):
                if _omitir(path) or path.name.lower() in vistos:
                    continue
                vistos.add(path.name.lower())
                docs.extend(_leer_archivo(path))
            continue
        for path in raiz.rglob("*"):
            if not path.is_file() or _omitir(path):
                continue
            if path.suffix.lower() not in {".html", ".htm", ".md", ".txt"}:
                continue
            key = _clave(path, raiz)
            if key in vistos:
                continue
            vistos.add(key)
            docs.extend(_leer_archivo(path))
    if guardar:
        INDEX_PATH.write_text(
            "\n".join(json.dumps(d, ensure_ascii=False) for d in docs) + ("\n" if docs else ""),
            encoding="utf-8",
        )
    return docs


def cargar_indice() -> list[dict[str, str]]:
    if not INDEX_PATH.exists() or INDEX_PATH.stat().st_size < 50:
        return reconstruir_indice(guardar=True)
    out: list[dict[str, str]] = []
    for line in INDEX_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(item, dict) and item.get("text"):
            out.append({
                "source": str(item.get("source") or ""),
                "title": str(item.get("title") or ""),
                "text": str(item.get("text") or ""),
            })
    return out


def _sin_acento(texto: str) -> str:
    nfd = unicodedata.normalize("NFD", texto or "")
    return "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")


def _tokens(texto: str) -> set[str]:
    plano = _sin_acento(texto or "").lower()
    plano = plano.replace("formulaimporte", "formula importe")
    plano = plano.replace("formulacantidad", "formula cantidad")
    plano = plano.replace("formulavalor", "formula valor")
    found = {t for t in _TOKEN.findall(plano)}
    extra: set[str] = set()
    for t in list(found):
        if t.endswith("s") and len(t) > 4:
            extra.add(t[:-1])
    return found | extra


def _es_consulta_formula(pregunta: str) -> bool:
    p = _sin_acento(pregunta or "").lower()
    return bool(
        re.search(
            r"formula|concepto\s*\d+|sueldo basico|sueldo proporcional|liquidacion|"
            r"en limpio|copi[aeo]|pegar|para pegar",
            p,
        )
    )


def _pide_pegar(pregunta: str) -> bool:
    p = _sin_acento(pregunta or "").lower()
    return bool(
        re.search(r"en limpio|copi[aeo]|pegar|para pegar|lista para|solo la formula|solamente la formula", p)
    )


def _texto_busqueda(pregunta: str, historial: list[dict[str, str]] | None) -> str:
    partes = [pregunta or ""]
    for m in (historial or [])[-6:]:
        partes.append(str(m.get("content") or ""))
    return "\n".join(partes)


def _nro_concepto(pregunta: str) -> str:
    m = re.search(r"concepto\s*(\d+)", _sin_acento(pregunta or ""), re.I)
    return m.group(1) if m else ""


def recuperar(pregunta: str, docs: list[dict[str, str]], k: int = 8) -> list[dict[str, str]]:
    q = _tokens(pregunta)
    if not q or not docs:
        return docs[:k]
    nro = _nro_concepto(pregunta) if _es_consulta_formula(pregunta) else ""
    scored: list[tuple[float, dict[str, str]]] = []
    for doc in docs:
        blob = f"{doc.get('title', '')} {doc.get('text', '')}"
        t = _tokens(blob)
        if not t:
            continue
        hit = len(q & t)
        if hit == 0:
            continue
        score = float(hit)
        title = (doc.get("title") or "").lower()
        source = (doc.get("source") or "").lower()
        if any(w in title for w in q):
            score += 2.0
        if "formulas_completo.csv" in source or "sueldos_formulas" in source:
            score += 12.0
        if nro and re.search(rf'codigo["\']?\s*:\s*["\']?{nro}\b', _sin_acento(blob).lower()):
            score += 25.0
        if nro and re.search(rf"^sueldos {nro} ", title):
            score += 4.0
        if "sueldo basico" in _sin_acento(title) and "basico" in q:
            score += 18.0
        if "sicoss" in source or "siap" in source:
            if "sicoss" not in q and "siap" not in q:
                score -= 8.0
        scored.append((score, doc))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [d for _, d in scored[:k]]


def _campo_csv(texto: str, nombre: str) -> str:
    m = re.search(rf"{re.escape(nombre)}:\s*(.*?)(?:\s+\|\s+[A-Za-zÁÉÍÓÚÑ][\w]*:|$)", texto, re.S)
    if not m:
        return ""
    return m.group(1).strip().strip("|").strip()


def _nro_de_hit(hit: dict[str, str]) -> str:
    m = re.search(
        r'codigo["\']?\s*:\s*["\']?(\d+)',
        _sin_acento(hit.get("text") or ""),
        re.I,
    )
    return m.group(1) if m else ""


def _nombre_concepto(hit: dict[str, str]) -> str:
    texto = hit.get("text") or ""
    desc = _campo_csv(texto, "Descripcion")
    if not desc:
        titulo = hit.get("title") or "Concepto"
        desc = re.sub(r"^Sueldos\s+\d+\s+", "", titulo, flags=re.I).strip() or titulo
    nro = _nro_de_hit(hit)
    if nro:
        return f"Concepto {nro} — {desc}"
    return desc


def _formula_para_pegar(hit: dict[str, str]) -> str:
    texto = hit.get("text") or ""
    importe = _campo_csv(texto, "FormulaImporte")
    cantidad = _campo_csv(texto, "FormulaCantidad")
    lineas = [f"**{_nombre_concepto(hit)}**", "", "Listo para pegar en Tango:"]
    if importe:
        lineas.extend(["", "**Importe**", "", "```", importe, "```"])
    if cantidad:
        lineas.extend(["", "**Cantidad**", "", "```", cantidad, "```"])
    if not importe and not cantidad:
        return _explicar_formula(hit)
    return "\n".join(lineas)


def _explicar_formula(hit: dict[str, str]) -> str:
    texto = hit.get("text") or ""
    importe = _campo_csv(texto, "FormulaImporte")
    cantidad = _campo_csv(texto, "FormulaCantidad")
    tipo = _campo_csv(texto, "TipoConcepto")
    lineas = [f"**{_nombre_concepto(hit)}**"]
    if tipo:
        lineas.append(f"Tipo: {tipo}")
    if importe:
        lineas.extend(["", "**Importe** (pegá en FormulaImporte):", "", "```", importe, "```"])
        if "USUELD" in importe:
            lineas.append("`USUELD` = sueldo básico del legajo")
        if "CANTIDAD" in importe:
            lineas.append("`CANTIDAD` = días/horas del concepto")
    if cantidad:
        lineas.extend(["", "**Cantidad** (pegá en FormulaCantidad):", "", "```", cantidad, "```"])
    if not importe and not cantidad:
        lineas.append(texto[:800])
    return "\n".join(lineas)


_GROK_MODELOS = ("grok-4.6", "grok-4", "grok-3")


def _resolver_llm(*, api_key: str = "", model: str = "") -> tuple[str, str, str]:
    """Devuelve (api_key, url, model). Prioriza Grok (xAI)."""
    xai = (api_key or os.environ.get("XAI_API_KEY") or "").strip()
    if xai:
        elegido = (model or os.environ.get("XAI_MODEL") or "grok-4.6").strip() or "grok-4.6"
        return xai, "https://api.x.ai/v1/chat/completions", elegido
    openai = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if openai:
        return openai, "https://api.openai.com/v1/chat/completions", (os.environ.get("OPENAI_MODEL") or "gpt-4o-mini").strip()
    groq = (os.environ.get("GROQ_API_KEY") or "").strip()
    if groq:
        return groq, "https://api.groq.com/openai/v1/chat/completions", (os.environ.get("GROQ_MODEL") or "llama-3.3-70b-versatile").strip()
    return "", "", ""


def _llamar_llm(
    system: str,
    messages: list[dict[str, Any]],
    *,
    api_key: str = "",
    model: str = "",
    imagenes: list[dict[str, str]] | None = None,
) -> str:
    key, url, modelo = _resolver_llm(api_key=api_key, model=model)
    if not key or not url:
        return ""
    if imagenes:
        if "x.ai" in url:
            modelo = (os.environ.get("XAI_VISION_MODEL") or modelo or "grok-4.6").strip()
        elif "openai.com" in url:
            modelo = (os.environ.get("OPENAI_VISION_MODEL") or "gpt-4o-mini").strip()
        elif "groq.com" in url:
            modelo = (
                os.environ.get("GROQ_VISION_MODEL")
                or "meta-llama/llama-4-scout-17b-16e-instruct"
            ).strip()
        last = messages[-1]
        texto = last.get("content") if isinstance(last.get("content"), str) else str(last.get("content") or "")
        partes: list[dict[str, Any]] = [{"type": "text", "text": texto}]
        for img in imagenes[:4]:
            partes.append({
                "type": "image_url",
                "image_url": {"url": img["data_url"]},
            })
        messages = [*messages[:-1], {"role": "user", "content": partes}]
    candidatos = [modelo]
    if "x.ai" in url:
        for alt in _GROK_MODELOS:
            if alt not in candidatos:
                candidatos.append(alt)
    ultimo = ""
    for modelo_try in candidatos:
        payload = {
            "model": modelo_try,
            "temperature": 0.2,
            "messages": [{"role": "system", "content": system}, *messages],
        }
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json",
                "User-Agent": "EstudioContable-TangoBot/1.0",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")[:400]
            ultimo = f"IA HTTP {exc.code}: {body}"
            if exc.code in {400, 404} and "model" in body.lower() and modelo_try != candidatos[-1]:
                continue
            raise RuntimeError(ultimo) from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"No se llegó a Grok (xAI): {exc.reason}") from exc
        choices = data.get("choices") or []
        if not choices:
            continue
        return str((choices[0].get("message") or {}).get("content") or "").strip()
    if ultimo:
        raise RuntimeError(ultimo)
    return ""


_SYSTEM = """Sos Grok, el agente de Tango Estudios (Axoft 26ar) del Estudio Contable.
Respondé y formulá. No vuelques CSV ni pegues ayudas crudas.

Cómo operás:
1. Usá el contexto solo como fuente. Contestá con tus palabras, corto y útil.
2. Si hay captura: qué pantalla/error se ve, y después la solución.
3. Si piden fórmula: bloques listos para pegar (Importe y Cantidad). Sintaxis Axoft:
   SI(), NOVCA, NOVCAG, USUELD, CANTIDAD, SUELDO, ABS, ACUCTA, MOVCTA.
4. Asiento por comprobante (Liquidador de IVA) ≠ determinación mensual DETIVA/DETIIBB/DETTISH.
5. No mezcles SIAp/SICOSS salvo que lo pidan.
6. FormulaImporte / FormulaCantidad del export es la fuente de verdad.
7. Concepto 1 = Sueldo básico `USUELD/30*CANTIDAD` (no el proporcional).
8. Si no está en el contexto ni en la imagen, decilo. No inventes menús.
9. Nunca pidas ni escribas claves SQL, TANGO.INI ni contraseñas.
10. “En limpio” / copiar / pegar: SOLO las fórmulas en bloques de código.
Respondé en español, como un compañero del estudio.
"""


def _es_doc_formula(doc: dict[str, str]) -> bool:
    source = (doc.get("source") or "").lower()
    return "formulas_completo.csv" in source or "sueldos_formulas" in source


def _docs_formula(docs: list[dict[str, str]]) -> list[dict[str, str]]:
    return [d for d in docs if _es_doc_formula(d)]


def _codigo_en_texto(texto: str, nro: str) -> bool:
    plano = _sin_acento(texto or "").lower()
    return bool(re.search(rf'codigo["\']?\s*:\s*["\']?{re.escape(nro)}\b', plano))


def _mejor_formula(pregunta: str, hits: list[dict[str, str]]) -> dict[str, str] | None:
    csv_hits = _docs_formula(hits)
    if not csv_hits:
        return None
    p = _sin_acento(pregunta or "").lower()
    nro = _nro_concepto(pregunta)

    def titulo(h: dict[str, str]) -> str:
        return _sin_acento(h.get("title") or "").lower()

    if "basico" in p:
        for h in csv_hits:
            t = titulo(h)
            if "sueldo basico" in t and "proporcional" not in t:
                return h
    if nro:
        for h in csv_hits:
            if _codigo_en_texto(h.get("text") or "", nro) and "sueldo basico" in titulo(h):
                return h
        for h in csv_hits:
            if _codigo_en_texto(h.get("text") or "", nro):
                return h
    return csv_hits[0]


def _fallback(pregunta: str, hits: list[dict[str, str]]) -> str:
    if _es_consulta_formula(pregunta) or _pide_pegar(pregunta):
        hit = _mejor_formula(pregunta, hits)
        if hit:
            return _formula_para_pegar(hit) if _pide_pegar(pregunta) else _explicar_formula(hit)
    for h in hits:
        if _es_doc_formula(h):
            return _explicar_formula(h)
    if not hits:
        return (
            "No lo tengo en las ayudas Tango del estudio. "
            "Probá con el nombre de la pantalla (Asientos, Modelos, Sueldos, IVA)."
        )
    bloques = []
    for h in hits[:2]:
        titulo = (h.get("title") or "Ayuda").strip()
        texto = (h.get("text") or "").strip()
        texto = re.sub(r"\s+\|\s+", "\n", texto)[:600]
        bloques.append(f"**{titulo}**\n\n{texto}")
    return "Según las ayudas del estudio:\n\n" + "\n\n".join(bloques)


def preparar_imagen(raw: bytes, nombre: str = "captura.png") -> dict[str, str]:
    """Comprime la captura para visión (JPEG). Devuelve data_url."""
    mime = "image/jpeg"
    data = raw
    try:
        if Image is not None:
            im = Image.open(BytesIO(raw))
            im = im.convert("RGB")
            im.thumbnail((1600, 1600))
            buf = BytesIO()
            im.save(buf, format="JPEG", quality=85)
            data = buf.getvalue()
        else:
            raise RuntimeError("sin Pillow")
    except Exception:
        data = raw
        lower = (nombre or "").lower()
        if lower.endswith(".png"):
            mime = "image/png"
        elif lower.endswith(".webp"):
            mime = "image/webp"
        else:
            mime = "image/jpeg"
    b64 = base64.b64encode(data).decode("ascii")
    return {"data_url": f"data:{mime};base64,{b64}", "nombre": Path(nombre).name}


def responder(
    pregunta: str,
    historial: list[dict[str, str]] | None = None,
    *,
    api_key: str = "",
    model: str = "",
    imagenes: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    docs = cargar_indice()
    consulta = pregunta or ""
    if imagenes and not consulta.strip():
        consulta = "Leé la captura de Tango: qué pantalla es, qué error o fórmula muestra, y cómo se resuelve."
    busqueda = _texto_busqueda(consulta, historial)
    hits = recuperar(busqueda if _pide_pegar(consulta) else consulta, docs, k=8)
    formula_hit = None
    if _es_consulta_formula(busqueda) or _pide_pegar(consulta):
        formula_hit = _mejor_formula(busqueda, _docs_formula(docs) or hits)
        if formula_hit:
            hits = [formula_hit] + [h for h in hits if h is not formula_hit][:7]
    key, _url, _modelo = _resolver_llm(api_key=api_key, model=model)
    if formula_hit and (_pide_pegar(consulta) or not key):
        return {
            "texto": _formula_para_pegar(formula_hit) if _pide_pegar(consulta) else _explicar_formula(formula_hit),
            "fuentes": [
                {"title": formula_hit.get("title") or "", "source": Path(str(formula_hit.get("source") or "")).name}
            ],
            "docs": len(docs),
            "uso_grok": False,
        }
    if not key:
        return {
            "texto": (
                "Este chat responde con **Grok**. Falta la clave xAI.\n\n"
                "En **Tango → Ajustes** pegá `XAI_API_KEY` (empieza con `xai-`) "
                "o en Streamlit **Manage app → Secrets**:\n\n"
                "`XAI_API_KEY = \"xai-...\"`\n\n"
                "La clave se crea en https://console.x.ai"
            ),
            "fuentes": [],
            "docs": len(docs),
            "uso_grok": False,
        }
    bloques_ctx: list[str] = []
    if formula_hit:
        bloques_ctx.append(
            "### Fórmula del export Tango Sueldos (fuente de verdad)\n"
            + _explicar_formula(formula_hit)
        )
    extra_hits = [h for h in hits if h is not formula_hit][:3]
    for h in extra_hits:
        texto_h = (h.get("text") or "").strip()
        if _es_doc_formula(h):
            texto_h = _explicar_formula(h)
        else:
            texto_h = re.sub(r"\s+\|\s+", "\n", texto_h)[:700]
        bloques_ctx.append(
            f"### {h.get('title')}\nFuente: {Path(str(h.get('source') or '')).name}\n{texto_h}"
        )
    contexto = "\n\n".join(bloques_ctx)
    msgs: list[dict[str, Any]] = []
    for m in (historial or [])[-8:]:
        role = m.get("role") or "user"
        if role not in {"user", "assistant"}:
            continue
        content = str(m.get("content") or "").strip()
        if content:
            msgs.append({"role": role, "content": content})
    extra_img = ""
    if imagenes:
        extra_img = (
            f"\n\nEl usuario adjuntó {len(imagenes)} captura(s) de Tango. "
            "Leelas y usalas para formular o corregir."
        )
    msgs.append({
        "role": "user",
        "content": f"Pregunta:\n{consulta}{extra_img}\n\nContexto:\n{contexto}",
    })
    ia = ""
    error = ""
    if imagenes and not key:
        return {
            "texto": (
                "Para leer capturas el agente necesita IA con visión. "
                "En **Gestionar la aplicación → Secrets** pegá `XAI_API_KEY` o `OPENAI_API_KEY`. "
                "Mientras, escribí qué dice la pantalla o pegá la fórmula en texto."
            ),
            "fuentes": [],
            "docs": len(docs),
            "uso_grok": False,
        }
    try:
        ia = _llamar_llm(
            _SYSTEM,
            msgs,
            api_key=api_key,
            model=model,
            imagenes=imagenes,
        )
    except Exception as exc:
        error = str(exc)
    if ia:
        texto = ia
    elif formula_hit:
        texto = _explicar_formula(formula_hit)
        if error:
            texto = f"{texto}\n\n_(Grok no respondió: {error[:180]})_"
    else:
        texto = (
            f"Grok no pudo responder: {error[:240]}"
            if error
            else _fallback(consulta, hits)
        )
    return {
        "texto": texto,
        "fuentes": [
            {"title": h.get("title") or "", "source": Path(str(h.get("source") or "")).name}
            for h in hits[:6]
        ],
        "docs": len(docs),
        "uso_grok": bool(ia),
    }
