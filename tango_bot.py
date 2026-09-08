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


def _es_seguimiento_visual(pregunta: str) -> bool:
    """«desde acá cómo sigo» y similares: manda la captura, no el tema anterior."""
    p = _sin_acento(pregunta or "").lower()
    return bool(
        re.search(
            r"\b(desde aca|de aca|aca como|como sigo|y ahora|esta pantalla|"
            r"esta ventana|esto que se ve|que se ve)\b",
            p,
        )
    ) or (len(_tokens(pregunta)) <= 4 and bool(re.search(r"\b(aca|aqui|esto|sigo)\b", p)))


def _es_consulta_proceso(pregunta: str) -> bool:
    p = _sin_acento(pregunta or "").lower()
    return bool(
        re.search(
            r"importar|paso a paso|como (hago|puedo|se |hago)|proceso|"
            r"portal iva|arca|mis comprobantes|libro (de )?iva|"
            r"menu|pantalla|asistente",
            p,
        )
    )


def _es_consulta_formula(pregunta: str) -> bool:
    p = _sin_acento(pregunta or "").lower()
    if _es_consulta_proceso(p) and not re.search(r"formula|concepto\s*\d+|sueldo basico", p):
        return False
    return bool(
        re.search(
            r"formula|concepto\s*\d+|sueldo basico|sueldo proporcional|"
            r"en limpio|solo la formula|solamente la formula",
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
        es_formula_csv = (
            "formulas_completo.csv" in source
            or "sueldos_formulas" in source
            or "huerfanas" in source
            or "huerfanas" in title
        )
        if es_formula_csv:
            if _es_consulta_formula(pregunta):
                score += 12.0
            else:
                score -= 25.0
        if _es_consulta_proceso(pregunta):
            if "portal iva" in source or "arca" in source or "importar" in source or "importacion" in _sin_acento(title):
                score += 20.0
            if "sueldos" in source and es_formula_csv:
                score -= 15.0
        if nro and re.search(rf'codigo["\']?\s*:\s*["\']?{nro}\b', _sin_acento(blob).lower()):
            score += 25.0
        if nro and re.search(rf"^sueldos {nro} ", title):
            score += 4.0
        if "sueldo basico" in _sin_acento(title) and "basico" in q:
            score += 18.0
        if "sicoss" in source or "siap" in source:
            if "sicoss" not in q and "siap" not in q:
                score -= 8.0
        if "base_marti" in source or "tango_conocimiento" in source:
            score += 8.0
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
_CLAUDE_MODELOS = ("claude-sonnet-5", "claude-sonnet-4-6", "claude-sonnet-4-5")
_DATA_URL_IMG = re.compile(
    r"^data:(image/(?:jpeg|png|gif|webp));base64,(.+)$",
    re.I | re.S,
)


def _proveedor_de_clave(key: str) -> str:
    k = (key or "").strip()
    if k.startswith("sk-ant-"):
        return "anthropic"
    if k.startswith("xai-"):
        return "xai"
    if k.startswith("gsk_"):
        return "groq"
    if k.startswith("sk-"):
        return "openai"
    return ""


def _es_modelo_claude(model: str) -> bool:
    return "claude" in (model or "").lower()


def _resolver_llm(*, api_key: str = "", model: str = "") -> tuple[str, str, str, str]:
    """Devuelve (provider, api_key, url, model). La clave pegada pisa Secrets."""
    pasted = (api_key or "").strip()
    tipo = _proveedor_de_clave(pasted)
    if tipo == "anthropic":
        elegido = (
            (model if _es_modelo_claude(model) else "")
            or os.environ.get("ANTHROPIC_MODEL")
            or "claude-sonnet-5"
        ).strip() or "claude-sonnet-5"
        return "anthropic", pasted, "https://api.anthropic.com/v1/messages", elegido
    if tipo == "xai":
        elegido = (
            (model if str(model).startswith("grok") else "")
            or os.environ.get("XAI_MODEL")
            or "grok-4.6"
        ).strip() or "grok-4.6"
        return "xai", pasted, "https://api.x.ai/v1/chat/completions", elegido
    if tipo == "openai":
        return (
            "openai",
            pasted,
            "https://api.openai.com/v1/chat/completions",
            (os.environ.get("OPENAI_MODEL") or "gpt-4o-mini").strip(),
        )
    if tipo == "groq":
        return (
            "groq",
            pasted,
            "https://api.groq.com/openai/v1/chat/completions",
            (os.environ.get("GROQ_MODEL") or "llama-3.3-70b-versatile").strip(),
        )
    anthropic = (os.environ.get("ANTHROPIC_API_KEY") or "").strip()
    xai = (os.environ.get("XAI_API_KEY") or "").strip()
    openai = (os.environ.get("OPENAI_API_KEY") or "").strip()
    groq = (os.environ.get("GROQ_API_KEY") or "").strip()
    if pasted:
        if _es_modelo_claude(model) or not xai:
            anthropic = pasted
        else:
            xai = pasted
    if anthropic:
        elegido = (
            (model if _es_modelo_claude(model) else "")
            or os.environ.get("ANTHROPIC_MODEL")
            or "claude-sonnet-5"
        ).strip() or "claude-sonnet-5"
        return "anthropic", anthropic, "https://api.anthropic.com/v1/messages", elegido
    if xai:
        elegido = (model or os.environ.get("XAI_MODEL") or "grok-4.6").strip() or "grok-4.6"
        return "xai", xai, "https://api.x.ai/v1/chat/completions", elegido
    if openai:
        return (
            "openai",
            openai,
            "https://api.openai.com/v1/chat/completions",
            (os.environ.get("OPENAI_MODEL") or "gpt-4o-mini").strip(),
        )
    if groq:
        return (
            "groq",
            groq,
            "https://api.groq.com/openai/v1/chat/completions",
            (os.environ.get("GROQ_MODEL") or "llama-3.3-70b-versatile").strip(),
        )
    return "", "", "", ""


def _texto_respuesta_claude(data: dict[str, Any]) -> str:
    partes: list[str] = []
    for bloque in data.get("content") or []:
        if isinstance(bloque, dict) and bloque.get("type") == "text":
            texto = str(bloque.get("text") or "").strip()
            if texto:
                partes.append(texto)
    return "\n\n".join(partes).strip()


def _bloques_imagen_claude(imagenes: list[dict[str, str]] | None) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for img in (imagenes or [])[:4]:
        m = _DATA_URL_IMG.match(str(img.get("data_url") or ""))
        if not m:
            continue
        out.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": m.group(1).lower(),
                "data": m.group(2),
            },
        })
    return out


def _mensajes_claude(
    messages: list[dict[str, Any]],
    imagenes: list[dict[str, str]] | None,
) -> list[dict[str, Any]]:
    limpios: list[dict[str, Any]] = []
    for m in messages:
        role = m.get("role")
        if role not in {"user", "assistant"}:
            continue
        content = m.get("content")
        if isinstance(content, str):
            content = content.strip()
            if not content:
                continue
        elif not content:
            continue
        if limpios and limpios[-1]["role"] == role and isinstance(limpios[-1]["content"], str) and isinstance(content, str):
            limpios[-1]["content"] = limpios[-1]["content"] + "\n\n" + content
        else:
            limpios.append({"role": role, "content": content})
    if limpios and limpios[0]["role"] != "user":
        limpios.insert(0, {"role": "user", "content": "(inicio)"})
    imgs = _bloques_imagen_claude(imagenes)
    if imgs:
        texto = "Leé la captura de Tango."
        if limpios and limpios[-1]["role"] == "user" and isinstance(limpios[-1]["content"], str):
            texto = limpios[-1]["content"] or texto
            limpios[-1]["content"] = [*imgs, {"type": "text", "text": texto}]
        else:
            limpios.append({"role": "user", "content": [*imgs, {"type": "text", "text": texto}]})
    if not limpios:
        limpios.append({"role": "user", "content": "Respondé la consulta de Tango."})
    return limpios


def _llamar_anthropic(
    system: str,
    messages: list[dict[str, Any]],
    *,
    key: str,
    modelo: str,
    imagenes: list[dict[str, str]] | None,
) -> str:
    candidatos = [modelo] if modelo else []
    for alt in _CLAUDE_MODELOS:
        if alt not in candidatos:
            candidatos.append(alt)
    msgs = _mensajes_claude(messages, imagenes)
    ultimo = ""
    for modelo_try in candidatos:
        payload = {
            "model": modelo_try,
            "max_tokens": 8192,
            "system": system,
            "messages": msgs,
        }
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "x-api-key": key,
                "anthropic-version": "2023-06-01",
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
            raise RuntimeError(f"No se llegó a Anthropic: {exc.reason}") from exc
        texto = _texto_respuesta_claude(data)
        if texto:
            return texto
    if ultimo:
        raise RuntimeError(ultimo)
    return ""


def _llamar_llm(
    system: str,
    messages: list[dict[str, Any]],
    *,
    api_key: str = "",
    model: str = "",
    imagenes: list[dict[str, str]] | None = None,
) -> str:
    provider, key, url, modelo = _resolver_llm(api_key=api_key, model=model)
    if not key or not url:
        return ""
    if provider == "anthropic":
        return _llamar_anthropic(
            system,
            messages,
            key=key,
            modelo=modelo,
            imagenes=imagenes,
        )
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
        partes: list[dict[str, Any]] = []
        for img in imagenes[:4]:
            partes.append({
                "type": "image_url",
                "image_url": {"url": img["data_url"]},
            })
        partes.append({"type": "text", "text": texto})
        messages = [*messages[:-1], {"role": "user", "content": partes}]
    candidatos = [modelo]
    if "x.ai" in url:
        for alt in _GROK_MODELOS:
            if alt not in candidatos:
                candidatos.append(alt)
    extras: list[dict[str, Any]] = [{}]
    if "x.ai" in url:
        esfuerzo = "low" if imagenes else "medium"
        extras = [{"reasoning_effort": esfuerzo}, {}]
    else:
        extras = [{"max_tokens": 4096}]
    ultimo = ""
    timeout = 90 if imagenes else 150
    for modelo_try in candidatos:
        salto_modelo = False
        for extra in extras:
            payload = {
                "model": modelo_try,
                "temperature": 0.7,
                "messages": [{"role": "system", "content": system}, *messages],
            }
            payload.update(extra)
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
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as exc:
                body = exc.read().decode("utf-8", errors="replace")[:400]
                ultimo = f"IA HTTP {exc.code}: {body}"
                if exc.code == 413:
                    raise RuntimeError(ultimo) from exc
                if extra.get("reasoning_effort") and "reasoning" in body.lower():
                    continue
                if exc.code in {400, 404} and "model" in body.lower():
                    salto_modelo = True
                    break
                if modelo_try != candidatos[-1] or extra != extras[-1]:
                    continue
                raise RuntimeError(ultimo) from exc
            except urllib.error.URLError as exc:
                raise RuntimeError(f"No se llegó a la IA: {exc.reason}") from exc
            choices = data.get("choices") or []
            if not choices:
                ultimo = "Grok no devolvió opciones"
                continue
            choice = choices[0] if isinstance(choices[0], dict) else {}
            texto_out = _texto_mensaje_llm(choice.get("message") or {})
            if texto_out:
                return texto_out
            fr = str(choice.get("finish_reason") or "vacio")
            ultimo = f"Grok no devolvió texto (finish_reason={fr})"
        if salto_modelo:
            continue
    if ultimo:
        raise RuntimeError(ultimo)
    return ""


_MEMORIA_ESTUDIO = """
Información del estudio. Formulá la respuesta con esto, con tus palabras. No lo copies de corrido:
- Tango Sueldos: SI(cond;verdadero;falso), OR = O, USUELD (nunca USUELO), códigos entre comillas ("099").
- CTRATO "099" y "048" se excluyen envolviendo Importe. "001" = 50% de IMPOR, "008" = 100%. El % va en Importe, no en Cantidad.
- Concepto 1 sueldo básico = USUELD/30*CANTIDAD. Si cambia una fórmula, refrescar el concepto y reliquidar.
- Papeles de bancos: solo el primer mes toma el saldo inicial del extracto; el resto arrastra. No forzar la diferencia a 0. No inventar meses sin extracto. Transferencias propias: una vez del lado que recibe.
- Monotributo: el período es fecha desde/hasta, no la fecha de emisión.
"""

_SYSTEM = """Sos Grok en el chat del Estudio Contable.

Tenés IA: pensá y explicá como en grok.com, en español. La respuesta la ARMÁS vos usando la información del estudio y el material que te pasamos. No inventes menús ni fórmulas si ya está en ese material. No pegues manuales ni URLs.

No pidas ni escribas claves, TANGO.INI ni contraseñas.
""" + _MEMORIA_ESTUDIO

_SYSTEM_FORMULA = _SYSTEM + """
Si piden fórmula de sueldos, devolvé FormulaImporte y FormulaCantidad en bloques ``` listos para pegar, usando el material (no reescribas una fórmula del export si ya viene).
"""

_SYSTEM_CAPTURA = _SYSTEM + """
Hay una captura: mirala y formulá cómo seguir desde esa pantalla, con la información del estudio.
"""


def _texto_mensaje_llm(message: dict[str, Any] | None) -> str:
    """Grok a veces devuelve content como lista, no como texto."""
    if not isinstance(message, dict):
        return ""
    content = message.get("content")
    if isinstance(content, str) and content.strip():
        return content.strip()
    extra = message.get("output_text")
    if isinstance(extra, str) and extra.strip():
        return extra.strip()
    partes: list[str] = []
    if isinstance(content, list):
        for bloque in content:
            if isinstance(bloque, str) and bloque.strip():
                partes.append(bloque.strip())
                continue
            if not isinstance(bloque, dict):
                continue
            if bloque.get("type") in {"text", "output_text", None}:
                t = str(bloque.get("text") or bloque.get("content") or "").strip()
                if t:
                    partes.append(t)
    return "\n\n".join(partes).strip()


def _es_doc_formula(doc: dict[str, str]) -> bool:
    source = (doc.get("source") or "").lower()
    title = (doc.get("title") or "").lower()
    return (
        "formulas_completo.csv" in source
        or "sueldos_formulas" in source
        or "huerfanas" in source
        or "huerfanas" in title
    )


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
        return None
    q = _tokens(pregunta)
    mejor: tuple[float, dict[str, str]] | None = None
    for h in csv_hits:
        blob = f"{h.get('title', '')} {h.get('text', '')}"
        score = float(len(q & _tokens(blob)))
        if any(w in titulo(h) for w in q):
            score += 3.0
        if mejor is None or score > mejor[0]:
            mejor = (score, h)
    if mejor and mejor[0] >= 3.0:
        return mejor[1]
    return None


def _pide_formula(pregunta: str) -> bool:
    return _es_consulta_formula(pregunta) or _pide_pegar(pregunta)


def _completar_formula_copiable(
    texto: str,
    formula_hit: dict[str, str] | None,
    pregunta: str,
) -> str:
    """Si la IA formuló sin bloques ```, agrega Importe/Cantidad listos para pegar."""
    if not _pide_formula(pregunta):
        return texto
    if "```" in (texto or ""):
        return texto
    if not formula_hit:
        return texto
    extra = _formula_para_pegar(formula_hit)
    if (texto or "").strip():
        return texto.rstrip() + "\n\n" + extra
    return extra


def _error_ia_amigable(error: str) -> str:
    e = (error or "").lower()
    if "incorrect api key" in e or "invalid api key" in e or "authentication" in e:
        return (
            "Grok no arrancó: la clave de xAI no vale o no está. "
            "En Tango, pegá una clave nueva de https://console.x.ai (empieza con xai-) y tocá Activar Grok."
        )
    if error:
        return f"La IA no respondió: {error[:160]}"
    return ""


def _respuesta_proceso_arca() -> str:
    return (
        "El proceso es así, en el chat (no hace falta ningún PDF):\n\n"
        "**1.** En ARCA entrá a **Portal IVA** con clave fiscal.\n"
        "**2.** Nueva declaración jurada → elegí el período → **Libro de IVA**.\n"
        "**3.** Compras o Ventas → **Importar comprobantes desde ARCA** → descargá el **CSV** "
        "(si viene en ZIP, descomprimilo; Tango no importa el ZIP).\n"
        "**4.** En Tango: **Liquidador de IVA → Comprobantes → Importación de comprobantes desde ARCA**.\n"
        "**5.** Origen: Libro IVA compras o ventas. Examiná y elegí el `.csv`.\n"
        "**6.** Completá las equivalencias de tipos de comprobante ARCA ↔ Tango "
        "(la primera vez, o si aparece un tipo nuevo).\n"
        "**7.** Si hay duplicados: solo los nuevos, o nuevos + actualizar.\n"
        "**8.** Confirmá. Al final Tango arma un Excel de qué entró y qué se rechazó.\n\n"
        "Importante: **esta importación no genera el asiento**. Después tenés que ir a "
        "**Generación de asientos contables** (modelos PIVAC compras / PIVAV ventas).\n\n"
        "Si un comprobante no entra, casi siempre es equivalencia de tipo, cliente/proveedor, "
        "o una percepción que hay que reimputar a mano en la registración."
    )


def _fallback(
    pregunta: str,
    hits: list[dict[str, str]],
    imagenes: list[dict[str, str]] | None = None,
) -> str:
    if imagenes:
        return (
            "Vi la captura, pero esta vez no pude leerla. "
            "Mandala de nuevo o decime el menú de arriba de Tango "
            "(módulo y pantalla) y te digo el siguiente click."
        )
    p = _sin_acento(pregunta or "").lower()
    if _es_consulta_proceso(pregunta) and re.search(r"arca|portal iva|importar", p):
        return _respuesta_proceso_arca()
    if _pide_formula(pregunta):
        hit = _mejor_formula(pregunta, hits)
        if hit:
            return (
                "Esta es la fórmula, lista para pegar en Tango (no es un PDF):\n\n"
                + _formula_para_pegar(hit)
            )
    utiles = [h for h in hits if not _es_doc_formula(h)]
    if not utiles:
        return (
            "No pude armar la respuesta con Grok ahora. "
            "Preguntame de nuevo o decime el menú de Tango (IVA, Sueldos, Compras)."
        )
    return (
        "No pude armar la respuesta con Grok ahora. "
        "Preguntame de nuevo con un poco más de detalle (pantalla o concepto)."
    )


def preparar_imagen(raw: bytes, nombre: str = "captura.png") -> dict[str, str]:
    """Comprime la captura a JPEG chico para visión. Devuelve data_url."""
    data = raw
    try:
        if Image is None:
            raise RuntimeError("sin Pillow")
        im = Image.open(BytesIO(raw)).convert("RGB")
        im.thumbnail((1280, 1280))
        data = b""
        for calidad in (80, 70, 58, 45):
            buf = BytesIO()
            im.save(buf, format="JPEG", quality=calidad, optimize=True)
            data = buf.getvalue()
            if len(data) <= 350_000:
                break
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
    b64 = base64.b64encode(data).decode("ascii")
    return {"data_url": f"data:image/jpeg;base64,{b64}", "nombre": Path(nombre).name}


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
        consulta = "Leé la captura de Tango: qué pantalla es, y cómo sigo desde acá."
    hay_foto = bool(imagenes)
    seguimiento = hay_foto and _es_seguimiento_visual(consulta)
    busqueda = consulta if hay_foto else _texto_busqueda(consulta, historial)
    hits = recuperar(busqueda if _pide_pegar(consulta) else consulta, docs, k=4 if hay_foto else 8)
    formula_hit = None
    pide_formula = _pide_formula(busqueda) or _pide_formula(consulta)
    if pide_formula:
        formula_hit = _mejor_formula(busqueda, _docs_formula(docs) or hits)
        if formula_hit:
            hits = [formula_hit] + [h for h in hits if h is not formula_hit][:7]
    _provider, key, _url, _modelo = _resolver_llm(api_key=api_key, model=model)
    if not key:
        extra = ""
        if formula_hit:
            extra = (
                "\n\nMientras tanto, del export de sueldos (sin Grok):\n\n"
                + _formula_para_pegar(formula_hit)
            )
        return {
            "texto": (
                "Este chat **todavía no está usando Grok**. Falta la clave de xAI.\n\n"
                "En **Tango**, pegá `XAI_API_KEY` (empieza con `xai-`) y tocá **Activar Grok**. "
                "La clave se crea en https://console.x.ai"
                + extra
            ),
            "fuentes": [
                {"title": formula_hit.get("title") or "", "source": Path(str(formula_hit.get("source") or "")).name}
            ] if formula_hit else [],
            "docs": len(docs),
            "uso_ia": False,
        }
    bloques_ctx: list[str] = []
    if formula_hit:
        bloques_ctx.append(
            "### Fórmula del export Tango Sueldos (fuente de verdad)\n"
            + _explicar_formula(formula_hit)
        )
    extra_hits = [h for h in hits if h is not formula_hit]
    if hay_foto:
        extra_hits = [] if seguimiento else extra_hits[:1]
    else:
        extra_hits = extra_hits[:3]
    for h in extra_hits:
        texto_h = (h.get("text") or "").strip()
        if _es_doc_formula(h):
            texto_h = _explicar_formula(h)
        else:
            texto_h = re.sub(r"\s+\|\s+", "\n", texto_h)[:600]
        bloques_ctx.append(
            f"### {h.get('title')}\n{texto_h}"
        )
    contexto = "\n\n".join(bloques_ctx)
    msgs: list[dict[str, Any]] = []
    hist_usar = (historial or [])[-4:] if hay_foto else (historial or [])[-10:]
    for m in hist_usar:
        role = m.get("role") or "user"
        if role not in {"user", "assistant"}:
            continue
        content = str(m.get("content") or "").strip()
        if hay_foto and role == "assistant":
            content = content[:800]
        if content:
            msgs.append({"role": role, "content": content})
    if imagenes and not consulta.strip():
        consulta = "¿Qué pantalla es y cómo sigo desde acá?"
    msgs.append({"role": "user", "content": consulta})
    if hay_foto:
        sistema = _SYSTEM_CAPTURA
    elif pide_formula:
        sistema = _SYSTEM_FORMULA
    else:
        sistema = _SYSTEM
    if contexto:
        sistema = (
            sistema
            + "\n\nMaterial del estudio para FORMULAR la respuesta "
            "(no lo pegues literal; usalo):\n"
            + contexto
        )
    ia = ""
    error = ""
    try:
        ia = _llamar_llm(
            sistema,
            msgs,
            api_key=api_key,
            model=model,
            imagenes=imagenes,
        )
    except Exception as exc:
        error = str(exc)
    if ia:
        texto = _completar_formula_copiable(ia, formula_hit, consulta)
    elif formula_hit and not hay_foto:
        texto = _formula_para_pegar(formula_hit)
    else:
        texto = _fallback(consulta, hits, imagenes)
        aviso = _error_ia_amigable(error)
        err_l = (error or "").lower()
        if aviso and ("incorrect api key" in err_l or "invalid api key" in err_l or "authentication" in err_l):
            texto = aviso
        elif aviso:
            texto = f"{texto}\n\n_({aviso})_"
    return {
        "texto": texto,
        "fuentes": [
            {"title": h.get("title") or "", "source": Path(str(h.get("source") or "")).name}
            for h in hits[:6]
        ],
        "docs": len(docs),
        "uso_ia": bool(ia),
    }
