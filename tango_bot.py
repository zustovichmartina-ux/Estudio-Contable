# -*- coding: utf-8 -*-
"""Bot Tango: índice de ayudas Axoft + reglas del estudio. Nunca guarda claves."""
from __future__ import annotations

import csv
import json
import os
import re
import urllib.error
import urllib.request
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

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
    with path.open(encoding="utf-8", errors="replace", newline="") as fh:
        reader = csv.DictReader(fh)
        for i, row in enumerate(reader):
            if i > 400:
                break
            partes = [f"{k}: {v}" for k, v in row.items() if str(v or "").strip()]
            if not partes:
                continue
            codigo = str(row.get("Codigo") or row.get("TIPO_CONCEPTO") or i)
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


def _tokens(texto: str) -> set[str]:
    return {t.lower() for t in _TOKEN.findall(texto or "")}


def recuperar(pregunta: str, docs: list[dict[str, str]], k: int = 8) -> list[dict[str, str]]:
    q = _tokens(pregunta)
    if not q or not docs:
        return docs[:k]
    scored: list[tuple[float, dict[str, str]]] = []
    for doc in docs:
        blob = f"{doc.get('title', '')} {doc.get('text', '')}"
        t = _tokens(blob)
        if not t:
            continue
        hit = len(q & t)
        if hit == 0:
            continue
        extra = 1.5 if any(w in (doc.get("title") or "").lower() for w in q) else 1.0
        scored.append((hit * extra, doc))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [d for _, d in scored[:k]]


def _llamar_grok(
    system: str,
    messages: list[dict[str, str]],
    *,
    api_key: str = "",
    model: str = "",
) -> str:
    key = (api_key or os.environ.get("XAI_API_KEY") or "").strip()
    if not key:
        return ""
    model = (model or os.environ.get("XAI_MODEL") or "grok-4").strip() or "grok-4"
    payload = {
        "model": model,
        "temperature": 0.2,
        "messages": [{"role": "system", "content": system}, *messages],
    }
    req = urllib.request.Request(
        "https://api.x.ai/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "User-Agent": "EstudioContable-TangoBot/1.0",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"Grok HTTP {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"No se llegó a xAI: {exc.reason}") from exc
    choices = data.get("choices") or []
    if not choices:
        return ""
    return str((choices[0].get("message") or {}).get("content") or "").strip()


_SYSTEM = """Sos el asistente Tango del Estudio Contable (Axoft Tango Estudios, ayudas 26ar).
Respondé en español, claro y ordenado, como un compañero del estudio.
Usá SOLO el contexto que te pasan (ayudas Axoft recolectadas + reglas del estudio).
Si el menú o la fórmula no está en el contexto, decilo: no inventes pantallas.
Nunca pidas ni escribas claves SQL, TANGO.INI ni contraseñas.
Si hay dos caminos (asiento por comprobante vs determinación mensual), distinguilos.
Cuando cites un menú, usá la ruta tal cual (ej. Liquidador de IVA > Archivos > …).
"""


def _fallback(pregunta: str, hits: list[dict[str, str]]) -> str:
    if not hits:
        return (
            "No lo tengo en las ayudas Tango del estudio. "
            "Probá con el nombre de la pantalla (Asientos, Modelos, Sueldos, IVA)."
        )
    bloques = []
    for h in hits[:3]:
        titulo = (h.get("title") or "Ayuda").strip()
        texto = (h.get("text") or "").strip()[:800]
        bloques.append(f"**{titulo}**\n\n{texto}")
    return "Según las ayudas del estudio:\n\n" + "\n\n".join(bloques)


def responder(
    pregunta: str,
    historial: list[dict[str, str]] | None = None,
    *,
    api_key: str = "",
    model: str = "",
) -> dict[str, Any]:
    docs = cargar_indice()
    hits = recuperar(pregunta, docs, k=8)
    contexto = "\n\n".join(
        f"### {h.get('title')}\nFuente: {Path(str(h.get('source') or '')).name}\n{h.get('text')}"
        for h in hits
    )
    msgs: list[dict[str, str]] = []
    for m in (historial or [])[-8:]:
        role = m.get("role") or "user"
        if role not in {"user", "assistant"}:
            continue
        content = str(m.get("content") or "").strip()
        if content:
            msgs.append({"role": role, "content": content})
    msgs.append({
        "role": "user",
        "content": f"Pregunta:\n{pregunta}\n\nContexto:\n{contexto}",
    })
    grok = ""
    error = ""
    try:
        grok = _llamar_grok(_SYSTEM, msgs, api_key=api_key, model=model)
    except Exception as exc:
        error = str(exc)
    texto = grok or _fallback(pregunta, hits)
    if error and not grok:
        texto = f"{texto}\n\n_(Grok no respondió: {error[:180]})_"
    return {
        "texto": texto,
        "fuentes": [
            {"title": h.get("title") or "", "source": Path(str(h.get("source") or "")).name}
            for h in hits[:6]
        ],
        "docs": len(docs),
        "uso_grok": bool(grok),
    }
