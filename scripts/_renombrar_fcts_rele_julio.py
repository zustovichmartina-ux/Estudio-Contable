# -*- coding: utf-8 -*-
"""Renombra facturas de compra RELE jul-2026: FCCA13-157 PROVEEDOR."""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

FOLDER = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026\07-2026"
)

CLIENTE_TOKENS = (
    "OFTALMOLOGIA RELE",
    "OFTALMOLOGIA",
    "RELE MAR DEL PLATA",
)

AFIP_COD = {
    1: ("FCC", "A"),
    2: ("NDC", "A"),
    3: ("NCC", "A"),
    6: ("FCC", "B"),
    7: ("NDC", "B"),
    8: ("NCC", "B"),
    11: ("FCC", "C"),
    12: ("NDC", "C"),
    13: ("NCC", "C"),
    51: ("FCC", "M"),
    201: ("FCC", "A"),
    203: ("NCC", "A"),
}

INVALID = re.compile(r'[\\/:*?"<>|]+')
RE_AFIP_NAME = re.compile(r"(?<!\d)(\d{11})_(\d{2,3})_(\d{1,5})_(\d{1,8})")

SKIP_OCR = True
_OCR = None


def _norm(s: str) -> str:
    t = unicodedata.normalize("NFKD", str(s or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip()


def _safe_prov(nombre: str) -> str:
    t = _norm(nombre).upper()
    t = t.replace("S.R.L.", "SRL").replace("S.A.", "SA")
    t = re.split(r"\bFECHA\s+DE\s+EMISION\b", t, flags=re.I)[0]
    t = re.split(r"\bCUIT\b", t, flags=re.I)[0]
    t = re.sub(r"[,.;:]+", " ", t)
    t = INVALID.sub(" ", t)
    t = re.sub(r"\s+", " ", t).strip(" -_")
    t = re.sub(r"\bHOJA\s*\d+\b", "", t).strip()
    return t[:80]


def _hoja(nombre: str) -> str:
    m = re.search(r"HOJA\s*(\d+)", nombre, re.I)
    return f" HOJA {m.group(1)}" if m else ""


def _texto_pdf(path: Path) -> str:
    try:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages[:3])
    except Exception:
        pass
    try:
        import fitz

        doc = fitz.open(path)
        t = "\n".join(page.get_text("text") or "" for page in doc)
        doc.close()
        return t
    except Exception:
        return ""


def _texto_imagen(path: Path) -> str:
    if SKIP_OCR:
        return ""


def _desde_nombre(nombre: str) -> dict:
    stem = Path(nombre).stem
    out: dict = {}
    compact = re.sub(r"\s+", "", stem)
    m = re.search(r"DDCA(\d{5})(\d{6,8})", compact, re.I)
    if m:
        out["tipo"], out["letra"] = "NDC", "A"
        out["pv"] = int(m.group(1))
        out["nro"] = int(m.group(2))
        out["prov"] = "BIOMAT SRL"

    m = RE_AFIP_NAME.search(nombre.replace(" ", "_"))
    if m:
        out["tipo"], out["letra"] = AFIP_COD.get(int(m.group(2)), ("FCC", "A"))
        out["pv"] = int(m.group(3))
        out["nro"] = int(m.group(4))
        pref = _safe_prov(re.sub(r"\d+", " ", nombre[: m.start()]))
        resto = nombre[m.end() :]
        resto = re.sub(r"(?i)^[\s_\-]*FC\s*", " ", resto)
        resto = _safe_prov(Path(resto).stem)
        if pref and pref not in {"DR", "DRA"}:
            out["prov"] = pref
        elif resto:
            out["prov"] = resto
        return out

    m = re.search(
        r"(?:F-A|FA-A|FC\s*ELEC\s*A|FC\s*A|FA\s*A|FC-A|DDCA)"
        r"[\s_-]*0*(\d{1,5})[\s_-]+0*(\d{3,8})",
        stem,
        re.I,
    )
    if m:
        out["tipo"] = "NCC" if re.search(r"(?i)\bNC", stem) else "FCC"
        if re.search(r"(?i)DDC", stem):
            out["tipo"] = "NDC"
        out["letra"] = "A"
        out["pv"] = int(m.group(1))
        out["nro"] = int(m.group(2))
    else:
        m = re.search(r"FC\s*A\s*(\d{1,5})-(\d{3,8})", stem, re.I)
        if m:
            out["tipo"], out["letra"] = "FCC", "A"
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
        else:
            m = re.search(
                r"(?:FC|FA|NC)[\s_-]*[A-C]?[\s_-]*0*(\d{4,8})\b",
                stem,
                re.I,
            )
            if m:
                out["tipo"] = "NCC" if re.search(r"(?i)\bNC", stem) else "FCC"
                out["letra"] = "A"
                out["nro"] = int(m.group(1))
                out["pv"] = 0

    m = re.search(r"(\d{1,5})-A-(\d{1,5})-(\d{5,8})", stem, re.I)
    if m:
        out["tipo"], out["letra"] = "FCC", "A"
        out["pv"], out["nro"] = int(m.group(2)), int(m.group(3))

    m = re.search(r"006_(\d{5})_(\d{8,})", stem)
    if m:
        out["tipo"], out["letra"] = "FCC", "B"
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))

    # proveedor al final o al inicio
    m = re.search(
        r"(?:[-–]|FC\s+|FA\s+)\s*"
        r"(BIOMAT|ALCON|VETRA|OACI|REBORA|RABBIONE|QUALITY CLEAN|QUALITY|"
        r"IMPLANTEC|PAYWAY|MED SRL|MSZ|GSJ SA|GSJ|OXFORD|CAMOGA|LIMPA|"
        r"CITYSSAN|GAONA|INTEGRAL PACK|INTEGRAL|ALEMAN|CISNEROS|BRAVO|"
        r"LAMARE|OLSINA|SUAREZ|LARA|RODRIGUEZ DIEGO|RODRIGUEZ|GONZALEZ HUGO|"
        r"GONZALEZ|BARROS|VIA CARGO|MAR DEL PLATA SODA|SODA|LIBRERIA OXFORD|"
        r"PAPELERA CAMOGA|FARMACIA MAGISTER|GAONA 3050SRL)",
        stem,
        re.I,
    )
    if m:
        out["prov"] = _safe_prov(m.group(1))
    if "prov" not in out:
        m = re.search(r"^([A-Za-zÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑ .&]{2,40})\s*[-–]", stem)
        if m and not m.group(1).upper().startswith("FC"):
            out["prov"] = _safe_prov(m.group(1))
    if "prov" not in out:
        m = re.search(r"[-–]\s*([A-Za-zÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑ0-9 .&]{2,50})$", stem)
        if m:
            out["prov"] = _safe_prov(m.group(1))
    if "SODA" in stem.upper():
        out["prov"] = "MAR DEL PLATA SODA"
    return out


def _desde_texto(texto: str) -> dict:
    out: dict = {}
    m = re.search(r"\bCod(?:igo)?\.?\s*(\d{1,3})\b", texto[:2500], re.I)
    if m:
        pair = AFIP_COD.get(int(m.group(1)))
        if pair:
            out["tipo"], out["letra"] = pair
    n = _norm(texto[:2500]).upper()
    if "tipo" not in out:
        if re.search(r"NOTA DE CREDITO", n):
            out["tipo"] = "NCC"
            out["letra"] = "C" if "CREDITO C" in n or "COD 013" in n else (
                "B" if "CREDITO B" in n or "COD 8" in n else "A"
            )
        elif re.search(r"NOTA DE DEBITO", n):
            out["tipo"] = "NDC"
            out["letra"] = "A"
        elif "FACTURA C" in n or "COD 011" in n:
            out["tipo"], out["letra"] = "FCC", "C"
        elif "FACTURA B" in n or "COD 006" in n:
            out["tipo"], out["letra"] = "FCC", "B"
        elif "FACTURA A" in n or "COD 001" in n:
            out["tipo"], out["letra"] = "FCC", "A"

    m = re.search(
        r"(?:Punto\s+de\s+Venta|Pto\.?\s*Vta)\s*:?\s*(\d{1,5})"
        r".{0,80}?"
        r"(?:Comp(?:robante)?\.?\s*N(?:ro|um)?\.?)\s*:?\s*(\d{1,8})",
        texto,
        re.I | re.S,
    )
    if m:
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
    if "nro" not in out:
        m = re.search(r"\b(\d{4,5})\s*[-–]\s*(\d{5,8})\b", texto[:1500])
        if m:
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))

    for m in re.finditer(r"Raz[oó]n\s+Social\s*:?\s*([^\n]{3,80})", texto, re.I):
        cand = _safe_prov(m.group(1))
        if cand and not any(tok in cand for tok in CLIENTE_TOKENS):
            out["prov"] = cand
            break
    if "prov" not in out:
        m = re.search(r"Apellido\s+y\s+Nombre\s*:?\s*([^\n]{3,80})", texto, re.I)
        if m:
            cand = _safe_prov(m.group(1))
            if cand and not any(tok in cand for tok in CLIENTE_TOKENS):
                out["prov"] = cand
    return out


def armar_nombre(info: dict, orig: str) -> str:
    tipo = info.get("tipo") or "FCC"
    letra = info.get("letra") or "A"
    pv = int(info.get("pv") or 0)
    nro = int(info.get("nro") or 0)
    prov = info.get("prov") or "PROVEEDOR"
    hoja = _hoja(orig)
    ext = Path(orig).suffix.lower()
    if pv and nro:
        base = f"{tipo}{letra}{pv}-{nro} {prov}{hoja}"
    elif nro:
        base = f"{tipo}{letra}{nro} {prov}{hoja}"
    else:
        base = f"{tipo}{letra} {prov}{hoja}"
    base = INVALID.sub(" ", base)
    return re.sub(r"\s+", " ", base).strip() + ext


def parsear(path: Path) -> dict:
    nombre = path.name
    info = _desde_nombre(nombre)
    need_txt = not (info.get("nro") and info.get("prov") and info.get("tipo"))
    texto = ""
    if need_txt:
        if path.suffix.lower() == ".pdf":
            texto = _texto_pdf(path)
        else:
            texto = _texto_imagen(path)
        extra = _desde_texto(texto)
        for k, v in extra.items():
            if k not in info or not info.get(k):
                info[k] = v
    nuevo = armar_nombre(info, nombre)
    return {
        "path": path,
        "orig": nombre,
        "nuevo": nuevo,
        "info": info,
        "len_texto": len(texto),
        "completo": bool(info.get("nro") and info.get("prov")),
    }


def unique_target(folder: Path, name: str, used: set[str]) -> str:
    if name.lower() not in used:
        used.add(name.lower())
        return name
    stem = Path(name).stem
    ext = Path(name).suffix
    i = 2
    while True:
        cand = f"{stem} ({i}){ext}"
        if cand.lower() not in used:
            used.add(cand.lower())
            return cand
        i += 1


def main() -> None:
    files = [
        p
        for p in FOLDER.iterdir()
        if p.is_file() and p.suffix.lower() in {".pdf", ".jpeg", ".jpg", ".png", ".webp"}
    ]
    print("carpeta", FOLDER)
    print("archivos", len(files))
    rows = []
    for p in sorted(files, key=lambda x: x.name.lower()):
        row = parsear(p)
        rows.append(row)
        flag = "OK" if row["completo"] else "REVISAR"
        print(f"[{flag}] {row['orig'][:52]:<52} -> {row['nuevo']}")

    used: set[str] = set()
    # no chocar con nombres que no vamos a tocar
    dests = []
    for row in rows:
        dest = unique_target(FOLDER, row["nuevo"], used)
        dests.append((row, dest))

    renamed = 0
    for row, dest in dests:
        src: Path = row["path"]
        dst = FOLDER / dest
        if src.name.lower() == dest.lower():
            continue
        if dst.exists() and dst.resolve() != src.resolve():
            dest = unique_target(FOLDER, dest, used)
            dst = FOLDER / dest
        print(f"REN {src.name} => {dest}")
        src.rename(dst)
        renamed += 1
    print(f"renombrados={renamed} total={len(rows)}")


if __name__ == "__main__":
    main()
