# -*- coding: utf-8 -*-
"""Renombra FCC/NCC en Facturas/2026 (excepto 07-2026 ya hecho)."""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

ROOT = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026"
)
SKIP_MESES = {"07-2026"}
CLIENTE_TOKENS = ("OFTALMOLOGIA RELE", "OFTALMOLOGIA", "RELE MAR DEL PLATA")

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
    15: ("FCC", "C"),
    51: ("FCC", "M"),
    81: ("FCC", "A"),
    201: ("FCC", "A"),
    203: ("NCC", "A"),
}

CUIT_PROV = {
    "20175365290": "CISNEROS ALEJANDRO DANIEL",
    "20206306913": "DEBIASSI",
    "20257163408": "TRUJILLO HERNAN FEDERICO",
    "20267692085": "ANESTESIA",
    "20272016985": "MARINO",
    "20272723878": "CONTENEDOR",
    "20288161225": "ANESTESIA",
    "20300059814": "ANESTESIA",
    "20343673605": "SUAREZ FACUNDO",
    "20347593657": "RODRIGUEZ DIEGO",
    "23229161539": "ADRIAN",
    "23435091989": "JUAN CRUZ",
    "27245392848": "REBORA",
    "27261473033": "ALEMAN",
    "27278627727": "OLSINA",
    "27280544669": "LAMARE",
    "27284539821": "BRAVO",
    "27415691292": "BARROS ABRIL",
    "27459230446": "LARA LEVINGSTON",
    "30516842667": "ALCON",
    "30525827433": "ASTATEC SA",
    "30541190518": "CENTRO MEDICO DE MAR DEL PLATA",
    "30545778137": "COLON SA",
    "30561820240": "MAR DEL PLATA SODA",
    "30591087793": "OACI SA",
    "30605410002": "MED SRL",
    "30617008153": "VETRA",
    "30634473129": "CLUB",
    "30707016023": "IMPLANTEC",
    "30708278722": "BIOMAT SRL",
    "30710463901": "QUALITY CLEAN",
    "30711425124": "GSJ SA",
    "30717664406": "PAYWAY",
}

INVALID = re.compile(r'[\\/:*?"<>|]+')
RE_AFIP_GUION = re.compile(r"(?<!\d)(\d{11})[-_](\d{2,3})[-_](\d{1,5})[-_](\d{5,11})")
RE_YA = re.compile(r"^(FCC|NCC|NDC)[ABC]", re.I)

SKIP_NAMES = re.compile(
    r"thumbs\.db|comprobantes faltantes|\.xlsx$|adobe scan",
    re.I,
)


def _norm(s: str) -> str:
    t = unicodedata.normalize("NFKD", str(s or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip()


def _safe_prov(nombre: str) -> str:
    t = _norm(nombre).upper()
    t = t.replace("S.R.L.", "SRL").replace("S.A.", "SA")
    t = re.split(r"\bFECHA\s+DE\s+EMISION\b", t, flags=re.I)[0]
    t = re.split(r"\bCUIT\b", t, flags=re.I)[0]
    t = re.sub(r"(?i)\bcopia\b|\(\d+\)", " ", t)
    t = re.sub(r"[,.;:]+", " ", t)
    t = INVALID.sub(" ", t)
    t = re.sub(r"\s+", " ", t).strip(" -_")
    t = re.sub(r"\bHOJA\s*\d+\b", "", t).strip()
    for tok in CLIENTE_TOKENS:
        if tok in t:
            t = t.replace(tok, " ").strip()
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


def _desde_nombre(nombre: str) -> dict:
    stem = Path(nombre).stem
    out: dict = {}
    m = RE_AFIP_GUION.search(nombre.replace(" ", "_"))
    if m:
        cuit, cod, pv, nro = m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4))
        out["tipo"], out["letra"] = AFIP_COD.get(cod, ("FCC", "A"))
        out["pv"], out["nro"] = pv, nro
        if cuit in CUIT_PROV:
            out["prov"] = CUIT_PROV[cuit]
        pref = _safe_prov(re.sub(r"\d+", " ", nombre[: m.start()]))
        resto = _safe_prov(Path(nombre[m.end() :]).stem)
        if pref and pref not in {"DR", "DRA", "FACU"}:
            out["prov"] = pref
        elif resto and len(resto) > 2:
            out.setdefault("prov", resto)
        return out

    m = re.search(r"VT-FA0*(\d+)-(\d+)", stem, re.I)
    if m:
        out["tipo"], out["letra"] = "FCC", "A"
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
        out["prov"] = "ASTATEC SA"
        return out

    m = re.search(r"FAC-A-0*(\d{1,5})-0*(\d{3,8})", stem, re.I)
    if m:
        out["tipo"], out["letra"] = "FCC", "A"
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))

    m = re.search(r"FACA0*(\d)0*(\d{5,8})", stem, re.I)
    if m and "pv" not in out:
        out["tipo"], out["letra"] = "FCC", "A"
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))

    for label, prov in (
        ("ALCON", "ALCON"),
        ("BIOMAT", "BIOMAT SRL"),
        ("VETRA", "VETRA"),
        ("OACI", "OACI SA"),
        ("REBORA", "REBORA"),
        ("IMPLANTEC", "IMPLANTEC"),
        ("PAYWAY", "PAYWAY"),
        ("QUALITY", "QUALITY CLEAN"),
        ("COLON", "COLON SA"),
        ("ASTATEC", "ASTATEC SA"),
        ("CISNEROS", "CISNEROS"),
        ("BRAVO", "BRAVO"),
        ("LAMARE", "LAMARE"),
        ("OLSINA", "OLSINA"),
        ("BARROS", "BARROS ABRIL"),
        ("SUAREZ", "SUAREZ FACUNDO"),
        ("LARA", "LARA LEVINGSTON"),
        ("ALEMAN", "ALEMAN"),
        ("RODRIGUEZ", "RODRIGUEZ DIEGO"),
        ("SODA", "MAR DEL PLATA SODA"),
        ("CLINICA", "COLON SA"),
    ):
        if re.search(label, stem, re.I):
            out.setdefault("prov", prov)
            break
    return out


def _desde_texto(texto: str) -> dict:
    out: dict = {}
    m = re.search(r"\bCod(?:igo)?\.?\s*:?\s*N?[ºo°]?\s*(\d{1,3})\b", texto[:2500], re.I)
    if m:
        pair = AFIP_COD.get(int(m.group(1)))
        if pair:
            out["tipo"], out["letra"] = pair
    n = _norm(texto[:2500]).upper()
    if "tipo" not in out:
        if re.search(r"NOTA DE CREDITO", n):
            out["tipo"], out["letra"] = "NCC", "A"
        elif re.search(r"NOTA DE DEBITO", n):
            out["tipo"], out["letra"] = "NDC", "A"
        elif "FACTURA C" in n or "COD 011" in n:
            out["tipo"], out["letra"] = "FCC", "C"
        elif "FACTURA B" in n:
            out["tipo"], out["letra"] = "FCC", "B"
        elif "FACTURA" in n:
            out["tipo"], out["letra"] = "FCC", "A"

    m = re.search(
        r"(?:Punto\s+de\s+Venta|Pto\.?\s*Vta)\s*:?\s*(\d{1,5})"
        r".{0,80}?(?:Cpte|Comp(?:robante)?)\.?\s*N(?:ro|um|[ºo°])?\.?\s*:?\s*(\d{1,8})",
        texto,
        re.I | re.S,
    )
    if m:
        out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
    if "nro" not in out:
        m = re.search(r"N[°º]?:?\s*A\s+(\d{4,5})\s+(\d{5,8})", texto, re.I)
        if m:
            out["tipo"], out["letra"] = "FCC", "A"
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
    if "nro" not in out:
        m = re.search(r"Factura\s+N[ªºo°]?\s*0*(\d{1,5})\s+0*(\d{3,8})", texto, re.I)
        if m:
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
    if "nro" not in out:
        m = re.search(r"N[ºo°]\s*:?\s*0*(\d{1,5})\s*-\s*0*(\d{3,8})", texto[:800], re.I)
        if m:
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))
    if "nro" not in out:
        m = re.search(r"\b0*(\d{4,5})\s*[-–]\s*0*(\d{5,8})\b", texto[:1200])
        if m:
            out["pv"], out["nro"] = int(m.group(1)), int(m.group(2))

    for m in re.finditer(r"Raz[oó]n\s+Social\s*:?\s*([^\n]{3,80})", texto, re.I):
        cand = _safe_prov(m.group(1))
        if cand and not any(tok in cand for tok in CLIENTE_TOKENS):
            out["prov"] = cand
            break
    if "COLON" in n and "prov" not in out:
        out["prov"] = "COLON SA"
    if "ASTATEC" in n:
        out["prov"] = "ASTATEC SA"
    if "CENTRO MEDICO" in n and "prov" not in out:
        out["prov"] = "CENTRO MEDICO DE MAR DEL PLATA"
    if "30700667606" in re.sub(r"\D", "", texto):
        out["prov"] = "VEGA"
    if "I.N.S.S.J.P" in n or "INSSJP" in n:
        out["prov"] = "PAMI"
        m = re.search(r"N[ºo°]\s*(\d{6,12})", texto)
        if m and "nro" not in out:
            out["tipo"], out["letra"] = "FCC", "A"
            out["nro"] = int(m.group(1))
            out["pv"] = 0
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
    return re.sub(r"\s+", " ", INVALID.sub(" ", base)).strip() + ext


def parsear(path: Path) -> dict:
    info = _desde_nombre(path.name)
    texto = ""
    if not (info.get("nro") and info.get("tipo") and info.get("prov")):
        if path.suffix.lower() == ".pdf":
            texto = _texto_pdf(path)
            extra = _desde_texto(texto)
            for k, v in extra.items():
                if not info.get(k):
                    info[k] = v
    return {
        "path": path,
        "orig": path.name,
        "nuevo": armar_nombre(info, path.name),
        "info": info,
        "completo": bool(info.get("nro") and info.get("prov")),
    }


def unique_target(folder: Path, name: str, used: set[str], current: str | None = None) -> str:
    low = name.lower()
    if current and low == current.lower():
        used.add(low)
        return name
    if low not in used and not (folder / name).exists():
        used.add(low)
        return name
    stem, ext = Path(name).stem, Path(name).suffix
    i = 2
    while True:
        cand = f"{stem} ({i}){ext}"
        if cand.lower() not in used and not (folder / cand).exists():
            used.add(cand.lower())
            return cand
        i += 1


def procesar_mes(folder: Path) -> tuple[int, int, list[str]]:
    files = [
        p
        for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".pdf", ".jpeg", ".jpg", ".png"}
        and not SKIP_NAMES.search(p.name)
        and not RE_YA.match(p.name)
    ]
    used = {p.name.lower() for p in folder.iterdir() if p.is_file()}
    renamed = 0
    revisar = []
    for p in sorted(files, key=lambda x: x.name.lower()):
        row = parsear(p)
        dest = unique_target(folder, row["nuevo"], used, current=p.name)
        if not row["completo"]:
            revisar.append(f"{p.name} -> {dest}")
        if p.name.lower() == dest.lower():
            continue
        dst = folder / dest
        try:
            p.rename(dst)
            renamed += 1
            print(f"  REN {p.name[:55]} => {dest}")
        except Exception as e:
            print(f"  ERR {p.name}: {e}")
            revisar.append(f"ERR {p.name}: {e}")
    return renamed, len(files), revisar


def main() -> None:
    meses = sorted(
        p for p in ROOT.iterdir() if p.is_dir() and p.name not in SKIP_MESES
    )
    tot_ren = tot_files = 0
    all_rev: list[str] = []
    for mes in meses:
        print("====", mes.name)
        r, n, rev = procesar_mes(mes)
        tot_ren += r
        tot_files += n
        all_rev.extend(f"{mes.name}: {x}" for x in rev)
        print(f"   archivos={n} renombrados={r} revisar={len(rev)}")
    print("TOTAL archivos", tot_files, "renombrados", tot_ren, "revisar", len(all_rev))
    for x in all_rev:
        print("REVISAR", x)


if __name__ == "__main__":
    main()
