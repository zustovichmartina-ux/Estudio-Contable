# -*- coding: utf-8 -*-
"""Créditos nominados y cheques — CC 067-0850315 Gastro MDP (cierre 31-08-2026)."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from excel_formato_estudio import guardar_informe_excel

CACHE = ROOT / "_cache_extractos_gastro"
OUT_DIR = Path(
    r"\\TANGOSRV\Compartido\CLIENTES"
    r"\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
    r"\Balances\Cierre 31-08-2026\Bancos\Cuenta Corriente Nº 067-0850315"
)
OUT_XLSX = OUT_DIR / "Creditos_y_Cheques_CC_067-0850315.xlsx"
OUT_DESKTOP = Path(r"C:\Users\recep\Desktop") / "Creditos_y_Cheques_CC_067-0850315.xlsx"

CUIT_ORIGEN = {
    "30541190518": "CENTRO MEDICO DE MAR DEL PLATA",
    "30529900747": "LABORATORIO DOMINGUEZ S.A.",
    "27113020674": "CIGARRETA MARIA CRISTINA",
    "27290677543": "VALLES MARIA CECILIA",
    "27329200596": "KATUNSKIS LUCIA",
    "27258988634": "VALERIO LILIANA PATRICIA",
    "20082860607": "RINALDI JORGE ALFREDO",
    "27401438683": "GOMEZ SAMANTA AGUSTINA",
    "23331029084": "SANTOS FERNANDA MARIEL",
    "20367811057": "DEMITRE RODRIGO NAHUEL",
    "20142037395": "RONCALLO LUIS FERNANDO",
    "20179821789": "RODRIGUEZ ROBERTO DANIEL",
    "27167792451": "TUSAR SUSANA BEATRIZ",
    "20084834522": "PRESSELLO JORGE LUIS",
    "20386281018": "LABORDE AMADOR JUAN C",
    "20280166325": "SABUDA RODRIGO JAVIER",
    "20382911424": "POLAKO LEANDRO",
    "20165045891": "FOGOLINI HORACIO OMAR",
    "27130717573": "RUSSO NORA AMELIA",
    "20143754740": "BERCERUELO HECTOR ABEL",
    "20263461674": "ESPENDE PABLO DANIEL",
}

RE_CUIT = re.compile(r"(?<!\d)(\d{11})(?!\d)")
RE_DE = re.compile(
    r"\bde\s+([a-záéíóúüñ][a-záéíóúüñ ,\.;']{3,80}?)(?=\s*(?:varios|/|- var|var\b|transf|cuit|fac\b|\d{11}|$))",
    re.I,
)
RE_CHQ_NRO = re.compile(r"\b(\d{2,8})\s+(?:cheque|deposito ch|dep[oó]sito ch)", re.I)

NOISE_CREDITO = (
    "impuesto ley",
    "retencion arba",
    "comision",
    "debito automatico",
    "pago haberes",
    "pago de haberes",
    "pago de servicios",
    "iva",
    "debito transf",
    "estructura jur",
    "acuerdo de giro",
    "los adquiridos por endoso",
)


def _norm(s: str) -> str:
    t = (s or "").upper()
    for a, b in (("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ñ", "N")):
        t = t.replace(a, b)
    t = t.replace("\n", " ").replace(",", " ")
    return re.sub(r"\s+", " ", t).strip()


def _blob(row: pd.Series) -> str:
    return " ".join(
        str(row.get(c) or "")
        for c in ("Descripcion", "Detalle", "Concepto unificado", "Clasificacion")
    )


def _es_ruido(blob: str) -> bool:
    n = _norm(blob).lower()
    if n.startswith("anul"):
        return False
    return any(x in n for x in NOISE_CREDITO)


def _es_cheque(blob: str) -> bool:
    n = _norm(blob)
    return bool(
        re.search(r"\bCHEQUE\b|\bECHEQ\b|\bDEPOSITO CH\b|\bDEPÓSITO CH\b", n)
        or "COMP ELECT CAM LOCAL" in n
        or "NODO DE RECAUDACIONES" in n
    )


def _es_ingreso(blob: str, credito: float) -> bool:
    if credito <= 0.009:
        return False
    n = _norm(blob)
    if _es_ruido(blob) and "ANUL" not in n:
        return False
    if re.search(r"\bA [A-Z]", n) and "RECIBID" not in n:
        # "A FLORRENCIA..." / "A OMNIASALUD" = salida mal marcada como crédito
        if re.search(r"\bA (FLORENCIA|OMNIASALUD|LISTORTI|LOZZI|MEDRANO)\b", n):
            return False
    keys = (
        "RECIBID",
        "TRANSF MINORISTA",
        "TRANSFERENCIA RECIBIDA",
        "TRANSF RECIBIDA",
        "PAGO A PROVEEDORES",
        "PAGO HONORARIOS",
        "CHEQUE",
        "DEPOSITO CH",
        "RESCATE",
        "SUPERFONDO",
        "MOBILE BANKING",
        "DE PABLO",
        "SIN DESCRIPC",
        "CVU",
    )
    return any(k in n for k in keys) or n.startswith("DE ")


def _origen(blob: str, desc: str) -> tuple[str, str]:
    """(origen, cuit)."""
    n = _norm(blob)
    digits = re.sub(r"\D", "", blob)
    cuit = ""
    for i in range(0, max(0, len(digits) - 10)):
        cand = digits[i : i + 11]
        if cand in CUIT_ORIGEN:
            cuit = cand
            break
    if not cuit:
        m = RE_CUIT.search(blob.replace("\n", " "))
        if m:
            cuit = m.group(1)

    if cuit and cuit in CUIT_ORIGEN:
        return CUIT_ORIGEN[cuit], cuit
    if "CENTRO MED" in n:
        return "CENTRO MEDICO DE MAR DEL PLATA", cuit or "30541190518"
    if "LABORATORIO DOMINGUEZ" in n or "LABORATORIO DOMINGUE" in n:
        return "LABORATORIO DOMINGUEZ S.A.", cuit or "30529900747"
    if "SUPERFONDO" in n or "RESCATE" in n:
        return "RESCATE SUPERFONDO RENTA", cuit
    if _es_cheque(blob):
        return "CHEQUE (sin librador en extracto)", cuit

    m = RE_DE.search((blob or "").replace("\n", " "))
    if m:
        nombre = re.sub(r"\s+", " ", m.group(1)).strip(" ,.;/-")
        nombre = re.sub(r"\b(var|varios|transf|hon|fac)\b", "", nombre, flags=re.I).strip(" ,.;/-")
        if len(nombre) >= 4:
            return nombre.upper(), cuit

    desc_n = _norm(desc)
    if desc_n.startswith("DE "):
        return desc_n[3:].split(" [")[0].strip(), cuit
    if "SIN DESCRIPC" in n:
        return "SIN DENOMINACION EN EXTRACTO", cuit
    return "SIN DENOMINACION EN EXTRACTO", cuit


def _tipo(blob: str, desc: str) -> str:
    n = _norm(blob)
    if _es_cheque(blob):
        return "Cheque recibido"
    if "SUPERFONDO" in n or "RESCATE" in n:
        return "Rescate FCI"
    if "HONORARIO" in n:
        return "Pago honorarios"
    if "PAGO A PROVEEDORES" in n:
        return "Pago recibido"
    if "RECIBID" in n or "TRANSF" in n or "MOBILE" in n or _norm(desc).startswith("DE "):
        return "Transferencia recibida"
    return "Credito"


def _nro_cheque(blob: str) -> str:
    m = RE_CHQ_NRO.search(blob.replace("\n", " "))
    return m.group(1) if m else ""


def _fmt_fecha(v) -> str:
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    s = str(v or "").strip()
    return s


def cargar() -> pd.DataFrame:
    frames = [pd.read_pickle(p) for p in sorted(CACHE.glob("*.pkl"))]
    if not frames:
        raise SystemExit("No hay cache de extractos Gastro")
    df = pd.concat(frames, ignore_index=True)
    df["Credito"] = pd.to_numeric(df.get("Credito"), errors="coerce")
    df["Debito"] = pd.to_numeric(df.get("Debito"), errors="coerce")
    df["Importe"] = pd.to_numeric(df.get("Importe"), errors="coerce")
    df["_fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
    # Cierre 31-08-2026: sep-2025 a ago-2026 (no hay extracto 08-2026)
    df = df[(df["_fecha"] >= "2025-09-01") & (df["_fecha"] <= "2026-08-31")].copy()
    # no duplicar el unificado
    if "Archivo origen" in df.columns:
        df = df[~df["Archivo origen"].astype(str).str.contains("unificado", case=False, na=False)]
    return df


def extraer_creditos(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        cred = float(r["Credito"] or 0) if pd.notna(r["Credito"]) else 0.0
        blob = _blob(r)
        if not _es_ingreso(blob, cred):
            continue
        origen, cuit = _origen(blob, str(r.get("Descripcion") or ""))
        rows.append(
            {
                "Fecha": _fmt_fecha(r.get("Fecha")),
                "Mes": r.get("Mes etiqueta") or r.get("Mes") or "",
                "Tipo": _tipo(blob, str(r.get("Descripcion") or "")),
                "Origen / denominacion": origen,
                "CUIT": cuit,
                "Nro cheque": _nro_cheque(blob) if _es_cheque(blob) else "",
                "Descripcion": str(r.get("Descripcion") or "").replace("\n", " ").strip(),
                "Detalle": str(r.get("Detalle") or "").replace("\n", " ").strip(),
                "Importe": round(cred, 2),
                "Archivo": r.get("Archivo origen") or "",
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["_ord"] = pd.to_datetime(out["Fecha"], dayfirst=True, errors="coerce")
    return out.sort_values(["_ord", "Origen / denominacion"]).drop(columns=["_ord"]).reset_index(drop=True)


def extraer_cheques(df: pd.DataFrame, creditos: pd.DataFrame) -> pd.DataFrame:
    """Cheques recibidos (crédito) + e-cheqs debitados (canje interno)."""
    rec = creditos[creditos["Tipo"] == "Cheque recibido"].copy()
    rec.insert(2, "Sentido", "Credito (depositado)")

    rows = []
    for _, r in df.iterrows():
        blob = _blob(r)
        n = _norm(blob)
        if "ECHEQ" not in n:
            continue
        if "ENDOSO" in n or "ACUERDO DE GIRO" in n:
            continue
        deb = float(r["Debito"] or 0) if pd.notna(r["Debito"]) else 0.0
        cred = float(r["Credito"] or 0) if pd.notna(r["Credito"]) else 0.0
        if deb <= 0.009 and cred <= 0.009:
            continue
        sentido = "Debito (e-cheq / canje)" if deb > 0.009 else "Credito (depositado)"
        monto = deb if deb > 0.009 else cred
        rows.append(
            {
                "Fecha": _fmt_fecha(r.get("Fecha")),
                "Mes": r.get("Mes etiqueta") or r.get("Mes") or "",
                "Sentido": sentido,
                "Tipo": "E-cheq",
                "Origen / denominacion": "E-CHEQ (sin librador en extracto)",
                "CUIT": "",
                "Nro cheque": str(r.get("Comprobante") or "").strip(),
                "Descripcion": str(r.get("Descripcion") or "").replace("\n", " ").strip(),
                "Detalle": str(r.get("Detalle") or "").replace("\n", " ").strip(),
                "Importe": round(monto, 2),
                "Archivo": r.get("Archivo origen") or "",
            }
        )
    ech = pd.DataFrame(rows)
    cols = [
        "Fecha",
        "Mes",
        "Sentido",
        "Tipo",
        "Origen / denominacion",
        "CUIT",
        "Nro cheque",
        "Descripcion",
        "Detalle",
        "Importe",
        "Archivo",
    ]
    out = pd.concat([rec.reindex(columns=cols), ech.reindex(columns=cols)], ignore_index=True)
    if out.empty:
        return out
    out["_ord"] = pd.to_datetime(out["Fecha"], dayfirst=True, errors="coerce")
    return out.sort_values(["_ord", "Sentido"]).drop(columns=["_ord"]).reset_index(drop=True)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    df = cargar()
    print(f"Movimientos cache: {len(df)}")
    cred = extraer_creditos(df)
    chq = extraer_cheques(df, cred)
    print(f"Creditos: {len(cred)}  total={cred['Importe'].sum():,.2f}" if not cred.empty else "Creditos: 0")
    print(f"Cheques: {len(chq)}  total={chq['Importe'].sum():,.2f}" if not chq.empty else "Cheques: 0")

    por_origen = (
        cred.groupby("Origen / denominacion", dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
        .sort_values("Importe", ascending=False)
    )
    por_tipo = (
        cred.groupby("Tipo", dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
        .sort_values("Importe", ascending=False)
    )
    chq_mes = (
        chq.groupby(["Mes", "Sentido"], dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
        .sort_values("Mes")
        if not chq.empty
        else pd.DataFrame()
    )

    kpis = [
        ("Creditos extraidos", float(len(cred))),
        ("Total creditos", float(cred["Importe"].sum()) if not cred.empty else 0.0),
        ("Cheques / e-cheqs", float(len(chq))),
        (
            "Total cheques recibidos",
            float(chq.loc[chq["Sentido"].astype(str).str.contains("Credito", na=False), "Importe"].sum())
            if not chq.empty
            else 0.0,
        ),
        (
            "Total e-cheqs debitados",
            float(chq.loc[chq["Sentido"].astype(str).str.contains("Debito", na=False), "Importe"].sum())
            if not chq.empty
            else 0.0,
        ),
    ]

    guardar_informe_excel(
        OUT_XLSX,
        titulo="Creditos nominados y cheques",
        subtitulo=(
            "GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A. — "
            "Cta. Cte. 067-0850315"
        ),
        periodo="Cierre 31/08/2026 (extractos sep-2025 a jul-2026; sin 08-2026 en carpeta)",
        kpis=kpis,
        resumenes=[
            ("Creditos por origen / denominacion", por_origen),
            ("Creditos por tipo", por_tipo),
            ("Cheques por mes", chq_mes),
        ],
        detalle=cred,
        hoja_detalle="Creditos",
        hojas_adicionales=[("Cheques", chq)],
        col_moneda=["Importe"],
        col_fecha=["Fecha"],
        total_col="Importe",
    )
    try:
        import shutil

        shutil.copy2(OUT_XLSX, OUT_DESKTOP)
    except OSError:
        guardar_informe_excel(
            OUT_DESKTOP,
            titulo="Creditos nominados y cheques",
            subtitulo=(
                "GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A. — "
                "Cta. Cte. 067-0850315"
            ),
            periodo="Cierre 31/08/2026 (extractos sep-2025 a jul-2026; sin 08-2026 en carpeta)",
            kpis=kpis,
            resumenes=[
                ("Creditos por origen / denominacion", por_origen),
                ("Creditos por tipo", por_tipo),
                ("Cheques por mes", chq_mes),
            ],
            detalle=cred,
            hoja_detalle="Creditos",
            hojas_adicionales=[("Cheques", chq)],
            col_moneda=["Importe"],
            col_fecha=["Fecha"],
            total_col="Importe",
        )
    print(f"Excel: {OUT_XLSX}")
    print(f"Copia: {OUT_DESKTOP}")
    if not por_origen.empty:
        print("\nTop origenes:")
        print(por_origen.head(12).to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
