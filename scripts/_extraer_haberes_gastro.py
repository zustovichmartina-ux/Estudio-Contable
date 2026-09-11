#!/usr/bin/env python3
"""Extrae débitos de haberes y transferencias nominadas — Gastro MDP (Provincia)."""

from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

import pandas as pd

from excel_formato_estudio import guardar_informe_excel
from procesador import procesar_extractos_bancarios_pdfs

BANCOS_PARENT = Path(
    r"\\TANGOSRV\Compartido\CLIENTES"
    r"\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
    r"\Balances\31-08-2026\Bancos"
)

CACHE_DIR = Path(r"C:\Users\recep\Desktop\Estudio Contable\_cache_extractos_gastro")
OUT_XLSX = Path(
    r"C:\Users\recep\Desktop\Debitos_Haberes_y_Transferencias_Gastro_sep25_jul26.xlsx"
)

DESTINATARIOS = [
    ("LOZZI", "ADRIANA", "Lozzi Adriana"),
    ("LISTORTI", "ANTONELLA", "Listorti Antonella"),
    ("MEDRANO", "CARLOS", "Medrano Carlos"),
    ("ARCE", "ENRIQUE", "Arce Enrique"),
    ("PASTORINO", "MARTIN", "Pastorino Martin"),
]

# CUITs observados en extractos (recuperan transferencias sin nombre / OCR cortado)
CUIT_DESTINATARIO = {
    "23160237384": "Listorti Antonella",
    "23328483424": "Lozzi Adriana",
    "20955912102": "Medrano Carlos",
    "20286081283": "Pastorino Martin",
    "20302079030": "Arce Enrique",
}

MAX_IMPORTE_HABERES = 5_000_000.0
MAX_IMPORTE_TRANSF = 20_000_000.0


class _PdfUpload:
    def __init__(self, path: Path) -> None:
        self.name = path.name
        self._path = path

    def getvalue(self) -> bytes:
        return self._path.read_bytes()


def _norm(s: str) -> str:
    t = (s or "").upper()
    for a, b in (
        ("Á", "A"),
        ("É", "E"),
        ("Í", "I"),
        ("Ó", "O"),
        ("Ú", "U"),
        ("Ñ", "N"),
        (",", " "),
    ):
        t = t.replace(a, b)
    return re.sub(r"\s+", " ", t).strip()


def encontrar_carpeta() -> Path:
    if not BANCOS_PARENT.is_dir():
        raise FileNotFoundError(f"No existe: {BANCOS_PARENT}")
    for p in BANCOS_PARENT.iterdir():
        if p.is_dir() and "0850315" in p.name:
            cand = p / "2026"
            return cand if cand.is_dir() else p
    raise FileNotFoundError("No se encontró la cuenta 067-0850315")


def match_destinatario(*textos: str) -> str | None:
    blob = _norm(" ".join(textos))
    for ape, nom, label in DESTINATARIOS:
        if ape in blob and nom in blob:
            return label
    # apellido inequívoco en transferencias "A nombre apellido"
    if "LOZZI" in blob:
        return "Lozzi Adriana"
    if "LISTORTI" in blob:
        return "Listorti Antonella"
    if "MEDRANO" in blob:
        return "Medrano Carlos"
    if "PASTORINO" in blob:
        return "Pastorino Martin"
    if "ARCE" in blob and ("ENRIQUE" in blob or "ENRIQ" in blob):
        return "Arce Enrique"
    return None


def es_haberes(*textos: str) -> bool:
    blob = _norm(" ".join(textos))
    if "HABER" not in blob and "SUELDO" not in blob:
        return False
    return True


def es_anulacion(*textos: str) -> bool:
    blob = _norm(" ".join(textos))
    return blob.startswith("ANUL") or " ANUL " in f" {blob} " or "ANULACION" in blob


def cache_path_for(pdf: Path) -> Path:
    h = hashlib.sha1(pdf.read_bytes()).hexdigest()[:16]
    return CACHE_DIR / f"{pdf.stem}_{h}.pkl"


def procesar_con_cache(pdfs: list[Path]) -> tuple[pd.DataFrame, list[dict]]:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    frames: list[pd.DataFrame] = []
    errores: list[dict] = []
    pendientes: list[Path] = []

    for pdf in pdfs:
        cp = cache_path_for(pdf)
        if cp.exists():
            print(f"  cache HIT {pdf.name}")
            frames.append(pd.read_pickle(cp))
        else:
            pendientes.append(pdf)

    if pendientes:
        print(f"Procesando {len(pendientes)} PDF(s) (puede OCR)...")
        for pdf in pendientes:
            print(f"  -> {pdf.name} ({pdf.stat().st_size // 1024} KB)")
            df, _meta, errs = procesar_extractos_bancarios_pdfs([_PdfUpload(pdf)])
            errores.extend(errs or [])
            if errs:
                for e in errs:
                    print(f"     ERROR: {e}")
            cp = cache_path_for(pdf)
            df.to_pickle(cp)
            print(f"     filas={len(df)} -> cache {cp.name}")
            frames.append(df)

    if not frames:
        return pd.DataFrame(), errores
    out = pd.concat(frames, ignore_index=True)
    return out, errores


def filtrar(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict] = []
    for _, r in df.iterrows():
        desc = str(r.get("Descripcion") or "")
        det = str(r.get("Detalle") or "")
        clas = str(r.get("Clasificacion") or "")
        textos = (desc, det, clas)

        haber = es_haberes(*textos) or "haberes" in clas.lower()
        dest = match_destinatario(*textos)
        if not haber and not dest:
            continue

        # Solo salidas (débitos). Anulaciones de haberes suelen ser crédito: marcar igual.
        try:
            importe = float(r.get("Importe") or 0)
        except (TypeError, ValueError):
            importe = 0.0
        debito = r.get("Debito")
        try:
            debito_f = float(debito) if debito not in (None, "") else None
        except (TypeError, ValueError):
            debito_f = None

        anul = es_anulacion(*textos)
        # Preferir columna Débito si existe
        if debito_f is not None and debito_f > 0:
            monto = debito_f
            sentido = "Debito"
        elif importe < 0:
            monto = abs(importe)
            sentido = "Debito"
        elif importe > 0 and (haber or dest):
            # En Provincia a veces el importe del débito viene positivo en Importe
            # y Credito vacío. Si es haberes/transf a persona, tomar como débito
            # salvo anulación explícita.
            if anul:
                monto = abs(importe)
                sentido = "Credito (anulacion)"
            else:
                credito = r.get("Credito")
                try:
                    credito_f = float(credito) if credito not in (None, "") else 0.0
                except (TypeError, ValueError):
                    credito_f = 0.0
                if credito_f and credito_f > 0:
                    monto = credito_f
                    sentido = "Credito"
                else:
                    monto = abs(importe)
                    sentido = "Debito"
        else:
            continue

        if dest and haber:
            categoria = "Haberes"
            destinatario = dest
            tipo = f"Debito haberes / {dest}"
        elif dest:
            categoria = "Transferencia"
            destinatario = dest
            tipo = f"Transferencia — {dest}"
        else:
            categoria = "Haberes"
            destinatario = ""
            tipo = "Debito haberes" + (" (anulacion)" if anul else "")

        rows.append(
            {
                "Fecha": r.get("Fecha"),
                "Mes": r.get("Mes") or "",
                "Categoria": categoria,
                "Tipo": tipo,
                "Destinatario": destinatario,
                "Descripcion": desc,
                "Detalle": det,
                "Importe": round(float(monto), 2),
                "Sentido": sentido,
                "Clasificacion banco": clas,
                "Archivo": r.get("Archivo origen") or "",
            }
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    # ordenar por fecha
    out["_ord"] = pd.to_datetime(out["Fecha"], dayfirst=True, errors="coerce")
    out = out.sort_values(["_ord", "Categoria", "Destinatario"]).drop(columns=["_ord"])
    return out.reset_index(drop=True)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    base = encontrar_carpeta()
    pdfs = sorted({p.resolve(): p for p in list(base.glob("*.pdf")) + list(base.glob("*.PDF"))}.values(), key=lambda p: p.name)
    print(f"Carpeta: {base}")
    print(f"PDFs: {len(pdfs)}")
    for p in pdfs:
        print(f"  - {p.name}")

    df_all, errores = procesar_con_cache(pdfs)
    print(f"\nMovimientos totales parseados: {len(df_all)}")
    if errores:
        print(f"Errores: {errores}")

    sel = filtrar(df_all)
    print(f"Seleccionados (haberes + nominados): {len(sel)}")

    if sel.empty:
        print("No se encontraron movimientos. Abortando Excel.")
        return 1

    # Resumen
    resumen_hab = (
        sel[sel["Categoria"] == "Haberes"]
        .groupby("Mes", dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
    )
    resumen_hab.insert(0, "Concepto", "Haberes")

    resumen_tr = (
        sel[sel["Categoria"] == "Transferencia"]
        .groupby(["Destinatario", "Mes"], dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
        .rename(columns={"Destinatario": "Concepto"})
    )

    por_persona = (
        sel.assign(
            Concepto=lambda d: d.apply(
                lambda r: r["Destinatario"] if r["Categoria"] == "Transferencia" else "Haberes",
                axis=1,
            )
        )
        .groupby("Concepto", dropna=False)["Importe"]
        .agg(Cantidad="count", Importe="sum")
        .reset_index()
        .sort_values("Concepto")
    )

    kpis = [
        ("PDFs procesados", float(len(pdfs))),
        ("Movimientos incluidos", float(len(sel))),
        (
            "Total débitos haberes",
            float(sel.loc[sel["Categoria"] == "Haberes", "Importe"].sum()),
        ),
        (
            "Total transferencias nominadas",
            float(sel.loc[sel["Categoria"] == "Transferencia", "Importe"].sum()),
        ),
    ]
    for _, row in por_persona.iterrows():
        kpis.append((f"Total {row['Concepto']}", float(row["Importe"])))

    detalle = sel.copy()
    # Fechas como texto DD/MM/YYYY si vienen datetime
    if "Fecha" in detalle.columns:
        detalle["Fecha"] = detalle["Fecha"].map(
            lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else str(x or "")
        )

    guardar_informe_excel(
        OUT_XLSX,
        titulo="Débitos de haberes y transferencias nominadas",
        subtitulo=(
            "GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A. — "
            "Cta. Cte. 067-0850315 (Banco Provincia)"
        ),
        periodo="Sep/2025 – Jul/2026 (11 extractos en carpeta; sin 08-2025)",
        kpis=kpis,
        resumenes=[
            ("Totales por concepto", por_persona),
            ("Haberes por mes", resumen_hab),
            ("Transferencias por persona/mes", resumen_tr),
        ],
        detalle=detalle,
        hoja_detalle="Movimientos",
        col_moneda=["Importe"],
        col_fecha=["Fecha"],
        total_col="Importe",
    )
    print(f"\nExcel: {OUT_XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
