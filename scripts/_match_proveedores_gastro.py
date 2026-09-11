"""Match débitos extracto Gastro vs Tango Gastr.xlsx (proveedores)."""

from __future__ import annotations

import hashlib
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from procesador import (  # noqa: E402
    _corregir_filas_extracto_por_saldos,
    cargar_debitos_desde_extracto_df,
    cargar_facturas_proveedores_excel,
    enriquecer_df_extracto_formato_banco,
    exportar_match_proveedores_excel,
    filtrar_facturas_match_proveedores,
    matchear_debitos_con_facturas,
    procesar_extractos_bancarios_pdfs,
)

BANCOS = Path(
    r"\\TANGOSRV\Compartido\CLIENTES"
    r"\GASTROENTEROLOGIA Y ENDOSCOPIA DIGESTIVA MAR DEL PLATA S.A"
    r"\Balances\31-08-2026\Bancos\Cuenta Corriente Nº 067-0850315\2026"
)
CACHE = Path(r"C:\Users\recep\Desktop\Estudio Contable\_cache_extractos_gastro")
TANGO = Path(r"C:\Users\recep\Desktop\Gastr.xlsx")
OUT = Path(r"C:\Users\recep\Desktop\Match_Proveedores_Gastro.xlsx")


class _PdfUpload:
    def __init__(self, path: Path) -> None:
        self.name = path.name
        self._path = path

    def getvalue(self) -> bytes:
        return self._path.read_bytes()


def cache_path_for(pdf: Path) -> Path:
    h = hashlib.sha1(pdf.read_bytes()).hexdigest()[:16]
    return CACHE / f"{pdf.stem}_{h}.pkl"


def cargar_extractos() -> pd.DataFrame:
    CACHE.mkdir(parents=True, exist_ok=True)
    pdfs = [
        p
        for p in sorted(BANCOS.glob("*.pdf"))
        if "unificado" not in p.name.lower()
    ]
    frames: list[pd.DataFrame] = []
    for pdf in pdfs:
        cp = cache_path_for(pdf)
        if cp.exists():
            print(f"  cache {pdf.name}")
            frames.append(pd.read_pickle(cp))
            continue
        print(f"  OCR/parse {pdf.name}")
        df, _meta, errs = procesar_extractos_bancarios_pdfs([_PdfUpload(pdf)])
        if errs:
            for e in errs:
                print(f"    ERR {e}")
        df.to_pickle(cp)
        frames.append(df)
    if not frames:
        raise SystemExit("No hay extractos")
    ext = pd.concat(frames, ignore_index=True)
    return pd.DataFrame(_corregir_filas_extracto_por_saldos(ext.to_dict("records")))


def main() -> None:
    print("1) Extractos…")
    ext = cargar_extractos()
    print(f"   filas extracto {len(ext)} cols {list(ext.columns)[:12]}")
    print("2) Match vs Gastr.xlsx (solo 21101 Proveedores)...")
    ext_enr = enriquecer_df_extracto_formato_banco(ext)
    debitos = cargar_debitos_desde_extracto_df(ext_enr)
    facturas = cargar_facturas_proveedores_excel(TANGO)
    if "cuenta" in facturas.columns:
        facturas = facturas[
            facturas["cuenta"].astype(str).str.contains("21101", na=False)
        ].copy()
        facturas = facturas.reset_index(drop=True)
        facturas["factura_id"] = facturas.index.astype(int)
    facturas, excluidas_banco = filtrar_facturas_match_proveedores(facturas)
    resultado = matchear_debitos_con_facturas(debitos, facturas)
    if excluidas_banco is not None and not excluidas_banco.empty:
        resultado["facturas_excluidas_banco"] = excluidas_banco
    meta_pipe = {
        "n_debitos": len(debitos),
        "n_facturas": len(facturas),
        "n_excluidas_banco": len(excluidas_banco),
    }
    meta = {
        "cliente": "Gastroenterología y Endoscopia Digestiva Mar del Plata S.A.",
        "cuit": "30-71784781-0",
        "origen_extracto": "CC Banco Provincia 067-0850315 (sep-2025 a jul-2026)",
        "origen_facturas": str(TANGO) + " · cuenta 21101 Proveedores",
        **meta_pipe,
    }
    xlsx = exportar_match_proveedores_excel(resultado, meta)
    OUT.write_bytes(xlsx)
    cal = resultado.get("calzados")
    sin = resultado.get("pagos_sin_factura")
    imp = resultado.get("facturas_impagas")
    print("OK", OUT)
    print("calzados", 0 if cal is None else len(cal))
    print("pagos sin factura", 0 if sin is None else len(sin))
    print("facturas impagas", 0 if imp is None else len(imp))


if __name__ == "__main__":
    main()
