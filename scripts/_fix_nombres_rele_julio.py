from pathlib import Path
import pdfplumber

folder = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026\07-2026"
)

fixes = {
    "FCCA GSJ SA.jpeg": "FCCA60-70515 GSJ SA.jpeg",
    "FCCA OACI.jpeg": "FCCA29-51941 OACI SA.jpeg",
    "FCCA RABBIONE.jpeg": "FCCA21-29406 RABBIONE.jpeg",
    "FCCA VETRA.jpeg": "FCCA4-5559 VETRA.jpeg",
    "FCCA4-5549 VETRA F.pdf": "FCCA4-5549 VETRA.pdf",
    "FCCA3-1257193 PAYWAY FC.pdf": "FCCA3-1257193 PAYWAY.pdf",
    "FCCA4-46610 MED LAMPARA FA.pdf": "FCCA4-46610 MED SRL.pdf",
}

for old, new in fixes.items():
    src = folder / old
    dst = folder / new
    if not src.exists():
        print("FALTA", old)
        continue
    if dst.exists() and dst.resolve() != src.resolve():
        print("CHOQUE", old, "->", new)
        continue
    src.rename(dst)
    print("OK", new)

via = folder / "FCCA8349-2579 VIA CARGO.pdf"
soda = folder / "FCCA14-312328 MAR DEL PLATA SODA.pdf"
for p in (via, soda):
    print("=" * 40, p.name)
    try:
        with pdfplumber.open(p) as pdf:
            t = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2])
        print(t[:1200])
    except Exception as e:
        print(e)
