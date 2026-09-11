from pathlib import Path
import re
import pdfplumber

folder = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026\07-2026"
)
keys = re.compile(
    r"percep|per\.?\s*rg|iibb|i\.i\.b\.b|ingresos\s+brutos|rg\s*5329|rg\s*3337|perc\.",
    re.I,
)
for p in sorted(folder.glob("*.pdf")):
    with pdfplumber.open(p) as pdf:
        t = "\n".join((pg.extract_text() or "") for pg in pdf.pages)
    hits = [ln for ln in t.splitlines() if keys.search(ln)]
    if hits:
        print("=" * 70)
        print(p.name)
        for ln in hits:
            print(" |", ln[:200])
