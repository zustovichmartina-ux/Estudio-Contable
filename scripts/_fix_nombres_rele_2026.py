# -*- coding: utf-8 -*-
from pathlib import Path

ROOT = Path(
    r"\\TANGOSRV\Compartido\CLIENTES\OFTALMOLOGIA RELE MAR DEL PLATA SRL"
    r"\Facturas\2026"
)

# dest actual -> nombre correcto (recuperado del nombre original)
FIXES = {
    "05-2026": [
        ("FCCA ALCON.jpeg", "FCCA20-252524 ALCON.jpeg"),
        ("FCCA ALCON (2).jpeg", "FCCA20-252945 ALCON.jpeg"),
        ("FCCA ALCON (3).jpeg", "FCCA20-253362 ALCON.jpeg"),
        ("FCCA ASTATEC SA.jpeg", "FCCA3-23068 ASTATEC SA.jpeg"),
        ("FCCA BIOMAT SRL.jpeg", "FCCA4-70704 BIOMAT SRL.jpeg"),
        ("FCCA11-21813 PROVEEDOR.pdf", "FCCA11-21813 BERGEON.pdf"),
        ("FCCA PROVEEDOR.jpeg", "FCCA30939 INTEGRAL PACK.jpeg"),
        ("FCCA PROVEEDOR (4).jpeg", "FCCA13-4268 OXFORD MARINO.jpeg"),
        ("FCCA PROVEEDOR (5).jpeg", "FCCA13-4313 OXFORD MARINO.jpeg"),
        ("FCCA PROVEEDOR (7).jpeg", "FCCA4-46229 MED SRL.jpeg"),
        ("FCCA PROVEEDOR (8).jpeg", "FCCA16-9516 LIMPA SA.jpeg"),
        ("FCCA PROVEEDOR (2).jpeg", "FCCA LANDA HOJA 2.jpeg"),
        ("FCCA PROVEEDOR (3).jpeg", "FCCA LANDA HOJA 1.jpeg"),
        ("FCCA PROVEEDOR (6).jpeg", "FCCA MATERIALES OBRA.jpeg"),
    ],
    "06-2026": [
        ("FCCA ALCON HOJA 1.jpeg", "FCCA20-255367 ALCON HOJA 1.jpeg"),
        ("FCCA ALCON HOJA 2.jpeg", "FCCA20-255367 ALCON HOJA 2.jpeg"),
        ("FCCA COLON SA.jpeg", "FCCA28-30715 COLON SA.jpeg"),
        ("FCCA PROVEEDOR.jpeg", "FCCA3-4072 DE LA SANTE.jpeg"),
        ("FCCA PROVEEDOR.pdf", "FCCA4-10779 SANITARIOS CASTELLI.pdf"),
        ("FCCA BIOMAT SRL.pdf", "FCCA4-71786 BIOMAT SRL.pdf"),
        ("FCCA BIOMAT SRL (2).pdf", "NDCA4-14865 BIOMAT SRL.pdf"),
        ("FCCA BIOMAT SRL (3).pdf", "NDCA4-14887 BIOMAT SRL.pdf"),
        ("FCCA BIOMAT SRL (4).pdf", "NDCA4-14888 BIOMAT SRL.pdf"),
        ("FCCC SOLIS SERRANO RODRIGO EDUARDO.pdf", "FCCC334 SOLIS SERRANO.pdf"),
        ("FCCC SOLIS SERRANO RODRIGO EDUARDO (2).pdf", "FCCC335 SOLIS SERRANO.pdf"),
        ("FCCA COLON SA (2).jpeg", "RETENCIONES COLON SA.jpeg"),
    ],
    "02-2026": [
        ("FCCA11-21113 PROVEEDOR.pdf", "FCCA11-21113 LANDA.pdf"),
        ("FCCA6-68018202 PROVEEDOR.pdf", "FCCA6-68018 ROSINOV.pdf"),
    ],
    "01-2026": [
        ("FCCA4-45550 PROVEEDOR.pdf", "FCCA4-45550 MED SRL.pdf"),
        ("FCCA3-21044 PROVEEDOR.pdf", "FCCA3-21044 QUALITY CLEAN.pdf"),
        ("FCCC2-7 SUAREZ FACUNDO OK.pdf", "FCCC2-7 SUAREZ FACUNDO.pdf"),
        ("NCCC2-1 SUAREZ NC.pdf", "NCCC2-1 SUAREZ FACUNDO.pdf"),
    ],
    "03-2026": [
        ("FCCA3-21315 PROVEEDOR.pdf", "FCCA3-21315 QUALITY CLEAN.pdf"),
        ("FCCA2-8789 COLON SA.pdf", "FCCA2-8789 MASTER CALCOMANIAS.pdf"),
    ],
    "04-2026": [
        ("FCCA3-21742 PROVEEDOR.pdf", "FCCA3-21742 QUALITY CLEAN.pdf"),
        ("FCCA6-68018 PROVEEDOR.pdf", "FCCA6-68018 ROSINOV.pdf"),
        ("FCCA4-5419 PROVEEDOR.pdf", "FCCA4-5419 VETRA (dup).pdf"),
        ("FCCA4-5419 PROVEEDOR (2).pdf", "FCCA4-5419 VETRA (dup 2).pdf"),
        ("FCCA4-5419 PROVEEDOR (3).pdf", "FCCA4-5419 VETRA (dup 3).pdf"),
    ],
}


def ren(folder: Path, old: str, new: str) -> None:
    src, dst = folder / old, folder / new
    if not src.exists():
        print("FALTA", folder.name, old)
        return
    if src.name == new:
        return
    if dst.exists():
        print("CHOQUE", folder.name, new)
        return
    src.rename(dst)
    print("OK", folder.name, new)


def main() -> None:
    for mes, pares in FIXES.items():
        folder = ROOT / mes
        for old, new in pares:
            if not new:
                continue
            src, dst = folder / old, folder / new
            if not src.exists():
                print("FALTA", folder.name, old)
                continue
            if dst.exists():
                stem, ext = Path(new).stem, Path(new).suffix
                i = 2
                while (folder / f"{stem} ({i}){ext}").exists():
                    i += 1
                dst = folder / f"{stem} ({i}){ext}"
            src.rename(dst)
            print("OK", folder.name, dst.name)


if __name__ == "__main__":
    main()
