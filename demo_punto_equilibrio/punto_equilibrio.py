"""Punto de equilibrio diario de una cafetería a partir del export de ventas.

Uso:
    python punto_equilibrio.py                      # usa datos_ejemplo/
    python punto_equilibrio.py --ventas export.csv  # export real del sistema

Lee:
  - el export de ventas (detecta solas las columnas de fecha, artículo, cantidad e importe)
  - gastos_fijos.csv, costos_productos.csv y parametros.csv (los carga el estudio)

Escribe el informe en pantalla y el aviso para el dueño en avisos/aviso_AAAA-MM-DD.txt.
"""
from __future__ import annotations

import argparse
import csv
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

AQUI = Path(__file__).parent
COSTO_POR_DEFECTO = 0.32  # artículos sin costo de receta cargado
MARGEN_MINIMO = 0.55      # por debajo se marca para revisar precio
DIAS = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]


# ---------------------------------------------------------------- lectura
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return s.strip().lower()


def _num(s: str) -> float:
    """Acepta '1.234.567,50', '1234567.50', '$ 3.800'."""
    s = re.sub(r"[^\d,.\-]", "", str(s))
    if re.search(r",\d{1,2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    return float(s) if s not in ("", "-", ".") else 0.0


def _fecha(s: str) -> date | None:
    s = s.strip()[:10]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _leer_csv(ruta: Path) -> list[dict]:
    texto = ruta.read_text(encoding="utf-8-sig")
    sep = ";" if texto.splitlines()[0].count(";") >= texto.splitlines()[0].count(",") else ","
    return list(csv.DictReader(texto.splitlines(), delimiter=sep))


@dataclass
class Venta:
    fecha: date
    articulo: str
    cantidad: float
    importe: float


def leer_ventas(ruta: Path) -> list[Venta]:
    filas = _leer_csv(ruta)
    if not filas:
        raise SystemExit(f"{ruta.name}: el archivo no tiene ventas.")
    cols = {_norm(c): c for c in filas[0]}

    def buscar(*claves: str) -> str | None:
        for clave in claves:
            for n, original in cols.items():
                if clave in n:
                    return original
        return None

    c_fecha = buscar("fecha")
    c_art = buscar("articulo", "producto", "descripcion", "item")
    c_cant = buscar("cantidad", "cant")
    c_imp = buscar("importe", "total", "monto")
    if not c_fecha or not c_imp:
        raise SystemExit(f"{ruta.name}: no encontré columnas de fecha e importe. Columnas: {list(filas[0])}")

    ventas = []
    for f in filas:
        d = _fecha(f[c_fecha])
        if d is None:
            continue
        ventas.append(Venta(d, f[c_art].strip() if c_art else "Venta",
                            _num(f[c_cant]) if c_cant else 1, _num(f[c_imp])))
    return ventas


# ---------------------------------------------------------------- cálculo
@dataclass
class Producto:
    articulo: str
    unidades: float = 0
    venta: float = 0
    costo: float = 0
    costo_estimado: bool = False

    @property
    def margen(self) -> float:
        return 1 - self.costo / self.venta if self.venta else 0


@dataclass
class Resultado:
    ventas_mes: float
    por_dia: dict[date, float]
    fijos: float
    insumos_pct: float
    variables_pct: float
    margen_contribucion: float
    pe_mes: float
    pe_dia: float
    dias_apertura: int
    productos: list[Producto] = field(default_factory=list)

    @property
    def ultimo_dia(self) -> date:
        return max(self.por_dia)

    @property
    def dias_cargados(self) -> int:
        return len(self.por_dia)

    @property
    def proyeccion(self) -> float:
        return self.ventas_mes / self.dias_cargados * self.dias_apertura

    @property
    def falta(self) -> float:
        return self.pe_mes - self.ventas_mes

    @property
    def dias_restantes(self) -> int:
        return max(0, self.dias_apertura - self.dias_cargados)

    @property
    def margen_seguridad(self) -> float:
        return (self.proyeccion - self.pe_mes) / self.proyeccion

    @property
    def resultado_proyectado(self) -> float:
        return self.proyeccion * self.margen_contribucion - self.fijos


def calcular(ventas: list[Venta], fijos: float, costos: dict[str, float],
             comision_pct: float, iibb_pct: float, dias_apertura: int) -> Resultado:
    por_dia: dict[date, float] = defaultdict(float)
    prods: dict[str, Producto] = {}
    for v in ventas:
        por_dia[v.fecha] += v.importe
        p = prods.setdefault(v.articulo, Producto(v.articulo))
        p.unidades += v.cantidad
        p.venta += v.importe
        costo_u = costos.get(_norm(v.articulo))
        if costo_u is None:
            p.costo += v.importe * COSTO_POR_DEFECTO
            p.costo_estimado = True
        else:
            p.costo += costo_u * v.cantidad

    ventas_mes = sum(por_dia.values())
    insumos = sum(p.costo for p in prods.values()) / ventas_mes
    variables = insumos + comision_pct / 100 + iibb_pct / 100
    mc = 1 - variables
    pe = fijos / mc
    return Resultado(ventas_mes, dict(sorted(por_dia.items())), fijos, insumos, variables, mc,
                     pe, pe / dias_apertura, dias_apertura,
                     sorted(prods.values(), key=lambda p: p.venta * (1 - (variables - insumos)) - p.costo, reverse=True))


# ---------------------------------------------------------------- salida
def pesos(n: float) -> str:
    return "$ " + f"{n:,.0f}".replace(",", ".")


def corto(n: float) -> str:
    a, s = abs(n), "-" if n < 0 else ""
    if a >= 1e6:
        return f"{s}$ {a / 1e6:,.1f} M".replace(".", ",")
    if a >= 1e3:
        return f"{s}$ {a / 1e3:,.0f} mil"
    return f"{s}{pesos(a)}"


def pct(x: float) -> str:
    return f"{x * 100:.1f} %".replace(".", ",")


def fecha_corta(d: date) -> str:
    return f"{DIAS[d.weekday()]} {d.day}/{d.month}"


def aviso(r: Resultado) -> str:
    d = r.ultimo_dia
    venta = r.por_dia[d]
    lineas = [f"Buen día. Resumen de ayer, {fecha_corta(d)}:",
              f"Vendiste {corto(venta)}. Para cubrir los gastos necesitás {corto(r.pe_dia)} por día. "
              + ("Día cubierto." if venta >= r.pe_dia else f"Faltaron {corto(r.pe_dia - venta)}."),
              "",
              f"En el mes llevás {corto(r.ventas_mes)} de {corto(r.pe_mes)} necesarios."]
    if r.falta <= 0:
        lineas.append("Ya cubriste todos los gastos del mes.")
    elif r.dias_restantes:
        lineas.append(f"Te faltan {corto(r.falta)}: unos {corto(r.falta / r.dias_restantes)} por día "
                      f"en los {r.dias_restantes} días que quedan.")
    peor = min(r.productos, key=lambda p: p.margen)
    if peor.margen < MARGEN_MINIMO:
        lineas += ["", f"Ojo: {peor.articulo} deja solo {peor.margen * 100:.0f} % de margen. Conviene revisar el precio."]
    return "\n".join(lineas)


def informe(r: Resultado) -> str:
    sep = "─" * 64
    out = [sep, f"PUNTO DE EQUILIBRIO · {r.ultimo_dia:%m/%Y} · {r.dias_cargados} de {r.dias_apertura} días cargados", sep,
           f"Ventas del mes            {pesos(r.ventas_mes):>18}",
           f"Gastos fijos              {pesos(r.fijos):>18}",
           f"Mercadería (por receta)   {pct(r.insumos_pct):>18}",
           f"Total variables           {pct(r.variables_pct):>18}",
           f"Margen de contribución    {pct(r.margen_contribucion):>18}",
           f"Punto de equilibrio mes   {pesos(r.pe_mes):>18}",
           f"Punto de equilibrio día   {pesos(r.pe_dia):>18}",
           f"Proyección del mes        {pesos(r.proyeccion):>18}",
           f"Resultado proyectado      {pesos(r.resultado_proyectado):>18}",
           f"Margen de seguridad       {pct(r.margen_seguridad):>18}",
           "", "Días por debajo del equilibrio:"]
    flojos = [f"  {fecha_corta(d):<10} {pesos(v):>14}   faltó {corto(r.pe_dia - v)}"
              for d, v in r.por_dia.items() if v < r.pe_dia]
    out += flojos or ["  ninguno"]
    out += ["", f"{'Artículo':<20}{'Unid.':>7}{'Venta':>14}{'Margen':>9}"]
    for p in r.productos:
        marca = "  <- revisar precio" if p.margen < MARGEN_MINIMO else ""
        est = " *" if p.costo_estimado else ""
        out.append(f"{p.articulo[:20]:<20}{p.unidades:>7.0f}{pesos(p.venta):>14}{p.margen * 100:>8.0f}%{est}{marca}")
    if any(p.costo_estimado for p in r.productos):
        out.append(f"* sin costo de receta cargado, se estimó {COSTO_POR_DEFECTO * 100:.0f} %")
    return "\n".join(out)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--datos", type=Path, default=AQUI / "datos_ejemplo", help="carpeta con gastos, costos y parámetros")
    ap.add_argument("--ventas", type=Path, help="export de ventas (por defecto el de ejemplo)")
    args = ap.parse_args()

    ruta_ventas = args.ventas or args.datos / "ventas_bistrosoft_noviembre.csv"
    ventas = leer_ventas(ruta_ventas)
    fijos = sum(_num(f["importe_mensual"]) for f in _leer_csv(args.datos / "gastos_fijos.csv"))
    costos = {_norm(f["articulo"]): _num(f["costo_receta"]) for f in _leer_csv(args.datos / "costos_productos.csv")}
    par = {f["parametro"]: _num(f["valor"]) for f in _leer_csv(args.datos / "parametros.csv")}

    r = calcular(ventas, fijos, costos, par["comision_medios_pago_pct"], par["ingresos_brutos_pct"],
                 int(par["dias_apertura_mes"]))
    print(f"Leídas {len(ventas)} líneas de {ruta_ventas.name}\n")
    print(informe(r))
    texto = aviso(r)
    print("\n" + "─" * 64 + "\nAVISO PARA EL DUEÑO\n" + "─" * 64 + "\n" + texto)

    carpeta = AQUI / "avisos"
    carpeta.mkdir(exist_ok=True)
    destino = carpeta / f"aviso_{r.ultimo_dia.isoformat()}.txt"
    destino.write_text(texto + "\n", encoding="utf-8")
    print(f"\nAviso guardado en {destino.relative_to(AQUI)}")


if __name__ == "__main__":
    main()
