import { Gasto, LineaVenta } from "./tipos";

/** Datos INVENTADOS para probar la app. Se borran desde Gastos > Borrar datos del mes. */
export const CATALOGO_EJEMPLO: { n: string; p: number; c: number; u: number }[] = [
  { n: "Espresso", p: 3200, c: 700, u: 40 },
  { n: "Cortado", p: 3800, c: 900, u: 45 },
  { n: "Café con leche", p: 4500, c: 1200, u: 35 },
  { n: "Capuchino", p: 5200, c: 1500, u: 30 },
  { n: "Medialuna", p: 1500, c: 450, u: 90 },
  { n: "Tostado J&Q", p: 8500, c: 3000, u: 15 },
  { n: "Avocado toast", p: 11000, c: 4200, u: 8 },
  { n: "Porción de torta", p: 6500, c: 2200, u: 12 },
  { n: "Jugo de naranja", p: 5500, c: 1900, u: 15 },
  { n: "Agua mineral", p: 2500, c: 1400, u: 20 },
  { n: "Brunch completo", p: 18000, c: 7000, u: 5 },
];

export const GASTOS_EJEMPLO: Gasto[] = [
  { concepto: "Alquiler", importe: 3500000 },
  { concepto: "Sueldos y cargas sociales", importe: 9000000 },
  { concepto: "Luz, gas, agua e internet", importe: 1200000 },
  { concepto: "Honorarios del estudio", importe: 600000 },
  { concepto: "Sistema de gestión", importe: 150000 },
  { concepto: "Seguros", importe: 250000 },
  { concepto: "Mantenimiento y limpieza", importe: 400000 },
  { concepto: "Tasas municipales", importe: 500000 },
];

const FACTOR = [0.75, 0.8, 0.85, 0.9, 1.1, 1.35, 1.25]; // lunes..domingo
const MEDIOS: [string, number][] = [["Efectivo", 30], ["Débito", 25], ["Crédito", 15], ["QR Mercado Pago", 30]];

function azar(semilla: number) {
  let a = semilla;
  return () => {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Ventas ticket por ticket del mes (AAAA-MM) hasta el día indicado. */
export function ventasEjemplo(mes: string, hastaDia: number): LineaVenta[] {
  const [a, m] = mes.split("-").map(Number);
  const rnd = azar(a * 100 + m);
  const elegir = <T,>(ops: [T, number][]) => {
    const tot = ops.reduce((s, o) => s + o[1], 0);
    let x = rnd() * tot;
    for (const [v, w] of ops) if ((x -= w) <= 0) return v;
    return ops[0][0];
  };
  const out: LineaVenta[] = [];
  let ticket = 10000;
  for (let d = 1; d <= hastaDia; d++) {
    const f = new Date(a, m - 1, d);
    const fecha = `${a}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    const objetivo = 1_000_000 * FACTOR[(f.getDay() + 6) % 7] * (d % 7 === 3 ? 0.8 : 1) * (0.9 + rnd() * 0.2);
    let vendido = 0;
    while (vendido < objetivo) {
      ticket++;
      const medio = elegir(MEDIOS);
      const items = new Map<string, number>();
      const n = elegir<number>([[1, 2], [2, 3], [3, 1], [4, 1]]);
      for (let i = 0; i < n; i++) {
        const p = elegir(CATALOGO_EJEMPLO.map((c) => [c, c.u] as [typeof c, number]));
        items.set(p.n, (items.get(p.n) ?? 0) + 1);
      }
      for (const [nombre, cant] of items) {
        const p = CATALOGO_EJEMPLO.find((c) => c.n === nombre)!;
        out.push({ fecha, articulo: nombre, cantidad: cant, importe: cant * p.p, medio, ticket: String(ticket) });
        vendido += cant * p.p;
      }
    }
  }
  return out;
}
