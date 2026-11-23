import { Hoja, aFechaISO, aNumero } from "./lector";
import { Compra, Config, MesVentas, Proveedor, norm } from "./tipos";

/**
 * Compras desde ARCA: "Mis Comprobantes Recibidos" (xlsx o csv), formato viejo o nuevo.
 * Cada factura se toma SIN IVA ni percepciones (si el café es Responsable Inscripto esos
 * importes son crédito, no costo). Las notas de crédito restan.
 */

const CODIGOS_NC = new Set([3, 8, 13, 21, 38, 43, 44, 48, 53, 110, 112, 113, 114, 119, 203, 208, 213]);

function col(enc: string[], ...pruebas: ((h: string) => boolean)[]): number {
  for (const p of pruebas) {
    const i = enc.findIndex(p);
    if (i >= 0) return i;
  }
  return -1;
}

function columnas(h: Hoja) {
  const n = h.encabezados.map((e) => norm(e));
  return {
    fecha: col(n, (e) => e.startsWith("fecha")),
    tipo: col(n, (e) => e === "tipo" || e.startsWith("tipo de comprobante") || (e.startsWith("tipo") && !e.includes("doc") && !e.includes("cambio"))),
    pv: col(n, (e) => e.startsWith("punto de venta")),
    numero: col(n, (e) => e.startsWith("numero desde") || e.startsWith("nro. desde"), (e) => (e.startsWith("numero") || e.startsWith("nro")) && !e.includes("hasta") && !e.includes("doc")),
    cuit: col(n, (e) => /^(nro\.?|numero) (de )?doc/.test(e) && /vendedor|emisor/.test(e), (e) => e.startsWith("cuit")),
    nombre: col(n, (e) => e.startsWith("denominacion vendedor") || e.startsWith("denominacion emisor"), (e) => e.includes("razon social")),
    total: col(n, (e) => e.startsWith("imp. total") || e.startsWith("importe total"), (e) => e === "total"),
    iva: col(n, (e) => e === "total iva" || e === "imp. total iva", (e) => e === "iva" || e === "imp. iva"),
    tributos: col(n, (e) => e.includes("otros tributos")),
    moneda: col(n, (e) => e === "moneda"),
    cambio: col(n, (e) => e.startsWith("tipo cambio") || e.startsWith("tipo de cambio")),
  };
}

export function esArcaCompras(h: Hoja): boolean {
  const c = columnas(h);
  return c.fecha >= 0 && c.total >= 0 && c.nombre >= 0 && c.cuit >= 0;
}

export function leerComprasArca(h: Hoja, ivaComputable: boolean): Compra[] {
  const c = columnas(h);
  const out: Compra[] = [];
  for (const fila of h.filas.slice(h.filaEncabezado + 1)) {
    const fecha = aFechaISO(fila[c.fecha]);
    let total = aNumero(fila[c.total]);
    if (!fecha || !isFinite(total)) continue;
    const tipo = String(fila[c.tipo] ?? "").trim();
    const codigo = parseInt(tipo, 10);
    const esNC = /cr[eé]dito/i.test(tipo) || CODIGOS_NC.has(codigo);
    const tc = c.cambio >= 0 ? aNumero(fila[c.cambio]) : 1;
    const enPesos = c.moneda < 0 || /^(pes|\$|ars)/i.test(String(fila[c.moneda] ?? "PES"));
    const factor = !enPesos && isFinite(tc) && tc > 0 ? tc : 1;
    const iva = c.iva >= 0 ? aNumero(fila[c.iva]) || 0 : 0;
    const trib = c.tributos >= 0 ? aNumero(fila[c.tributos]) || 0 : 0;
    let costo = ivaComputable ? total - iva - trib : total;
    total *= factor;
    costo *= factor;
    const signo = esNC ? -1 : 1;
    out.push({
      fecha,
      cuit: String(fila[c.cuit] ?? "").replace(/\D/g, "") || "sin-cuit",
      proveedor: String(fila[c.nombre] ?? "").trim() || "Sin nombre",
      tipo: tipo.replace(/^\d+\s*-\s*/, ""),
      numero: [c.pv >= 0 ? fila[c.pv] : "", c.numero >= 0 ? fila[c.numero] : ""].filter((x) => x !== "" && x != null).join("-"),
      costo: Math.round(signo * Math.abs(costo) * 100) / 100,
      total: Math.round(signo * Math.abs(total) * 100) / 100,
    });
  }
  return out;
}

/** Guarda las compras por mes; reemplaza las que ya había en el rango de fechas del archivo. */
export function incorporarCompras(meses: Record<string, MesVentas>, compras: Compra[], archivo: string): Record<string, MesVentas> {
  if (!compras.length) return meses;
  const fechas = compras.map((c) => c.fecha).sort();
  const desde = fechas[0], hasta = fechas[fechas.length - 1];
  const out = { ...meses };
  const porMes = new Map<string, Compra[]>();
  for (const c of compras) {
    const m = c.fecha.slice(0, 7);
    if (!porMes.has(m)) porMes.set(m, []);
    porMes.get(m)!.push(c);
  }
  for (const [mes, cs] of porMes) {
    const prev = meses[mes];
    const quedan = (prev?.compras ?? []).filter((c) => c.fecha < desde || c.fecha > hasta);
    out[mes] = {
      mes,
      items: prev?.items ?? [],
      medios: prev?.medios ?? [],
      tickets: prev?.tickets ?? {},
      archivos: [...new Set([...(prev?.archivos ?? []), archivo])],
      actualizado: new Date().toISOString(),
      compras: [...quedan, ...cs],
    };
  }
  return out;
}

// ---------------------------------------------------------------- clasificación de proveedores
export const CONCEPTOS_FIJOS = [
  "Alquiler",
  "Luz, gas, agua e internet",
  "Honorarios del estudio",
  "Sistema de gestión",
  "Seguros",
  "Mantenimiento y limpieza",
  "Tasas municipales",
  "Otros gastos fijos",
];

const REGLAS: [RegExp, Proveedor["categoria"], string][] = [
  [/\b(edea|edenor|edesur|epec|camuzzi|metrogas|naturgy|litoral gas|osse|aysa|aguas|telecom|personal|movistar|claro|fibertel|telefonica|cablevision|iplan|starlink|telered)\b/, "fijo", "Luz, gas, agua e internet"],
  [/\b(seguro|seguros|sancor|federacion patronal|la caja|mapfre|allianz|zurich|rivadavia|la segunda|mercantil andina)\b/, "fijo", "Seguros"],
  [/\b(bistrosoft|fudo|google|microsoft|meta platforms|canva|spotify)\b/, "fijo", "Sistema de gestión"],
  [/\b(inmobiliaria|propiedades|alquiler)\b/, "fijo", "Alquiler"],
  [/\b(limpieza|fumig|mantenimiento|matafuego)\b/, "fijo", "Mantenimiento y limpieza"],
  [/\b(estudio contable|contador)\b/, "fijo", "Honorarios del estudio"],
  [/\b(municipalidad|arba|agip|afip|arca)\b/, "ignorar", ""],
];

export function sugerirProveedor(nombre: string): Proveedor {
  const n = norm(nombre);
  for (const [re, categoria, concepto] of REGLAS) if (re.test(n)) return { nombre, categoria, concepto };
  return { nombre, categoria: "mercaderia", concepto: "" };
}

export function clasificar(config: Config, cuit: string, nombre: string): Proveedor & { sugerido: boolean } {
  const p = config.proveedores?.[cuit];
  return p ? { ...p, sugerido: false } : { ...sugerirProveedor(nombre), sugerido: true };
}

export type ResumenCompras = {
  cantidad: number;
  mercaderia: number;
  fijos: number;
  ignorado: number;
  gastosFijos: { concepto: string; importe: number }[];
  proveedores: { cuit: string; nombre: string; comprobantes: number; costo: number; categoria: Proveedor["categoria"]; concepto: string; sugerido: boolean }[];
};

export function resumirCompras(compras: Compra[] | undefined, config: Config): ResumenCompras {
  const r: ResumenCompras = { cantidad: 0, mercaderia: 0, fijos: 0, ignorado: 0, gastosFijos: [], proveedores: [] };
  const prov = new Map<string, ResumenCompras["proveedores"][number]>();
  const fijos = new Map<string, number>();
  for (const c of compras ?? []) {
    r.cantidad++;
    const k = clasificar(config, c.cuit, c.proveedor);
    if (k.categoria === "mercaderia") r.mercaderia += c.costo;
    else if (k.categoria === "fijo") {
      r.fijos += c.costo;
      const concepto = k.concepto || "Otros gastos fijos";
      fijos.set(concepto, (fijos.get(concepto) ?? 0) + c.costo);
    } else r.ignorado += c.costo;
    const p = prov.get(c.cuit) ?? { cuit: c.cuit, nombre: c.proveedor, comprobantes: 0, costo: 0, categoria: k.categoria, concepto: k.concepto, sugerido: k.sugerido };
    p.comprobantes++;
    p.costo += c.costo;
    prov.set(c.cuit, p);
  }
  r.gastosFijos = [...fijos.entries()].map(([concepto, importe]) => ({ concepto, importe })).sort((a, b) => b.importe - a.importe);
  r.proveedores = [...prov.values()].sort((a, b) => b.costo - a.costo);
  return r;
}
