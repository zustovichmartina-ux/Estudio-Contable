import * as XLSX from "xlsx";
import { LineaVenta, MesVentas, norm } from "./tipos";

export type Campo = "fecha" | "articulo" | "cantidad" | "importe" | "medio" | "ticket";
export type Mapeo = Record<Campo, number>; // índice de columna, -1 = no está

export type Hoja = {
  nombre: string;
  filas: unknown[][];
  filaEncabezado: number;
  encabezados: string[];
  mapeo: Mapeo;
};

export const CAMPOS: { campo: Campo; etiqueta: string; obligatorio: boolean }[] = [
  { campo: "fecha", etiqueta: "Fecha", obligatorio: true },
  { campo: "importe", etiqueta: "Importe / total", obligatorio: true },
  { campo: "articulo", etiqueta: "Artículo / producto", obligatorio: false },
  { campo: "cantidad", etiqueta: "Cantidad", obligatorio: false },
  { campo: "medio", etiqueta: "Medio de pago", obligatorio: false },
  { campo: "ticket", etiqueta: "Nº de ticket / comanda", obligatorio: false },
];

const PATRONES: Record<Campo, RegExp[]> = {
  fecha: [/^fecha/, /fecha/, /^dia$/, /date/],
  importe: [/^importe/, /^total/, /importe/, /total/, /monto/, /neto/, /venta/],
  articulo: [/articulo/, /producto/, /descripcion/, /^item/, /plato/, /nombre/],
  cantidad: [/^cant/, /cantidad/, /unidades/, /^qty/],
  medio: [/medio de pago/, /forma de pago/, /medio/, /pago/],
  ticket: [/ticket/, /comanda/, /comprobante/, /^nro/, /numero/, /orden/],
};

function detectarMapeo(enc: string[]): Mapeo {
  const n = enc.map((e) => norm(String(e ?? "")));
  const usados = new Set<number>();
  const m = {} as Mapeo;
  for (const campo of ["fecha", "articulo", "cantidad", "medio", "ticket", "importe"] as Campo[]) {
    m[campo] = -1;
    for (const re of PATRONES[campo]) {
      const i = n.findIndex((h, idx) => !usados.has(idx) && re.test(h) && !(campo === "importe" && /unit|precio/.test(h)));
      if (i >= 0) {
        m[campo] = i;
        usados.add(i);
        break;
      }
    }
  }
  return m;
}

/** Busca la fila de títulos: la primera que tenga algo parecido a fecha y a importe. */
function detectarEncabezado(filas: unknown[][]): number {
  for (let i = 0; i < Math.min(filas.length, 30); i++) {
    const n = (filas[i] ?? []).map((c) => norm(String(c ?? "")));
    if (n.some((h) => /fecha|dia/.test(h)) && n.some((h) => /importe|total|monto|venta/.test(h))) return i;
  }
  for (let i = 0; i < Math.min(filas.length, 30); i++) {
    if ((filas[i] ?? []).filter((c) => typeof c === "string" && c.trim()).length >= 3) return i;
  }
  return 0;
}

export async function leerArchivo(archivo: File): Promise<Hoja[]> {
  const buf = await archivo.arrayBuffer();
  const esTexto = /\.(csv|txt)$/i.test(archivo.name);
  const wb = esTexto
    ? XLSX.read(new TextDecoder("utf-8").decode(buf).replace(/^﻿/, ""), { type: "string", cellDates: true, raw: false, FS: detectarSeparador(buf) })
    : XLSX.read(buf, { type: "array", cellDates: true });
  return wb.SheetNames.map((nombre) => {
    const filas = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[nombre], { header: 1, raw: true, defval: null, blankrows: false });
    const filaEncabezado = detectarEncabezado(filas);
    const encabezados = (filas[filaEncabezado] ?? []).map((c, i) => (c == null || c === "" ? `Columna ${i + 1}` : String(c).trim()));
    return { nombre, filas, filaEncabezado, encabezados, mapeo: detectarMapeo(encabezados) };
  }).filter((h) => h.filas.length > 1);
}

function detectarSeparador(buf: ArrayBuffer): string {
  const primera = new TextDecoder("utf-8").decode(buf.slice(0, 2000)).split(/\r?\n/)[0] ?? "";
  return (primera.match(/;/g)?.length ?? 0) >= (primera.match(/,/g)?.length ?? 0) ? ";" : ",";
}

// ---------------------------------------------------------------- conversión de valores
const iso = (a: number, m: number, d: number) => `${a}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;

export function aFechaISO(v: unknown): string | null {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v.getTime())) return iso(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = XLSX.SSF.parse_date_code(v);
    return d ? iso(d.y, d.m, d.d) : null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/);
  if (m) {
    let a = +m[3];
    if (a < 100) a += 2000;
    return +m[2] >= 1 && +m[2] <= 12 ? iso(a, +m[2], +m[1]) : null;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return iso(+m[1], +m[2], +m[3]);
  return null;
}

export function aNumero(v: unknown): number {
  if (typeof v === "number") return v;
  if (v == null) return NaN;
  let s = String(v).replace(/[^\d,.\-]/g, "");
  if (!s) return NaN;
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else if (/\.\d{3}(\.|$)/.test(s) && !/,/.test(s)) s = s.replace(/\./g, "");
  else s = s.replace(/,/g, "");
  return parseFloat(s);
}

export type ResultadoLectura = { lineas: LineaVenta[]; descartadas: number; desde?: string; hasta?: string };

export function convertir(h: Hoja, fechaFija?: string): ResultadoLectura {
  const m = h.mapeo;
  const lineas: LineaVenta[] = [];
  let descartadas = 0;
  for (const fila of h.filas.slice(h.filaEncabezado + 1)) {
    const fecha = m.fecha >= 0 ? aFechaISO(fila[m.fecha]) : fechaFija ?? null;
    const importe = m.importe >= 0 ? aNumero(fila[m.importe]) : NaN;
    const articulo = m.articulo >= 0 ? String(fila[m.articulo] ?? "").trim() : "Venta";
    // filas de totales o subtotales: sin fecha, sin importe o con artículo "total"
    if (!fecha || !isFinite(importe) || /^(sub)?total/i.test(articulo)) {
      descartadas++;
      continue;
    }
    const cantidad = m.cantidad >= 0 ? aNumero(fila[m.cantidad]) : 1;
    lineas.push({
      fecha,
      articulo: articulo || "Sin detalle",
      cantidad: isFinite(cantidad) ? cantidad : 1,
      importe,
      medio: m.medio >= 0 && fila[m.medio] != null ? String(fila[m.medio]).trim() : undefined,
      ticket: m.ticket >= 0 && fila[m.ticket] != null ? String(fila[m.ticket]).trim() : undefined,
    });
  }
  const fechas = lineas.map((l) => l.fecha).sort();
  return { lineas, descartadas, desde: fechas[0], hasta: fechas[fechas.length - 1] };
}

/** Agrupa las líneas por mes y las suma en los meses existentes, reemplazando los días que vienen en el archivo. */
export function incorporar(meses: Record<string, MesVentas>, lineas: LineaVenta[], archivo: string): Record<string, MesVentas> {
  const porMes = new Map<string, LineaVenta[]>();
  for (const l of lineas) {
    const mes = l.fecha.slice(0, 7);
    if (!porMes.has(mes)) porMes.set(mes, []);
    porMes.get(mes)!.push(l);
  }
  const out = { ...meses };
  for (const [mes, ls] of porMes) {
    const fechasNuevas = new Set(ls.map((l) => l.fecha));
    const prev = meses[mes];
    const items = new Map<string, { fecha: string; articulo: string; cantidad: number; importe: number }>();
    const medios = new Map<string, { fecha: string; medio: string; importe: number }>();
    const tickets: Record<string, number> = {};
    for (const it of prev?.items ?? []) if (!fechasNuevas.has(it.fecha)) items.set(`${it.fecha}|${norm(it.articulo)}`, { ...it });
    for (const md of prev?.medios ?? []) if (!fechasNuevas.has(md.fecha)) medios.set(`${md.fecha}|${md.medio}`, { ...md });
    for (const [f, n] of Object.entries(prev?.tickets ?? {})) if (!fechasNuevas.has(f)) tickets[f] = n;

    const tk = new Map<string, Set<string>>();
    for (const l of ls) {
      const k = `${l.fecha}|${norm(l.articulo)}`;
      const it = items.get(k) ?? { fecha: l.fecha, articulo: l.articulo, cantidad: 0, importe: 0 };
      it.cantidad += l.cantidad;
      it.importe += l.importe;
      items.set(k, it);
      if (l.medio) {
        const km = `${l.fecha}|${l.medio}`;
        const md = medios.get(km) ?? { fecha: l.fecha, medio: l.medio, importe: 0 };
        md.importe += l.importe;
        medios.set(km, md);
      }
      if (l.ticket) {
        if (!tk.has(l.fecha)) tk.set(l.fecha, new Set());
        tk.get(l.fecha)!.add(l.ticket);
      }
    }
    for (const [f, s] of tk) tickets[f] = s.size;
    out[mes] = {
      mes,
      items: [...items.values()].map((i) => ({ ...i, importe: Math.round(i.importe * 100) / 100 })),
      medios: [...medios.values()],
      tickets,
      archivos: [...new Set([...(prev?.archivos ?? []), archivo])],
      actualizado: new Date().toISOString(),
    };
  }
  return out;
}
