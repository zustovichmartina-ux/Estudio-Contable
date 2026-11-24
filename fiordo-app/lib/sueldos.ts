import { Hoja, aNumero } from "./lector";
import { Gasto, norm } from "./tipos";

/**
 * Liquidación de sueldos del estudio (Excel con una hoja por mes, nombrada AAAAMM, ej. "202610").
 * De cada hoja se toma la columna TOTALES:
 *  - Sueldos = TOTAL HABERES REMUNERATIVOS + TOTAL NO REMUNERATIVOS (lo que cuesta el sueldo bruto).
 *  - Cargas sociales = contribuciones patronales que están después de SUELDO NETO (previsional, PAMI,
 *    fondo de empleo, ANSSAL, asignaciones familiares, obra social, ART, seguro colectivo, seguro de vida,
 *    fondo convenio, contribución adicional). Los aportes del empleado no se suman: ya están en el bruto.
 */
export type SueldosMes = { mes: string; sueldos: number; cargas: number; empleados: number };

const CARGAS = [
  /^cont\.? previsional/,
  /^cont\.? pami/,
  /^fdo\.? nac/,
  /^fondo nac/,
  /^cont\.? anssal/,
  /^cont\.? asign/,
  /^cont\.? obra social/,
  /^art\b/,
  /^svco/,
  /^seguro de vida/,
  /^fondo convenio/,
  /^contribucion adicional/,
];

function etiqueta(fila: unknown[]): { texto: string; col: number } {
  for (let i = 0; i < Math.min(fila.length, 4); i++) {
    const v = fila[i];
    if (typeof v === "string" && v.trim()) return { texto: norm(v), col: i };
  }
  return { texto: "", col: -1 };
}

export function esLiquidacionSueldos(h: Hoja): boolean {
  return /^\d{6}$/.test(h.nombre.trim()) && h.filas.some((f) => etiqueta(f).texto.startsWith("total haberes remunerativos"));
}

export function leerLiquidacion(h: Hoja): SueldosMes | null {
  const nombre = h.nombre.trim();
  const mes = `${nombre.slice(0, 4)}-${nombre.slice(4, 6)}`;
  // columna TOTALES: la de la fila de encabezado que dice "TOTALES"
  let colTot = -1;
  let empleados = 0;
  for (const f of h.filas) {
    const i = f.findIndex((c) => typeof c === "string" && norm(c) === "totales");
    if (i >= 0) {
      colTot = i;
      const e = etiqueta(f);
      empleados = f.slice(e.col + 1, i).filter((c) => typeof c === "string" && c.trim()).length;
      break;
    }
  }
  if (colTot < 0) return null;
  const valor = (f: unknown[]) => {
    const n = aNumero(f[colTot]);
    return isFinite(n) ? n : 0;
  };
  let rem = 0, noRem = 0, cargas = 0, pasoNeto = false;
  for (const f of h.filas) {
    const { texto } = etiqueta(f);
    if (!texto) continue;
    if (texto.startsWith("total haberes remunerativos")) rem = valor(f);
    else if (texto.startsWith("total no remunerativos")) noRem = valor(f);
    else if (texto.startsWith("sueldo neto")) pasoNeto = true;
    else if (pasoNeto && CARGAS.some((re) => re.test(texto))) cargas += valor(f);
  }
  const sueldos = rem + noRem;
  if (!sueldos) return null;
  return { mes, sueldos: Math.round(sueldos), cargas: Math.round(cargas), empleados };
}

/** Reemplaza en la lista de gastos del mes los renglones de sueldos y cargas por los de la liquidación. */
export function aplicarSueldos(gastos: Gasto[], s: SueldosMes): Gasto[] {
  const otros = gastos.filter((g) => !/sueldo|cargas sociales/i.test(g.concepto));
  return [{ concepto: "Sueldos", importe: s.sueldos }, { concepto: "Cargas sociales", importe: s.cargas }, ...otros];
}
