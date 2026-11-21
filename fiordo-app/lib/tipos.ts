/** Una línea de venta ya normalizada (la unidad mínima que guardamos). */
export type LineaVenta = {
  fecha: string; // AAAA-MM-DD
  articulo: string;
  cantidad: number;
  importe: number;
  medio?: string;
  ticket?: string;
};

/** Ventas de un mes, agregadas por día y artículo para que pesen poco. */
export type MesVentas = {
  mes: string; // AAAA-MM
  items: { fecha: string; articulo: string; cantidad: number; importe: number }[];
  medios: { fecha: string; medio: string; importe: number }[];
  tickets: Record<string, number>; // fecha -> cantidad de tickets
  archivos: string[];
  actualizado: string; // ISO
};

export type Gasto = { concepto: string; importe: number };

export type Parametros = {
  comisionPct: number; // tarjetas, QR, posnet
  iibbPct: number; // ingresos brutos
  diasCerrados: number; // días del mes que el local no abre (feriados, vacaciones)
  costoDefectoPct: number; // artículos sin costo de receta
  margenMinimoPct: number; // por debajo se marca para revisar
};

export type Config = {
  nombreLocal: string;
  parametros: Parametros;
  costos: Record<string, number>; // artículo (normalizado) -> costo unitario
  nombresArticulos: Record<string, string>; // normalizado -> como se escribe
  gastosBase: Gasto[]; // se copian a cada mes nuevo
  gastosPorMes: Record<string, Gasto[]>; // AAAA-MM -> gastos de ese mes
};

export type Datos = { config: Config; meses: Record<string, MesVentas> };

export const PARAMETROS_INICIALES: Parametros = {
  comisionPct: 3.5,
  iibbPct: 3.0,
  diasCerrados: 0,
  costoDefectoPct: 32,
  margenMinimoPct: 55,
};

export const GASTOS_INICIALES: Gasto[] = [
  { concepto: "Alquiler", importe: 0 },
  { concepto: "Sueldos y cargas sociales", importe: 0 },
  { concepto: "Luz, gas, agua e internet", importe: 0 },
  { concepto: "Honorarios del estudio", importe: 0 },
  { concepto: "Sistema de gestión", importe: 0 },
  { concepto: "Seguros", importe: 0 },
  { concepto: "Mantenimiento y limpieza", importe: 0 },
  { concepto: "Tasas municipales", importe: 0 },
];

export const CONFIG_INICIAL: Config = {
  nombreLocal: "Fiordo Café",
  parametros: PARAMETROS_INICIALES,
  costos: {},
  nombresArticulos: {},
  gastosBase: GASTOS_INICIALES,
  gastosPorMes: {},
};

export const norm = (s: string) =>
  s.normalize("NFKD").replace(/[̀-ͯ]/g, "").trim().toLowerCase().replace(/\s+/g, " ");
