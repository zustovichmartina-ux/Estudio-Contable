/** Una línea de venta ya normalizada (la unidad mínima que guardamos). */
export type LineaVenta = {
  fecha: string; // AAAA-MM-DD
  articulo: string;
  cantidad: number;
  importe: number;
  medio?: string;
  ticket?: string;
  tickets?: number; // cantidad de tickets del día, cuando el reporte la trae resumida
};

/** Ventas de un mes, agregadas por día y artículo para que pesen poco. */
export type MesVentas = {
  mes: string; // AAAA-MM
  items: { fecha: string; articulo: string; cantidad: number; importe: number }[];
  medios: { fecha: string; medio: string; importe: number }[];
  tickets: Record<string, number>; // fecha -> cantidad de tickets
  archivos: string[];
  actualizado: string; // ISO
  compras?: Compra[]; // facturas de proveedores del mes (ARCA, Mis Comprobantes Recibidos)
};

/** Factura de compra. costo = importe sin IVA ni percepciones (negativo si es nota de crédito). */
export type Compra = { fecha: string; cuit: string; proveedor: string; tipo: string; numero: string; costo: number; total: number };

/** Cómo se usa cada proveedor: mercadería (costo variable), gasto fijo (con su rubro) o no se incluye. */
export type Proveedor = { nombre: string; categoria: "mercaderia" | "fijo" | "ignorar"; concepto: string };

export type Gasto = { concepto: string; importe: number; origen?: "arca" };

export type Parametros = {
  comisionPct: number; // tarjetas, QR, posnet
  iibbPct: number; // ingresos brutos
  diasCerrados: number; // días del mes que el local no abre (feriados, vacaciones)
  costoDefectoPct: number; // artículos sin costo de receta
  margenMinimoPct: number; // por debajo se marca para revisar
  mercaderiaSegun: "auto" | "compras" | "recetas"; // de dónde sale el costo de mercadería
  ivaComputable: boolean; // Responsable Inscripto: el IVA de compras no es costo
};

export type Config = {
  nombreLocal: string;
  parametros: Parametros;
  costos: Record<string, number>; // artículo (normalizado) -> costo unitario
  nombresArticulos: Record<string, string>; // normalizado -> como se escribe
  gastosBase: Gasto[]; // se copian a cada mes nuevo
  gastosPorMes: Record<string, Gasto[]>; // AAAA-MM -> gastos de ese mes
  proveedores: Record<string, Proveedor>; // CUIT -> clasificación
};

export type Datos = { config: Config; meses: Record<string, MesVentas> };

export const PARAMETROS_INICIALES: Parametros = {
  comisionPct: 3.5,
  iibbPct: 3.0,
  diasCerrados: 0,
  costoDefectoPct: 32,
  margenMinimoPct: 55,
  mercaderiaSegun: "auto",
  ivaComputable: true,
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
  proveedores: {},
};

export const norm = (s: string) =>
  s.normalize("NFKD").replace(/[̀-ͯ]/g, "").trim().toLowerCase().replace(/\s+/g, " ");
