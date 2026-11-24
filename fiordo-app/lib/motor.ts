import { Config, Gasto, MesVentas, norm } from "./tipos";
import { VENTA_SIN_DETALLE } from "./lector";
import { ResumenCompras, resumirCompras } from "./compras";
import { DIAS_SEMANA, corto, diaSemana, diasDelMes, fechaCorta, pct } from "./formato";

export type Dia = { fecha: string; venta: number; tickets: number; cubrio: boolean; diferencia: number };
export type Producto = {
  articulo: string;
  unidades: number;
  venta: number;
  precioProm: number;
  participacion: number;
};
export type DiaSemana = { nombre: string; promedio: number; dias: number; cubiertos: number; vsNecesario: number };
export type Medio = { medio: string; venta: number; participacion: number; tasa: number; comision: number };

/** Agrupa los nombres que usa Bistrosoft en cada reporte ("Tarjeta", "Tarjetas", "Qr", "QR"…). */
export function medioCanonico(m: string): string {
  const n = norm(m);
  const partes = n.split(/\s*,\s*/).filter(Boolean);
  if (partes.length > 1) return "Pago combinado";
  if (/efectivo/.test(n)) return "Efectivo";
  if (/tarjeta|debito|credito|posnet/.test(n)) return "Tarjetas";
  if (/\bqr\b/.test(n)) return "QR";
  if (/online|link|transferencia/.test(n)) return "Online";
  return m.trim().charAt(0).toUpperCase() + m.trim().slice(1).toLowerCase();
}

/** Tasa de comisión de un medio: la cargada para ese medio, o la general para los que no tienen una. */
export function tasaMedio(medio: string, p: Config["parametros"]): number {
  const t = p.comisionesPorMedio?.[medio];
  if (typeof t === "number") return t;
  if (medio === "Pago combinado") {
    const ts = ["Efectivo", "QR"].map((k) => p.comisionesPorMedio?.[k] ?? 0);
    return ts.reduce((a, b) => a + b, 0) / ts.length;
  }
  return p.comisionPct;
}
export type Estado = "ok" | "warn" | "bad";

export type Resultado = {
  mes: string;
  hayDatos: boolean;
  ventasMes: number;
  fijos: number;
  gastos: Gasto[];
  insumosPct: number; // mercadería: compras de ARCA clasificadas como mercadería ÷ ventas
  comisionPct: number; // comisión promedio ponderada por medio de pago
  variablesPct: number;
  mc: number;
  peMes: number;
  peDia: number;
  diasApertura: number;
  dias: Dia[];
  productos: Producto[];
  semana: DiaSemana[];
  medios: Medio[];
  diasSinDetalle: number;
  compras: ResumenCompras;
  fijosManuales: number;
  sinMercaderia: boolean; // hay ventas pero no hay compras de mercadería cargadas en el mes
  // derivados
  ultimo?: Dia;
  promedioDia: number;
  proyeccion: number;
  falta: number;
  diasRestantes: number;
  necesarioRestante: number;
  margenSeguridad: number;
  resultadoProyectado: number;
  estadoDia: Estado;
  ticketPromedio: number;
};

export function gastosDelMes(config: Config, mes: string): Gasto[] {
  return config.gastosPorMes[mes] ?? config.gastosBase;
}

export function calcular(mesVentas: MesVentas | undefined, mes: string, config: Config): Resultado {
  const p = config.parametros;
  const manuales = gastosDelMes(config, mes);
  const compras = resumirCompras(mesVentas?.compras, config);
  const gastos: Gasto[] = [...manuales, ...compras.gastosFijos.map((g) => ({ ...g, origen: "arca" as const }))];
  const fijosManuales = manuales.reduce((s, g) => s + (Number(g.importe) || 0), 0);
  const fijos = fijosManuales + compras.fijos;
  // Comisiones según cómo se cobró cada venta (efectivo 0 %, tarjetas, QR, online…)
  const pm = new Map<string, number>();
  for (const m of mesVentas?.medios ?? []) {
    const c = medioCanonico(m.medio);
    pm.set(c, (pm.get(c) ?? 0) + m.importe);
  }
  const totalMedios = [...pm.values()].reduce((s, v) => s + v, 0);
  const medios: Medio[] = [...pm.entries()]
    .map(([medio, venta]) => {
      const tasa = tasaMedio(medio, p);
      return { medio, venta, participacion: totalMedios ? venta / totalMedios : 0, tasa, comision: (venta * tasa) / 100 };
    })
    .sort((a, b) => b.venta - a.venta);
  const comisionPct = totalMedios > 0 ? medios.reduce((s, m) => s + m.comision, 0) / totalMedios * 100 : p.comisionPct;
  const otrosVar = (comisionPct + p.iibbPct) / 100;
  const diasApertura = Math.max(1, diasDelMes(mes) - (p.diasCerrados || 0));
  const items = mesVentas?.items ?? [];

  const porDia = new Map<string, number>();
  const porProd = new Map<string, { articulo: string; unidades: number; venta: number }>();
  let ventasMes = 0;

  for (const it of items) {
    ventasMes += it.importe;
    porDia.set(it.fecha, (porDia.get(it.fecha) ?? 0) + it.importe);
    const clave = norm(it.articulo);
    const prev = porProd.get(clave) ?? { articulo: config.nombresArticulos[clave] ?? it.articulo, unidades: 0, venta: 0 };
    prev.unidades += it.cantidad;
    prev.venta += it.importe;
    porProd.set(clave, prev);
  }

  // Mercadería: lo comprado en el mes según ARCA (proveedores marcados "Mercadería") ÷ lo vendido
  const insumosPct = ventasMes > 0 ? Math.max(0, compras.mercaderia) / ventasMes : 0;
  const variablesPct = insumosPct + otrosVar;
  const mc = 1 - variablesPct;
  const peMes = mc > 0 ? fijos / mc : Infinity;
  const peDia = peMes / diasApertura;

  const dias: Dia[] = [...porDia.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([fecha, venta]) => ({
      fecha,
      venta,
      tickets: mesVentas?.tickets[fecha] ?? 0,
      cubrio: venta >= peDia,
      diferencia: venta - peDia,
    }));

  const productos: Producto[] = [...porProd.values()]
    .filter((x) => x.venta > 0 && x.articulo !== VENTA_SIN_DETALLE)
    .map((x) => ({
      ...x,
      precioProm: x.unidades ? x.venta / x.unidades : 0,
      participacion: ventasMes ? x.venta / ventasMes : 0,
    }))
    .sort((a, b) => b.venta - a.venta);

  const semana: DiaSemana[] = DIAS_SEMANA.map((nombre, i) => {
    const ds = dias.filter((d) => diaSemana(d.fecha) === i);
    const promedio = ds.length ? ds.reduce((s, d) => s + d.venta, 0) / ds.length : 0;
    return { nombre, promedio, dias: ds.length, cubiertos: ds.filter((d) => d.cubrio).length, vsNecesario: peDia ? promedio / peDia - 1 : 0 };
  }).filter((d) => d.dias > 0);

  const ultimo = dias[dias.length - 1];
  const promedioDia = dias.length ? ventasMes / dias.length : 0;
  const proyeccion = promedioDia * diasApertura;
  const falta = peMes - ventasMes;
  const diasRestantes = Math.max(0, diasApertura - dias.length);
  const ratio = ultimo && peDia ? ultimo.venta / peDia : 0;
  const totalTickets = dias.reduce((s, d) => s + d.tickets, 0);

  return {
    mes,
    hayDatos: ventasMes > 0,
    ventasMes,
    fijos,
    gastos,
    insumosPct,
    variablesPct,
    mc,
    peMes,
    peDia,
    diasApertura,
    dias,
    productos,
    semana,
    medios,
    diasSinDetalle: new Set(items.filter((i) => i.articulo === VENTA_SIN_DETALLE).map((i) => i.fecha)).size,
    compras,
    fijosManuales,
    comisionPct,
    sinMercaderia: ventasMes > 0 && compras.mercaderia <= 0,
    ultimo,
    promedioDia,
    proyeccion,
    falta,
    diasRestantes,
    necesarioRestante: diasRestantes && falta > 0 ? falta / diasRestantes : 0,
    margenSeguridad: proyeccion ? (proyeccion - peMes) / proyeccion : 0,
    resultadoProyectado: proyeccion * mc - fijos,
    estadoDia: ratio >= 1 ? "ok" : ratio >= 0.85 ? "warn" : "bad",
    ticketPromedio: totalTickets ? ventasMes / totalTickets : 0,
  };
}

export function avisoWhatsapp(r: Resultado, nombreLocal: string): string {
  if (!r.hayDatos || !r.ultimo) return "Todavía no hay ventas cargadas para este mes.";
  const u = r.ultimo;
  const l: string[] = [
    `Buen día ☕️ Resumen de ${fechaCorta(u.fecha)}${nombreLocal ? ` en ${nombreLocal}` : ""}:`,
    `Vendiste ${corto(u.venta)}. Para cubrir los gastos necesitás ${corto(r.peDia)} por día. ` +
      (u.cubrio ? "Día cubierto 🤎" : `Faltaron ${corto(-u.diferencia)}.`),
    "",
    `En el mes llevás ${corto(r.ventasMes)} de ${corto(r.peMes)} necesarios.`,
  ];
  if (r.falta <= 0) l.push("Ya cubriste todos los gastos del mes ✨");
  else if (r.diasRestantes)
    l.push(`Te faltan ${corto(r.falta)}: unos ${corto(r.necesarioRestante)} por día en los ${r.diasRestantes} días que quedan.`);
  const flojo = [...r.semana].filter((d) => d.vsNecesario < 0).sort((a, b) => a.vsNecesario - b.vsNecesario)[0];
  if (flojo)
    l.push("", `Los ${flojo.nombre.toLowerCase()} son el día más flojo: venden en promedio ${corto(flojo.promedio)}, ${pct(-flojo.vsNecesario, 0)} menos de lo necesario.`);
  return l.join("\n");
}

export function recomendaciones(r: Resultado): string[] {
  if (!r.hayDatos) return [];
  const out: string[] = [];
  if (r.fijos <= 0) out.push("Todavía no hay gastos cargados para este mes (sueldos, alquiler, servicios): el punto de equilibrio no es real hasta cargarlos.");
  if (r.sinMercaderia)
    out.push("Todavía no hay compras de mercadería cargadas para este mes: se calcula con 0 % de mercadería y el punto de equilibrio queda más bajo de lo real. Subí las compras de ARCA.");
  if (r.margenSeguridad < 0)
    out.push(`Al ritmo actual el mes cierra con pérdida de ${corto(-r.resultadoProyectado)}. Hace falta vender ${corto(r.peDia)} por día en promedio o bajar gastos.`);
  else if (r.margenSeguridad < 0.1)
    out.push(`El margen de seguridad es bajo (${pct(r.margenSeguridad)}): una caída chica de ventas deja el mes en pérdida.`);
  else out.push(`Las ventas proyectadas superan el punto de equilibrio en ${pct(r.margenSeguridad)}.`);
  const flojos = r.semana.filter((d) => d.vsNecesario < 0);
  if (flojos.length)
    out.push(`Días que en promedio no cubren los gastos: ${flojos.map((d) => d.nombre.toLowerCase()).join(", ")}. Conviene revisar horarios, personal o promociones.`);
  return out;
}

export type EstadoResultados = {
  proyectado: boolean; // el mes no terminó: se usa la venta proyectada al cierre
  ventas: number;
  cmv: number;
  comisiones: number;
  iibb: number;
  gastos: { concepto: string; importe: number }[]; // todos los gastos del mes en una sola lista
  totalGastos: number;
  resultado: number;
  margen: number; // resultado ÷ ventas
  llegaEquilibrio: boolean;
};

/**
 * Ingresos menos gastos del mes, en una sola lista. Por dentro, mercadería, ingresos brutos y comisiones
 * se calculan como % de la venta (suben si se vende más), para que el punto de equilibrio sea correcto.
 */
export function estadoResultados(r: Resultado, iibbPct: number): EstadoResultados {
  const proyectado = r.diasRestantes > 0;
  const ventas = proyectado ? r.proyeccion : r.ventasMes;
  const cmv = ventas * r.insumosPct;
  const comisiones = (ventas * r.comisionPct) / 100;
  const iibb = (ventas * iibbPct) / 100;
  const gastos = [
    { concepto: r.sinMercaderia ? "Mercadería (falta subir compras)" : "Mercadería", importe: cmv },
    ...r.gastos.filter((g) => Number(g.importe)).map((g) => ({ concepto: g.concepto, importe: Number(g.importe) })),
    { concepto: `Ingresos brutos (${iibbPct} %)`, importe: iibb },
    { concepto: "Comisiones de tarjetas, QR y online", importe: comisiones },
  ];
  const totalGastos = gastos.reduce((s, g) => s + g.importe, 0);
  const resultado = ventas - totalGastos;
  return {
    proyectado,
    ventas,
    cmv,
    comisiones,
    iibb,
    gastos,
    totalGastos,
    resultado,
    margen: ventas ? resultado / ventas : 0,
    llegaEquilibrio: isFinite(r.peMes) && ventas >= r.peMes,
  };
}
