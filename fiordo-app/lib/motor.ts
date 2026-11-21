import { Config, Gasto, MesVentas, norm } from "./tipos";
import { DIAS_SEMANA, corto, diaSemana, diasDelMes, fechaCorta, pct } from "./formato";

export type Dia = { fecha: string; venta: number; tickets: number; cubrio: boolean; diferencia: number };
export type Producto = {
  articulo: string;
  unidades: number;
  venta: number;
  costo: number;
  precioProm: number;
  costoU: number;
  margen: number;
  aporte: number;
  participacion: number;
  conCosto: boolean;
  revisar: boolean;
};
export type DiaSemana = { nombre: string; promedio: number; dias: number; cubiertos: number; vsNecesario: number };
export type Medio = { medio: string; venta: number; participacion: number };
export type Estado = "ok" | "warn" | "bad";

export type Resultado = {
  mes: string;
  hayDatos: boolean;
  ventasMes: number;
  fijos: number;
  gastos: Gasto[];
  insumosPct: number;
  variablesPct: number;
  mc: number;
  peMes: number;
  peDia: number;
  diasApertura: number;
  dias: Dia[];
  productos: Producto[];
  semana: DiaSemana[];
  medios: Medio[];
  sinCosto: string[];
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
  const gastos = gastosDelMes(config, mes);
  const fijos = gastos.reduce((s, g) => s + (Number(g.importe) || 0), 0);
  const otrosVar = (p.comisionPct + p.iibbPct) / 100;
  const diasApertura = Math.max(1, diasDelMes(mes) - (p.diasCerrados || 0));
  const items = mesVentas?.items ?? [];

  const porDia = new Map<string, number>();
  const porProd = new Map<string, { articulo: string; unidades: number; venta: number; costo: number; conCosto: boolean }>();
  let ventasMes = 0;
  let costoTotal = 0;

  for (const it of items) {
    ventasMes += it.importe;
    porDia.set(it.fecha, (porDia.get(it.fecha) ?? 0) + it.importe);
    const clave = norm(it.articulo);
    const cu = config.costos[clave];
    const tieneCosto = typeof cu === "number" && cu > 0;
    // un descuento baja la venta pero no el costo de la mercadería
    const costo = it.importe <= 0 ? 0 : tieneCosto ? cu * it.cantidad : (it.importe * p.costoDefectoPct) / 100;
    costoTotal += costo;
    const prev = porProd.get(clave) ?? { articulo: config.nombresArticulos[clave] ?? it.articulo, unidades: 0, venta: 0, costo: 0, conCosto: true };
    prev.unidades += it.cantidad;
    prev.venta += it.importe;
    prev.costo += costo;
    prev.conCosto = prev.conCosto && tieneCosto;
    porProd.set(clave, prev);
  }

  const insumosPct = ventasMes ? costoTotal / ventasMes : 0;
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
    .filter((x) => x.venta > 0)
    .map((x) => {
      const margen = x.venta ? 1 - x.costo / x.venta : 0;
      return {
        ...x,
        precioProm: x.unidades ? x.venta / x.unidades : 0,
        costoU: x.unidades ? x.costo / x.unidades : 0,
        margen,
        aporte: x.venta * (1 - otrosVar) - x.costo,
        participacion: ventasMes ? x.venta / ventasMes : 0,
        revisar: margen < p.margenMinimoPct / 100,
      };
    })
    .sort((a, b) => b.aporte - a.aporte);

  const semana: DiaSemana[] = DIAS_SEMANA.map((nombre, i) => {
    const ds = dias.filter((d) => diaSemana(d.fecha) === i);
    const promedio = ds.length ? ds.reduce((s, d) => s + d.venta, 0) / ds.length : 0;
    return { nombre, promedio, dias: ds.length, cubiertos: ds.filter((d) => d.cubrio).length, vsNecesario: peDia ? promedio / peDia - 1 : 0 };
  }).filter((d) => d.dias > 0);

  const pm = new Map<string, number>();
  for (const m of mesVentas?.medios ?? []) pm.set(m.medio, (pm.get(m.medio) ?? 0) + m.importe);
  const totalMedios = [...pm.values()].reduce((s, v) => s + v, 0);
  const medios: Medio[] = [...pm.entries()]
    .map(([medio, venta]) => ({ medio, venta, participacion: totalMedios ? venta / totalMedios : 0 }))
    .sort((a, b) => b.venta - a.venta);

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
    sinCosto: productos.filter((x) => !x.conCosto).map((x) => x.articulo).sort(),
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
  const peor = r.productos.filter((x) => x.revisar).sort((a, b) => a.margen - b.margen)[0];
  if (peor) l.push("", `Ojo: ${peor.articulo} deja solo ${pct(peor.margen, 0)} de margen. Conviene revisar el precio.`);
  return l.join("\n");
}

export function recomendaciones(r: Resultado): string[] {
  if (!r.hayDatos) return [];
  const out: string[] = [];
  if (r.fijos <= 0) out.push("Todavía no hay gastos fijos cargados para este mes: el punto de equilibrio no es real hasta cargarlos.");
  if (r.margenSeguridad < 0)
    out.push(`Al ritmo actual el mes cierra con pérdida de ${corto(-r.resultadoProyectado)}. Hace falta vender ${corto(r.peDia)} por día en promedio o bajar gastos fijos.`);
  else if (r.margenSeguridad < 0.1)
    out.push(`El margen de seguridad es bajo (${pct(r.margenSeguridad)}): una caída chica de ventas deja el mes en pérdida.`);
  else out.push(`Las ventas proyectadas superan el punto de equilibrio en ${pct(r.margenSeguridad)}.`);
  const flojos = r.semana.filter((d) => d.vsNecesario < 0);
  if (flojos.length)
    out.push(`Días que en promedio no cubren los gastos: ${flojos.map((d) => d.nombre.toLowerCase()).join(", ")}. Conviene revisar horarios, personal o promociones.`);
  for (const x of r.productos.filter((x) => x.revisar)) out.push(`${x.articulo}: margen de ${pct(x.margen, 0)}. Revisar precio o costo de receta.`);
  if (r.sinCosto.length)
    out.push(`${r.sinCosto.length} artículos sin costo de receta cargado; se estimó su costo. Cargarlos en Productos mejora la precisión.`);
  return out;
}
