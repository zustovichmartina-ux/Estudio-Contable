import { Resultado, recomendaciones } from "./motor";
import { corto, fechaCorta, nombreMes, pct, pesos } from "./formato";
import { MARCA, hexARgb } from "./tema";

/** Descarga un archivo generado en el navegador. */
function descargar(blob: Blob, nombre: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

const slug = (s: string) => s.normalize("NFKD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]+/g, "_").toLowerCase();

// ---------------------------------------------------------------- PDF mensual
export async function reportePDF(r: Resultado, nombreLocal: string) {
  const { jsPDF } = await import("jspdf");
  const autoTable = (await import("jspdf-autotable")).default;
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const W = doc.internal.pageSize.getWidth();
  const M = 14;
  const prim = hexARgb(MARCA.primario);
  const gris = hexARgb(MARCA.gris);
  const tinta = hexARgb(MARCA.tinta);

  // Encabezado
  doc.setFillColor(...prim);
  doc.rect(0, 0, W, 30, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text(nombreLocal || "Punto de equilibrio", M, 14);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.text(`Informe de punto de equilibrio · ${nombreMes(r.mes)}`, M, 22);
  doc.setFontSize(8);
  doc.text(`Generado el ${new Date().toLocaleDateString("es-AR")} · ${r.dias.length} de ${r.diasApertura} días con ventas cargadas`, W - M, 22, { align: "right" });

  // Indicadores
  const kpis: [string, string][] = [
    ["Ventas del mes", pesos(r.ventasMes)],
    ["Punto de equilibrio", pesos(r.peMes)],
    ["Necesario por día", pesos(r.peDia)],
    ["Proyección del mes", pesos(r.proyeccion)],
    ["Margen de contribución", pct(r.mc)],
    ["Margen de seguridad", pct(r.margenSeguridad)],
    ["Resultado proyectado", pesos(r.resultadoProyectado)],
    ["Ticket promedio", r.ticketPromedio ? pesos(r.ticketPromedio) : "–"],
  ];
  let y = 38;
  const cw = (W - 2 * M - 9) / 4;
  kpis.forEach(([k, v], i) => {
    const x = M + (i % 4) * (cw + 3);
    const yy = y + Math.floor(i / 4) * 19;
    doc.setDrawColor(...hexARgb(MARCA.linea));
    doc.roundedRect(x, yy, cw, 16, 2, 2, "S");
    doc.setFontSize(7);
    doc.setTextColor(...gris);
    doc.text(k.toUpperCase(), x + 3, yy + 5);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...tinta);
    doc.text(v, x + 3, yy + 12);
    doc.setFont("helvetica", "normal");
  });
  y += 42;

  // Gráfico de ventas por día
  if (r.dias.length) {
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.text("Ventas por día contra lo necesario", M, y);
    doc.setFont("helvetica", "normal");
    const gx = M + 12, gy = y + 4, gw = W - 2 * M - 12, gh = 42;
    const crudo = Math.max(isFinite(r.peDia) ? r.peDia : 0, ...r.dias.map((d) => d.venta)) * 1.1;
    const paso = crudo > 3e6 ? 1e6 : crudo > 1.5e6 ? 5e5 : 25e4;
    const max = Math.ceil(crudo / paso) * paso;
    const n = r.diasApertura;
    const bw = gw / n;
    doc.setFontSize(6);
    doc.setTextColor(...gris);
    for (let v = 0; v <= max; v += paso) {
      const yy = gy + gh - (gh * v) / max;
      doc.setDrawColor(230, 230, 230);
      doc.line(gx, yy, gx + gw, yy);
      doc.text(corto(v).replace("$ ", ""), gx - 1, yy + 1, { align: "right" });
    }
    for (const d of r.dias) {
      const i = Number(d.fecha.slice(8, 10)) - 1;
      const h = (gh * d.venta) / max;
      doc.setFillColor(...hexARgb(d.cubrio ? MARCA.ok : MARCA.mal));
      doc.rect(gx + i * bw + bw * 0.15, gy + gh - h, bw * 0.7, h, "F");
    }
    for (let i = 0; i < n; i += 5) doc.text(String(i + 1), gx + i * bw + bw / 2, gy + gh + 4, { align: "center" });
    const yl = gy + gh - (gh * r.peDia) / max;
    doc.setDrawColor(...tinta);
    doc.setLineDashPattern([1.5, 1], 0);
    doc.line(gx, yl, gx + gw, yl);
    doc.setLineDashPattern([], 0);
    doc.setTextColor(...tinta);
    doc.text(`Necesario: ${corto(r.peDia)} por día`, gx + gw, yl - 1.5, { align: "right" });
    y = gy + gh + 10;
  }

  // Conclusiones
  const recs = recomendaciones(r);
  if (recs.length) {
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...tinta);
    doc.text("Conclusiones", M, y);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    y += 5;
    for (const t of recs) {
      const lines = doc.splitTextToSize(`•  ${t}`, W - 2 * M);
      doc.text(lines, M, y);
      y += lines.length * 4.2 + 1;
    }
    y += 3;
  }

  const estilo = {
    headStyles: { fillColor: prim, textColor: 255, fontSize: 8 },
    styles: { fontSize: 8, cellPadding: 1.6 },
    alternateRowStyles: { fillColor: hexARgb(MARCA.primarioClaro) },
    margin: { left: M, right: M },
  };
  const ultimaY = () => (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY;

  autoTable(doc, {
    ...estilo,
    startY: y,
    head: [["Gastos fijos del mes", "Importe"]],
    body: [...r.gastos.filter((g) => g.importe).map((g) => [g.concepto, pesos(g.importe)]), ["Total", pesos(r.fijos)]],
    columnStyles: { 1: { halign: "right" } },
  });
  autoTable(doc, {
    ...estilo,
    startY: ultimaY() + 6,
    head: [["Día de la semana", "Venta promedio", "Días", "Cubrieron", "Contra lo necesario"]],
    body: r.semana.map((d) => [d.nombre, pesos(d.promedio), d.dias, d.cubiertos, pct(d.vsNecesario, 0)]),
    columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } },
  });
  autoTable(doc, {
    ...estilo,
    startY: ultimaY() + 6,
    head: [["Producto", "Unidades", "Venta", "Precio prom.", "Costo u.", "Margen", "Aporte"]],
    body: r.productos.map((p) => [p.articulo + (p.conCosto ? "" : " *"), Math.round(p.unidades), pesos(p.venta), pesos(p.precioProm), pesos(p.costoU), pct(p.margen, 0), pesos(p.aporte)]),
    columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" }, 5: { halign: "right" }, 6: { halign: "right" } },
    didParseCell: (c) => {
      if (c.section === "body" && c.column.index === 5 && r.productos[c.row.index]?.revisar) c.cell.styles.textColor = hexARgb(MARCA.mal);
    },
  });
  autoTable(doc, {
    ...estilo,
    startY: ultimaY() + 6,
    head: [["Fecha", "Venta", "Necesario", "Diferencia", "Tickets"]],
    body: r.dias.map((d) => [fechaCorta(d.fecha), pesos(d.venta), pesos(r.peDia), pesos(d.diferencia), d.tickets || "–"]),
    columnStyles: { 1: { halign: "right" }, 2: { halign: "right" }, 3: { halign: "right" }, 4: { halign: "right" } },
    didParseCell: (c) => {
      if (c.section === "body" && c.column.index === 3) c.cell.styles.textColor = hexARgb(r.dias[c.row.index]?.cubrio ? MARCA.ok : MARCA.mal);
    },
  });

  const paginas = doc.getNumberOfPages();
  for (let i = 1; i <= paginas; i++) {
    doc.setPage(i);
    doc.setFontSize(7);
    doc.setTextColor(...gris);
    doc.text("* costo estimado: falta cargar el costo de receta.  Punto de equilibrio = gastos fijos ÷ margen de contribución.", M, 290);
    doc.text(`${i} / ${paginas}`, W - M, 290, { align: "right" });
  }
  descargar(doc.output("blob"), `punto_equilibrio_${slug(nombreLocal)}_${r.mes}.pdf`);
}

// ---------------------------------------------------------------- Excel
export async function reporteExcel(r: Resultado, nombreLocal: string) {
  const XLSX = await import("xlsx");
  const wb = XLSX.utils.book_new();
  const hoja = (filas: (string | number)[][], anchos: number[]) => {
    const ws = XLSX.utils.aoa_to_sheet(filas);
    ws["!cols"] = anchos.map((w) => ({ wch: w }));
    return ws;
  };
  XLSX.utils.book_append_sheet(wb, hoja([
    [`${nombreLocal} · ${nombreMes(r.mes)}`],
    [],
    ["Ventas del mes", Math.round(r.ventasMes)],
    ["Gastos fijos", Math.round(r.fijos)],
    ["Mercadería / insumos %", +(r.insumosPct * 100).toFixed(2)],
    ["Comisiones medios de pago % (promedio)", +r.comisionPct.toFixed(2)],
    ["Total variables %", +(r.variablesPct * 100).toFixed(2)],
    ["Margen de contribución %", +(r.mc * 100).toFixed(2)],
    ["Punto de equilibrio mensual", Math.round(r.peMes)],
    ["Punto de equilibrio diario", Math.round(r.peDia)],
    ["Días con ventas cargadas", r.dias.length],
    ["Proyección del mes", Math.round(r.proyeccion)],
    ["Resultado proyectado", Math.round(r.resultadoProyectado)],
    ["Margen de seguridad %", +(r.margenSeguridad * 100).toFixed(2)],
    [],
    ["Conclusiones"],
    ...recomendaciones(r).map((t) => [t]),
  ], [34, 18]), "Resumen");
  XLSX.utils.book_append_sheet(wb, hoja([
    ["Fecha", "Día", "Venta", "Necesario", "Diferencia", "Cubrió", "Tickets"],
    ...r.dias.map((d) => [d.fecha, fechaCorta(d.fecha), Math.round(d.venta), Math.round(r.peDia), Math.round(d.diferencia), d.cubrio ? "Sí" : "No", d.tickets]),
  ], [12, 10, 14, 14, 14, 8, 8]), "Por día");
  XLSX.utils.book_append_sheet(wb, hoja([
    ["Producto", "Unidades", "Venta", "Precio promedio", "Costo unitario", "Costo total", "Margen %", "Aporte", "% de la venta", "Costo cargado"],
    ...r.productos.map((p) => [p.articulo, Math.round(p.unidades), Math.round(p.venta), Math.round(p.precioProm), Math.round(p.costoU), Math.round(p.costo), +(p.margen * 100).toFixed(1), Math.round(p.aporte), +(p.participacion * 100).toFixed(1), p.conCosto ? "Sí" : "Estimado"]),
  ], [26, 10, 14, 14, 14, 14, 10, 14, 12, 12]), "Por producto");
  XLSX.utils.book_append_sheet(wb, hoja([
    ["Día", "Venta promedio", "Días", "Cubrieron", "Contra lo necesario %"],
    ...r.semana.map((d) => [d.nombre, Math.round(d.promedio), d.dias, d.cubiertos, +(d.vsNecesario * 100).toFixed(1)]),
  ], [12, 16, 8, 10, 20]), "Por día de semana");
  if (r.medios.length)
    XLSX.utils.book_append_sheet(wb, hoja([
      ["Medio de pago", "Venta", "% de la venta", "Comisión %", "Costo comisión"],
      ...r.medios.map((m) => [m.medio, Math.round(m.venta), +(m.participacion * 100).toFixed(1), m.tasa, Math.round(m.comision)]),
    ], [20, 14, 14, 12, 16]), "Medios de pago");
  XLSX.utils.book_append_sheet(wb, hoja([["Concepto", "Importe"], ...r.gastos.map((g) => [g.concepto, g.importe]), ["Total", r.fijos]], [32, 14]), "Gastos fijos");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  descargar(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `punto_equilibrio_${slug(nombreLocal)}_${r.mes}.xlsx`);
}

/** Planilla para completar el costo de receta de cada producto y volver a subirla. */
export async function planillaCostos(r: Resultado, costos: Record<string, number>, nombreLocal: string) {
  const XLSX = await import("xlsx");
  const { norm } = await import("./tipos");
  const filas: (string | number)[][] = [
    ["Producto", "Precio promedio de venta", "Unidades vendidas en el mes", "Costo"],
    ...[...r.productos]
      .sort((a, b) => b.venta - a.venta)
      .map((p) => [p.articulo, Math.round(p.precioProm), Math.round(p.unidades), costos[norm(p.articulo)] ?? ""]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(filas);
  ws["!cols"] = [{ wch: 32 }, { wch: 22 }, { wch: 24 }, { wch: 12 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Costos");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  descargar(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `costos_${slug(nombreLocal)}.xlsx`);
}
