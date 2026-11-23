"use client";
import { useMemo, useState } from "react";
import { CAMPOS, Campo, Hoja, convertir, leerArchivo } from "@/lib/lector";
import { esArcaCompras, leerComprasArca } from "@/lib/compras";
import { Compra, LineaVenta, MesVentas } from "@/lib/tipos";
import { fechaCorta, nombreMes, pesos } from "@/lib/formato";
import { IcoCargar } from "./Iconos";

export type Lote = { lineas: LineaVenta[]; archivo: string; soloCompletar?: boolean };
export type LoteCompras = { compras: Compra[]; archivo: string };
type Props = {
  meses: Record<string, MesVentas>;
  ivaComputable: boolean;
  guardar: (lotes: Lote[], compras: LoteCompras[]) => Promise<void>;
  deshacer: (() => Promise<void>) | null;
  verResultado: (mensaje: string, vista?: "hoy" | "gastos") => void;
};

export default function Cargar({ meses, ivaComputable, guardar, deshacer, verResultado }: Props) {
  const [archivo, setArchivo] = useState<string>("");
  const [hojas, setHojas] = useState<Hoja[]>([]);
  const [iHoja, setIHoja] = useState(0);
  const [fechaFija, setFechaFija] = useState("");
  const [estado, setEstado] = useState<{ tipo: "ok" | "bad" | "info"; texto: string } | null>(null);
  const [sobre, setSobre] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [listo, setListo] = useState(false);

  const hoja = hojas[iHoja];
  const lectura = useMemo(() => (hoja ? convertir(hoja, fechaFija || undefined) : null), [hoja, fechaFija]);
  const total = lectura?.lineas.reduce((s, l) => s + l.importe, 0) ?? 0;

  async function abrir(lista: FileList | File[] | null | undefined) {
    const archivos = Array.from(lista ?? []);
    if (!archivos.length) return;
    setListo(false);
    setHojas([]);
    setEstado({ tipo: "info", texto: archivos.length === 1 ? `Leyendo ${archivos[0].name}…` : `Leyendo ${archivos.length} archivos…` });
    const lotes: Lote[] = [];
    const lotesCompras: LoteCompras[] = [];
    const errores: string[] = [];
    let manual: { nombre: string; hojas: Hoja[] } | null = null;
    for (const f of archivos) {
      try {
        const hs = await leerArchivo(f);
        if (!hs.length) throw new Error("no tiene datos");
        const arca = hs.find(esArcaCompras);
        const reconocida = arca ? undefined : hs.find((h) => convertir(h).formato);
        if (arca) {
          const cs = leerComprasArca(arca, ivaComputable);
          if (cs.length) lotesCompras.push({ compras: cs, archivo: f.name });
          else errores.push(`${f.name}: no tiene comprobantes`);
        } else if (reconocida) {
          const r = convertir(reconocida);
          if (r.lineas.length) lotes.push({ lineas: r.lineas, archivo: f.name, soloCompletar: r.soloCompletar });
          else errores.push(`${f.name}: no tiene ventas`);
        } else if (!manual) manual = { nombre: f.name, hojas: hs };
        else errores.push(`${f.name}: formato no reconocido, subilo solo`);
      } catch (e) {
        errores.push(`${f.name}: no se pudo leer (${(e as Error).message})`);
      }
    }
    if (lotes.length || lotesCompras.length) {
      setGuardando(true);
      try {
        await guardar(lotes, lotesCompras);
        const partes: string[] = [];
        if (lotes.length) {
          const lineas = lotes.flatMap((l) => l.lineas);
          const fechas = lineas.map((l) => l.fecha).sort();
          const total = lineas.reduce((s, l) => s + l.importe, 0);
          partes.push(`Se leyeron ventas por ${pesos(total)} del ${fechaCorta(fechas[0])} al ${fechaCorta(fechas[fechas.length - 1])}.` +
            (lotes.some((l) => l.soloCompletar) ? " Del reporte de ventas por día solo se tomaron los días que no tenían el detalle por producto." : ""));
        }
        if (lotesCompras.length) {
          const cs = lotesCompras.flatMap((l) => l.compras);
          const fechas = cs.map((c) => c.fecha).sort();
          partes.push(`Se leyeron ${cs.length} facturas de compras de ARCA por ${pesos(cs.reduce((s, c) => s + c.costo, 0))} sin IVA, del ${fechaCorta(fechas[0])} al ${fechaCorta(fechas[fechas.length - 1])}. Revisá cómo quedó clasificado cada proveedor.`);
        }
        const texto = partes.join(" ") + (errores.length ? ` No se cargó: ${errores.join("; ")}.` : "");
        setEstado({ tipo: "ok", texto });
        setListo(true);
        if (!manual) verResultado(texto, lotesCompras.length && !lotes.length ? "gastos" : "hoy");
      } catch (e) {
        setEstado({ tipo: "bad", texto: (e as Error).message });
      } finally {
        setGuardando(false);
      }
    } else if (!manual) {
      setEstado({ tipo: "bad", texto: `No se pudo cargar: ${errores.join("; ")}. Probá exportarlo de nuevo como Excel (.xlsx).` });
    } else {
      setEstado(errores.length ? { tipo: "bad", texto: `No se cargó: ${errores.join("; ")}.` } : null);
    }
    if (manual) {
      setArchivo(manual.nombre);
      setHojas(manual.hojas);
      setIHoja(0);
    }
  }

  function cambiarHoja(cambio: Partial<Hoja>) {
    setHojas((hs) => hs.map((h, i) => (i === iHoja ? { ...h, ...cambio } : h)));
  }
  function cambiarMapeo(campo: Campo, col: number) {
    if (!hoja) return;
    cambiarHoja({ mapeo: { ...hoja.mapeo, [campo]: col } });
  }
  function cambiarEncabezado(n: number) {
    if (!hoja) return;
    const fila = Math.max(0, Math.min(hoja.filas.length - 1, n));
    const encabezados = (hoja.filas[fila] ?? []).map((c, i) => (c == null || c === "" ? `Columna ${i + 1}` : String(c).trim()));
    cambiarHoja({ filaEncabezado: fila, encabezados });
  }

  async function confirmar() {
    if (!lectura?.lineas.length) return;
    setGuardando(true);
    try {
      await guardar([{ lineas: lectura.lineas, archivo }], []);
      setListo(true);
      verResultado(`Se cargaron ventas por ${pesos(total)} del ${fechaCorta(lectura.desde!)} al ${fechaCorta(lectura.hasta!)}.`);
      setEstado({ tipo: "ok", texto: `Listo: se guardaron ${lectura.lineas.length.toLocaleString("es-AR")} líneas por ${pesos(total)}, del ${fechaCorta(lectura.desde!)} al ${fechaCorta(lectura.hasta!)}. Si esos días ya estaban cargados, se reemplazaron.` });
      setHojas([]);
    } catch (e) {
      setEstado({ tipo: "bad", texto: (e as Error).message });
    } finally {
      setGuardando(false);
    }
  }

  const faltaFecha = hoja && !lectura?.formato && hoja.mapeo.fecha < 0 && !fechaFija;
  const faltaImporte = hoja && !lectura?.formato && hoja.mapeo.importe < 0;

  return (
    <>
      <section className="card">
        <div className="card-head"><h2>Cargar ventas desde Excel</h2></div>
        <label
          className={`drop${sobre ? " sobre" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setSobre(true); }}
          onDragLeave={() => setSobre(false)}
          onDrop={(e) => { e.preventDefault(); setSobre(false); abrir(e.dataTransfer.files); }}
        >
          <IcoCargar />
          <strong>{guardando ? "Guardando…" : "Subí el Detalle de Transacciones de Bistrosoft"}</strong>
          <span className="small muted">También las compras de ARCA (Mis Comprobantes Recibidos). Se guarda solo y podés elegir varios archivos a la vez.</span>
          <input id="archivo" type="file" multiple accept=".xlsx,.xls,.csv,.txt" onChange={(e) => { abrir(e.target.files); e.target.value = ""; }} />
        </label>
        {estado && <div className={`aviso ${estado.tipo}`} style={{ marginTop: 12 }} role="status">{estado.texto}</div>}
        {listo && (
          <div className="fila" style={{ marginTop: 12 }}>
            <button className="btn" onClick={() => verResultado("")}>Ver cómo vamos</button>
            {deshacer && (
              <button className="btn sec" onClick={async () => { await deshacer(); setListo(false); setEstado({ tipo: "info", texto: "Se deshizo la última carga." }); }}>Deshacer</button>
            )}
          </div>
        )}
      </section>

      {hoja && lectura && (
        <section className="card" style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div className="card-head" style={{ marginBottom: 0 }}>
            <h2>Revisá cómo se leyó</h2>
            <span className="small muted">{archivo}</span>
          </div>

          <div className="kpis" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))" }}>
            {hojas.length > 1 && (
              <div className="campo">
                <label htmlFor="hoja">Hoja</label>
                <select id="hoja" className="input" value={iHoja} onChange={(e) => setIHoja(+e.target.value)}>
                  {hojas.map((h, i) => <option key={h.nombre} value={i}>{h.nombre}</option>)}
                </select>
              </div>
            )}
            <div className="campo">
              <label htmlFor="fila-enc">Fila de títulos</label>
              <input id="fila-enc" className="input" type="number" min={1} value={hoja.filaEncabezado + 1} onChange={(e) => cambiarEncabezado(+e.target.value - 1)} />
            </div>
            {!lectura.formato && CAMPOS.map((c) => (
              <div className="campo" key={c.campo}>
                <label htmlFor={`map-${c.campo}`}>{c.etiqueta}{c.obligatorio ? " *" : ""}</label>
                <select id={`map-${c.campo}`} className="input" value={hoja.mapeo[c.campo]} onChange={(e) => cambiarMapeo(c.campo, +e.target.value)}>
                  <option value={-1}>— No está —</option>
                  {hoja.encabezados.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
              </div>
            ))}
            {lectura.formato && (
              <div className="campo" style={{ gridColumn: "1 / -1" }}>
                <span className="pill ok" style={{ alignSelf: "flex-start", fontSize: 13, padding: "4px 12px" }}>Formato reconocido: {lectura.formato}</span>
                <span className="small muted">Se toman los artículos, combos y descuentos de cada ticket. Los movimientos de caja (apertura, retiros, cierre) no son ventas y se ignoran.</span>
              </div>
            )}
            {!lectura.formato && hoja.mapeo.fecha < 0 && (
              <div className="campo">
                <label htmlFor="fecha-fija">Fecha de todas las filas *</label>
                <input id="fecha-fija" className="input" type="date" value={fechaFija} onChange={(e) => setFechaFija(e.target.value)} />
              </div>
            )}
          </div>

          {faltaImporte && <div className="aviso bad">Elegí qué columna tiene el importe de cada venta.</div>}
          {faltaFecha && <div className="aviso bad">El archivo no tiene columna de fecha. Elegila arriba o indicá a qué día corresponden las ventas.</div>}

          {!faltaImporte && !faltaFecha && (
            <>
              <div className="aviso info">
                {lectura.lineas.length ? (
                  <>Se leyeron <strong>{lectura.lineas.length.toLocaleString("es-AR")} líneas</strong> por <strong>{pesos(total)}</strong>, del {fechaCorta(lectura.desde!)} al {fechaCorta(lectura.hasta!)}.
                  {lectura.descartadas > 0 && !lectura.formato && <> {lectura.descartadas === 1 ? "Se ignoró 1 fila" : `Se ignoraron ${lectura.descartadas} filas`} sin fecha o importe (títulos, totales).</>}</>
                ) : (
                  <>No encontré filas con fecha e importe válidos. Revisá la fila de títulos y las columnas elegidas.</>
                )}
              </div>
              {lectura.lineas.length > 0 && (
                <div className="tabla-wrap">
                  <table>
                    <thead><tr><th>Fecha</th><th>Artículo</th><th className="r">Cant.</th><th className="r">Importe</th><th>Medio</th></tr></thead>
                    <tbody>
                      {lectura.lineas.slice(0, 8).map((l, i) => (
                        <tr key={i}><td>{fechaCorta(l.fecha)}</td><td>{l.articulo}</td><td className="r num">{l.cantidad}</td><td className="r num">{pesos(l.importe)}</td><td>{l.medio ?? "–"}</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              <div className="fila">
                <button className="btn" disabled={!lectura.lineas.length || guardando} onClick={confirmar}>{guardando ? "Guardando…" : "Guardar estas ventas"}</button>
                <button className="btn sec" onClick={() => setHojas([])}>Cancelar</button>
              </div>
            </>
          )}
        </section>
      )}

      <div className="grid-2">
        <section className="card">
          <div className="card-head"><h2>Qué bajar de Bistrosoft</h2></div>
          <ol style={{ margin: 0, paddingLeft: 20, display: "flex", flexDirection: "column", gap: 6 }}>
            <li>Entrá a Bistro Web (webaccess.bistrosoft.com) con el usuario del local.</li>
            <li>Abrí el reporte <strong>Detalle de Transacciones</strong> (trae cada producto vendido).</li>
            <li>Elegí desde el <strong>día 1 del mes hasta hoy</strong>.</li>
            <li>Exportalo a <strong>Excel</strong> y subilo acá.</li>
          </ol>
          <p className="small muted" style={{ marginTop: 10 }}>
            También sirve el <strong>Reporte de Ventas por Día</strong>: es más liviano y alcanza para el semáforo y el punto de equilibrio, pero no dice qué productos se vendieron. Si un día ya tiene el detalle, el reporte por día no lo pisa.
            Podés subir los archivos todos los días o una vez por semana: los días repetidos se reemplazan, no se duplican.
          </p>
          <h3 style={{ marginTop: 14 }}>Compras de ARCA</h3>
          <ol style={{ margin: "6px 0 0", paddingLeft: 20, display: "flex", flexDirection: "column", gap: 6 }}>
            <li>Entrá a <strong>ARCA</strong> con la clave fiscal del café y abrí <strong>Mis Comprobantes</strong>.</li>
            <li>Elegí <strong>Recibidos</strong>, poné las fechas del mes y tocá <strong>Buscar</strong>.</li>
            <li>Exportá a <strong>Excel</strong> (o CSV) y subilo acá.</li>
          </ol>
        </section>
        <section className="card">
          <div className="card-head"><h2>Meses cargados</h2></div>
          {Object.keys(meses).length === 0 ? (
            <p className="muted">Todavía no subiste ventas.</p>
          ) : (
            <div className="tabla-wrap">
              <table>
                <thead><tr><th>Mes</th><th className="r">Días</th><th className="r">Venta</th><th className="r">Compras</th><th>Archivos</th></tr></thead>
                <tbody>
                  {Object.values(meses).sort((a, b) => b.mes.localeCompare(a.mes)).map((m) => (
                    <tr key={m.mes}>
                      <td>{nombreMes(m.mes)}</td>
                      <td className="r num">{new Set(m.items.map((i) => i.fecha)).size}</td>
                      <td className="r num">{pesos(m.items.reduce((s, i) => s + i.importe, 0))}</td>
                      <td className="r num">{m.compras?.length ? `${m.compras.length} fact.` : "–"}</td>
                      <td className="small muted" style={{ whiteSpace: "normal" }}>{m.archivos.join(", ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      </div>
    </>
  );
}
