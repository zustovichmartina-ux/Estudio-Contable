"use client";
import { useMemo, useState } from "react";
import { CAMPOS, Campo, Hoja, convertir, leerArchivo } from "@/lib/lector";
import { LineaVenta, MesVentas } from "@/lib/tipos";
import { fechaCorta, nombreMes, pesos } from "@/lib/formato";
import { IcoCargar } from "./Iconos";

type Props = { meses: Record<string, MesVentas>; guardar: (lineas: LineaVenta[], archivo: string) => Promise<void> };

export default function Cargar({ meses, guardar }: Props) {
  const [archivo, setArchivo] = useState<string>("");
  const [hojas, setHojas] = useState<Hoja[]>([]);
  const [iHoja, setIHoja] = useState(0);
  const [fechaFija, setFechaFija] = useState("");
  const [estado, setEstado] = useState<{ tipo: "ok" | "bad" | "info"; texto: string } | null>(null);
  const [sobre, setSobre] = useState(false);
  const [guardando, setGuardando] = useState(false);

  const hoja = hojas[iHoja];
  const lectura = useMemo(() => (hoja ? convertir(hoja, fechaFija || undefined) : null), [hoja, fechaFija]);
  const total = lectura?.lineas.reduce((s, l) => s + l.importe, 0) ?? 0;

  async function abrir(f: File | undefined) {
    if (!f) return;
    setEstado({ tipo: "info", texto: `Leyendo ${f.name}…` });
    try {
      const hs = await leerArchivo(f);
      if (!hs.length) throw new Error("El archivo no tiene datos.");
      setArchivo(f.name);
      setHojas(hs);
      setIHoja(0);
      setEstado(null);
    } catch (e) {
      setHojas([]);
      setEstado({ tipo: "bad", texto: `No pude leer ${f.name}. Probá exportarlo de nuevo como Excel (.xlsx) o CSV. (${(e as Error).message})` });
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
      await guardar(lectura.lineas, archivo);
      setEstado({ tipo: "ok", texto: `Listo: se guardaron ${lectura.lineas.length.toLocaleString("es-AR")} líneas por ${pesos(total)}, del ${fechaCorta(lectura.desde!)} al ${fechaCorta(lectura.hasta!)}. Si esos días ya estaban cargados, se reemplazaron.` });
      setHojas([]);
    } catch (e) {
      setEstado({ tipo: "bad", texto: (e as Error).message });
    } finally {
      setGuardando(false);
    }
  }

  const faltaFecha = hoja && hoja.mapeo.fecha < 0 && !fechaFija;
  const faltaImporte = hoja && hoja.mapeo.importe < 0;

  return (
    <>
      <section className="card">
        <div className="card-head"><h2>Cargar ventas desde Excel</h2></div>
        <label
          className={`drop${sobre ? " sobre" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setSobre(true); }}
          onDragLeave={() => setSobre(false)}
          onDrop={(e) => { e.preventDefault(); setSobre(false); abrir(e.dataTransfer.files[0]); }}
        >
          <IcoCargar />
          <strong>Elegí el archivo de ventas</strong>
          <span className="small muted">Excel (.xlsx, .xls) o CSV. En la compu también podés arrastrarlo acá.</span>
          <input id="archivo" type="file" accept=".xlsx,.xls,.csv,.txt" onChange={(e) => { abrir(e.target.files?.[0]); e.target.value = ""; }} />
        </label>
        {estado && <div className={`aviso ${estado.tipo}`} style={{ marginTop: 12 }} role="status">{estado.texto}</div>}
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
            {CAMPOS.map((c) => (
              <div className="campo" key={c.campo}>
                <label htmlFor={`map-${c.campo}`}>{c.etiqueta}{c.obligatorio ? " *" : ""}</label>
                <select id={`map-${c.campo}`} className="input" value={hoja.mapeo[c.campo]} onChange={(e) => cambiarMapeo(c.campo, +e.target.value)}>
                  <option value={-1}>— No está —</option>
                  {hoja.encabezados.map((h, i) => <option key={i} value={i}>{h}</option>)}
                </select>
              </div>
            ))}
            {hoja.mapeo.fecha < 0 && (
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
                  {lectura.descartadas > 0 && <> {lectura.descartadas === 1 ? "Se ignoró 1 fila" : `Se ignoraron ${lectura.descartadas} filas`} sin fecha o importe (títulos, totales).</>}</>
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
          <div className="card-head"><h2>Cómo sacar el Excel de Bistrosoft</h2></div>
          <ol style={{ margin: 0, paddingLeft: 20, display: "flex", flexDirection: "column", gap: 6 }}>
            <li>Entrá a Bistro Web (webaccess.bistrosoft.com) con el usuario del local.</li>
            <li>Andá a la sección de <strong>Reportes</strong> o <strong>Ventas</strong>.</li>
            <li>Elegí el reporte con más detalle: <strong>ventas por artículo</strong> o <strong>detalle de ventas / tickets</strong>.</li>
            <li>Poné el rango de fechas (por ejemplo, del 1 al último día del mes).</li>
            <li>Tocá <strong>Exportar → Excel</strong> y subí ese archivo acá.</li>
          </ol>
          <p className="small muted" style={{ marginTop: 10 }}>Lo ideal es que tenga una fila por artículo vendido con fecha, cantidad e importe. Si además trae medio de pago y número de ticket, la app los usa.</p>
        </section>
        <section className="card">
          <div className="card-head"><h2>Meses cargados</h2></div>
          {Object.keys(meses).length === 0 ? (
            <p className="muted">Todavía no subiste ventas.</p>
          ) : (
            <div className="tabla-wrap">
              <table>
                <thead><tr><th>Mes</th><th className="r">Días</th><th className="r">Venta</th><th>Archivos</th></tr></thead>
                <tbody>
                  {Object.values(meses).sort((a, b) => b.mes.localeCompare(a.mes)).map((m) => (
                    <tr key={m.mes}>
                      <td>{nombreMes(m.mes)}</td>
                      <td className="r num">{new Set(m.items.map((i) => i.fecha)).size}</td>
                      <td className="r num">{pesos(m.items.reduce((s, i) => s + i.importe, 0))}</td>
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
