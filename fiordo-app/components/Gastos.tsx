"use client";
import { useEffect, useState } from "react";
import { Resultado } from "@/lib/motor";
import { Config, Gasto, Parametros, Proveedor } from "@/lib/tipos";
import { CONCEPTOS_FIJOS } from "@/lib/compras";
import { tasaMedio } from "@/lib/motor";
import { diasDelMes, nombreMes, pct, pesos } from "@/lib/formato";
import { Modo } from "@/lib/almacen";
import { IcoBorrar, IcoMas } from "./Iconos";

type Props = {
  r: Resultado;
  config: Config;
  modo: Modo;
  guardarGastos: (gastos: Gasto[], tambienBase: boolean) => void;
  guardarParametros: (p: Parametros, nombreLocal: string) => void;
  borrarMes: () => void;
  clasificarProveedor: (cuit: string, p: Proveedor) => void;
  irACargar: () => void;
};

export default function Gastos({ r, config, modo, guardarGastos, guardarParametros, borrarMes, clasificarProveedor, irACargar }: Props) {
  const manuales = r.gastos.filter((g) => !g.origen);
  const deArca = r.gastos.filter((g) => g.origen === "arca");
  const [gastos, setGastos] = useState<Gasto[]>(manuales);
  const [base, setBase] = useState(true);
  const [p, setP] = useState<Parametros>(config.parametros);
  const [nombre, setNombre] = useState(config.nombreLocal);
  const [confirmarBorrado, setConfirmarBorrado] = useState(false);
  const [guardado, setGuardado] = useState("");

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => setGastos(r.gastos.filter((g) => !g.origen)), [r.mes, r.gastos]);
  useEffect(() => { setP(config.parametros); setNombre(config.nombreLocal); }, [config.parametros, config.nombreLocal]);

  const total = gastos.reduce((s, g) => s + (Number(g.importe) || 0), 0);
  // medios con ventas del mes + los habituales aunque no tengan ventas
  const mediosTabla = [
    ...r.medios,
    ...["Efectivo", "Tarjetas", "QR", "Online"].filter((m) => !r.medios.some((x) => x.medio === m)).map((medio) => ({ medio, venta: 0, participacion: 0 })),
  ];
  const totalMedios = r.medios.reduce((s, m) => s + m.venta, 0);
  const comisionPrev = totalMedios ? r.medios.reduce((s, m) => s + m.venta * tasaMedio(m.medio, p), 0) / totalMedios : p.comisionPct;
  const avisar = (t: string) => { setGuardado(t); setTimeout(() => setGuardado(""), 2500); };

  return (
    <>
      <div className="grid-2">
        <section className="card">
          <div className="card-head"><h2>Gastos fijos de {nombreMes(r.mes)}</h2><span className="small muted">Se pagan vendas o no</span></div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {gastos.map((g, i) => (
              <div key={i} className="fila" style={{ flexWrap: "nowrap", gap: 8 }}>
                <input className="input" aria-label="Concepto" value={g.concepto} onChange={(e) => setGastos(gastos.map((x, j) => (j === i ? { ...x, concepto: e.target.value } : x)))} />
                <input className="input r" aria-label={`Importe de ${g.concepto}`} style={{ maxWidth: 150 }} type="number" inputMode="numeric" min={0} value={g.importe || ""} placeholder="0" onChange={(e) => setGastos(gastos.map((x, j) => (j === i ? { ...x, importe: Number(e.target.value) || 0 } : x)))} />
                <button className="btn sec" style={{ minHeight: 40, padding: "0 10px", borderColor: "var(--line)", color: "var(--muted)" }} aria-label={`Quitar ${g.concepto}`} onClick={() => setGastos(gastos.filter((_, j) => j !== i))}><IcoBorrar /></button>
              </div>
            ))}
          </div>
          <div className="fila" style={{ justifyContent: "space-between", marginTop: 12 }}>
            <button className="btn sec" onClick={() => setGastos([...gastos, { concepto: "", importe: 0 }])}><IcoMas /> Agregar gasto</button>
            <strong className="num">{deArca.length ? "Cargados a mano" : "Total"}: {pesos(total)}</strong>
          </div>
          {deArca.length > 0 && (
            <div className="arca-fijos">
              <span className="eyebrow">Facturados en ARCA (sin IVA)</span>
              {deArca.map((g) => (
                <div key={g.concepto} className="fila" style={{ justifyContent: "space-between" }}>
                  <span>{g.concepto}</span><span className="num">{pesos(g.importe)}</span>
                </div>
              ))}
              <div className="fila" style={{ justifyContent: "space-between", borderTop: "1px solid var(--line)", paddingTop: 8 }}>
                <strong>Total gastos fijos</strong><strong className="num">{pesos(total + r.compras.fijos)}</strong>
              </div>
              <span className="small muted">No cargues a mano lo que ya viene facturado en ARCA: se sumaría dos veces.</span>
            </div>
          )}
          <label className="fila small" style={{ gap: 6, marginTop: 12 }}>
            <input type="checkbox" checked={base} onChange={(e) => setBase(e.target.checked)} /> Usar estos gastos también para los meses siguientes
          </label>
          <div className="fila" style={{ marginTop: 12 }}>
            <button className="btn" onClick={() => { guardarGastos(gastos.filter((g) => g.concepto.trim() || g.importe), base); avisar("Gastos guardados"); }}>Guardar gastos</button>
            {guardado && <span className="small" style={{ color: "var(--ok)" }} role="status">{guardado}</span>}
          </div>
        </section>

        <section className="card">
          <div className="card-head"><h2>Gastos variables y datos del local</h2></div>
          <p className="small muted" style={{ marginBottom: 12 }}>
            Mercadería este mes: <strong>{pct(r.insumosPct)}</strong> de la venta, {r.mercaderiaSegun === "compras" ? "según las compras de ARCA" : "según el costo de cada producto"}. Acá van los demás gastos que suben con cada venta.
          </p>
          <div className="kpis" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))" }}>
            <div className="campo"><label htmlFor="p-iibb">Ingresos brutos (%)</label><input id="p-iibb" className="input r" type="number" step="0.1" inputMode="decimal" value={p.iibbPct} onChange={(e) => setP({ ...p, iibbPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-dias">Días que cierra en el mes</label><input id="p-dias" className="input r" type="number" min={0} max={30} inputMode="numeric" value={p.diasCerrados} onChange={(e) => setP({ ...p, diasCerrados: Math.max(0, Math.min(30, Number(e.target.value) || 0)) })} /></div>
            <div className="campo"><label htmlFor="p-def">Costo estimado si falta receta (%)</label><input id="p-def" className="input r" type="number" step="1" inputMode="decimal" value={p.costoDefectoPct} onChange={(e) => setP({ ...p, costoDefectoPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-min">Margen mínimo aceptable (%)</label><input id="p-min" className="input r" type="number" step="1" inputMode="decimal" value={p.margenMinimoPct} onChange={(e) => setP({ ...p, margenMinimoPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo">
              <label htmlFor="p-merc">Costo de mercadería según</label>
              <select id="p-merc" className="input" value={p.mercaderiaSegun} onChange={(e) => setP({ ...p, mercaderiaSegun: e.target.value as Parametros["mercaderiaSegun"] })}>
                <option value="auto">Automático (compras si hay)</option>
                <option value="compras">Compras de ARCA</option>
                <option value="recetas">Costo de recetas</option>
              </select>
            </div>
            <div className="campo"><label htmlFor="p-nombre">Nombre del local</label><input id="p-nombre" className="input" value={nombre} onChange={(e) => setNombre(e.target.value)} /></div>
          </div>
          <h3 style={{ marginTop: 16 }}>Comisiones por medio de pago</h3>
          <p className="small muted" style={{ margin: "4px 0 8px" }}>
            Todas las ventas cuentan igual, se cobren como se cobren. La comisión se calcula sobre lo vendido con cada medio, según los reportes de Bistrosoft.
          </p>
          <div className="tabla-wrap">
            <table>
              <thead><tr><th>Medio</th><th className="r">Ventas</th><th className="r">Comisión %</th><th className="r">Costo</th></tr></thead>
              <tbody>
                {mediosTabla.map((m) => (
                  <tr key={m.medio}>
                    <td>{m.medio}</td>
                    <td className="r num">{m.venta ? pct(m.participacion, 0) : "–"}</td>
                    <td className="r">
                      <input className="input r" style={{ width: 80 }} type="number" step="0.1" min={0} inputMode="decimal" aria-label={`Comisión de ${m.medio}`}
                        value={tasaMedio(m.medio, p)}
                        onChange={(e) => setP({ ...p, comisionesPorMedio: { ...(p.comisionesPorMedio ?? {}), [m.medio]: Number(e.target.value) || 0 } })} />
                    </td>
                    <td className="r num">{m.venta ? pesos((m.venta * tasaMedio(m.medio, p)) / 100) : "–"}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot><tr><td>Promedio sobre las ventas</td><td></td><td className="r num">{pct(comisionPrev / 100)}</td><td></td></tr></tfoot>
            </table>
          </div>
          <label className="fila small" style={{ gap: 6, marginTop: 10 }}>
            <input type="checkbox" checked={p.ivaComputable} onChange={(e) => setP({ ...p, ivaComputable: e.target.checked })} /> Es Responsable Inscripto: el IVA de las compras no es costo (aplica a las compras que se suban desde ahora)
          </label>
          <div className="fila" style={{ marginTop: 12 }}>
            <button className="btn" onClick={() => { guardarParametros(p, nombre); avisar("Datos guardados"); }}>Guardar</button>
            <span className="small muted">Abre {r.diasApertura} de {diasDelMes(r.mes)} días en {nombreMes(r.mes).split(" ")[0].toLowerCase()}.</span>
          </div>
          <div className="formula aviso info" style={{ marginTop: 14 }}>
            <strong>Punto de equilibrio</strong> = gastos fijos ÷ margen de contribución<br />
            {pesos(r.fijos)} ÷ {pct(r.mc)} = <strong>{isFinite(r.peMes) ? pesos(r.peMes) : "–"}</strong> por mes
          </div>
        </section>
      </div>

      <section className="card">
        <div className="card-head">
          <h2>Compras de ARCA de {nombreMes(r.mes)}</h2>
          <button className="btn sec" style={{ minHeight: 36 }} onClick={irACargar}>Subir compras</button>
        </div>
        {r.compras.cantidad === 0 ? (
          <p className="muted">Todavía no hay compras de este mes. Bajá <strong>Mis Comprobantes Recibidos</strong> de ARCA y subilo en Cargar: la mercadería y los gastos facturados se calculan solos.</p>
        ) : (
          <>
            <div className="kpis" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", marginBottom: 14 }}>
              <div className="kpi"><span className="eyebrow">Mercadería</span><span className="val num">{pesos(r.compras.mercaderia)}</span><span className="hint">{r.hayDatos ? `${pct(r.compras.mercaderia / r.ventasMes)} de la venta` : "Sin ventas cargadas"}</span></div>
              <div className="kpi"><span className="eyebrow">Gastos fijos</span><span className="val num">{pesos(r.compras.fijos)}</span><span className="hint">Se suman a los gastos del mes</span></div>
              <div className="kpi"><span className="eyebrow">No incluido</span><span className="val num">{pesos(r.compras.ignorado)}</span><span className="hint">{r.compras.cantidad} facturas en total, sin IVA</span></div>
            </div>
            <p className="small muted" style={{ marginBottom: 10 }}>Revisá cada proveedor una sola vez: la app lo recuerda para los meses siguientes. Los marcados con • son sugerencias automáticas.</p>
            <div className="tabla-wrap">
              <table>
                <thead><tr><th>Proveedor</th><th className="r">Fact.</th><th className="r">Sin IVA</th><th>Es</th><th>Rubro</th></tr></thead>
                <tbody>
                  {r.compras.proveedores.map((pv) => (
                    <tr key={pv.cuit}>
                      <td style={{ whiteSpace: "normal", minWidth: 160 }}>{pv.sugerido && <span title="Sugerencia automática" style={{ color: "var(--accent)" }}>• </span>}{pv.nombre}<br /><span className="small muted">CUIT {pv.cuit}</span></td>
                      <td className="r num">{pv.comprobantes}</td>
                      <td className="r num">{pesos(pv.costo)}</td>
                      <td>
                        <select className="input" aria-label={`Clasificación de ${pv.nombre}`} value={pv.categoria} onChange={(e) => {
                          const categoria = e.target.value as Proveedor["categoria"];
                          clasificarProveedor(pv.cuit, { nombre: pv.nombre, categoria, concepto: categoria === "fijo" ? pv.concepto || "Otros gastos fijos" : "" });
                        }}>
                          <option value="mercaderia">Mercadería</option>
                          <option value="fijo">Gasto fijo</option>
                          <option value="ignorar">No incluir</option>
                        </select>
                      </td>
                      <td>
                        {pv.categoria === "fijo" ? (
                          <select className="input" aria-label={`Rubro de ${pv.nombre}`} value={pv.concepto || "Otros gastos fijos"} onChange={(e) => clasificarProveedor(pv.cuit, { nombre: pv.nombre, categoria: "fijo", concepto: e.target.value })}>
                            {CONCEPTOS_FIJOS.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        ) : <span className="muted">–</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </section>

      <section className="card">
        <div className="card-head"><h2>Dónde se guardan los datos</h2></div>
        {modo === "nube" ? (
          <p>En la base de datos de la app: lo que cargues desde la compu se ve en el celular y al revés.</p>
        ) : (
          <p>Solo en este navegador. Para ver lo mismo en el celular y en la compu, hay que conectar la base de datos en Vercel (ver README).</p>
        )}
        {r.hayDatos && (
          <div className="fila" style={{ marginTop: 12 }}>
            {!confirmarBorrado ? (
              <button className="btn peligro" onClick={() => setConfirmarBorrado(true)}><IcoBorrar /> Borrar ventas de {nombreMes(r.mes)}</button>
            ) : (
              <>
                <span className="small">¿Seguro? Se borran todas las ventas cargadas de {nombreMes(r.mes)}.</span>
                <button className="btn peligro" onClick={() => { borrarMes(); setConfirmarBorrado(false); }}>Sí, borrar</button>
                <button className="btn sec" onClick={() => setConfirmarBorrado(false)}>Cancelar</button>
              </>
            )}
          </div>
        )}
      </section>
    </>
  );
}
