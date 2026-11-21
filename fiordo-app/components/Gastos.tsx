"use client";
import { useEffect, useState } from "react";
import { Resultado } from "@/lib/motor";
import { Config, Gasto, Parametros } from "@/lib/tipos";
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
};

export default function Gastos({ r, config, modo, guardarGastos, guardarParametros, borrarMes }: Props) {
  const [gastos, setGastos] = useState<Gasto[]>(r.gastos);
  const [base, setBase] = useState(true);
  const [p, setP] = useState<Parametros>(config.parametros);
  const [nombre, setNombre] = useState(config.nombreLocal);
  const [confirmarBorrado, setConfirmarBorrado] = useState(false);
  const [guardado, setGuardado] = useState("");

  useEffect(() => setGastos(r.gastos), [r.mes, r.gastos]);
  useEffect(() => { setP(config.parametros); setNombre(config.nombreLocal); }, [config.parametros, config.nombreLocal]);

  const total = gastos.reduce((s, g) => s + (Number(g.importe) || 0), 0);
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
            <strong className="num">Total: {pesos(total)}</strong>
          </div>
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
          <p className="small muted" style={{ marginBottom: 12 }}>La mercadería sale sola del costo de cada producto ({pct(r.insumosPct)} este mes). Acá van los demás gastos que suben con cada venta.</p>
          <div className="kpis" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))" }}>
            <div className="campo"><label htmlFor="p-com">Comisiones tarjetas y QR (%)</label><input id="p-com" className="input r" type="number" step="0.1" inputMode="decimal" value={p.comisionPct} onChange={(e) => setP({ ...p, comisionPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-iibb">Ingresos brutos (%)</label><input id="p-iibb" className="input r" type="number" step="0.1" inputMode="decimal" value={p.iibbPct} onChange={(e) => setP({ ...p, iibbPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-dias">Días que abre en el mes</label><input id="p-dias" className="input r" type="number" min={1} max={31} inputMode="numeric" value={p.diasApertura} onChange={(e) => setP({ ...p, diasApertura: Math.max(1, Math.min(31, Number(e.target.value) || 1)) })} /></div>
            <div className="campo"><label htmlFor="p-def">Costo estimado si falta receta (%)</label><input id="p-def" className="input r" type="number" step="1" inputMode="decimal" value={p.costoDefectoPct} onChange={(e) => setP({ ...p, costoDefectoPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-min">Margen mínimo aceptable (%)</label><input id="p-min" className="input r" type="number" step="1" inputMode="decimal" value={p.margenMinimoPct} onChange={(e) => setP({ ...p, margenMinimoPct: Number(e.target.value) || 0 })} /></div>
            <div className="campo"><label htmlFor="p-nombre">Nombre del local</label><input id="p-nombre" className="input" value={nombre} onChange={(e) => setNombre(e.target.value)} /></div>
          </div>
          <div className="fila" style={{ marginTop: 12 }}>
            <button className="btn" onClick={() => { guardarParametros(p, nombre); avisar("Datos guardados"); }}>Guardar</button>
            <button className="btn sec" onClick={() => setP({ ...p, diasApertura: diasDelMes(r.mes) })}>Abre todos los días ({diasDelMes(r.mes)})</button>
          </div>
          <div className="formula aviso info" style={{ marginTop: 14 }}>
            <strong>Punto de equilibrio</strong> = gastos fijos ÷ margen de contribución<br />
            {pesos(r.fijos)} ÷ {pct(r.mc)} = <strong>{isFinite(r.peMes) ? pesos(r.peMes) : "–"}</strong> por mes
          </div>
        </section>
      </div>

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
