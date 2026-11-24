"use client";
import { Resultado, estadoResultados } from "@/lib/motor";
import { corto, fechaCorta, nombreMes, pct, pesos } from "@/lib/formato";
import Grafico from "./Grafico";

type Props = { r: Resultado; iibbPct: number; irA: (v: "cargar" | "productos" | "gastos") => void; probarEjemplo: () => void };

function EstadoMes({ r, iibbPct }: { r: Resultado; iibbPct: number }) {
  const e = estadoResultados(r, iibbPct);
  const gana = e.resultado >= 0;
  const fila = (concepto: string, importe: number, opc: { menos?: boolean; sub?: boolean; total?: boolean } = {}) => (
    <tr key={concepto} className={opc.total ? "er-total" : opc.sub ? "er-sub" : undefined}>
      <td>{concepto}</td>
      <td className="r num">{opc.menos && importe ? "– " : ""}{pesos(Math.abs(importe))}</td>
      <td className="r num muted">{e.ventas ? pct(importe / e.ventas, 1) : "–"}</td>
    </tr>
  );
  return (
    <section className="card">
      <div className="card-head">
        <h2>Resultado de {nombreMes(r.mes).toLowerCase()}</h2>
        <span className="small muted">{e.proyectado ? `Proyectado al cierre (${r.dias.length} de ${r.diasApertura} días cargados)` : "Mes completo"}</span>
      </div>
      <div className="tabla-wrap">
        <table className="er">
          <thead><tr><th>Concepto</th><th className="r">Importe</th><th className="r">% venta</th></tr></thead>
          <tbody>
            {fila(e.proyectado ? "Ingresos Bistrosoft (proyectados al cierre)" : "Ingresos Bistrosoft", e.ventas, { sub: true })}
            {e.gastos.map((g) => fila(g.concepto, g.importe, { menos: true }))}
            {fila("Total gastos", e.totalGastos, { menos: true, sub: true })}
          </tbody>
          <tfoot>
            <tr className={`er-total ${gana ? "ok" : "bad"}`}>
              <td>{gana ? "Ganancia" : "Pérdida"}</td>
              <td className="r num">{pesos(e.resultado)}</td>
              <td className="r num">{pct(e.margen, 1)}</td>
            </tr>
          </tfoot>
        </table>
      </div>
      <div className="er-veredicto">
        <span className={`pill ${gana ? "ok" : "bad"}`}>{gana ? "El ingreso cubre los gastos" : "El ingreso no cubre los gastos"}</span>
        <span className="small">Margen {pct(e.margen, 1)}</span>
      </div>
      {r.fijos > 0 && isFinite(r.peMes) && (
        <p className="er-equilibrio">
          Para cubrir todos los gastos hay que vender <strong>{pesos(r.peMes)}</strong> en el mes (<strong>{corto(r.peDia)}</strong> por día).{" "}
          {e.llegaEquilibrio
            ? <>{e.proyectado ? "Al ritmo actual vas" : "Se vendió"} <strong style={{ color: "var(--ok)" }}>{pct(e.ventas / r.peMes - 1, 0)} por encima</strong>.</>
            : <>{e.proyectado ? "Al ritmo actual vas" : "Se vendió"} <strong style={{ color: "var(--bad)" }}>{pct(1 - e.ventas / r.peMes, 0)} por debajo</strong>.</>}
        </p>
      )}
      {r.fijos <= 0 && <p className="er-equilibrio">Cargá los gastos del mes (sueldos, alquiler, servicios) para ver cuánto hay que vender para cubrirlos.</p>}
    </section>
  );
}

export default function Hoy({ r, iibbPct, irA, probarEjemplo }: Props) {
  if (!r.hayDatos)
    return (
      <div className="card vacio">
        <h2>No hay ventas cargadas en {nombreMes(r.mes)}</h2>
        <p className="muted">Subí el Excel de ventas que exportás de Bistrosoft y acá vas a ver si el café cubre sus gastos día a día.</p>
        <div className="fila" style={{ justifyContent: "center" }}>
          <button className="btn" onClick={() => irA("cargar")}>Cargar Excel de ventas</button>
          <button className="btn sec" onClick={probarEjemplo}>Probar con datos de ejemplo</button>
        </div>
      </div>
    );

  const u = r.ultimo!;
  const sinGastos = r.fijos <= 0;
  const pasos: { titulo: string; detalle: string; hecho: boolean; ir: "cargar" | "productos" | "gastos"; boton: string }[] = [
    { titulo: "Subir las ventas", detalle: "El Detalle de Transacciones de Bistrosoft.", hecho: r.hayDatos, ir: "cargar", boton: "Subir" },
    { titulo: "Subir las compras", detalle: "El listado por imputación contable de Tango (o Mis Comprobantes Recibidos de ARCA).", hecho: !r.sinMercaderia, ir: "cargar", boton: "Subir" },
    { titulo: "Cargar sueldos y otros gastos", detalle: "Sueldos, cargas y amortizaciones a mano; lo facturado sale de las compras.", hecho: !sinGastos, ir: "gastos", boton: "Cargar" },
  ];
  const pendientes = pasos.filter((p) => !p.hecho).length;
  // días de atraso: solo si el mes que se mira es el actual
  const hoy = new Date();
  const esMesActual = r.mes === `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, "0")}`;
  const atraso = esMesActual ? Math.floor((hoy.getTime() - new Date(u.fecha + "T00:00:00").getTime()) / 86400000) : 0;
  const escala = Math.max(r.peMes, r.proyeccion, r.ventasMes) * 1.05 || 1;
  const veredicto = sinGastos
    ? "Cargá los gastos del mes para saber si este día cubrió los costos."
    : r.estadoDia === "ok" ? `Cubriste los gastos del día, con ${corto(u.diferencia)} de sobra.`
    : r.estadoDia === "warn" ? `Quedaste cerca: faltaron ${corto(-u.diferencia)} para cubrir el día.`
    : `No alcanzó: faltaron ${corto(-u.diferencia)} para cubrir el día.`;

  return (
    <>
      {pendientes > 0 && (
        <section className="card pasos" aria-label="Primeros pasos">
          <div className="card-head" style={{ marginBottom: 8 }}>
            <h2>Primeros pasos</h2>
            <span className="small muted">{3 - pendientes} de 3 listos</span>
          </div>
          <ul>
            {pasos.map((p) => (
              <li key={p.titulo} className={p.hecho ? "hecho" : ""}>
                <span className="check" aria-hidden="true">{p.hecho ? "✓" : ""}</span>
                <div>
                  <strong>{p.titulo}</strong>
                  <span className="small muted">{p.detalle}</span>
                </div>
                {!p.hecho && <button className="btn sec" onClick={() => irA(p.ir)}>{p.boton}</button>}
              </li>
            ))}
          </ul>
        </section>
      )}

      <div className="al-dia">
        <span>Ventas cargadas hasta el <strong>{fechaCorta(u.fecha)}</strong>{atraso > 1 && <span className="muted"> · faltan {atraso - 1} días</span>}</span>
        <button className="btn sec" onClick={() => irA("cargar")}>Actualizar ventas</button>
      </div>

      <div className="grid-2">
        <section className={`semaforo ${sinGastos ? "neutro" : r.estadoDia}`} aria-label="Resultado del último día">
          <div className="fila" style={{ gap: 12 }}>
            <span className="luz" aria-hidden="true" />
            <span className="eyebrow">Último día cargado · {fechaCorta(u.fecha)}</span>
          </div>
          <p className="veredicto">{veredicto}</p>
          <div className="cifras">
            <div><span className="eyebrow">Vendiste</span><span className="cifra num">{corto(u.venta)}</span></div>
            <div><span className="eyebrow">Necesitás por día</span><span className="cifra num">{sinGastos ? "–" : corto(r.peDia)}</span></div>
          </div>
        </section>

        <section className="card" aria-label="Avance del mes" style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <div>
            <span className="eyebrow">Vendido en {nombreMes(r.mes)}</span>
            <p className="num" style={{ fontFamily: "var(--font-display)", fontSize: 24, fontWeight: 700, lineHeight: 1.15 }}>{pesos(r.ventasMes)}</p>
          </div>
          <div className="barra" role="img" aria-label={`Vendido ${pesos(r.ventasMes)} de ${pesos(r.peMes)} necesarios`}>
            <div className="barra-fill" style={{ width: `${Math.min(100, (r.ventasMes / escala) * 100)}%` }} />
            {!sinGastos && isFinite(r.peMes) && <div className="barra-marca" style={{ left: `${Math.min(100, (r.peMes / escala) * 100)}%` }} />}
          </div>
          <div className="fila small muted" style={{ justifyContent: "space-between" }}>
            <span>{r.dias.length} de {r.diasApertura} días</span>
            <span>{sinGastos ? "Equilibrio: falta cargar gastos" : `Equilibrio: ${corto(r.peMes)}`}</span>
          </div>
          <p>
            {sinGastos ? (
              <>Promedio por día: <strong>{corto(r.promedioDia)}</strong>. Con los gastos del mes cargados vas a ver cuánto falta para cubrirlos.</>
            ) : r.falta <= 0 ? (
              <><strong>Ya cubriste los gastos del mes.</strong> Desde acá, de cada $ 100 que vendas quedan $ {Math.round(r.mc * 100)} de ganancia antes de impuestos.</>
            ) : r.diasRestantes > 0 ? (
              <>Te faltan <strong>{corto(r.falta)}</strong>. En los {r.diasRestantes} días que quedan necesitás vender <strong>{corto(r.necesarioRestante)} por día</strong>.</>
            ) : (
              <>El mes cerró <strong>{corto(r.falta)} por debajo</strong> del punto de equilibrio.</>
            )}
          </p>
        </section>
      </div>

      <section className="kpis" aria-label="Indicadores">
        <div className="kpi"><span className="eyebrow">Punto de equilibrio</span><span className="val num">{sinGastos ? "–" : corto(r.peMes)}</span><span className="hint">Lo que hay que vender para cubrir los gastos</span></div>
        <div className="kpi"><span className="eyebrow">Gastos del mes</span><span className="val num">{sinGastos ? "–" : corto(estadoResultados(r, iibbPct).totalGastos)}</span><span className="hint">Todos, al ritmo de venta actual</span></div>
        <div className="kpi"><span className="eyebrow">Proyección</span><span className="val num">{corto(r.proyeccion)}</span><span className="hint">{sinGastos ? "Al ritmo actual, mes completo" : `Resultado estimado: ${corto(r.resultadoProyectado)}`}</span></div>
        <div className="kpi">
          <span className="eyebrow">Margen de seguridad</span>
          <span className="val num" style={{ color: sinGastos ? undefined : r.margenSeguridad >= 0.15 ? "var(--ok)" : r.margenSeguridad >= 0 ? "var(--warn)" : "var(--bad)" }}>{sinGastos ? "–" : pct(r.margenSeguridad)}</span>
          <span className="hint">Cuánto pueden caer las ventas sin perder</span>
        </div>
      </section>

      <EstadoMes r={r} iibbPct={iibbPct} />

      <section className="card">
        <div className="card-head">
          <h2>Ventas por día</h2>
          {!sinGastos && <div className="leyenda">
            <span><i style={{ background: "var(--ok)" }} />Cubrió</span>
            <span><i style={{ background: "var(--bad)" }} />No llegó</span>
            <span><i style={{ background: "var(--ink)", height: 2 }} />Necesario</span>
          </div>}
        </div>
        <Grafico r={r} neutro={sinGastos} />
      </section>

      <div className="grid-2">
        <section className="card">
          <div className="card-head"><h2>Por día de la semana</h2><span className="small muted">Promedio de venta</span></div>
          <div className="tabla-wrap">
            <table>
              <thead><tr><th>Día</th><th className="r">Promedio</th><th className="r">Contra lo necesario</th></tr></thead>
              <tbody>
                {r.semana.map((d) => (
                  <tr key={d.nombre}>
                    <td>{d.nombre}</td>
                    <td className="r num">{corto(d.promedio)}</td>
                    <td className="r">{sinGastos ? <span className="muted">–</span> : <span className={`pill ${d.vsNecesario >= 0 ? "ok" : d.vsNecesario > -0.15 ? "warn" : "bad"}`}>{d.vsNecesario >= 0 ? "+" : ""}{pct(d.vsNecesario, 0)}</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
        <section className="card">
          <div className="card-head"><h2>Lo que más se vende</h2><button className="btn sec" style={{ minHeight: 34 }} onClick={() => irA("productos")}>Ver todos</button></div>
          <div className="tabla-wrap">
            <table>
              <thead><tr><th>Producto</th><th className="r">Venta</th><th className="r">% de la venta</th></tr></thead>
              <tbody>
                {r.productos.slice(0, 6).map((p) => (
                  <tr key={p.articulo}>
                    <td>{p.articulo}</td>
                    <td className="r num">{corto(p.venta)}</td>
                    <td className="r num">{pct(p.participacion, 0)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {r.medios.length > 0 && (
            <p className="small muted" style={{ marginTop: 10 }}>
              Medios de pago: {r.medios.slice(0, 4).map((m) => `${m.medio} ${pct(m.participacion, 0)}`).join(" · ")}
            </p>
          )}
        </section>
      </div>
    </>
  );
}
