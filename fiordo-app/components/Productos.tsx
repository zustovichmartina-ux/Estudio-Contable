"use client";
import { useState } from "react";
import { Resultado } from "@/lib/motor";
import { corto, pct, pesos } from "@/lib/formato";
import { norm } from "@/lib/tipos";

type Props = { r: Resultado; costos: Record<string, number>; margenMinimo: number; cambiarCosto: (articulo: string, costo: number | null) => void };

export default function Productos({ r, costos, margenMinimo, cambiarCosto }: Props) {
  const [soloSinCosto, setSoloSinCosto] = useState(false);
  const [busca, setBusca] = useState("");
  if (!r.hayDatos) return <div className="card vacio"><h2>Sin productos todavía</h2><p className="muted">Aparecen solos cuando cargás el Excel de ventas.</p></div>;

  const lista = r.productos.filter((p) => (!soloSinCosto || !p.conCosto) && p.articulo.toLowerCase().includes(busca.toLowerCase()));

  return (
    <>
      <section className="card">
        <div className="card-head">
          <h2>Qué deja cada producto</h2>
          <span className="small muted">Mercadería promedio: {pct(r.insumosPct)} de la venta</span>
        </div>
        <p className="muted small" style={{ marginBottom: 12 }}>
          Cargá el <strong>costo de receta</strong> de cada producto: lo que cuestan los insumos de una unidad. Los productos sin costo se estiman y se marcan con un asterisco. Por debajo de {margenMinimo} % de margen se marcan para revisar el precio.
        </p>
        <div className="fila" style={{ marginBottom: 12 }}>
          <input className="input" style={{ maxWidth: 260 }} placeholder="Buscar producto" aria-label="Buscar producto" value={busca} onChange={(e) => setBusca(e.target.value)} />
          <label className="fila small" style={{ gap: 6 }}>
            <input type="checkbox" checked={soloSinCosto} onChange={(e) => setSoloSinCosto(e.target.checked)} /> Solo sin costo ({r.sinCosto.length})
          </label>
        </div>
        <div className="tabla-wrap">
          <table>
            <thead>
              <tr><th>Producto</th><th className="r">Vendidos</th><th className="r">Precio</th><th className="r">Costo unitario</th><th className="r">Margen</th><th className="r">Aporte</th></tr>
            </thead>
            <tbody>
              {lista.map((p) => {
                const clave = norm(p.articulo);
                const cargado = costos[clave];
                return (
                  <tr key={p.articulo}>
                    <td>{p.articulo}{!p.conCosto && <span className="muted"> *</span>}</td>
                    <td className="r num">{Math.round(p.unidades).toLocaleString("es-AR")}</td>
                    <td className="r num">{pesos(p.precioProm)}</td>
                    <td className="r">
                      <input
                        key={`${clave}-${cargado ?? "x"}`}
                        className="input r"
                        style={{ width: 110 }}
                        type="number"
                        inputMode="decimal"
                        min={0}
                        aria-label={`Costo de ${p.articulo}`}
                        placeholder={String(Math.round(p.costoU))}
                        defaultValue={cargado ?? ""}
                        onBlur={(e) => {
                          const v = e.target.value.trim();
                          const n = v === "" ? null : Number(v);
                          if (n !== (cargado ?? null)) cambiarCosto(p.articulo, n);
                        }}
                        onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }}
                      />
                    </td>
                    <td className="r"><span className={`pill ${p.revisar ? "bad" : p.margen < 0.65 ? "warn" : "ok"}`}>{pct(p.margen, 0)}</span></td>
                    <td className="r num">{corto(p.aporte)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <p className="small muted" style={{ marginTop: 10 }}>El aporte es lo que deja cada producto para pagar los gastos fijos, después de insumos, comisiones e ingresos brutos.</p>
      </section>
    </>
  );
}
