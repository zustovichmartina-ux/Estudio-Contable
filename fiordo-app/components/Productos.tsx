"use client";
import { useState } from "react";
import { Resultado } from "@/lib/motor";
import { pct, pesos } from "@/lib/formato";

type Props = { r: Resultado };

export default function Productos({ r }: Props) {
  const [busca, setBusca] = useState("");
  if (!r.hayDatos) return <div className="card vacio"><h2>Sin productos todavía</h2><p className="muted">Aparecen solos cuando cargás el Excel de ventas.</p></div>;

  const lista = r.productos.filter((p) => p.articulo.toLowerCase().includes(busca.toLowerCase()));

  return (
    <section className="card">
      <div className="card-head">
        <h2>Qué se vende más</h2>
        <span className="small muted">{r.productos.length} productos</span>
      </div>
      {r.diasSinDetalle > 0 && (
        <p className="aviso info small" style={{ marginBottom: 12 }}>
          {r.diasSinDetalle} {r.diasSinDetalle === 1 ? "día se cargó" : "días se cargaron"} con el reporte de ventas por día, que no trae productos: esas ventas no aparecen en esta lista. Subí el Detalle de Transacciones de esos días para verlas acá.
        </p>
      )}
      <div className="fila" style={{ marginBottom: 12 }}>
        <input className="input" style={{ maxWidth: 260 }} placeholder="Buscar producto" aria-label="Buscar producto" value={busca} onChange={(e) => setBusca(e.target.value)} />
      </div>
      <div className="tabla-wrap">
        <table>
          <thead>
            <tr><th>Producto</th><th className="r">Vendidos</th><th className="r">Precio promedio</th><th className="r">Venta</th><th className="r">% de la venta</th></tr>
          </thead>
          <tbody>
            {lista.map((p) => (
              <tr key={p.articulo}>
                <td>{p.articulo}</td>
                <td className="r num">{Math.round(p.unidades).toLocaleString("es-AR")}</td>
                <td className="r num">{pesos(p.precioProm)}</td>
                <td className="r num">{pesos(p.venta)}</td>
                <td className="r num">{pct(p.participacion, 1)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
