"use client";
import { useState } from "react";
import { Resultado } from "@/lib/motor";
import { corto, pct, pesos } from "@/lib/formato";
import { norm } from "@/lib/tipos";
import { leerCostos } from "@/lib/lector";
import { planillaCostos } from "@/lib/reportes";
import { IcoCargar, IcoDescarga } from "./Iconos";

type Props = {
  r: Resultado;
  costos: Record<string, number>;
  margenMinimo: number;
  nombreLocal: string;
  cambiarCosto: (articulo: string, costo: number | null) => void;
  importarCostos: (lista: { articulo: string; costo: number }[]) => void;
};

export default function Productos({ r, costos, margenMinimo, nombreLocal, cambiarCosto, importarCostos }: Props) {
  const [msg, setMsg] = useState<{ tipo: "ok" | "bad"; texto: string } | null>(null);
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
        <div className="aviso info" style={{ marginBottom: 12, display: "flex", flexDirection: "column", gap: 10 }}>
          <span><strong>Cargar muchos costos juntos:</strong> bajá la planilla con todos los productos, completá la columna Costo en Excel y subila. También acepta el reporte &quot;Costo de mercaderías&quot; de Bistrosoft.</span>
          <div className="fila">
            <button className="btn sec" onClick={() => planillaCostos(r, costos, nombreLocal)}><IcoDescarga /> Bajar planilla de costos</button>
            <label className="btn" style={{ position: "relative" }}>
              <IcoCargar /> Subir costos
              <input id="costos-archivo" type="file" accept=".xlsx,.xls,.csv" style={{ position: "absolute", opacity: 0, width: 1, height: 1 }}
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  e.target.value = "";
                  if (!f) return;
                  try {
                    const lista = await leerCostos(f);
                    if (!lista.length) {
                      setMsg({ tipo: "bad", texto: "El archivo no tiene ningún costo cargado. Si es el reporte de Bistrosoft, primero hay que cargar los costos allá; si no, completá la columna Costo de la planilla." });
                      return;
                    }
                    importarCostos(lista);
                    const vendidos = new Set(r.productos.map((p) => norm(p.articulo)));
                    const coinciden = lista.filter((x) => vendidos.has(norm(x.articulo))).length;
                    setMsg({ tipo: "ok", texto: `Se cargaron ${lista.length} costos (${coinciden} coinciden con productos vendidos este mes).` });
                  } catch (err) {
                    setMsg({ tipo: "bad", texto: (err as Error).message });
                  }
                }} />
            </label>
          </div>
          {msg && <span className={`aviso ${msg.tipo}`} role="status">{msg.texto}</span>}
        </div>
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
