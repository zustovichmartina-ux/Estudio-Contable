"use client";
import { useState } from "react";
import { Resultado, avisoWhatsapp, recomendaciones } from "@/lib/motor";
import { nombreMes } from "@/lib/formato";
import { reporteExcel, reportePDF } from "@/lib/reportes";
import { IcoCopiar, IcoDescarga, IcoWhats } from "./Iconos";

export default function Reportes({ r, nombreLocal }: { r: Resultado; nombreLocal: string }) {
  const [estado, setEstado] = useState("");
  const [generando, setGenerando] = useState<"" | "pdf" | "xlsx">("");
  const texto = avisoWhatsapp(r, nombreLocal);
  const recs = recomendaciones(r);

  async function generar(tipo: "pdf" | "xlsx") {
    setGenerando(tipo);
    try {
      if (tipo === "pdf") await reportePDF(r, nombreLocal);
      else await reporteExcel(r, nombreLocal);
    } catch (e) {
      setEstado(`No se pudo generar el archivo: ${(e as Error).message}`);
    } finally {
      setGenerando("");
    }
  }

  async function copiar() {
    try {
      await navigator.clipboard.writeText(texto);
      setEstado("Mensaje copiado.");
    } catch {
      setEstado("No se pudo copiar automáticamente: mantené apretado el texto para copiarlo.");
    }
    setTimeout(() => setEstado(""), 3000);
  }

  if (!r.hayDatos) return <div className="card vacio"><h2>Sin datos para reportar</h2><p className="muted">Cargá las ventas de {nombreMes(r.mes)} para generar los reportes.</p></div>;

  return (
    <>
      <div className="grid-2">
        <section className="card" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <div className="card-head" style={{ marginBottom: 0 }}><h2>Informe de {nombreMes(r.mes)}</h2></div>
          <p className="muted">El PDF trae indicadores, gráfico, conclusiones y el detalle por día y por producto. El Excel trae todo en hojas separadas para seguir trabajando.</p>
          <div className="fila">
            <button className="btn" onClick={() => generar("pdf")} disabled={!!generando}><IcoDescarga /> {generando === "pdf" ? "Generando…" : "Descargar PDF"}</button>
            <button className="btn sec" onClick={() => generar("xlsx")} disabled={!!generando}><IcoDescarga /> {generando === "xlsx" ? "Generando…" : "Descargar Excel"}</button>
          </div>
          <h3 style={{ marginTop: 6 }}>Conclusiones del mes</h3>
          <ul className="lista-recs">{recs.map((t, i) => <li key={i}>{t}</li>)}</ul>
        </section>

        <section className="card" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <div className="card-head" style={{ marginBottom: 0 }}><h2>Aviso para el dueño</h2><span className="small muted">Para mandar por WhatsApp</span></div>
          <div className="burbuja">{texto}</div>
          <div className="fila">
            <a className="btn" href={`https://wa.me/?text=${encodeURIComponent(texto)}`} target="_blank" rel="noreferrer"><IcoWhats /> Enviar por WhatsApp</a>
            <button className="btn sec" onClick={copiar}><IcoCopiar /> Copiar</button>
          </div>
        </section>
      </div>
      {estado && <div className="aviso info" role="status">{estado}</div>}
    </>
  );
}
