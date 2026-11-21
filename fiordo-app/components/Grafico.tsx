"use client";
import { useEffect, useRef, useState } from "react";
import { Resultado } from "@/lib/motor";
import { corto, fechaCorta, pesos } from "@/lib/formato";

/** Barras de venta por día (verde = cubrió, rojo = no) contra la línea de lo necesario. */
export default function Grafico({ r }: { r: Resultado }) {
  const ref = useRef<HTMLDivElement>(null);
  const [w, setW] = useState(640);
  useEffect(() => {
    if (!ref.current) return;
    const ro = new ResizeObserver(([e]) => setW(Math.max(300, e.contentRect.width)));
    ro.observe(ref.current);
    return () => ro.disconnect();
  }, []);

  const H = w < 500 ? 200 : 250;
  const padL = 44, padR = 8, padT = 18, padB = 26;
  const n = Math.max(r.diasApertura, r.dias.length ? Number(r.dias[r.dias.length - 1].fecha.slice(8)) : 0);
  const max = Math.max(isFinite(r.peDia) ? r.peDia : 0, ...r.dias.map((d) => d.venta), 1) * 1.12;
  const paso = max > 3e6 ? 1e6 : max > 1.5e6 ? 5e5 : 25e4;
  const top = Math.ceil(max / paso) * paso;
  const y = (v: number) => padT + (H - padT - padB) * (1 - v / top);
  const bw = (w - padL - padR) / n;
  const porDia = new Map(r.dias.map((d) => [Number(d.fecha.slice(8)), d]));
  const marcas: number[] = [];
  for (let t = 0; t <= top; t += paso) marcas.push(t);

  return (
    <div ref={ref} className="chart">
      <svg width={w} height={H} role="img" aria-label="Ventas por día contra lo necesario">
        {marcas.map((t) => (
          <g key={t}>
            <line x1={padL} x2={w - padR} y1={y(t)} y2={y(t)} stroke="var(--line)" />
            <text x={padL - 6} y={y(t) + 4} textAnchor="end">{corto(t).replace("$ ", "")}</text>
          </g>
        ))}
        {Array.from({ length: n }, (_, i) => {
          const dia = i + 1;
          const d = porDia.get(dia);
          const x = padL + i * bw + bw * 0.15;
          const bwi = Math.max(1, bw * 0.7);
          return (
            <g key={dia}>
              {d ? (
                <rect x={x} y={y(d.venta)} width={bwi} height={y(0) - y(d.venta)} rx={2} fill={d.cubrio ? "var(--ok)" : "var(--bad)"}>
                  <title>{`${fechaCorta(d.fecha)}: ${pesos(d.venta)}`}</title>
                </rect>
              ) : (
                <rect x={x} y={y(0) - 3} width={bwi} height={3} rx={1} fill="var(--bar-empty)" />
              )}
              {(dia === 1 || dia % 5 === 0) && (
                <text x={x + bwi / 2} y={H - padB + 15} textAnchor="middle">{dia}</text>
              )}
            </g>
          );
        })}
        {isFinite(r.peDia) && r.peDia > 0 && (
          <>
            <line x1={padL} x2={w - padR} y1={y(r.peDia)} y2={y(r.peDia)} stroke="var(--ink)" strokeWidth={2} strokeDasharray="6 4" />
            <text x={w - padR} y={y(r.peDia) - 6} textAnchor="end" style={{ fill: "var(--ink)", fontWeight: 600 }}>
              {`${corto(r.peDia)} por día`}
            </text>
          </>
        )}
      </svg>
    </div>
  );
}
