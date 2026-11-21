const nf = new Intl.NumberFormat("es-AR", { maximumFractionDigits: 0 });

export const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const DIAS_CORTOS = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"];
export const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

export const pesos = (n: number) => "$ " + nf.format(Math.round(n));

export function corto(n: number): string {
  const a = Math.abs(n);
  const s = n < 0 ? "-" : "";
  if (!isFinite(a)) return "–";
  if (a >= 1e6) return `${s}$ ${(a / 1e6).toLocaleString("es-AR", { maximumFractionDigits: 1 })} M`;
  if (a >= 1e3) return `${s}$ ${Math.round(a / 1e3).toLocaleString("es-AR")} mil`;
  return s + pesos(a);
}

export const pct = (x: number, dec = 1) =>
  isFinite(x) ? `${(x * 100).toLocaleString("es-AR", { maximumFractionDigits: dec, minimumFractionDigits: dec })} %` : "–";

/** "2026-11-21" -> Date local (sin corrimiento por zona horaria) */
export const aFecha = (iso: string) => {
  const [a, m, d] = iso.split("-").map(Number);
  return new Date(a, m - 1, d);
};

/** Lunes = 0 ... Domingo = 6 */
export const diaSemana = (iso: string) => (aFecha(iso).getDay() + 6) % 7;

export const fechaCorta = (iso: string) => {
  const f = aFecha(iso);
  return `${DIAS_CORTOS[diaSemana(iso)]} ${f.getDate()}/${f.getMonth() + 1}`;
};

export const nombreMes = (mes: string) => {
  const [a, m] = mes.split("-").map(Number);
  return `${MESES[m - 1]} ${a}`;
};

export const diasDelMes = (mes: string) => {
  const [a, m] = mes.split("-").map(Number);
  return new Date(a, m, 0).getDate();
};
