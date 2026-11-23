import { CONFIG_INICIAL, Config, Datos, MesVentas, PARAMETROS_INICIALES } from "./tipos";

/** "nube" = base compartida en Vercel (celular y compu ven lo mismo); "local" = solo este navegador. */
export type Modo = "nube" | "local";

const CLAVE_LOCAL = "fiordo-datos-v1";

function leerLocal(): Datos {
  try {
    const raw = localStorage.getItem(CLAVE_LOCAL);
    if (raw) return completar(JSON.parse(raw));
  } catch {
    /* navegador sin almacenamiento: arrancamos vacío */
  }
  return { config: CONFIG_INICIAL, meses: {} };
}

function escribirLocal(d: Datos) {
  try {
    localStorage.setItem(CLAVE_LOCAL, JSON.stringify(d));
  } catch {
    /* sin espacio o bloqueado */
  }
}

function completar(d: Partial<Datos> & { config?: Partial<Config> | null }): Datos {
  const c: Partial<Config> = d.config ?? {};
  return {
    config: {
      ...CONFIG_INICIAL,
      ...c,
      parametros: migrarParametros({ ...PARAMETROS_INICIALES, ...(c.parametros ?? {}) }, c.parametros),
    } as Config,
    meses: (d.meses ?? {}) as Record<string, MesVentas>,
  };
}

export async function cargar(): Promise<{ datos: Datos; modo: Modo }> {
  try {
    const r = await fetch("/api/datos", { cache: "no-store" });
    if (r.ok) {
      const j = await r.json();
      const datos = completar(j);
      escribirLocal(datos); // copia de resguardo
      return { datos, modo: "nube" };
    }
  } catch {
    /* sin conexión: usamos lo local */
  }
  return { datos: leerLocal(), modo: "local" };
}

async function put(body: unknown) {
  const r = await fetch("/api/datos", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error("No se pudo guardar en la nube. Revisá la conexión e intentá de nuevo.");
}

export async function guardarConfig(datos: Datos, modo: Modo) {
  escribirLocal(datos);
  if (modo === "nube") await put({ tipo: "config", data: datos.config });
}

export async function guardarMeses(datos: Datos, mesesCambiados: string[], modo: Modo) {
  escribirLocal(datos);
  if (modo === "nube") for (const mes of mesesCambiados) await put({ tipo: "mes", mes, data: datos.meses[mes] });
}

export async function borrarMes(datos: Datos, mes: string, modo: Modo) {
  escribirLocal(datos);
  if (modo === "nube") await fetch(`/api/datos?mes=${mes}`, { method: "DELETE" });
}

/** Ajustes de versiones anteriores: IIBB pasó de 3 % (referencia) a 4 %, y las comisiones van por medio de pago. */
function migrarParametros(p: Config["parametros"], guardado?: Partial<Config["parametros"]>): Config["parametros"] {
  const out = { ...p };
  if (guardado && guardado.comisionesPorMedio === undefined) {
    if (guardado.iibbPct === 3) out.iibbPct = 4;
    out.comisionPct = 0; // antes era una tasa pareja para todo; ahora se usa por medio
  }
  return out;
}
