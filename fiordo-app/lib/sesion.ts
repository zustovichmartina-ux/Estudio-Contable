/** Sesión simple: una cookie con un resumen (SHA-256) de usuario + contraseña. Funciona igual en el navegador y en la app instalada en el celular. */
export const COOKIE = "fiordo_sesion";

export function credenciales() {
  const clave = process.env.APP_PASSWORD;
  if (!clave) return null;
  return { usuario: (process.env.APP_USER ?? "fiordo").trim(), clave };
}

export async function fichaSesion(usuario: string, clave: string): Promise<string> {
  const datos = new TextEncoder().encode(`fiordo|${usuario}|${clave}`);
  const h = await crypto.subtle.digest("SHA-256", datos);
  return Array.from(new Uint8Array(h), (b) => b.toString(16).padStart(2, "0")).join("");
}
