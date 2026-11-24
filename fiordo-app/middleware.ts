import { NextRequest, NextResponse } from "next/server";
import { COOKIE, credenciales, fichaSesion } from "./lib/sesion";

/**
 * Si APP_PASSWORD está definida, hay que entrar con usuario y contraseña en /entrar.
 * Queda recordado en el dispositivo (cookie) por un año. En localhost sin variable, la app abre directo.
 */
export async function middleware(req: NextRequest) {
  const cred = credenciales();
  if (!cred) return NextResponse.next();
  const { pathname } = req.nextUrl;
  if (pathname === "/entrar" || pathname === "/api/entrar") return NextResponse.next();

  const ficha = req.cookies.get(COOKIE)?.value;
  if (ficha && ficha === (await fichaSesion(cred.usuario, cred.clave))) return NextResponse.next();

  // también acepta el usuario y contraseña "a la antigua" (cartel del navegador)
  const auth = req.headers.get("authorization") ?? "";
  if (auth.startsWith("Basic ")) {
    const txt = atob(auth.slice(6));
    const i = txt.indexOf(":");
    if (i >= 0 && txt.slice(0, i) === cred.usuario && txt.slice(i + 1) === cred.clave) return NextResponse.next();
  }

  if (pathname.startsWith("/api/")) return NextResponse.json({ error: "sin sesión" }, { status: 401 });
  const url = req.nextUrl.clone();
  url.pathname = "/entrar";
  url.search = "";
  url.hash = "";
  return NextResponse.redirect(url);
}

export const config = { matcher: ["/((?!_next/static|_next/image|favicon.ico|icon|apple-touch-icon|manifest.webmanifest|sw.js).*)"] };
