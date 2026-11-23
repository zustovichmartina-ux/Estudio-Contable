import { NextRequest, NextResponse } from "next/server";

/**
 * Si APP_PASSWORD está definida, pide usuario y contraseña para entrar.
 * En localhost sin variable, la app abre directo.
 */
export function middleware(req: NextRequest) {
  const clave = process.env.APP_PASSWORD;
  if (!clave) return NextResponse.next();
  const usuario = process.env.APP_USER ?? "fiordo";
  const auth = req.headers.get("authorization") ?? "";
  if (auth.startsWith("Basic ")) {
    const [u, p] = atob(auth.slice(6)).split(":");
    if (u === usuario && p === clave) return NextResponse.next();
  }
  return new NextResponse("Ingresá usuario y contraseña", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="Fiordo", charset="UTF-8"' },
  });
}

export const config = { matcher: ["/((?!_next/static|_next/image|favicon.ico|icon|apple-touch-icon|manifest.webmanifest|sw.js).*)"] };
