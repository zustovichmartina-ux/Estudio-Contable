import { NextRequest, NextResponse } from "next/server";
import { COOKIE, credenciales, fichaSesion } from "@/lib/sesion";

export async function POST(req: NextRequest) {
  const cred = credenciales();
  if (!cred) return NextResponse.json({ ok: true });
  const { usuario = "", clave = "" } = await req.json().catch(() => ({}));
  const u = String(usuario).trim().toLowerCase();
  if (u !== cred.usuario.toLowerCase() || String(clave) !== cred.clave) {
    return NextResponse.json({ ok: false }, { status: 401 });
  }
  const res = NextResponse.json({ ok: true });
  res.cookies.set(COOKIE, await fichaSesion(cred.usuario, cred.clave), {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 365,
  });
  return res;
}
