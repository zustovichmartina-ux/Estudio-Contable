import { Redis } from "@upstash/redis";
import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

const PREFIJO = "fiordo";

function cliente(): Redis | null {
  const url = process.env.KV_REST_API_URL ?? process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN ?? process.env.UPSTASH_REDIS_REST_TOKEN;
  return url && token ? new Redis({ url, token }) : null;
}

const sinBase = () =>
  NextResponse.json({ error: "No hay base de datos conectada; los datos se guardan en este navegador." }, { status: 501 });

/** Devuelve la configuración y todos los meses cargados. */
export async function GET() {
  const r = cliente();
  if (!r) return sinBase();
  const [config, meses] = await Promise.all([r.get(`${PREFIJO}:config`), r.smembers(`${PREFIJO}:meses`)]);
  const datos = meses.length ? await r.mget(...meses.map((m) => `${PREFIJO}:mes:${m}`)) : [];
  const porMes: Record<string, unknown> = {};
  meses.forEach((m, i) => {
    if (datos[i]) porMes[m] = datos[i];
  });
  return NextResponse.json({ config: config ?? null, meses: porMes });
}

/** Guarda la configuración ({tipo:"config", data}) o un mes ({tipo:"mes", mes, data}). */
export async function PUT(req: Request) {
  const r = cliente();
  if (!r) return sinBase();
  const body = await req.json().catch(() => null);
  if (body?.tipo === "config" && body.data) {
    await r.set(`${PREFIJO}:config`, body.data);
    return NextResponse.json({ ok: true });
  }
  if (body?.tipo === "mes" && /^\d{4}-\d{2}$/.test(body.mes ?? "") && body.data) {
    await Promise.all([r.set(`${PREFIJO}:mes:${body.mes}`, body.data), r.sadd(`${PREFIJO}:meses`, body.mes)]);
    return NextResponse.json({ ok: true });
  }
  return NextResponse.json({ error: "Pedido inválido" }, { status: 400 });
}

/** Borra las ventas de un mes (?mes=AAAA-MM). */
export async function DELETE(req: Request) {
  const r = cliente();
  if (!r) return sinBase();
  const mes = new URL(req.url).searchParams.get("mes") ?? "";
  if (!/^\d{4}-\d{2}$/.test(mes)) return NextResponse.json({ error: "Mes inválido" }, { status: 400 });
  await Promise.all([r.del(`${PREFIJO}:mes:${mes}`), r.srem(`${PREFIJO}:meses`, mes)]);
  return NextResponse.json({ ok: true });
}
