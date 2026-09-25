import { NextResponse } from "next/server";
import { runSync } from "@/lib/syncEngine";
 
export const maxDuration = 60; // segundos, ajustable segun plan de Vercel
 
export async function GET(request) {
  const authHeader = request.headers.get("authorization");
  const url = new URL(request.url);
  const secretParam = url.searchParams.get("secret");
 
  const autorizado =
    authHeader === `Bearer ${process.env.CRON_SECRET}` ||
    secretParam === process.env.CRON_SECRET;
 
  if (!autorizado) {
    return NextResponse.json({ error: "No autorizado" }, { status: 401 });
  }
 
  try {
    const resultado = await runSync();
    return NextResponse.json({ ok: true, ...resultado });
  } catch (err) {
    console.error("Error en sync:", err);
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
