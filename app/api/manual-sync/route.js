import { NextResponse } from "next/server";
import { runSync } from "@/lib/syncEngine";

export const maxDuration = 60;

export async function POST() {
  try {
    const resultado = await runSync();
    return NextResponse.json({ ok: true, ...resultado });
  } catch (err) {
    console.error("Error en sync manual:", err);
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
