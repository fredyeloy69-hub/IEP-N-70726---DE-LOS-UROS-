import { NextResponse } from "next/server";
import { adminDb } from "@/lib/firebaseAdmin";
import { getDriveClient } from "@/lib/googleDrive";

export const maxDuration = 60;

// Corrige la hora de los eventos YA guardados (los que quedaron con la hora de
// la sincronización) usando la hora real que guarda Drive en cada archivo.
// Uso: /api/reparar-horas?secret=TU_CRON_SECRET   (opcional: &n=100 = cuántos eventos revisar)
export async function GET(request) {
  const url = new URL(request.url);
  const autorizado =
    request.headers.get("authorization") === `Bearer ${process.env.CRON_SECRET}` ||
    url.searchParams.get("secret") === process.env.CRON_SECRET;
  if (!autorizado) return NextResponse.json({ error: "No autorizado" }, { status: 401 });

  const n = Math.min(Number(url.searchParams.get("n")) || 60, 150);
  try {
    const drive = getDriveClient();
    const snap = await adminDb.collection("eventos").orderBy("timestamp", "desc").limit(n).get();
    const cambios = [];
    let revisados = 0;

    for (const d of snap.docs) {
      const e = d.data();
      if (!e.itemId || !e.timestamp || !e.timestamp.toDate) continue;
      revisados++;
      let iso = null;
      try {
        const r = await drive.files.get({
          fileId: e.itemId,
          fields: "createdTime,modifiedTime,trashedTime",
          supportsAllDrives: true,
        });
        if (/subido|copiado|creada/.test(e.tipo)) iso = r.data.createdTime;
        else if (/reemplazado/.test(e.tipo)) iso = r.data.modifiedTime;
        else if (/borrad/.test(e.tipo)) iso = r.data.trashedTime;
      } catch {
        // archivo borrado del todo o sin acceso: se deja como está
      }
      if (!iso) continue;
      const real = new Date(iso);
      const actual = e.timestamp.toDate();
      const difMin = (actual.getTime() - real.getTime()) / 60000;
      // Solo se corrige si la hora guardada es posterior a la real (hora de sync) y es de hoy/ayer
      if (difMin > 2 && difMin < 60 * 48) {
        await d.ref.update({ timestamp: real });
        cambios.push({ item: e.item, tipo: e.tipo, antes: actual.toISOString(), ahora: real.toISOString() });
      }
    }
    return NextResponse.json({ ok: true, revisados, corregidos: cambios.length, cambios });
  } catch (err) {
    console.error("Error reparando horas:", err);
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
