import { NextResponse } from "next/server";
import { adminDb } from "../../../lib/firebaseAdmin";
import { analizarPdf } from "../../../lib/googleDrive";

export const maxDuration = 60;

const MAX_POR_LLAMADA = 8; // PDFs que se analizan al mismo tiempo en una llamada
const TIMEOUT_PDF_MS = 40000;

function conTiempoMaximo(promesa, ms) {
  return Promise.race([promesa, new Promise((res) => setTimeout(() => res(null), ms))]);
}

// Analiza el formato REAL (A4/A3/A2/A1/A0) de cada hoja de varios PDFs y lo
// guarda en el documento de su carpeta, para que el botón "Material de
// impresión" no tenga que adivinar por el nombre del archivo.
//
// Solo analiza PDFs que YA están registrados en Firestore (colección
// "carpetas") — no se puede pedir analizar un archivo cualquiera de Drive.
//
// Body: { items: [{ carpetaId, fileId }, ...] }  (máx. 8 por llamada)
// Respuesta: { ok, resultados: { [fileId]: { paginas, formatos, aprox } | null } }
export async function POST(request) {
  try {
    const { items } = await request.json();
    if (!Array.isArray(items) || items.length === 0) {
      return NextResponse.json({ ok: false, error: "Falta items" }, { status: 400 });
    }
    const pedidos = items
      .filter((i) => i && typeof i.carpetaId === "string" && typeof i.fileId === "string")
      .slice(0, MAX_POR_LLAMADA);

    // Leer cada carpeta una sola vez
    const idsCarpetas = [...new Set(pedidos.map((p) => p.carpetaId))];
    const docs = {};
    await Promise.all(
      idsCarpetas.map(async (id) => {
        const snap = await adminDb.collection("carpetas").doc(id).get();
        if (snap.exists) docs[id] = { ref: snap.ref, archivos: snap.data().archivos || [] };
      })
    );

    // Solo PDFs que de verdad existen en esa carpeta
    const validos = pedidos.filter((p) => {
      const d = docs[p.carpetaId];
      return d && d.archivos.some((a) => a.id === p.fileId && a.tipo === "pdf");
    });

    const resultados = {};
    await Promise.all(
      validos.map(async (p) => {
        resultados[p.fileId] = await conTiempoMaximo(analizarPdf(p.fileId, TIMEOUT_PDF_MS - 2000), TIMEOUT_PDF_MS);
      })
    );

    // Guardar (1 escritura por carpeta)
    const carpetasConCambios = new Set();
    for (const p of validos) {
      const r = resultados[p.fileId];
      if (!r) continue;
      const d = docs[p.carpetaId];
      d.archivos = d.archivos.map((a) =>
        a.id === p.fileId ? { ...a, paginas: r.paginas, formatos: r.formatos, aprox: r.aprox || null } : a
      );
      carpetasConCambios.add(p.carpetaId);
    }
    await Promise.all(
      [...carpetasConCambios].map((id) => docs[id].ref.update({ archivos: docs[id].archivos }))
    );

    return NextResponse.json({ ok: true, resultados });
  } catch (err) {
    console.error("Error en analizar-formatos:", err);
    return NextResponse.json({ ok: false, error: err.message }, { status: 500 });
  }
}
