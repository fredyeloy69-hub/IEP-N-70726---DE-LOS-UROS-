import { NextResponse } from "next/server";
import { adminDb } from "../../../lib/firebaseAdmin";
import { FieldValue } from "firebase-admin/firestore";
import { getAuth } from "firebase-admin/auth";
import { fechaLimaISO } from "../../../lib/syncEngine";

// Marca (o desmarca) una carpeta manualmente como "completa" o "incompleta" —
// para casos excepcionales donde no aplica tener un editable (ej. documentos
// escaneados donde solo existe el PDF del trámite), o para revertir un
// estado calculado automáticamente que el evaluador considera incorrecto.
//
// `estado` acepta: "completa" | "incompleta" | null (quita el forzado y deja
// que el próximo sync recalcule el estado real según los archivos).
//
// IMPORTANTE: el usuario NUNCA se toma de un texto que mande el navegador —
// se exige un idToken de Firebase Auth (login real con Google) y se verifica
// en el servidor con Firebase Admin. Así el nombre/correo que queda guardado
// es siempre el verdadero, igual que Drive hace con sus propios archivos.
//
// Guarda la excepción en la colección "carpetasForzadas" (para que sobreviva
// a los syncs automáticos, que la vuelven a leer en cada corrida) y también
// actualiza el documento en "carpetas" al toque, para que se vea el cambio
// en el dashboard de todos los evaluadores sin esperar el próximo sync.
//
// Además registra un evento en "eventos" (igual que subidas/borrados de Drive)
// y suma al contador agregado de actividad por día, para que quede visible en
// "Actividad reciente" y en el mapa de calor — así cualquier evaluador que
// entre después puede ver quién marcó o desmarcó qué, y por qué.
export async function POST(request) {
  try {
    const { folderId, estado, motivo, idToken, folderName, folderRuta } = await request.json();

    if (!folderId) {
      return NextResponse.json({ error: "Falta folderId" }, { status: 400 });
    }
    if (estado !== "completa" && estado !== "incompleta" && estado !== null) {
      return NextResponse.json(
        { error: "estado debe ser 'completa', 'incompleta' o null" },
        { status: 400 }
      );
    }
    if (!idToken) {
      return NextResponse.json({ error: "Falta iniciar sesión con Google" }, { status: 401 });
    }

    // --- Verificar quién es de verdad, del lado del servidor ---
    let decoded;
    try {
      decoded = await getAuth(adminDb.app).verifyIdToken(idToken);
    } catch (err) {
      return NextResponse.json({ error: "Sesión inválida o vencida, vuelve a iniciar sesión" }, { status: 401 });
    }
    const nombreDecoded = decoded.name;
    const correoDecoded = decoded.email;
    const nombreUsuario =
      nombreDecoded && correoDecoded
        ? `${nombreDecoded} (${correoDecoded})`
        : nombreDecoded || correoDecoded || "Cuenta de Google verificada";

    const nombreCarpeta = folderName || "(carpeta)";
    const rutaCarpeta = folderRuta || nombreCarpeta;
    const overrideRef = adminDb.collection("carpetasForzadas").doc(folderId);
    const now = new Date();
    const nowISO = now.toISOString();

    if (estado === "completa" || estado === "incompleta") {
      await overrideRef.set({
        forzada: true,
        estadoForzado: estado,
        motivo: motivo || "",
        marcadoPor: nombreUsuario,
        marcadoPorEmail: decoded.email || null,
        marcadoEn: nowISO,
      });
      await adminDb
        .collection("carpetas")
        .doc(folderId)
        .set(
          {
            estado,
            detalle: `Marcada manualmente como ${estado}${motivo ? ` — ${motivo}` : ""}`,
            forzada: true,
            estadoForzado: estado,
            marcadoPor: nombreUsuario,
            marcadoPorEmail: decoded.email || null,
            motivo: motivo || "",
            marcadoEn: nowISO,
          },
          { merge: true }
        );
    } else {
      await overrideRef.delete();
      await adminDb
        .collection("carpetas")
        .doc(folderId)
        .set(
          {
            forzada: false,
            estadoForzado: null,
            marcadoPor: null,
            marcadoPorEmail: null,
            motivo: null,
            marcadoEn: null,
          },
          { merge: true }
        );
      // Nota: el estado real (completa/incompleta/vacía) según los archivos
      // se vuelve a calcular recién en el próximo sync, no acá al toque.
    }

    // --- Registrar el evento para que se vea en "Actividad reciente" ---
    const tipoEvento =
      estado === "completa"
        ? "carpeta_marcada_completa"
        : estado === "incompleta"
        ? "carpeta_marcada_incompleta"
        : "carpeta_desmarcada";
    await adminDb.collection("eventos").add({
      tipo: tipoEvento,
      item: nombreCarpeta,
      ruta: rutaCarpeta,
      usuario: nombreUsuario,
      motivo: motivo || null,
      folderId,
      timestamp: FieldValue.serverTimestamp(),
    });

    // --- Sumar al contador agregado del día (para el mapa de calor) ---
    const fechaHoy = fechaLimaISO(now);
    await adminDb
      .collection("_meta")
      .doc("actividadPorDia")
      .set({ [fechaHoy]: { [tipoEvento]: FieldValue.increment(1) } }, { merge: true });

    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: err.message || "Error desconocido" }, { status: 500 });
  }
}
