import { google } from "googleapis";
import { PDFDocument, PDFName, PDFNumber } from "pdf-lib";
import { clasificarHoja } from "./formatosHoja";

function getAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n");

  if (!email) {
    throw new Error(
      "Falta GOOGLE_SERVICE_ACCOUNT_EMAIL"
    );
  }

  if (!privateKey) {
    throw new Error(
      "Falta GOOGLE_PRIVATE_KEY"
    );
  }

  return new google.auth.JWT({
    email,
    key: privateKey,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/drive.activity.readonly",
    ],
  });
}

export function getDriveClient() {
  return google.drive({
    version: "v3",
    auth: getAuth(),
  });
}

export function getActivityClient() {
  return google.driveactivity({
    version: "v2",
    auth: getAuth(),
  });
}

const FIELDS =
  "files(id,name,mimeType,parents,trashed,modifiedTime,md5Checksum,size,shortcutDetails(targetId,targetMimeType),webViewLink,thumbnailLink,lastModifyingUser(displayName,emailAddress))";

async function listFolderContents(drive, folderId) {
  const items = [];
  let pageToken = undefined;

  do {
    const res = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false`,
      fields: `nextPageToken, ${FIELDS}`,
      pageSize: 1000,
      pageToken,
    });

    items.push(...(res.data.files || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);

  return items;
}

export async function scanDriveTree(rootId) {
  if (!rootId) {
    throw new Error(
      "Falta DRIVE_ROOT_FOLDER_ID"
    );
  }

  const drive = getDriveClient();

  const allItems = [];
  const visited = new Set([rootId]);

  let currentLevel = [rootId];

  while (currentLevel.length > 0) {
    const results = await Promise.all(
      currentLevel.map((folderId) =>
        listFolderContents(drive, folderId)
      )
    );

    const nextLevel = [];

    for (const items of results) {
      for (const original of items) {
        let file = original;

        // Acceso directo de Drive a una carpeta: se trata como si fuera la carpeta
        // real (se entra a su contenido original) y se le quita el ".lnk" del nombre.
        if (
          file.mimeType === "application/vnd.google-apps.shortcut" &&
          file.shortcutDetails?.targetMimeType === "application/vnd.google-apps.folder" &&
          file.shortcutDetails?.targetId
        ) {
          file = {
            ...file,
            id: file.shortcutDetails.targetId,
            name: String(file.name || "").replace(/\.lnk$/i, "").trim(),
            mimeType: "application/vnd.google-apps.folder",
          };
        }

        allItems.push(file);

        if (
          file.mimeType ===
            "application/vnd.google-apps.folder" &&
          !visited.has(file.id)
        ) {
          visited.add(file.id);
          nextLevel.push(file.id);
        }
      }
    }

    currentLevel = nextLevel;
  }

  return allItems;
}

export async function getUltimoActorDeItem(itemId) {
  // 1) Drive guarda quién mandó a la papelera cada archivo/carpeta (trashingUser).
  try {
    const drive = getDriveClient();
    const r = await drive.files.get({
      fileId: itemId,
      fields: "trashed,trashedTime,trashingUser(displayName,emailAddress)",
      supportsAllDrives: true,
    });
    const u = r.data && r.data.trashingUser;
    if (u && (u.displayName || u.emailAddress)) {
      return u.displayName && u.emailAddress ? `${u.displayName} (${u.emailAddress})` : u.emailAddress || u.displayName;
    }
  } catch (err) {
    console.log(`[actor] papelera ${itemId}: ${String((err && err.message) || err).slice(0, 160)}`);
  }

  // 2) Drive Activity (si la cuenta de servicio tiene acceso al historial)
  try {
    const activity = getActivityClient();

    const res = await activity.activity.query({
      requestBody: {
        itemName: `items/${itemId}`,
        pageSize: 10,
      },
    });

    const activities = res.data.activities || [];

    for (const act of activities) {
      const actor = (act.actors || []).find((a) => a?.user?.knownUser);

      if (actor) {
        const ku = actor.user.knownUser;
        return ku.personName || null;
      }
    }

    return null;
  } catch (err) {
    console.log(`[actor] activity ${itemId}: ${String((err && err.message) || err).slice(0, 160)}`);
    return null;
  }
}

export async function contarPaginasPdf(fileId) {
  try {
    const drive = getDriveClient();

    const descarga = drive.files.get(
      {
        fileId,
        alt: "media",
      },
      {
        responseType: "arraybuffer",
      }
    );

    const timeout = new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error("timeout")),
        25000
      )
    );

    const res = await Promise.race([
      descarga,
      timeout,
    ]);

    const bytes = Buffer.from(res.data);

    const pdf = await PDFDocument.load(bytes, {
      ignoreEncryption: true,
    });

    return pdf.getPageCount();
  } catch (err) {
    console.error(
      `[contarPaginasPdf] Falló para el archivo ${fileId}:`,
      err.message || err
    );

    return null;
  }
}

/**
 * Descarga un PDF de Drive y devuelve, además de la cantidad de páginas, el
 * FORMATO REAL (A4/A3/A2/A1/A0) de cada una según su tamaño físico — así un
 * mismo PDF con hojas de varios formatos se cuenta bien, sin importar si el
 * nombre del archivo dice A1, A3, etc.
 *
 * @returns {Promise<{paginas:number, formatos:{A4:number,A3:number,A2:number,A1:number,A0:number}, aprox:Array<{mm:string,formato:string,n:number}>}|null>}
 *   `aprox` lista las hojas con proporción rara (se les asignó el formato A más
 *   cercano en área). Si el PDF no se pudo abrir devuelve { error: "motivo" }.
 */
export async function analizarPdf(fileId, timeoutMs = 25000) {
  try {
    const drive = getDriveClient();

    const descarga = drive.files.get(
      { fileId, alt: "media" },
      { responseType: "arraybuffer", maxContentLength: Infinity, maxBodyLength: Infinity }
    );
    const timeout = new Promise((_, reject) =>
      setTimeout(() => reject(new Error("timeout")), timeoutMs)
    );
    const res = await Promise.race([descarga, timeout]);
    // Sin copiar el buffer: con PDFs de hasta 50 MB se ahorra memoria.
    const bytes = new Uint8Array(res.data);

    const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    const paginasPdf = pdf.getPages();

    const formatos = { A4: 0, A3: 0, A2: 0, A1: 0, A0: 0 };
    const aproxMap = new Map();
    const tamanosMap = new Map(); // "corto x largo" (mm) -> cantidad de hojas
    let numPag = 0;
    for (const pagina of paginasPdf) {
      numPag += 1;
      // Se mide la HOJA COMPLETA (MediaBox). El CropBox puede ser más chico si
      // alguien recortó los márgenes en pantalla, pero el papel sigue siendo
      // el de la hoja original (ej. un A4 recortado a 172×275 mm).
      let ancho, alto;
      try {
        const box = pagina.getMediaBox();
        ancho = box.width;
        alto = box.height;
      } catch {
        const s = pagina.getSize();
        ancho = s.width;
        alto = s.height;
      }
      // UserUnit: algunos PDFs indican que 1 punto no es 1/72" sino un múltiplo
      try {
        const uu = pagina.node.lookup(PDFName.of("UserUnit"));
        if (uu instanceof PDFNumber && uu.asNumber() > 0) {
          ancho *= uu.asNumber();
          alto *= uu.asNumber();
        }
      } catch {
        // sin UserUnit: se usa el tamaño tal cual
      }
      if (!(ancho > 0) || !(alto > 0)) {
        formatos.A4 += 1;
        continue;
      }
      const c = clasificarHoja(ancho, alto);
      formatos[c.formato] += 1;
      const claveTam = `${c.mm[0]}x${c.mm[1]}`;
      const tm = tamanosMap.get(claveTam) || { n: 0, pags: [] };
      tm.n += 1;
      // Se anota el número de página solo de las hojas de medida rara (máx. 12 por medida)
      if (!c.exacto && tm.pags.length < 12) tm.pags.push(numPag);
      tamanosMap.set(claveTam, tm);
      if (!c.exacto) {
        const clave = `${c.mm[0]}x${c.mm[1]}→${c.formato}`;
        const previo = aproxMap.get(clave) || { mm: `${c.mm[0]}x${c.mm[1]}`, formato: c.formato, n: 0 };
        previo.n += 1;
        aproxMap.set(clave, previo);
      }
    }

    return {
      v: 4, // versión del método de medición (4 = MediaBox + UserUnit + números de página de hojas raras)
      paginas: paginasPdf.length,
      formatos,
      aprox: [...aproxMap.values()].slice(0, 6),
      // Medidas reales (mm) y cuántas hojas hay de cada una. Permite recalcular
      // el formato si cambia la regla de tolerancia, sin volver a descargar.
      tamanos: [...tamanosMap.entries()]
        .map(([k, t]) => ({ mm: k.split("x").map(Number), n: t.n, ...(t.pags.length ? { pags: t.pags } : {}) }))
        .sort((a, b) => b.n - a.n)
        .slice(0, 40),
    };
  } catch (err) {
    console.error(`[analizarPdf] Falló para el archivo ${fileId}:`, err.message || err);
    const msg = String(err && err.message ? err.message : err);
    let motivo = "no se pudo leer el PDF (archivo dañado o con formato no estándar)";
    if (/timeout/i.test(msg)) motivo = "la descarga desde Drive tardó demasiado";
    else if (/encrypt|password/i.test(msg)) motivo = "PDF protegido con contraseña";
    else if (/403|404|not found|permission/i.test(msg)) motivo = "sin acceso al archivo en Drive";
    return { error: motivo };
  }
}
