import { google } from "googleapis";
import { PDFDocument } from "pdf-lib";

function getAuth() {
  const key = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  return new google.auth.JWT({
    email: key.client_email,
    key: key.private_key,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/drive.activity.readonly",
    ],
  });
}

export function getDriveClient() {
  return google.drive({ version: "v3", auth: getAuth() });
}

export function getActivityClient() {
  return google.driveactivity({ version: "v2", auth: getAuth() });
}

const FIELDS =
  "files(id,name,mimeType,parents,trashed,modifiedTime,md5Checksum,webViewLink,thumbnailLink,lastModifyingUser(displayName,emailAddress))";

// Trae todos los items (archivos y subcarpetas) dentro de UNA carpeta.
// Maneja paginacion si hay mas de 1000 items en esa carpeta.
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

// Trae TODOS los items (carpetas y archivos) debajo de rootId, recursivamente.
// Version paralelizada: procesa por niveles (BFS) en vez de uno por uno (DFS secuencial).
// En cada nivel, pide TODAS las carpetas de ese nivel al mismo tiempo con Promise.all,
// en vez de esperar cada respuesta antes de pedir la siguiente carpeta.
export async function scanDriveTree(rootId) {
  const drive = getDriveClient();
  const allItems = [];
  const visited = new Set([rootId]);

  let currentLevel = [rootId];

  while (currentLevel.length > 0) {
    // Pide todas las carpetas del nivel actual EN PARALELO
    const results = await Promise.all(
      currentLevel.map((folderId) => listFolderContents(drive, folderId))
    );

    const nextLevel = [];

    for (const items of results) {
      for (const file of items) {
        allItems.push(file);
        if (
          file.mimeType === "application/vnd.google-apps.folder" &&
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

// Intento best-effort de identificar quien borro un archivo, usando Drive Activity API.
// LIMITACION CONOCIDA: en cuentas de Gmail personales (sin dominio de Google Workspace),
// Google generalmente no permite resolver el nombre/correo de la persona por privacidad,
// aunque el evento de borrado si haya quedado registrado. En esos casos esta funcion
// devuelve null y el codigo que la llama debe usar "Desconocido" como respaldo.
export async function getUltimoActorDeItem(itemId) {
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
        // A veces Google resuelve un nombre visible directamente; si no, no hay forma
        // de obtener mas informacion sin la People API (que requiere permisos que un
        // service account normalmente no tiene sobre cuentas personales ajenas).
        return ku.personName || null;
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

// Descarga el contenido real de un PDF de Drive y cuenta sus páginas —
// necesario para el reporte de material de impresión: un archivo puede tener
// 1 página o 200, y "1 archivo = 1 hoja" estaría mal. Devuelve null si el
// PDF no se pudo leer (protegido con contraseña, corrupto, un PDF escaneado
// mal generado, etc.) para que el que llama lo marque como "revisar a mano"
// en vez de contarlo incorrectamente.
export async function contarPaginasPdf(fileId) {
  try {
    const drive = getDriveClient();
    const descarga = drive.files.get(
      { fileId, alt: "media" },
      { responseType: "arraybuffer" }
    );
    // Límite de tiempo por archivo — la ruta que llama a esto tiene
    // maxDuration=60 configurado en Vercel (60 segundos reales, no los 10 por
    // defecto), así que hay bastante margen. Como todos los PDFs de la
    // corrida se descargan EN PARALELO, este límite no se suma entre
    // archivos — el tiempo total depende del más lento, no de la cantidad.
    const timeout = new Promise((_, reject) =>
      setTimeout(() => reject(new Error("timeout")), 25000)
    );
    const res = await Promise.race([descarga, timeout]);
    const bytes = Buffer.from(res.data);
    const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
    return pdf.getPageCount();
  } catch (err) {
    console.error(`[contarPaginasPdf] Falló para el archivo ${fileId}:`, err.message || err);
    return null;
  }
}
