import { google } from "googleapis";
import { PDFDocument } from "pdf-lib";
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
  try {
    const activity = getActivityClient();

    const res = await activity.activity.query({
      requestBody: {
        itemName: `items/${itemId}`,
        pageSize: 10,
      },
    });

    const activities =
      res.data.activities || [];

    for (const act of activities) {
      const actor = (act.actors || []).find(
        (a) => a?.user?.knownUser
      );

      if (actor) {
        const ku = actor.user.knownUser;

        return ku.personName || null;
      }
    }

    return null;
  } catch (err) {
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
 *   `aprox` lista los tamaños que no son de la serie A (se les asignó la hoja
 *   en la que caben). null si el PDF no se pudo abrir.
 */
export async function analizarPdf(fileId, timeoutMs = 25000) {
  try {
    const drive = getDriveClient();

    const descarga = drive.files.get(
      { fileId, alt: "media" },
      { responseType: "arraybuffer" }
    );
    const timeout = new Promise((_, reject) =>
      setTimeout(() => reject(new Error("timeout")), timeoutMs)
    );
    const res = await Promise.race([descarga, timeout]);
    const bytes = Buffer.from(res.data);

    const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    const paginasPdf = pdf.getPages();

    const formatos = { A4: 0, A3: 0, A2: 0, A1: 0, A0: 0 };
    const aproxMap = new Map();
    for (const pagina of paginasPdf) {
      let ancho, alto;
      try {
        const box = pagina.getCropBox();
        ancho = box.width;
        alto = box.height;
      } catch {
        const s = pagina.getSize();
        ancho = s.width;
        alto = s.height;
      }
      if (!(ancho > 0) || !(alto > 0)) {
        formatos.A4 += 1;
        continue;
      }
      const c = clasificarHoja(ancho, alto);
      formatos[c.formato] += 1;
      if (!c.exacto) {
        const clave = `${c.mm[0]}x${c.mm[1]}→${c.formato}`;
        const previo = aproxMap.get(clave) || { mm: `${c.mm[0]}x${c.mm[1]}`, formato: c.formato, n: 0 };
        previo.n += 1;
        aproxMap.set(clave, previo);
      }
    }

    return {
      paginas: paginasPdf.length,
      formatos,
      aprox: [...aproxMap.values()].slice(0, 6),
    };
  } catch (err) {
    console.error(`[analizarPdf] Falló para el archivo ${fileId}:`, err.message || err);
    return null;
  }
}
