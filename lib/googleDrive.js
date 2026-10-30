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
  "files(id,name,mimeType,parents,trashed,createdTime,modifiedTime,md5Checksum,size,shortcutDetails(targetId,targetMimeType),webViewLink,thumbnailLink,lastModifyingUser(displayName,emailAddress,permissionId),sharingUser(displayName,emailAddress,permissionId),owners(displayName,emailAddress,permissionId))";

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

    // Se prefiere la actividad de BORRAR o MOVER (esa es la que dice quién fue);
    // si no hay, la más reciente que tenga una persona identificada.
    const esAccionDeBorrarMover = (act) => {
      const k = Object.keys(act.primaryActionDetail || {})[0];
      return k === "delete" || k === "move";
    };
    const ordenadas = [...activities.filter(esAccionDeBorrarMover), ...activities.filter((a) => !esAccionDeBorrarMover(a))];
    for (const act of ordenadas) {
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

// Personas con acceso a las carpetas (permisos): sirve para traducir el código
// "people/123456" que entrega Drive Activity a "Nombre (correo)". El id de
// permiso de Drive es el mismo número que el de la persona.
export async function getPersonasConAcceso(carpetaIds) {
  const drive = getDriveClient();
  const mapa = new Map();
  await Promise.all(
    carpetaIds.map(async (id) => {
      try {
        let pageToken;
        do {
          const r = await drive.permissions.list({
            fileId: id,
            fields: "nextPageToken,permissions(id,type,displayName,emailAddress)",
            pageSize: 100,
            pageToken,
            supportsAllDrives: true,
          });
          for (const p of r.data.permissions || []) {
            if (p.type === "user" && p.id && (p.displayName || p.emailAddress)) {
              mapa.set(
                String(p.id),
                p.displayName && p.emailAddress ? `${p.displayName} (${p.emailAddress})` : p.emailAddress || p.displayName
              );
            }
          }
          pageToken = r.data.nextPageToken;
        } while (pageToken);
      } catch (err) {
        console.log(`[personas] ${id}: ${String((err && err.message) || err).slice(0, 120)}`);
      }
    })
  );
  return mapa;
}

// Historial REAL de actividad de Drive (Drive Activity API): cada acción con su
// hora exacta y quién la hizo. Se consulta una vez por carpeta principal, desde
// la fecha indicada. Devuelve [{ id, tipo, time, personName }] donde tipo es
// create | edit | move | delete | rename | restore, y personName es el código
// "people/123…" de la persona (no trae el correo).
export async function getActividadDrive(carpetaIds, desde, maxMs = 15000) {
  const t0 = Date.now();
  const resultados = [];
  const vistos = new Set();
  const activity = getActivityClient();

  async function consultar(folderId) {
    let pageToken;
    let paginas = 0;
    try {
      do {
        if (Date.now() - t0 > maxMs) return;
        const res = await activity.activity.query({
          requestBody: {
            ancestorName: `items/${folderId}`,
            filter: `time >= "${desde.toISOString()}"`,
            pageSize: 100,
            pageToken,
          },
        });
        for (const act of res.data.activities || []) {
          const tipo = Object.keys(act.primaryActionDetail || {})[0] || null;
          const t = act.timestamp || (act.timeRange && act.timeRange.endTime) || null;
          const actor = (act.actors || []).find((x) => x?.user?.knownUser);
          const personName = actor ? actor.user.knownUser.personName || null : null;
          const a0 = (act.actors || [])[0] || {};
          const actorTipo = actor ? "usuario" : a0.administrator ? "administrador" : a0.system ? "sistema" : a0.impersonation ? "suplantacion" : a0.user && a0.user.unknownUser ? "desconocido" : "otro";
          for (const tg of act.targets || []) {
            const nombre = (tg.driveItem && tg.driveItem.name) || "";
            const id = nombre.replace(/^items\//, "");
            if (!id || !t || !tipo) continue;
            const clave = `${id}|${tipo}|${t}`;
            if (vistos.has(clave)) continue;
            vistos.add(clave);
            resultados.push({ id, tipo, time: new Date(t), personName, actorTipo });
          }
        }
        pageToken = res.data.nextPageToken;
        paginas++;
      } while (pageToken && paginas < 10);
    } catch (err) {
      console.log(`[actividad] ${folderId}: ${String((err && err.message) || err).slice(0, 160)}`);
    }
  }

  await Promise.all(carpetaIds.map((id) => consultar(id)));
  return resultados;
}

// Acción de BORRAR o MOVER de un elemento concreto, según Drive Activity.
// Devuelve { tipo: "delete"|"move"|null, time, personName, actorTipo, total } o { error }.
// Solo cuenta acciones de borrar/mover: nunca se toma a quien lo subió como autor.
export async function getAccionBorrarMover(itemId) {
  try {
    const activity = getActivityClient();
    const res = await activity.activity.query({
      requestBody: { itemName: `items/${itemId}`, pageSize: 20 },
    });
    const acts = res.data.activities || [];
    for (const act of acts) {
      const tipo = Object.keys(act.primaryActionDetail || {})[0] || null;
      if (tipo !== "delete" && tipo !== "move") continue;
      const t = act.timestamp || (act.timeRange && act.timeRange.endTime) || null;
      const actor = (act.actors || []).find((x) => x?.user?.knownUser);
      const a0 = (act.actors || [])[0] || {};
      const actorTipo = actor ? "usuario" : a0.administrator ? "administrador" : a0.system ? "sistema" : a0.impersonation ? "suplantacion" : a0.user && a0.user.unknownUser ? "desconocido" : "otro";
      return {
        tipo,
        time: t ? new Date(t) : null,
        personName: actor ? actor.user.knownUser.personName || null : null,
        actorTipo,
        total: acts.length,
      };
    }
    return { tipo: null, time: null, personName: null, actorTipo: null, total: acts.length };
  } catch (err) {
    return { error: String((err && err.message) || err).slice(0, 120) };
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
