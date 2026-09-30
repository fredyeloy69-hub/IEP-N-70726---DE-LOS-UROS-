import { google } from "googleapis";
import { PDFDocument } from "pdf-lib";

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
  "files(id,name,mimeType,parents,trashed,modifiedTime,md5Checksum,webViewLink,thumbnailLink,lastModifyingUser(displayName,emailAddress))";

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
      for (const file of items) {
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
