import { initializeApp, getApps, cert } from "firebase-admin/app";
import { getFirestore } from "firebase-admin/firestore";

function getAdminApp() {
  if (getApps().length) {
    return getApps()[0];
  }

  const serviceAccountJson =
    process.env.FIREBASE_SERVICE_ACCOUNT_KEY;

  if (!serviceAccountJson) {
    throw new Error(
      "Falta FIREBASE_SERVICE_ACCOUNT_KEY"
    );
  }

  let serviceAccount;

  try {
    serviceAccount = JSON.parse(serviceAccountJson);
  } catch (error) {
    throw new Error(
      "FIREBASE_SERVICE_ACCOUNT_KEY no contiene un JSON válido"
    );
  }

  if (!serviceAccount.project_id) {
    throw new Error(
      "FIREBASE_SERVICE_ACCOUNT_KEY no contiene project_id"
    );
  }

  if (!serviceAccount.client_email) {
    throw new Error(
      "FIREBASE_SERVICE_ACCOUNT_KEY no contiene client_email"
    );
  }

  if (!serviceAccount.private_key) {
    throw new Error(
      "FIREBASE_SERVICE_ACCOUNT_KEY no contiene private_key"
    );
  }

  return initializeApp({
    credential: cert(serviceAccount),
  });
}

export const adminDb =
  getFirestore(getAdminApp());
