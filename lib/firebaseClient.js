import { initializeApp, getApps } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDBDykR2wKiLv6meb_GOIokj7i17MLUFGk",
  authDomain: "iep-de-los-uros.firebaseapp.com",
  databaseURL: "https://iep-de-los-uros-default-rtdb.firebaseio.com",
  projectId: "iep-de-los-uros",
  storageBucket: "iep-de-los-uros.firebasestorage.app",
  messagingSenderId: "224589046311",
  appId: "1:224589046311:web:b7eb0e187ec9695a8aa402",
  measurementId: "G-BSJDEF1SEZ"
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApps()[0];
export const auth = getAuth(app);
export const db = getFirestore(app);
