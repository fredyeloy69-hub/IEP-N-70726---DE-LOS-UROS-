import { initializeApp, getApps } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBRn71tbELaxkhtpoukvKj_e8D2VHbNgtM",
  authDomain: "exp-cs-acaso-i-2.firebaseapp.com",
  projectId: "exp-cs-acaso-i-2",
  storageBucket: "exp-cs-acaso-i-2.firebasestorage.app",
  messagingSenderId: "122866621409",
  appId: "1:122866621409:web:6b41dd09fe3f7f7a4f50a0",
  measurementId: "G-ER10ZXNEWF"
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApps()[0];
export const auth = getAuth(app);
export const db = getFirestore(app);
