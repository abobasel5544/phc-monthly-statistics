import { initializeApp, getApps, getApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyD3B2XrqLtDfMScmeBalJDs-x1yv18PCGA",
  authDomain: "phc-monthly-statistics.firebaseapp.com",
  projectId: "phc-monthly-statistics",
  storageBucket: "phc-monthly-statistics.firebasestorage.app",
  messagingSenderId: "910166212485",
  appId: "1:910166212485:web:57a5610707d36f797bceec",
};

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);

export const db = getFirestore(app);
