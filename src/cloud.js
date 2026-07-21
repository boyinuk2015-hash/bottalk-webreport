// ── Firebase / Firestore cloud sync ───────────────────────────────────────────
// Stores the whole report (sales + area data) in one document so every device
// that opens the site sees the same, latest data in real time.
import { initializeApp } from "firebase/app";
import { getFirestore, doc, onSnapshot, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyARLimPay2cg7cQoE758YEw5O7i-bDG7LA",
  authDomain: "bottalk-report.firebaseapp.com",
  projectId: "bottalk-report",
  storageBucket: "bottalk-report.firebasestorage.app",
  messagingSenderId: "872281424010",
  appId: "1:872281424010:web:d9f4262f402f88d6fe2d73",
};

let reportDoc = null;
try {
  const app = initializeApp(firebaseConfig);
  const db = getFirestore(app);
  reportDoc = doc(db, "bottalk", "report");
} catch (e) {
  reportDoc = null; // cloud disabled (e.g. blocked network) — app still works locally
}

export const cloudEnabled = !!reportDoc;

// Subscribe to the shared report. onData(payloadObj|null) fires on every change.
export function subscribeReport(onData, onError) {
  if (!reportDoc) return () => {};
  return onSnapshot(
    reportDoc,
    (snap) => {
      if (!snap.exists()) { onData(null); return; }
      const raw = snap.data() || {};
      let payload = null;
      try { payload = raw.payload ? JSON.parse(raw.payload) : null; } catch (e) { payload = null; }
      onData(payload);
    },
    (err) => { if (onError) onError(err); }
  );
}

// Save the full report payload to the cloud (last write wins).
export function saveReport(payloadObj) {
  if (!reportDoc) return Promise.resolve();
  return setDoc(reportDoc, { payload: JSON.stringify(payloadObj), updatedAt: Date.now() }, { merge: true });
}
