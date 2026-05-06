import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore, doc, getDoc, setDoc, onSnapshot, collection, addDoc, deleteDoc, updateDoc, getDocs } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyAIvUvWCInpYNDRdSkUO7YLW-F4POxxgFg",
  authDomain: "portcity-logistics-dashboard.firebaseapp.com",
  projectId: "portcity-logistics-dashboard",
  storageBucket: "portcity-logistics-dashboard.firebasestorage.app",
  messagingSenderId: "1038238701860",
  appId: "1:1038238701860:web:b21c837d9d6750ae263bb1"
};

const app = initializeApp(firebaseConfig);
const db  = getFirestore(app);

// ── Sync status helper ────────────────────────────────────────────────────────
function setSyncStatus(state) {
  const el = document.getElementById('sync-status');
  el.className = 'synced';
  if (state === 'saving') { el.textContent = '● Saving…'; el.className = 'saving'; }
  else if (state === 'error') { el.textContent = '● Error'; el.className = 'error'; }
  else { el.textContent = '● Live'; el.className = 'synced'; }
}

async function saveDoc(path, data) {
  setSyncStatus('saving');
  try {
    await setDoc(doc(db, path), data, { merge: true });
    setSyncStatus('synced');
  } catch(e) { setSyncStatus('error'); console.error(e); }
}

// ── Expose to global scope ───────────────────────────────────────────────────
window._db = db;
window._doc = doc;
window._setDoc = setDoc;
window._addDoc = addDoc;
window._deleteDoc = deleteDoc;
window._updateDoc = updateDoc;
window._getDocs = getDocs;
window._collection = collection;
window._onSnapshot = onSnapshot;
window._getDoc = getDoc;
window._saveDoc = saveDoc;
window._setSyncStatus = setSyncStatus;

// ── Init all listeners once Firebase is ready ─────────────────────────────────
window.addEventListener('firebase-ready', () => {
  initDashboard();
});
window.dispatchEvent(new Event('firebase-ready'));