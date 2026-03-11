import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, signOut, onAuthStateChanged }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { getFirestore, doc, setDoc, getDoc, onSnapshot }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

// ⚠️ FIX#3: Firebase config — these keys are client-side public keys (safe for web),
// BUT restrict them in Firebase Console:
// 1. Authentication → Authorized domains (add your domain only)
// 2. Firestore → Rules (users can only read/write their own doc)
// 3. API Key → Application restrictions (HTTP referrers: your domain)
const firebaseConfig = {
  apiKey: window.__FIREBASE_API_KEY__ || "AIzaSyDnKud5gR8a_Fyq8cNdzgHMNQw4GMuX0-Q",
  authDomain: "azem-ad49b.firebaseapp.com",
  projectId: "azem-ad49b",
  storageBucket: "azem-ad49b.firebasestorage.app",
  messagingSenderId: "509862794540",
  appId: "1:509862794540:web:aeeec0aa12c5b859a98889",
  measurementId: "G-GHZVGY5ZWB"
};

const app  = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db   = getFirestore(app);

let _fbUid = null;
let _syncUnsubscribe = null;
let _syncDebounce = null;

// ── Expose to global scope ──
window.firebaseSignIn = async function() {
  try {
    const btn = document.getElementById('google-signin-btn');
    if (btn) { btn.textContent = '⏳ جارٍ التسجيل...'; btn.disabled = true; }
    const provider = new GoogleAuthProvider();
    // Use redirect on standalone PWA, popup on desktop browser
    const isStandalone = window.matchMedia('(display-mode: standalone)').matches
      || window.navigator.standalone === true;
    if (isStandalone) {
      await signInWithRedirect(auth, provider);
    } else {
      const result = await signInWithPopup(auth, provider);
      if (result?.user) {
        showMiniToast(\`☁️ مرحباً \${result.user.displayName?.split(' ')[0] || ''}!\`);
      }
    }
  } catch(e) {
    if (e.code === 'auth/popup-closed-by-user') {
      const btn = document.getElementById('google-signin-btn');
      if (btn) { btn.textContent = 'تسجيل الدخول بـ Google'; btn.disabled = false; }
      return;
    }
    showMiniToast('⚠️ فشل تسجيل الدخول: ' + (e.message || e.code));
    const btn = document.getElementById('google-signin-btn');
    if (btn) { btn.textContent = 'تسجيل الدخول بـ Google'; btn.disabled = false; }
  }
};

window.firebaseSignOut = async function() {
  if (_syncUnsubscribe) { _syncUnsubscribe(); _syncUnsubscribe = null; }
  await signOut(auth);
  _fbUid = null;
  updateAuthUI(null);
  showMiniToast('👋 تم تسجيل الخروج');
};

window.firebaseSyncNow = async function() {
  if (!_fbUid) { showMiniToast('⚠️ سجّل دخولك أولاً'); return; }
  await pushToCloud();
};

async function pushToCloud() {
  if (!_fbUid) return;
  try {
    // Strip large base64 images to avoid Firestore 1MB limit
    const payload = JSON.parse(JSON.stringify(S));
    if (payload.customImages) {
      Object.keys(payload.customImages).forEach(k => {
        if (payload.customImages[k] && payload.customImages[k].length > 10000) {
          delete payload.customImages[k];
        }
      });
    }
    payload._syncedAt = Date.now();
    await setDoc(doc(db, 'users', _fbUid), { state: payload }, { merge: true });
    const statusEl = document.getElementById('firebase-sync-status');
    if (statusEl) {
      const t = new Date().toLocaleTimeString('ar-SA');
      statusEl.textContent = `✅ مزامن · ${t}`;
    }
  } catch(e) {
    console.warn('Firebase push error:', e);
  }
}

async function pullFromCloud(uid) {
  try {
    const snap = await getDoc(doc(db, 'users', uid));
    if (snap.exists()) {
      const remote = snap.data().state;
      if (remote) {
        // Merge remote over local — remote wins on conflicts
        const merged = { ...S, ...remote };
        // Keep local unsaved progress if newer
        if ((S._localTs||0) > (remote._syncedAt||0)) {
          Object.assign(merged, { calories: Math.max(S.calories||0, remote.calories||0) });
        }
        Object.assign(S, merged);
        saveState();
        try { render(); } catch(e) {}
        showMiniToast('☁️ تم جلب بياناتك من السحاب!');
      }
    } else {
      // No cloud data yet — push local to cloud
      await pushToCloud();
    }
  } catch(e) {
    console.warn('Firebase pull error:', e);
  }
}

function updateAuthUI(user) {
  const signinBtn  = document.getElementById('google-signin-btn');
  const userArea   = document.getElementById('firebase-user-area');
  const nameEl     = document.getElementById('firebase-user-name');
  const photoEl    = document.getElementById('firebase-user-photo');
  // Header auth badge
  const hdrBtn     = document.getElementById('hdr-auth-btn');
  const hdrIcon    = document.getElementById('hdr-auth-icon');
  const hdrAvatar  = document.getElementById('hdr-user-avatar');

  if (user) {
    // Update settings sheet
    if (signinBtn) signinBtn.style.display = 'none';
    if (userArea) userArea.style.display = 'block';
    if (nameEl)  nameEl.textContent = user.displayName || user.email || '';
    if (photoEl && user.photoURL) photoEl.src = user.photoURL;
    // Update header badge — show avatar, hide icon
    if (hdrBtn) { hdrBtn.style.background='transparent'; hdrBtn.style.border='none'; hdrBtn.title='متصل: '+(user.displayName||user.email||''); }
    if (hdrIcon) hdrIcon.style.display = 'none';
    if (hdrAvatar && user.photoURL) { hdrAvatar.src=user.photoURL; hdrAvatar.style.display='block'; }
    else if (hdrAvatar) { hdrAvatar.style.display='none'; if(hdrIcon) hdrIcon.style.display='block'; }
  } else {
    // Update settings sheet
    if (signinBtn) { signinBtn.style.display = 'flex'; signinBtn.textContent = 'تسجيل الدخول بـ Google'; signinBtn.disabled = false; }
    if (userArea) userArea.style.display = 'none';
    // Update header badge — show Google icon
    if (hdrBtn) { hdrBtn.style.background='rgba(66,133,244,.15)'; hdrBtn.style.border='1.5px solid rgba(66,133,244,.5)'; hdrBtn.title='تسجيل الدخول بـ Google'; }
    if (hdrIcon) hdrIcon.style.display = 'block';
    if (hdrAvatar) hdrAvatar.style.display = 'none';
  }
}

// ── Intercept saveState to auto-sync ──
const _origSaveState = window.saveState;
window.saveState = function() {
  if (typeof _origSaveState === 'function') _origSaveState();
  if (_fbUid) {
    clearTimeout(_syncDebounce);
    _syncDebounce = setTimeout(pushToCloud, 2500); // debounce 2.5s
  }
};

// ── Handle redirect result (PWA standalone sign-in) ──
getRedirectResult(auth).then(result => {
  if (result?.user) {
    showMiniToast(`☁️ مرحباً ${result.user.displayName?.split(' ')[0] || ''}! تم تسجيل الدخول`);
  }
}).catch(e => {
  if (e.code !== 'auth/no-current-user') console.warn('Redirect result error:', e);
});

// ── Auth state listener ──
onAuthStateChanged(auth, async user => {
  if (user) {
    _fbUid = user.uid;
    updateAuthUI(user);
    await pullFromCloud(user.uid);
    // Show sync indicator in header if there's one
    showMiniToast(`☁️ مرحباً ${user.displayName?.split(' ')[0] || ''}! بياناتك تُزامن تلقائياً`);
  } else {
    _fbUid = null;
    updateAuthUI(null);
  }
});

// Re-run updateAuthUI each time settings sheet opens
const _origOpenSettings = window.openSettingsSheet;
window.openSettingsSheet = function() {
  if (typeof _origOpenSettings === 'function') _origOpenSettings();
  setTimeout(() => updateAuthUI(auth.currentUser), 100);
};
