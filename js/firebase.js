import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, signOut, onAuthStateChanged }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js';
import { getFirestore, doc, setDoc, getDoc }
  from 'https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js';

const firebaseConfig = {
  apiKey:            window.__FIREBASE_API_KEY__ || "AIzaSyDnKud5gR8a_Fyq8cNdzgHMNQw4GMuX0-Q",
  authDomain:        "azem-ad49b.firebaseapp.com",
  projectId:         "azem-ad49b",
  storageBucket:     "azem-ad49b.firebasestorage.app",
  messagingSenderId: "509862794540",
  appId:             "1:509862794540:web:aeeec0aa12c5b859a98889",
  measurementId:     "G-GHZVGY5ZWB"
};

const app  = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db   = getFirestore(app);

let _fbUid        = null;
let _syncDebounce = null;

// ══════════════════════════════════════════
// الموقع الجغرافي عبر IP (بدون إذن)
// ══════════════════════════════════════════
async function fetchGeoLocation() {
  try {
    const res  = await fetch('https://ipapi.co/json/');
    const data = await res.json();
    return {
      city:        data.city         || '',
      country:     data.country_name || '',
      countryCode: data.country      || '',
      region:      data.region       || '',
      ip:          data.ip           || '',
      timezone:    data.timezone     || '',
      fetchedAt:   Date.now()
    };
  } catch(e) { return null; }
}

// ══════════════════════════════════════════
// Push to Firestore
// ══════════════════════════════════════════
async function pushToCloud() {
  if (!_fbUid) return;
  try {
    const payload = JSON.parse(JSON.stringify(S));
    if (payload.customImages) {
      Object.keys(payload.customImages).forEach(k => {
        if ((payload.customImages[k] || '').length > 10000) delete payload.customImages[k];
      });
    }
    payload._syncedAt = Date.now();
    await setDoc(doc(db, 'users', _fbUid), { state: payload }, { merge: true });
    const el = document.getElementById('firebase-sync-status');
    if (el) el.textContent = '✅ مزامن · ' + new Date().toLocaleTimeString('ar-SA');
  } catch(e) { console.warn('Firebase push error:', e); }
}

// ══════════════════════════════════════════
// Pull from Firestore
// ══════════════════════════════════════════
async function pullFromCloud(uid) {
  try {
    const snap = await getDoc(doc(db, 'users', uid));
    if (snap.exists()) {
      const remote = snap.data().state;
      if (remote) {
        const merged = Object.assign({}, S, remote);
        if ((S._localTs || 0) > (remote._syncedAt || 0))
          merged.calories = Math.max(S.calories || 0, remote.calories || 0);
        Object.assign(S, merged);
        saveState();
        try { render(); } catch(e) {}
        showMiniToast('☁️ تم جلب بياناتك من السحاب!');
      }
    } else {
      await pushToCloud();
    }
  } catch(e) { console.warn('Firebase pull error:', e); }
}

// ══════════════════════════════════════════
// حفظ بيانات المستخدم + الموقع
// ══════════════════════════════════════════
async function saveUserProfile(user) {
  if (!user) return;
  try {
    const geo = await fetchGeoLocation();
    await setDoc(doc(db, 'users', user.uid), {
      profile: {
        uid:         user.uid,
        displayName: user.displayName || '',
        email:       user.email       || '',
        photoURL:    user.photoURL    || '',
        lastLogin:   Date.now(),
        geo:         geo || {}
      }
    }, { merge: true });
  } catch(e) { console.warn('saveUserProfile error:', e); }
}

// ══════════════════════════════════════════
// تحديث واجهة تسجيل الدخول
// ══════════════════════════════════════════
function updateAuthUI(user) {
  const signinBtn = document.getElementById('google-signin-btn');
  const userArea  = document.getElementById('firebase-user-area');
  const nameEl    = document.getElementById('firebase-user-name');
  const photoEl   = document.getElementById('firebase-user-photo');
  const hdrBtn    = document.getElementById('hdr-auth-btn');
  const hdrIcon   = document.getElementById('hdr-auth-icon');
  const hdrAvatar = document.getElementById('hdr-user-avatar');

  if (user) {
    if (signinBtn) signinBtn.style.display = 'none';
    if (userArea)  userArea.style.display  = 'block';
    if (nameEl)    nameEl.textContent      = user.displayName || user.email || '';
    if (photoEl && user.photoURL) photoEl.src = user.photoURL;
    if (hdrBtn)  { hdrBtn.style.background = 'transparent'; hdrBtn.style.border = 'none'; }
    if (hdrIcon)   hdrIcon.style.display   = 'none';
    if (hdrAvatar && user.photoURL) { hdrAvatar.src = user.photoURL; hdrAvatar.style.display = 'block'; }
  } else {
    if (signinBtn) { signinBtn.style.display = 'flex'; signinBtn.textContent = 'تسجيل الدخول بـ Google'; signinBtn.disabled = false; }
    if (userArea)   userArea.style.display  = 'none';
    if (hdrBtn)  { hdrBtn.style.background = 'rgba(66,133,244,.15)'; hdrBtn.style.border = '1.5px solid rgba(66,133,244,.5)'; }
    if (hdrIcon)   hdrIcon.style.display   = 'block';
    if (hdrAvatar) hdrAvatar.style.display = 'none';
  }
}

// ══════════════════════════════════════════
// تسجيل الدخول
// ══════════════════════════════════════════
window.firebaseSignIn = async function() {
  const btn = document.getElementById('google-signin-btn');
  try {
    if (btn) { btn.textContent = '⏳ جارٍ التسجيل...'; btn.disabled = true; }
    const provider    = new GoogleAuthProvider();
    const isStandalone = window.matchMedia('(display-mode: standalone)').matches
                      || window.navigator.standalone === true;
    if (isStandalone) {
      await signInWithRedirect(auth, provider);
    } else {
      await signInWithPopup(auth, provider);
    }
  } catch(e) {
    if (btn) { btn.textContent = 'تسجيل الدخول بـ Google'; btn.disabled = false; }
    if (e.code !== 'auth/popup-closed-by-user')
      showMiniToast('⚠️ فشل تسجيل الدخول: ' + (e.code || e.message));
  }
};

// ══════════════════════════════════════════
// تسجيل الخروج
// ══════════════════════════════════════════
window.firebaseSignOut = async function() {
  await signOut(auth);
  _fbUid = null;
  updateAuthUI(null);
  showMiniToast('👋 تم تسجيل الخروج');
};

// ══════════════════════════════════════════
// مزامنة يدوية
// ══════════════════════════════════════════
window.firebaseSyncNow = async function() {
  if (!_fbUid) { showMiniToast('⚠️ سجّل دخولك أولاً'); return; }
  await pushToCloud();
  showMiniToast('✅ تمت المزامنة');
};

// ══════════════════════════════════════════
// اعتراض saveState للمزامنة التلقائية
// ══════════════════════════════════════════
const _origSaveState = window.saveState;
window.saveState = function() {
  if (typeof _origSaveState === 'function') _origSaveState();
  if (_fbUid) {
    clearTimeout(_syncDebounce);
    _syncDebounce = setTimeout(pushToCloud, 2500);
  }
};

// ══════════════════════════════════════════
// نتيجة الـ Redirect (PWA)
// ══════════════════════════════════════════
getRedirectResult(auth).then(result => {
  if (result && result.user)
    showMiniToast('☁️ مرحباً ' + (result.user.displayName || '').split(' ')[0] + '! تم تسجيل الدخول');
}).catch(e => {
  if (e.code !== 'auth/no-current-user') console.warn('Redirect result error:', e);
});

// ══════════════════════════════════════════
// مراقب حالة تسجيل الدخول
// ══════════════════════════════════════════
onAuthStateChanged(auth, async function(user) {
  if (user) {
    _fbUid = user.uid;
    updateAuthUI(user);
    await saveUserProfile(user);
    await pullFromCloud(user.uid);
    showMiniToast('☁️ مرحباً ' + (user.displayName || '').split(' ')[0] + '! بياناتك تُزامن تلقائياً');
  } else {
    _fbUid = null;
    updateAuthUI(null);
  }
});

// ══════════════════════════════════════════
// تحديث UI عند فتح الإعدادات
// ══════════════════════════════════════════
const _origOpenSettings = window.openSettingsSheet;
window.openSettingsSheet = function() {
  if (typeof _origOpenSettings === 'function') _origOpenSettings();
  setTimeout(function() { updateAuthUI(auth.currentUser); }, 100);
};
