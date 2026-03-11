// ============================================================
// AZEM (عزم) — Service Worker v6
// Strategy: Cache-First (assets) + Network-First (fonts)
// ============================================================

const APP_CACHE  = 'azem-app-v12';
const FONT_CACHE = 'azem-fonts-v2';
const VERSION    = '7.0.0';

// Core files — must be cached at install time
const CORE_ASSETS = [
  './index.html',
  './gifs.js',
  './manifest.json',
  './sw.js',
  './icon-192-1.png',
  './icon-512-1.png',
  './css/styles.css',
  './js/data.js',
  './js/audio.js',
  './js/i18n.js',
  './js/timer.js',
  './js/session.js',
  './js/render.js',
  './js/editor.js',
  './js/ui.js',
  './js/coach.js',
  './js/firebase.js',
];

// ---- Install: cache all core assets ----
self.addEventListener('install', e => {
  console.log(`[SW ${VERSION}] Installing...`);
  e.waitUntil(
    caches.open(APP_CACHE)
      .then(cache => cache.addAll(CORE_ASSETS))
      .then(() => {
        console.log(`[SW ${VERSION}] Core assets cached ✓`);
        return self.skipWaiting();
      })
  );
});

// ---- Activate: delete old caches ----
self.addEventListener('activate', e => {
  console.log(`[SW ${VERSION}] Activating...`);
  e.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys
        .filter(k => k !== APP_CACHE && k !== FONT_CACHE)
        .map(k => {
          console.log(`[SW] Deleting old cache: ${k}`);
          return caches.delete(k);
        })
    )).then(() => self.clients.claim())
  );
});

// ---- Fetch: smart routing ----
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Skip non-GET and chrome-extension
  if(e.request.method !== 'GET') return;
  if(url.protocol === 'chrome-extension:') return;

  // Google Fonts — Stale-While-Revalidate
  if(url.hostname.includes('fonts.googleapis.com') ||
     url.hostname.includes('fonts.gstatic.com')){
    e.respondWith(staleWhileRevalidate(e.request, FONT_CACHE));
    return;
  }

  // Core app assets — Cache-First
  if(url.origin === self.location.origin){
    e.respondWith(cacheFirst(e.request));
    return;
  }

  // Everything else — Network with cache fallback
  e.respondWith(networkFirst(e.request));
});

// ---- Strategies ----

// Cache-First: serve from cache, background update
async function cacheFirst(req){
  const cached = await caches.match(req);
  if(cached){
    // Background update
    update(req, APP_CACHE);
    return cached;
  }
  return networkWithCacheSave(req, APP_CACHE);
}

// Network-First: try network, fall back to cache
async function networkFirst(req){
  try{
    const res = await fetch(req);
    if(res && res.ok){
      const cache = await caches.open(APP_CACHE);
      cache.put(req, res.clone());
    }
    return res;
  } catch(_){
    const cached = await caches.match(req);
    return cached || offlineFallback();
  }
}

// Stale-While-Revalidate: return cached immediately, update in background
async function staleWhileRevalidate(req, cacheName){
  const cache = await caches.open(cacheName);
  const cached = await cache.match(req);
  const fetchPromise = fetch(req).then(res => {
    if(res && res.ok) cache.put(req, res.clone());
    return res;
  }).catch(()=>{});
  return cached || fetchPromise;
}

async function networkWithCacheSave(req, cacheName){
  try{
    const res = await fetch(req);
    if(res && res.ok){
      const cache = await caches.open(cacheName);
      cache.put(req, res.clone());
    }
    return res;
  } catch(_){
    return offlineFallback();
  }
}

// Background cache update (don't wait)
function update(req, cacheName){
  fetch(req).then(res => {
    if(res && res.ok){
      caches.open(cacheName).then(c => c.put(req, res.clone()));
    }
  }).catch(()=>{});
}

// Offline fallback page
async function offlineFallback(){
  const cached = await caches.match('./index.html');
  if(cached) return cached;
  return new Response(
    `<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8">
    <title>AZEM — غير متصل</title>
    <style>body{background:#07090F;color:#D4A843;font-family:sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100vh;margin:0;text-align:center}h1{font-size:3em}p{color:#6B6357}</style>
    </head><body><h1>⚡</h1><h2>AZEM (عزم)</h2><p>لا يوجد اتصال بالإنترنت</p><p>سيعمل التطبيق عند إعادة الاتصال</p></body></html>`,
    { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
  );
}

// ---- Listen for messages from app ----
self.addEventListener('message', e => {
  if(e.data === 'SKIP_WAITING') self.skipWaiting();
  if(e.data === 'GET_VERSION') e.source.postMessage({ type: 'VERSION', version: VERSION });
  if (e.data?.type === 'SCHEDULE_REMINDER') {
    const { trainTime, name, day } = e.data;
    self._reminderConfig = { trainTime, name, day };
  }
});

// Periodic check for reminder (when SW is active)
self.addEventListener('periodicsync', e => {
  if (e.tag === 'azem-reminder') {
    e.waitUntil((async () => {
      const cfg = self._reminderConfig;
      if (!cfg) return;
      const now = new Date();
      const [h, m] = cfg.trainTime.split(':').map(Number);
      if (now.getHours() === h && now.getMinutes() === m) {
        await self.registration.showNotification('AZEM (عزم) 🔥', {
          body: `${cfg.name}، وقت تدريبك! اليوم ${cfg.day} في انتظارك 💪`,
          icon: './icon-192-1.png',
          tag: 'azem-reminder'
        });
      }
    })());
  }
});
