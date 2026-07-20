// Offline app-shell cache for the Capacitor Android app only (registered from
// src/main.jsx, gated to Capacitor.isNativePlatform() — never runs in a plain browser
// tab, so it can't collide with public/firebase-messaging-sw.js, which is registered at
// the same root scope for the web push path).
//
// The Android app normally loads live from madja.vercel.app (capacitor.config.json's
// server.url) instead of a bundled copy, so a web deploy applies to it automatically
// with no APK rebuild — but that also means with zero network at all, the WebView has
// nothing to render: there's no bundled index.html/JS to fall back to. This caches every
// same-origin GET response as it's requested (network-first, so an online device always
// gets the freshest build) and serves the cached copy once the network is unreachable —
// covering the app shell (index.html, the hashed JS/CSS bundle) and any other
// same-origin GET. It deliberately leaves cross-origin requests (Supabase, Firebase)
// alone — those have their own offline handling: useRemoteState mirrors every Supabase
// read/write into localStorage and falls back to that cache when a fetch fails, so
// already-opened data (kelas, jadwal, nilai, ...) stays visible offline too, not just
// the empty app shell.
const CACHE_NAME = "mdtnilai-offline-v1";

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      })
      .catch(() => caches.match(req).then((cached) => cached || caches.match("/index.html")))
  );
});
