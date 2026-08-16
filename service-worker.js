// Ento app service worker
// Strategy:
//  - Precache the app shell (everything needed to open the form with zero signal).
//  - Cache-first for all GET requests (including CDN libs like Chart.js/Leaflet/fonts)
//    so the first successful load banks a copy for later offline use.
//  - NEVER touch requests to script.google.com (Google Sheet sync) — those must hit
//    the real network so the app's own offline-queue/retry logic (in index.html)
//    keeps working exactly as it does today.
// Bump CACHE_NAME any time you redeploy so devices pick up the new files.
const CACHE_NAME = "ento-shell-v1";

const APP_SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./support.js",
  "./locations-core.json",
  "./locations-villages.json",
  "./icon-180.png",
  "./icon-512.png",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/styles.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/_ds_bundle.js",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/typography.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/colors.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/spacing.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/fonts.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/fonts/Trebuchet_MS.ttf",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/fonts/Trebuchet_MS_Bold.ttf",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/fonts/Trebuchet_MS_Italic.ttf",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/fonts/Trebuchet_MS_Bold_Italic.ttf",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // addAll fails the whole install if even one file 404s — go one by one instead
      // so a single missing/renamed asset can't block the app from becoming installable.
      Promise.all(
        APP_SHELL.map((url) =>
          cache.add(url).catch((err) => console.warn("[sw] precache skipped:", url, err))
        )
      )
    ).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = req.url;

  // Only handle GETs; POSTs (sheet sync) always go straight to network untouched.
  if (req.method !== "GET") return;

  // Google Apps Script sync/history/records calls: network only, never cached,
  // never intercepted. Offline failures here are handled by index.html itself.
  if (url.indexOf("script.google.com") !== -1) return;

  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((res) => {
          // Only cache good, same-type responses (covers same-origin + CORS'd CDN assets).
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          }
          return res;
        })
        .catch(() => cached); // offline and not cached yet -> nothing we can do

      // Cache-first: instant + works offline. Falls back to network if not cached yet.
      return cached || network;
    })
  );
});
