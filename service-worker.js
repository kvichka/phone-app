/**
 * Ento field app — service worker
 *
 * Cache-first for everything the app is made of, so it opens with no signal.
 * Anything that talks to the Google Sheet is network-only and never cached:
 * the app's own offline queue is what makes sync work, not this worker.
 *
 * BUMP CACHE_NAME ON EVERY PUSH. Skipping it leaves phones on stale files.
 */

const CACHE_NAME = "ento-shell-v11";

const SHELL = [
  "./",
  "./index.html",
  "./support.js",
  "./manifest.json",
  "./locations-core.json",
  "./locations-villages.json",
  "./icon-180.png",
  "./icon-512.png",
  "./dashboard.html",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/styles.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/_ds_bundle.js",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/colors.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/fonts.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/spacing.css",
  "./_ds/chai-design-system-eb475a38-cb32-41b6-a2ce-822ca75afe8d/tokens/typography.css",
];

// Fetched opportunistically: a miss here must never fail the install.
const VENDOR = [
  "https://cdn.jsdelivr.net/npm/chart.js@4.5.1",
  "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css",
  "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js",
  "https://fonts.googleapis.com/css2?family=Kantumruy+Pro:wght@400;500;700&display=swap",
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(SHELL);
    await Promise.all(VENDOR.map((url) =>
      cache.add(new Request(url, { mode: "no-cors" })).catch(() => null)));
    self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // The sheet: always live, never cached, never intercepted on failure.
  if (url.hostname.endsWith("script.google.com")
    || url.hostname.endsWith("script.googleusercontent.com")) {
    return;
  }

  // Map tiles: try the network, fall back to whatever was seen before.
  if (url.hostname.endsWith("tile.openstreetmap.org")) {
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((c) => c.put(req, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(req))
    );
    return;
  }

  event.respondWith((async () => {
    const hit = await caches.match(req, { ignoreSearch: false });
    if (hit) {
      // Refresh in the background so the next open is current.
      fetch(req).then((res) => {
        if (res && (res.ok || res.type === "opaque")) {
          caches.open(CACHE_NAME).then((c) => c.put(req, res.clone())).catch(() => {});
        }
      }).catch(() => {});
      return hit;
    }
    try {
      const res = await fetch(req);
      if (res && (res.ok || res.type === "opaque")) {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((c) => c.put(req, copy)).catch(() => {});
      }
      return res;
    } catch (err) {
      if (req.mode === "navigate") {
        const shell = await caches.match("./index.html");
        if (shell) return shell;
      }
      throw err;
    }
  })());
});
