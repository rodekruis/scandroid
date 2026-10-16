/* Scandroid PWA service worker — v40 */
const CACHE_VERSION = 'v40'; // bump on every deploy
const CACHE_NAME = `scandroid-cache-${CACHE_VERSION}`;

// NOTE: We deliberately do NOT precache /scan, /fsp-admin, /fsp-login here
// because they're auth-protected. During install the user is often
// unauthenticated, so precaching would store a redirect to /fsp-login
// instead of the real page. We let them be cached lazily on first
// successful (authenticated) navigation instead.
const PRECACHE_URLS = [
  "/", "/invalid-qr",
  "/beneficiary-offline", "/success-offline",
  "/static/scandroid.png", "/static/scandroid_banner.png",
  "/static/ns1.png", "/static/ns2.png",
  "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js"
];

console.log(`[ServiceWorker] Installing version: ${CACHE_NAME}`);

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      for (const url of PRECACHE_URLS) {
        try {
          await cache.add(url);
        } catch (err) {
          console.warn("[SW] Failed to cache:", url);
        }
      }
    })
  );
});

self.addEventListener("activate", (event) => {
  self.clients.claim();
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
});

// Routes that require auth on the server, but MUST work offline once cached.
// Strategy: network-first with very short timeout, fall back to cache.
// On success we update the cache, but ONLY if response is a real 200 HTML
// (never cache redirects — they would poison the offline experience).
const APP_SHELL_AUTH_ROUTES = ["/scan", "/fsp-admin", "/fsp-login"];

self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  if (req.method !== "GET") return;

  // -------------- /ping: network-only reachability probe -------------- //
  // The page reads response.ok to decide online/offline, so return the REAL
  // network response when it answers, and a 503 only when the fetch genuinely
  // fails (offline). Never cached.
  // (Previously this synthesised a 204 and overrode the request headers; that
  //  is removed in favour of passing the true status through, which is simpler
  //  and avoids the probe misreporting reachability.)
  if (url.pathname === "/ping") {
    event.respondWith(
      fetch(req, { cache: "no-store" })
        .catch(() => new Response("", { status: 503 }))
    );
    return;
  }

  // Special handling for Kobo sync ZIP (online-only, but cache last-good)
  if (url.pathname === "/api/offline/latest.zip") {
    event.respondWith(networkThenCache(req));
    return;
  }

  // -------- AUTH-PROTECTED APP SHELL ROUTES --------
  // These must work offline. Use network-first with short timeout
  // and only cache 200 HTML (not redirects).
  if (APP_SHELL_AUTH_ROUTES.includes(url.pathname)) {
    event.respondWith(networkFirstWithCacheUpdate(req, 1500));
    return;
  }

  // -------- OFFLINE-FIRST APP ROUTES --------
  const offlineFirstRoutes = [
    "/beneficiary-offline",
    "/success-offline",
    "/invalid-qr"
  ];

  if (offlineFirstRoutes.includes(url.pathname)) {
    event.respondWith(cacheFirst(req));
    return;
  }

  // -------- INSTANCE-STATIC (uploaded logos, etc.) --------
  // These files are mutable: they are overwritten in place on re-upload and
  // versioned with a ?v=<mtime> query string. The generic static handler below
  // uses cacheFirst with ignoreSearch, which would ignore ?v= and serve a stale
  // logo forever (only a hard refresh would show the new one). Use network-first
  // so a changed ?v= always fetches fresh, falling back to cache only offline.
  if (url.pathname.startsWith("/instance-static/")) {
    event.respondWith(networkThenCache(req));
    return;
  }

  // -------- NORMAL NAVIGATION --------
  const isNavigation =
    req.mode === "navigate" ||
    (req.headers.get("accept") || "").includes("text/html");

  if (isNavigation) {
    event.respondWith(networkFirstNavigation(req));
    return;
  }

  // Static assets
  if (isStatic(req)) {
    event.respondWith(cacheFirst(req));
    return;
  }
});


/* ---------- helpers ---------- */

function isStatic(req) {
  const url = new URL(req.url);
  return (
    (url.origin === location.origin &&
      (/\/static\//.test(url.pathname) ||
        /\.(js|css|png|jpg|jpeg|webp|svg|ico)$/i.test(url.pathname))) ||
    /cdn.jsdelivr.net$/.test(url.host)
  );
}

/**
 * Network-first with timeout for auth-protected app-shell routes.
 * - If network succeeds AND response is a real 200 (not a redirect),
 *   update the cache and return the fresh response.
 * - If network fails / times out, return whatever is in cache.
 * - Cache lookup ignores query strings so /scan?lang=fr still works.
 */
async function networkFirstWithCacheUpdate(req, timeoutMs) {
  const cache = await caches.open(CACHE_NAME);
  const url = new URL(req.url);

  const fetchPromise = (async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(req, { cache: "no-store", signal: controller.signal });
      clearTimeout(timer);
      // Only cache real 200 OK HTML — never redirects, never errors.
      if (res && res.ok && !res.redirected && res.status === 200) {
        try {
          // Cache under the bare path so query strings don't fragment cache.
          cache.put(url.pathname, res.clone());
        } catch (e) {
          console.warn("[SW] cache.put failed:", e);
        }
      }
      return res;
    } catch (e) {
      clearTimeout(timer);
      throw e;
    }
  })();

  try {
    return await fetchPromise;
  } catch (e) {
    // Offline / timeout — try cache (ignoreSearch so query strings are tolerated)
    let cached = await cache.match(req, { ignoreSearch: true });
    if (cached) return cached;
    cached = await cache.match(url.pathname);
    if (cached) return cached;
    return new Response(
      "<h1>Offline</h1><p>This page hasn't been cached yet. Please connect to the internet once and reload it before going offline.</p>",
      { status: 503, headers: { "Content-Type": "text/html" } }
    );
  }
}

async function networkFirstNavigation(req) {
  try {
    return await fetch(req, { cache: "no-store" });
  } catch (e) {
    const cache = await caches.open(CACHE_NAME);
    const url = new URL(req.url);

    let cached = await cache.match(req, { ignoreSearch: true });
    if (cached) return cached;

    let path = url.pathname;
    cached = await cache.match(path);
    if (cached) return cached;
    cached = await cache.match(path + '/');
    if (cached) return cached;
    cached = await cache.match(path + '.html');
    if (cached) return cached;

    const offline = await cache.match("/offline");
    return offline || new Response("Offline", { status: 503 });
  }
}

async function cacheFirst(req) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(req, { ignoreSearch: true });
  if (cached) return cached;

  try {
    const fresh = await fetch(req);
    if (fresh && fresh.ok && !fresh.redirected) {
      cache.put(req, fresh.clone());
    }
    return fresh;
  } catch (e) {
    return new Response(
      "<h1>Offline</h1><p>This page hasn't been cached yet.</p>",
      { status: 503, headers: { "Content-Type": "text/html" } }
    );
  }
}

async function networkThenCache(req) {
  try {
    const res = await fetch(req, { cache: "no-store" });
    const cache = await caches.open(CACHE_NAME);
    if (res && res.ok) cache.put(req, res.clone());
    return res;
  } catch (e) {
    const cache = await caches.open(CACHE_NAME);
    const cached = await cache.match(req, { ignoreSearch: true });
    if (cached) return cached;
    throw e;
  }
}