const CACHE_NAME = "merlin-shell-v4";
const APP_BASE = new URL("./", self.registration.scope).pathname;
const appAsset = (name = "") => `${APP_BASE}${name}`;
const APP_SHELL = [
  APP_BASE,
  appAsset("manifest.webmanifest"),
  appAsset("icon-192.png"),
  appAsset("icon-512.png"),
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter((key) => key !== CACHE_NAME)
            .map((key) => caches.delete(key)),
        ),
      ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);
  if (
    url.origin !== self.location.origin ||
    url.pathname.startsWith(appAsset("api/"))
  ) {
    return;
  }

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request).catch(() => caches.match(APP_BASE)),
    );
    return;
  }

  const cacheable =
    APP_SHELL.includes(url.pathname) ||
    url.pathname.startsWith(appAsset("assets/")) ||
    url.pathname.startsWith(appAsset("_next/"));
  if (!cacheable) return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request)),
  );
});
