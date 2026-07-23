const CACHE_NAME = "merlin-shell-v6";
const SHARE_CACHE_NAME = "merlin-share-inbox-v1";
const MAX_SHARED_IMAGES = 3;
const MAX_SHARED_IMAGE_BYTES = 8 * 1024 * 1024;
const SUPPORTED_SHARED_IMAGE_TYPES = new Set([
  "image/jpeg",
  "image/png",
  "image/webp",
]);
const APP_BASE = new URL("./", self.registration.scope).pathname;
const appAsset = (name = "") => `${APP_BASE}${name}`;
const SHARE_TARGET_PATH = appAsset("share-target");
const APP_SHELL = [
  APP_BASE,
  appAsset("manifest.webmanifest"),
  appAsset("icon-192.png"),
  appAsset("icon-512.png"),
];

function shareRedirect(parameters) {
  const url = new URL(APP_BASE, self.location.origin);
  for (const [name, value] of Object.entries(parameters)) {
    url.searchParams.set(name, String(value));
  }
  return Response.redirect(url.href, 303);
}

function safeSharedFileName(value, index) {
  const lastSegment = String(value || "")
    .split(/[\\/]/)
    .pop()
    .trim()
    .slice(0, 120);
  return lastSegment || `megosztott-kep-${index + 1}`;
}

function isSupportedSharedImage(value) {
  return (
    typeof File !== "undefined" &&
    value instanceof File &&
    SUPPORTED_SHARED_IMAGE_TYPES.has(value.type.toLowerCase()) &&
    value.size > 0 &&
    value.size <= MAX_SHARED_IMAGE_BYTES
  );
}

async function clearShareInbox(cache) {
  const keys = await cache.keys();
  await Promise.all(keys.map((request) => cache.delete(request)));
}

async function handleShareTarget(request) {
  const cache = await caches.open(SHARE_CACHE_NAME);

  try {
    const formData = await request.formData();
    const candidates = formData.getAll("images");
    const images = candidates
      .filter(isSupportedSharedImage)
      .slice(0, MAX_SHARED_IMAGES);

    await clearShareInbox(cache);
    if (images.length === 0) {
      return shareRedirect({ shareError: "unsupported-image" });
    }

    const shareId =
      globalThis.crypto?.randomUUID?.() ||
      `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    const entryBase = new URL(
      appAsset(`__share-inbox/${shareId}/`),
      self.location.origin,
    );
    const sharedText = [
      formData.get("text"),
      formData.get("url"),
    ]
      .filter((value) => typeof value === "string" && value.trim())
      .join("\n")
      .trim()
      .slice(0, 4000);
    const metadata = {
      createdAt: Date.now(),
      text: sharedText,
      rejectedCount: Math.max(0, candidates.length - images.length),
      files: images.map((file, index) => ({
        name: safeSharedFileName(file.name, index),
        type: file.type.toLowerCase(),
        size: file.size,
      })),
    };

    await Promise.all(
      images.map((file, index) =>
        cache.put(
          new Request(new URL(`image-${index}`, entryBase)),
          new Response(file, {
            headers: {
              "Content-Type": file.type,
              "Cache-Control": "no-store",
            },
          }),
        ),
      ),
    );
    await cache.put(
      new Request(new URL("meta.json", entryBase)),
      new Response(JSON.stringify(metadata), {
        headers: {
          "Content-Type": "application/json; charset=utf-8",
          "Cache-Control": "no-store",
        },
      }),
    );

    return shareRedirect({ shareId });
  } catch {
    await clearShareInbox(cache);
    return shareRedirect({ shareError: "read-failed" });
  }
}

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
  const url = new URL(event.request.url);
  if (
    event.request.method === "POST" &&
    url.origin === self.location.origin &&
    url.pathname === SHARE_TARGET_PATH
  ) {
    event.respondWith(handleShareTarget(event.request));
    return;
  }

  if (event.request.method !== "GET") return;

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
