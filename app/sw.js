const CACHE_NAME = "chuan-hoa-cong-to-v4";
const ASSETS = ["/", "/index.html", "/styles.css", "/app.js", "/manifest.webmanifest", "/icon.svg"];
const INCOMING_DB_NAME = "chuan-hoa-incoming-files";
const INCOMING_DB_STORE = "files";
const INCOMING_FILE_KEY = "latest-excel";

function openIncomingDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(INCOMING_DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(INCOMING_DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveIncomingFile(file) {
  const db = await openIncomingDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(INCOMING_DB_STORE, "readwrite");
    tx.objectStore(INCOMING_DB_STORE).put(
      {
        name: file.name || "file_chia_se.xlsx",
        type: file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        blob: file,
        updatedAt: new Date().toISOString(),
      },
      INCOMING_FILE_KEY,
    );
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))),
    ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method === "POST" && url.pathname === "/share-target") {
    event.respondWith(
      (async () => {
        const formData = await event.request.formData();
        const file = formData.getAll("excel").find((item) => item instanceof File && item.size > 0);
        if (file) await saveIncomingFile(file);
        return Response.redirect("/?shared=1", 303);
      })(),
    );
    return;
  }
  if (event.request.method !== "GET") return;
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});
