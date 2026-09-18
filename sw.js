// Tiene l'app disponibile anche senza rete.
// Strategia: prima la rete, saltando la memoria del browser (così un aggiornamento arriva subito),
// e la copia salvata solo se la rete manca.
const CACHE = "diario-assistenza-v2";
const FILES = ["./", "./index.html", "./app.js", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./icon-maskable-512.png"];

self.addEventListener("install", (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).catch(() => {}));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    // cache: "reload" ignora la copia tenuta dal browser e chiede il file al server
    fetch(req, { cache: "reload" })
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
        return res;
      })
      .catch(() => fetch(req).catch(() => caches.match(req).then((hit) => hit || caches.match("./index.html"))))
  );
});
