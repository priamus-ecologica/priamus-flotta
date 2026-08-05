// ============================================================
//  Service worker Priamus - versione con AGGIORNAMENTO AUTOMATICO
//
//  Il precedente serviva sempre la copia salvata: quando
//  aggiornavamo l'app, i telefoni restavano indietro finche'
//  non la si disinstallava e reinstallava.
//
//  Questo invece: prova sempre a scaricare la versione nuova,
//  e usa la copia salvata SOLO se manca la connessione.
// ============================================================

const CACHE = "priamus-v2";
const FILES = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", e => {
  // skipWaiting: la versione nuova entra in servizio subito,
  // senza aspettare che l'utente chiuda tutte le schede
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(k => Promise.all(k.filter(n => n !== CACHE).map(n => caches.delete(n))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const url = e.request.url;

  // Le chiamate al backend non si mettono MAI in cache
  if (url.indexOf("/.netlify/functions/") !== -1) return;
  if (url.indexOf("script.google.com") !== -1) return;
  if (e.request.method !== "GET") return;

  // RETE PER PRIMA: cosi' un aggiornamento arriva subito.
  // La copia salvata interviene solo se la rete non risponde.
  e.respondWith(
    fetch(e.request)
      .then(resp => {
        const copia = resp.clone();
        caches.open(CACHE).then(c => c.put(e.request, copia)).catch(() => {});
        return resp;
      })
      .catch(() =>
        caches.match(e.request).then(r => r || caches.match("./index.html"))
      )
  );
});

// Permette all'app di chiedere l'aggiornamento immediato
self.addEventListener("message", e => {
  if (e.data === "AGGIORNA") self.skipWaiting();
});
