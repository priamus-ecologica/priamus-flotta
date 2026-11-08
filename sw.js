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

// ---------- notifiche dei viaggi assegnati ----------
self.addEventListener("push", e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; }
  catch (x) { d = { body: e.data ? e.data.text() : "" }; }
  e.waitUntil(
    self.registration.showNotification(d.title || "Priamus", {
      body: d.body || "",
      icon: "icon-192.png",
      badge: "icon-192.png",
      tag: d.tag || "viaggi",
      renotify: true,
      vibrate: [200, 100, 200],
      data: { url: d.url || "./" }
    })
  );
});

// tocco sulla notifica: apre l'app (o la porta in primo piano)
self.addEventListener("notificationclick", e => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || "./";
  e.waitUntil(
    clients.matchAll({ type: "window", includeUncontrolled: true }).then(lista => {
      for (const c of lista) { if ("focus" in c) return c.focus(); }
      return clients.openWindow(url);
    })
  );
});
