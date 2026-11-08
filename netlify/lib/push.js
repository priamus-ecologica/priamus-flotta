// ============================================================
//  PRIAMUS - Notifiche sul telefono degli autisti (Web Push)
//
//  File: netlify/lib/push.js
//  Lo usano:
//    - netlify/functions/api.js       (iscrizione dal telefono, prova)
//    - netlify/functions/notifiche.js (controllo ogni 2 minuti)
//
//  Come funziona:
//    1. l'autista tocca "Attiva notifiche" nell'app: il telefono
//       crea un "indirizzo" per le notifiche, che salviamo nel
//       foglio NOTIFICHE_ISCRIZIONI insieme al nome dell'autista;
//    2. ogni 2 minuti notifiche.js guarda REGISTRO_VIAGGI: per ogni
//       viaggio "Assegnato" (di oggi o dei giorni dopo) che non ha
//       ancora avvisato, manda la notifica all'autista e lo segna
//       in NOTIFICHE_INVIATE, cosi' non lo manda due volte.
//       Vale per i viaggi creati dall'app e da "Crea i viaggi" nel foglio.
//
//  Le chiavi VAPID (servono a firmare le notifiche) stanno nel foglio
//  NOTIFICHE_CHIAVI, NON nel codice: il repository e' pubblico.
//  Se su Netlify esistono VAPID_PUBLIC_KEY e VAPID_PRIVATE_KEY,
//  hanno la precedenza su quelle del foglio.
// ============================================================

const webpush = require("web-push");

const F_ISCR = "NOTIFICHE_ISCRIZIONI";
const F_INV = "NOTIFICHE_INVIATE";
const F_CHIAVI = "NOTIFICHE_CHIAVI";

const RIGHE_REGISTRO = 1500;   // righe recenti di REGISTRO_VIAGGI da guardare
const MAX_INVIATE = 6000;      // oltre questo numero il foglio INVIATE viene accorciato
const TIENI_INVIATE = 3000;
const CONTATTO = "https://priamus-flotta.netlify.app";

// ------------------------------------------------------------
//  piccoli aiuti sul foglio
// ------------------------------------------------------------
async function leggi(s, id, range) {
  try {
    const r = await s.spreadsheets.values.get({
      spreadsheetId: id, range,
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "SERIAL_NUMBER"
    });
    return r.data.values || [];
  } catch (e) {
    if (/Unable to parse range/i.test(String(e.message))) return null;   // foglio mancante
    throw e;
  }
}

async function creaFoglio(s, id, nome, intestazione) {
  await s.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addSheet: { properties: { title: nome, hidden: true } } }] }
  });
  await s.spreadsheets.values.update({
    spreadsheetId: id, range: nome + "!A1",
    valueInputOption: "RAW",
    requestBody: { values: [intestazione] }
  });
}

async function accoda(s, id, nome, righe) {
  if (!righe.length) return;
  await s.spreadsheets.values.append({
    spreadsheetId: id, range: nome + "!A1",
    valueInputOption: "RAW", insertDataOption: "INSERT_ROWS",
    requestBody: { values: righe }
  });
}

async function riscrivi(s, id, nome, intestazione, righe) {
  await s.spreadsheets.values.clear({ spreadsheetId: id, range: nome + "!A:Z" });
  await s.spreadsheets.values.update({
    spreadsheetId: id, range: nome + "!A1",
    valueInputOption: "RAW",
    requestBody: { values: [intestazione].concat(righe) }
  });
}

function oraItalia() {
  return new Intl.DateTimeFormat("it-IT", {
    timeZone: "Europe/Rome", day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: false
  }).format(new Date());
}

// numero seriale del foglio per "oggi" in Italia (senza ore)
function oggiSeriale() {
  const p = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit"
  }).format(new Date()).split("-");
  return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000 + 25569;
}

const GIORNI = ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"];
function dataBreve(n) {
  if (typeof n !== "number") return String(n || "");
  const d = new Date(Math.round((n - 25569) * 86400000));
  return GIORNI[d.getUTCDay()] + " " +
    String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0");
}

// ------------------------------------------------------------
//  chiavi VAPID
// ------------------------------------------------------------
let _chiavi = null;
async function chiavi(s, id) {
  if (_chiavi) return _chiavi;
  if (process.env.VAPID_PUBLIC_KEY && process.env.VAPID_PRIVATE_KEY) {
    _chiavi = { pubblica: process.env.VAPID_PUBLIC_KEY, privata: process.env.VAPID_PRIVATE_KEY };
  } else {
    const v = await leggi(s, id, F_CHIAVI + "!A2:B3");
    const m = {};
    (v || []).forEach(r => { m[String(r[0] || "").trim()] = String(r[1] || "").trim(); });
    if (!m.pubblica || !m.privata) throw new Error("Chiavi delle notifiche mancanti (foglio " + F_CHIAVI + ").");
    _chiavi = m;
  }
  webpush.setVapidDetails(CONTATTO, _chiavi.pubblica, _chiavi.privata);
  return _chiavi;
}

async function chiavePubblica(s, id) {
  return (await chiavi(s, id)).pubblica;
}

// ------------------------------------------------------------
//  iscrizioni: una riga per telefono (endpoint), con l'autista
// ------------------------------------------------------------
const INT_ISCR = ["Autista", "Endpoint", "Iscrizione", "Telefono", "Aggiornata"];

async function leggiIscrizioni(s, id) {
  let v = await leggi(s, id, F_ISCR + "!A2:E");
  if (v === null) { await creaFoglio(s, id, F_ISCR, INT_ISCR); v = []; }
  return v.filter(r => r[0] && r[1] && r[2]).map(r => ({
    autista: String(r[0]).trim(), endpoint: String(r[1]), json: String(r[2]),
    telefono: String(r[3] || ""), data: String(r[4] || "")
  }));
}

function telefonoDa(ua) {
  ua = String(ua || "");
  if (/iPhone|iPad/i.test(ua)) return "iPhone";
  const m = ua.match(/Android[^;)]*;\s*([^;)]+?)(?:\sBuild|\))/i);
  if (m) return ("Android " + m[1]).slice(0, 60);
  if (/Android/i.test(ua)) return "Android";
  return ua.slice(0, 60);
}

// Salva (o aggiorna) l'iscrizione di un telefono.
// Se sullo stesso telefono entra un altro autista, la riga passa a lui.
async function iscrivi(s, id, autista, iscrizione, ua) {
  if (!iscrizione || !iscrizione.endpoint || !iscrizione.keys) throw new Error("Iscrizione non valida.");
  const tutte = await leggiIscrizioni(s, id);
  const altre = tutte.filter(x => x.endpoint !== iscrizione.endpoint);
  const riga = [autista, iscrizione.endpoint, JSON.stringify(iscrizione), telefonoDa(ua), oraItalia()];
  if (altre.length === tutte.length) {
    await accoda(s, id, F_ISCR, [riga]);
  } else {
    await riscrivi(s, id, F_ISCR, INT_ISCR,
      altre.map(x => [x.autista, x.endpoint, x.json, x.telefono, x.data]).concat([riga]));
  }
  return { ok: true };
}

// Manda un messaggio a tutti i telefoni di un autista.
// Restituisce quanti telefoni l'hanno ricevuto e gli endpoint scaduti.
async function mandaAutista(s, id, iscrizioni, autista, messaggio) {
  await chiavi(s, id);
  const sue = iscrizioni.filter(x => x.autista.toLowerCase() === autista.toLowerCase());
  let ok = 0;
  const scadute = [];
  for (const x of sue) {
    try {
      await webpush.sendNotification(JSON.parse(x.json), JSON.stringify(messaggio), { TTL: 86400 });
      ok++;
    } catch (e) {
      // 404/410: il telefono ha tolto il permesso o ha disinstallato l'app
      if (e.statusCode === 404 || e.statusCode === 410) scadute.push(x.endpoint);
      else console.log("Notifica non partita per " + autista + ": " + (e.statusCode || "") + " " + e.message);
    }
  }
  return { telefoni: sue.length, ok, scadute };
}

async function togliScadute(s, id, iscrizioni, scadute) {
  if (!scadute.length) return;
  const restano = iscrizioni.filter(x => scadute.indexOf(x.endpoint) < 0);
  await riscrivi(s, id, F_ISCR, INT_ISCR, restano.map(x => [x.autista, x.endpoint, x.json, x.telefono, x.data]));
}

// Notifica di prova (dal pannello del responsabile)
async function prova(s, id, autista) {
  const iscr = await leggiIscrizioni(s, id);
  const r = await mandaAutista(s, id, iscr, autista, {
    title: "🔔 Prova notifiche Priamus",
    body: "Se leggi questo messaggio, le notifiche dei viaggi funzionano.",
    tag: "prova"
  });
  await togliScadute(s, id, iscr, r.scadute);
  return r;
}

// ------------------------------------------------------------
//  il controllo periodico: viaggi assegnati non ancora avvisati
// ------------------------------------------------------------
const INT_INV = ["Chiave", "Autista", "Data_viaggio", "Inviata", "Telefoni_raggiunti"];

async function inviaPendenti(s, id) {
  // REGISTRO_VIAGGI: A ID, B Data, C Autista, E Targa, P Stato, S Tratta
  const reg = (await leggi(s, id, "REGISTRO_VIAGGI!A2:S10001")) || [];
  const recenti = reg.slice(Math.max(0, reg.length - RIGHE_REGISTRO));
  const oggi = oggiSeriale();

  const candidati = [];
  recenti.forEach(r => {
    const vid = String(r[0] || "").trim();
    const autista = String(r[2] || "").trim();
    const stato = String(r[15] || "").trim();
    const data = r[1];
    if (!vid || !autista || stato !== "Assegnato") return;
    if (typeof data !== "number" || Math.floor(data) < oggi) return;
    candidati.push({ chiave: vid + "|" + autista, vid, autista, data, tratta: String(r[18] || r[9] || "").trim() });
  });

  // primo avvio: segno come gia' avvisati i viaggi che esistono adesso
  let inv = await leggi(s, id, F_INV + "!A2:A");
  if (inv === null) {
    await creaFoglio(s, id, F_INV, INT_INV);
    await accoda(s, id, F_INV, candidati.map(c => [c.chiave, c.autista, dataBreve(c.data), "primo avvio", 0]));
    return { primoAvvio: true, segnati: candidati.length };
  }
  const gia = new Set(inv.map(r => String(r[0] || "")));
  const nuovi = candidati.filter(c => !gia.has(c.chiave));
  if (!nuovi.length) return { nuovi: 0 };

  // raggruppo per autista: un solo messaggio anche se ha piu' viaggi
  const perAutista = {};
  nuovi.forEach(c => { (perAutista[c.autista] = perAutista[c.autista] || []).push(c); });

  const iscr = await leggiIscrizioni(s, id);
  let scadute = [];
  const righe = [];
  const quando = oraItalia();

  for (const autista of Object.keys(perAutista)) {
    const v = perAutista[autista].sort((a, b) => a.data - b.data);
    let titolo, testo;
    if (v.length === 1) {
      titolo = "🚛 Nuovo viaggio assegnato";
      testo = dataBreve(v[0].data) + ": " + (v[0].tratta || "apri l'app per i dettagli");
    } else {
      titolo = "🚛 " + v.length + " nuovi viaggi assegnati";
      testo = v.map(x => dataBreve(x.data) + ": " + (x.tratta || "?")).join("\n");
    }
    let r = { telefoni: 0, ok: 0, scadute: [] };
    if (iscr.some(x => x.autista.toLowerCase() === autista.toLowerCase())) {
      r = await mandaAutista(s, id, iscr, autista, { title: titolo, body: testo, tag: "viaggi", url: "./" });
      scadute = scadute.concat(r.scadute);
    }
    v.forEach(x => righe.push([x.chiave, autista, dataBreve(x.data), quando, r.ok]));
  }

  await accoda(s, id, F_INV, righe);
  await togliScadute(s, id, iscr, scadute);

  // il foglio degli avvisi non deve crescere all'infinito
  if (gia.size + righe.length > MAX_INVIATE) {
    const tutte = (await leggi(s, id, F_INV + "!A2:E")) || [];
    await riscrivi(s, id, F_INV, INT_INV, tutte.slice(-TIENI_INVIATE));
  }

  return { nuovi: nuovi.length, autisti: Object.keys(perAutista).length, scadute: scadute.length };
}

module.exports = { chiavePubblica, iscrivi, prova, inviaPendenti };
