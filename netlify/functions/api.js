// ============================================================
//  PRIAMUS - Backend su Netlify Functions
//  Parla con Google Sheets tramite l'API ufficiale:
//  niente Apps Script, niente avvii a freddo.
//
//  File: netlify/functions/api.js
//
//  Variabili d'ambiente da impostare su Netlify (Site config):
//    GOOGLE_CLIENT_EMAIL  -> email del robot (account di servizio)
//    GOOGLE_PRIVATE_KEY   -> chiave privata dal file JSON scaricato
//    SHEET_ID             -> id del foglio database
//    API_SEGRETO          -> stessa frase segreta scritta in index.html
//    PASSWORD_ADMIN       -> password del responsabile
// ============================================================

const { google } = require("googleapis");

const SHEET_ID = process.env.SHEET_ID;
const SEGRETO = process.env.API_SEGRETO;
const PASSWORD_ADMIN = process.env.PASSWORD_ADMIN;

// Quante righe recenti del registro considerare per i viaggi attivi
const FINESTRA = 1500;     // righe recenti usate per i viaggi attivi
const FINESTRA_MAX = 5000; // tetto di righe lette dal registro

// ---- connessione a Google (riusata tra le richieste finche' la funzione resta calda)
let _sheets = null;
function sheets() {
  if (_sheets) return _sheets;
  const auth = new google.auth.JWT(
    process.env.GOOGLE_CLIENT_EMAIL,
    null,
    (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    ["https://www.googleapis.com/auth/spreadsheets"]
  );
  _sheets = google.sheets({ version: "v4", auth });
  return _sheets;
}

// ---- cache in memoria per le anagrafiche (90 secondi)
let _base = null;
let _baseTime = 0;

async function leggi(range) {
  const res = await sheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: range,
    valueRenderOption: "UNFORMATTED_VALUE",
    dateTimeRenderOption: "SERIAL_NUMBER"
  });
  return res.data.values || [];
}

async function leggiTesto(range) {
  const res = await sheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: range
  });
  return res.data.values || [];
}

function serialToData(n) {
  // I fogli Google contano i giorni dal 30/12/1899
  if (typeof n !== "number") return String(n || "");
  const ms = Math.round((n - 25569) * 86400 * 1000);
  const d = new Date(ms);
  const gg = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return gg + "/" + mm + "/" + d.getUTCFullYear();
}

function serialToISO(n) {
  if (typeof n !== "number") return "";
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  return d.getUTCFullYear() + "-" +
         String(d.getUTCMonth() + 1).padStart(2, "0") + "-" +
         String(d.getUTCDate()).padStart(2, "0");
}

// "2026-08" a partire dal numero seriale del foglio
function serialToMese(n) {
  if (typeof n !== "number") return "";
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0");
}

function oraLocale() {
  // Ora italiana in formato HH:mm
  return new Intl.DateTimeFormat("it-IT", {
    timeZone: "Europe/Rome", hour: "2-digit", minute: "2-digit", hour12: false
  }).format(new Date());
}

function dataSerial(iso) {
  // "2026-08-04" -> numero seriale del foglio, a mezzogiorno
  const p = String(iso || "").split("-");
  let d;
  if (p.length === 3) d = Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
  else d = Date.now();
  return d / 86400000 + 25569;
}

async function base() {
  const ora = Date.now();
  if (_base && ora - _baseTime < 90000) return _base;

  // UNA sola richiesta per tutte e quattro le anagrafiche.
  // Prima erano 4 richieste separate: con piu' autisti insieme
  // si superava il limite di Google e arrivavano gli errori.
  const risp = await sheets().spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID,
    ranges: [
      "AUTISTI!A2:C200",
      "ANAGRAFICA_MEZZI!A2:E200",
      "ANAGRAFICA_RIMORCHI!A2:B200",
      "REGOLE_VIAGGI!A2:C300"
    ]
  });
  const blocchi = (risp.data.valueRanges || []).map(v => v.values || []);
  const vA = blocchi[0] || [], vM = blocchi[1] || [];
  const vR = blocchi[2] || [], vReg = blocchi[3] || [];

  const autisti = [], pin = {};
  vA.forEach(r => {
    const nome = String(r[0] || "").trim();
    if (nome) { autisti.push(nome); pin[nome] = String(r[2] || "").trim(); }
  });

  const mezzi = {};
  const rigaMezzo = {};   // targa -> riga nel foglio ANAGRAFICA_MEZZI
  vM.forEach((r, i) => {
    const targa = String(r[1] || "").trim();
    if (targa) {
      mezzi[targa] = String(r[3] || "").trim() || targa;
      rigaMezzo[targa] = i + 2;          // +2 perche' partiamo dalla riga 2
    }
  });

  const rimorchi = {};
  vR.forEach(r => {
    const targa = String(r[0] || "").trim();
    if (targa) rimorchi[targa] = String(r[1] || "").trim() || targa;
  });

  const regole = [], chiavePerArrivo = {};
  vReg.forEach(r => {
    const chiave = String(r[0] || "").trim();
    const arrivo = String(r[2] || "").trim();
    if (chiave) {
      regole.push({ chiave, partenza: String(r[1] || "").trim(), arrivo });
      if (arrivo && !chiavePerArrivo[arrivo]) chiavePerArrivo[arrivo] = chiave;
    }
  });

  _base = { autisti, pin, mezzi, rimorchi, regole, chiavePerArrivo, rigaMezzo };
  _baseTime = ora;
  return _base;
}

function svuotaCache() { _base = null; _baseTime = 0; }

// Cache brevissima del registro: piu' azioni della stessa schermata
// non rileggono lo stesso dato due volte.
let _coda = null;
let _codaTime = 0;

async function coda(fresca) {
  const ora = Date.now();
  if (!fresca && _coda && ora - _codaTime < 8000) return _coda;

  // UNA sola richiesta: prima ne servivano due.
  const tutte = await leggi("REGISTRO_VIAGGI!A2:V" + (FINESTRA_MAX + 1));
  const n = tutte.length;
  const da = Math.max(0, n - FINESTRA);
  const righe = tutte.slice(da);

  _coda = { righe: righe, primaRiga: 2 + da };
  _codaTime = ora;
  return _coda;
}

function svuotaCoda() { _coda = null; _codaTime = 0; }

function etich(mappa, targa) {
  const t = String(targa || "").trim();
  return t ? (mappa[t] || t) : "";
}

// ------------------------------------------------------------
//  Azioni
// ------------------------------------------------------------
async function azElenchi() {
  const b = await base();
  const c = await coda();

  const ultimeDestinazioni = {};
  for (let i = c.righe.length - 1; i >= 0; i--) {
    const targa = String(c.righe[i][4] || "").trim();
    const arrivo = String(c.righe[i][9] || "").trim();
    if (targa && arrivo && !ultimeDestinazioni[targa]) ultimeDestinazioni[targa] = arrivo;
  }

  return {
    autisti: b.autisti,
    mezzi: Object.keys(b.mezzi).map(t => ({ targa: t, label: b.mezzi[t] })),
    rimorchi: Object.keys(b.rimorchi).map(t => ({ targa: t, label: b.rimorchi[t] })),
    regole: b.regole,
    ultimeDestinazioni
  };
}

async function azViaggiAutista(req) {
  const b = await base();
  const c = await coda();
  const cercato = String(req.autista || "").trim().toLowerCase();

  // PRIMO PASSAGGIO: per ogni camion trovo l'ultimo KM di arrivo registrato.
  // Va fatto su TUTTE le righe prima di costruire l'elenco, altrimenti il
  // viaggio da avviare (che e' fra i piu' recenti) non trova ancora il dato.
  const ultimiKm = {};
  for (let i = c.righe.length - 1; i >= 0; i--) {
    const r = c.righe[i];
    const targa = String(r[4] || "").trim();
    const km = r[6] !== undefined && r[6] !== "" ? String(r[6]) : "";
    if (targa && km && !ultimiKm[targa]) ultimiKm[targa] = km;
  }

  // SECONDO PASSAGGIO: i viaggi dell'autista
  const viaggi = [];
  for (let i = c.righe.length - 1; i >= 0; i--) {
    const r = c.righe[i];
    const targa = String(r[4] || "").trim();

    if (String(r[2] || "").trim().toLowerCase() === cercato) {
      const stato = String(r[15] || "").trim();
      if (stato !== "Completato") {
        const arrivo = String(r[9] || "").trim();
        viaggi.push({
          id: String(r[0]),
          data: serialToData(r[1]),
          mezzo: etich(b.mezzi, targa),
          rimorchio: etich(b.rimorchi, String(r[10] || "").trim()),
          targaMezzo: targa,
          chiave: String(r[18] || "") || b.chiavePerArrivo[arrivo] || arrivo,
          stato: stato || "Assegnato",
          kmPrecedente: ultimiKm[targa] || "",
          nota: String(r[17] || "")
        });
      }
    }
  }
  return viaggi.reverse();
}

async function azStatoFlotta() {
  const b = await base();
  const c = await coda();
  const flotta = [];
  for (let i = c.righe.length - 1; i >= 0; i--) {
    const r = c.righe[i];
    const stato = String(r[15] || "").trim();
    if (stato === "Assegnato" || stato === "In Corso") {
      flotta.push({
        id: String(r[0]),
        autista: String(r[2] || ""),
        mezzo: etich(b.mezzi, r[4]),
        rimorchio: etich(b.rimorchi, r[10]),
        partenza: String(r[8] || ""),
        arrivo: String(r[9] || ""),
        chiave: String(r[18] || ""),
        nota: String(r[17] || ""),
        stato,
        data: serialToData(r[1]),
        // serve al modulo di modifica dell'admin
        targaMezzo: String(r[4] || "").trim(),
        targaRimorchio: String(r[10] || "").trim(),
        dataISO: serialToISO(r[1])
      });
    }
  }
  return flotta;
}

async function azLogin(req) {
  if (req.tipo === "responsabile") {
    if (req.password === PASSWORD_ADMIN) return { success: true, role: "admin" };
    return { success: false, msg: "Password Responsabile errata!" };
  }
  if (req.tipo === "autista") {
    const b = await base();
    const nome = String(req.utente || "").trim();
    if (!(nome in b.pin)) return { success: false, msg: "Autista non trovato nel database." };
    if (b.pin[nome] && b.pin[nome] === String(req.password)) {
      return { success: true, role: "autista", nome, viaggi: await azViaggiAutista({ autista: nome }) };
    }
    return { success: false, msg: "PIN errato!" };
  }
  return { success: false, msg: "Errore di sistema." };
}

async function trovaRiga(id) {
  const c = await coda();
  for (let i = c.righe.length - 1; i >= 0; i--) {
    if (String(c.righe[i][0]) === String(id)) {
      return { riga: c.primaRiga + i, dati: c.righe[i] };
    }
  }
  throw new Error("Viaggio non trovato.");
}

async function scrivi(range, valori) {
  await sheets().spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: valori }
  });
}

// Scrive PIU' celle in UNA sola chiamata.
// Prima le scritture partivano in parallelo e Google ne rifiutava qualcuna:
// era il motivo per cui l'Ora_Inizio a volte non veniva salvata.
async function scriviBlocco(elenco) {
  await sheets().spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: "USER_ENTERED",
      data: elenco.map(x => ({ range: x.range, values: x.values }))
    }
  });
}

async function azAssegna(req) {
  // anti-doppione
  if (req.idViaggio) {
    const c = await coda();
    for (const r of c.righe) {
      if (String(r[0]) === String(req.idViaggio)) {
        return { msg: "Viaggio gia' assegnato a " + req.autista };
      }
    }
  }
  const id = String(req.idViaggio || ("v-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8)));

  await sheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "REGISTRO_VIAGGI!A1:T1",
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        id, dataSerial(req.dataViaggio), req.autista, "",
        req.targaMezzo, "", "", "", req.partenza, req.arrivo,
        req.targaRimorchio, "", "", "", "", "Assegnato", "",
        req.nota || "", req.chiave || "", ""
      ]]
    }
  });
  svuotaCoda();
  return { msg: "Viaggio assegnato a " + req.autista };
}

async function azAvvia(req) {
  const t = await trovaRiga(req.id);
  const ora = oraLocale();
  await scriviBlocco([
    { range: "REGISTRO_VIAGGI!F" + t.riga, values: [[Number(req.kmPartenza)]] },  // KM partenza
    { range: "REGISTRO_VIAGGI!P" + t.riga, values: [["In Corso"]] },              // Stato
    { range: "REGISTRO_VIAGGI!T" + t.riga, values: [[ora]] }                      // Ora_Inizio
  ]);
  svuotaCoda();
  return { msg: "Viaggio avviato!", oraInizio: ora };
}

async function azTermina(req) {
  // I KM di arrivo sono OBBLIGATORI: senza, il viaggio non si chiude.
  if (req.kmArrivo === "" || req.kmArrivo == null || isNaN(Number(req.kmArrivo))) {
    throw new Error("Inserisci i KM di arrivo per chiudere il viaggio.");
  }

  const t = await trovaRiga(req.id);
  const kmPart = Number(t.dati[5]) || 0;
  const kmArr = Number(req.kmArrivo);
  const kmViaggio = (kmPart !== 0) ? (kmArr - kmPart) : "";

  const scritture = [
    { range: "REGISTRO_VIAGGI!G" + t.riga + ":H" + t.riga, values: [[kmArr, kmViaggio]] },
    { range: "REGISTRO_VIAGGI!L" + t.riga + ":M" + t.riga, values: [[
        (req.tonnellate !== "" && req.tonnellate != null) ? Number(req.tonnellate) : "",
        req.oreLavoro || ""
      ]] },
    { range: "REGISTRO_VIAGGI!P" + t.riga + ":Q" + t.riga, values: [["Completato", oraLocale()]] },
    // NUOVO: U = numero di viaggi svolti, V = nota scritta dall'autista
    { range: "REGISTRO_VIAGGI!U" + t.riga + ":V" + t.riga, values: [[
        (req.numeroViaggi !== "" && req.numeroViaggi != null) ? Number(req.numeroViaggi) : "",
        req.notaAutista || ""
      ]] }
  ];

  // NUOVO: aggiorno i KM attuali del camion in ANAGRAFICA_MEZZI (colonna E)
  if (kmArr !== "" && !isNaN(kmArr)) {
    const b = await base();
    const targa = String(t.dati[4] || "").trim();
    const rigaAnagrafica = b.rigaMezzo ? b.rigaMezzo[targa] : null;
    if (rigaAnagrafica) {
      scritture.push({
        range: "ANAGRAFICA_MEZZI!E" + rigaAnagrafica,
        values: [[kmArr]]
      });
    }
  }

  await scriviBlocco(scritture);
  svuotaCoda();
  return { msg: "Viaggio chiuso correttamente!" };
}

async function azElimina(req) {
  const t = await trovaRiga(req.id);
  if (String(t.dati[15] || "").trim() === "Completato") {
    throw new Error("Viaggio Completato: non si elimina dallo storico.");
  }
  const autista = String(t.dati[2] || "");

  // per cancellare la riga serve l'id numerico della scheda
  const meta = await sheets().spreadsheets.get({ spreadsheetId: SHEET_ID });
  const scheda = meta.data.sheets.find(s => s.properties.title === "REGISTRO_VIAGGI");
  await sheets().spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId: scheda.properties.sheetId,
            dimension: "ROWS",
            startIndex: t.riga - 1,
            endIndex: t.riga
          }
        }
      }]
    }
  });
  svuotaCoda();
  return { msg: "Viaggio di " + autista + " eliminato." };
}

async function azRifornimento(req) {
  if (req.idRifornimento) {
    const colA = await leggiTesto("REGISTRO_RIFORNIMENTI!A2:A100000");
    for (const r of colA) {
      if (String(r[0]) === String(req.idRifornimento)) return { msg: "Rifornimento gia' salvato!" };
    }
  }
  const id = String(req.idRifornimento || ("r-" + Date.now()));
  const adesso = Date.now() / 86400000 + 25569;

  await sheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "REGISTRO_RIFORNIMENTI!A1:G1",
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[id, adesso, req.targaMezzo, Number(req.litri), Number(req.costo), Number(req.km), ""]]
    }
  });

  // NUOVO: anche il rifornimento aggiorna i KM attuali del camion
  try {
    const b = await base();
    const targa = String(req.targaMezzo || "").trim();
    const riga = b.rigaMezzo ? b.rigaMezzo[targa] : null;
    if (riga && req.km !== "" && !isNaN(Number(req.km))) {
      await scrivi("ANAGRAFICA_MEZZI!E" + riga, [[Number(req.km)]]);
    }
  } catch (e) {}

  return { msg: "Rifornimento salvato!", id };
}

// Le foto restano su Apps Script? No: le carichiamo su Drive via API... richiede scope Drive.
// Soluzione semplice: la foto viene salvata dentro il foglio come link NO.
// Qui: carichiamo su Drive con l'API ufficiale usando lo stesso robot.
async function azFoto(req) {
  if (!req.fotoData) return { msg: "Nessuna foto." };

  const auth = new google.auth.JWT(
    process.env.GOOGLE_CLIENT_EMAIL,
    null,
    (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    ["https://www.googleapis.com/auth/drive.file"]
  );
  const drive = google.drive({ version: "v3", auth });

  const base64 = String(req.fotoData).split(",")[1] || req.fotoData;
  const buffer = Buffer.from(base64, "base64");
  const { Readable } = require("stream");

  const file = await drive.files.create({
    requestBody: { name: req.fotoNome || "foto.jpg" },
    media: { mimeType: req.fotoMime || "image/jpeg", body: Readable.from(buffer) },
    fields: "id, webViewLink"
  });
  try {
    await drive.permissions.create({
      fileId: file.data.id,
      requestBody: { role: "reader", type: "anyone" }
    });
  } catch (e) {}
  const url = file.data.webViewLink;

  // aggancia il link alla riga giusta
  const nomeFoglio = (req.tipo === "rifornimento") ? "REGISTRO_RIFORNIMENTI" : "REGISTRO_VIAGGI";
  const colonna    = (req.tipo === "rifornimento") ? "G" : "N";
  const colA = await leggiTesto(nomeFoglio + "!A1:A100000");
  for (let i = colA.length - 1; i >= 1; i--) {
    if (String(colA[i][0]) === String(req.id)) {
      const cella = nomeFoglio + "!" + colonna + (i + 1);
      // NUOVO: se ci sono gia' altre foto, il nuovo link si aggiunge in coda
      let esistente = "";
      try {
        const attuale = await leggiTesto(cella);
        esistente = (attuale[0] && attuale[0][0]) ? String(attuale[0][0]).trim() : "";
      } catch (e) {}
      const finale = esistente ? (esistente + "\n" + url) : url;
      await scrivi(cella, [[finale]]);
      return { msg: "Foto salvata." };
    }
  }
  return { msg: "Foto caricata ma riga non trovata.", url };
}


// ============================================================
//  MODIFICA di un viaggio gia' assegnato (solo admin)
//  Ammessa su "Assegnato" e "In Corso", mai sui Completati.
// ============================================================
async function azModifica(req) {
  const t = await trovaRiga(req.id);
  const stato = String(t.dati[15] || "").trim();
  if (stato === "Completato") {
    throw new Error("Viaggio gia' completato: non si puo' modificare.");
  }

  const scritture = [];

  if (req.dataViaggio) {
    scritture.push({ range: "REGISTRO_VIAGGI!B" + t.riga, values: [[dataSerial(req.dataViaggio)]] });
  }
  if (req.autista) {
    scritture.push({ range: "REGISTRO_VIAGGI!C" + t.riga, values: [[req.autista]] });
  }
  if (req.targaMezzo) {
    scritture.push({ range: "REGISTRO_VIAGGI!E" + t.riga, values: [[req.targaMezzo]] });
  }
  if (req.targaRimorchio != null) {
    scritture.push({ range: "REGISTRO_VIAGGI!K" + t.riga, values: [[req.targaRimorchio]] });
  }
  if (req.chiave) {
    scritture.push({ range: "REGISTRO_VIAGGI!I" + t.riga + ":J" + t.riga,
                     values: [[req.partenza || "", req.arrivo || ""]] });
    scritture.push({ range: "REGISTRO_VIAGGI!S" + t.riga, values: [[req.chiave]] });
  }
  if (req.nota != null) {
    scritture.push({ range: "REGISTRO_VIAGGI!R" + t.riga, values: [[req.nota]] });
  }

  if (!scritture.length) return { msg: "Nessuna modifica da salvare." };

  await scriviBlocco(scritture);
  svuotaCoda();
  return { msg: "Viaggio aggiornato." };
}

// ============================================================
//  RESOCONTI: consumi e attivita' per camion, mese per mese
// ============================================================
// ============================================================
//  CONTROLLI  (solo admin): cerca errori e dati sospetti
//  nei viaggi e nei rifornimenti del periodo scelto.
// ============================================================

// Numero da cella del foglio: numero vero, "84020,5", "194500", vuoto
function numC(v) {
  if (typeof v === "number") return v;
  const t = String(v == null ? "" : v).replace(/\s/g, "").replace(",", ".");
  if (t === "") return null;
  const n = Number(t);
  return isNaN(n) ? null : n;
}

// Data da cella: numero seriale oppure testo "05/10/2026 14.03.47"
function serialC(v) {
  if (typeof v === "number") return v;
  const m = String(v || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2})[.:](\d{2})(?:[.:](\d{2}))?)?/);
  if (!m) return null;
  const ms = Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  return ms / 86400000 + 25569;
}

function dataC(serial) {
  const d = new Date(Math.round((serial - 25569) * 86400000));
  return String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0");
}

function kmC(n) {
  return n == null ? "—" : Math.round(n).toLocaleString("it-IT");
}

function oggiSerialRoma() {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone: "Europe/Rome", year: "numeric", month: "2-digit", day: "2-digit" })
    .format(new Date()).split("-");
  return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000 + 25569;
}

/**
 * righeV: REGISTRO_VIAGGI da riga 2 (colonne A..AD)
 * righeR: REGISTRO_RIFORNIMENTI da riga 2 (colonne A..G)
 * mezzi:  { targa: nome } dall'anagrafica
 * ignorati: { chiave: true }
 * dal, al: numeri seriali (giorno intero, compresi)
 */
function calcolaControlli(righeV, righeR, mezzi, ignorati, dal, al, oggi) {
  const out = [];
  const nel = s => s != null && Math.floor(s) >= dal && Math.floor(s) <= al;
  const nome = t => (mezzi[t] ? mezzi[t] + " · " : "") + t;
  function segnala(o) { if (!ignorati[o.chiave]) out.push(o); else nIgnorati++; }
  let nIgnorati = 0;

  // ---------------- VIAGGI ----------------
  const viaggi = [];
  righeV.forEach((r, i) => {
    const s = serialC(r[1]);
    if (s == null) return;
    viaggi.push({
      riga: i + 2, id: String(r[0] || "riga" + (i + 2)), giorno: Math.floor(s),
      autista: String(r[2] || ""), targa: String(r[4] || "").trim(),
      kmP: numC(r[5]), kmA: numC(r[6]), km: numC(r[7]),
      tratta: String(r[18] || r[9] || ""), stato: String(r[15] || ""),
      ricavo: numC(r[29])
    });
  });
  const descV = v => dataC(v.giorno) + " · " + v.autista + " · " + v.tratta;
  const notte = v => /notte/i.test(v.tratta);

  viaggi.forEach(v => {
    if (!nel(v.giorno)) return;
    if (v.stato !== "Completato" && v.giorno < oggi) {
      segnala({ gruppo: "Viaggi", gravita: "media", chiave: "V-APERTO|" + v.id,
        titolo: "Viaggio rimasto " + (v.stato || "senza stato") + " · " + nome(v.targa),
        dettaglio: descV(v) + ". Non è stato chiuso: km e ricavo potrebbero mancare.",
        foglio: "REGISTRO_VIAGGI", riga: v.riga });
      return;
    }
    if (v.stato !== "Completato") return;
    if (v.ricavo == null || v.ricavo === 0) {
      segnala({ gruppo: "Viaggi", gravita: "alta", chiave: "V-RICAVO|" + v.id,
        titolo: "Viaggio senza ricavo · " + nome(v.targa),
        dettaglio: descV(v) + ". Totale ricavi a 0: controlla la tratta o il ricavo manuale.",
        foglio: "REGISTRO_VIAGGI", riga: v.riga });
    }
    if (v.km != null && v.km > 900) {
      segnala({ gruppo: "Viaggi", gravita: "alta", chiave: "V-KMALTI|" + v.id,
        titolo: "Km del viaggio troppo alti · " + nome(v.targa),
        dettaglio: descV(v) + ". " + kmC(v.km) + " km (da " + kmC(v.kmP) + " a " + kmC(v.kmA) + ").",
        foglio: "REGISTRO_VIAGGI", riga: v.riga });
    } else if (v.km != null && v.km <= 0 && !notte(v)) {
      segnala({ gruppo: "Viaggi", gravita: "media", chiave: "V-KMZERO|" + v.id,
        titolo: "Viaggio con 0 km · " + nome(v.targa),
        dettaglio: descV(v) + ". Km partenza " + kmC(v.kmP) + ", arrivo " + kmC(v.kmA) + ".",
        foglio: "REGISTRO_VIAGGI", riga: v.riga });
    }
  });

  // viaggi doppi: stesso camion, stesso giorno, stessi km di partenza e arrivo
  const visti = {};
  viaggi.forEach(v => {
    if (v.stato !== "Completato" || !v.targa || !v.km || v.kmP == null) return;
    const k = v.targa + "|" + v.giorno + "|" + v.kmP + "|" + v.kmA;
    if (visti[k] && nel(v.giorno)) {
      segnala({ gruppo: "Viaggi", gravita: "alta", chiave: "V-DOPPIO|" + v.id,
        titolo: "Possibile viaggio doppio · " + nome(v.targa),
        dettaglio: descV(v) + ". Stessi km (" + kmC(v.kmP) + " → " + kmC(v.kmA) + ") del viaggio alla riga " + visti[k].riga + ".",
        foglio: "REGISTRO_VIAGGI", riga: v.riga });
    } else if (!visti[k]) visti[k] = v;
  });

  // contachilometri: ogni viaggio deve partire dove era arrivato il precedente dello stesso camion
  const perTarga = {};
  viaggi.forEach(v => {
    if (v.stato !== "Completato" || !v.targa || v.kmP == null || v.kmA == null || !v.kmP) return;
    (perTarga[v.targa] = perTarga[v.targa] || []).push(v);
  });
  Object.keys(perTarga).forEach(t => {
    const l = perTarga[t].sort((a, b) => a.giorno - b.giorno || a.kmP - b.kmP);
    // togli i doppi già segnalati
    const u = l.filter((v, i) => !(i > 0 && v.giorno === l[i - 1].giorno && v.kmP === l[i - 1].kmP && v.kmA === l[i - 1].kmA));
    const salta = {};
    // viaggio "fuori sequenza": salto enorme prima e dopo, in senso opposto
    for (let i = 1; i < u.length - 1; i++) {
      const g1 = u[i].kmP - u[i - 1].kmA, g2 = u[i + 1].kmP - u[i].kmA;
      if (Math.abs(g1) > 500 && Math.abs(g2) > 500 && Math.sign(g1) !== Math.sign(g2)) {
        salta[i] = true; salta[i + 1] = true;
        const v = u[i];
        if (nel(v.giorno)) segnala({ gruppo: "Viaggi", gravita: "alta", chiave: "V-FUORISEQ|" + v.id,
          titolo: "Contachilometri fuori sequenza · " + nome(t),
          dettaglio: descV(v) + ". Km " + kmC(v.kmP) + " → " + kmC(v.kmA) + ", ma i viaggi prima e dopo sono intorno a " +
            kmC(u[i - 1].kmA) + ". Forse è il camion sbagliato o un errore di battitura.",
          foglio: "REGISTRO_VIAGGI", riga: v.riga });
      }
    }
    for (let i = 1; i < u.length; i++) {
      if (salta[i]) continue;
      const a = u[i - 1], b = u[i], g = b.kmP - a.kmA;
      if (!nel(b.giorno)) continue;
      if (g < -5) {
        const poco = g > -30;
        segnala({ gruppo: "Viaggi", gravita: poco ? "bassa" : "alta", chiave: "V-KMINDIETRO|" + b.id,
          titolo: (poco ? "Piccola sovrapposizione di km · " : "Km che tornano indietro · ") + nome(t),
          dettaglio: descV(b) + ". Parte da " + kmC(b.kmP) + " ma il viaggio precedente (" + dataC(a.giorno) +
            ", riga " + a.riga + ") era arrivato a " + kmC(a.kmA) + ": " + kmC(-g) + " km sovrapposti.",
          foglio: "REGISTRO_VIAGGI", riga: b.riga });
      } else if (g > 500) {
        segnala({ gruppo: "Viaggi", gravita: "media", chiave: "V-KMBUCO|" + b.id,
          titolo: "Km non registrati · " + nome(t),
          dettaglio: kmC(g) + " km tra l'arrivo del " + dataC(a.giorno) + " (" + kmC(a.kmA) + ", riga " + a.riga +
            ") e la partenza del " + dataC(b.giorno) + " (" + kmC(b.kmP) + "). Manca un viaggio o c'è un errore.",
          foglio: "REGISTRO_VIAGGI", riga: b.riga });
      }
    }
  });

  // ---------------- RIFORNIMENTI ----------------
  const rif = [];
  righeR.forEach((r, i) => {
    const s = serialC(r[1]);
    if (s == null) return;
    const targa = String(r[2] || "").trim();
    rif.push({
      riga: i + 2, s: s, giorno: Math.floor(s), targa: targa,
      id: String(r[0] || "") || ("R" + Math.round(s * 1440) + targa + String(r[4])),
      litri: numC(r[3]), costo: numC(r[4]), km: numC(r[5])
    });
  });
  const descR = x => dataC(x.giorno) + " · " + (x.litri != null ? x.litri.toLocaleString("it-IT") + " L" : "litri mancanti") +
    " · " + (x.costo != null ? x.costo.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €" : "importo mancante");

  // prezzo al litro "normale": mediana degli ultimi 150 rifornimenti
  const prezzi = rif.filter(x => x.litri > 0 && x.costo > 0).slice(-150).map(x => x.costo / x.litri).sort((a, b) => a - b);
  const mediana = prezzi.length ? prezzi[Math.floor(prezzi.length / 2)] : null;
  const sconosciute = {};

  rif.forEach(x => {
    if (!nel(x.giorno)) return;
    if (x.targa && !mezzi[x.targa]) {
      const e = sconosciute[x.targa] = sconosciute[x.targa] || { n: 0, euro: 0, righe: [] };
      e.n++; e.euro += x.costo || 0; e.righe.push(x.riga);
    }
    if (!x.litri || !x.costo || x.litri <= 0 || x.costo <= 0) {
      segnala({ gruppo: "Rifornimenti", gravita: "alta", chiave: "R-VUOTO|" + x.id,
        titolo: "Litri o importo mancanti · " + nome(x.targa),
        dettaglio: descR(x) + ".", foglio: "REGISTRO_RIFORNIMENTI", riga: x.riga });
    } else if (mediana) {
      const p = x.costo / x.litri, scarto = p / mediana - 1;
      if (Math.abs(scarto) > 0.08 || x.litri > 800) {
        segnala({ gruppo: "Rifornimenti", gravita: "alta", chiave: "R-PREZZO|" + x.id,
          titolo: "Litri o importo sbagliati · " + nome(x.targa),
          dettaglio: descR(x) + " = " + p.toLocaleString("it-IT", { maximumFractionDigits: 3 }) + " €/L, di solito " +
            mediana.toLocaleString("it-IT", { maximumFractionDigits: 2 }) + " €/L." +
            (scarto > 0.08 ? " Forse c'è dentro anche l'AdBlue." : "") +
            (x.litri > 800 ? " Litri impossibili per un pieno." : ""),
          foglio: "REGISTRO_RIFORNIMENTI", riga: x.riga });
      }
    }
    if (!x.km) {
      segnala({ gruppo: "Rifornimenti", gravita: "bassa", chiave: "R-SENZAKM|" + x.id,
        titolo: "Rifornimento senza km · " + nome(x.targa),
        dettaglio: descR(x) + ". Senza i km il consumo del camion non torna.",
        foglio: "REGISTRO_RIFORNIMENTI", riga: x.riga });
    }
  });

  // targhe che non sono in anagrafica: una sola segnalazione per targa
  Object.keys(sconosciute).forEach(t => {
    const e = sconosciute[t];
    segnala({ gruppo: "Rifornimenti", gravita: "media", chiave: "R-TARGA|" + t,
      titolo: "Targa non in anagrafica · " + t,
      dettaglio: e.n + (e.n === 1 ? " rifornimento" : " rifornimenti") + " nel periodo per " +
        e.euro.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
        " €. Questa targa non è tra i camion: errore di battitura o mezzo esterno? Righe " + e.righe.join(", ") + ".",
      foglio: "REGISTRO_RIFORNIMENTI", riga: e.righe[0] });
  });

  // km dei rifornimenti dello stesso camion: devono solo salire
  const rPerTarga = {};
  rif.forEach(x => { if (x.targa && x.km > 0) (rPerTarga[x.targa] = rPerTarga[x.targa] || []).push(x); });
  Object.keys(rPerTarga).forEach(t => {
    const l = rPerTarga[t].sort((a, b) => a.s - b.s);
    const salta = {};
    for (let i = 1; i < l.length - 1; i++) {
      const g1 = l[i].km - l[i - 1].km, g2 = l[i + 1].km - l[i].km;
      if (Math.abs(g1) > 3000 && Math.abs(g2) > 3000 && Math.sign(g1) !== Math.sign(g2)) {
        salta[i] = true; salta[i + 1] = true;
        if (nel(l[i].giorno)) segnala({ gruppo: "Rifornimenti", gravita: "alta", chiave: "R-FUORISEQ|" + l[i].id,
          titolo: "Km del rifornimento fuori sequenza · " + nome(t),
          dettaglio: descR(l[i]) + ". Km " + kmC(l[i].km) + ", ma prima e dopo il camion era intorno a " + kmC(l[i - 1].km) + ".",
          foglio: "REGISTRO_RIFORNIMENTI", riga: l[i].riga });
      }
    }
    for (let i = 1; i < l.length; i++) {
      if (salta[i] || !nel(l[i].giorno)) continue;
      const g = l[i].km - l[i - 1].km;
      if (g < -5) {
        segnala({ gruppo: "Rifornimenti", gravita: "alta", chiave: "R-KMINDIETRO|" + l[i].id,
          titolo: "Km del rifornimento più bassi del precedente · " + nome(t),
          dettaglio: descR(l[i]) + ". Km " + kmC(l[i].km) + ", ma il " + dataC(l[i - 1].giorno) + " (riga " + l[i - 1].riga +
            ") erano " + kmC(l[i - 1].km) + ".",
          foglio: "REGISTRO_RIFORNIMENTI", riga: l[i].riga });
      }
    }
  });

  // ---------------- CAMION ----------------
  if (al - dal >= 4) {
    const kmPer = {}, rifPer = {};
    viaggi.forEach(v => { if (v.stato === "Completato" && nel(v.giorno) && v.targa) kmPer[v.targa] = (kmPer[v.targa] || 0) + (v.km || 0); });
    rif.forEach(x => { if (nel(x.giorno) && x.targa) rifPer[x.targa] = (rifPer[x.targa] || 0) + 1; });
    Object.keys(kmPer).forEach(t => {
      if (kmPer[t] >= 1000 && !rifPer[t]) {
        segnala({ gruppo: "Camion", gravita: "media", chiave: "C-SENZARIF|" + t + "|" + dal + "|" + al,
          titolo: "Nessun rifornimento registrato · " + nome(t),
          dettaglio: kmC(kmPer[t]) + " km di viaggi nel periodo ma nessun rifornimento: manca una ricevuta o è sotto un'altra targa?",
          foglio: "REGISTRO_RIFORNIMENTI", riga: null });
      }
    });
  }

  const peso = { alta: 0, media: 1, bassa: 2 };
  out.sort((a, b) => peso[a.gravita] - peso[b.gravita]);
  return {
    segnalazioni: out,
    conteggi: { alta: out.filter(o => o.gravita === "alta").length, media: out.filter(o => o.gravita === "media").length,
                bassa: out.filter(o => o.gravita === "bassa").length },
    ignorati: nIgnorati
  };
}

async function azControlli(req) {
  const b = await base();
  const oggi = oggiSerialRoma();
  const dal = req.dal ? Math.floor(dataSerial(req.dal)) : oggi - 6;
  const al = req.al ? Math.floor(dataSerial(req.al)) : oggi;

  const risp = await sheets().spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID,
    ranges: ["REGISTRO_VIAGGI!A2:AD" + (FINESTRA_MAX + 1), "REGISTRO_RIFORNIMENTI!A2:G5000", "CONTROLLI_IGNORATI!A2:A5000"],
    valueRenderOption: "UNFORMATTED_VALUE",
    dateTimeRenderOption: "SERIAL_NUMBER"
  }).catch(async e => {
    // se il foglio CONTROLLI_IGNORATI non c'è ancora, leggo solo gli altri due
    const r2 = await sheets().spreadsheets.values.batchGet({
      spreadsheetId: SHEET_ID,
      ranges: ["REGISTRO_VIAGGI!A2:AD" + (FINESTRA_MAX + 1), "REGISTRO_RIFORNIMENTI!A2:G5000"],
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "SERIAL_NUMBER"
    });
    r2.data.valueRanges.push({ values: [] });
    return r2;
  });
  const blocchi = (risp.data.valueRanges || []).map(v => v.values || []);
  const ignorati = {};
  (blocchi[2] || []).forEach(r => { if (r[0]) ignorati[String(r[0])] = true; });

  const ris = calcolaControlli(blocchi[0], blocchi[1], b.mezzi, ignorati, dal, al, oggi);
  ris.dal = serialToISO(dal);
  ris.al = serialToISO(al);
  return ris;
}

// Camion al lavoro e camion fermi in una giornata (data "AAAA-MM-GG", vuota = oggi)
async function azGiornata(req) {
  const b = await base();
  const c = await coda(true);
  const giorno = req.data ? Math.floor(dataSerial(req.data)) : oggiSerialRoma();

  const lavoro = {}, ultimo = {};
  c.righe.forEach(r => {
    const targa = String(r[4] || "").trim();
    const s = serialC(r[1]);
    if (!targa || s == null) return;
    const g = Math.floor(s);
    if (g < giorno && (!ultimo[targa] || g > ultimo[targa])) ultimo[targa] = g;
    if (g !== giorno) return;
    (lavoro[targa] = lavoro[targa] || []).push({
      autista: String(r[2] || ""),
      tratta: String(r[18] || "") || (String(r[8] || "") + " ➔ " + String(r[9] || "")),
      stato: String(r[15] || "").trim(),
      rimorchio: etich(b.rimorchi, r[10]),
      ora: typeof r[19] === "number" ? Math.round(r[19] * 1440) : String(r[19] || "")
    });
  });

  const nomeDi = t => b.mezzi[t] || t;
  const alLavoro = Object.keys(lavoro).map(t => ({ targa: t, nome: nomeDi(t), viaggi: lavoro[t] }))
    .sort((x, y) => x.nome.localeCompare(y.nome));
  const fermi = Object.keys(b.mezzi).filter(t => !lavoro[t]).map(t => ({
    targa: t, nome: nomeDi(t),
    ultimo: ultimo[t] ? serialToData(ultimo[t]) : "",
    giorni: ultimo[t] ? giorno - ultimo[t] : null
  })).sort((x, y) => (y.giorni == null ? 9999 : y.giorni) - (x.giorni == null ? 9999 : x.giorni));

  return { data: serialToISO(giorno), alLavoro: alLavoro, fermi: fermi, totaleCamion: Object.keys(b.mezzi).length };
}

async function azIgnoraControllo(req) {
  if (!req.chiave) throw new Error("Segnalazione mancante.");
  await sheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "CONTROLLI_IGNORATI!A1:D1",
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [[String(req.chiave), oraLocale() + " " + serialToData(oggiSerialRoma()), String(req.titolo || ""), String(req.nota || "")]] }
  });
  return { msg: "Segnalazione ignorata." };
}

// ============================================================
//  MODIFICA VIAGGI (ufficio): cerca e corregge i viaggi, anche chiusi.
//  Si scrive solo nelle colonne "a mano" (A..V e AB): le colonne
//  con le formule (W, X, Y, Z, AA, AC, AD) non vengono mai toccate.
// ============================================================

// campo -> colonna e tipo
const CAMPI_VIAGGIO = {
  data:      { col: "B", i: 1,  tipo: "data",  nome: "Data" },
  autista:   { col: "C", i: 2,  tipo: "testo", nome: "Autista" },
  targa:     { col: "E", i: 4,  tipo: "testo", nome: "Camion" },
  kmP:       { col: "F", i: 5,  tipo: "num",   nome: "KM partenza" },
  kmA:       { col: "G", i: 6,  tipo: "num",   nome: "KM arrivo" },
  rimorchio: { col: "K", i: 10, tipo: "testo", nome: "Rimorchio" },
  ton:       { col: "L", i: 11, tipo: "num",   nome: "Tonnellate" },
  ore:       { col: "M", i: 12, tipo: "num",   nome: "Ore di lavoro" },
  stato:     { col: "P", i: 15, tipo: "testo", nome: "Stato" },
  oraFine:   { col: "Q", i: 16, tipo: "ora",   nome: "Ora fine" },
  nota:      { col: "R", i: 17, tipo: "testo", nome: "Nota ufficio" },
  tratta:    { col: "S", i: 18, tipo: "testo", nome: "Tratta" },
  oraInizio: { col: "T", i: 19, tipo: "ora",   nome: "Ora inizio" },
  nViaggi:   { col: "U", i: 20, tipo: "num",   nome: "Numero viaggi" },
  notaAut:   { col: "V", i: 21, tipo: "testo", nome: "Nota autista" },
  ricMan:    { col: "AB", i: 27, tipo: "num",  nome: "Ricavo manuale" }
};
const STATI_VIAGGIO = ["Assegnato", "In Corso", "Completato"];

// valore della cella -> valore "pulito" per il confronto e per l'app
function valoreCampo(def, v) {
  if (def.tipo === "data") { const s = serialC(v); return s == null ? "" : serialToISO(Math.floor(s)); }
  if (def.tipo === "num") { const n = numC(v); return n == null ? "" : n; }
  if (def.tipo === "ora") {
    if (typeof v === "number") { const m = Math.round(v * 1440) % 1440; return String(Math.floor(m / 60)).padStart(2, "0") + ":" + String(m % 60).padStart(2, "0"); }
    return String(v == null ? "" : v).trim();
  }
  return String(v == null ? "" : v).trim();
}

function viaggioDaRiga(r, riga, b) {
  const v = { id: String(r[0] || ""), riga: riga };
  Object.keys(CAMPI_VIAGGIO).forEach(k => { v[k] = valoreCampo(CAMPI_VIAGGIO[k], r[CAMPI_VIAGGIO[k].i]); });
  v.km = numC(r[7]);
  v.partenza = String(r[8] || "");
  v.arrivo = String(r[9] || "");
  v.foto = String(r[13] || "").split(/\s+/).filter(x => /^https?:/.test(x));
  v.ricavoTariffa = numC(r[25]);
  v.ricavoTotale = numC(r[29]);
  v.tipo = String(r[26] || "");
  v.categoria = String(r[28] || "");
  v.nomeCamion = b.mezzi[v.targa] || v.targa;
  return v;
}

async function leggiRegistroCompleto() {
  return leggi("REGISTRO_VIAGGI!A2:AD" + (FINESTRA_MAX + 1));
}

async function azCercaViaggi(req) {
  const b = await base();
  const righe = await leggiRegistroCompleto();
  const oggi = oggiSerialRoma();
  const dal = req.dal ? Math.floor(dataSerial(req.dal)) : oggi - 1;
  const al = req.al ? Math.floor(dataSerial(req.al)) : oggi;
  const testo = String(req.testo || "").trim().toLowerCase();

  const out = [];
  righe.forEach((r, i) => {
    if (!r[0]) return;
    const s = serialC(r[1]);
    if (s == null) return;
    const g = Math.floor(s);
    if (g < dal || g > al) return;
    if (req.targa && String(r[4] || "").trim() !== req.targa) return;
    if (req.autista && String(r[2] || "").trim() !== req.autista) return;
    if (req.stato && String(r[15] || "").trim() !== req.stato) return;
    if (testo && [r[2], r[4], r[18], r[8], r[9], r[17], r[21]].join(" ").toLowerCase().indexOf(testo) < 0) return;
    out.push(viaggioDaRiga(r, i + 2, b));
  });
  out.sort((x, y) => (x.data < y.data ? 1 : x.data > y.data ? -1 : y.riga - x.riga));
  return { viaggi: out.slice(0, 300), troppi: out.length > 300, totale: out.length };
}

async function azSalvaViaggio(req) {
  const b = await base();
  const chi = String(req.chi || "").trim();
  if (!chi) throw new Error("Scrivi il tuo nome in \"Chi modifica\".");
  const dopo = req.dopo || {}, prima = req.prima || {};
  const campi = Object.keys(dopo).filter(k => CAMPI_VIAGGIO[k]);
  if (!campi.length) return { msg: "Nessuna modifica da salvare." };

  // trovo la riga per ID su tutto il registro (le righe possono spostarsi)
  const righe = await leggiRegistroCompleto();
  let idx = -1;
  for (let i = righe.length - 1; i >= 0; i--) if (String(righe[i][0]) === String(req.id)) { idx = i; break; }
  if (idx < 0) throw new Error("Viaggio non trovato: forse è stato eliminato. Ricarica l'elenco.");
  const r = righe[idx], riga = idx + 2;

  // qualcun altro lo ha cambiato nel frattempo?
  const cambiati = campi.filter(k => String(valoreCampo(CAMPI_VIAGGIO[k], r[CAMPI_VIAGGIO[k].i])) !== String(prima[k] == null ? "" : prima[k]));
  if (cambiati.length) {
    throw new Error("Nel frattempo qualcuno ha cambiato " + cambiati.map(k => CAMPI_VIAGGIO[k].nome).join(", ") +
      " di questo viaggio. Riapri il viaggio e rifai la modifica.");
  }

  // controlli sui valori nuovi
  const err = [];
  const nuovo = {};
  campi.forEach(k => {
    const def = CAMPI_VIAGGIO[k];
    let v = dopo[k];
    if (def.tipo === "num") {
      if (v === "" || v == null) { nuovo[k] = ""; return; }
      v = numC(v);
      if (v == null || v < 0) { err.push(def.nome + ": numero non valido"); return; }
    } else if (def.tipo === "data") {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v))) { err.push("Data non valida"); return; }
    } else if (def.tipo === "ora") {
      v = String(v || "").trim();
      if (v && !/^\d{1,2}:\d{2}$/.test(v)) { err.push(def.nome + ": scrivi l'ora come 06:30"); return; }
    } else {
      v = String(v == null ? "" : v).trim();
    }
    nuovo[k] = v;
  });
  if ("autista" in nuovo && b.autisti.indexOf(nuovo.autista) < 0) err.push("Autista non presente nell'elenco AUTISTI");
  if ("targa" in nuovo && !b.mezzi[nuovo.targa]) err.push("Camion non presente in ANAGRAFICA_MEZZI");
  if ("rimorchio" in nuovo && nuovo.rimorchio && !b.rimorchi[nuovo.rimorchio]) err.push("Rimorchio non presente in ANAGRAFICA_RIMORCHI");
  if ("stato" in nuovo && STATI_VIAGGIO.indexOf(nuovo.stato) < 0) err.push("Stato non valido");
  let regola = null;
  if ("tratta" in nuovo) {
    regola = b.regole.find(x => x.chiave === nuovo.tratta);
    if (!regola) err.push("Tratta non presente in REGOLE_VIAGGI");
  }
  const kmP = "kmP" in nuovo ? nuovo.kmP : numC(r[5]);
  const kmA = "kmA" in nuovo ? nuovo.kmA : numC(r[6]);
  if (kmP !== "" && kmA !== "" && kmP != null && kmA != null && kmA < kmP) err.push("I KM di arrivo sono più bassi dei KM di partenza");
  if (err.length) throw new Error(err.join(". ") + ".");

  // scritture (RAW: niente conversioni strane di testi, date e numeri)
  const dati = [];
  const metti = (col, val) => dati.push({ range: "REGISTRO_VIAGGI!" + col + riga, values: [[val]] });
  campi.forEach(k => {
    const def = CAMPI_VIAGGIO[k];
    let val = nuovo[k];
    if (def.tipo === "data") val = Math.floor(dataSerial(val));
    metti(def.col, val);
  });
  if ("kmP" in nuovo || "kmA" in nuovo) {
    metti("H", (kmP !== "" && kmA !== "" && kmP != null && kmA != null && kmP > 0) ? Math.round((kmA - kmP) * 10) / 10 : "");
  }
  if (regola) {
    metti("J", regola.arrivo || String(r[9] || ""));
    if (regola.partenza) metti("I", regola.partenza);
  }
  await sheets().spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { valueInputOption: "RAW", data: dati }
  });

  // registro delle modifiche (se il foglio non c'e', la modifica resta valida)
  const quando = serialToData(oggiSerialRoma()) + " " + oraLocale();
  const dataV = valoreCampo(CAMPI_VIAGGIO.data, r[1]);
  try {
    await sheets().spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "MODIFICHE_VIAGGI!A1:I1",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: campi.map(k => [quando, chi, String(req.id), riga, dataV.split("-").reverse().join("/"),
        String(r[2] || ""), CAMPI_VIAGGIO[k].nome, String(prima[k] == null ? "" : prima[k]), String(nuovo[k])]) }
    });
  } catch (e) {}

  svuotaCoda();
  const aggiornato = (await leggi("REGISTRO_VIAGGI!A" + riga + ":AD" + riga))[0] || [];
  return { msg: "Viaggio aggiornato (" + campi.length + (campi.length === 1 ? " campo)." : " campi)."), viaggio: viaggioDaRiga(aggiornato, riga, b) };
}

// Le ore possono arrivare come "4:30", "4,5" o numero: le riporto a decimale
function parseOre(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v < 1 ? v * 24 : v;   // se e' un orario del foglio
  const t = String(v).trim().replace(",", ".");
  if (t.indexOf(":") >= 0) {
    const p = t.split(":");
    return (Number(p[0]) || 0) + (Number(p[1]) || 0) / 60;
  }
  const n = Number(t);
  return isNaN(n) ? 0 : n;
}

// ------------------------------------------------------------
//  Punto di ingresso
// ------------------------------------------------------------
exports.handler = async function (event) {
  const rispondi = (corpo, codice) => ({
    statusCode: codice || 200,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Allow-Methods": "POST, OPTIONS"
    },
    body: JSON.stringify(corpo)
  });

  if (event.httpMethod === "OPTIONS") return rispondi({ ok: true });
  if (event.httpMethod === "GET") return rispondi({ ok: true, msg: "API Priamus (Netlify) attiva." });

  let req;
  try { req = JSON.parse(event.body); }
  catch (e) { return rispondi({ ok: false, msg: "Richiesta non valida." }); }

  if (req.segreto !== SEGRETO) return rispondi({ ok: false, msg: "Accesso non autorizzato." });

  try {
    let dati;
    switch (req.azione) {
      case "ping":          dati = { pong: true }; break;
      case "elenchi":       dati = await azElenchi(); break;
      case "aggiorna":      svuotaCache(); svuotaCoda(); dati = await azElenchi(); break;
      case "login":         dati = await azLogin(req); break;
      case "viaggiAutista": dati = await azViaggiAutista(req); break;
      case "statoFlotta":   dati = await azStatoFlotta(); break;
      case "assegna":       dati = await azAssegna(req); break;
      case "avvia":         dati = await azAvvia(req); break;
      case "termina":       dati = await azTermina(req); break;
      case "elimina":       dati = await azElimina(req); break;
      case "rifornimento":  dati = await azRifornimento(req); break;
      case "foto":          dati = await azFoto(req); break;
      case "modifica":      dati = await azModifica(req); break;
      case "controlli":     dati = await azControlli(req); break;
      case "ignoraControllo": dati = await azIgnoraControllo(req); break;
      case "giornata":      dati = await azGiornata(req); break;
      case "cercaViaggi":   dati = await azCercaViaggi(req); break;
      case "salvaViaggio":  dati = await azSalvaViaggio(req); break;
      default: return rispondi({ ok: false, msg: "Azione sconosciuta: " + req.azione });
    }
    return rispondi({ ok: true, dati });
  } catch (err) {
    return rispondi({ ok: false, msg: String(err.message || err) });
  }
};
