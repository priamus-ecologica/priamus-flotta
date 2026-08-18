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
async function azResoconti(req) {
  const b = await base();

  const [righeV, righeR] = await Promise.all([
    leggi("REGISTRO_VIAGGI!A2:V" + (FINESTRA_MAX + 1)),
    leggi("REGISTRO_RIFORNIMENTI!A2:G5000")
  ]);

  // elenco dei mesi che hanno almeno un dato
  const insieme = {};
  righeV.forEach(r => { const m = serialToMese(r[1]); if (m) insieme[m] = true; });
  righeR.forEach(r => { const m = serialToMese(r[1]); if (m) insieme[m] = true; });
  const oggi = new Date();
  const meseCorrente = oggi.getFullYear() + "-" + String(oggi.getMonth() + 1).padStart(2, "0");
  insieme[meseCorrente] = true;
  const mesi = Object.keys(insieme).sort().reverse();
  const mese = req.mese || (mesi.indexOf(meseCorrente) >= 0 ? meseCorrente : (mesi[0] || meseCorrente));

  function vuoto() {
    return { viaggi: 0, rotazioni: 0, km: 0, tonnellate: 0, ore: 0, litri: 0, spesa: 0 };
  }

  const perTarga = {};
  const storico = {};   // targa -> { mese -> {km, litri} }

  righeV.forEach(r => {
    const targa = String(r[4] || "").trim();
    if (!targa) return;
    const m = serialToMese(r[1]);
    if (!m) return;

    const km = Number(r[7]) || 0;
    if (!storico[targa]) storico[targa] = {};
    if (!storico[targa][m]) storico[targa][m] = { km: 0, litri: 0 };
    storico[targa][m].km += km;

    if (m !== mese) return;
    if (!perTarga[targa]) perTarga[targa] = vuoto();
    const p = perTarga[targa];
    p.viaggi += 1;
    p.rotazioni += Number(r[20]) || 0;
    p.km += km;
    p.tonnellate += Number(r[11]) || 0;
    const ore = parseOre(r[12]);
    p.ore += ore;
  });

  righeR.forEach(r => {
    const targa = String(r[2] || "").trim();
    if (!targa) return;
    const m = serialToMese(r[1]);
    if (!m) return;

    const litri = Number(r[3]) || 0;
    if (!storico[targa]) storico[targa] = {};
    if (!storico[targa][m]) storico[targa][m] = { km: 0, litri: 0 };
    storico[targa][m].litri += litri;

    if (m !== mese) return;
    if (!perTarga[targa]) perTarga[targa] = vuoto();
    perTarga[targa].litri += litri;
    perTarga[targa].spesa += Number(r[4]) || 0;
  });

  const camion = Object.keys(perTarga).map(targa => {
    const p = perTarga[targa];
    return {
      targa: targa,
      nome: b.mezzi[targa] || targa,
      viaggi: p.viaggi,
      rotazioni: p.rotazioni,
      km: Math.round(p.km),
      tonnellate: Math.round(p.tonnellate * 10) / 10,
      ore: Math.round(p.ore * 10) / 10,
      litri: Math.round(p.litri * 10) / 10,
      spesa: Math.round(p.spesa * 100) / 100,
      consumo: (p.km > 0 && p.litri > 0) ? Math.round(p.litri / p.km * 1000) / 10 : null,
      euroKm: (p.km > 0 && p.spesa > 0) ? Math.round(p.spesa / p.km * 1000) / 1000 : null
    };
  }).sort((a, b2) => b2.km - a.km);

  // andamento degli ultimi 6 mesi, solo dove il consumo e' calcolabile
  const ultimi = mesi.slice().sort().slice(-6);
  const andamento = {};
  Object.keys(storico).forEach(targa => {
    andamento[targa] = ultimi.map(m => {
      const d = storico[targa][m];
      const c = (d && d.km > 0 && d.litri > 0) ? Math.round(d.litri / d.km * 1000) / 10 : null;
      return { mese: m, consumo: c };
    });
  });

  const tot = camion.reduce((a, c) => ({
    viaggi: a.viaggi + c.viaggi, rotazioni: a.rotazioni + c.rotazioni,
    km: a.km + c.km, tonnellate: a.tonnellate + c.tonnellate, ore: a.ore + c.ore,
    litri: a.litri + c.litri, spesa: a.spesa + c.spesa
  }), { viaggi: 0, rotazioni: 0, km: 0, tonnellate: 0, ore: 0, litri: 0, spesa: 0 });

  tot.tonnellate = Math.round(tot.tonnellate * 10) / 10;
  tot.ore = Math.round(tot.ore * 10) / 10;
  tot.litri = Math.round(tot.litri * 10) / 10;
  tot.spesa = Math.round(tot.spesa * 100) / 100;
  tot.consumo = (tot.km > 0 && tot.litri > 0) ? Math.round(tot.litri / tot.km * 1000) / 10 : null;
  tot.euroKm = (tot.km > 0 && tot.spesa > 0) ? Math.round(tot.spesa / tot.km * 1000) / 1000 : null;

  return { mese: mese, mesi: mesi, camion: camion, totali: tot, mesiGrafico: ultimi, andamento: andamento };
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
      case "resoconti":     dati = await azResoconti(req); break;
      default: return rispondi({ ok: false, msg: "Azione sconosciuta: " + req.azione });
    }
    return rispondi({ ok: true, dati });
  } catch (err) {
    return rispondi({ ok: false, msg: String(err.message || err) });
  }
};
