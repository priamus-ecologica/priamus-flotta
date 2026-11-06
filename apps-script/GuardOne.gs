// ============================================================
//  SATELLITARE GUARDONE (ATLAS / WTrack) - scarico dati via API
//
//  Ogni sera scarica dal satellitare, per tutti i camion:
//    - GPS_GIORNI : km del giorno e contachilometri (una riga per camion e giorno)
//    - GPS_SOSTE  : le soste di almeno 10 minuti, con coordinate
//    - GPS_VIAGGI : i tragitti, con orari, km e coordinate
//  I dati servono per controllare il REGISTRO_VIAGGI.
//
//  Usa SOLO chiamate di lettura (Login + GetList...). Non manda
//  comandi ai camion e non modifica niente sul portale.
//
//  Credenziali: NON stanno nel codice (il repository GitHub è
//  pubblico). Si inseriscono dal menu
//  🚛 Priamus > 🛰️ Satellitare > Inserisci utente e password
//  e finiscono nelle Proprietà script (GUARDONE_USER, GUARDONE_PASSWORD).
//
//  Menu: 🚛 Priamus > 🛰️ Satellitare
// ============================================================

const GO_BASE = "https://getposition0.guardoneitalia.it/WebAPI/API/";
const GO_TZ = "Europe/Rome";
const GO_ORA = 22;                       // ora dello scarico serale
const GO_FUNZIONE = "GPS_SCARICO_SERALE";
const GO_SOSTA_MIN = 10;                 // soste più corte non vengono salvate
const GO_TRAGITTO_MIN_KM = 0.5;          // tragitti più corti non vengono salvati

const GO_FOGLI = {
  GPS_GIORNI: ["Data", "Targa", "PerifId", "Km_GPS", "Km_Contatore_Inizio", "Km_Contatore_Fine",
               "Km_Contatore", "Prima_Accensione", "Ultimo_Spegnimento", "Minuti_Attivita", "Aggiornato"],
  GPS_SOSTE: ["Data", "Targa", "Inizio", "Fine", "Minuti", "Lat", "Lon", "Km_GPS", "Mappa"],
  GPS_VIAGGI: ["Data", "Targa", "Inizio", "Fine", "Km", "Lat_Partenza", "Lon_Partenza",
               "Lat_Arrivo", "Lon_Arrivo", "Km_GPS_Fine", "Vel_Max"]
};

// ------------------------------------------------------------
//  MENU
// ------------------------------------------------------------

/** Menu: salva utente e password GuardOne nelle Proprietà script. */
function GPS_CREDENZIALI() {
  const ui = SpreadsheetApp.getUi();
  const u = ui.prompt("Satellitare GuardOne", "Username dell'utente sviluppatore (es. 41494DEV):",
    ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK || !u.getResponseText().trim()) return;
  const p = ui.prompt("Satellitare GuardOne", "Password (quella nuova, impostata al primo accesso):",
    ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK || !p.getResponseText()) return;

  const props = PropertiesService.getScriptProperties();
  props.setProperty("GUARDONE_USER", u.getResponseText().trim());
  props.setProperty("GUARDONE_PASSWORD", p.getResponseText());
  props.deleteProperty("GUARDONE_LOGIN_MODO");
  CacheService.getScriptCache().remove("GO_TOKEN");

  try {
    goToken_();
    ui.alert("Credenziali salvate e login riuscito ✅\n\nOra puoi usare \"Scarica ieri (prova)\".");
  } catch (e) {
    ui.alert("Credenziali salvate, ma il login NON è riuscito:\n\n" + e.message +
      "\n\nControlla username e password e reinseriscili.");
  }
}

/** Menu: prova completa sul giorno di ieri. */
function GPS_PROVA() {
  const ieri = goGiorno_(-1);
  const r = goScaricaGiorno_(ieri);
  SpreadsheetApp.getUi().alert("Scarico del " + goDataIt_(ieri) + " completato ✅\n\n" +
    "Camion: " + r.camion + "\nSoste: " + r.soste + "\nTragitti: " + r.viaggi +
    "\n\nI dati sono nei fogli GPS_GIORNI, GPS_SOSTE e GPS_VIAGGI.");
}

/** Menu: scarica un giorno a scelta. */
function GPS_SCARICA_DATA() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Satellitare GuardOne", "Giorno da scaricare (gg/mm/aaaa):", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) { ui.alert("Data non valida. Esempio: 05/11/2026"); return; }
  const giorno = m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  const x = goScaricaGiorno_(giorno);
  ui.alert("Scarico del " + goDataIt_(giorno) + " completato ✅\n\nCamion: " + x.camion +
    "\nSoste: " + x.soste + "\nTragitti: " + x.viaggi);
}

/** Menu: attiva lo scarico automatico ogni sera. */
function GPS_ATTIVA() {
  goTogliTrigger_();
  ScriptApp.newTrigger(GO_FUNZIONE).timeBased().atHour(GO_ORA).nearMinute(30)
    .everyDays(1).inTimezone(GO_TZ).create();
  SpreadsheetApp.getUi().alert("Scarico automatico attivato: ogni sera verso le " + GO_ORA +
    ":30 scarica i dati di oggi e ricontrolla quelli di ieri.");
}

/** Menu: spegne lo scarico automatico. */
function GPS_DISATTIVA() {
  const n = goTogliTrigger_();
  SpreadsheetApp.getUi().alert(n ? "Scarico automatico disattivato." : "Lo scarico automatico non era attivo.");
}

/** Lanciata dal trigger ogni sera: scarica oggi e riscarica ieri (viaggi notturni). */
function GPS_SCARICO_SERALE() {
  goScaricaGiorno_(goGiorno_(-1));
  goScaricaGiorno_(goGiorno_(0));
}

// ------------------------------------------------------------
//  SCARICO DI UN GIORNO
// ------------------------------------------------------------

/** Scarica un giorno (yyyy-MM-dd) e lo scrive nei tre fogli, sostituendo i dati vecchi dello stesso giorno. */
function goScaricaGiorno_(giorno) {
  const q = { DateFrom: giorno, DateTo: giorno };
  const riepilogo = goCall_("AnalysisData/GetListDailySummary", Object.assign({ GroupPeriod: "Day" }, q));
  const contatori = goCall_("MessageData/GetListCounters", q);
  const soste = goCall_("HistoryData/GetListStopSummary", q);
  const tragitti = goCall_("HistoryData/GetListTripSummary", q);

  const targa = x => String(x.RegistrationPlate || x.Name || "").trim().toUpperCase();
  const camion = {};
  const prendi = t => (camion[t] = camion[t] || { perif: "", km: "", cIni: "", cFin: "", on: "", off: "" });

  riepilogo.forEach(x => {
    const c = prendi(targa(x));
    c.perif = x.PerifId || c.perif;
    c.km = goNum_(x.Distance);
  });
  // i contatori arrivano in ora UTC: li porto in ora italiana
  contatori
    .slice()
    .sort((a, b) => String(a.GenerationDate).localeCompare(String(b.GenerationDate)))
    .forEach(x => {
      const c = prendi(targa(x));
      c.perif = x.PerifId || c.perif;
      const km = goNum_(x.Kms);
      const ora = goOraDaUtc_(x.GenerationDate);
      if (km !== "") { if (c.cIni === "") c.cIni = km; c.cFin = km; }
      if (x.FromMessage === "KEYON" && !c.on) c.on = ora;
      if (x.FromMessage === "KEYOFF") c.off = ora;
    });

  const adesso = Utilities.formatDate(new Date(), GO_TZ, "dd/MM/yyyy HH:mm");
  const righeGiorni = Object.keys(camion).sort().map(t => {
    const c = camion[t];
    const kmCont = (c.cIni !== "" && c.cFin !== "") ? Math.round((c.cFin - c.cIni) * 10) / 10 : "";
    const minuti = (c.on && c.off) ? goMinuti_(c.off) - goMinuti_(c.on) : "";
    return [giorno, t, c.perif, c.km, c.cIni, c.cFin, kmCont, c.on, c.off, minuti, adesso];
  });

  // soste e tragitti arrivano già in ora italiana
  const righeSoste = soste
    .filter(x => Number(x.TimeSec || 0) >= GO_SOSTA_MIN * 60)
    .map(x => {
      const lat = goNum_(x.BeginLatitude, 5), lon = goNum_(x.BeginLongitude, 5);
      return [giorno, targa(x), goOra_(x.BeginDateTime), goOra_(x.EndDateTime),
        Math.round(Number(x.TimeSec) / 60), lat, lon, goNum_(x.KmFromStart),
        (lat !== "" && lon !== "") ? "https://www.google.com/maps?q=" + lat + "," + lon : ""];
    });

  const righeViaggi = tragitti
    .filter(x => Number(x.DistanceMt || 0) >= GO_TRAGITTO_MIN_KM * 1000)
    .map(x => [giorno, targa(x), goOra_(x.BeginDateTime), goOra_(x.EndDateTime),
      Math.round(Number(x.DistanceMt) / 100) / 10,
      goNum_(x.BeginLatitude, 5), goNum_(x.BeginLongitude, 5),
      goNum_(x.EndLatitude, 5), goNum_(x.EndLongitude, 5),
      goNum_(x.KmFromStart), goNum_(x.MaxSpeed)]);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  goScrivi_(ss, "GPS_GIORNI", giorno, righeGiorni);
  goScrivi_(ss, "GPS_SOSTE", giorno, righeSoste);
  goScrivi_(ss, "GPS_VIAGGI", giorno, righeViaggi);

  return { camion: righeGiorni.length, soste: righeSoste.length, viaggi: righeViaggi.length };
}

/** Toglie le righe già presenti per quel giorno e aggiunge quelle nuove in fondo. */
function goScrivi_(ss, nome, giorno, righe) {
  let sh = ss.getSheetByName(nome);
  const intest = GO_FOGLI[nome];
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, intest.length).setValues([intest]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const ultima = sh.getLastRow();
  if (ultima > 1) {
    const date = sh.getRange(2, 1, ultima - 1, 1).getDisplayValues();
    // cancello dal basso verso l'alto, a blocchi contigui
    let fine = -1;
    for (let i = date.length - 1; i >= -1; i--) {
      const uguale = i >= 0 && date[i][0] === giorno;
      if (uguale && fine < 0) fine = i;
      if (!uguale && fine >= 0) {
        sh.deleteRows(i + 3, fine - i);
        fine = -1;
      }
    }
  }
  if (!righe.length) return;
  const da = sh.getLastRow() + 1;
  const r = sh.getRange(da, 1, righe.length, intest.length);
  r.setNumberFormat("@").setValues(righe.map(x => x.map(v => (v === null || v === undefined) ? "" : String(v))));
  // numeri come numeri (tranne data, targa e orari)
  const testo = { Data: 1, Targa: 1, Inizio: 1, Fine: 1, Prima_Accensione: 1, Ultimo_Spegnimento: 1, Mappa: 1, Aggiornato: 1 };
  const formato = h => /^(PerifId|Minuti|Minuti_Attivita)$/.test(h) ? "0"
                     : /^(Lat|Lon)/.test(h) ? "0.00000" : "0.0";
  intest.forEach((h, j) => {
    if (testo[h]) return;
    const col = sh.getRange(da, j + 1, righe.length, 1);
    col.setNumberFormat(formato(h)).setValues(righe.map(x => [x[j] === "" ? "" : Number(x[j])]));
  });
}

// ------------------------------------------------------------
//  API GUARDONE
// ------------------------------------------------------------

/** Chiamata POST di lettura. Ritorna l'array "data". Se il token è scaduto, rifà il login una volta. */
function goCall_(endpoint, corpo, giaRiprovato) {
  const token = goToken_();
  const risp = UrlFetchApp.fetch(GO_BASE + endpoint, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(Object.assign({ token: token }, corpo)),
    muteHttpExceptions: true
  });
  const codice = risp.getResponseCode();
  let j = null;
  try { j = JSON.parse(risp.getContentText()); } catch (e) {}
  if ((codice === 401 || codice === 403) && !giaRiprovato) {
    CacheService.getScriptCache().remove("GO_TOKEN");
    return goCall_(endpoint, corpo, true);
  }
  if (codice !== 200 || !j) {
    throw new Error("GuardOne " + endpoint + ": risposta " + codice + " " +
      String(risp.getContentText()).slice(0, 200));
  }
  return Array.isArray(j.data) ? j.data : [];
}

/** Token valido (in cache per qualche ora). */
function goToken_() {
  const cache = CacheService.getScriptCache();
  const t = cache.get("GO_TOKEN");
  if (t) return t;

  const props = PropertiesService.getScriptProperties();
  const user = props.getProperty("GUARDONE_USER");
  const pass = props.getProperty("GUARDONE_PASSWORD");
  if (!user || !pass) {
    throw new Error("Mancano le credenziali: usa 🚛 Priamus > 🛰️ Satellitare > Inserisci utente e password.");
  }

  // Il login vuole Group / Name / Password: provo le combinazioni possibili
  // e ricordo quella che funziona.
  const modi = [
    { Group: null, Name: user },
    { Group: user, Name: user },
    { Group: "41494", Name: user },
    { Group: "GO", Name: user }
  ];
  const salvato = props.getProperty("GUARDONE_LOGIN_MODO");
  const ordine = salvato !== null ? [Number(salvato)].concat(modi.map((_, i) => i).filter(i => i !== Number(salvato)))
                                  : modi.map((_, i) => i);
  let ultimoErrore = "";
  for (const i of ordine) {
    const corpo = Object.assign({ Password: pass, CodeMFA: null, UserOptions: null }, modi[i]);
    const risp = UrlFetchApp.fetch(GO_BASE + "Login/Login", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify(corpo), muteHttpExceptions: true
    });
    let j = null;
    try { j = JSON.parse(risp.getContentText()); } catch (e) {}
    if (risp.getResponseCode() === 200 && j && j.token) {
      props.setProperty("GUARDONE_LOGIN_MODO", String(i));
      const durata = Math.min(Number(j.validity || 14400) - 600, 21000);  // margine di 10 minuti
      cache.put("GO_TOKEN", j.token, Math.max(durata, 600));
      return j.token;
    }
    ultimoErrore = "risposta " + risp.getResponseCode() + " " + String(risp.getContentText()).slice(0, 150);
  }
  throw new Error("Login GuardOne non riuscito (" + ultimoErrore + ")");
}

// ------------------------------------------------------------
//  UTILITÀ
// ------------------------------------------------------------

function goTogliTrigger_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === GO_FUNZIONE) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

/** Giorno in formato yyyy-MM-dd, spostato di "delta" giorni da oggi. */
function goGiorno_(delta) {
  const d = new Date();
  d.setDate(d.getDate() + delta);
  return Utilities.formatDate(d, GO_TZ, "yyyy-MM-dd");
}

function goDataIt_(g) { return g.slice(8, 10) + "/" + g.slice(5, 7) + "/" + g.slice(0, 4); }

/** "2026-11-06T05:57:00" (già ora italiana) -> "05:57" */
function goOra_(s) { return s ? String(s).slice(11, 16) : ""; }

/** "2026-11-06T04:07:00" in UTC -> "05:07" in ora italiana */
function goOraDaUtc_(s) {
  if (!s) return "";
  const d = new Date(String(s).replace(/Z?$/, "Z"));
  return isNaN(d) ? "" : Utilities.formatDate(d, GO_TZ, "HH:mm");
}

function goMinuti_(hhmm) { const p = String(hhmm).split(":"); return Number(p[0]) * 60 + Number(p[1]); }

function goNum_(v, dec) {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  if (!isFinite(n)) return "";
  const f = Math.pow(10, dec === undefined ? 1 : dec);
  return Math.round(n * f) / f;
}
