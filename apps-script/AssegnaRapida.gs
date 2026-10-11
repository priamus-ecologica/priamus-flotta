// ============================================================
//  PRIAMUS - ASSEGNAZIONE RAPIDA DA MESSAGGIO
//
//  Incolli il messaggio del capo, il programma lo legge e
//  prepara i viaggi. Tu controlli, correggi, e li crei.
//
//  DOVE VA: progetto Apps Script del DATABASE
//  (Estensioni > Apps Script > + > Script, nome: AssegnaRapida)
//
//  PRIMA VOLTA: menu "Priamus" > "Prepara le schede"
// ============================================================

const FOGLIO_LAVORO = "ASSEGNA_RAPIDA";
const FOGLIO_SINONIMI = "SINONIMI";

const RIGA_INIZIO_TESTO = 4;   // da qui in giu' si incolla il messaggio (colonna A)
const RIGA_INIZIO_GRIGLIA = 4; // da qui in giu' compaiono i viaggi (colonna C)

const COL_TESTO = 1;           // A
const COL_AUT = 3;             // C
const COL_CAM = 4;             // D
const COL_RIM = 5;             // E
const COL_TRA = 6;             // F
const COL_NOTA = 7;            // G
const COL_ESITO = 8;           // H
const COL_RAW = 9;             // I..L nascoste: testo originale dei 4 campi

const GIALLO = "#fff3cd";
const ROSSO = "#f8d7da";
const VERDE = "#d4edda";
const BIANCO = "#ffffff";

// ------------------------------------------------------------
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const assegna = ui.createMenu("📝 Assegna rapida")
    .addItem("1. Leggi il messaggio", "LEGGI_MESSAGGIO")
    .addItem("2. Crea i viaggi", "CREA_VIAGGI")
    .addSeparator()
    .addItem("Svuota tutto", "SVUOTA")
    .addItem("Prepara le schede (prima volta)", "PREPARA");

  const costi = ui.createMenu("💶 Costi dei mezzi")
    .addItem("Importa i file nuovi", "IMPORTA_COSTI")
    .addItem("Sistema i mesi", "SISTEMA_MESI")
    .addItem("Prepara il foglio (prima volta)", "PREPARA_COSTI")
    .addSeparator()
    .addItem("Attiva importazione automatica", "ATTIVA_AUTOMATISMO")
    .addItem("Disattiva importazione automatica", "DISATTIVA_AUTOMATISMO");

  const schede = ui.createMenu("🚚 Schede mezzi")
    .addItem("Prepara il confronto (prima volta)", "PREPARA_CONFRONTO")
    .addItem("Sistema il mese nelle schede", "SISTEMA_MESE_SCHEDE")
    .addSeparator()
    .addItem("Crea il registro del mese", "CREA_REGISTRO_MESE")
    .addItem("Compila €/km", "COMPILA_EURO_KM")
    .addSeparator()
    .addItem("Crea/aggiorna le schede", "CREA_SCHEDE")
    .addItem("Rifai una scheda sola", "RIFAI_SCHEDA")
    .addItem("Rifai TUTTE le schede", "RIFAI_TUTTE_LE_SCHEDE")
    .addItem("Crea il registro del giorno", "CREA_REGISTRO_GIORNO")
    .addItem("Crea il registro mezzo", "CREA_REGISTRO_MEZZO")
    .addItem("Aggiorna schede per il pianale", "AGGIORNA_SCHEDE_PIANALE")
    .addItem("Crea il registro anno", "CREA_REGISTRO_ANNO");
  

  const mail = ui.createMenu("📧 Mail serale")
    .addItem("Attiva la mail delle 20", "ATTIVA_MAIL_SERALE")
    .addItem("Mandala adesso (prova)", "PROVA_MAIL_SERALE")
    .addItem("Disattiva", "DISATTIVA_MAIL_SERALE");

  ui.createMenu("🚛 Priamus")
    .addSubMenu(assegna)
    .addSubMenu(costi)
    .addSubMenu(schede)
    .addSubMenu(mail)
    .addToUi();
}

// ------------------------------------------------------------
//  Lettura anagrafiche
// ------------------------------------------------------------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function norm_(t) {
  return String(t || "")
    .toLowerCase()
    .replace(/[àáâä]/g, "a").replace(/[èéêë]/g, "e").replace(/[ìíîï]/g, "i")
    .replace(/[òóôö]/g, "o").replace(/[ùúûü]/g, "u")
    .replace(/[^a-z0-9]/g, "");
}

function parole_(t) {
  return String(t || "")
    .toLowerCase()
    .replace(/[àáâä]/g, "a").replace(/[èéêë]/g, "e").replace(/[ìíîï]/g, "i")
    .replace(/[òóôö]/g, "o").replace(/[ùúûü]/g, "u")
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(p => p && p !== "x" && p !== "il" && p !== "la" && p !== "di");
}

// distanza di edit, per riconoscere i refusi (poliservice / poloservice)
function distanza_(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prec = [];
  for (let j = 0; j <= n; j++) prec[j] = j;
  for (let i = 1; i <= m; i++) {
    let cur = [i];
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prec[j] + 1, cur[j - 1] + 1,
                        prec[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prec = cur;
  }
  return prec[n];
}

function anagrafiche_() {
  const s = ss_();

  const autisti = [];
  const shA = s.getSheetByName("AUTISTI");
  const nA = Math.max(shA.getLastRow() - 1, 0);
  if (nA > 0) {
    shA.getRange(2, 1, nA, 1).getValues().forEach(r => {
      const nome = String(r[0] || "").trim();
      if (nome) autisti.push(nome);
    });
  }

  const mezzi = [];   // {targa, modello}
  const shM = s.getSheetByName("ANAGRAFICA_MEZZI");
  const nM = Math.max(shM.getLastRow() - 1, 0);
  if (nM > 0) {
    shM.getRange(2, 1, nM, 4).getValues().forEach(r => {
      const targa = String(r[1] || "").trim();
      const modello = String(r[3] || "").trim();
      if (targa) mezzi.push({ targa: targa, modello: modello || targa, categoria: String(r[2] || "").trim() });
    });
  }

  const rimorchi = [];  // {targa, nome}
  const shR = s.getSheetByName("ANAGRAFICA_RIMORCHI");
  const nR = Math.max(shR.getLastRow() - 1, 0);
  if (nR > 0) {
    shR.getRange(2, 1, nR, 2).getValues().forEach(r => {
      const targa = String(r[0] || "").trim();
      const nome = String(r[1] || "").trim();
      if (targa) rimorchi.push({ targa: targa, nome: nome || targa });
    });
  }

  const regole = [];  // {chiave, partenza, arrivo}
  const shReg = s.getSheetByName("REGOLE_VIAGGI");
  const nReg = Math.max(shReg.getLastRow() - 1, 0);
  if (nReg > 0) {
    shReg.getRange(2, 1, nReg, 3).getValues().forEach(r => {
      const chiave = String(r[0] || "").trim();
      if (chiave) regole.push({
        chiave: chiave,
        partenza: String(r[1] || "").trim(),
        arrivo: String(r[2] || "").trim()
      });
    });
  }

  return { autisti: autisti, mezzi: mezzi, rimorchi: rimorchi, regole: regole };
}

function sinonimi_() {
  const sh = ss_().getSheetByName(FOGLIO_SINONIMI);
  const mappa = { AUTISTA: {}, CAMION: {}, RIMORCHIO: {}, TRATTA: {} };
  if (!sh) return mappa;
  const n = Math.max(sh.getLastRow() - 1, 0);
  if (n === 0) return mappa;
  sh.getRange(2, 1, n, 3).getValues().forEach(r => {
    const tipo = String(r[0] || "").trim().toUpperCase();
    const scritto = norm_(r[1]);
    const vale = String(r[2] || "").trim();
    if (mappa[tipo] && scritto && vale) mappa[tipo][scritto] = vale;
  });
  return mappa;
}

// ------------------------------------------------------------
//  Riconoscimento dei singoli campi
//  esito: "ok" | "forse" | "no"
// ------------------------------------------------------------
function trovaAutista_(testo, ana, sin) {
  const t = norm_(testo);
  if (!t) return { valore: "", esito: "no" };

  // nome completo (l'anagrafica vince sempre sui sinonimi)
  for (const a of ana.autisti) if (norm_(a) === t) return { valore: a, esito: "ok" };

  if (sin.AUTISTA[t]) return { valore: sin.AUTISTA[t], esito: "ok" };

  // solo nome (ultima parola) oppure solo cognome (tutto il resto)
  const perNome = ana.autisti.filter(a => norm_(a.split(/\s+/).pop()) === t);
  const perCognome = ana.autisti.filter(a => {
    const p = a.split(/\s+/); p.pop();
    return norm_(p.join("")) === t;
  });

  if (perCognome.length === 1) return { valore: perCognome[0], esito: "ok" };
  if (perNome.length === 1) return { valore: perNome[0], esito: "ok" };
  if (perNome.length > 1) return { valore: perNome[0], esito: "forse" };   // nome doppio
  if (perCognome.length > 1) return { valore: perCognome[0], esito: "forse" };

  // ultimo tentativo: somiglianza
  let best = null, bestD = 99;
  ana.autisti.forEach(a => {
    const d = Math.min(distanza_(t, norm_(a)), distanza_(t, norm_(a.split(/\s+/).pop())));
    if (d < bestD) { bestD = d; best = a; }
  });
  if (best && bestD <= 2) return { valore: best, esito: "forse" };
  return { valore: "", esito: "no" };
}

function trovaCamion_(testo, ana, sin) {
  const t = norm_(testo);
  if (!t) return { valore: "", esito: "no" };

  for (const m of ana.mezzi) {
    if (norm_(m.modello) === t || norm_(m.targa) === t) return { valore: m.targa, esito: "ok" };
  }

  if (sin.CAMION[t]) return { valore: sin.CAMION[t], esito: "ok" };

  // "Pierandelo 4 assi Renault": tolgo una parola alla volta davanti finche' trovo il camion
  const parole = String(testo).trim().split(/\s+/);
  for (let k = 1; k <= parole.length - 2; k++) {
    const coda = norm_(parole.slice(k).join(" "));
    for (const m of ana.mezzi) if (norm_(m.modello) === coda) return { valore: m.targa, esito: "ok" };
    if (sin.CAMION[coda]) return { valore: sin.CAMION[coda], esito: "ok" };
  }
  let best = null, bestD = 99;
  ana.mezzi.forEach(m => {
    const d = distanza_(t, norm_(m.modello));
    if (d < bestD) { bestD = d; best = m; }
  });
  if (best && bestD <= 2) return { valore: best.targa, esito: "forse" };
  return { valore: "", esito: "no" };
}

function trovaRimorchio_(testo, ana, sin) {
  const t = norm_(testo);
  if (!t) return { valore: "", esito: "no" };

  if (sin.RIMORCHIO[t]) return { valore: sin.RIMORCHIO[t], esito: "ok" };

  for (const r of ana.rimorchi) {
    if (norm_(r.nome) === t || norm_(r.targa) === t) return { valore: r.targa, esito: "ok" };
  }
  return { valore: "", esito: "no" };
}

function trovaTratta_(testo, ana, sin) {
  const t = norm_(testo);
  if (!t) return { valore: "", esito: "no" };

  for (const r of ana.regole) if (norm_(r.chiave) === t) return { valore: r.chiave, esito: "ok" };

  if (sin.TRATTA[t]) return { valore: sin.TRATTA[t], esito: "ok" };

  // punteggio per parole in comune + somiglianza
  const pt = parole_(testo);
  let best = null, bestScore = 0;
  ana.regole.forEach(r => {
    const pr = parole_(r.chiave);
    if (!pr.length) return;
    let comuni = 0;
    pr.forEach(x => {
      if (pt.indexOf(x) >= 0) comuni++;
      else if (pt.some(y => y.length > 4 && distanza_(x, y) <= 2)) comuni += 0.9;
    });
    // conta le parole in comune rispetto al TOTALE delle parole coinvolte:
    // cosi' "ecotec x lanciano carico" sceglie "Ecotec x Lanciano" e non "Lanciano"
    const score = comuni / Math.max(pr.length, pt.length);
    if (score > bestScore) { bestScore = score; best = r.chiave; }
  });

  if (bestScore >= 0.99) return { valore: best, esito: "ok" };
  if (bestScore >= 0.5) return { valore: best, esito: "forse" };
  return { valore: "", esito: "no" };
}

// ------------------------------------------------------------
//  Lettura del messaggio
// ------------------------------------------------------------
function LEGGI_MESSAGGIO() {
  const sh = ss_().getSheetByName(FOGLIO_LAVORO);
  if (!sh) { PREPARA(); return; }

  const ultima = sh.getLastRow();
  const righe = ultima >= RIGA_INIZIO_TESTO
    ? sh.getRange(RIGA_INIZIO_TESTO, COL_TESTO, ultima - RIGA_INIZIO_TESTO + 1, 1)
        .getValues().map(r => String(r[0] || "").trim())
    : [];

  if (!righe.filter(String).length) {
    SpreadsheetApp.getUi().alert("Incolla il messaggio nella colonna A, a partire dalla riga " + RIGA_INIZIO_TESTO + ".");
    return;
  }

  const ana = anagrafiche_();
  const sin = sinonimi_();

  let data = null;
  const viaggi = [];
  const scartate = [];
  let autistaCorr = "", autistaRaw = "", camionCorr = "", camionRaw = "";
  let esitoAut = "no", esitoCam = "no";

  function nuovoViaggio_(rimorchio, rimRaw, esitoRim, est, tra) {
    return {
      autista: autistaCorr, autistaRaw: autistaRaw, esitoAut: esitoAut,
      camion: camionCorr, camionRaw: camionRaw, esitoCam: esitoCam,
      rimorchio: rimorchio, rimorchioRaw: rimRaw, esitoRim: esitoRim,
      tratta: tra.valore, trattaRaw: est.testo, esitoTra: tra.esito,
      nota: est.orario ? ("Orario " + est.orario) : ""
    };
  }
  function impostaAutista_(div) {
    autistaRaw = div.autistaRaw; camionRaw = div.camionRaw;
    const a = trovaAutista_(div.autistaRaw, ana, sin);
    const c = trovaCamion_(div.camionRaw, ana, sin);
    autistaCorr = a.valore; esitoAut = a.esito;
    camionCorr = c.valore;  esitoCam = c.esito;
  }
  // senza rimorchio va bene solo per i 4 assi
  function esitoSenzaRim_() { return eQuattroAssi_(camionCorr, ana) ? "ok" : "no"; }

  righe.forEach(riga => {
    if (!riga) return;

    // --- data: solo nelle righe prima del primo autista (es. "Viaggi lunedi' 12 ottobre" o "12/10")
    if (!data && !autistaRaw && !viaggi.length) {
      const d = leggiData_(riga);
      if (d) { data = d; return; }
    }

    // --- nota esplicita
    if (/^note\s*:/i.test(riga)) {
      const testo = riga.replace(/^note\s*:/i, "").trim();
      if (viaggi.length) {
        viaggi[viaggi.length - 1].nota =
          (viaggi[viaggi.length - 1].nota ? viaggi[viaggi.length - 1].nota + " - " : "") + testo;
      } else {
        scartate.push(riga + "  (nota senza viaggio sopra)");
      }
      return;
    }

    // --- riga di viaggio: comincia con il nome di un rimorchio
    const pezzi = riga.split(/\s+/);
    const rim = trovaRimorchio_(pezzi[0], ana, sin);
    if (rim.esito === "ok" && pezzi.length > 1) {
      const est = estraiOrario_(pezzi.slice(1).join(" "));
      const tra = trovaTratta_(est.testo, ana, sin);
      viaggi.push(nuovoViaggio_(rim.valore, pezzi[0], "ok", est, tra));
      return;
    }

    // --- autista + camion riconosciuti entrambi (anche "Delle Monache 4 assi iveco (2)")
    const div = dividiAutistaCamion_(riga, ana, sin, true);
    if (div) { impostaAutista_(div); return; }

    // --- viaggio SENZA rimorchio (4 assi): "Roda: via ripoli x coget" sotto l'autista
    const est0 = estraiOrario_(riga);
    const tra0 = trovaTratta_(est0.testo, ana, sin);
    if (autistaRaw && tra0.esito === "ok") { viaggi.push(nuovoViaggio_("", "", esitoSenzaRim_(), est0, tra0)); return; }

    // --- riga che e' SOLO un camion: non assegnabile
    const soloCamion = trovaCamion_(riga, ana, sin);
    if (soloCamion.esito === "ok") { scartate.push(riga); return; }

    // --- autista riconosciuto, camion no
    const div2 = dividiAutistaCamion_(riga, ana, sin, false);
    if (div2) { impostaAutista_(div2); return; }

    // --- tratta incerta senza rimorchio
    if (autistaRaw && tra0.esito === "forse") { viaggi.push(nuovoViaggio_("", "", esitoSenzaRim_(), est0, tra0)); return; }

    scartate.push(riga);
  });

  scriviGriglia_(sh, data, viaggi, scartate, ana);
}

// "21/08", "21/08/2026", "21-8", "Viaggi lunedi' 12 ottobre", "12 ottobre 2026"
function leggiData_(riga) {
  const t = String(riga).toLowerCase();
  const MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
                "agosto", "settembre", "ottobre", "novembre", "dicembre"];
  let gg, mm, aa = null;
  let m = t.match(/(?:^|[^\d.])(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?(?!\d)/) ||
          t.match(/^\s*(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\s*$/);
  if (m) { gg = Number(m[1]); mm = Number(m[2]); if (m[3]) aa = Number(m[3]); }
  else {
    m = t.match(new RegExp("(?:^|\\D)(\\d{1,2})\\s+(" + MESI.join("|") + ")(?:\\s+(\\d{4}))?"));
    if (!m) return null;
    gg = Number(m[1]); mm = MESI.indexOf(m[2]) + 1; if (m[3]) aa = Number(m[3]);
  }
  if (gg < 1 || gg > 31 || mm < 1 || mm > 12) return null;
  if (aa !== null && aa < 100) aa += 2000;
  if (aa === null) {
    // anno non scritto: quello piu' vicino a oggi (a fine dicembre "2 gennaio" e' l'anno dopo)
    const oggi = new Date();
    aa = oggi.getFullYear();
    const giorni = (new Date(aa, mm - 1, gg) - oggi) / 86400000;
    if (giorni < -180) aa++;
    else if (giorni > 180) aa--;
  }
  return new Date(aa, mm - 1, gg, 12, 0, 0);
}

function eQuattroAssi_(camion, ana) {
  const m = ana.mezzi.filter(x => x.targa === camion || x.modello === camion)[0];
  return !!m && /4\s*assi/i.test(m.categoria + " " + m.modello);
}

// toglie l'orario dalla tratta: "isernia 8,30" -> {testo:"isernia", orario:"8:30"}
function estraiOrario_(testo) {
  let t = String(testo);
  let orario = "";

  let m = t.match(/(?:\bore\s*)?\b(\d{1,2})[,.:](\d{2})\b/);
  if (m) {
    orario = Number(m[1]) + ":" + m[2];
    t = t.replace(m[0], " ");
  } else {
    m = t.match(/\bore\s*(\d{1,2})\b/i);
    if (m) { orario = Number(m[1]) + ":00"; t = t.replace(m[0], " "); }
  }
  return { testo: t.replace(/\s+/g, " ").trim(), orario: orario };
}

// "Massimo Mercedes 510 (2)" -> autista "Massimo", camion "Mercedes 510 (2)"
function dividiAutistaCamion_(riga, ana, sin, soloSeCompleto) {
  const pezzi = riga.split(/\s+/);
  if (pezzi.length < 2) return null;

  // provo tutti i punti di taglio, tengo quello in cui il camion e' riconosciuto
  for (let i = 1; i < pezzi.length; i++) {
    const sinistra = pezzi.slice(0, i).join(" ");
    const destra = pezzi.slice(i).join(" ");
    const c = trovaCamion_(destra, ana, sin);
    const a = trovaAutista_(sinistra, ana, sin);
    if (soloSeCompleto ? (c.esito === "ok" && a.esito === "ok") : (c.esito !== "no" && a.esito !== "no")) {
      return { autistaRaw: sinistra, camionRaw: destra };
    }
  }
  if (soloSeCompleto) return null;
  // camion non riconosciuto: prendo la prima parola come autista
  const a1 = trovaAutista_(pezzi[0], ana, sin);
  if (a1.esito !== "no") {
    return { autistaRaw: pezzi[0], camionRaw: pezzi.slice(1).join(" ") };
  }
  return null;
}

// ------------------------------------------------------------
//  Scrittura della griglia di controllo
// ------------------------------------------------------------
function scriviGriglia_(sh, data, viaggi, scartate, ana) {
  // pulizia area griglia
  const ultima = Math.max(sh.getLastRow(), RIGA_INIZIO_GRIGLIA + 60);
  sh.getRange(RIGA_INIZIO_GRIGLIA, COL_AUT, ultima, 10).clearContent().clearDataValidations()
    .setBackground(BIANCO).setNote(null);

  sh.getRange(2, COL_CAM).setValue(data || new Date());

  if (!viaggi.length) {
    SpreadsheetApp.getUi().alert("Non ho trovato viaggi nel messaggio.\nControlla di averlo incollato dalla riga " + RIGA_INIZIO_TESTO + ".");
    return;
  }

  const valori = [], sfondi = [], note = [];
  viaggi.forEach(v => {
    valori.push([v.autista, v.camion ? etichettaCamion_(v.camion, ana) : "",
                 v.rimorchio ? etichettaRim_(v.rimorchio, ana) : "",
                 v.tratta, v.nota, testoEsito_(v),
                 v.autistaRaw, v.camionRaw, v.rimorchioRaw, v.trattaRaw]);
    const peggio = peggiore_([v.esitoAut, v.esitoCam, v.esitoRim, v.esitoTra]);
    const col = peggio === "no" ? ROSSO : (peggio === "forse" ? GIALLO : VERDE);
    sfondi.push(new Array(10).fill(col));
    note.push(new Array(10).fill(null));
  });

  const r = sh.getRange(RIGA_INIZIO_GRIGLIA, COL_AUT, valori.length, 10);
  r.setValues(valori);
  r.setBackgrounds(sfondi);

  // tendine per correggere
  tendina_(sh, COL_AUT, valori.length, ana.autisti);
  tendina_(sh, COL_CAM, valori.length, ana.mezzi.map(m => m.modello));
  tendina_(sh, COL_RIM, valori.length, ana.rimorchi.map(x => x.nome));
  tendina_(sh, COL_TRA, valori.length, ana.regole.map(x => x.chiave));

  // righe scartate, sotto la griglia
  const rScarti = RIGA_INIZIO_GRIGLIA + valori.length + 1;
  sh.getRange(rScarti, COL_AUT).setValue("RIGHE NON RICONOSCIUTE (non verranno assegnate)")
    .setFontWeight("bold").setFontColor("#8a6d3b");
  if (scartate.length) {
    sh.getRange(rScarti + 1, COL_AUT, scartate.length, 1)
      .setValues(scartate.map(x => [x])).setFontColor("#8a6d3b").setFontStyle("italic");
  } else {
    sh.getRange(rScarti + 1, COL_AUT).setValue("nessuna").setFontStyle("italic");
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "Lette " + valori.length + " righe di viaggio.\n\n" +
    "VERDE = riconosciuto\nGIALLO = da controllare\nROSSO = manca qualcosa\n\n" +
    (scartate.length ? ("Righe ignorate: " + scartate.length + " (elencate sotto la griglia)\n\n") : "") +
    "Correggi le celle colorate con le tendine, poi:\nmenu Priamus > 2. Crea i viaggi"
  );
}

function peggiore_(elenco) {
  if (elenco.indexOf("no") >= 0) return "no";
  if (elenco.indexOf("forse") >= 0) return "forse";
  return "ok";
}

function testoEsito_(v) {
  const mancanti = [];
  if (v.esitoAut !== "ok") mancanti.push("autista");
  if (v.esitoCam !== "ok") mancanti.push("camion");
  if (v.esitoRim !== "ok") mancanti.push("rimorchio");
  if (v.esitoTra !== "ok") mancanti.push("tratta");
  return mancanti.length ? ("controlla: " + mancanti.join(", ")) : "ok";
}

function etichettaCamion_(targa, ana) {
  const m = ana.mezzi.filter(x => x.targa === targa)[0];
  return m ? m.modello : targa;
}
function etichettaRim_(targa, ana) {
  const r = ana.rimorchi.filter(x => x.targa === targa)[0];
  return r ? r.nome : targa;
}

function tendina_(sh, colonna, quante, elenco) {
  if (!elenco.length) return;
  const reg = SpreadsheetApp.newDataValidation()
    .requireValueInList(elenco, true).setAllowInvalid(true).build();
  sh.getRange(RIGA_INIZIO_GRIGLIA, colonna, quante, 1).setDataValidation(reg);
}

// ------------------------------------------------------------
//  Creazione dei viaggi nel registro
// ------------------------------------------------------------

// Ultima riga con un ID in colonna A.
// NON si puo' usare getLastRow(): le formule dei ricavi in W:Z
// occupano il foglio fino in fondo e sposterebbero i viaggi nuovi
// a migliaia di righe piu' giu'.
function ultimaRigaViaggi_(foglio) {
  const massimo = foglio.getMaxRows();
  if (massimo < 2) return 1;
  const ids = foglio.getRange(2, 1, massimo - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() !== "") return i + 2;
  }
  return 1;
}

function CREA_VIAGGI() {
  const s = ss_();
  const sh = s.getSheetByName(FOGLIO_LAVORO);
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert("Scheda di lavoro non trovata: esegui prima 'Prepara le schede'."); return; }

  const ana = anagrafiche_();
  const sin = sinonimi_();

  const data = sh.getRange(2, COL_CAM).getValue();
  if (!(data instanceof Date)) { ui.alert("Metti la data dei viaggi nella cella D2."); return; }

  // leggo la griglia fino alla prima riga vuota
  const righe = [];
  let r = RIGA_INIZIO_GRIGLIA;
  while (true) {
    const v = sh.getRange(r, COL_AUT, 1, 10).getValues()[0];
    if (!String(v[0] || "").trim() && !String(v[3] || "").trim()) break;
    righe.push({ riga: r, v: v });
    r++;
    if (r > RIGA_INIZIO_GRIGLIA + 300) break;
  }

  if (!righe.length) { ui.alert("Nessun viaggio da creare."); return; }

  // controllo che sia tutto completo
  const problemi = [];
  const pronti = [];
  righe.forEach(x => {
    const autista = String(x.v[0] || "").trim();
    const camionEt = String(x.v[1] || "").trim();
    const rimEt = String(x.v[2] || "").trim();
    const tratta = String(x.v[3] || "").trim();
    const nota = String(x.v[4] || "").trim();

    const mezzo = ana.mezzi.filter(m => m.modello === camionEt || m.targa === camionEt)[0];
    const rim = ana.rimorchi.filter(m => m.nome === rimEt || m.targa === rimEt)[0];
    const reg = ana.regole.filter(m => m.chiave === tratta)[0];

    const mancano = [];
    if (ana.autisti.indexOf(autista) < 0) mancano.push("autista");
    if (!mezzo) mancano.push("camion");
    if (!rim && !(rimEt === "" && mezzo && eQuattroAssi_(mezzo.targa, ana))) mancano.push("rimorchio");
    if (!reg) mancano.push("tratta");

    if (mancano.length) {
      problemi.push("riga " + x.riga + ": manca " + mancano.join(", "));
    } else {
      pronti.push({ riga: x.riga, autista: autista, mezzo: mezzo, rim: rim, reg: reg, nota: nota, raw: x.v });
    }
  });

  if (problemi.length) {
    ui.alert("Non posso creare i viaggi:\n\n" + problemi.slice(0, 12).join("\n") +
             (problemi.length > 12 ? "\n... e altri " + (problemi.length - 12) : "") +
             "\n\nCorreggi le celle colorate e riprova.");
    return;
  }

  const conferma = ui.alert("Creo " + pronti.length + " viaggi per il " +
    Utilities.formatDate(data, s.getSpreadsheetTimeZone(), "dd/MM/yyyy") + ".\n\nProcedo?",
    ui.ButtonSet.YES_NO);
  if (conferma !== ui.Button.YES) return;

  // partenza: quella della regola, altrimenti l'ultimo arrivo di quel mezzo
  const reg = s.getSheetByName("REGISTRO_VIAGGI");
  const ultimeDest = {};
  const ur = ultimaRigaViaggi_(reg);
  if (ur >= 2) {
    const dati = reg.getRange(2, 5, ur - 1, 6).getValues();   // E..J
    dati.forEach(rr => {
      const targa = String(rr[0] || "").trim();
      const arrivo = String(rr[5] || "").trim();
      if (targa && arrivo) ultimeDest[targa] = arrivo;
    });
  }

  const nuove = [];
  pronti.forEach(p => {
    const partenza = p.reg.partenza || ultimeDest[p.mezzo.targa] || "";
    nuove.push([
      Utilities.getUuid(), data, p.autista, "",
      p.mezzo.targa, "", "", "", partenza, p.reg.arrivo,
      p.rim ? p.rim.targa : "", "", "", "", "", "Assegnato", "",
      p.nota, p.reg.chiave, "", "", ""
    ]);
    ultimeDest[p.mezzo.targa] = p.reg.arrivo;
  });

  reg.getRange(ultimaRigaViaggi_(reg) + 1, 1, nuove.length, 22).setValues(nuove);

  const imparati = imparaSinonimi_(pronti, ana);

  SpreadsheetApp.flush();
  ui.alert("FATTO!\n\n" + nuove.length + " viaggi creati nel registro." +
    (imparati ? ("\n" + imparati + " nuovi sinonimi imparati: la prossima volta li riconoscero' da solo.") : "") +
    "\n\nGli autisti li vedono subito nell'app.");
}

// Se hai corretto una cella, memorizzo la corrispondenza per la volta dopo
function imparaSinonimi_(pronti, ana) {
  const sh = ss_().getSheetByName(FOGLIO_SINONIMI);
  if (!sh) return 0;

  const gia = {};
  const n = Math.max(sh.getLastRow() - 1, 0);
  if (n > 0) {
    sh.getRange(2, 1, n, 2).getValues().forEach(r => {
      gia[String(r[0]).toUpperCase() + "|" + norm_(r[1])] = true;
    });
  }

  const nuovi = [];
  const agg = (tipo, scritto, vale) => {
    const s = String(scritto || "").trim();
    if (!s || !vale) return;
    if (norm_(s) === norm_(vale)) return;
    const k = tipo + "|" + norm_(s);
    if (gia[k]) return;
    gia[k] = true;
    nuovi.push([tipo, s, vale]);
  };

  pronti.forEach(p => {
    agg("AUTISTA", p.raw[6], p.autista);
    agg("CAMION", p.raw[7], p.mezzo.modello);
    if (p.rim) agg("RIMORCHIO", p.raw[8], p.rim.nome);
    agg("TRATTA", p.raw[9], p.reg.chiave);
  });

  if (nuovi.length) sh.getRange(sh.getLastRow() + 1, 1, nuovi.length, 3).setValues(nuovi);
  return nuovi.length;
}

// ------------------------------------------------------------
function SVUOTA() {
  const sh = ss_().getSheetByName(FOGLIO_LAVORO);
  if (!sh) return;
  const ultima = Math.max(sh.getLastRow(), 100);
  sh.getRange(RIGA_INIZIO_TESTO, COL_TESTO, ultima, 1).clearContent();
  sh.getRange(RIGA_INIZIO_GRIGLIA, COL_AUT, ultima, 10)
    .clearContent().clearDataValidations().setBackground(BIANCO)
    .setFontColor(null).setFontStyle("normal").setFontWeight("normal");
  SpreadsheetApp.getUi().alert("Svuotato. Puoi incollare un nuovo messaggio.");
}

// ------------------------------------------------------------
//  Creazione delle schede (da eseguire una volta sola)
// ------------------------------------------------------------
function PREPARA() {
  const s = ss_();

  let sin = s.getSheetByName(FOGLIO_SINONIMI);
  if (!sin) {
    sin = s.insertSheet(FOGLIO_SINONIMI);
    sin.getRange(1, 1, 1, 3).setValues([["Tipo", "Come lo scrive il capo", "Corrisponde a"]])
      .setBackground("#0b192c").setFontColor("#ffffff").setFontWeight("bold");
    sin.setColumnWidth(1, 110); sin.setColumnWidth(2, 260); sin.setColumnWidth(3, 260);
    sin.setFrozenRows(1);

    const iniziali = [
      ["TRATTA", "scarico pettoranello", "Res Pettoranello"],
      ["TRATTA", "carico pettoranello", "Res Pettoranello"],
      ["TRATTA", "gma plast x deco", "GMA Plast"],
      ["TRATTA", "sri x deco", "S.R.I."],
      ["TRATTA", "poliservice x deco", "Poloservice x Deco"],
      ["TRATTA", "carico edison x teramo", "Edison x Teramo"],
      ["TRATTA", "ecotec x lanciano carico", "Ecotec x Lanciano"],
      ["CAMION", "mercedes 630 (3)", "Merceds 630 (3)"],
      ["CAMION", "iveco 480 (2)", "Iveco 580(2)"]
    ];
    sin.getRange(2, 1, iniziali.length, 3).setValues(iniziali);
  }

  let sh = s.getSheetByName(FOGLIO_LAVORO);
  if (!sh) sh = s.insertSheet(FOGLIO_LAVORO);
  sh.clear();

  sh.getRange(1, 1).setValue("1) INCOLLA QUI SOTTO IL MESSAGGIO DEL CAPO")
    .setFontWeight("bold").setFontSize(12).setFontColor("#0b192c");
  sh.getRange(2, 1).setValue("(dalla riga 4 in giu', una riga per riga)")
    .setFontStyle("italic").setFontColor("#64748b");

  sh.getRange(1, 3).setValue("2) CONTROLLA E CORREGGI")
    .setFontWeight("bold").setFontSize(12).setFontColor("#0b192c");
  sh.getRange(2, 3).setValue("Data viaggi:").setFontWeight("bold");
  sh.getRange(2, 4).setValue(new Date()).setNumberFormat("dd/mm/yyyy")
    .setBackground("#fff7ed").setBorder(true, true, true, true, false, false);

  const intest = ["Autista", "Camion", "Rimorchio", "Tratta", "Nota", "Esito",
                  "(orig. autista)", "(orig. camion)", "(orig. rim.)", "(orig. tratta)"];
  sh.getRange(3, COL_AUT, 1, intest.length).setValues([intest])
    .setBackground("#0b192c").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.setColumnWidth(1, 380);
  sh.setColumnWidth(COL_AUT, 200); sh.setColumnWidth(COL_CAM, 170);
  sh.setColumnWidth(COL_RIM, 100); sh.setColumnWidth(COL_TRA, 200);
  sh.setColumnWidth(COL_NOTA, 240); sh.setColumnWidth(COL_ESITO, 190);
  for (let c = COL_RAW; c < COL_RAW + 4; c++) sh.setColumnWidth(c, 130);
  sh.hideColumns(COL_RAW, 4);
  sh.setFrozenRows(3);
  sh.setTabColor("#ff7800");

  s.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    "Schede pronte!\n\n" +
    "1. Incolla il messaggio nella colonna A dalla riga 4\n" +
    "2. Menu Priamus > 1. Leggi il messaggio\n" +
    "3. Correggi le righe colorate\n" +
    "4. Menu Priamus > 2. Crea i viaggi"
  );
}