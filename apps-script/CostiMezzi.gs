// ============================================================
//  PRIAMUS - COSTI DEI MEZZI
//  (tutti i nomi iniziano con CM_ / cm per non scontrarsi
//   con le costanti degli altri file del progetto)
//
//  Legge i file Excel dei costi che l'ufficio mette in una
//  cartella di Drive e li trascrive nel foglio COSTI_MEZZI.
//
//  DOVE VA: progetto Apps Script del DATABASE
//  (+ > Script, nome: CostiMezzi)
//
//  PRIMA DI USARLO servono due cose:
//   1) nell'editor, colonna di sinistra: Servizi (+) > "Drive API" > Aggiungi
//   2) su Drive, una cartella con il nome scritto qui sotto
// ============================================================

const CM_CARTELLA = "Costi mezzi";       // cartella su Drive dove arrivano i file
const CM_FOGLIO = "COSTI_MEZZI";

// Le colonne NON sono sempre nella stessa posizione:
// in alcuni mesi manca "Avere" e tutto si sposta.
// Lo script le riconosce da solo, vedi cmTrovaColonne_().

// Voci scritte in modi diversi ma che sono la stessa cosa
const CM_VOCI_UGUALI = {
  "CONTRATTO MANUTENZIONE": "CONTR. MANUTENZIONE",
  "CONTRATTO DI MANUTENZIONE": "CONTR. MANUTENZIONE",
  "RIPARAZIONI - MANUTENZIONI": "RIPARAZIONI",
  "RIPARAZIONI E RICAMBI": "RIPARAZIONI",
  "RICAMBI - RIPARAZIONI": "RIPARAZIONI",
  "RIPARAZIONE": "RIPARAZIONI",
  "RIPAZIONI": "RIPARAZIONI",          // refuso ricorrente nei file
  "ASSICURAZIONI": "ASSICURAZIONE",
  "-KASKO": "KASKO",
  "MUTUO SELLA": "MUTUO",
  "ANTICIPO LEASING": "ANTICIPO",
  "MUTUO": "MUTUO"
};

// Targhe scritte male nei file della contabilita' -> targa del database
// (cosi' i costi finiscono sulla scheda del mezzo giusto)
const CM_TARGHE_UGUALI = {
  "HA879PG": "HA878PG",   // Volvo 540
  "HE973NF": "HE972FN",   // Man 520 (2): targa vera HE973FN, nel database HE972FN
  "HE973FN": "HE972FN",
  "XA155FT": "XA115FT",   // rimorchio V.10
  "XA452ZD": "XA452ZA",   // rimorchio V.5
  "HF943DL": "GE098XW"    // Mercedes 530 (1): nuova targa dopo lo smarrimento
};

const CM_MESI = {
  "GENNAIO": 1, "FEBBRAIO": 2, "MARZO": 3, "APRILE": 4, "MAGGIO": 5, "GIUGNO": 6,
  "LUGLIO": 7, "AGOSTO": 8, "SETTEMBRE": 9, "OTTOBRE": 10, "NOVEMBRE": 11, "DICEMBRE": 12
};

// ------------------------------------------------------------
//  NOTA: qui NON deve esistere una funzione onOpen().
//  Il menu si costruisce tutto dentro AssegnaRapida.gs:
//  in Apps Script puo' esserci un solo onOpen in tutto il progetto.
// ------------------------------------------------------------
function menuCosti_() {
  SpreadsheetApp.getUi()
    .createMenu("💶 Costi")
    .addItem("Importa i file nuovi", "IMPORTA_COSTI")
    .addItem("Prepara il foglio (prima volta)", "PREPARA_COSTI")
    .addSeparator()
    .addItem("Attiva importazione automatica", "ATTIVA_AUTOMATISMO")
    .addItem("Disattiva importazione automatica", "DISATTIVA_AUTOMATISMO")
    .addToUi();
}

// ------------------------------------------------------------
function PREPARA_COSTI() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  let sh = s.getSheetByName(CM_FOGLIO);
  if (!sh) sh = s.insertSheet(CM_FOGLIO);

  if (sh.getLastRow() === 0) {
    const intest = ["Mese", "Targa", "Tipo", "Voce", "Importo",
                    "Voce originale", "Annotazione", "File di origine"];
    sh.getRange(1, 1, 1, intest.length).setValues([intest])
      .setBackground("#0b192c").setFontColor("#FFFFFF").setFontWeight("bold")
      .setHorizontalAlignment("center");
    sh.setColumnWidth(1, 90);  sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 105);
    sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 190);
    sh.setColumnWidth(7, 300); sh.setColumnWidth(8, 220);
    sh.setFrozenRows(1);
    sh.setTabColor("#0b192c");
  }

  // la cartella su Drive
  const c = DriveApp.getFoldersByName(CM_CARTELLA);
  const cartella = c.hasNext() ? c.next() : DriveApp.createFolder(CM_CARTELLA);

  SpreadsheetApp.getUi().alert(
    "Pronto.\n\n" +
    "Foglio: " + CM_FOGLIO + "\n" +
    "Cartella su Drive: " + cartella.getName() + "\n\n" +
    "Metti li' dentro i file dei costi, poi:\n" +
    "menu Costi > Importa i file nuovi.\n\n" +
    "NOME DEI FILE: meglio se contengono anno e mese,\n" +
    "per esempio  costi_2026-07.xlsx"
  );
}

// ------------------------------------------------------------
//  Importazione
// ------------------------------------------------------------
function IMPORTA_COSTI() {
  const esito = cmImportaTutto_();
  SpreadsheetApp.getUi().alert(esito);
}

// Versione silenziosa, usata dall'automatismo
function IMPORTA_COSTI_AUTO() {
  cmImportaTutto_();
}

function cmImportaTutto_() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const sh = s.getSheetByName(CM_FOGLIO);
  if (!sh) return "Foglio " + CM_FOGLIO + " non trovato: esegui prima 'Prepara il foglio'.";

  const c = DriveApp.getFoldersByName(CM_CARTELLA);
  if (!c.hasNext()) return "Cartella '" + CM_CARTELLA + "' non trovata su Drive.";
  const cartella = c.next();

  // file gia' importati
  const gia = {};
  const ultima = sh.getLastRow();
  if (ultima > 1) {
    sh.getRange(2, 8, ultima - 1, 1).getValues().forEach(r => {
      const n = String(r[0] || "").trim();
      if (n) gia[n] = true;
    });
  }

  const tipi = cmTipiVeicolo_();
  const nuove = [];
  const fatti = [], saltati = [], problemi = [];

  const file = cartella.getFiles();
  while (file.hasNext()) {
    const f = file.next();
    const nome = f.getName();

    if (gia[nome]) { saltati.push(nome); continue; }
    if (!/\.(xlsx|xls|csv)$/i.test(nome) && f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    let righe;
    try {
      righe = cmLeggiFile_(f);
    } catch (e) {
      problemi.push(nome + ": " + e.message);
      continue;
    }

    const mese = cmMeseDalNome_(nome);
    const col = cmTrovaColonne_(righe);
    if (col.etichetta < 0) { problemi.push(nome + ": non trovo la colonna con le targhe"); continue; }
    let contate = 0;

    righe.forEach(r => {
      const etichetta = String(r[col.etichetta] || "").trim();
      const m = etichetta.match(/^([A-Z]{2}\d{3}[A-Z]{2})\s+(.*)$/);
      if (!m) return;                       // salta intestazioni e riga di totale

      const importo = cmNumero_(r[col.importo]);
      if (!importo) return;

      const targa = CM_TARGHE_UGUALI[m[1]] || m[1];
      const voceOrig = m[2].trim().toUpperCase();
      const voce = CM_VOCI_UGUALI[voceOrig] || voceOrig;

      nuove.push([
        mese, targa, tipi[targa] || "SCONOSCIUTO", voce, importo,
        voceOrig, (col.nota >= 0 ? String(r[col.nota] || "").trim() : ""), nome
      ]);
      contate++;
    });

    if (contate) fatti.push(nome + " (" + mese + "): " + contate + " voci");
    else problemi.push(nome + ": nessuna riga riconosciuta");
  }

  if (nuove.length) {
    const daRiga = sh.getLastRow() + 1;
    // La colonna Mese DEVE restare testo: scrivendo "2026-08" senza questo,
    // Google lo converte in data e i confronti nelle schede non funzionano piu'.
    sh.getRange(daRiga, 1, nuove.length, 1).setNumberFormat("@");
    sh.getRange(daRiga, 1, nuove.length, 8).setValues(nuove);
    sh.getRange(2, 5, sh.getLastRow() - 1, 1).setNumberFormat('#,##0.00 "€"');
    SpreadsheetApp.flush();
  }

  let testo = "IMPORTAZIONE COSTI\n\n";
  testo += fatti.length ? ("Importati:\n- " + fatti.join("\n- ") + "\n\n")
                        : "Nessun file nuovo da importare.\n\n";
  testo += "Righe aggiunte: " + nuove.length + "\n";
  if (saltati.length) testo += "Gia' importati (saltati): " + saltati.length + "\n";
  if (problemi.length) testo += "\nATTENZIONE:\n- " + problemi.join("\n- ");

  const sconosciute = {};
  nuove.forEach(r => { if (r[2] === "SCONOSCIUTO") sconosciute[r[1]] = true; });
  const elenco = Object.keys(sconosciute);
  if (elenco.length) {
    testo += "\n\nTarghe non presenti in anagrafica:\n" + elenco.join(", ") +
             "\n(le voci sono state importate lo stesso)";
  }
  return testo;
}

// ------------------------------------------------------------
//  Riconoscimento delle colonne
//  In alcuni mesi manca la colonna "Avere" e tutto slitta:
//  invece di fidarmi della posizione, le cerco.
// ------------------------------------------------------------
function cmTrovaColonne_(righe) {
  const col = { importo: 1, nota: -1, etichetta: -1 };
  if (!righe.length) return col;

  // dalle intestazioni: "Dare" e "Annotazioni"
  const testa = righe[0].map(x => String(x || "").trim().toLowerCase());
  const iDare = testa.indexOf("dare");
  if (iDare >= 0) col.importo = iDare;
  const iNota = testa.indexOf("annotazioni");
  if (iNota >= 0) col.nota = iNota;

  // la colonna delle targhe: quella dove compaiono piu' volte
  const quante = [];
  for (let c = 0; c < 15; c++) quante[c] = 0;
  for (let r = 1; r < Math.min(righe.length, 40); r++) {
    for (let c = 0; c < Math.min(righe[r].length, 15); c++) {
      if (/^[A-Z]{2}\d{3}[A-Z]{2}\s+/.test(String(righe[r][c] || "").trim())) quante[c]++;
    }
  }
  let best = -1, max = 0;
  quante.forEach((n, c) => { if (n > max) { max = n; best = c; } });
  col.etichetta = max > 0 ? best : -1;

  return col;
}

// Gli importi arrivano come numero ("40.98") o come testo all'italiana ("1.450,00")
function cmNumero_(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;

  let t = String(v).trim().replace(/[€\s]/g, "");
  if (!t) return 0;

  if (t.indexOf(",") >= 0) {
    t = t.replace(/\./g, "").replace(",", ".");   // 1.450,00 -> 1450.00
  }
  const n = Number(t);
  return isNaN(n) ? 0 : n;
}

// ------------------------------------------------------------
//  Lettura del file: xlsx e xls vengono convertiti al volo
// ------------------------------------------------------------
function cmLeggiFile_(f) {
  const tipo = f.getMimeType();

  if (tipo === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(f.getId()).getSheets()[0].getDataRange().getValues();
  }

  if (tipo === MimeType.CSV || /\.csv$/i.test(f.getName())) {
    return Utilities.parseCsv(f.getBlob().getDataAsString());
  }

  // Excel: ne faccio una copia in formato Google, la leggo e la cestino
  let idCopia;
  try {
    idCopia = Drive.Files.copy({ name: "tmp-costi", mimeType: MimeType.GOOGLE_SHEETS }, f.getId()).id;
  } catch (e1) {
    try {
      idCopia = Drive.Files.copy({ title: "tmp-costi", mimeType: MimeType.GOOGLE_SHEETS }, f.getId()).id;
    } catch (e2) {
      throw new Error("non riesco a convertire il file. Hai aggiunto il servizio 'Drive API'? (" + e2.message + ")");
    }
  }

  try {
    return SpreadsheetApp.openById(idCopia).getSheets()[0].getDataRange().getValues();
  } finally {
    try { DriveApp.getFileById(idCopia).setTrashed(true); } catch (e) {}
  }
}

// "costi_2026-07.xlsx" -> "2026-07"   |   "costi_camion_LUGLIO.xlsx" -> anno corrente
function cmMeseDalNome_(nome) {
  const m = nome.match(/(20\d{2})[-_\.]?(0[1-9]|1[0-2])/);
  if (m) return m[1] + "-" + m[2];

  const su = nome.toUpperCase();
  for (const k in CM_MESI) {
    if (su.indexOf(k) >= 0) {
      const anno = (su.match(/20\d{2}/) || [new Date().getFullYear()])[0];
      return anno + "-" + ("0" + CM_MESI[k]).slice(-2);
    }
  }
  const oggi = new Date();
  return oggi.getFullYear() + "-" + ("0" + (oggi.getMonth() + 1)).slice(-2);
}

// targa -> CAMION oppure RIMORCHIO
function cmTipiVeicolo_() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const tipi = {};

  const shM = s.getSheetByName("ANAGRAFICA_MEZZI");
  if (shM && shM.getLastRow() > 1) {
    shM.getRange(2, 2, shM.getLastRow() - 1, 1).getValues().forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) tipi[t] = "CAMION";
    });
  }

  const shR = s.getSheetByName("ANAGRAFICA_RIMORCHI");
  if (shR && shR.getLastRow() > 1) {
    shR.getRange(2, 1, shR.getLastRow() - 1, 1).getValues().forEach(r => {
      const t = String(r[0] || "").trim().toUpperCase();
      if (t) tipi[t] = "RIMORCHIO";
    });
  }
  return tipi;
}

// ------------------------------------------------------------
//  Automatismo: controlla la cartella ogni giorno
// ------------------------------------------------------------
function ATTIVA_AUTOMATISMO() {
  DISATTIVA_AUTOMATISMO();
  ScriptApp.newTrigger("IMPORTA_COSTI_AUTO")
    .timeBased().atHour(7).everyDays(1).create();
  SpreadsheetApp.getUi().alert(
    "Fatto.\n\nOgni mattina verso le 7 il programma guarda nella cartella\n'" +
    CM_CARTELLA + "' e importa i file che non ha ancora letto.\n\n" +
    "Puoi comunque importare a mano quando vuoi."
  );
}

function DISATTIVA_AUTOMATISMO() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "IMPORTA_COSTI_AUTO") ScriptApp.deleteTrigger(t);
  });
}


// ------------------------------------------------------------
//  SISTEMA I CM_MESI GIA' SCRITTI
//  Google aveva trasformato "2026-08" in una data: qui torna
//  a essere testo, cosi' le schede dei mezzi lo riconoscono.
//  Da eseguire una volta. Riesegubile senza danni.
// ------------------------------------------------------------
function SISTEMA_MESI() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CM_FOGLIO);
  const ui = SpreadsheetApp.getUi();
  if (!sh) { ui.alert("Foglio " + CM_FOGLIO + " non trovato."); return; }

  const ultima = sh.getLastRow();
  if (ultima < 2) { ui.alert("Nessun dato da sistemare."); return; }

  const rng = sh.getRange(2, 1, ultima - 1, 1);
  const valori = rng.getValues();
  const fuso = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  let cambiati = 0;
  const nuovi = valori.map(r => {
    const v = r[0];
    if (v instanceof Date) {
      cambiati++;
      return [Utilities.formatDate(v, fuso, "yyyy-MM")];
    }
    const t = String(v || "").trim();
    return [t];
  });

  rng.setNumberFormat("@");     // prima il formato, poi i valori
  rng.setValues(nuovi);
  SpreadsheetApp.flush();

  // conta i mesi risultanti, per controllo
  const conteggio = {};
  nuovi.forEach(r => { if (r[0]) conteggio[r[0]] = (conteggio[r[0]] || 0) + 1; });
  const elenco = Object.keys(conteggio).sort()
    .map(m => "   " + m + ": " + conteggio[m] + " righe").join("\n");

  ui.alert(
    "CM_MESI SISTEMATI\n\n" +
    "Celle convertite da data a testo: " + cambiati + "\n\n" +
    "Mesi presenti ora:\n" + elenco + "\n\n" +
    "Riapri una scheda mezzo: le spese devono comparire."
  );
}