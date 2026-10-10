// ============================================================
//  GESTIONE FLOTTA - PRIAMUS  |  Codice.gs
//  Aggiornato: i menu mostrano MODELLO / NOME RIMORCHIO,
//  ma sui fogli continua a essere scritta la TARGA.
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gestione Flotta - Priamus')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ------------------------------------------------------------
//  NUOVO: dizionari targa -> etichetta da mostrare
//  ANAGRAFICA_MEZZI    -> col B (indice 1) = Targa, col D (indice 3) = Modello
//  ANAGRAFICA_RIMORCHI -> col A (indice 0) = Targa, col B (indice 1) = Nome
// ------------------------------------------------------------
function mappeFlotta_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mezzi = {};
  const rimorchi = {};

  const fMezzi = ss.getSheetByName("ANAGRAFICA_MEZZI").getDataRange().getValues();
  for (let i = 1; i < fMezzi.length; i++) {
    const targa = String(fMezzi[i][1]).trim();
    const modello = String(fMezzi[i][3]).trim();
    if (targa !== "") mezzi[targa] = (modello !== "" ? modello : targa);
  }

  const fRim = ss.getSheetByName("ANAGRAFICA_RIMORCHI").getDataRange().getValues();
  for (let i = 1; i < fRim.length; i++) {
    const targa = String(fRim[i][0]).trim();
    const nome = String(fRim[i][1]).trim();
    if (targa !== "") rimorchi[targa] = (nome !== "" ? nome : targa);
  }

  return { mezzi: mezzi, rimorchi: rimorchi };
}

// Traduce una targa nell'etichetta; se non la trova, restituisce la targa.
function etichetta_(mappa, targa) {
  const t = String(targa).trim();
  if (t === "") return "";
  return mappa[t] || t;
}

function getElenchiIniziali() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mappe = mappeFlotta_();

    // 1. Elenchi base
    const fAutisti = ss.getSheetByName("AUTISTI").getDataRange().getValues();
    const autisti = fAutisti.slice(1).map(r => String(r[0]).trim()).filter(String);

    // NUOVO: liste di oggetti {targa, label} invece di semplici targhe
    const mezzi = Object.keys(mappe.mezzi).map(t => ({ targa: t, label: mappe.mezzi[t] }));
    const rimorchi = Object.keys(mappe.rimorchi).map(t => ({ targa: t, label: mappe.rimorchi[t] }));

    // 2. Regole viaggi
    let regole = [];
    const foglioRegole = ss.getSheetByName("REGOLE_VIAGGI");
    if (foglioRegole) {
      const fRegole = foglioRegole.getDataRange().getValues();
      for (let i = 1; i < fRegole.length; i++) {
        let chiave = String(fRegole[i][0]).trim();
        if (chiave !== "") {
          regole.push({
            chiave: chiave,
            partenza: String(fRegole[i][1]).trim(),
            arrivo: String(fRegole[i][2]).trim()
          });
        }
      }
    }

    // 3. Ultima destinazione di ogni camion (chiave = TARGA, non cambia)
    let ultimeDestinazioni = {};
    const foglioViaggi = ss.getSheetByName("REGISTRO_VIAGGI");
    if (foglioViaggi) {
      const datiViaggi = foglioViaggi.getDataRange().getValues();
      for (let i = datiViaggi.length - 1; i > 0; i--) {
        let targa = String(datiViaggi[i][4]).trim();
        let arrivo = String(datiViaggi[i][9]).trim();
        if (targa !== "" && arrivo !== "" && !ultimeDestinazioni[targa]) {
          ultimeDestinazioni[targa] = arrivo;
        }
      }
    }

    return {
      autisti: autisti,
      mezzi: mezzi,
      rimorchi: rimorchi,
      regole: regole,
      ultimeDestinazioni: ultimeDestinazioni
    };
  } catch (e) {
    throw new Error("Errore nel caricamento elenchi: " + e.message);
  }
}

function verificaAccesso(tipo, utente, password) {
  if (tipo === "responsabile") {
    // Il valore NON sta nel codice (il repository GitHub è pubblico):
    // Impostazioni progetto > Proprietà script > PASSWORD_ADMIN
    const PASSWORD_ADMIN = PropertiesService.getScriptProperties().getProperty('PASSWORD_ADMIN');
    if (PASSWORD_ADMIN && password === PASSWORD_ADMIN) return { success: true, role: "admin" };
    return { success: false, msg: "Password Responsabile errata!" };
  }
  if (tipo === "autista") {
    const fAutisti = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AUTISTI").getDataRange().getValues();
    for (let i = 1; i < fAutisti.length; i++) {
      let nomeCorrente = String(fAutisti[i][0]).trim();
      let pinCorrente = String(fAutisti[i][2]).trim();
      if (nomeCorrente === utente) {
        if (pinCorrente === password && pinCorrente !== "") return { success: true, role: "autista", nome: nomeCorrente };
        return { success: false, msg: "PIN errato!" };
      }
    }
    return { success: false, msg: "Autista non trovato nel database." };
  }
  return { success: false, msg: "Errore di sistema." };
}

function getViaggiAttiviAutista(nomeAutista) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const foglio = ss.getSheetByName("REGISTRO_VIAGGI");
    const dati = foglio.getDataRange().getValues();
    const mappe = mappeFlotta_();   // NUOVO
    const viaggi = [];
    const nomeCercato = String(nomeAutista).trim().toLowerCase();

    // NUOVO: mappa Luogo Arrivo -> Viaggio_Chiave (dalle REGOLE_VIAGGI).
    // Cosi' l'autista legge la chiave (es. "Ostia") invece di partenza/arrivo.
    let chiavePerArrivo = {};
    const foglioRegole = ss.getSheetByName("REGOLE_VIAGGI");
    if (foglioRegole) {
      const fRegole = foglioRegole.getDataRange().getValues();
      for (let i = 1; i < fRegole.length; i++) {
        const chiave = String(fRegole[i][0]).trim();
        const arrivoRegola = String(fRegole[i][2]).trim();
        if (chiave !== "" && arrivoRegola !== "" && !chiavePerArrivo[arrivoRegola]) {
          chiavePerArrivo[arrivoRegola] = chiave;
        }
      }
    }

    let ultimiKmCamion = {};
    for (let i = dati.length - 1; i > 0; i--) {
      let targaCorrente = String(dati[i][4]).trim();
      let kmArrivo = String(dati[i][6]).trim();
      if (targaCorrente !== "" && kmArrivo !== "" && !ultimiKmCamion[targaCorrente]) {
        ultimiKmCamion[targaCorrente] = kmArrivo;
      }
    }

    for (let i = 1; i < dati.length; i++) {
      const autistaRiga = String(dati[i][2]).trim().toLowerCase();
      const statoRiga = String(dati[i][15]).trim();
      if (autistaRiga === nomeCercato && statoRiga !== "Completato") {
        let targaMezzo = String(dati[i][4]).trim();
        let targaRimorchio = String(dati[i][10]).trim();
        let kmSuggerito = ultimiKmCamion[targaMezzo] || "";
        let dataFormattata = "";
        if (dati[i][1]) {
          dataFormattata = dati[i][1] instanceof Date
            ? Utilities.formatDate(dati[i][1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm")
            : String(dati[i][1]);
        }
        viaggi.push({
          id: String(dati[i][0]),
          data: dataFormattata,
          mezzo: etichetta_(mappe.mezzi, targaMezzo),          // NUOVO: modello
          rimorchio: etichetta_(mappe.rimorchi, targaRimorchio), // NUOVO: nome
          targaMezzo: targaMezzo,                               // serve al rifornimento
          partenza: String(dati[i][8]),
          arrivo: String(dati[i][9]),
          chiave: String(dati[i][18] || "") || chiavePerArrivo[String(dati[i][9]).trim()] || String(dati[i][9]), // NUOVO: prima la chiave salvata (col S), poi la traduzione
          stato: statoRiga || "Assegnato",
          kmPrecedente: kmSuggerito,
          nota: String(dati[i][17] || "")   // NUOVO: nota dell'admin (colonna R)
        });
      }
    }
    return viaggi;
  } catch (e) {
    throw new Error("Errore nel server: " + e.message);
  }
}

function assegnaNuovoViaggio(dati) {
  // dati.targaMezzo e dati.targaRimorchio arrivano gia' come TARGHE: nessuna modifica.
  // NUOVO: la colonna B (Data) contiene la DATA DEL VIAGGIO scelta dall'admin,
  // non piu' il momento dell'assegnazione. E' il giorno in cui il viaggio
  // compare nel Programma Giornaliero. Ore 12:00 per evitare problemi di fuso.
  const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTRO_VIAGGI");

  let dataViaggio = new Date();
  if (dati.dataViaggio) {
    const p = String(dati.dataViaggio).split("-"); // formato yyyy-mm-dd
    if (p.length === 3) {
      dataViaggio = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
    }
  }

  foglio.appendRow([
    Utilities.getUuid(), dataViaggio, dati.autista, "",
    dati.targaMezzo, "", "", "", dati.partenza, dati.arrivo,
    dati.targaRimorchio, "", "", "", "", "Assegnato", "",
    dati.nota || "", dati.chiave || ""   // R = Nota, S = Viaggio_Chiave
  ]);
  return "Viaggio assegnato correttamente a " + dati.autista;
}

function avviaViaggioAutista(idViaggio, kmPartenza) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foglio = ss.getSheetByName("REGISTRO_VIAGGI");
  const dati = foglio.getDataRange().getValues();
  for (let i = 1; i < dati.length; i++) {
    if (String(dati[i][0]) === String(idViaggio)) {
      // NON tocchiamo la colonna B (Data): contiene la data del
      // viaggio pianificata dall'admin e sposterebbe il viaggio nel Programma.
      foglio.getRange(i + 1, 6).setValue(Number(kmPartenza));
      foglio.getRange(i + 1, 16).setValue("In Corso");
      // NUOVO: colonna T (20) = Ora_Inizio, momento in cui l'autista parte
      foglio.getRange(i + 1, 20).setValue(
        Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm")
      );
      return "Viaggio avviato con successo!";
    }
  }
  throw new Error("Viaggio non trovato.");
}

function terminaViaggioAutista(dati) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const foglio = ss.getSheetByName("REGISTRO_VIAGGI");
    const righeDati = foglio.getDataRange().getValues();
    const oraAttuale = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "HH:mm");

    for (let i = 1; i < righeDati.length; i++) {
      if (String(righeDati[i][0]) === String(dati.idViaggio)) {
        const riga = i + 1;
        const kmPartenza = Number(righeDati[i][5]) || 0;
        const kmArrivo = dati.kmArrivo !== "" ? Number(dati.kmArrivo) : "";
        const kmViaggio = (kmArrivo !== "" && kmPartenza !== 0) ? (kmArrivo - kmPartenza) : "";
        const tonnellate = dati.tonnellate !== "" ? Number(dati.tonnellate) : "";
        const oreLavoro = dati.oreLavoro || "";

        let fotoUrl = "";
        if (dati.fotoData && dati.fotoNome) {
          try {
            const decodificato = Utilities.base64Decode(dati.fotoData.split(",")[1] || dati.fotoData);
            const blob = Utilities.newBlob(decodificato, dati.fotoMime || 'image/jpeg', dati.fotoNome);
            const file = DriveApp.getRootFolder().createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
            fotoUrl = file.getUrl();
          } catch (e) { fotoUrl = "Errore salvataggio foto"; }
        }

        foglio.getRange(riga, 7).setValue(kmArrivo);
        foglio.getRange(riga, 8).setValue(kmViaggio);
        foglio.getRange(riga, 12).setValue(tonnellate);
        foglio.getRange(riga, 13).setValue(oreLavoro);
        if (fotoUrl !== "") foglio.getRange(riga, 14).setValue(fotoUrl);
        foglio.getRange(riga, 16).setValue("Completato");
        foglio.getRange(riga, 17).setValue(oraAttuale);
        return "Viaggio chiuso correttamente!";
      }
    }
    throw new Error("Viaggio non trovato.");
  } catch (err) {
    throw new Error(err.message);
  }
}

function aggiungiRifornimento(dati) {
  // dati.targaMezzo arriva gia' come TARGA: sul foglio finisce la targa.
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTRO_RIFORNIMENTI");
    const idRif = Utilities.getUuid();
    const dataRif = new Date();
    let fotoUrl = "";

    if (dati.fotoData && dati.fotoNome) {
      try {
        const decodificato = Utilities.base64Decode(dati.fotoData.split(",")[1] || dati.fotoData);
        const blob = Utilities.newBlob(decodificato, dati.fotoMime || 'image/jpeg', dati.fotoNome);
        const file = DriveApp.getRootFolder().createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
        fotoUrl = file.getUrl();
      } catch (e) { fotoUrl = "Errore salvataggio foto"; }
    }

    foglio.appendRow([idRif, dataRif, dati.targaMezzo, Number(dati.litri), Number(dati.costo), Number(dati.km), fotoUrl]);
    return "Rifornimento salvato con successo nel registro!";
  } catch (err) {
    throw new Error("Errore nel salvataggio rifornimento: " + err.message);
  }
}

function getStatoFlotta() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const foglio = ss.getSheetByName("REGISTRO_VIAGGI");
    const dati = foglio.getDataRange().getValues();
    const mappe = mappeFlotta_();   // NUOVO
    const flotta = [];

    for (let i = dati.length - 1; i > 0; i--) {
      const stato = String(dati[i][15]).trim();
      if (stato === "Assegnato" || stato === "In Corso") {
        let dataFormattata = "";
        if (dati[i][1]) {
          dataFormattata = dati[i][1] instanceof Date
            ? Utilities.formatDate(dati[i][1], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy")
            : String(dati[i][1]);
        }
        flotta.push({
          id: String(dati[i][0]),                              // NUOVO: serve al pulsante Elimina
          autista: String(dati[i][2]),
          mezzo: etichetta_(mappe.mezzi, dati[i][4]),          // NUOVO
          rimorchio: etichetta_(mappe.rimorchi, dati[i][10]),  // NUOVO
          partenza: String(dati[i][8]),
          arrivo: String(dati[i][9]),
          stato: stato,
          dataAggiornamento: dataFormattata
        });
      }
    }
    return flotta;
  } catch (e) {
    throw new Error("Errore nel recupero flotta: " + e.message);
  }
}

// NUOVO: elimina un viaggio assegnato per errore (solo Assegnato o In Corso).
// I viaggi Completati non si toccano: sono storico.
function eliminaViaggio(idViaggio) {
  const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTRO_VIAGGI");
  const dati = foglio.getDataRange().getValues();
  for (let i = 1; i < dati.length; i++) {
    if (String(dati[i][0]) === String(idViaggio)) {
      const stato = String(dati[i][15]).trim();
      if (stato === "Completato") {
        throw new Error("Questo viaggio e' Completato: non si puo' eliminare dallo storico.");
      }
      const autista = String(dati[i][2]);
      foglio.deleteRow(i + 1);
      return "Viaggio di " + autista + " eliminato correttamente.";
    }
  }
  throw new Error("Viaggio non trovato: forse e' gia' stato eliminato.");
}