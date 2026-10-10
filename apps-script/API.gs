// ============================================================
//  PRIAMUS - API.gs  (SERVIZIO FOTO)  v2 - due cartelle separate
//
//  Le foto dei VIAGGI vanno in una cartella,
//  le RICEVUTE dei rifornimenti in un'altra.
//  Dentro ciascuna, una sottocartella per mese.
//
//  Lo script gira a NOME TUO e usa il TUO spazio Drive.
//
//  DOVE VA: sostituisce il contenuto del file API.gs nel progetto
//  Apps Script del database.
//  DOPO: Esegui il deployment > Gestisci deployment > API PWA >
//        Nuova versione. L'URL non cambia, quindi su GitHub
//        non serve toccare niente.
// ============================================================

// >>> DEVE essere IDENTICO al segreto scritto in index.html <<<
// Il valore NON sta nel codice (il repository GitHub è pubblico):
// Impostazioni progetto > Proprietà script > API_SEGRETO
const API_SEGRETO = PropertiesService.getScriptProperties().getProperty('API_SEGRETO');

// Nomi delle due cartelle (vengono create da sole la prima volta)
const CARTELLA_FORMULARI = "Priamus - Formulari";
const CARTELLA_RICEVUTE  = "Priamus - Ricevute carburante";

// true  = dentro ogni cartella crea "2026-09", "2026-10", ...
// false = tutte le foto sfuse nella cartella principale
const SOTTOCARTELLE_MESE = true;

// ------------------------------------------------------------
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);

    if (!API_SEGRETO || req.segreto !== API_SEGRETO) {
      return rispondi_({ ok: false, msg: "Accesso non autorizzato." });
    }
    if (req.azione !== "foto") {
      return rispondi_({ ok: false, msg: "Questo servizio gestisce solo le foto." });
    }

    return rispondi_({ ok: true, dati: salvaFoto_(req) });

  } catch (err) {
    return rispondi_({ ok: false, msg: String(err.message || err) });
  }
}

function doGet() {
  return rispondi_({ ok: true, msg: "Servizio foto Priamus attivo." });
}

function rispondi_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
//  Salvataggio della foto + aggancio alla riga del registro
// ------------------------------------------------------------
function salvaFoto_(req) {
  if (!req.fotoData) return { msg: "Nessuna foto." };

  const rifornimento = (req.tipo === "rifornimento");

  // 1) carico il file nella cartella giusta
  const grezzo = String(req.fotoData).split(",")[1] || req.fotoData;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(grezzo),
    req.fotoMime || "image/jpeg",
    req.fotoNome || "foto.jpg"
  );

  const file = cartellaPer_(rifornimento).createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  const url = file.getUrl();

  // 2) aggancio il link alla riga giusta
  const foglio = rifornimento ? "REGISTRO_RIFORNIMENTI" : "REGISTRO_VIAGGI";
  const colonna = rifornimento ? 7 : 14;   // G oppure N

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(foglio);
  if (!sh) return { msg: "Foto caricata, ma foglio non trovato.", url: url };

  const ultima = sh.getLastRow();
  if (ultima >= 2) {
    const ids = sh.getRange(2, 1, ultima - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(req.id)) {
        const cella = sh.getRange(i + 2, colonna);
        // se ci sono gia' altre foto, il nuovo link si aggiunge in coda
        const esistente = String(cella.getValue() || "").trim();
        const elenco = esistente ? esistente.split("\n") : [];
        elenco.push(url);
        scriviLink_(cella, elenco);
        return { msg: "Foto salvata." };
      }
    }
  }
  return { msg: "Foto caricata, ma riga non trovata.", url: url };
}

// Scrive piu' indirizzi nella stessa cella rendendoli TUTTI cliccabili.
// Con setValue() Google li lascerebbe come testo semplice: riconosce
// il collegamento da solo solo quando la cella contiene un indirizzo unico.
function scriviLink_(cella, elenco) {
  const puliti = elenco.map(x => String(x || "").trim()).filter(String);
  if (!puliti.length) { cella.clearContent(); return; }

  const testo = puliti.join("\n");
  const b = SpreadsheetApp.newRichTextValue().setText(testo);

  let pos = 0;
  puliti.forEach(u => {
    if (/^https?:\/\//i.test(u)) b.setLinkUrl(pos, pos + u.length, u);
    pos += u.length + 1;   // +1 per l'a capo
  });

  cella.setRichTextValue(b.build());
}

// Restituisce la cartella dove mettere il file
function cartellaPer_(rifornimento) {
  const nomeBase = rifornimento ? CARTELLA_RICEVUTE : CARTELLA_FORMULARI;
  const base = trovaOCrea_(DriveApp.getRootFolder(), nomeBase);
  if (!SOTTOCARTELLE_MESE) return base;

  const oggi = new Date();
  const mese = oggi.getFullYear() + "-" + ("0" + (oggi.getMonth() + 1)).slice(-2);
  return trovaOCrea_(base, mese);
}

function trovaOCrea_(padre, nome) {
  const trovate = padre.getFoldersByName(nome);
  return trovate.hasNext() ? trovate.next() : padre.createFolder(nome);
}

// ------------------------------------------------------------
//  Prova: crea le due cartelle e verifica che Drive risponda
// ------------------------------------------------------------
function PROVA_DRIVE() {
  const f = cartellaPer_(false);
  const r = cartellaPer_(true);
  const prova = f.createFile(Utilities.newBlob("prova", "text/plain", "prova.txt"));
  prova.setTrashed(true);

  SpreadsheetApp.getUi().alert(
    "Drive funziona.\n\n" +
    "Formulari  ->  " + CARTELLA_FORMULARI + " / " + f.getName() + "\n" +
    "Ricevute   ->  " + CARTELLA_RICEVUTE + " / " + r.getName() + "\n\n" +
    "Le cartelle sono state create (se non c'erano gia')."
  );
}

// ------------------------------------------------------------
//  RIORDINA le foto gia' caricate prima di questa modifica.
//  Le ricevute si riconoscono dal nome del file ("Ricevuta_...").
//  Da eseguire UNA VOLTA, poi non serve piu'.
// ------------------------------------------------------------
function RIORDINA_VECCHIE_FOTO() {
  const vecchia = DriveApp.getRootFolder().getFoldersByName("Formulari Priamus");
  if (!vecchia.hasNext()) {
    SpreadsheetApp.getUi().alert(
      "Non trovo la vecchia cartella 'Formulari Priamus'.\n" +
      "Se le foto erano gia' a posto, non c'e' nulla da fare."
    );
    return;
  }

  let spostateRic = 0, spostateForm = 0;
  const radice = vecchia.next();

  const trattaCartella = (cart, nomeMese) => {
    const file = cart.getFiles();
    while (file.hasNext()) {
      const f = file.next();
      const nome = f.getName();
      const ricevuta = nome.indexOf("Ricevuta_") === 0;
      const dest = SOTTOCARTELLE_MESE
        ? trovaOCrea_(trovaOCrea_(DriveApp.getRootFolder(),
            ricevuta ? CARTELLA_RICEVUTE : CARTELLA_FORMULARI), nomeMese || "vecchie")
        : trovaOCrea_(DriveApp.getRootFolder(),
            ricevuta ? CARTELLA_RICEVUTE : CARTELLA_FORMULARI);
      try {
        f.moveTo(dest);
        if (ricevuta) spostateRic++; else spostateForm++;
      } catch (e) {}
    }
  };

  trattaCartella(radice, null);
  const sotto = radice.getFolders();
  while (sotto.hasNext()) {
    const s = sotto.next();
    trattaCartella(s, s.getName());
  }

  SpreadsheetApp.getUi().alert(
    "RIORDINO FATTO\n\n" +
    "Formulari spostati: " + spostateForm + "\n" +
    "Ricevute spostate: " + spostateRic + "\n\n" +
    "I link nel registro continuano a funzionare: spostare un file\n" +
    "su Drive non cambia il suo indirizzo.\n\n" +
    "La vecchia cartella 'Formulari Priamus' ora dovrebbe essere vuota:\n" +
    "puoi cancellarla a mano."
  );
}


// ------------------------------------------------------------
//  SISTEMA I LINK GIA' SCRITTI COME TESTO
//  Ripassa le foto caricate finora e le rende cliccabili.
//  Da eseguire UNA VOLTA. Non modifica gli indirizzi, solo il formato.
// ------------------------------------------------------------
function RENDI_CLICCABILI_LE_FOTO() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  let sistemate = 0;

  const passa = (nomeFoglio, colonna) => {
    const sh = s.getSheetByName(nomeFoglio);
    if (!sh) return;
    const ultima = sh.getLastRow();
    if (ultima < 2) return;

    const valori = sh.getRange(2, colonna, ultima - 1, 1).getValues();
    for (let i = 0; i < valori.length; i++) {
      const testo = String(valori[i][0] || "").trim();
      if (!testo || testo.indexOf("http") < 0) continue;
      const elenco = testo.split("\n").map(x => x.trim()).filter(String);
      if (elenco.length < 2) continue;      // con un solo link ci pensa Google
      scriviLink_(sh.getRange(i + 2, colonna), elenco);
      sistemate++;
    }
  };

  passa("REGISTRO_VIAGGI", 14);          // colonna N
  passa("REGISTRO_RIFORNIMENTI", 7);     // colonna G

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "FATTO\n\nCelle sistemate: " + sistemate + "\n\n" +
    "Ora ogni foto e' un collegamento cliccabile."
  );
}