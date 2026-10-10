// ============================================================
//  TEST ORA INIZIO
//  Esegue direttamente la funzione del pulsante "Inizia Tratta",
//  saltando l'app. Serve a capire se il problema e' nel CODICE
//  oppure nella DISTRIBUZIONE.
//
//  Aggiungi come file NUOVO nel progetto del database (+ > Script).
//  Esegui:  TEST_ORA_INIZIO()
// ============================================================

function TEST_ORA_INIZIO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("REGISTRO_VIAGGI");
  const dati = sh.getDataRange().getValues();

  // cerca l'ultimo viaggio ancora "Assegnato"
  let riga = -1, id = "";
  for (let i = dati.length - 1; i > 0; i--) {
    if (String(dati[i][15]).trim() === "Assegnato") {
      riga = i + 1;
      id = String(dati[i][0]);
      break;
    }
  }

  if (riga === -1) {
    SpreadsheetApp.getUi().alert(
      "Nessun viaggio in stato 'Assegnato'.\n\n" +
      "Assegnane uno dall'app (senza avviarlo) e riesegui questo test."
    );
    return;
  }

  const autista = String(dati[riga - 1][2]);
  const kmPrima = sh.getRange(riga, 6).getDisplayValue();
  const oraPrima = sh.getRange(riga, 20).getDisplayValue();

  // >>> chiama ESATTAMENTE la funzione del pulsante <<<
  let esito = "ok";
  try {
    avviaViaggioAutista(id, 99999);
  } catch (e) {
    esito = "ERRORE: " + e.message;
  }

  SpreadsheetApp.flush();
  const kmDopo = sh.getRange(riga, 6).getDisplayValue();
  const statoDopo = sh.getRange(riga, 16).getDisplayValue();
  const oraDopo = sh.getRange(riga, 20).getDisplayValue();

  SpreadsheetApp.getUi().alert(
    "VIAGGIO DI PROVA: " + autista + " (riga " + riga + ")\n" +
    "Esito chiamata: " + esito + "\n\n" +
    "KM partenza:  prima=" + kmPrima + "  dopo=" + kmDopo + "\n" +
    "Stato:        dopo=" + statoDopo + "\n" +
    "ORA INIZIO:   prima=[" + oraPrima + "]  dopo=[" + oraDopo + "]\n\n" +
    (oraDopo ?
      "IL CODICE FUNZIONA.\nIl problema e' nella distribuzione o nel link usato dagli autisti." :
      "IL CODICE NON SCRIVE L'ORA.\nIl Codice.gs nel progetto non e' la versione nuova.") +
    "\n\nRicorda: questo test ha messo 99999 nei KM e stato 'In Corso'.\nCorreggi o elimina la riga " + riga + "."
  );
}