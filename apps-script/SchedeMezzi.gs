// ============================================================
//  PRIAMUS - SCHEDE DEI MEZZI
//
//  Crea una scheda per ogni camion e per ogni rimorchio.
//  Tutto e' costruito con FORMULE: si aggiorna da solo e le
//  colonne che aggiungi tu (dalla T in poi) non vengono
//  mai toccate, nemmeno rilanciando lo script.
//
//  NOTA: tutti i nomi cominciano con SM_ / sm per non scontrarsi
//  con le costanti degli altri file del progetto (in Apps Script
//  i file condividono lo stesso spazio dei nomi).
//
//  DOVE VA: progetto Apps Script del DATABASE
//  (+ > Script, nome: SchedeMezzi)
//
//  USO: menu "🚚 Schede" > "Crea/aggiorna le schede"
//       Rilanciandolo crea solo quelle che mancano.
// ============================================================

const SM_PRIMO_ANNO = 2026;
const SM_ULTIMO_ANNO = 2029;

const SM_RIGA_DATI = 6;

// Voci di spesa, nell'ordine in cui compaiono nelle schede.
// Sono fisse perche' SUMIFS e' l'unica funzione che su questi dati
// confronta il mese in modo affidabile: QUERY e FILTER pretendono
// che il tipo della cella coincida, e li' si rompeva tutto.
const SM_VOCI = ["LEASING", "NOLEGGIO", "MUTUO", "ACQUISTO", "ANTICIPO",
  "ASSICURAZIONE", "KASKO", "BOLLO", "REVISIONE", "IMMATRICOLAZIONE",
  "INDICIZZAZIONE", "CONTR. MANUTENZIONE", "RIPARAZIONI", "RICAMBI",
  "PNEUMATICI", "LAVAGGIO"];

// Scrive il blocco spese: voci fisse + SUMIFS
function smSpese_(sh, colVoce) {
  const colImporto = colVoce + 1;
  const n = SM_VOCI.length;
  sh.getRange(SM_RIGA_DATI, colVoce, n, 1).setValues(SM_VOCI.map(v => [v]));

  const f = [];
  for (let i = 0; i < n; i++) {
    const r = SM_RIGA_DATI + i;
    f.push(['=SUMIFS(COSTI_MEZZI!$E:$E;COSTI_MEZZI!$B:$B;$B$2;' +
            'COSTI_MEZZI!$A:$A;$D$2;COSTI_MEZZI!$D:$D;$' + smLettera_(colVoce) + r + ')']);
  }
  sh.getRange(SM_RIGA_DATI, colImporto, n, 1).setFormulas(f)
    .setNumberFormat('#,##0.00 "€"');

  // totale sotto l'elenco
  const rTot = SM_RIGA_DATI + n;
  const LI = smLettera_(colImporto);
  sh.getRange(rTot, colVoce).setValue("TOTALE").setFontWeight("bold");
  sh.getRange(rTot, colImporto)
    .setFormula('=SUM(' + LI + SM_RIGA_DATI + ':' + LI + (rTot - 1) + ')')
    .setNumberFormat('#,##0.00 "€"').setFontWeight("bold").setBackground("#fef3c7");

  // le voci a zero restano grigie, cosi' saltano all'occhio quelle piene
  const grigio = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0).setFontColor("#cbd5e1")
    .setRanges([sh.getRange(SM_RIGA_DATI, colImporto, n, 1)]).build();
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([grigio]));

  return rTot;
}          // da qui in giu' compaiono i dati
const SM_BLU = "#0b192c";
const SM_ARANCIO = "#ff7800";
const SM_CHIARO = "#f4f7fa";

// ------------------------------------------------------------
function smMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("🚚 Schede")
    .addItem("Crea/aggiorna le schede", "CREA_SCHEDE")
    .addItem("Rifai una scheda sola", "RIFAI_SCHEDA")
    .addToUi();
}

// ------------------------------------------------------------
//  PREPARAZIONE (una volta sola)
//  - crea il foglio PARAMETRI con i valori che puoi cambiare
//  - aggiunge al REGISTRO_VIAGGI la colonna AA "Tipo_TR"
//    che dice se ogni viaggio e' terra (t) o rifiuti (r)
// ------------------------------------------------------------
function PREPARA_CONFRONTO() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- foglio PARAMETRI
  let p = s.getSheetByName("PARAMETRI");
  if (!p) {
    p = s.insertSheet("PARAMETRI");
    p.getRange(1, 1, 1, 3).setValues([["Parametro", "Valore", "A cosa serve"]])
      .setBackground("#0b192c").setFontColor("#FFFFFF").setFontWeight("bold");
    p.getRange(2, 1, 4, 3).setValues([
      ["Costo giornaliero",        320, "costo fisso per ogni giorno lavorato"],
      ["Costo al km",              0.9, "costo di esercizio per chilometro"],
      ["Utile giornaliero TERRA",  200, "margine atteso per giorno sui lavori di terra"],
      ["Utile giornaliero RIFIUTI",300, "margine atteso per giorno sui rifiuti"]
    ]);
    p.getRange(2, 2, 4, 1).setNumberFormat('#,##0.00 "€"')
      .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center");
    p.setColumnWidth(1, 220); p.setColumnWidth(2, 110); p.setColumnWidth(3, 330);
    p.setFrozenRows(1);
    p.setTabColor("#ff7800");
  }

  // --- colonna Tipo_TR nel registro
  const rv = s.getSheetByName("REGISTRO_VIAGGI");
  if (!rv) { ui.alert("REGISTRO_VIAGGI non trovato."); return; }
  if (rv.getMaxColumns() < 27) rv.insertColumnsAfter(rv.getMaxColumns(), 27 - rv.getMaxColumns());

  rv.getRange(1, 27).setValue("Tipo_TR")
    .setBackground("#0b192c").setFontColor("#FFFFFF").setFontWeight("bold")
    .setHorizontalAlignment("center");
  rv.getRange(2, 27).setFormula(
    '=ARRAYFORMULA(IF($A$2:$A$3000="";"";' +
    'IFERROR(LOWER(TRIM(VLOOKUP($S$2:$S$3000;REGOLE_VIAGGI!$A:$I;9;FALSE)));"")))'
  );

  SpreadsheetApp.flush();
  ui.alert(
    "PRONTO\n\n" +
    "1) Foglio PARAMETRI creato: li' cambi i 320, lo 0,9 e i 200/300\n" +
    "   e tutte le schede si aggiornano da sole.\n\n" +
    "2) REGISTRO_VIAGGI: colonna AA 'Tipo_TR' con terra/rifiuti\n" +
    "   presi da REGOLE_VIAGGI.\n\n" +
    "Ora rifai le schede dei camion per vedere il confronto."
  );
}

// ------------------------------------------------------------
function smElencoMesi_() {
  const m = [];
  for (let a = SM_PRIMO_ANNO; a <= SM_ULTIMO_ANNO; a++) {
    for (let i = 1; i <= 12; i++) m.push(a + "-" + ("0" + i).slice(-2));
  }
  return m;
}

function smMeseCorrente_() {
  const o = new Date();
  return o.getFullYear() + "-" + ("0" + (o.getMonth() + 1)).slice(-2);
}

function smTargaValida_(t) {
  return /^[A-Z]{2}\d{3}[A-Z]{2}$/.test(String(t || "").trim().toUpperCase());
}

// ------------------------------------------------------------
function CREA_SCHEDE() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const camion = [], rimorchi = [];

  const shM = s.getSheetByName("ANAGRAFICA_MEZZI");
  if (shM && shM.getLastRow() > 1) {
    shM.getRange(2, 1, shM.getLastRow() - 1, 4).getValues().forEach(r => {
      const targa = String(r[1] || "").trim().toUpperCase();
      const modello = String(r[3] || "").trim();
      if (smTargaValida_(targa)) camion.push({ targa: targa, nome: modello || targa });
    });
  }

  const shR = s.getSheetByName("ANAGRAFICA_RIMORCHI");
  if (shR && shR.getLastRow() > 1) {
    shR.getRange(2, 1, shR.getLastRow() - 1, 2).getValues().forEach(r => {
      const targa = String(r[0] || "").trim().toUpperCase();
      const nome = String(r[1] || "").trim();
      if (smTargaValida_(targa)) rimorchi.push({ targa: targa, nome: nome || targa });
    });
  }

  if (!camion.length && !rimorchi.length) {
    ui.alert("Non trovo mezzi in anagrafica."); return;
  }

  const inizio = Date.now();
  let create = 0, saltate = 0;
  const rimaste = [];

  const tutte = camion.map(x => ({ v: x, tipo: "CAMION" }))
         .concat(rimorchi.map(x => ({ v: x, tipo: "RIMORCHIO" })));

  tutte.forEach(x => {
    if (s.getSheetByName(x.v.targa)) { saltate++; return; }
    if (Date.now() - inizio > 260000) { rimaste.push(x.v.targa); return; }
    if (x.tipo === "CAMION") smCreaSchedaCamion_(s, x.v);
    else smCreaSchedaRimorchio_(s, x.v);
    create++;
  });

  SpreadsheetApp.flush();
  ui.alert(
    "SCHEDE MEZZI\n\n" +
    "Create ora: " + create + "\n" +
    "Gia' esistenti (lasciate come sono): " + saltate + "\n" +
    (rimaste.length
      ? ("\nNon ho fatto in tempo a creare: " + rimaste.length +
         "\nRiesegui 'Crea/aggiorna le schede' per completare.")
      : "\nFatto tutto.")
  );
}

// ------------------------------------------------------------
//  Rifa' TUTTE le schede dei mezzi da zero.
//  Serve dopo una correzione allo schema. Cancella anche le
//  colonne aggiunte a mano: usala solo se sai cosa stai facendo.
//  Se va in timeout, rieseguila: riprende da dove si era fermata.
// ------------------------------------------------------------
function RIFAI_TUTTE_LE_SCHEDE() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const r = ui.alert("Rifare TUTTE le schede dei mezzi?",
    "Le schede attuali vengono cancellate e ricostruite.\n" +
    "Le colonne che hai aggiunto a mano andranno perse.\n\n" +
    "Procedo?", ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const inizio = Date.now();
  let rifatte = 0;
  const rimaste = [];

  smElencoMezzi_(s).forEach(x => {
    if (Date.now() - inizio > 250000) { rimaste.push(x.targa); return; }
    const vecchia = s.getSheetByName(x.targa);
    if (vecchia) s.deleteSheet(vecchia);
    if (x.tipo === "CAMION") smCreaSchedaCamion_(s, x);
    else smCreaSchedaRimorchio_(s, x);
    rifatte++;
  });

  SpreadsheetApp.flush();
  ui.alert("Schede rifatte: " + rifatte +
    (rimaste.length ? ("\n\nNon ho fatto in tempo: " + rimaste.length +
                       "\nRiesegui 'Rifai tutte le schede' per completare.")
                    : "\n\nFatto tutto."));
}

// Elenco unico di camion e rimorchi
function smElencoMezzi_(s) {
  const elenco = [];

  const shM = s.getSheetByName("ANAGRAFICA_MEZZI");
  if (shM && shM.getLastRow() > 1) {
    shM.getRange(2, 1, shM.getLastRow() - 1, 4).getValues().forEach(r => {
      const targa = String(r[1] || "").trim().toUpperCase();
      if (smTargaValida_(targa)) {
        elenco.push({ targa: targa, nome: String(r[3] || "").trim() || targa, tipo: "CAMION" });
      }
    });
  }

  const shR = s.getSheetByName("ANAGRAFICA_RIMORCHI");
  if (shR && shR.getLastRow() > 1) {
    shR.getRange(2, 1, shR.getLastRow() - 1, 2).getValues().forEach(r => {
      const targa = String(r[0] || "").trim().toUpperCase();
      if (smTargaValida_(targa)) {
        elenco.push({ targa: targa, nome: String(r[1] || "").trim() || targa, tipo: "RIMORCHIO" });
      }
    });
  }
  return elenco;
}

// ------------------------------------------------------------
//  Corregge SOLO le formule dei giorni nelle schede gia' fatte.
//  Non tocca nient'altro: le colonne che hai aggiunto restano.
// ------------------------------------------------------------


// ------------------------------------------------------------
//  SISTEMA LA CELLA DEL MESE in tutte le schede gia' create.
//  Google aveva trasformato "2026-09" in una data: torna testo.
//  Non tocca nient'altro.
// ------------------------------------------------------------
function SISTEMA_MESE_SCHEDE() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const fuso = s.getSpreadsheetTimeZone();
  let sistemate = 0, giaOk = 0;

  smElencoMezzi_(s).forEach(x => {
    const sh = s.getSheetByName(x.targa);
    if (!sh) return;
    const cella = sh.getRange("D2");
    const v = cella.getValue();

    let testo;
    if (v instanceof Date) {
      testo = Utilities.formatDate(v, fuso, "yyyy-MM");
    } else if (typeof v === "number") {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      testo = d.getUTCFullYear() + "-" + ("0" + (d.getUTCMonth() + 1)).slice(-2);
    } else {
      testo = String(v || "").trim();
      if (/^\d{4}-\d{2}$/.test(testo)) { giaOk++; cella.setNumberFormat("@"); return; }
    }
    if (!/^\d{4}-\d{2}$/.test(testo)) testo = smMeseCorrente_();

    cella.setNumberFormat("@");   // prima il formato
    cella.setValue(testo);        // poi il valore
    sistemate++;
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "MESE DELLE SCHEDE SISTEMATO\n\n" +
    "Convertite da data a testo: " + sistemate + "\n" +
    "Gia' corrette: " + giaOk + "\n\n" +
    "Le spese dovrebbero comparire. Se cambi mese dalla tendina, resta testo."
  );
}

// Rifa' una scheda da zero (se l'hai rovinata)
function RIFAI_SCHEDA() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Rifai una scheda",
    "Scrivi la targa della scheda da rifare.\nATTENZIONE: la scheda attuale viene cancellata, comprese le colonne che hai aggiunto.",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const targa = String(r.getResponseText() || "").trim().toUpperCase();
  if (!smTargaValida_(targa)) { ui.alert("Targa non valida."); return; }

  const vecchia = s.getSheetByName(targa);
  if (vecchia) s.deleteSheet(vecchia);

  const dati = smCercaMezzo_(s, targa);
  if (!dati) { ui.alert("Targa non trovata in anagrafica."); return; }

  if (dati.tipo === "CAMION") smCreaSchedaCamion_(s, dati);
  else smCreaSchedaRimorchio_(s, dati);
  ui.alert("Scheda " + targa + " rifatta.");
}

function smCercaMezzo_(s, targa) {
  const shM = s.getSheetByName("ANAGRAFICA_MEZZI");
  if (shM && shM.getLastRow() > 1) {
    const v = shM.getRange(2, 1, shM.getLastRow() - 1, 4).getValues();
    for (const r of v) {
      if (String(r[1] || "").trim().toUpperCase() === targa) {
        return { targa: targa, nome: String(r[3] || "").trim() || targa, tipo: "CAMION" };
      }
    }
  }
  const shR = s.getSheetByName("ANAGRAFICA_RIMORCHI");
  if (shR && shR.getLastRow() > 1) {
    const v = shR.getRange(2, 1, shR.getLastRow() - 1, 2).getValues();
    for (const r of v) {
      if (String(r[0] || "").trim().toUpperCase() === targa) {
        return { targa: targa, nome: String(r[1] || "").trim() || targa, tipo: "RIMORCHIO" };
      }
    }
  }
  return null;
}

// ------------------------------------------------------------
//  Intestazione comune
// ------------------------------------------------------------
function smTestata_(sh, mezzo, colore) {
  sh.getRange("A1:G1").merge().setValue(mezzo.nome)
    .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground(colore).setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  sh.getRange("A2").setValue("Targa").setFontWeight("bold");
  sh.getRange("B2").setValue(mezzo.targa).setFontWeight("bold").setFontColor(colore);

  sh.getRange("C2").setValue("Mese").setFontWeight("bold");
  const cellaMese = sh.getRange("D2");
  cellaMese.setNumberFormat("@");            // TESTO: altrimenti "2026-09" diventa una data
  cellaMese.setValue(smMeseCorrente_())
    .setBackground("#fff7ed").setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  cellaMese.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(smElencoMesi_(), true)
      .setAllowInvalid(false).build());

  // estremi del mese, servono alle formule
  sh.getRange("E2").setValue("dal").setFontColor("#94a3b8").setHorizontalAlignment("right");
  sh.getRange("F2").setFormula('=DATE(VALUE(LEFT($D$2;4));VALUE(RIGHT($D$2;2));1)')
    .setNumberFormat("dd/mm/yyyy").setFontColor("#94a3b8");
  sh.getRange("G2").setValue("al").setFontColor("#94a3b8").setHorizontalAlignment("right");
  sh.getRange("H2").setFormula('=EOMONTH($F$2;0)').setNumberFormat("dd/mm/yyyy").setFontColor("#94a3b8");
}

function smBanda_(sh, cella, testo, colore) {
  sh.getRange(cella).merge().setValue(testo)
    .setBackground(colore).setFontColor("#FFFFFF").setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function smIntestazioni_(sh, riga, colonna, titoli) {
  sh.getRange(riga, colonna, 1, titoli.length).setValues([titoli])
    .setBackground(SM_CHIARO).setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setWrap(true)
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
}


// ============================================================
//  REGISTRO DEL MESE
//  Un elenco piatto di tutti i viaggi di tutti i camion,
//  ordinati per data. A:L, con i filtri sulle intestazioni.
//
//  Tutto a formule: si aggiorna da solo.
//  I filtri NASCONDONO le righe e funzionano; l'ordinamento
//  no, perche' le righe sono il risultato di una formula.
//  Per ordinare usa Dati > Visualizzazioni filtro.
// ============================================================
const SM_FOGLIO_REGISTRO = "REGISTRO_MESE";
const SM_ULTIMA_RIGA_REG = 2000;







// ------------------------------------------------------------
//  Confronto teorico/effettivo di TUTTA la flotta,
//  su un intervallo di date scelto da te (celle O2 e Q2).
//  I "giorni" sono giornate-camion: se in un giorno lavorano
//  nove mezzi valgono nove, perche' ognuno ha il suo costo fisso.
// ------------------------------------------------------------


// Elenco dei mesi che hanno almeno un viaggio, dal piu' vecchio
function smMesiConDati_(s) {
  const rv = s.getSheetByName("REGISTRO_VIAGGI");
  const fuso = s.getSpreadsheetTimeZone();
  const trovati = {};
  if (rv && rv.getLastRow() > 1) {
    rv.getRange(2, 2, rv.getLastRow() - 1, 1).getValues().forEach(r => {
      const d = r[0];
      if (d instanceof Date) trovati[Utilities.formatDate(d, fuso, "yyyy-MM")] = true;
    });
  }
  const elenco = Object.keys(trovati).sort();
  return elenco.length ? elenco : [smMeseCorrente_()];
}



// Colonna €/km: la scrivi tu, viaggio per viaggio. Vuota = valore di PARAMETRI.
// Confronto MESE PER MESE dentro il registro del mese: spento.
// Ora sta nel foglio REGISTRO_ANNO (comando "Crea il registro anno").
const SM_MOSTRA_MENSILE = false;

const SM_COL_EUROKM = 12;     // L  €/km, la scrivi tu
const SM_COL_COSTOG_MAN = 10; // J  costo giornaliero manuale, lo scrivi tu
const SM_COL_ID = 27;         // AA nascosta: id del viaggio, serve a non perdere le modifiche

function CREA_REGISTRO_MESE() {
  const t0 = Date.now();
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const U = SM_ULTIMA_RIGA_REG;
  smAssicuraParametri_(s);
  const par = s.getSheetByName("PARAMETRI");
  const pred = Number(par.getRange("B3").getValue()) || 0.9;
  // costo giornaliero dell'APPALTI NOTTE quando il camion ha fatto anche altri lavori
  if (String(par.getRange("A6").getValue()).trim() === "") {
    par.getRange("A6:C6").setValues([["Costo giornaliero APPALTI NOTTE", 180,
      "fisso dell'appalti notte quando quel giorno il camion ha fatto anche altri lavori"]]);
    par.getRange("B6").setNumberFormat('#,##0.00 "€"')
      .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center");
  }
  // soglia del MARGINE %: sotto diventa rosso
  if (String(par.getRange("A7").getValue()).trim() === "") {
    par.getRange("A7:C7").setValues([["Margine % minimo", 0.3,
      "sotto questa percentuale il margine % del registro diventa rosso"]]);
    par.getRange("B7").setNumberFormat("0%")
      .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center");
  }

  // --- salvo i €/km che hai cambiato (solo quelli diversi dal valore predefinito)
  // --- salvo quello che hai scritto a mano: €/km e costo giornaliero manuale.
  //     Le colonne le cerco dal NOME dell'intestazione, non dalla posizione:
  //     cosi' funziona anche quando la disposizione del registro cambia.
  const salvati = {}, salvatiCG = {};
  const vecchio = s.getSheetByName(SM_FOGLIO_REGISTRO);
  if (vecchio) {
    const ult = vecchio.getLastRow();
    if (ult > 1 && vecchio.getMaxColumns() >= SM_COL_ID) {
      const testa = vecchio.getRange(1, 1, Math.min(ult, 300), 2).getValues();
      let rV = -1;
      for (let x = 0; x < testa.length; x++) {
        if (String(testa[x][0]).trim() === "Mese" && String(testa[x][1]).trim() === "Data") { rV = x + 1; break; }
      }
      if (rV > 0) {
        const nomi = vecchio.getRange(rV, 1, 1, 25).getValues()[0]
          .map(x => String(x || "").replace(/\s+/g, " ").trim().toLowerCase());
        const cE = nomi.indexOf("€/km") + 1;
        const cC = nomi.indexOf("costo giornaliero manuale") + 1;
        const ids = vecchio.getRange(1, SM_COL_ID, ult, 1).getValues();
        const vE = cE > 0 ? vecchio.getRange(1, cE, ult, 1).getValues() : null;
        const vC = cC > 0 ? vecchio.getRange(1, cC, ult, 1).getValues() : null;
        ids.forEach((r, x) => {
          if (x < rV) return;
          const id = String(r[0] || "").trim();
          if (!id) return;
          if (vE) { const v = vE[x][0]; if (v !== "" && v != null && Number(v) !== pred) salvati[id] = Number(v); }
          if (vC) { const v = vC[x][0]; if (v !== "" && v != null) salvatiCG[id] = Number(v); }
        });
      }
    }
    // li rimette a posto "Compila €/km"
    PropertiesService.getDocumentProperties()
      .setProperty("SM_EUROKM_SALVATI", JSON.stringify(salvati))
      .setProperty("SM_COSTOG_SALVATI", JSON.stringify(salvatiCG));
  }

  // Il foglio NON viene cancellato: lo svuoto e lo riscrivo.
  // Cancellandolo, il registro del giorno e il registro mezzo (che leggono da qui)
  // perderebbero il collegamento e mostrerebbero errori.
  let sh;
  if (vecchio) {
    sh = vecchio;
    const filtro = sh.getFilter();
    if (filtro) filtro.remove();
    sh.setFrozenRows(0);
    const tutto = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    tutto.breakApart();
    tutto.clearDataValidations();
    sh.clearConditionalFormatRules();
    sh.clear();
    sh.showColumns(1, sh.getMaxColumns());
  } else {
    sh = s.insertSheet(SM_FOGLIO_REGISTRO, 0);
  }
  sh.setTabColor("#6d28d9");
  if (sh.getMaxColumns() < 44) sh.insertColumnsAfter(sh.getMaxColumns(), 44 - sh.getMaxColumns());

  const mesi = smMesiConDati_(s);
  const rMese = 11;                                       // (confronto mese per mese, se attivo)
  const rUltimoMese = rMese + 1 + mesi.length * 4;        // 4 righe per mese: terra, rifiuti, pianale, totale
  const rTest = SM_MOSTRA_MENSILE ? rUltimoMese + 2 : 10;  // intestazioni elenco viaggi
  const R1 = rTest + 1;

  // ---------- testata ----------
  sh.getRange("A1:T1").merge().setValue("REGISTRO VIAGGI  —  tutti i mesi")
    .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground("#6d28d9").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);

  const oggi = new Date();
  const primo = new Date(oggi.getFullYear(), oggi.getMonth(), 1);
  sh.getRange("A2").setValue("Dal").setFontWeight("bold").setHorizontalAlignment("right");
  sh.getRange("B2").setValue(primo).setNumberFormat("dd/mm/yyyy")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange("C2").setValue("al").setFontWeight("bold").setHorizontalAlignment("right");
  sh.getRange("D2").setValue(oggi).setNumberFormat("dd/mm/yyyy")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  // ---------- TRATTA: filtra tutta la tabella (elenco e confronto) ----------
  // AP (nascosta): "Tutte" + le tratte presenti nel registro, per la tendina
  sh.getRange("AP1").setValue("Tutte");
  sh.getRange("AP2").setFormula('=IFERROR(SORT(UNIQUE(FILTER($F$' + R1 + ':$F$' + U + ';$F$' + R1 + ':$F$' + U + '<>"")));"")');
  sh.getRange("F2").setValue("Tratta").setFontWeight("bold").setHorizontalAlignment("right");
  sh.getRange("G2:I2").merge().setValue("Tutte")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(sh.getRange("AP1:AP300"), true).setAllowInvalid(false).build());
  sh.getRange("K2").setValue("← Dal/al e Tratta filtrano tutta la tabella: elenco dei viaggi e confronto. (Tratta: Tutte = nessun filtro)")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);

  // ---------- dati grezzi, nascosti (AB:AL) ----------
  sh.getRange(R1, 28).setFormula(
    '=IFERROR(LET(q;QUERY(REGISTRO_VIAGGI!$A$2:$AB;' +
    '"select A,B,E,C,S,AA,H,L,M,Z,AB where E is not null and B is not null";0);' +
    'SORT(q;2;TRUE;3;TRUE;ARRAYFORMULA(ISNUMBER(SEARCH("notte";INDEX(q;;5))));TRUE;1;TRUE));"")');

  const rg = (L) => '$' + L + '$' + R1 + ':$' + L + '$' + U;
  const D = rg('AC');
  const copia = (col, src) =>
    sh.getRange(R1, col).setFormula('=ARRAYFORMULA(IF(' + D + '="";"";' + src + '))');

  // =====================================================================
  //  COLONNE VISIBILI
  //  A Mese   B Data   C Targa   D Categoria   E Autista   F Tratta   G T/R
  //  H Costo AUTISTA (automatico)   I Costi FISSI del mezzo (automatico)
  //  J Costo giornaliero MANUALE (tua: sostituisce autista + fissi, conta tutto come autista)
  //  K Km   L €/km (tua)   M Totale costo
  //  N Tonn.  O €/ton   P Ore   Q €/ora
  //  R Ricavo €   S Margine reale del giorno   T Teorico del giorno
  //
  //  COLONNE NASCOSTE (aiuti di calcolo)
  //  U costo giornaliero usato (manuale se c'e', altrimenti autista + fissi)
  //  V 1=viaggio normale/0=appalti   W costo della riga
  //  X chiave gruppo   Y chiave giorno   Z ricavo usato (manuale o calcolato)
  //  AA id   AB:AL dati grezzi (AL = Totale_Ricavi_Manuale)
  //  AM riga visibile   AN/AO costo e ricavo della giornata   AP elenco tratte
  //  AQ parte AUTISTA usata   AR parte FISSI usata (con il manuale: tutto autista, fissi 0)
  // =====================================================================
  sh.getRange(R1, 1).setFormula('=ARRAYFORMULA(IF(' + D + '="";"";TEXT(' + D + ';"yyyy-mm")))');
  copia(2, D);                 // B Data
  copia(3, rg('AD'));          // C Targa
  sh.getRange(R1, 4).setFormula(   // D Categoria
    '=ARRAYFORMULA(IF(' + D + '="";"";IFERROR(VLOOKUP(' + rg('AD') + ';ANAGRAFICA_MEZZI!$B:$C;2;FALSE);"")))');
  copia(5, rg('AE'));          // E Autista
  copia(6, rg('AF'));          // F Tratta
  copia(7, rg('AG'));          // G T/R
  copia(11, rg('AH'));         // K Km
  copia(14, rg('AI'));         // N Tonn.
  copia(16, rg('AJ'));         // P Ore
  copia(27, rg('AB'));         // AA id (nascosta)

  // --- O €/ton: la tariffa a tonnellata della tratta (vuota se la tratta non si paga a tonnellata)
  const tTon = 'IFERROR(VLOOKUP(' + rg('AF') + ';REGOLE_VIAGGI!$A:$K;6;FALSE);0)';
  sh.getRange(R1, 15).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + tTon + '>0;' + tTon + ';"")))');

  // --- Q €/ora: la tariffa oraria APPLICATA, con quella dei 4Assi se esiste
  const t4A = 'IFERROR(VLOOKUP(' + rg('AF') + ';REGOLE_VIAGGI!$A:$K;11;FALSE);0)';
  const tOra = 'IFERROR(VLOOKUP(' + rg('AF') + ';REGOLE_VIAGGI!$A:$K;7;FALSE);0)';
  const tOraApplicata = 'IF((LOWER(' + rg('D') + ')="4assi")*(' + t4A + '>0);' + t4A + ';' + tOra + ')';
  sh.getRange(R1, 17).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + tOraApplicata + '>0;' + tOraApplicata + ';"")))');

  // --- aiuti nascosti per i gruppi (stesso camion, stesso giorno; appalti notte a parte)
  // LAVORO NOTTURNO: qualunque tratta con "notte" nel nome (Appalti notte, Slim notte, ...)
  // e' trattata come l'appalti notte: gruppo a se', 180 se il camion ha fatto anche altro.
  const isApp = '(ISNUMBER(SEARCH("notte";' + rg('AF') + ')))';
  sh.getRange(R1, 22).setFormula(          // V: 1 normale, 0 appalti notte
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + isApp + ';0;1)))');
  sh.getRange(R1, 25).setFormula(          // Y: chiave del giorno
    '=ARRAYFORMULA(IF(' + D + '="";"";TEXT(' + D + ';"yyyymmdd")&"|"&' + rg('AD') + '))');
  sh.getRange(R1, 24).setFormula(          // X: chiave del gruppo
    '=ARRAYFORMULA(IF(' + D + '="";"";' + rg('Y') + '&"|"&' + rg('V') + '))');

  const chiave = rg('X');
  const chiaveGiorno = rg('Y');
  const primo_ = '(' + chiave + '<>$X$' + (R1 - 1) + ':$X$' + (U - 1) + ')';
  // utile atteso della giornata: terra B4, rifiuti B5, PIANALE B9
  const utile = 'IF(' + rg('G') + '="r";PARAMETRI!$B$5;IF(' + rg('G') + '="t";PARAMETRI!$B$4;IF(' + rg('G') + '="p";PARAMETRI!$B$9;0)))';
  const nNormali = 'SUMIF(' + chiaveGiorno + ';' + chiaveGiorno + ';' + rg('V') + ')';

  // H Costo AUTISTA e I Costi FISSI del mezzo, AUTOMATICI, sulla prima riga di ogni gruppo:
  //  SABATO e DOMENICA -> autista 100, fissi 0 (per qualsiasi gruppo, appalti notte compreso)
  //  dal lunedi' al venerdi':
  //    gruppo normale e appalti notte da solo -> autista 202 + fissi 118 (= 320)
  //    appalti notte con altri lavori         -> autista 180, fissi 0 (il mezzo e' gia' pagato di giorno)
  //  202 e 118 stanno in PARAMETRI B10 e B11 (B2 = B10 + B11).
  sh.getRange(R1, 8).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + primo_ + ';' +
    'IF(WEEKDAY(' + D + ';2)>=6;PARAMETRI!$B$8;' +
    'IF(' + isApp + ';IF(' + nNormali + '>0;PARAMETRI!$B$6;PARAMETRI!$B$10);PARAMETRI!$B$10));0)))');
  sh.getRange(R1, 9).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + primo_ + ';' +
    'IF(WEEKDAY(' + D + ';2)>=6;0;' +
    'IF(' + isApp + ';IF(' + nNormali + '>0;0;PARAMETRI!$B$11);PARAMETRI!$B$11));0)))');

  // U (nascosta): costo giornaliero USATO = il tuo (colonna J) se l'hai scritto, altrimenti autista + fissi
  sh.getRange(R1, 21).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + rg('J') + '="";N(' + rg('H') + ')+N(' + rg('I') + ');' + rg('J') + ')))');

  // AQ / AR (nascoste): parte AUTISTA e parte FISSI del costo usato.
  // Senza manuale sono H e I. Con il manuale, il manuale va TUTTO al costo autista (fissi 0).
  sh.getRange(R1, 43).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + rg('J') + '="";N(' + rg('H') + ');N(' + rg('J') + '))))');
  sh.getRange(R1, 44).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";N(' + rg('U') + ')-N(' + rg('AQ') + ')))');

  // W (nascosta): costo della singola riga = costo giornaliero usato + km x €/km
  sh.getRange(R1, 23).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";N(' + rg('U') + ')+N(' + rg('K') + ')*' +
    'IF(' + rg('L') + '="";PARAMETRI!$B$3;' + rg('L') + ')))');

  // Z (nascosta): ricavo del viaggio per il REGISTRO =
  //   Totale_Ricavi (colonna Z di REGISTRO_VIAGGI, qui in AK) + Totale_Ricavi_Manuale (colonna AB, qui in AL).
  // Il manuale si SOMMA sempre al calcolato (supplementi; per togliere si scrive un negativo).
  sh.getRange(R1, 26).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";N(' + rg('AK') + ')+N(' + rg('AL') + ')))');

  const costoGruppo = 'SUMIF(' + chiave + ';' + chiave + ';' + rg('W') + ')';
  const ricavoGruppo = 'SUMIF(' + chiave + ';' + chiave + ';' + rg('Z') + ')';

  // Con "Tutte" (o vuota): costo, ricavo e margine della GIORNATA sulla prima riga del gruppo.
  // Con una tratta scelta: i valori del SINGOLO viaggio (il fisso resta sul primo viaggio del giorno).
  const tutte = '(OR($G$2="";$G$2="Tutte"))';
  sh.getRange(R1, 13).setFormula(          // M Totale costo
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + tutte + ';IF(' + primo_ + ';' + costoGruppo + ';"");' + rg('W') + ')))');
  sh.getRange(R1, 18).setFormula(          // R Ricavo
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + tutte + ';IF(' + primo_ + ';' + ricavoGruppo + ';"");' + rg('Z') + ')))');
  sh.getRange(R1, 19).setFormula(          // S Margine reale = ricavo - costo
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + tutte + ';IF(' + primo_ + ';' + ricavoGruppo + '-' + costoGruppo + ';"");' +
    rg('Z') + '-' + rg('W') + ')))');
  // AN / AO (nascoste): costo e ricavo della giornata SEMPRE sulla prima riga, qualunque tratta sia scelta.
  sh.getRange(R1, 40).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + primo_ + ';' + costoGruppo + ';"")))');
  sh.getRange(R1, 41).setFormula(
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + primo_ + ';' + ricavoGruppo + ';"")))');
  // AM (nascosta): 1 se la riga e' visibile (non nascosta dal filtro). Serve al confronto in alto.
  const vis = [];
  for (let r = R1; r <= U; r++) vis.push(['=IF($B' + r + '="";0;SUBTOTAL(103;$C' + r + '))']);
  sh.getRange(R1, 39, vis.length, 1).setFormulas(vis);
  sh.getRange(R1, 20).setFormula(          // T Teorico = costo + utile atteso
    '=ARRAYFORMULA(IF(' + D + '="";"";IF(' + primo_ + ';' + costoGruppo + '+' + utile + ';"")))');

  // ---------- intestazioni dell'elenco ----------
  const titoli = ["Mese", "Data", "Targa", "Categoria", "Autista", "Tratta", "T/R",
    "Costo\nautista", "Costi fissi\nmezzo", "Costo giornaliero\nmanuale", "Km", "€/km", "Totale\ncosto",
    "Tonn.", "€/ton", "Ore", "€/ora", "Ricavo €", "Margine reale\ndel giorno", "Teorico\ndel giorno"];
  smIntestazioni_(sh, rTest, 1, titoli);
  sh.getRange(rTest, 1, 1, 20).setBackground("#0b192c").setFontColor("#FFFFFF");
  // le due colonne in cui scrivi tu: arancioni
  [SM_COL_COSTOG_MAN, SM_COL_EUROKM].forEach(c => {
    sh.getRange(rTest, c).setBackground("#b45309");
    sh.getRange(R1, c, U - R1 + 1, 1).setBackground("#fffbeb");
  });
  // le due tariffe si leggono meglio con uno sfondo tenue
  [15, 17].forEach(c => sh.getRange(R1, c, U - R1 + 1, 1).setBackground("#f0f9ff"));
  sh.setRowHeight(rTest, 38);

  [[1,80],[2,90],[3,95],[4,95],[5,165],[6,150],[7,45],[8,85],[9,85],[10,105],[11,65],[12,65],
   [13,100],[14,70],[15,65],[16,60],[17,65],[18,100],[19,115],[20,105]]
    .forEach(x => sh.setColumnWidth(x[0], x[1]));
  sh.getRange(R1, 2, U - R1 + 1, 1).setNumberFormat("dd/mm/yyyy");
  sh.getRange(R1, 11, U - R1 + 1, 1).setNumberFormat("#,##0");
  [12, 15, 17].forEach(c => sh.getRange(R1, c, U - R1 + 1, 1).setNumberFormat('0.00 "€"'));
  [14, 16].forEach(c => sh.getRange(R1, c, U - R1 + 1, 1).setNumberFormat("#,##0.0"));
  [8, 9, 10, 13, 18, 19, 20].forEach(c => sh.getRange(R1, c, U - R1 + 1, 1).setNumberFormat('#,##0.00 "€"'));
  [1, 2, 7, 12, 15, 17].forEach(c => sh.getRange(R1, c, U - R1 + 1, 1).setHorizontalAlignment("center"));

  // margine del giorno: rosso sotto l'utile atteso (200 terra / 300 rifiuti), verde sopra
  const zona = [sh.getRange(R1, 19, U - R1 + 1, 1)];
  // Le regole di colore NON possono leggere un altro foglio direttamente:
  // Google le rifiuta e lo script si blocca. INDIRECT e' l'unico modo ammesso.
  const soglia = 'IF($G' + R1 + '="r";INDIRECT("PARAMETRI!B5");IF($G' + R1 + '="p";INDIRECT("PARAMETRI!B9");INDIRECT("PARAMETRI!B4")))';
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($S' + R1 + ');$S' + R1 + '<' + soglia + ')')
      .setFontColor("#b91c1c").setBold(true).setBackground("#fee2e2").setRanges(zona).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($S' + R1 + ');$S' + R1 + '>=' + soglia + ')')
      .setFontColor("#15803d").setBold(true).setBackground("#dcfce7").setRanges(zona).build()
  ]));

  smRiepiloghiV22_(sh, R1, U, mesi, rMese, rUltimoMese);

  sh.hideColumns(21, 24);   // da U ad AR: solo aiuti di calcolo
  sh.setFrozenRows(rTest);
  // NIENTE setFrozenColumns: il titolo in riga 1 e' una cella unita A:T,
  // e Google rifiuta di bloccare colonne che la taglierebbero a meta'.
  sh.getRange(rTest, 1, U - rTest + 1, 20).createFilter();
  smInstallaTriggerModifica_();             // Dal/al e Tratta applicano il filtro da soli
  smApplicaFiltri_(sh);                     // il periodo vale subito anche per l'elenco
  SpreadsheetApp.flush();

  // --- €/km: li compilo subito solo se resta tempo. Il calcolo del foglio
  //     puo' essere lungo e Google concede al massimo 6 minuti.
  let compilati = -1;
  if (Date.now() - t0 < 180000) {
    compilati = smCompilaEuroKm_(sh, R1, U, t0);
  }

  s.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    "REGISTRO creato.\n\n" +
    (compilati >= 0
      ? ("Colonna €/km compilata (" + compilati + " viaggi).")
      : ("Il foglio sta ancora calcolando.\n" +
         "Tra un minuto lancia:  Priamus > Schede mezzi > Compila €/km")) +
    "\n\nLa colonna €/km (arancione) e' tua: cambiala viaggio per viaggio.");
}


// ------------------------------------------------------------
//  Riepiloghi in cima al registro.
//  Il costo si somma riga per riga (colonna Totale costo), perche'
//  il €/km puo' cambiare da un viaggio all'altro.
// ------------------------------------------------------------
function smRiepiloghiV22_(sh, R1, U, mesi, rMese, rUltimoMese) {
  const rg = (L) => '$' + L + '$' + R1 + ':$' + L + '$' + U;
  // FISSO = costo giornaliero USATO (colonna nascosta U: manuale se c'e', altrimenti autista + fissi)
  // AUT / FIS = le due parti del costo usato (colonne nascoste AQ e AR)
  // Valori del SINGOLO viaggio (colonne nascoste): W costo della riga, Z ricavo della riga.
  // La somma e' identica a quella per giornata, ma cosi' il confronto funziona anche filtrando per tratta.
  const DATE = rg('B'), TARG = rg('C'), TR = rg('G'), FISSO = rg('U'),
        KM = rg('K'), TOT = rg('W'), RIC = rg('Z'), MESE = rg('A'),
        AUT = rg('AQ'), FIS = rg('AR');
  const VIS = rg('AM');     // 1 = riga visibile (non nascosta dal filtro della tendina Tratta)
  // giornate VERE: camion + data. L'appalti notte non aggiunge giornate:
  // il suo utile in piu' entra nel teorico, sommando la colonna "Teorico del giorno".
  const chiave = rg('Y');   // chiave del giorno, colonna nascosta Y
  const TEOR = rg('T');

  // Ordine delle colonne (scostamento dalla prima colonna numerica):
  //  +0 Giornate  +1 Km  +2 Costo AUTISTA  +3 Costi FISSI  +4 Costo giornaliero medio
  //  +5 €/km medio  +6 Costo km  +7 TOTALE COSTO
  //  +8 RICAVO EFFETTIVO  +9 MARGINE  +10 MARGINE %
  //  +11 MARGINE REALE per giornata  +12 Utile atteso per giornata  +13 RICAVO TEORICO
  //  +14 DIFFERENZA
  const N_COL = 15;          // colonne numeriche
  const intest = ["Giornate\ncamion", "Km", "Costo\nautista", "Costi fissi\nmezzo", "Costo giornaliero\nmedio",
    "€/km\nmedio", "Costo km",
    "TOTALE\nCOSTO", "RICAVO\nEFFETTIVO", "MARGINE", "MARGINE\n%",
    "MARGINE REALE\nper giornata", "Utile atteso\nper giornata", "RICAVO\nTEORICO", "DIFFERENZA"];

  // formule comuni alle righe di dettaglio e ai totali (calcolate dalle celle della stessa riga)
  function derivate(r, c0) {
    const L = (n) => smLettera_(c0 + n);
    // costo giornaliero MEDIO: i fissi pagati (autista + fissi; 100 nel weekend, 180 appalti) / giornate
    sh.getRange(r, c0 + 4).setFormula('=IFERROR((' + L(7) + r + '-' + L(6) + r + ')/' + L(0) + r + ';"")');
    sh.getRange(r, c0 + 5).setFormula('=IFERROR(' + L(6) + r + '/' + L(1) + r + ';"")');           // €/km medio
    sh.getRange(r, c0 + 9).setFormula('=' + L(8) + r + '-' + L(7) + r);                            // margine
    sh.getRange(r, c0 + 10).setFormula('=IFERROR(' + L(9) + r + '/' + L(8) + r + ';"")');          // margine %
    sh.getRange(r, c0 + 11).setFormula('=IFERROR(' + L(9) + r + '/' + L(0) + r + ';"")');          // per giornata
    // utile atteso per giornata REALE: 200/300 nei giorni normali, di piu' con l'appalti notte
    sh.getRange(r, c0 + 12).setFormula('=IFERROR((' + L(13) + r + '-' + L(7) + r + ')/' + L(0) + r + ';"")');
    sh.getRange(r, c0 + 14).setFormula('=' + L(8) + r + '-' + L(13) + r);                          // differenza
  }

  function riga(r, c0, cond, crit, par) {
    // conta solo le righe visibili
    cond = cond + ';' + VIS + '=1';
    crit = crit + ';' + VIS + ';1';
    sh.getRange(r, c0).setFormula(
      '=IF(COUNTIFS(' + crit + ')=0;0;COUNTUNIQUE(FILTER(' + chiave + ';' + cond + ')))');
    sh.getRange(r, c0 + 1).setFormula('=SUMIFS(' + KM + ';' + crit + ')');
    sh.getRange(r, c0 + 2).setFormula('=SUMIFS(' + AUT + ';' + crit + ')');      // costo autista
    sh.getRange(r, c0 + 3).setFormula('=SUMIFS(' + FIS + ';' + crit + ')');      // costi fissi
    sh.getRange(r, c0 + 6).setFormula('=SUMIFS(' + TOT + ';' + crit + ')-SUMIFS(' + FISSO + ';' + crit + ')');
    sh.getRange(r, c0 + 7).setFormula('=SUMIFS(' + TOT + ';' + crit + ')');
    sh.getRange(r, c0 + 8).setFormula('=SUMIFS(' + RIC + ';' + crit + ')');
    // teorico = costo + giornate x utile FISSO della riga (200 terra, 300 rifiuti).
    // Cosi' l'utile atteso per giornata (calcolato in derivate) vale esattamente 200 o 300.
    sh.getRange(r, c0 + 13).setFormula('=' + smLettera_(c0 + 7) + r + '+' + smLettera_(c0) + r + '*PARAMETRI!' + par);
    derivate(r, c0);
  }

  // riga di totale: somma le grandezze, poi ricalcola i rapporti (che non si sommano)
  function totale(r, c0, a, b) {
    const L = (n) => smLettera_(c0 + n);
    [0, 1, 2, 3, 6, 7, 8, 13].forEach(n =>
      sh.getRange(r, c0 + n).setFormula('=SUM(' + L(n) + a + ':' + L(n) + b + ')'));
    derivate(r, c0);
  }

  function stile(r0, n, c0, ncol) {
    sh.getRange(r0, c0, n, 2).setNumberFormat("#,##0");
    sh.getRange(r0, c0 + 2, n, 3).setNumberFormat('#,##0.00 "€"');   // autista, fissi, costo giornaliero medio
    sh.getRange(r0, c0 + 5, n, 1).setNumberFormat('0.00 "€"');
    sh.getRange(r0, c0 + 6, n, 4).setNumberFormat('#,##0.00 "€"');   // costo km, tot costo, effettivo, margine
    sh.getRange(r0, c0 + 10, n, 1).setNumberFormat('0.0%');           // margine %
    sh.getRange(r0, c0 + 11, n, 4).setNumberFormat('#,##0.00 "€"');  // per giornata, utile, teorico, differenza
    sh.getRange(r0, 1, n, ncol).setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  }

  function colori(r0, n, c0) {
    // margine e differenza: rosso sotto zero, verde sopra
    const segno = [c0 + 9, c0 + 14].map(c => sh.getRange(r0, c, n, 1));
    // margine %: rosso sotto la soglia di PARAMETRI (B7, di base 30%), verde da li' in su.
    // INDIRECT: le regole di colore non possono leggere un altro foglio direttamente.
    const LP = smLettera_(c0 + 10);
    const zP = [sh.getRange(r0, c0 + 10, n, 1)];
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISNUMBER($' + LP + r0 + ');$' + LP + r0 + '<INDIRECT("PARAMETRI!B7"))')
        .setFontColor("#b91c1c").setBold(true).setBackground("#fee2e2").setRanges(zP).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISNUMBER($' + LP + r0 + ');$' + LP + r0 + '>=INDIRECT("PARAMETRI!B7"))')
        .setFontColor("#15803d").setBold(true).setBackground("#dcfce7").setRanges(zP).build()
    ]));
    // margine per giornata: rosso sotto l'utile atteso, verde da li' in su
    const cM = c0 + 11, cU = c0 + 12;
    const LM = smLettera_(cM), LU = smLettera_(cU);
    const zM = [sh.getRange(r0, cM, n, 1)];
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
        .setFontColor("#b91c1c").setBold(true).setRanges(segno).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setFontColor("#15803d").setBold(true).setRanges(segno).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISNUMBER($' + LM + r0 + ');ISNUMBER($' + LU + r0 + ');$' + LM + r0 + '<$' + LU + r0 + ')')
        .setFontColor("#b91c1c").setBold(true).setBackground("#fee2e2").setRanges(zM).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(ISNUMBER($' + LM + r0 + ');ISNUMBER($' + LU + r0 + ');$' + LM + r0 + '>=$' + LU + r0 + ')')
        .setFontColor("#15803d").setBold(true).setBackground("#dcfce7").setRanges(zM).build()
    ]));
  }

  // ====== periodo scelto (righe 3-8): Tipo + 15 colonne = A:P ======
  const ultP = smLettera_(1 + N_COL);
  smBanda_(sh, "A3:" + ultP + "3", "CONFRONTO TEORICO / EFFETTIVO  —  PERIODO SCELTO", "#6d28d9");
  smIntestazioni_(sh, 4, 1, ["Tipo"].concat(intest));
  [{ r: 5, nome: "Terra", cod: "t", par: "$B$4" }, { r: 6, nome: "Rifiuti", cod: "r", par: "$B$5" },
   { r: 7, nome: "Pianale", cod: "p", par: "$B$9" }].forEach(x => {
    sh.getRange(x.r, 1).setValue(x.nome).setFontWeight("bold");
    riga(x.r, 2,
      // fine periodo: "< giorno dopo", perche' le date dell'app hanno anche l'ora (es. 13/10 ore 12)
      TR + '="' + x.cod + '";' + DATE + '>=$B$2;' + DATE + '<$D$2+1',
      TR + ';"' + x.cod + '";' + DATE + ';">="&$B$2;' + DATE + ';"<"&($D$2+1)', x.par);
  });
  sh.getRange(8, 1).setValue("TOTALE").setFontWeight("bold");
  totale(8, 2, 5, 7);
  // giornate REALI del periodo: un camion che nello stesso giorno fa due tipi di lavoro
  // conta una volta sola (nelle righe dei tipi invece compare in entrambe)
  const unTipo = '((' + TR + '="t")+(' + TR + '="r")+(' + TR + '="p"))>0';
  sh.getRange(8, 2).setFormula('=IFERROR(ROWS(UNIQUE(FILTER(' + chiave + ';' + DATE + '>=$B$2;' + DATE + '<$D$2+1;' +
    VIS + '=1;' + unTipo + ')));0)');
  sh.getRange(8, 1, 1, 1 + N_COL).setFontWeight("bold").setBackground("#ede9fe");
  stile(5, 4, 2, 1 + N_COL);
  colori(5, 4, 2);

  // ====== mese per mese (spento con SM_MOSTRA_MENSILE = false): Mese + Tipo + 15 = A:Q ======
  if (!SM_MOSTRA_MENSILE) return;
  const ultM = smLettera_(2 + N_COL);
  smBanda_(sh, "A10:" + ultM + "10", "CONFRONTO MESE PER MESE", "#0369a1");
  smIntestazioni_(sh, rMese, 1, ["Mese", "Tipo"].concat(intest));
  mesi.forEach((m, i) => {
    const base = rMese + 1 + i * 4;
    sh.getRange(base, 1).setNumberFormat("@").setValue(m).setFontWeight("bold");
    [{ r: base, nome: "Terra", cod: "t", par: "$B$4" }, { r: base + 1, nome: "Rifiuti", cod: "r", par: "$B$5" },
     { r: base + 2, nome: "Pianale", cod: "p", par: "$B$9" }].forEach(x => {
      sh.getRange(x.r, 2).setValue(x.nome);
      riga(x.r, 3,
        MESE + '=$A$' + base + ';' + TR + '="' + x.cod + '"',
        MESE + ';$A$' + base + ';' + TR + ';"' + x.cod + '"', x.par);
    });
    const t = base + 3;
    sh.getRange(t, 2).setValue("TOTALE").setFontWeight("bold");
    totale(t, 3, base, base + 2);
    sh.getRange(t, 3).setFormula('=IFERROR(ROWS(UNIQUE(FILTER(' + chiave + ';' + MESE + '=$A$' + base + ';' +
      '((' + TR + '="t")+(' + TR + '="r")+(' + TR + '="p"))>0)));0)');
    sh.getRange(t, 1, 1, 2 + N_COL).setFontWeight("bold").setBackground("#dbeafe");
  });

  // totale generale: somma le righe TOTALE dei singoli mesi, poi ricalcola i rapporti
  const rGran = rUltimoMese, daR = rMese + 1, aR = rGran - 1;
  sh.getRange(rGran, 2).setValue("TOTALE GENERALE").setFontWeight("bold");
  [0, 1, 2, 3, 6, 7, 8, 13].forEach(n => {
    const L = smLettera_(3 + n);
    sh.getRange(rGran, 3 + n).setFormula(
      '=SUMIF($B$' + daR + ':$B$' + aR + ';"TOTALE";' + L + '$' + daR + ':' + L + '$' + aR + ')');
  });
  derivate(rGran, 3);
  sh.getRange(rGran, 1, 1, 2 + N_COL).setFontWeight("bold").setBackground("#93c5fd");

  const nR = rGran - rMese;
  stile(rMese + 1, nR, 3, 2 + N_COL);
  colori(rMese + 1, nR, 3);
}



// ------------------------------------------------------------
//  COMPILA €/km
//  Riempie la colonna €/km del registro: rimette i valori che
//  avevi personalizzato e mette il predefinito dove e' vuoto.
//  Non tocca mai una cella che hai gia' scritto.
//  Utile anche quando arrivano viaggi nuovi: basta rilanciarlo.
// ------------------------------------------------------------
function COMPILA_EURO_KM() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sh = s.getSheetByName(SM_FOGLIO_REGISTRO);
  if (!sh) { ui.alert("Registro non trovato: crealo prima dal menu."); return; }

  // trovo la riga delle intestazioni dell'elenco (Mese / Data)
  const testa = sh.getRange(1, 1, Math.min(sh.getLastRow(), 300), 2).getValues();
  let rTest = -1;
  for (let i = 0; i < testa.length; i++) {
    if (String(testa[i][0]).trim() === "Mese" && String(testa[i][1]).trim() === "Data") { rTest = i + 1; break; }
  }
  if (rTest < 0) { ui.alert("Non trovo l'elenco dei viaggi nel registro."); return; }

  const n = smCompilaEuroKm_(sh, rTest + 1, SM_ULTIMA_RIGA_REG, Date.now());
  ui.alert(n >= 0 ? ("€/km compilati: " + n + " viaggi.")
                  : "Il foglio non ha ancora finito di calcolare. Riprova tra un minuto.");
}

function smCompilaEuroKm_(sh, R1, U, t0) {
  const s = sh.getParent();
  const pred = Number(s.getSheetByName("PARAMETRI").getRange("B3").getValue()) || 0.9;
  let salvati = {};
  try {
    salvati = JSON.parse(PropertiesService.getDocumentProperties()
      .getProperty("SM_EUROKM_SALVATI") || "{}");
  } catch (e) {}

  const n = U - R1 + 1;
  let ids = [];
  // gli id vengono da una formula: se non sono pronti riprovo, ma senza superare il tempo
  while (true) {
    SpreadsheetApp.flush();
    ids = sh.getRange(R1, SM_COL_ID, n, 1).getValues();
    if (ids.some(r => String(r[0] || "").trim() !== "")) break;
    if (Date.now() - t0 > 240000) return -1;
    Utilities.sleep(2000);
  }

  const attuali = sh.getRange(R1, SM_COL_EUROKM, n, 1).getValues();
  let scritti = 0;
  const out = ids.map((r, i) => {
    const id = String(r[0] || "").trim();
    const v = attuali[i][0];
    if (!id) return [""];
    if (v !== "" && v != null) return [v];                       // gia' scritto: non lo tocco
    scritti++;
    return [salvati[id] !== undefined ? salvati[id] : pred];     // personalizzato o predefinito
  });
  sh.getRange(R1, SM_COL_EUROKM, n, 1).setValues(out);

  // costo giornaliero manuale: rimetto i valori che avevi scritto (solo nelle celle vuote)
  let salvatiCG = {};
  try {
    salvatiCG = JSON.parse(PropertiesService.getDocumentProperties()
      .getProperty("SM_COSTOG_SALVATI") || "{}");
  } catch (e) {}
  if (Object.keys(salvatiCG).length) {
    const attCG = sh.getRange(R1, SM_COL_COSTOG_MAN, n, 1).getValues();
    const outCG = ids.map((r, i) => {
      const id = String(r[0] || "").trim();
      const v = attCG[i][0];
      if (v !== "" && v != null) return [v];
      return [id && salvatiCG[id] !== undefined ? salvatiCG[id] : ""];
    });
    sh.getRange(R1, SM_COL_COSTOG_MAN, n, 1).setValues(outCG);
  }
  // rimessi a posto: svuoto la memoria, cosi' un valore che poi cancelli tu non ritorna
  PropertiesService.getDocumentProperties()
    .setProperty("SM_EUROKM_SALVATI", "{}").setProperty("SM_COSTOG_SALVATI", "{}");
  return scritti;
}


// ============================================================
//  REGISTRO GIORNO
//  Lo stesso confronto teorico/effettivo del registro del mese,
//  ma una riga per ogni giornata, senza dividere terra e rifiuti.
//  Legge i dati dal REGISTRO_MESE, quindi i numeri coincidono.
//  Per aggiungere le giornate nuove basta rilanciare il comando.
// ============================================================
const SG_FOGLIO = "REGISTRO_GIORNO";
const SG_RIGHE = 400;      // giornate lavorative preparate: oltre un anno

function CREA_REGISTRO_GIORNO() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rm = s.getSheetByName(SM_FOGLIO_REGISTRO);
  if (!rm) { ui.alert("Crea prima il registro del mese."); return; }
  smAssicuraParametri_(s);

  // dove parte l'elenco viaggi nel registro del mese (intestazioni Mese / Data)
  const testa = rm.getRange(1, 1, Math.min(rm.getLastRow(), 300), 2).getValues();
  let rTest = -1;
  for (let i = 0; i < testa.length; i++) {
    if (String(testa[i][0]).trim() === "Mese" && String(testa[i][1]).trim() === "Data") { rTest = i + 1; break; }
  }
  if (rTest < 0) { ui.alert("Non trovo l'elenco dei viaggi nel registro del mese."); return; }
  const R1 = rTest + 1, U = SM_ULTIMA_RIGA_REG;

  // Le giornate NON sono piu' scritte dallo script: le elenca una formula che legge
  // il registro del mese, quindi i giorni nuovi (ottobre, novembre...) compaiono da soli.
  // Preparo le righe per SG_RIGHE giornate lavorative: oltre un anno.
  const N = SG_RIGHE;

  let sh = s.getSheetByName(SG_FOGLIO);
  const pos = sh ? sh.getIndex() - 1 : 1;
  if (sh) s.deleteSheet(sh);
  sh = s.insertSheet(SG_FOGLIO, pos);
  sh.setTabColor("#6d28d9");
  if (sh.getMaxColumns() < 27) sh.insertColumnsAfter(sh.getMaxColumns(), 27 - sh.getMaxColumns());

  // intervalli del registro del mese
  const X = (L) => SM_FOGLIO_REGISTRO + "!$" + L + "$" + R1 + ":$" + L + "$" + U;
  const DATE = X('B'), TR = X('G'), FISSO = X('U'), KM = X('K'),
        TOT = X('W'), RIC = X('Z'), CHG = X('Y');     // V/Y: costo e ricavo del SINGOLO viaggio (diviso giusto tra terra, rifiuti e pianale)

  // ------------------------------------------------------------------
  //  DISPOSIZIONE
  //  riga 1 titolo | riga 3 Dal/al | riga 4 TOTALE (viola) | riga 5 intestazioni | dalla 6 i giorni
  //  A Data  B Giorno  C Giornate  D Km  E Costo giornaliero  F €/km medio  G Costo km
  //  H Totale costo  I Ricavo effettivo  J Margine  K Margine %  L Margine per giornata
  //  M Margine TERRA  N Margine RIFIUTI  O Margine PIANALE
  //  P Utile atteso  Q Ricavo teorico  R Differenza
  //  nascoste:  T giornate terra  U giornate rifiuti  V giornate pianale
  //             W margine terra €  X margine rifiuti €  Y margine pianale €
  //             Z 1 = giorno da contare nel TOTALE (visibile coi filtri E dentro il periodo Dal/al)
  // ------------------------------------------------------------------
  const rT = 4, rH = 5, r0 = 6;
  const rUlt = r0 + N - 1;
  const NV = 18;                                   // colonne visibili A:R

  sh.getRange("A1:R1").merge().setValue("REGISTRO GIORNO  —  confronto teorico / effettivo, giorno per giorno")
    .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground("#6d28d9").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);
  sh.getRange("A2").setValue("Si aggiorna da solo con i giorni nuovi.")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);

  // ---------- periodo del TOTALE ----------
  const oggi = new Date();
  const cellaData = (a1, v) => sh.getRange(a1).setValue(v).setNumberFormat("dd/mm/yyyy")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange("A3").setValue("Dal").setFontWeight("bold").setHorizontalAlignment("right");
  cellaData("B3", new Date(oggi.getFullYear(), 0, 1));
  sh.getRange("C3").setValue("al").setFontWeight("bold").setHorizontalAlignment("right");
  cellaData("D3", oggi);
  sh.getRange("E3").setValue("← periodo del TOTALE viola. Il totale rispetta anche i filtri sulle intestazioni.")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);

  const titoli = ["Data", "Giorno", "Giornate\ncamion", "Km", "Costo\ngiornaliero", "€/km\nmedio", "Costo km",
    "TOTALE\nCOSTO", "RICAVO\nEFFETTIVO", "MARGINE", "MARGINE\n%", "MARGINE REALE\nper giornata",
    "MARGINE REALE\nper giornata\nTERRA", "MARGINE REALE\nper giornata\nRIFIUTI", "MARGINE REALE\nper giornata\nPIANALE",
    "Utile atteso\nper giornata", "RICAVO\nTEORICO", "DIFFERENZA"];
  smIntestazioni_(sh, rH, 1, titoli);
  sh.getRange(rH, 1, 1, titoli.length).setBackground("#0b192c").setFontColor("#FFFFFF");
  sh.setRowHeight(rH, 52);

  // ---------- una riga per giornata ----------
  sh.getRange(r0, 1, N, 1).setNumberFormat("dd/mm/yyyy");
  sh.getRange(r0, 1).setFormula(
    '=IFERROR(SORT(UNIQUE(FILTER(ARRAYFORMULA(INT(' + DATE + '));' + DATE + '<>"")));"")');

  // utili attesi: terra B4, rifiuti B5, pianale B9
  const P4 = 'PARAMETRI!$B$4', P5 = 'PARAMETRI!$B$5', P9 = 'PARAMETRI!$B$9';
  const righe = [];
  for (let i = 0; i < N; i++) {
    const r = r0 + i;
    const crit = DATE + ';">="&$A' + r + ';' + DATE + ';"<"&($A' + r + '+1)';
    const cond = DATE + '>=$A' + r + ';' + DATE + '<$A' + r + '+1';
    const giornate = (extra, extraCond) =>
      '=IF(COUNTIFS(' + crit + extra + ')=0;0;COUNTUNIQUE(FILTER(' + CHG + ';' + cond + extraCond + ')))';
    const margine = (tipo) =>
      '=SUMIFS(' + RIC + ';' + crit + ';' + TR + ';"' + tipo + '")-SUMIFS(' + TOT + ';' + crit + ';' + TR + ';"' + tipo + '")';
    const riga = [
      '=PROPER(TEXT($A' + r + ';"dddd"))',                                   // B giorno della settimana
      giornate('', ''),                                                      // C giornate
      '=SUMIFS(' + KM + ';' + crit + ')',                                    // D km
      '=IF(WEEKDAY($A' + r + ';2)>=6;PARAMETRI!$B$8;PARAMETRI!$B$2)',       // E costo giornaliero: 100 nel weekend
      '=IFERROR(G' + r + '/D' + r + ';"")',                                  // F €/km medio
      '=SUMIFS(' + TOT + ';' + crit + ')-SUMIFS(' + FISSO + ';' + crit + ')', // G costo km
      '=SUMIFS(' + TOT + ';' + crit + ')',                                   // H totale costo
      '=SUMIFS(' + RIC + ';' + crit + ')',                                   // I ricavo effettivo
      '=I' + r + '-H' + r,                                                   // J margine
      '=IFERROR(J' + r + '/I' + r + ';"")',                                  // K margine %
      '=IFERROR(J' + r + '/C' + r + ';"")',                                  // L margine per giornata
      '=IFERROR(W' + r + '/T' + r + ';"")',                                  // M terra per giornata
      '=IFERROR(X' + r + '/U' + r + ';"")',                                  // N rifiuti per giornata
      '=IFERROR(Y' + r + '/V' + r + ';"")',                                  // O pianale per giornata
      // P utile atteso per giornata: utile atteso del giorno (teorico - costo) / giornate reali
      '=IFERROR((Q' + r + '-H' + r + ')/C' + r + ';"")',
      '=H' + r + '+T' + r + '*' + P4 + '+U' + r + '*' + P5 + '+V' + r + '*' + P9,   // Q teorico
      '=I' + r + '-Q' + r,                                                   // R differenza
      '',                                                                    // S (vuota)
      giornate(';' + TR + ';"t"', ';' + TR + '="t"'),                        // T giornate terra
      giornate(';' + TR + ';"r"', ';' + TR + '="r"'),                        // U giornate rifiuti
      giornate(';' + TR + ';"p"', ';' + TR + '="p"'),                        // V giornate pianale
      margine('t'),                                                          // W margine terra €
      margine('r'),                                                          // X margine rifiuti €
      margine('p')                                                           // Y margine pianale €
    ];
    // riga senza data -> vuota (e Google non calcola il resto: IF salta il ramo non usato)
    righe.push(riga.map(fx => fx === '' ? '' : '=IF($A' + r + '="";"";' + fx.substring(1) + ')'));
  }
  sh.getRange(r0, 2, N, 24).setFormulas(righe);
  const conta = [];
  for (let i = 0; i < N; i++) {
    const r = r0 + i;
    // SUBTOTAL(103) vale 0 se il filtro nasconde la riga; poi il controllo sul periodo
    conta.push(['=IF($A' + r + '="";0;SUBTOTAL(103;$A' + r + ')*($A' + r + '>=$B$3)*($A' + r + '<=$D$3))']);
  }
  sh.getRange(r0, 26, N, 1).setFormulas(conta);

  // ---------- TOTALE viola, SOPRA le intestazioni: conta solo i giorni visibili e nel periodo ----------
  const Z = '$Z$' + r0 + ':$Z$' + rUlt;
  const sub = (L) => '=SUMIF(' + Z + ';1;' + L + r0 + ':' + L + rUlt + ')';
  const t = rT;
  sh.getRange(t, 1).setValue("TOTALE");
  sh.getRange(t, 2, 1, 24).setFormulas([[
    '', sub('C'), sub('D'), '=IFERROR((H' + t + '-G' + t + ')/C' + t + ';"")', '=IFERROR(G' + t + '/D' + t + ';"")',
    sub('G'), sub('H'), sub('I'), '=I' + t + '-H' + t, '=IFERROR(J' + t + '/I' + t + ';"")',
    '=IFERROR(J' + t + '/C' + t + ';"")',
    '=IFERROR(W' + t + '/T' + t + ';"")', '=IFERROR(X' + t + '/U' + t + ';"")', '=IFERROR(Y' + t + '/V' + t + ';"")',
    '=IFERROR((Q' + t + '-H' + t + ')/C' + t + ';"")',
    sub('Q'), '=I' + t + '-Q' + t, '',
    sub('T'), sub('U'), sub('V'), sub('W'), sub('X'), sub('Y')
  ]]);
  sh.getRange(t, 1, 1, NV).setFontWeight("bold").setBackground("#ede9fe")
    .setBorder(true, true, true, true, true, true, "#6d28d9", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(t, 28);

  // ---------- formati (riga del totale + giorni) ----------
  const fmt = (c, n, f) => { sh.getRange(t, c, 1, n).setNumberFormat(f); sh.getRange(r0, c, N, n).setNumberFormat(f); };
  fmt(3, 2, "#,##0");
  fmt(5, 1, '#,##0.00 "€"');
  fmt(6, 1, '0.00 "€"');
  fmt(7, 4, '#,##0.00 "€"');
  fmt(11, 1, "0.0%");
  fmt(12, 7, '#,##0.00 "€"');
  sh.getRange(t, 1, 1, NV).setHorizontalAlignment("center");
  sh.getRange(r0, 1, N, NV).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  [[1,95],[2,90],[3,75],[4,75],[5,90],[6,75],[7,100],[8,105],[9,105],[10,100],[11,75],
   [12,110],[13,110],[14,110],[15,110],[16,105],[17,105],[18,105]].forEach(x => sh.setColumnWidth(x[0], x[1]));

  // ---------- colori (dalla riga del TOTALE fino all'ultimo giorno) ----------
  // Le intestazioni stanno in mezzo ma sono testo: le regole su numeri le ignorano.
  // INDIRECT: le regole di colore non possono leggere un altro foglio direttamente.
  const n = rUlt - t + 1;
  const z = (c) => [sh.getRange(t, c, n, 1)];
  const regola = (f, rosso, c) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(f)
    .setFontColor(rosso ? "#b91c1c" : "#15803d").setBold(true)
    .setBackground(rosso ? "#fee2e2" : "#dcfce7").setRanges(z(c)).build();
  const cond = (col, op, soglia) => '=AND(ISNUMBER($' + col + t + ');$' + col + t + op + soglia + ')';
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($A' + r0 + ');WEEKDAY($A' + r0 + ';2)>=6)')
      .setFontColor("#b45309").setBold(true).setBackground("#fff7ed")
      .setRanges([sh.getRange(r0, 2, N, 1), sh.getRange(r0, 5, N, 1)]).build(),
    // margine e differenza: segno
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#b91c1c").setBold(true)
      .setRanges([sh.getRange(t, 10, n, 1), sh.getRange(t, 18, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor("#15803d").setBold(true)
      .setRanges([sh.getRange(t, 10, n, 1), sh.getRange(t, 18, n, 1)]).build(),
    // margine %: soglia di PARAMETRI (30%)
    regola(cond('K', '<', 'INDIRECT("PARAMETRI!B7")'), true, 11),
    regola(cond('K', '>=', 'INDIRECT("PARAMETRI!B7")'), false, 11),
    // margine per giornata: contro l'utile atteso della stessa riga
    regola('=AND(ISNUMBER($L' + t + ');ISNUMBER($P' + t + ');$L' + t + '<$P' + t + ')', true, 12),
    regola('=AND(ISNUMBER($L' + t + ');ISNUMBER($P' + t + ');$L' + t + '>=$P' + t + ')', false, 12),
    // terra B4 | rifiuti B5 | pianale B9
    regola(cond('M', '<', 'INDIRECT("PARAMETRI!B4")'), true, 13),
    regola(cond('M', '>=', 'INDIRECT("PARAMETRI!B4")'), false, 13),
    regola(cond('N', '<', 'INDIRECT("PARAMETRI!B5")'), true, 14),
    regola(cond('N', '>=', 'INDIRECT("PARAMETRI!B5")'), false, 14),
    regola(cond('O', '<', 'INDIRECT("PARAMETRI!B9")'), true, 15),
    regola(cond('O', '>=', 'INDIRECT("PARAMETRI!B9")'), false, 15)
  ]);

  sh.hideColumns(19, 8);                        // S:Z appoggi di calcolo
  sh.setFrozenRows(rH);                         // titolo, TOTALE e intestazioni restano fermi
  sh.getRange(rH, 1, rUlt - rH + 1, NV).createFilter();   // il TOTALE sta sopra, fuori dal filtro

  SpreadsheetApp.flush();
  s.setActiveSheet(sh);
  ui.alert("REGISTRO GIORNO creato.\n\n" +
    "Si aggiorna da solo: i giorni nuovi compaiono man mano che arrivano i viaggi.\n" +
    "Non serve piu' rilanciare il comando.");
}
function smLettera_(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}


// ============================================================
//  SCHEDA CAMION
//  A-G  viaggi        H-I  teorico e differenza del GIORNO
//  K-N  rifornimenti  P-Q  spese      S-T  riepilogo
//  V-AF confronto del mese     AH in poi: spazio libero tuo
// ============================================================


// ------------------------------------------------------------
//  Confronto del MESE, diviso tra terra e rifiuti (V:AF)
// ------------------------------------------------------------


function smLettera_(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}


// ============================================================
//  SCHEDA CAMION
//  A-H  viaggi (con il tipo t/r accanto alla tratta)
//  I    teorico del giorno    J  differenza    K  margine reale
//  M-P  rifornimenti   R-S  spese   U-V  riepilogo
//  X-AI confronto del mese    AK in poi: spazio libero tuo
// ============================================================
function smCreaSchedaCamion_(s, mezzo) {
  const sh = s.insertSheet(mezzo.targa);
  if (sh.getMaxColumns() < 48) sh.insertColumnsAfter(sh.getMaxColumns(), 48 - sh.getMaxColumns());
  sh.setTabColor(SM_BLU);
  smTestata_(sh, mezzo, SM_BLU);

  const dal = '"&TEXT($F$2;"yyyy-mm-dd")&"';
  const al  = '"&TEXT($H$2+1;"yyyy-mm-dd")&"';
  const ULT = 205;

  // ---- VIAGGI (A:H) — il tipo t/r arriva dalla colonna AA del registro
  smBanda_(sh, "A4:K4", "VIAGGI DEL MESE", SM_BLU);
  smIntestazioni_(sh, 5, 1, ["Data", "Autista", "Tratta", "T/R", "Km", "Tonn.", "Ore", "Ricavo €",
                             "Teorico\ndel giorno", "Differenza\ndel giorno", "Margine reale\ndel giorno"]);
  sh.getRange(SM_RIGA_DATI, 1).setFormula(
    // ricavo = colonna AD di REGISTRO_VIAGGI (Totale_Ricavi + Totale_Ricavi_Manuale)
    '=IFERROR(QUERY(REGISTRO_VIAGGI!$A$2:$AD;' +
    '"select B,C,S,AA,H,L,M,AD where E = \'" & $B$2 & "\' ' +
    'and B >= date \'' + dal + '\' and B < date \'' + al + '\' order by B";0);"")'
  );

  smFormuleGiornoScheda_(sh, ULT, "I");   // I teorico del giorno (J e K piu' sotto usano lo stesso utile)
  const utile = smUtileScheda_(ULT);
  // J = quello che ha incassato davvero MENO quello che avrebbe dovuto
  sh.getRange(SM_RIGA_DATI, 10).setFormula(
    '=ARRAYFORMULA(IF($I$6:$I$' + ULT + '="";"";' +
    'SUMIF($A$6:$A$' + ULT + ';$A$6:$A$' + ULT + ';$H$6:$H$' + ULT + ')-$I$6:$I$' + ULT + '))'
  );
  smFormuleGiornoScheda_(sh, ULT, "K");   // K margine reale = differenza + utile atteso

  // ---- RIFORNIMENTI (M:P)
  smBanda_(sh, "M4:P4", "RIFORNIMENTI", "#0369a1");
  smIntestazioni_(sh, 5, 13, ["Data", "Litri", "Costo €", "Km"]);
  sh.getRange(SM_RIGA_DATI, 13).setFormula(
    '=IFERROR(QUERY(REGISTRO_RIFORNIMENTI!$A$2:$G;' +
    '"select B,D,E,F where C = \'" & $B$2 & "\' ' +
    'and B >= date \'' + dal + '\' and B < date \'' + al + '\' order by B";0);"")'
  );

  // ---- SPESE (R:S)
  smBanda_(sh, "R4:S4", "SPESE DEL MESE", "#b45309");
  smIntestazioni_(sh, 5, 18, ["Voce", "Importo €"]);
  smSpese_(sh, 18);

  // ---- RIEPILOGO (U:V)
  smBanda_(sh, "U4:V4", "RIEPILOGO", "#15803d");
  const voci = [
    ["Viaggi",            '=IF(A6="";0;COUNTA(A6:A' + ULT + '))', "#,##0"],
    ["Km percorsi",       '=SUM(E6:E' + ULT + ')', "#,##0"],
    ["Tonnellate",        '=SUM(F6:F' + ULT + ')', "#,##0.0"],
    ["Ore",               '=SUM(G6:G' + ULT + ')', "#,##0.0"],
    ["", "", ""],
    ["RICAVI",            '=SUM(H6:H' + ULT + ')', '#,##0.00 "€"'],
    ["Gasolio",           '=SUM(O6:O' + ULT + ')', '#,##0.00 "€"'],
    ["Spese",             '=SUM(S6:S21)', '#,##0.00 "€"'],
    ["MARGINE",           '=V10-V11-V12', '#,##0.00 "€"'],
    ["", "", ""],
    ["Litri",             '=SUM(N6:N' + ULT + ')', "#,##0.0"],
    ["Consumo L/100km",   '=IFERROR(V15/V6*100;"")', "0.0"],
    ["Costo totale €/km", '=IFERROR((V11+V12)/V6;"")', '0.00 "€"'],
    ["Ricavo €/km",       '=IFERROR(V10/V6;"")', '0.00 "€"']
  ];
  smScriviRiepilogo_(sh, voci, 21);

  smConfronto_(sh);

  smFormatta_(sh, [
    { col: 1, fmt: "dd/mm/yyyy", w: 85 }, { col: 2, w: 155 }, { col: 3, w: 145 },
    { col: 4, w: 45 },
    { col: 5, fmt: "#,##0", w: 60 }, { col: 6, fmt: "#,##0.0", w: 60 },
    { col: 7, fmt: "#,##0.0", w: 55 }, { col: 8, fmt: '#,##0.00 "€"', w: 90 },
    { col: 9, fmt: '#,##0.00 "€"', w: 95 }, { col: 10, fmt: '#,##0.00 "€"', w: 95 },
    { col: 11, fmt: '#,##0.00 "€"', w: 100 },
    { col: 12, w: 25 },
    { col: 13, fmt: "dd/mm/yyyy", w: 85 }, { col: 14, fmt: "#,##0.0", w: 70 },
    { col: 15, fmt: '#,##0.00 "€"', w: 85 }, { col: 16, fmt: "#,##0", w: 80 },
    { col: 17, w: 25 },
    { col: 18, w: 175 }, { col: 19, fmt: '#,##0.00 "€"', w: 95 },
    { col: 20, w: 25 },
    { col: 21, w: 145 }, { col: 22, w: 110 },
    { col: 23, w: 25 }
  ]);
  sh.getRange(SM_RIGA_DATI, 4, ULT - 5, 1).setHorizontalAlignment("center");

  // rosso sotto zero, verde sopra, sulle due colonne di differenza
  [10, 11].forEach(c => {
    const zona = [sh.getRange(SM_RIGA_DATI, c, ULT - 5, 1)];
    sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
        .setFontColor("#b91c1c").setBold(true).setRanges(zona).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setFontColor("#15803d").setRanges(zona).build()
    ]));
  });

  sh.setFrozenRows(5);
  sh.getRange("AK4").setValue("← da qui in poi lo spazio e' tuo: quello che scrivi non viene mai sovrascritto")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);
}


// ------------------------------------------------------------
//  Utile atteso del giorno nelle schede camion, dal tipo in colonna D:
//  terra B4, rifiuti B5, PIANALE B9
// ------------------------------------------------------------
function smUtileScheda_(ULT) {
  const D = '$D$6:$D$' + ULT;
  return 'IF(' + D + '="r";PARAMETRI!$B$5;IF(' + D + '="t";PARAMETRI!$B$4;IF(' + D + '="p";PARAMETRI!$B$9;0)))';
}
// Scrive le formule del giorno della scheda camion che dipendono dall'utile atteso
//   "I" teorico del giorno   "K" margine reale del giorno
function smFormuleGiornoScheda_(sh, ULT, quale) {
  const utile = smUtileScheda_(ULT);
  if (quale === "I") {
    sh.getRange(SM_RIGA_DATI, 9).setFormula(
      '=ARRAYFORMULA(IF($A$6:$A$' + ULT + '="";"";' +
      'IF($A$6:$A$' + ULT + '<>$A$5:$A$' + (ULT - 1) + ';' +
      'PARAMETRI!$B$2' +
      '+SUMIF($A$6:$A$' + ULT + ';$A$6:$A$' + ULT + ';$E$6:$E$' + ULT + ')*PARAMETRI!$B$3' +
      '+' + utile + ';"")))');
  } else {
    sh.getRange(SM_RIGA_DATI, 11).setFormula(
      '=ARRAYFORMULA(IF($J$6:$J$' + ULT + '="";"";$J$6:$J$' + ULT + '+' + utile + '))');
  }
}

// ------------------------------------------------------------
//  Confronto del MESE, diviso tra terra, rifiuti e pianale (X:AI)
// ------------------------------------------------------------
function smConfronto_(sh) {
  smBanda_(sh, "X4:AI4", "CONFRONTO DEL MESE  —  TEORICO / EFFETTIVO", "#6d28d9");
  smIntestazioni_(sh, 5, 24, [
    "Tipo", "Giorni", "Km", "Costo\ngiornaliero", "Costo\n€/km", "Costo km",
    "TOTALE\nCOSTO", "Utile\ngiornaliero", "RICAVO\nTEORICO",
    "RICAVO\nEFFETTIVO", "DIFFERENZA", "MARGINE REALE\nper giorno"
  ]);

  const R = "REGISTRO_VIAGGI!";
  [{ r: 6, nome: "Terra", cod: "t", par: "$B$4" },
   { r: 7, nome: "Rifiuti", cod: "r", par: "$B$5" },
   { r: 8, nome: "Pianale", cod: "p", par: "$B$9" }].forEach(x => {
    const cond = R + '$E$2:$E=$B$2;' + R + '$AA$2:$AA="' + x.cod + '";' +
                 R + '$B$2:$B>=$F$2;' + R + '$B$2:$B<$H$2+1';
    const crit = R + '$E:$E;$B$2;' + R + '$AA:$AA;"' + x.cod + '";' +
                 R + '$B:$B;">="&$F$2;' + R + '$B:$B;"<"&$H$2+1';

    sh.getRange(x.r, 24).setValue(x.nome).setFontWeight("bold");
    sh.getRange(x.r, 25).setFormula(
      '=IF(COUNTIFS(' + crit + ')=0;0;COUNTUNIQUE(FILTER(INT(' + R + '$B$2:$B);' + cond + ')))');
    sh.getRange(x.r, 26).setFormula('=SUMIFS(' + R + '$H:$H;' + crit + ')');
    sh.getRange(x.r, 27).setFormula('=PARAMETRI!$B$2');
    sh.getRange(x.r, 28).setFormula('=PARAMETRI!$B$3');
    sh.getRange(x.r, 29).setFormula('=Z' + x.r + '*AB' + x.r);
    sh.getRange(x.r, 30).setFormula('=Y' + x.r + '*AA' + x.r + '+AC' + x.r);
    sh.getRange(x.r, 31).setFormula('=PARAMETRI!' + x.par);
    sh.getRange(x.r, 32).setFormula('=AD' + x.r + '+Y' + x.r + '*AE' + x.r);
    sh.getRange(x.r, 33).setFormula('=SUMIFS(' + R + '$AD:$AD;' + crit + ')');   // calcolato + manuale
    sh.getRange(x.r, 34).setFormula('=AG' + x.r + '-AF' + x.r);
    // (ricavo effettivo - totale costo) diviso i giorni lavorati
    sh.getRange(x.r, 35).setFormula('=IFERROR((AG' + x.r + '-AD' + x.r + ')/Y' + x.r + ';"")');
  });

  sh.getRange(9, 24).setValue("TOTALE").setFontWeight("bold");
  [25, 26, 29, 30, 32, 33, 34].forEach(c => {
    const L = smLettera_(c);
    sh.getRange(9, c).setFormula('=SUM(' + L + '6:' + L + '8)').setFontWeight("bold");
  });
  // totale: (ricavo effettivo - costo totale) diviso i giorni complessivi
  sh.getRange(9, 35).setFormula('=IFERROR((AG9-AD9)/Y9;"")').setFontWeight("bold");
  // margine atteso medio, pesato sui giorni: e' la soglia della riga totale
  sh.getRange(9, 31).setFormula('=IFERROR((Y6*AE6+Y7*AE7+Y8*AE8)/Y9;"")').setFontWeight("bold");

  sh.getRange(6, 25, 4, 2).setNumberFormat("#,##0");
  sh.getRange(6, 27, 4, 1).setNumberFormat('#,##0.00 "€"');
  sh.getRange(6, 28, 4, 1).setNumberFormat('0.00 "€"');
  sh.getRange(6, 29, 4, 7).setNumberFormat('#,##0.00 "€"');
  sh.getRange(6, 24, 4, 12).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(6, 32, 4, 3).setBackground("#f5f3ff");
  sh.getRange(6, 35, 4, 1).setBackground("#ecfdf5");
  sh.getRange(9, 24, 1, 12).setBackground("#ede9fe");

  // DIFFERENZA: rosso sotto zero, verde sopra
  const zonaDiff = [sh.getRange(6, 34, 4, 1)];
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setFontColor("#b91c1c").setBold(true).setRanges(zonaDiff).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor("#15803d").setBold(true).setRanges(zonaDiff).build()
  ]));

  // MARGINE REALE PER GIORNO: il confronto non e' con lo zero ma con
  // l'utile atteso della riga (200 terra, 300 rifiuti, media pesata sul totale).
  // Uso AE della stessa riga, cosi' se cambi i PARAMETRI le soglie si spostano da sole.
  const zonaMarg = [sh.getRange(6, 35, 4, 1)];
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($AI6);ISNUMBER($AE6);$AI6<$AE6)')
      .setFontColor("#b91c1c").setBold(true).setBackground("#fee2e2")
      .setRanges(zonaMarg).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($AI6);ISNUMBER($AE6);$AI6>=$AE6)')
      .setFontColor("#15803d").setBold(true).setBackground("#dcfce7")
      .setRanges(zonaMarg).build()
  ]));

  [[24,90],[25,70],[26,80],[27,95],[28,80],[29,100],[30,110],[31,95],[32,110],[33,110],[34,110],[35,120]]
    .forEach(c => sh.setColumnWidth(c[0], c[1]));
}

function smLettera_(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

// ------------------------------------------------------------
//  SCHEDA RIMORCHIO  (niente rifornimenti, niente ricavi)
// ------------------------------------------------------------
function smCreaSchedaRimorchio_(s, mezzo) {
  const sh = s.insertSheet(mezzo.targa);
  sh.setTabColor(SM_ARANCIO);
  smTestata_(sh, mezzo, SM_ARANCIO);

  const dal = '"&TEXT($F$2;"yyyy-mm-dd")&"';
  const al  = '"&TEXT($H$2+1;"yyyy-mm-dd")&"';

  // ---- VIAGGI (colonne A:F) - senza ricavo: il guadagno e' del camion
  smBanda_(sh, "A4:F4", "VIAGGI DEL MESE", SM_ARANCIO);
  smIntestazioni_(sh, 5, 1, ["Data", "Autista", "Camion", "Tratta", "Km", "Tonn."]);
  sh.getRange(SM_RIGA_DATI, 1).setFormula(
    '=IFERROR(QUERY(REGISTRO_VIAGGI!$A$2:$Z;' +
    '"select B,C,E,S,H,L where K = \'" & $B$2 & "\' ' +
    'and B >= date \'' + dal + '\' and B < date \'' + al + '\' order by B";0);"")'
  );

  // ---- SPESE (colonne H:I)
  smBanda_(sh, "H4:I4", "SPESE DEL MESE", "#b45309");
  smIntestazioni_(sh, 5, 8, ["Voce", "Importo €"]);
  smSpese_(sh, 8);

  // ---- RIEPILOGO (colonne K:L)
  smBanda_(sh, "K4:L4", "RIEPILOGO", "#15803d");
  const voci = [
    ["Viaggi",              '=IF(A' + SM_RIGA_DATI + '="";0;COUNTA(A' + SM_RIGA_DATI + ':A))', "#,##0"],
    ["Giorni di utilizzo",  '=IF(A' + SM_RIGA_DATI + '="";0;COUNTUNIQUE(A' + SM_RIGA_DATI + ':A))', "#,##0"],
    ["Km",                  '=SUM(E' + SM_RIGA_DATI + ':E)', "#,##0"],
    ["Tonnellate",          '=SUM(F' + SM_RIGA_DATI + ':F)', "#,##0.0"],
    ["", "", ""],
    ["SPESE DEL MESE",      '=SUM(I6:I21)', '#,##0.00 "€"'],
    ["", "", ""],
    ["Costo per viaggio",   '=IFERROR(L10/L5;"")', '#,##0.00 "€"'],
    ["Costo per giorno",    '=IFERROR(L10/L6;"")', '#,##0.00 "€"'],
    ["Costo per tonnellata",'=IFERROR(L10/L8;"")', '#,##0.00 "€"']
  ];
  smScriviRiepilogo_(sh, voci, 11);

  smFormatta_(sh, [
    { col: 1, fmt: "dd/mm/yyyy", w: 85 }, { col: 2, w: 165 }, { col: 3, w: 135 },
    { col: 4, w: 150 }, { col: 5, fmt: "#,##0", w: 65 }, { col: 6, fmt: "#,##0.0", w: 65 },
    { col: 7, w: 25 },
    { col: 8, w: 175 }, { col: 9, fmt: '#,##0.00 "€"', w: 95 },
    { col: 10, w: 25 },
    { col: 11, w: 160 }, { col: 12, w: 110 }
  ]);

  sh.setFrozenRows(5);
  sh.getRange("N4").setValue("← da qui in poi lo spazio e' tuo")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);
}

// ------------------------------------------------------------
function smScriviRiepilogo_(sh, voci, colonnaEtichetta) {
  let r = 5;
  voci.forEach(v => {
    if (v[0] === "") { r++; return; }
    const grassetto = (v[0] === "MARGINE" || v[0] === "RICAVI" || v[0] === "SPESE DEL MESE");
    sh.getRange(r, colonnaEtichetta).setValue(v[0])
      .setFontWeight(grassetto ? "bold" : "normal").setFontSize(10);
    const cella = sh.getRange(r, colonnaEtichetta + 1);
    cella.setFormula(v[1]).setNumberFormat(v[2])
      .setFontWeight(grassetto ? "bold" : "normal")
      .setHorizontalAlignment("right")
      .setBackground(grassetto ? "#dcfce7" : null);
    r++;
  });
}

function smFormatta_(sh, colonne) {
  colonne.forEach(c => {
    sh.setColumnWidth(c.col, c.w);
    if (c.fmt) sh.getRange(SM_RIGA_DATI, c.col, 400, 1).setNumberFormat(c.fmt);
  });
}


// ------------------------------------------------------------
//  Aggiunge in PARAMETRI le righe introdotte dopo la prima
//  installazione, se mancano. Non tocca quelle gia' presenti.
// ------------------------------------------------------------
function smAssicuraParametri_(s) {
  const par = s.getSheetByName("PARAMETRI");
  if (!par) return;
  const righe = [
    [6, "Costo giornaliero LAVORI NOTTE", 180, "fisso dei lavori notturni (tratte con \"notte\" nel nome) quando quel giorno il camion ha fatto anche altri lavori", '#,##0.00 "€"'],
    [7, "Margine % minimo", 0.3, "sotto questa percentuale il margine % del registro diventa rosso", "0%"],
    [9, "Utile giornaliero PIANALE", "=B4", "margine atteso per giorno sui lavori con il pianale (di base uguale alla terra: scrivi un numero per cambiarlo)", '#,##0.00 "€"'],
    [8, "Costo giornaliero SABATO e DOMENICA", 100, "costo fisso nel fine settimana: sostituisce sia il costo giornaliero sia quello dell'appalti notte (tutto costo autista)", '#,##0.00 "€"'],
    [10, "Costo giornaliero AUTISTA", 202, "parte del costo giornaliero per l'autista (dal lunedi' al venerdi')", '#,##0.00 "€"'],
    [11, "Costo giornaliero FISSI MEZZO", 118, "parte del costo giornaliero per i costi fissi del mezzo (dal lunedi' al venerdi')", '#,##0.00 "€"']
  ];
  let nuovaDivisione = false;
  righe.forEach(x => {
    if (String(par.getRange(x[0], 1).getValue()).trim() !== "") return;
    par.getRange(x[0], 1, 1, 3).setValues([[x[1], x[2], x[3]]]);
    par.getRange(x[0], 2).setNumberFormat(x[4])
      .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center");
    if (x[0] === 10) nuovaDivisione = true;
  });
  // il costo giornaliero (B2) diventa autista + fissi: si cambiano le righe 10 e 11
  if (nuovaDivisione) {
    par.getRange("B2").setFormula("=B10+B11");
    par.getRange("C2").setValue("somma di autista (riga 10) + fissi del mezzo (riga 11): cambia quelle due righe");
  }
}

// ============================================================
//  REGISTRO MEZZO
//  Una riga per ogni CAMION e MESE: conti del modello (costo
//  giornaliero + €/km) affiancati ai conti REALI (gasolio dei
//  rifornimenti + spese di COSTI_MEZZI).
//  Si allunga da solo: un camion compare nel mese in cui ha
//  fatto almeno un viaggio.
// ============================================================
const SZ_FOGLIO = "REGISTRO_MEZZO";
const SZ_RIGHE = 600;      // righe camion-mese preparate: 22 camion per oltre 2 anni

function CREA_REGISTRO_MEZZO() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rm = s.getSheetByName(SM_FOGLIO_REGISTRO);
  if (!rm) { ui.alert("Crea prima il registro del mese."); return; }
  smAssicuraParametri_(s);

  // dove parte l'elenco viaggi nel registro del mese
  const testa = rm.getRange(1, 1, Math.min(rm.getLastRow(), 300), 2).getValues();
  let rTest = -1;
  for (let i = 0; i < testa.length; i++) {
    if (String(testa[i][0]).trim() === "Mese" && String(testa[i][1]).trim() === "Data") { rTest = i + 1; break; }
  }
  if (rTest < 0) { ui.alert("Non trovo l'elenco dei viaggi nel registro del mese."); return; }
  const R1 = rTest + 1, U = SM_ULTIMA_RIGA_REG, N = SZ_RIGHE;

  let sh = s.getSheetByName(SZ_FOGLIO);
  const pos = sh ? sh.getIndex() - 1 : 2;
  if (sh) s.deleteSheet(sh);
  sh = s.insertSheet(SZ_FOGLIO, pos);
  sh.setTabColor("#15803d");
  if (sh.getMaxColumns() < 37) sh.insertColumnsAfter(sh.getMaxColumns(), 37 - sh.getMaxColumns());

  // dal registro del mese
  const RM = (L) => SM_FOGLIO_REGISTRO + "!$" + L + "$" + R1 + ":$" + L + "$" + U;
  const RM_MESE = RM('A'), RM_DATA = RM('B'), RM_TARGA = RM('C'), RM_TR = RM('G'), RM_KM = RM('K'),
        RM_TOT = RM('W'), RM_RIC = RM('Z'), RM_CHG = RM('Y');     // W/Z: per singolo viaggio
  // dai rifornimenti: B data, C targa, D litri, E costo
  const RF_DATA = 'REGISTRO_RIFORNIMENTI!$B:$B', RF_TARGA = 'REGISTRO_RIFORNIMENTI!$C:$C',
        RF_LITRI = 'REGISTRO_RIFORNIMENTI!$D:$D', RF_COSTO = 'REGISTRO_RIFORNIMENTI!$E:$E';

  // ------------------------------------------------------------------
  //  COLONNE VISIBILI (A:W)
  //  A Targa  B Modello  C Tipo  D Mese  E Giorni terra  F Giorni rifiuti  G Giorni pianale  H Tot giorni
  //  I Km  J Litri  K Litri/km  L Km/litro  M Costo al km  N TOTALE COSTO  O RICAVO EFFETTIVO
  //  P Margine  Q % Margine/Ricavo  R Costo gasolio  S Spese  T Costo reale mese
  //  U Margine reale mese  V % Margine reale/Ricavo  W Costo reale al km
  //  COLONNE NASCOSTE: i nomi qui sotto, cosi' le lettere sono scritte in un posto solo
  // ------------------------------------------------------------------
  const NV = 23;                        // colonne visibili A:W
  const H = {
    chiave: 'Y',     // mese|targa (elenco automatico)
    mese: 'Z',       // "2026-09"
    targa: 'AA',
    dal: 'AB',       // primo giorno del mese
    al: 'AC',        // primo giorno del mese dopo
    conta: 'AD',     // 1 = riga da contare nel TOTALE (visibile coi filtri)
    kmTutti: 'AE',   // km del camion nella finestra, tutti i tipi
    ggTutti: 'AF',   // giornate del camion nella finestra, tutti i tipi
    da: 'AI',        // inizio finestra (incluso) = parte del mese dentro Dal/al
    a: 'AJ'          // fine finestra (esclusa)
  };
  const COD = '$AG$3', CRIT_TIPO = '$AH$3';   // tipo di lavoro scelto nella tendina
  const col = (L) => { let n = 0; for (const c of L) n = n * 26 + (c.charCodeAt(0) - 64); return n; };

  const rT = 4, rH = 5, r0 = 6, rUlt = r0 + N - 1;

  sh.getRange("A1:W1").merge().setValue("REGISTRO MEZZO  —  conti del modello e conti reali, camion per camion")
    .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground("#15803d").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);
  sh.getRange("A2").setValue("Si aggiorna da solo. Costo reale = gasolio dei rifornimenti + spese di COSTI_MEZZI.")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);

  // ---------- periodo e tipo di lavoro ----------
  const oggi = new Date();
  const cellaData = (a1, v) => sh.getRange(a1).setValue(v).setNumberFormat("dd/mm/yyyy")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange("A3").setValue("Dal").setFontWeight("bold").setHorizontalAlignment("right");
  cellaData("B3", new Date(oggi.getFullYear(), 0, 1));
  sh.getRange("C3").setValue("al").setFontWeight("bold").setHorizontalAlignment("right");
  cellaData("D3", oggi);
  sh.getRange("F3").setValue("Lavoro").setFontWeight("bold").setHorizontalAlignment("right");
  sh.getRange("G3").setValue("Tutti")
    .setBackground("#fff7ed").setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, SM_ARANCIO, SpreadsheetApp.BorderStyle.SOLID_THICK)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["Tutti", "Terra", "Rifiuti", "Pianale"], true).setAllowInvalid(false).build());
  sh.getRange("I3").setValue("← Dal/al: periodo esatto al giorno, per tutta la tabella e il TOTALE.  Lavoro: solo terra, rifiuti o pianale.")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);
  // AG3: codice del tipo scelto ("" = tutti)   AH3: lo stesso, come criterio per SUMIFS
  sh.getRange("AG3").setFormula('=IF($G$3="Terra";"t";IF($G$3="Rifiuti";"r";IF($G$3="Pianale";"p";"")))');
  sh.getRange("AH3").setFormula('=IF($AG$3="";"<>###";$AG$3)');

  const titoli = ["Targa", "Modello", "Tipo", "Mese", "Giorni\nTerra", "Giorni\nRifiuti", "Giorni\nPianale", "Tot\nGiorni",
    "Km\npercorsi", "Litri\nGasolio", "Litri/km", "Km/litro", "Costo\nal km", "TOTALE\nCOSTO", "RICAVO\nEFFETTIVO",
    "Margine", "% Margine/\nRicavo effettivo", "Costo\nGasolio", "Spese", "Costo reale\nmese",
    "Margine reale\nmese", "% Margine reale/\nRicavo effettivo", "Costo reale\nal km"];
  smIntestazioni_(sh, rH, 1, titoli);
  sh.getRange(rH, 1, 1, titoli.length).setBackground("#0b192c").setFontColor("#FFFFFF");
  sh.getRange(rH, 18, 1, 6).setBackground("#15803d");      // conti reali R:W
  sh.setRowHeight(rH, 52);

  // chiave: elenco automatico mese|targa, ordinato per mese e poi per targa.
  // Solo i camion con viaggi DENTRO il periodo Dal/al e del tipo di lavoro scelto.
  sh.getRange(r0, col(H.chiave)).setFormula(
    '=IFERROR(SORT(UNIQUE(FILTER(ARRAYFORMULA(' + RM_MESE + '&"|"&' + RM_TARGA + ');' + RM_TARGA + '<>"";' +
    RM_DATA + '>=$B$3;' + RM_DATA + '<$D$3+1;' +
    '((' + COD + '="")+(' + RM_TR + '=' + COD + '))>0)));"")');

  const se = (r, fx) => '=IF($' + H.chiave + r + '="";"";' + fx + ')';
  const visibili = [], appoggi = [];
  for (let i = 0; i < N; i++) {
    const r = r0 + i;
    const MESE = '$' + H.mese + r, TARGA = '$' + H.targa + r, DAL = '$' + H.dal + r, AL = '$' + H.al + r;
    const DA = '$' + H.da + r, A = '$' + H.a + r;
    const critBase = RM_DATA + ';">="&' + DA + ';' + RM_DATA + ';"<"&' + A + ';' + RM_TARGA + ';' + TARGA;
    const condBase = RM_DATA + '>=' + DA + ';' + RM_DATA + '<' + A + ';' + RM_TARGA + '=' + TARGA;
    // con il tipo di lavoro scelto nella tendina (Tutti = nessun limite)
    const crit = critBase + ';' + RM_TR + ';' + CRIT_TIPO;
    const cond = condBase + ';((' + COD + '="")+(' + RM_TR + '=' + COD + '))>0';
    const giornate = (tipo) => tipo
      ? 'IF(COUNTIFS(' + crit + ';' + RM_TR + ';"' + tipo + '")=0;0;COUNTUNIQUE(FILTER(' + RM_CHG + ';' + cond + ';' + RM_TR + '="' + tipo + '")))'
      : 'IF(COUNTIFS(' + crit + ')=0;0;COUNTUNIQUE(FILTER(' + RM_CHG + ';' + cond + ')))';
    // quote per ripartire i costi reali quando si sceglie un solo tipo di lavoro
    const quotaKm = 'IF(' + COD + '="";1;IFERROR(I' + r + '/$' + H.kmTutti + r + ';0))';
    const quotaGiorni = 'IF(' + COD + '="";1;IFERROR(H' + r + '/$' + H.ggTutti + r + ';0))';
    const rif = (c) => 'SUMIFS(' + c + ';' + RF_TARGA + ';' + TARGA + ';' + RF_DATA + ';">="&' + DA + ';' + RF_DATA + ';"<"&' + A + ')';
    // spese mensili: per un mese preso solo in parte, la quota dei giorni compresi nel periodo
    const quotaPeriodo = 'IFERROR((' + A + '-' + DA + ')/(' + AL + '-' + DAL + ');0)';
    visibili.push([
      se(r, TARGA),                                                             // A targa
      se(r, 'IFERROR(VLOOKUP(' + TARGA + ';ANAGRAFICA_MEZZI!$B:$D;3;FALSE);"")'),  // B modello
      se(r, 'IFERROR(VLOOKUP(' + TARGA + ';ANAGRAFICA_MEZZI!$B:$C;2;FALSE);"")'),  // C tipo
      se(r, 'TEXT(' + DAL + ';"mmmm yyyy")'),                                   // D mese
      se(r, giornate('t')),                                                     // E giorni terra
      se(r, giornate('r')),                                                     // F giorni rifiuti
      se(r, giornate('p')),                                                     // G giorni pianale
      se(r, giornate('')),                                                      // H tot giorni
      se(r, 'SUMIFS(' + RM_KM + ';' + crit + ')'),                              // I km
      se(r, rif(RF_LITRI) + '*' + quotaKm),                                    // J litri (quota per km)
      se(r, 'IFERROR(J' + r + '/I' + r + ';"")'),                               // K litri/km
      se(r, 'IFERROR(I' + r + '/J' + r + ';"")'),                               // L km/litro
      se(r, 'IFERROR(N' + r + '/I' + r + ';"")'),                               // M costo al km
      se(r, 'SUMIFS(' + RM_TOT + ';' + crit + ')'),                             // N totale costo
      se(r, 'SUMIFS(' + RM_RIC + ';' + crit + ')'),                             // O ricavo effettivo
      se(r, 'O' + r + '-N' + r),                                                // P margine
      se(r, 'IFERROR(P' + r + '/O' + r + ';"")'),                               // Q % margine
      se(r, rif(RF_COSTO) + '*' + quotaKm),                                    // R costo gasolio (quota per km)
      // S spese: SUMIFS e' tollerante sul mese (in COSTI_MEZZI puo' essere testo o data)
      se(r, 'SUMIFS(COSTI_MEZZI!$E:$E;COSTI_MEZZI!$B:$B;' + TARGA + ';COSTI_MEZZI!$A:$A;' + MESE + ')*' + quotaGiorni + '*' + quotaPeriodo),
      se(r, 'R' + r + '+S' + r),                                                // T costo reale
      se(r, 'O' + r + '-T' + r),                                                // U margine reale
      se(r, 'IFERROR(U' + r + '/O' + r + ';"")'),                               // V % margine reale
      se(r, 'IFERROR(T' + r + '/I' + r + ';"")')                                // W costo reale al km
    ]);
    appoggi.push([
      se(r, 'LEFT($' + H.chiave + r + ';7)'),                                   // Z mese "2026-09"
      se(r, 'MID($' + H.chiave + r + ';9;20)'),                                 // AA targa
      se(r, 'DATE(VALUE(LEFT(' + MESE + ';4));VALUE(RIGHT(' + MESE + ';2));1)'),   // AB primo giorno
      se(r, 'EDATE(' + DAL + ';1)'),                                            // AC mese dopo
      // AD: SUBTOTAL(103) vale 0 se il filtro nasconde la riga (il periodo e' gia' nell'elenco)
      '=IF($' + H.chiave + r + '="";0;SUBTOTAL(103;$A' + r + '))',
      se(r, 'SUMIFS(' + RM_KM + ';' + critBase + ')'),                          // AE km tutti i tipi
      se(r, 'IF(COUNTIFS(' + critBase + ')=0;0;COUNTUNIQUE(FILTER(' + RM_CHG + ';' + condBase + ')))'),  // AF giornate tutti i tipi
      '', '',                                                                   // AG, AH: usate solo in riga 3
      se(r, 'MAX(' + DAL + ';$B$3)'),                                           // AI inizio finestra (incluso)
      se(r, 'MIN(' + AL + ';$D$3+1)')                                           // AJ fine finestra (esclusa)
    ]);
  }
  sh.getRange(r0, 1, N, NV).setFormulas(visibili);
  sh.getRange(r0, col(H.mese), N, 11).setFormulas(appoggi);                 // Z:AJ
  sh.getRange(r0, col(H.mese), N, 1).setNumberFormat("@");

  // ---------- TOTALE sopra le intestazioni: solo le righe visibili ----------
  const CONTA = '$' + H.conta + '$' + r0 + ':$' + H.conta + '$' + rUlt;
  const sub = (L) => '=SUMIF(' + CONTA + ';1;' + L + r0 + ':' + L + rUlt + ')';
  const media = (L) => '=IFERROR(AVERAGEIF(' + CONTA + ';1;' + L + r0 + ':' + L + rUlt + ');"")';
  const t = rT;
  sh.getRange(t, 1).setFormula('="TOTALE"&IF($G$3="Tutti";"";" "&UPPER($G$3))');
  sh.getRange(t, 5, 1, 19).setFormulas([[
    sub('E'), sub('F'), sub('G'), sub('H'), sub('I'), sub('J'),
    media('K'), media('L'),                                   // litri/km e km/litro: MEDIA dei camion visibili
    '=IFERROR(N' + t + '/I' + t + ';"")',
    sub('N'), sub('O'), sub('P'), '=IFERROR(P' + t + '/O' + t + ';"")',
    sub('R'), sub('S'), sub('T'), sub('U'), '=IFERROR(U' + t + '/O' + t + ';"")',
    '=IFERROR(T' + t + '/I' + t + ';"")'
  ]]);
  sh.getRange(t, 1, 1, NV).setFontWeight("bold").setBackground("#dcfce7")
    .setBorder(true, true, true, true, true, true, "#15803d", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(t, 28);

  // ---------- formati ----------
  const fmt = (c, n, f) => { sh.getRange(t, c, 1, n).setNumberFormat(f); sh.getRange(r0, c, N, n).setNumberFormat(f); };
  fmt(5, 5, "#,##0");                 // E:I giorni e km
  fmt(10, 1, "#,##0.0");              // J litri
  fmt(11, 1, "0.000");                // K litri/km
  fmt(12, 1, "0.00");                 // L km/litro
  fmt(13, 1, '0.00 "€"');             // M costo al km
  fmt(14, 3, '#,##0.00 "€"');         // N:P
  fmt(17, 1, "0.0%");                 // Q
  fmt(18, 4, '#,##0.00 "€"');         // R:U
  fmt(22, 1, "0.0%");                 // V
  fmt(23, 1, '0.00 "€"');             // W
  sh.getRange(t, 1, 1, NV).setHorizontalAlignment("center");
  sh.getRange(r0, 1, N, NV).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r0, 18, N, 6).setBackground("#f0fdf4");      // conti reali R:W
  [[1,90],[2,140],[3,70],[4,110],[5,65],[6,65],[7,65],[8,60],[9,85],[10,85],[11,70],[12,70],[13,75],
   [14,105],[15,105],[16,100],[17,95],[18,95],[19,95],[20,105],[21,105],[22,105],[23,90]]
    .forEach(x => sh.setColumnWidth(x[0], x[1]));

  // ---------- colori: margini sul segno, percentuali sulla soglia di PARAMETRI (30%) ----------
  const n = rUlt - t + 1;
  const pct = (L, c, rosso) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($' + L + t + ');$' + L + t + (rosso ? '<' : '>=') + 'INDIRECT("PARAMETRI!B7"))')
    .setFontColor(rosso ? "#b91c1c" : "#15803d").setBold(true)
    .setBackground(rosso ? "#fee2e2" : "#dcfce7")
    .setRanges([sh.getRange(t, c, n, 1)]).build();
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#b91c1c").setBold(true)
      .setRanges([sh.getRange(t, 16, n, 1), sh.getRange(t, 21, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor("#15803d").setBold(true)
      .setRanges([sh.getRange(t, 16, n, 1), sh.getRange(t, 21, n, 1)]).build(),
    pct('Q', 17, true), pct('Q', 17, false), pct('V', 22, true), pct('V', 22, false)
  ]);

  sh.hideColumns(24, 13);                       // X:AJ appoggi di calcolo
  sh.setFrozenRows(rH);
  sh.getRange(rH, 1, rUlt - rH + 1, NV).createFilter();

  SpreadsheetApp.flush();
  s.setActiveSheet(sh);
  ui.alert("REGISTRO MEZZO creato.\n\n" +
    "Si aggiorna da solo: ogni camion compare nei mesi in cui ha lavorato.\n" +
    "Usa i filtri per vedere un camion solo o un mese solo.");
}
function AGGIORNA_RICAVI_SCHEDE() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  let fatte = 0, gia = 0;
  smElencoMezzi_(s).forEach(x => {
    if (x.tipo !== "CAMION") return;
    const sh = s.getSheetByName(x.targa);
    if (!sh) return;
    let cambiato = false;
    const a6 = sh.getRange("A6");
    const f = a6.getFormula();
    if (f.indexOf("M,Z where") >= 0) {
      a6.setFormula(f.replace("$A$2:$AA", "$A$2:$AD").replace("M,Z where", "M,AD where"));
      cambiato = true;
    }
    ["AG6", "AG7"].forEach(c => {
      const cella = sh.getRange(c);
      const g = cella.getFormula();
      if (g.indexOf("$Z:$Z") >= 0) { cella.setFormula(g.split("$Z:$Z").join("$AD:$AD")); cambiato = true; }
    });
    if (cambiato) fatte++; else gia++;
  });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Schede aggiornate: " + fatte + "\nGia' a posto: " + gia +
    "\n\nOra il ricavo delle schede e' calcolato + manuale.");
}


// ------------------------------------------------------------
//  DAL / AL e TENDINA TRATTA del registro del mese
//  Quando li cambi, nascondono i viaggi fuori periodo o di altre tratte.
//  Le righe non si spostano: €/km e costo manuale restano attaccati
//  al viaggio giusto. "Tutte" toglie il filtro sulla tratta.
//  Gira con un attivatore installato da "Crea il registro del mese".
// ------------------------------------------------------------
function SM_SU_MODIFICA(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SM_FOGLIO_REGISTRO) return;
  // riga 2: B2 (dal), D2 (al), G2:I2 (tratta)
  if (e.range.getRow() !== 2 || e.range.getColumn() > 9) return;
  smApplicaFiltri_(sh);
}

// Applica all'ELENCO dei viaggi il periodo Dal/al e la tratta scelta.
// Le righe vengono nascoste, non spostate.
function smApplicaFiltri_(sh) {
  const filtro = sh.getFilter();
  if (!filtro) return;
  const r1 = filtro.getRange().getRow() + 1;     // prima riga dei viaggi

  // --- periodo (colonna B): dal giorno B2 incluso al giorno D2 incluso
  const dal = sh.getRange("B2").getValue(), al = sh.getRange("D2").getValue();
  if (dal instanceof Date && al instanceof Date) {
    filtro.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria()
      .whenFormulaSatisfied('=AND(ISNUMBER($B' + r1 + ');$B' + r1 + '>=$B$2;$B' + r1 + '<$D$2+1)').build());
  } else {
    filtro.removeColumnFilterCriteria(2);
  }

  // --- tratta (colonna F)
  const scelta = String(sh.getRange("G2").getValue() || "").trim();
  if (scelta === "" || scelta === "Tutte") {
    filtro.removeColumnFilterCriteria(6);
  } else {
    filtro.setColumnFilterCriteria(6,
      SpreadsheetApp.newFilterCriteria().whenTextEqualTo(scelta).build());
  }
}

// Installa (una volta sola) l'attivatore che fa funzionare la tendina
function smInstallaTriggerModifica_() {
  const gia = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "SM_SU_MODIFICA");
  if (!gia) {
    ScriptApp.newTrigger("SM_SU_MODIFICA")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  }
}


// ------------------------------------------------------------
//  AGGIORNA SCHEDE PER IL PIANALE
//  Nelle schede dei camion gia' esistenti: teorico e margine del giorno
//  con l'utile del pianale, e confronto del mese con la riga Pianale.
//  Non tocca nient'altro (le colonne aggiunte a mano restano).
// ------------------------------------------------------------
function AGGIORNA_SCHEDE_PIANALE() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const inizio = Date.now();
  let fatte = 0;
  const rimaste = [];
  smElencoMezzi_(s).forEach(x => {
    if (x.tipo !== "CAMION") return;
    const sh = s.getSheetByName(x.targa);
    if (!sh) return;
    if (Date.now() - inizio > 270000) { rimaste.push(x.targa); return; }
    // formule del giorno
    smFormuleGiornoScheda_(sh, 205, "I");
    smFormuleGiornoScheda_(sh, 205, "K");
    // confronto del mese: tolgo le vecchie regole di colore della zona X:AI e lo riscrivo
    const regole = sh.getConditionalFormatRules().filter(r =>
      !r.getRanges().some(g => g.getColumn() >= 24 && g.getColumn() <= 35 && g.getRow() >= 4 && g.getRow() <= 9));
    sh.setConditionalFormatRules(regole);
    const zona = sh.getRange(4, 24, 6, 12);
    zona.breakApart(); zona.clear();
    smConfronto_(sh);
    fatte++;
  });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Schede aggiornate per il pianale: " + fatte +
    (rimaste.length ? "\n\nNon ho fatto in tempo: " + rimaste.length + "\nRilancia il comando per completare."
                    : "\n\nFatto tutto."));
}


// ============================================================
//  REGISTRO ANNO
//  Il confronto teorico / effettivo MESE PER MESE, diviso tra
//  terra, rifiuti e pianale, con il totale di ogni mese e il
//  TOTALE GENERALE in cima.
//  Si allunga da solo con i mesi nuovi (spazio per SA_MESI mesi).
//  Legge dal registro del mese le colonne "per giornata" (AM/AN),
//  quindi non cambia con la tendina Tratta del registro del mese.
// ============================================================
const SA_FOGLIO = "REGISTRO_ANNO";
const SA_MESI = 36;          // 3 anni

function CREA_REGISTRO_ANNO() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const rm = s.getSheetByName(SM_FOGLIO_REGISTRO);
  if (!rm) { ui.alert("Crea prima il registro del mese."); return; }
  smAssicuraParametri_(s);

  const testa = rm.getRange(1, 1, Math.min(rm.getLastRow(), 300), 2).getValues();
  let rTest = -1;
  for (let i = 0; i < testa.length; i++) {
    if (String(testa[i][0]).trim() === "Mese" && String(testa[i][1]).trim() === "Data") { rTest = i + 1; break; }
  }
  if (rTest < 0) { ui.alert("Non trovo l'elenco dei viaggi nel registro del mese."); return; }
  const R1 = rTest + 1, U = SM_ULTIMA_RIGA_REG;

  let sh = s.getSheetByName(SA_FOGLIO);
  const pos = sh ? sh.getIndex() - 1 : 1;
  if (sh) s.deleteSheet(sh);
  sh = s.insertSheet(SA_FOGLIO, pos);
  sh.setTabColor("#0369a1");
  const r0 = 6, nR = SA_MESI * 4, rUlt = r0 + nR - 1;
  if (sh.getMaxRows() < rUlt + 5) sh.insertRowsAfter(sh.getMaxRows(), rUlt + 5 - sh.getMaxRows());
  if (sh.getMaxColumns() < 18) sh.insertColumnsAfter(sh.getMaxColumns(), 18 - sh.getMaxColumns());

  // dal registro del mese (colonne per giornata, stabili)
  const X = (L) => SM_FOGLIO_REGISTRO + "!$" + L + "$" + R1 + ":$" + L + "$" + U;
  // V/Y: costo e ricavo del SINGOLO viaggio, cosi' la divisione tra terra, rifiuti e pianale e' esatta
  const MESE = X('A'), TR = X('G'), FISSO = X('U'), KM = X('K'), TOT = X('W'), RIC = X('Z'), CHG = X('Y');

  // ------------------------------------------------------------------
  //  A Mese  B Tipo  C Giornate  D Km  E Costo giornaliero medio  F €/km medio  G Costo km
  //  H TOTALE COSTO  I RICAVO EFFETTIVO  J MARGINE  K MARGINE %  L MARGINE REALE per giornata
  //  M Utile atteso per giornata  N RICAVO TEORICO  O DIFFERENZA
  //  Q (nascosta): chiave del mese "2026-09", sulla prima riga di ogni blocco
  //  Ogni mese = 4 righe: Terra, Rifiuti, Pianale, TOTALE
  // ------------------------------------------------------------------
  sh.getRange("A1:O1").merge().setValue("REGISTRO ANNO  —  confronto teorico / effettivo, mese per mese")
    .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground("#0369a1").setVerticalAlignment("middle");
  sh.setRowHeight(1, 36);
  sh.getRange("A2").setValue("Si aggiorna da solo: i mesi nuovi compaiono man mano che arrivano i viaggi.")
    .setFontStyle("italic").setFontColor("#94a3b8").setFontSize(9);

  const rT = 4, rH = 5;
  const titoli = ["Mese", "Tipo", "Giornate\ncamion", "Km", "Costo giornaliero\nmedio", "€/km\nmedio", "Costo km",
    "TOTALE\nCOSTO", "RICAVO\nEFFETTIVO", "MARGINE", "MARGINE\n%", "MARGINE REALE\nper giornata",
    "Utile atteso\nper giornata", "RICAVO\nTEORICO", "DIFFERENZA"];
  smIntestazioni_(sh, rH, 1, titoli);
  sh.getRange(rH, 1, 1, 15).setBackground("#0b192c").setFontColor("#FFFFFF");
  sh.setRowHeight(rH, 44);

  // formule che si calcolano dalla riga stessa (uguali per tipi, totali di mese e totale generale)
  const derivate = (r) => ({
    E: '=IFERROR((H' + r + '-G' + r + ')/C' + r + ';"")',
    F: '=IFERROR(G' + r + '/D' + r + ';"")',
    J: '=I' + r + '-H' + r,
    K: '=IFERROR(J' + r + '/I' + r + ';"")',
    L: '=IFERROR(J' + r + '/C' + r + ';"")',
    M: '=IFERROR((N' + r + '-H' + r + ')/C' + r + ';"")',
    O: '=I' + r + '-N' + r
  });
  const ordina = (o) => ['C','D','E','F','G','H','I','J','K','L','M','N','O'].map(L => o[L]);
  // ogni formula del blocco vale solo se il mese esiste
  const se = (base, fx) => fx === '' ? '' : '=IF($Q$' + base + '="";"";' + fx.substring(1) + ')';

  const tipi = [
    { nome: "Terra", cod: "t", par: "$B$4" },
    { nome: "Rifiuti", cod: "r", par: "$B$5" },
    { nome: "Pianale", cod: "p", par: "$B$9" }
  ];
  const righe = [], chiavi = [];
  for (let k = 0; k < SA_MESI; k++) {
    const base = r0 + k * 4, Q = '$Q$' + base;
    // chiave del mese numero k+1 (dal piu' vecchio)
    chiavi.push(['=IFERROR(INDEX(SORT(UNIQUE(FILTER(' + MESE + ';' + MESE + '<>"")));' + (k + 1) + ');"")'], [''], [''], ['']);

    tipi.forEach((x, j) => {
      const r = base + j;
      const crit = MESE + ';' + Q + ';' + TR + ';"' + x.cod + '"';
      const cond = MESE + '=' + Q + ';' + TR + '="' + x.cod + '"';
      const o = Object.assign(derivate(r), {
        C: '=IF(COUNTIFS(' + crit + ')=0;0;COUNTUNIQUE(FILTER(' + CHG + ';' + cond + ')))',
        D: '=SUMIFS(' + KM + ';' + crit + ')',
        G: '=SUMIFS(' + TOT + ';' + crit + ')-SUMIFS(' + FISSO + ';' + crit + ')',
        H: '=SUMIFS(' + TOT + ';' + crit + ')',
        I: '=SUMIFS(' + RIC + ';' + crit + ')',
        N: '=H' + r + '+C' + r + '*PARAMETRI!' + x.par       // teorico: costo + giornate x utile fisso
      });
      const a = j === 0
        ? '=PROPER(TEXT(DATE(VALUE(LEFT(' + Q + ';4));VALUE(RIGHT(' + Q + ';2));1);"mmmm yyyy"))'
        : '';
      righe.push([se(base, a), se(base, '="' + x.nome + '"')].concat(ordina(o).map(f => se(base, f))));
    });

    // TOTALE del mese: somma dei tre tipi, poi i rapporti
    const r = base + 3;
    const somma = (L) => '=SUM(' + L + base + ':' + L + (base + 2) + ')';
    // giornate REALI del mese: un camion con due tipi di lavoro nello stesso giorno conta una volta
    const giornateReali = '=IFERROR(ROWS(UNIQUE(FILTER(' + CHG + ';' + MESE + '=' + Q + ';' +
      '((' + TR + '="t")+(' + TR + '="r")+(' + TR + '="p"))>0)));0)';
    const o = Object.assign(derivate(r), { C: giornateReali, D: somma('D'), G: somma('G'), H: somma('H'), I: somma('I'), N: somma('N') });
    righe.push(['', se(base, '="TOTALE"')].concat(ordina(o).map(f => se(base, f))));
  }
  sh.getRange(r0, 1, nR, 15).setFormulas(righe);
  sh.getRange(r0, 17, nR, 1).setFormulas(chiavi);

  // ---------- TOTALE GENERALE in cima: somma dei TOTALE dei mesi ----------
  const somGen = (L) => '=SUMIF($B$' + r0 + ':$B$' + rUlt + ';"TOTALE";' + L + r0 + ':' + L + rUlt + ')';
  const og = Object.assign(derivate(rT), { C: somGen('C'), D: somGen('D'), G: somGen('G'), H: somGen('H'), I: somGen('I'), N: somGen('N') });
  sh.getRange(rT, 1).setValue("TOTALE GENERALE");
  sh.getRange(rT, 3, 1, 13).setFormulas([ordina(og)]);
  sh.getRange(rT, 1, 1, 15).setFontWeight("bold").setBackground("#93c5fd")
    .setBorder(true, true, true, true, true, true, "#0369a1", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(rT, 28);

  // ---------- formati ----------
  const fmt = (c, n, f) => sh.getRange(rT, c, rUlt - rT + 1, n).setNumberFormat(f);
  fmt(3, 2, "#,##0");
  fmt(5, 1, '#,##0.00 "€"');
  fmt(6, 1, '0.00 "€"');
  fmt(7, 4, '#,##0.00 "€"');
  fmt(11, 1, "0.0%");
  fmt(12, 4, '#,##0.00 "€"');
  sh.getRange(rT, 1, rUlt - rT + 1, 15).setHorizontalAlignment("center");
  sh.getRange(r0, 1, nR, 15).setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r0, 1, nR, 1).setFontWeight("bold");
  [[1,130],[2,80],[3,75],[4,80],[5,105],[6,75],[7,100],[8,105],[9,105],[10,100],[11,75],
   [12,110],[13,105],[14,105],[15,105]].forEach(x => sh.setColumnWidth(x[0], x[1]));

  // ---------- colori ----------
  const zona = (c) => sh.getRange(rT, c, rUlt - rT + 1, 1);
  const regola = (f, rosso, c) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)
    .setFontColor(rosso ? "#b91c1c" : "#15803d").setBold(true)
    .setBackground(rosso ? "#fee2e2" : "#dcfce7").setRanges([zona(c)]).build();
  sh.setConditionalFormatRules([
    // riga TOTALE di ogni mese
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B' + r0 + '="TOTALE"')
      .setBold(true).setBackground("#dbeafe").setRanges([sh.getRange(r0, 1, nR, 15)]).build(),
    // margine e differenza: segno
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor("#b91c1c").setBold(true)
      .setRanges([zona(10), zona(15)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor("#15803d").setBold(true)
      .setRanges([zona(10), zona(15)]).build(),
    // margine %: soglia di PARAMETRI (30%). INDIRECT: le regole non possono leggere altri fogli
    regola('=AND(ISNUMBER($K' + rT + ');$K' + rT + '<INDIRECT("PARAMETRI!B7"))', true, 11),
    regola('=AND(ISNUMBER($K' + rT + ');$K' + rT + '>=INDIRECT("PARAMETRI!B7"))', false, 11),
    // margine per giornata contro l'utile atteso della stessa riga
    regola('=AND(ISNUMBER($L' + rT + ');ISNUMBER($M' + rT + ');$L' + rT + '<$M' + rT + ')', true, 12),
    regola('=AND(ISNUMBER($L' + rT + ');ISNUMBER($M' + rT + ');$L' + rT + '>=$M' + rT + ')', false, 12)
  ]);

  sh.hideColumns(16, 2);                         // P:Q
  sh.setFrozenRows(rH);
  SpreadsheetApp.flush();
  s.setActiveSheet(sh);
  ui.alert("REGISTRO ANNO creato.\n\nSi aggiorna da solo con i mesi nuovi: non serve rilanciarlo.");
}