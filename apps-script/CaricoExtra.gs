// ============================================================
//  CARICO EXTRA dal menu 🚛 Priamus
//  Selezioni un viaggio (in REGISTRO_VIAGGI o in REGISTRO_MESE),
//  scegli la tratta (di solito "Saline a tonnellata") e le tonnellate:
//  viene aggiunta una riga SUBITO SOTTO quel viaggio, con lo stesso
//  giorno, autista e camion, km vuoti e stato Completato.
//  Cosi' il ricavo va al cliente giusto invece del "ricavo manuale".
//  E' la stessa cosa del bottone "➕ Carico extra" nella scheda Modifica dell'app.
// ============================================================

const CE_FOGLIO = "REGISTRO_VIAGGI";
const CE_TRATTA_PREDEFINITA = "Saline a tonnellata";

/** Menu: apre la finestra per il viaggio selezionato. */
function CARICO_EXTRA() {
  const ui = SpreadsheetApp.getUi();
  const v = ceViaggioSelezionato_();
  if (!v) {
    ui.alert("Carico extra",
      "Clicca prima su una riga di un viaggio in REGISTRO_VIAGGI o in REGISTRO_MESE, poi riapri il menu.\n\n" +
      "Il carico extra verra' aggiunto sullo stesso camion e nello stesso giorno di quel viaggio.",
      ui.ButtonSet.OK);
    return;
  }
  const tratte = ceTratte_();
  const html = HtmlService.createHtmlOutput(ceHtml_(v, tratte)).setWidth(420).setHeight(520);
  ui.showModalDialog(html, "➕ Carico extra");
}

/** Chiamata dalla finestra: aggiunge la riga. Restituisce un messaggio. */
function CE_SALVA(dati) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CE_FOGLIO);

    // ritrovo il viaggio per ID: nel frattempo le righe potrebbero essersi spostate
    const riga = ceRigaDaId_(sh, dati.id);
    if (!riga) throw new Error("Viaggio non trovato: chiudi la finestra e riprova.");
    const r = sh.getRange(riga, 1, 1, 22).getValues()[0];

    const tratta = String(dati.tratta || "").trim();
    const regola = ceTratte_().find(t => t.chiave === tratta);
    if (!regola) throw new Error("Scegli una tratta dell'elenco.");

    const num = (v, nome, intero) => {
      const s = String(v == null ? "" : v).trim().replace(",", ".");
      if (s === "") return "";
      const n = Number(s);
      if (isNaN(n) || n < 0 || (intero && Math.round(n) !== n)) throw new Error(nome + ": numero non valido.");
      return n;
    };
    const ton = num(dati.ton, "Tonnellate"), nv = num(dati.nv, "Numero viaggi", true), ore = num(dati.ore, "Ore");
    if (ton === "" && nv === "" && ore === "") throw new Error("Scrivi almeno le tonnellate, le ore o il numero di viaggi.");
    if (ton !== "" && ton > 1000) throw new Error(ton + " t sembrano kg: scrivile in tonnellate (es. 32,5).");

    // data del viaggio senza ora
    const d0 = r[1] instanceof Date ? r[1] : new Date(r[1]);
    if (isNaN(d0)) throw new Error("Il viaggio selezionato non ha una data valida.");
    const giorno = new Date(d0.getFullYear(), d0.getMonth(), d0.getDate());

    const id = "extra-" + Date.now() + "-" + Math.random().toString(36).slice(2, 6);
    const nota = String(dati.nota || "").trim() || ("Carico extra del viaggio qui sopra (" + String(r[18] || "") + ")");
    const valori = [id, giorno, String(r[2] || ""), "", String(r[4] || "").trim(), "", "", "",
      regola.partenza, regola.arrivo, "", ton, ore, "", "", "Completato", "", nota, regola.chiave, "", nv, ""];

    sh.insertRowAfter(riga);
    sh.getRange(riga + 1, 1, 1, valori.length).setValues([valori]);

    // registro delle modifiche (se qualcosa va storto qui, la riga resta comunque aggiunta)
    try {
      const log = ss.getSheetByName("MODIFICHE_VIAGGI");
      if (log) {
        // l'email di chi usa il menu: senza l'autorizzazione "userinfo.email" Google da' errore, quindi la provo e basta
        let chi = "";
        try { chi = Session.getActiveUser().getEmail(); } catch (e) {}
        const quando = Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm");
        log.appendRow(["'" + quando, chi || "foglio (menu)", id, riga + 1, Utilities.formatDate(giorno, "Europe/Rome", "dd/MM/yyyy"),
          String(r[2] || ""), "Carico extra", "",
          regola.chiave + (ton !== "" ? " · " + ton + " t" : "") + (nv !== "" ? " · " + nv + " viaggi" : "") + (ore !== "" ? " · " + ore + " ore" : "")]);
      }
    } catch (e) {}

    // il registro del mese nasconde le righe per posizione: dopo l'aggiunta rimetto a posto il filtro
    try {
      const rm = ss.getSheetByName(SM_FOGLIO_REGISTRO);
      if (rm) { SpreadsheetApp.flush(); smApplicaFiltri_(rm); }
    } catch (e) {}

    return "Aggiunta la riga " + (riga + 1) + ": " + regola.chiave +
      (ton !== "" ? ", " + String(ton).replace(".", ",") + " t" : "") + " sotto il viaggio di " + String(r[2] || "") + ".";
  } finally {
    lock.releaseLock();
  }
}

// ---------------- funzioni di supporto ----------------

// Il viaggio della riga selezionata: da REGISTRO_VIAGGI direttamente,
// da REGISTRO_MESE tramite l'ID (colonna AA nascosta).
function ceViaggioSelezionato_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const riga = sh.getActiveRange() ? sh.getActiveRange().getRow() : 0;
  let id = "";
  if (sh.getName() === CE_FOGLIO) {
    if (riga < 2) return null;
    id = String(sh.getRange(riga, 1).getValue() || "").trim();
  } else if (sh.getName() === SM_FOGLIO_REGISTRO) {
    if (riga < 11) return null;
    id = String(sh.getRange(riga, SM_COL_ID).getValue() || "").trim();
  } else {
    return null;
  }
  if (!id) return null;
  const reg = ss.getSheetByName(CE_FOGLIO);
  const rr = ceRigaDaId_(reg, id);
  if (!rr) return null;
  const r = reg.getRange(rr, 1, 1, 22).getValues()[0];
  const d = r[1] instanceof Date ? Utilities.formatDate(r[1], "Europe/Rome", "dd/MM/yyyy") : String(r[1] || "");
  return { id: id, riga: rr, data: d, autista: String(r[2] || ""), targa: String(r[4] || ""), tratta: String(r[18] || "") };
}

function ceRigaDaId_(sh, id) {
  const ult = sh.getLastRow();
  if (ult < 2) return 0;
  const f = sh.getRange(2, 1, ult - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}

function ceTratte_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGOLE_VIAGGI");
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  return sh.getRange(2, 1, ult - 1, 3).getValues()
    .filter(r => String(r[0] || "").trim() !== "")
    .map(r => ({ chiave: String(r[0]).trim(), partenza: String(r[1] || "").trim(), arrivo: String(r[2] || "").trim() }));
}

function ceEsc_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function ceHtml_(v, tratte) {
  const opz = tratte.map(t => "<option" + (t.chiave === CE_TRATTA_PREDEFINITA ? " selected" : "") + ">" + ceEsc_(t.chiave) + "</option>").join("");
  return "<!doctype html><html><head><base target='_top'><style>" +
    "body{font-family:Arial,sans-serif;font-size:13px;color:#0b192c;margin:0;padding:4px 2px}" +
    ".info{background:#f1f5f9;border-radius:8px;padding:10px;margin-bottom:12px;line-height:1.5}" +
    "label{display:block;font-weight:bold;font-size:12px;margin:10px 0 4px}" +
    "input,select{width:100%;box-sizing:border-box;padding:8px;font-size:14px;border:1px solid #cbd5e1;border-radius:6px}" +
    ".riga{display:flex;gap:8px}.riga>div{flex:1}" +
    "button{margin-top:14px;width:100%;padding:10px;font-size:14px;font-weight:bold;border:0;border-radius:8px;background:#0b192c;color:#fff;cursor:pointer}" +
    "button:disabled{opacity:.5}#msg{margin-top:10px;font-weight:bold}.err{color:#b91c1c}.ok{color:#15803d}" +
    ".nota{color:#64748b;font-size:11px;margin-top:6px}" +
    "</style></head><body>" +
    "<div class='info'><b>" + ceEsc_(v.data) + "</b> · " + ceEsc_(v.autista) + "<br>Camion <b>" + ceEsc_(v.targa) + "</b> · " + ceEsc_(v.tratta) +
    "<div class='nota'>La riga nuova va sotto questo viaggio, con km vuoti (i km restano sul viaggio principale).</div></div>" +
    "<label>Tratta</label><select id='tratta'>" + opz + "</select>" +
    "<div class='riga'><div><label>Tonnellate</label><input id='ton' inputmode='decimal' placeholder='es. 32,5'></div>" +
    "<div><label>N. viaggi</label><input id='nv' inputmode='numeric' placeholder='1'></div></div>" +
    "<label>Ore (solo per tratte a ore)</label><input id='ore' inputmode='decimal'>" +
    "<label>Nota</label><input id='nota' placeholder='facoltativa'>" +
    "<button id='ok' onclick='salva()'>Aggiungi</button><div id='msg'></div>" +
    "<script>" +
    "function salva(){var b=document.getElementById('ok');b.disabled=true;b.textContent='Salvataggio…';" +
    "var m=document.getElementById('msg');m.textContent='';" +
    "google.script.run.withSuccessHandler(function(t){m.className='ok';m.textContent='✅ '+t;b.textContent='Fatto';setTimeout(function(){google.script.host.close()},3000);})" +
    ".withFailureHandler(function(e){m.className='err';m.textContent='⚠️ '+e.message;b.disabled=false;b.textContent='Aggiungi';})" +
    ".CE_SALVA({id:" + JSON.stringify(v.id) + ",tratta:document.getElementById('tratta').value,ton:document.getElementById('ton').value," +
    "nv:document.getElementById('nv').value,ore:document.getElementById('ore').value,nota:document.getElementById('nota').value});}" +
    "document.getElementById('ton').focus();" +
    "</script></body></html>";
}
