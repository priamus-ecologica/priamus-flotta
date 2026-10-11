// ============================================================
//  MAIL SERALE: ogni sera verso le 20 manda una mail con i viaggi
//  di oggi (e dei giorni prima) non ancora avviati o non chiusi.
//  Menu: 🚛 Priamus > 📧 Mail serale
// ============================================================

const RS_ORA = 20;                 // ora di invio (fuso Europe/Rome)
const RS_FUNZIONE = "RIEPILOGO_SERALE";
const RS_RIGHE = 2000;             // righe recenti di REGISTRO_VIAGGI da controllare

/** Menu: attiva l'invio automatico ogni sera. */
function ATTIVA_MAIL_SERALE() {
  rsTogliTrigger_();
  ScriptApp.newTrigger(RS_FUNZIONE).timeBased().atHour(RS_ORA).nearMinute(0)
    .everyDays(1).inTimezone("Europe/Rome").create();
  SpreadsheetApp.getUi().alert("Mail serale attivata: ogni sera verso le " + RS_ORA +
    " arriverà a " + rsDestinatario_() + ".");
}

/** Menu: spegne l'invio automatico. */
function DISATTIVA_MAIL_SERALE() {
  const n = rsTogliTrigger_();
  SpreadsheetApp.getUi().alert(n ? "Mail serale disattivata." : "La mail serale non era attiva.");
}

/** Menu: manda subito la mail, per provarla. */
function PROVA_MAIL_SERALE() {
  RIEPILOGO_SERALE();
  SpreadsheetApp.getUi().alert("Mail mandata a " + rsDestinatario_() + ".");
}

/** Lanciata dal trigger ogni sera. */
function RIEPILOGO_SERALE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = "Europe/Rome";
  const oggi = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  // nomi dei camion
  const nomi = {};
  const an = ss.getSheetByName("ANAGRAFICA_MEZZI");
  if (an && an.getLastRow() > 1) {
    an.getRange(2, 1, an.getLastRow() - 1, 4).getValues().forEach(r => {
      const t = String(r[1] || "").trim();
      if (t) nomi[t] = String(r[3] || "").trim() || t;
    });
  }

  const sh = ss.getSheetByName("REGISTRO_VIAGGI");
  const ultima = sh.getLastRow();
  const da = Math.max(2, ultima - RS_RIGHE + 1);
  const righe = ultima >= 2 ? sh.getRange(da, 1, ultima - da + 1, 20).getValues() : [];

  const nonAvviati = [], nonChiusi = [];
  righe.forEach((r, i) => {
    const giorno = rsGiorno_(r[1], tz);
    if (!giorno || giorno > oggi) return;              // vuote o viaggi dei prossimi giorni
    const stato = String(r[15] || "").trim();
    if (stato === "Completato") return;
    const targa = String(r[4] || "").trim();
    const v = {
      giorno: giorno,
      riga: da + i,
      autista: String(r[2] || "").trim() || "(senza autista)",
      camion: (nomi[targa] || targa) + (targa && nomi[targa] ? " · " + targa : ""),
      tratta: String(r[18] || "").trim() || (String(r[8] || "") + " → " + String(r[9] || "")),
      ora: rsOra_(r[19], tz)
    };
    if (stato === "In Corso") nonChiusi.push(v); else nonAvviati.push(v);
  });

  const ord = (a, b) => (a.giorno < b.giorno ? -1 : a.giorno > b.giorno ? 1 : a.autista.localeCompare(b.autista));
  nonAvviati.sort(ord); nonChiusi.sort(ord);

  const dataLeggibile = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  const tot = nonAvviati.length + nonChiusi.length;
  const oggetto = tot
    ? "Priamus · " + dataLeggibile + " · " + tot + (tot === 1 ? " viaggio da sistemare" : " viaggi da sistemare")
    : "Priamus · " + dataLeggibile + " · tutti i viaggi chiusi ✅";

  let html = "<div style='font-family:Arial,sans-serif;font-size:14px;color:#0b192c'>";
  if (!tot) {
    html += "<p>✅ Tutti i viaggi di oggi e dei giorni precedenti sono stati avviati e chiusi.</p>";
  } else {
    html += rsTabella_("🔴 Non chiusi (In corso)", "L'autista ha avviato il viaggio ma non l'ha chiuso: mancano km di arrivo, tonnellate e ore.",
                       nonChiusi, oggi, true);
    html += rsTabella_("🟠 Non avviati (Assegnati)", "Il viaggio è stato assegnato ma l'autista non ha premuto \"Inizia tratta\".",
                       nonAvviati, oggi, false);
  }
  html += "<p style='color:#64748b;font-size:12px;margin-top:18px'>Mail automatica del foglio Priamus. " +
          "Per spegnerla: menu 🚛 Priamus → 📧 Mail serale → Disattiva.</p></div>";

  MailApp.sendEmail({ to: rsDestinatario_(), subject: oggetto, htmlBody: html, name: "Priamus Flotta" });
}

// ---------------- funzioni di supporto ----------------

function rsTabella_(titolo, spiegazione, elenco, oggi, conOra) {
  if (!elenco.length) return "";
  const cella = "padding:6px 8px;border-bottom:1px solid #e2e8f0;text-align:left;vertical-align:top";
  let h = "<h3 style='margin:18px 0 4px'>" + titolo + " — " + elenco.length + "</h3>" +
          "<p style='margin:0 0 8px;color:#64748b;font-size:12px'>" + spiegazione + "</p>" +
          "<table style='border-collapse:collapse;width:100%;font-size:13px'><tr style='background:#f1f5f9'>" +
          "<th style='" + cella + "'>Giorno</th><th style='" + cella + "'>Autista</th><th style='" + cella + "'>Camion</th>" +
          "<th style='" + cella + "'>Tratta</th>" + (conOra ? "<th style='" + cella + "'>Partito alle</th>" : "") +
          "<th style='" + cella + "'>Riga</th></tr>";
  elenco.forEach(v => {
    const vecchio = v.giorno < oggi;
    const g = v.giorno.split("-").reverse().slice(0, 2).join("/");
    h += "<tr" + (vecchio ? " style='background:#fef2f2'" : "") + ">" +
         "<td style='" + cella + "'>" + (vecchio ? "<b>" + g + "</b>" : "oggi") + "</td>" +
         "<td style='" + cella + "'>" + rsEsc_(v.autista) + "</td>" +
         "<td style='" + cella + "'>" + rsEsc_(v.camion) + "</td>" +
         "<td style='" + cella + "'>" + rsEsc_(v.tratta) + "</td>" +
         (conOra ? "<td style='" + cella + "'>" + rsEsc_(v.ora) + "</td>" : "") +
         "<td style='" + cella + "'>" + v.riga + "</td></tr>";
  });
  return h + "</table>";
}

function rsGiorno_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  const m = String(v || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2) : "";
}

function rsOra_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, "HH:mm");
  return String(v || "");
}

function rsEsc_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function rsDestinatario_() {
  return PropertiesService.getScriptProperties().getProperty("EMAIL_RIEPILOGO") ||
         Session.getEffectiveUser().getEmail();
}

function rsTogliTrigger_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === RS_FUNZIONE) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}
