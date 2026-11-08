// ============================================================
//  PRIAMUS - Controllo notifiche (funzione programmata)
//
//  File: netlify/functions/notifiche.js
//  Netlify la avvia da sola ogni 2 minuti (vedi netlify.toml).
//  Cerca i viaggi appena assegnati e avvisa gli autisti sul telefono.
//  La logica sta in netlify/lib/push.js.
// ============================================================

const { google } = require("googleapis");
const push = require("../lib/push");

exports.handler = async function () {
  try {
    const auth = new google.auth.JWT(
      process.env.GOOGLE_CLIENT_EMAIL,
      null,
      (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );
    const s = google.sheets({ version: "v4", auth });
    const esito = await push.inviaPendenti(s, process.env.SHEET_ID);
    console.log("Notifiche: " + JSON.stringify(esito));
    return { statusCode: 200, body: JSON.stringify(esito) };
  } catch (e) {
    console.log("Notifiche, errore: " + (e.message || e));
    return { statusCode: 500, body: String(e.message || e) };
  }
};
