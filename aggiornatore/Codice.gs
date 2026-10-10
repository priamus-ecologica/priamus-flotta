/**
 * AGGIORNATORE PRIAMUS
 * Progetto Apps Script SEPARATO dal foglio.
 * Prende il codice dalla cartella "apps-script" del repository GitHub
 * e lo installa nel progetto Apps Script collegato al foglio Priamus.
 * Prima di ogni modifica salva una copia su Drive, per poter tornare indietro.
 *
 * L'unica cosa da compilare è SCRIPT_ID_FOGLIO.
 */

// ID dello script del foglio: nel foglio Priamus → Estensioni → Apps Script →
// Impostazioni progetto (ingranaggio a sinistra) → "ID script".
const SCRIPT_ID_FOGLIO = '1p8hex7suF_Kqw-8NdaBr9oXvOz6Dl3DhKMaT-GjPbB4ZrvxrYogxnZfo';

const GH_OWNER   = 'priamus-ecologica';
const GH_REPO    = 'priamus-flotta';
const GH_BRANCH  = 'main';
const GH_CARTELLA = 'apps-script';

// Pubblicazioni (/exec) da portare all'ultima versione a ogni aggiornamento.
// Solo quella usata dall'app Netlify; le altre vecchie pubblicazioni restano ferme.
const PUBBLICAZIONI_DA_AGGIORNARE = [
  'AKfycbxhci3gX_VkNZ6EdIOYMLdMcDrZXm7W0SJ_YMjShFR0MSzIQyqjss32kpnayR6gEg39'
];

const CARTELLA_DRIVE  ='Priamus - copie codice Apps Script';
const FILE_ESPORTAZIONE = 'PRIAMUS_codice_attuale.json';
const COPIE_DA_TENERE = 30;

/* ============================== PAGINA ============================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Pagina')
    .setTitle('Aggiornatore Priamus')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ======================== AZIONI DELLA PAGINA ======================== */

/** Confronta il codice su GitHub con quello installato nel foglio, senza cambiare niente. */
function controllaNovita() {
  controllaConfig_();
  const attuali = leggiProgetto_();
  const nuovi = leggiGithub_();
  return confronta_(attuali, nuovi);
}

/**
 * Installa il codice di GitHub nel foglio.
 * @param {boolean} consentiRimozioni  true per eliminare dal foglio i file che non sono più su GitHub
 * @param {boolean} aggiornaPubblicata true per aggiornare anche l'app pubblicata (indirizzo /exec)
 */
function installaAggiornamento(consentiRimozioni, aggiornaPubblicata) {
  controllaConfig_();
  const attuali = leggiProgetto_();
  const nuovi = leggiGithub_();
  const diff = confronta_(attuali, nuovi);

  if (!diff.cambiamenti) return { esito: 'Niente da aggiornare: il codice è già uguale a GitHub.' };
  if (diff.rimossi.length && !consentiRimozioni) {
    throw new Error('Questi file verrebbero eliminati dal foglio: ' + diff.rimossi.join(', ') +
      '. Spunta "consenti eliminazione" se è voluto.');
  }

  const pubblicazioni = leggiPubblicazioni_();
  const copia = salvaCopia_(attuali, pubblicazioni, 'prima di aggiornamento');

  api_('put', 'projects/' + SCRIPT_ID_FOGLIO + '/content', { files: nuovi });

  let msgPubbl = '';
  const daSpostare = pubblicazioni.filter(function (p) { return PUBBLICAZIONI_DA_AGGIORNARE.indexOf(p.id) >= 0; });
  if (aggiornaPubblicata && daSpostare.length) {
    const v = api_('post', 'projects/' + SCRIPT_ID_FOGLIO + '/versions',
      { description: 'Aggiornatore ' + oraLeggibile_() });
    daSpostare.forEach(function (p) { spostaPubblicazione_(p, v.versionNumber); });
    msgPubbl = ' App pubblicata aggiornata alla versione ' + v.versionNumber + '.';
  }

  return {
    esito: 'Aggiornamento installato.' + msgPubbl,
    copia: copia,
    dettaglio: diff
  };
}

/** Salva su Drive il codice installato adesso nel foglio (serve a Claude per leggerlo). */
function esportaCodice() {
  controllaConfig_();
  const files = leggiProgetto_();
  const pubblicazioni = leggiPubblicazioni_();
  const dati = {
    esportato: oraLeggibile_(),
    scriptId: SCRIPT_ID_FOGLIO,
    pubblicazioni: pubblicazioni,
    files: files
  };
  const cartella = cartella_();
  const testo = JSON.stringify(dati, null, 2);
  const esistenti = cartella.getFilesByName(FILE_ESPORTAZIONE);
  if (esistenti.hasNext()) esistenti.next().setContent(testo);
  else cartella.createFile(FILE_ESPORTAZIONE, testo, MimeType.PLAIN_TEXT);
  return {
    esito: 'Codice esportato su Drive (' + files.length + ' file).',
    file: files.map(function (f) { return nomeCompleto_(f); })
  };
}

/** Elenco delle copie salvate, dalla più recente. */
function elencoCopie() {
  const it = cartella_().getFiles();
  const out = [];
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('copia_') === 0) out.push({ id: f.getId(), nome: f.getName(), data: f.getDateCreated().getTime() });
  }
  out.sort(function (a, b) { return b.data - a.data; });
  return out.map(function (c) { return { id: c.id, nome: c.nome }; });
}

/** Rimette nel foglio il codice di una copia (e l'app pubblicata alla versione di allora). */
function ripristinaCopia(idFile) {
  controllaConfig_();
  const dati = JSON.parse(DriveApp.getFileById(idFile).getBlob().getDataAsString());
  if (!dati.files || !dati.files.length) throw new Error('La copia è vuota o non valida.');

  const attuali = leggiProgetto_();
  const pubblicazioniOra = leggiPubblicazioni_();
  salvaCopia_(attuali, pubblicazioniOra, 'prima di ripristino');

  api_('put', 'projects/' + SCRIPT_ID_FOGLIO + '/content', { files: dati.files });

  let msg = '';
  (dati.pubblicazioni || []).forEach(function (p) {
    if (p.versione && PUBBLICAZIONI_DA_AGGIORNARE.indexOf(p.id) >= 0) { spostaPubblicazione_(p, p.versione); msg = ' App pubblicata riportata alla versione di allora.'; }
  });
  return { esito: 'Ripristinata la copia "' + dati.motivo + '" del ' + dati.salvata + '.' + msg };
}

/* ============================ FUNZIONI INTERNE ============================ */

function controllaConfig_() {
  if (!SCRIPT_ID_FOGLIO || SCRIPT_ID_FOGLIO.indexOf('INCOLLA') === 0) {
    throw new Error('Manca SCRIPT_ID_FOGLIO in cima al file Codice.gs.');
  }
}

function api_(metodo, percorso, corpo) {
  const opz = {
    method: metodo,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
    contentType: 'application/json'
  };
  if (corpo) opz.payload = JSON.stringify(corpo);
  const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/' + percorso, opz);
  const testo = r.getContentText();
  if (r.getResponseCode() >= 300) {
    let m = testo;
    try { m = JSON.parse(testo).error.message; } catch (e) {}
    if (/User has not enabled/i.test(m)) {
      m = 'API Apps Script spenta nelle tue impostazioni: apri script.google.com/home/usersettings e accendila. (' + m + ')';
    } else if (/has not been used in project|is disabled/i.test(m)) {
      m = 'API Apps Script non abilitata nel progetto Google Cloud collegato all\'Aggiornatore. (' + m + ')';
    }
    throw new Error('Google Apps Script API: ' + m);
  }
  return testo ? JSON.parse(testo) : {};
}

function leggiProgetto_() {
  const r = api_('get', 'projects/' + SCRIPT_ID_FOGLIO + '/content');
  return (r.files || []).map(function (f) {
    return { name: f.name, type: f.type, source: f.source };
  });
}

/** Pubblicazioni con versione fissa (es. indirizzo /exec). La pubblicazione "HEAD" (/dev) si aggiorna da sola. */
function leggiPubblicazioni_() {
  const r = api_('get', 'projects/' + SCRIPT_ID_FOGLIO + '/deployments?pageSize=50');
  return (r.deployments || [])
    .filter(function (d) { return d.deploymentConfig && d.deploymentConfig.versionNumber; })
    .map(function (d) {
      const web = (d.entryPoints || []).filter(function (e) { return e.entryPointType === 'WEB_APP'; })[0];
      return {
        id: d.deploymentId,
        versione: d.deploymentConfig.versionNumber,
        descrizione: d.deploymentConfig.description || '',
        url: web && web.webApp ? web.webApp.url : ''
      };
    });
}

function spostaPubblicazione_(p, versione) {
  api_('put', 'projects/' + SCRIPT_ID_FOGLIO + '/deployments/' + p.id, {
    deploymentConfig: {
      scriptId: SCRIPT_ID_FOGLIO,
      versionNumber: versione,
      manifestFileName: 'appsscript',
      description: p.descrizione || ('Aggiornatore ' + oraLeggibile_())
    }
  });
}

function leggiGithub_() {
  const base = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/';
  const elenco = ghFetch_(base + GH_CARTELLA + '?ref=' + GH_BRANCH, 'application/vnd.github+json');
  const voci = JSON.parse(elenco);
  if (!Array.isArray(voci)) throw new Error('Cartella "' + GH_CARTELLA + '" non trovata su GitHub.');

  const files = [];
  voci.filter(function (v) { return v.type === 'file'; }).forEach(function (v) {
    const m = v.name.match(/^(.*)\.(gs|js|html|json)$/i);
    if (!m) return; // ignora README e altri file
    const est = m[2].toLowerCase();
    let tipo;
    if (est === 'json') {
      if (m[1] !== 'appsscript') return; // l'unico JSON ammesso è il manifest
      tipo = 'JSON';
    } else if (est === 'html') tipo = 'HTML';
    else tipo = 'SERVER_JS';
    const testo = ghFetch_(base + GH_CARTELLA + '/' + encodeURIComponent(v.name) + '?ref=' + GH_BRANCH,
      'application/vnd.github.raw');
    files.push({ name: m[1], type: tipo, source: testo });
  });

  if (!files.some(function (f) { return f.type === 'JSON'; })) {
    throw new Error('Su GitHub manca appsscript.json: aggiornamento bloccato per sicurezza.');
  }
  if (!files.some(function (f) { return f.type === 'SERVER_JS'; })) {
    throw new Error('Su GitHub non ci sono file .gs: aggiornamento bloccato per sicurezza.');
  }
  return files;
}

function ghFetch_(url, accept) {
  const r = UrlFetchApp.fetch(url, {
    headers: { Accept: accept, 'X-GitHub-Api-Version': '2022-11-28' },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) {
    throw new Error('GitHub ha risposto ' + r.getResponseCode() + ' per ' + url.split('?')[0]);
  }
  return r.getContentText();
}

function nomeCompleto_(f) {
  return f.name + (f.type === 'JSON' ? '.json' : f.type === 'HTML' ? '.html' : '.gs');
}

function normalizza_(s) {
  return String(s || '').replace(/\r\n/g, '\n').replace(/\s+$/, '');
}

function confronta_(attuali, nuovi) {
  const mappa = {};
  attuali.forEach(function (f) { mappa[nomeCompleto_(f)] = f; });
  const aggiunti = [], modificati = [], uguali = [];
  nuovi.forEach(function (f) {
    const n = nomeCompleto_(f);
    if (!mappa[n]) aggiunti.push(n);
    else if (normalizza_(mappa[n].source) !== normalizza_(f.source)) modificati.push(n);
    else uguali.push(n);
    delete mappa[n];
  });
  const rimossi = Object.keys(mappa);
  return {
    aggiunti: aggiunti, modificati: modificati, rimossi: rimossi, uguali: uguali,
    cambiamenti: aggiunti.length + modificati.length + rimossi.length
  };
}

function cartella_() {
  const it = DriveApp.getFoldersByName(CARTELLA_DRIVE);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARTELLA_DRIVE);
}

function salvaCopia_(files, pubblicazioni, motivo) {
  const cartella = cartella_();
  const ora = new Date();
  const nome = 'copia_' + Utilities.formatDate(ora, 'Europe/Rome', 'yyyy-MM-dd_HH-mm-ss') + '.json';
  cartella.createFile(nome, JSON.stringify({
    salvata: oraLeggibile_(), motivo: motivo, pubblicazioni: pubblicazioni, files: files
  }), MimeType.PLAIN_TEXT);

  // tiene solo le copie più recenti
  const copie = elencoCopie();
  copie.slice(COPIE_DA_TENERE).forEach(function (c) { DriveApp.getFileById(c.id).setTrashed(true); });
  return nome;
}

function oraLeggibile_() {
  return Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy HH:mm');
}
