const CALENDAR_NAME = "Partite - AC";
const MAJOR_VERSION = 0;
const MINOR_VERSION = 8;
const PATCH_VERSION = 3;
const CURRENT_VERSION = `${MAJOR_VERSION}.${MINOR_VERSION}.${PATCH_VERSION}`;
const githubUrl = "https://github.com/matteocheccacci/AutoCalendar-for-TBT-FipavOnline-SOL-and-CED-Federvolley";

// Filtro Gmail condiviso da sincronizzazione automatica e manuale.
const MAIL_FILTER = [
  'from:tiebreaktech.com',
  'from:federvolley.it',
  'subject:"Designazione gara"',
  'subject:"VARIAZIONE GARA"',
  'subject:"SPOSTAMENTO GARA"',
  'subject:"modifiche alla gara"'
].join(" OR ");

const WHATS_NEW_TEXT = `
  <p><b>Versione 0.8.2</b></p>
  <ul>
    <li><b>Bot Telegram (opzionale):</b> AutoCalendar può scriverti su Telegram. Le notifiche restano attive anche senza il bot, questa è solo una via in più.</li>
    <li><b>Elenco delle nuove designazioni:</b> a ogni sincronizzazione automatica che trova qualcosa, il bot ti manda un messaggio con le gare aggiunte (e quelle variate), complete di data, ora, impianto, ruolo e codici referto.</li>
    <li><b>Promemoria anche su Telegram:</b> ogni promemoria del calendario ti arriva anche in chat, con gli stessi orari (24 ore prima, la mattina della gara alle 8:00 e le tue notifiche personalizzate). La mattina della gara ricevi l'elenco delle partite di giornata.</li>
    <li><b>Il bot è tuo:</b> lo crei in un minuto con <b>@BotFather</b> (<i>/newbot</i>, un nome, uno username che finisca per "bot") e la configurazione ti guida passo passo. Nessun bot condiviso con altri.</li>
    <li><b>Abbinamento in 3 passaggi:</b> incolli il token nella configurazione, lo script genera una key, tu scrivi <b>/start</b> al bot e incolli la key. Da quel momento il bot scrive <b>solo alla tua chat</b>.</li>
    <li><b>Si disattiva quando vuoi:</b> riapri <i>Configura Bot Telegram</i>, lascia il campo del token vuoto e premi OK.</li>
  </ul>
  <p><b>Versione 0.8.3</b></p>
  <ul>
    <li><b>Nome repository aggiornato:</b> il progetto su GitHub ha un nuovo nome; i link e il controllo aggiornamenti puntano al nuovo indirizzo.</li>
    <li><b>Stringa copyright aggiornata:</b> corrette le informazioni di copyright nella finestra <i>Informazioni su AutoCalendar</i>.</li>
  </ul>`;

const updtMailBody = `
    <p>È disponibile la nuova versione <b>${CURRENT_VERSION}</b> su GitHub.</p>
    <p>Installala copiando il codice da <a href="${githubUrl}">GitHub</a>.</p>
  `;

function getRawFileUrl_(fileName) {
  const base = githubUrl.replace("https://github.com/", "https://raw.githubusercontent.com/");
  return base + "/main/" + fileName + "?t=" + new Date().getTime();
}

function fetchRemoteCodeText_() {
  const urls = [getRawFileUrl_("Codice.gs"), getRawFileUrl_("Code.gs"), getRawFileUrl_("Code.js"), getRawFileUrl_("Codice.js")];
  let lastErr = null;
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = resp.getResponseCode();
      const text = resp.getContentText() || "";
      if (code >= 200 && code < 300 && /MAJOR_VERSION|MINOR_VERSION|PATCH_VERSION/.test(text)) {
        return { url, text, code };
      }
      lastErr = new Error("HTTP " + code + " for " + url);
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr || new Error("Impossibile scaricare le file remoto.");
}

function normalizeGaraId_(v) {
  if (v == null) return "";
  return String(v).toUpperCase().replace(/\s+/g, " ").replace(/\s*-\s*/g, " - ").trim();
}

function normalizeText_(s) {
  return String(s || "")
    .replace(/[\u2012\u2013\u2014\u2212]/g, "-")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s+/g, " ")
    .trim();
}

function pad2_(v) {
  return String(v).trim().length === 1 ? "0" + String(v).trim() : String(v).trim();
}

/**
 * Il ruolo che TU hai in quella gara (colonna L). Le mail dei comitati e del CED lo
 * dichiarano; quelle che elencano solo gli arbitri lo ricavano dal nome configurato.
 */
/**
 * Le mail inoltrate non vengono elaborate: nel foglio devono finire solo le TUE
 * designazioni, arrivate direttamente dal sistema federale alla tua casella.
 * La designazione di un collega girata per posta non e' tua.
 */
function isInoltrata_(subject, body) {
  if (/^\s*(fwd|fw|inoltr\w*|i)\s*:/i.test(String(subject || ""))) return true;
  return /-{3,}\s*Forwarded message|Begin forwarded message|messaggio inoltrato/i.test(String(body || ""));
}

function ruoloDesignato_(dataGara) {
  if (dataGara.ruolo) return normalizeText_(dataGara.ruolo);
  var myName = (PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "").trim().toLowerCase();
  if (!myName) return "";
  var a1 = String(dataGara.arb1 || "").trim().toLowerCase();
  var a2 = String(dataGara.arb2 || "").trim().toLowerCase();
  if (a1 && (a1.indexOf(myName) !== -1 || myName.indexOf(a1) !== -1)) return "1° Arbitro";
  if (a2 && (a2.indexOf(myName) !== -1 || myName.indexOf(a2) !== -1)) return "2° Arbitro";
  return "";
}

function capitalizeName_(name) {
  if (!name || typeof name !== 'string') return name;
  return name.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
}

function sortSheetByDateTime_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏐 AutoCalendar')
    .addItem('Sincronizzazione Manuale', 'manualSync')
    .addSeparator()
    .addItem('Configurazione Iniziale Completa', 'setupTrigger')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Impostazioni Singole')
      .addItem('Imposta il Tuo Nome', 'setUserName')
      .addItem('Imposta API Gemini', 'setGeminiKey')
      .addItem('Cambia Frequenza Aggiornamento', 'setFrequency')
      .addItem('Gestisci Invitati Calendario', 'setGuests')
      .addItem('Imposta Notifiche Extra', 'setExtraReminder')
      .addItem('Configura Bot Telegram', 'setTelegramBot')
      .addItem('Prova Notifica Telegram', 'telegramTest'))
    .addSeparator()
    .addItem('Verifica Aggiornamenti', 'checkUpdatesManual')
    .addItem('Novità', 'checkWhatsNew')
    .addItem('Info', 'showInfo')
    .addToUi();
  
  checkUpdates(false);
  Utilities.sleep(1000);
  checkWhatsNew(false);
}

function checkWhatsNew(isManual = true) {
  const props = PropertiesService.getScriptProperties();
  const lastVersion = props.getProperty('LAST_INSTALLED_VERSION');
  
  if (!isManual && lastVersion === CURRENT_VERSION) {
    return;
  }

  const html = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h2 style="color: #1a73e8;">✨ Novità Versione ${CURRENT_VERSION}</h2>
      ${WHATS_NEW_TEXT}
      <p style="font-size: 0.9em; color: #666;">Se non hai tempo ora, clicca il tasto sotto per ricevere le novità via email.</p>
      <div style="margin-top: 20px; text-align: center;">
        <input type="button" value="Inviamele via Email" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).sendWhatsNewEmail();" style="padding: 10px; background-color: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer;">
        <input type="button" value="Chiudi" onclick="google.script.host.close();" style="padding: 10px; margin-left: 10px; border: 1px solid #ccc; border-radius: 4px; cursor: pointer;">
      </div>
    </div>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(ui, "AutoCalendar - Novità");
  props.setProperty('LAST_INSTALLED_VERSION', CURRENT_VERSION);
}

function sendWhatsNewEmail() {
  const email = Session.getEffectiveUser().getEmail();
  const body = `<h2>AutoCalendar - Novità Versione ${CURRENT_VERSION}</h2>${WHATS_NEW_TEXT}<br><p>Segnala un Bug su GitHub: ${githubUrl}</p>`;
  GmailApp.sendEmail(email, `AutoCalendar: Novità versione ${CURRENT_VERSION}`, "", {htmlBody: body});
}

function checkUpdatesManual() { checkUpdates(true); }

function getUpdateAvailable() {
  try {
    const remote = fetchRemoteCodeText_();
    const content = remote.text;
    const mMaj = content.match(/const MAJOR_VERSION = (\d+);/);
    const mMin = content.match(/const MINOR_VERSION = (\d+);/);
    const mPat = content.match(/const PATCH_VERSION = (\d+);/);
    if (mMaj && mMin && mPat) {
      const rMaj = parseInt(mMaj[1], 10);
      const rMin = parseInt(mMin[1], 10);
      const rPat = parseInt(mPat[1], 10);
      const isNewer = rMaj > MAJOR_VERSION || (rMaj === MAJOR_VERSION && rMin > MINOR_VERSION) || (rMaj === MAJOR_VERSION && rMin === MINOR_VERSION && rPat > PATCH_VERSION);
      return { isNewer: isNewer, version: `${rMaj}.${rMin}.${rPat}` };
    }
  } catch (e) {}
  return { isNewer: false };
}

function checkUpdates(showIfUpdated) {
  const updateInfo = getUpdateAvailable();
  const ui = SpreadsheetApp.getUi();
  if (updateInfo.isNewer) {
    ui.alert('🚀 Aggiornamento Disponibile', `Nuova versione: ${updateInfo.version}\nVersione attuale: ${CURRENT_VERSION}\n\nScarica l'ultima versione da GitHub.`, ui.ButtonSet.OK);
  } else if (showIfUpdated) {
    ui.alert('🔝 Sei aggiornato!', `Versione attuale: ${CURRENT_VERSION}`, ui.ButtonSet.OK);
  }
}

function showInfo() {
  const anno = new Date().getFullYear();
  const versione = CURRENT_VERSION;
  const html = `<div style="font-family: sans-serif; line-height: 1.4; color: #333; text-align: center;">
    <h2 style="color: #1a73e8;">🏐 AutoCalendar</h2>
    <p><b>Versione:</b> ${versione}<br><b>Creatore:</b> Matteo Checcacci</p>
    <div style="font-size: 0.8em; background: #f9f9f9; padding: 10px; border-radius: 5px; text-align: left; margin: 10px 0;">
      <strong>© Info Copyright:</strong><br>
      • TBT, FWM, Fipav Web Manager sono copyright di TieBreakTech.<br>
      • FipavOnline è copyright di Manufacturing Point Software.<br>
      • SOL / SOL arbitri e i sistemi del CED FIPAV sono copyright della Federazione Italiana Pallavolo (FIPAV).

    </div>
    <p style="font-size: 0.85em;">Prodotto concesso in <b>Licenza MIT</b>.</p>
    <div style="margin: 15px 0;">
      <a href="${githubUrl}" target="_blank" style="background-color: #24292e; color: white; padding: 8px 16px; text-decoration: none; border-radius: 5px; font-weight: bold;">GitHub / Segnala Bug</a>
    </div>
    <hr style="border: 0; border-top: 1px solid #eee;">
    <p style="font-size: 0.75em; color: #666;">© ${anno} KekkoTech Softwares - <a href="https://refpublic.it/" target="_blank">RefPublic</a></p>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(350), 'Informazioni su AutoCalendar');
}

function setUserName() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "Non impostato";
  var res = ui.prompt('Impostazione Nome', 'Valore attuale: ' + current + '\n\nInserisci COGNOME NOME:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('USER_FULL_NAME', v);
  }
}

function setGeminiKey() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ? "Chiave presente (nascosta)" : "Non impostata";
  var res = ui.prompt('Impostazione API Gemini', 'Stato attuale: ' + current + '\n\nInserisci la nuova API Key:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var v = (res.getResponseText() || "").trim();
    if (v) PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', v);
  }
}

function setFrequency() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var currentFreq = "Non impostata";
  triggers.forEach(t => { if (t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') {
    currentFreq = "Trigger attivo"; 
  }});
  var res = ui.prompt('Frequenza Aggiornamento', 'Stato: ' + currentFreq + '\n\nInserisci ore (1, 2, 4, 6, 8, 12):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var freq = parseInt((res.getResponseText() || "").trim(), 10);
    var allowed = { 1: true, 2: true, 4: true, 6: true, 8: true, 12: true };
    if (!allowed[freq]) {
      ui.alert('Valore non valido', 'Inserisci solo: 1, 2, 4, 6, 8, 12', ui.ButtonSet.OK);
      return;
    }
    triggers.forEach(t => { if (t.getHandlerFunction() == 'syncGmailToSheetAndCalendar') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('syncGmailToSheetAndCalendar').timeBased().everyHours(freq).create();
  }
}

function setGuests() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS') || "Nessuno";
  var res = ui.prompt('Gestione Invitati', 'Attuali: ' + current + '\n\nInserisci Email (separate da virgola):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('CALENDAR_GUESTS', (res.getResponseText() || "").trim());
  }
}

function setExtraReminder() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('EXTRA_REMINDER_MINUTES') || "Nessuno";
  var res = ui.prompt('Notifiche Extra Personalizzate', 'Valori attuali (minuti prima): ' + current + '\n\nInserisci i minuti di anticipo separati dalla virgola (es: 60, 120, 1440):', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() == ui.Button.OK) {
    var val = (res.getResponseText() || "").trim();
    if (val) {
      PropertiesService.getScriptProperties().setProperty('EXTRA_REMINDER_MINUTES', val);
    }
  }
}

function setupTrigger() {
  setUserName();
  setGeminiKey();
  setFrequency();
  setGuests();
  setExtraReminder();
  setTelegramBot();
  getOrCreateCalendar();
  SpreadsheetApp.getUi().alert('✅ Configurazione completata.');
}

function manualSync() {
  const summary = syncGmailToSheetAndCalendar('(' + MAIL_FILTER + ') newer_than:30d');
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '✅ Sincronizzazione terminata',
    `Gare aggiunte a Excel: ${summary.added}\nGare aggiornate/spostate sul calendario: ${summary.moved}`,
    ui.ButtonSet.OK
  );
}

function syncGmailToSheetAndCalendar(customQuery) {
  // I trigger temporizzati passano un oggetto evento come primo argomento:
  // senza questo controllo la query di ricerca diventa un oggetto e la sincronizzazione automatica non trova nulla.
  var defaultQuery = 'is:unread (' + MAIL_FILTER + ') newer_than:30d';
  var query = (typeof customQuery === "string" && customQuery) ? customQuery : defaultQuery;
  var threads = GmailApp.search(query);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  if (!sheet.getRange(1, 12).getValue()) sheet.getRange(1, 12).setValue("Ruolo");
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var cal = getOrCreateCalendar();

  var addedCount = 0;
  var addedRows = [];
  var updatedRows = [];
  var isAuto = !(typeof customQuery === "string" && customQuery);
  var allMsgs = [];
  threads.forEach(thread => { thread.getMessages().forEach(m => allMsgs.push(m)); });
  allMsgs.sort((a, b) => a.getDate().getTime() - b.getDate().getTime());

  allMsgs.forEach(msg => {
    var htmlBody = msg.getBody();
    var from = (msg.getFrom() || "").toLowerCase();
    var subject = (msg.getSubject() || "").toUpperCase();
    if (isInoltrata_(msg.getSubject(), htmlBody)) return;
    var isRegionale = from.includes("tiebreaktech") || htmlBody.includes("TieBreak");

    var dataGara = null;
    if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO") || subject.includes("AUTORIZZAZIONE")) {
      dataGara = isRegionale ? parseSpostamentoTBT(htmlBody) : parseSpostamentoFipavOnline(htmlBody);
    } else {
      dataGara = isRegionale ? parseRegionaleStandard(htmlBody) : (parseFipavNazionale(htmlBody) || parseComitatoFipav(htmlBody) || parseTerritorialeStandard(htmlBody));
    }

    if (apiKey) {
      var aiData = callGeminiAI(htmlBody, apiKey);
      if (aiData) {
        if (!dataGara) dataGara = aiData;
        else {
          if (aiData.ora) dataGara.ora = aiData.ora;
          if (aiData.codA) dataGara.codA = aiData.codA;
          if (aiData.codF) dataGara.codF = aiData.codF;
          if (aiData.luogo) dataGara.luogo = aiData.luogo;
          if (aiData.squadraCasa) dataGara.squadraCasa = aiData.squadraCasa;
          if (aiData.squadraOspite) dataGara.squadraOspite = aiData.squadraOspite;
          if (aiData.data) dataGara.data = aiData.data;
          if (aiData.numeroGara) dataGara.numeroGara = aiData.numeroGara;
          if (aiData.categoria) dataGara.categoria = aiData.categoria;
          if (aiData.arb1) dataGara.arb1 = aiData.arb1;
          if (aiData.arb2) dataGara.arb2 = aiData.arb2;
        }
      }
    }

    if (dataGara && dataGara.numeroGara) {
      var dataValues = sheet.getDataRange().getValues();
      var rowIndex = -1;
      var targetId = normalizeGaraId_(dataGara.numeroGara);

      for (var i = 1; i < dataValues.length; i++) {
        var sheetId = normalizeGaraId_(dataValues[i][6]);
        if (sheetId && (sheetId === targetId || sheetId.includes(targetId) || targetId.includes(sheetId))) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        var oldRow = dataValues[rowIndex - 1];

        if (subject.includes("VARIAZIONE") || subject.includes("SPOSTAMENTO") || subject.includes("AUTORIZZAZIONE")) {
          deleteCalendarEventById_(cal, targetId);
        }

        var finalCodA = (dataGara.codA && dataGara.codA !== "" && dataGara.codA !== "-") ? dataGara.codA : oldRow[7];
        var finalCodF = (dataGara.codF && dataGara.codF !== "" && dataGara.codF !== "-") ? dataGara.codF : oldRow[8];

        var rowData = [
          dataGara.data || oldRow[0],
          dataGara.ora || oldRow[1],
          dataGara.luogo || oldRow[2],
          dataGara.squadraCasa || oldRow[3],
          dataGara.squadraOspite || oldRow[4],
          dataGara.categoria || oldRow[5],
          oldRow[6],
          finalCodA,
          finalCodF,
          capitalizeName_(dataGara.arb1 || oldRow[9]),
          capitalizeName_(dataGara.arb2 || oldRow[10]),
          ruoloDesignato_(dataGara) || oldRow[11] || ""
        ];
        
        var targetRange = sheet.getRange(rowIndex, 1, 1, rowData.length);
        sheet.getRange(rowIndex, 7, 1, 3).setNumberFormat("@");
        targetRange.setValues([rowData]);
        updatedRows.push(rowData);
        
      } else if (!subject.includes("VARIAZIONE") && !subject.includes("SPOSTAMENTO") && !subject.includes("AUTORIZZAZIONE")) {
        // Senza data, ora e squadre non e' una designazione: meglio nessuna riga che una
        // riga di spazzatura (e la mail resta da leggere, cosi' la si puo' controllare).
        if (!dataGara.data || !dataGara.ora || !dataGara.squadraCasa) return;
        var rowDataNew = [
          dataGara.data, dataGara.ora, dataGara.luogo,
          dataGara.squadraCasa, dataGara.squadraOspite,
          dataGara.categoria, dataGara.numeroGara,
          isRegionale ? "-" : (dataGara.codA || ""),
          isRegionale ? "-" : (dataGara.codF || ""),
          capitalizeName_(dataGara.arb1 || ""),
          capitalizeName_(dataGara.arb2 || ""),
          ruoloDesignato_(dataGara)
        ];
        var newRowIndex = sheet.getLastRow() + 1;
        // Formato testo prima della scrittura: altrimenti i numeri gara con lo zero
        // iniziale (es. 01018 delle serie nazionali) diventano numeri e lo perdono.
        sheet.getRange(newRowIndex, 7, 1, 3).setNumberFormat("@");
        sheet.getRange(newRowIndex, 1, 1, rowDataNew.length).setValues([rowDataNew]);
        addedRows.push(rowDataNew);
        addedCount++;
      }
      msg.markRead();
    }
  });

  SpreadsheetApp.flush();
  sortSheetByDateTime_();
  const movedCount = createCalendarEvents();
  if (isAuto) { try { telegramNotifySync_(addedRows, updatedRows); } catch (e) {} }
  try { checkUpdatesAutomated(); } catch (e) {}
  return { added: addedCount, moved: movedCount };
}

function deleteCalendarEventById_(cal, garaId) {
  if (!garaId) return;
  var targetId = normalizeGaraId_(garaId);

  var now = new Date();
  var startSearch = new Date(); startSearch.setDate(now.getDate() - 60);
  var endSearch = new Date(); endSearch.setMonth(now.getMonth() + 6);

  var events = cal.getEvents(startSearch, endSearch);
  events.forEach(ev => {
    var desc = ev.getDescription() || "";
    var match = desc.match(/N\. Gara:\s*([^\n\r]+)/i) || desc.match(/\[GARA:([^\]]+)\]/);
    if (match) {
      var eventId = normalizeGaraId_(match[1]);
      if (eventId === targetId) {
        ev.deleteEvent();
      }
    }
  });
}

function createCalendarEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  var cal = getOrCreateCalendar();
  var data = sheet.getDataRange().getValues();
  var extraReminder = PropertiesService.getScriptProperties().getProperty('EXTRA_REMINDER_MINUTES');

  var now = new Date();
  var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);

  var lookbackLimit = new Date();
  lookbackLimit.setDate(lookbackLimit.getDate() - 30);

  var futureLimit = new Date();
  futureLimit.setMonth(futureLimit.getMonth() + 6);

  var calId = cal.getId();
  var tz = cal.getTimeZone() || ss.getSpreadsheetTimeZone() || "Europe/Rome";

  // Rimuove eventuali videochiamate Meet aggiunte automaticamente in passato
  stripMeetFromRange_(calId, lookbackLimit, futureLimit);

  var allEvents = cal.getEvents(lookbackLimit, futureLimit);

  var movedCount = 0;
  var eventsMap = {};

  allEvents.forEach(ev => {
    var desc = ev.getDescription() || "";
    var match = desc.match(/N\. Gara:\s*([^\n\r]+)/i) || desc.match(/\[GARA:([^\]]+)\]/);
    if (match) {
      var id = normalizeGaraId_(match[1]);
      if (!eventsMap[id]) {
        eventsMap[id] = ev;
      } else {
        ev.deleteEvent();
      }
    }
  });

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var garaIdRaw = row[6];
    if (!garaIdRaw) continue;

    var garaIdNorm = normalizeGaraId_(garaIdRaw);
    var dOnly = parseDateCell(row[0]);
    var time = parseTimeCell(row[1]);

    if (!dOnly || !time) continue;

    if (dOnly < yesterday) {
      delete eventsMap[garaIdNorm];
      continue;
    }

    var start = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), time.hh, time.mm);
    var end = new Date(start.getTime() + 9000000);
    var title = "🏐 " + row[3] + " vs " + row[4];
    var searchTag = "N. Gara: " + garaIdRaw;

    var descLines = [searchTag, "Categoria: " + row[5]];
    if (row[11] && String(row[11]).trim() !== "") descLines.push("Ruolo: " + row[11]);
    var arb1 = String(row[9]).trim(), arb2 = String(row[10]).trim();
    var myName = (PropertiesService.getScriptProperties().getProperty('USER_FULL_NAME') || "").toLowerCase();
    var amI1 = myName !== "" && (arb1.toLowerCase().includes(myName) || myName.includes(arb1.toLowerCase()));
    var amI2 = myName !== "" && (arb2.toLowerCase().includes(myName) || myName.includes(arb2.toLowerCase()));

    if (amI1) { if (arb2 && arb2 !== "-" && arb2 !== "") descLines.push("Secondo arbitro: " + arb2); }
    else if (amI2) { if (arb1 && arb1 !== "-" && arb1 !== "") descLines.push("Primo arbitro: " + arb1); }
    else {
      if (arb1 && arb1 !== "-" && arb1 !== "") descLines.push("Primo arbitro: " + arb1);
      if (arb2 && arb2 !== "-" && arb2 !== "") descLines.push("Secondo arbitro: " + arb2);
    }

    if (row[7] && row[7] !== "-" && row[7] !== "") descLines.push("Codice attivazione: " + row[7]);
    if (row[8] && row[8] !== "-" && row[8] !== "") descLines.push("Codice firma: " + row[8]);

    var finalDesc = descLines.join("\n");

    if (eventsMap[garaIdNorm]) {
      var ev = eventsMap[garaIdNorm];
      if (ev.getStartTime().getTime() !== start.getTime() || ev.getLocation() !== row[2] || ev.getTitle() !== title || ev.getDescription() !== finalDesc) {
        movedCount++;
        ev.setTime(start, end);
        ev.setLocation(row[2]);
        ev.setDescription(finalDesc);
        ev.setTitle(title);
        
        ev.removeAllReminders();
        buildReminderOverrides_(start, dOnly, extraReminder).forEach(function(o) {
          ev.addPopupReminder(o.minutes);
        });
      }
      delete eventsMap[garaIdNorm];
    } else {
      var guests = PropertiesService.getScriptProperties().getProperty('CALENDAR_GUESTS');
      var overrides = buildReminderOverrides_(start, dOnly, extraReminder);
      createEventWithoutMeet_(calId, tz, {
        title: title,
        start: start,
        end: end,
        location: row[2],
        description: finalDesc,
        guests: guests,
        overrides: overrides
      });
    }
  }

  for (var id in eventsMap) {
    eventsMap[id].deleteEvent();
  }

  return movedCount;
}

/**
 * Chiamata generica all'API REST di Google Calendar.
 * Usa il token OAuth del progetto: non serve abilitare servizi avanzati.
 */
function calendarApiFetch_(method, path, payload, params) {
  var url = "https://www.googleapis.com/calendar/v3" + path;
  if (params) {
    var qs = Object.keys(params).map(function(k) {
      return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]);
    }).join("&");
    if (qs) url += (url.indexOf("?") === -1 ? "?" : "&") + qs;
  }
  var options = {
    method: method,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error("Calendar API " + code + ": " + text);
  return text ? JSON.parse(text) : {};
}

/**
 * Costruisce i promemoria popup: 24h prima, alle 08:00 del giorno gara ed extra personalizzati.
 * Google accetta al massimo 5 override per evento.
 */
function buildReminderOverrides_(start, dOnly, extraReminder) {
  var minutes = [1440];
  var morning = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), 8, 0);
  var rem = Math.round((start.getTime() - morning.getTime()) / 60000);
  if (rem > 0) minutes.push(rem);

  if (extraReminder) {
    extraReminder.split(',').forEach(function(m) {
      var val = parseInt(String(m).trim(), 10);
      if (!isNaN(val) && val > 0) minutes.push(val);
    });
  }

  var seen = {};
  var overrides = [];
  minutes.forEach(function(v) {
    if (v > 0 && v <= 40320 && !seen[v] && overrides.length < 5) {
      seen[v] = true;
      overrides.push({ method: "popup", minutes: v });
    }
  });
  return overrides;
}

/**
 * Crea l'evento SENZA videochiamata Meet.
 * conferenceDataVersion=1 senza conferenceData impedisce l'aggiunta automatica del Meet
 * quando l'evento ha degli invitati.
 */
function createEventWithoutMeet_(calId, tz, opt) {
  var attendees = String(opt.guests || "")
    .split(",")
    .map(function(e) { return e.trim(); })
    .filter(function(e) { return e !== ""; })
    .map(function(e) { return { email: e }; });

  var body = {
    summary: opt.title,
    location: opt.location ? String(opt.location) : "",
    description: opt.description,
    start: { dateTime: Utilities.formatDate(opt.start, tz, "yyyy-MM-dd'T'HH:mm:ss"), timeZone: tz },
    end: { dateTime: Utilities.formatDate(opt.end, tz, "yyyy-MM-dd'T'HH:mm:ss"), timeZone: tz },
    reminders: { useDefault: false, overrides: opt.overrides }
  };
  if (attendees.length > 0) body.attendees = attendees;

  try {
    return calendarApiFetch_(
      "post",
      "/calendars/" + encodeURIComponent(calId) + "/events",
      body,
      { conferenceDataVersion: 1, sendUpdates: "none" }
    );
  } catch (e) {
    // Fallback di sicurezza: se l'API REST non risponde, l'evento viene comunque creato.
    var cal = CalendarApp.getCalendarById(calId);
    var params = { location: opt.location, description: opt.description };
    if (attendees.length > 0) params.guests = attendees.map(function(a) { return a.email; }).join(",");
    var ev = cal.createEvent(opt.title, opt.start, opt.end, params);
    ev.removeAllReminders();
    opt.overrides.forEach(function(o) { ev.addPopupReminder(o.minutes); });
    return null;
  }
}

/**
 * Elimina la videochiamata Meet dagli eventi già esistenti nell'intervallo indicato.
 */
function stripMeetFromRange_(calId, timeMin, timeMax) {
  try {
    var pageToken = null;
    do {
      var params = {
        timeMin: timeMin.toISOString(),
        timeMax: timeMax.toISOString(),
        singleEvents: true,
        maxResults: 250
      };
      if (pageToken) params.pageToken = pageToken;

      var resp = calendarApiFetch_("get", "/calendars/" + encodeURIComponent(calId) + "/events", null, params);
      (resp.items || []).forEach(function(item) {
        if (item.hangoutLink || item.conferenceData) {
          try {
            calendarApiFetch_(
              "patch",
              "/calendars/" + encodeURIComponent(calId) + "/events/" + encodeURIComponent(item.id),
              { conferenceData: null },
              { conferenceDataVersion: 1, sendUpdates: "none" }
            );
          } catch (e) {}
        }
      });
      pageToken = resp.nextPageToken;
    } while (pageToken);
  } catch (e) {}
}

function splitTeamsSmart_(teamLineRaw) {
  var teamLine = normalizeText_(teamLineRaw || "");
  if (!teamLine || teamLine.indexOf(" - ") === -1) return null;

  var parts = teamLine.split(" - ").map(p => normalizeText_(p)).filter(p => p);
  if (parts.length < 2) return null;
  if (parts.length === 2) return { casa: parts[0], ospite: parts[1] };

  var best = null;
  for (var k = 1; k < parts.length; k++) {
    var left = parts.slice(0, k).join(" - ").trim();
    var right = parts.slice(k).join(" - ").trim();
    if (!left || !right) continue;

    var score = 0;
    var leftTokens = left.split(" ").filter(Boolean).length;
    var rightTokens = right.split(" ").filter(Boolean).length;
    score += Math.abs(leftTokens - rightTokens) * 5;
    score += Math.abs(left.length - right.length) * 0.5;

    if (/^[A-Z\s\d]+$/.test(parts[k-1]) && /[a-z]/.test(parts[k])) {
      score += 150; 
    }

    var mid = Math.floor(parts.length / 2);
    score += Math.abs(k - mid) * 10;

    if (best == null || score < best.score) best = { score: score, casa: left, ospite: right };
  }

  return best ? { casa: best.casa, ospite: best.ospite } : null;
}

function parseTerritorialeStandard(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    var matchGara = text.match(/gara (\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];

    var mData = text.match(/del (\d{2}\/\d{2}\/\d{4})/);
    if (mData) res.data = mData[1];

    var mOra = text.match(/ore (\d{1,2}[:\.]\d{2})/i);
    if (mOra) res.ora = mOra[1].replace(".", ":");

    var lines = text.split('\n').map(s => normalizeText_(s)).filter(s => s.length > 0);
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (l.includes(" - ") && !l.toLowerCase().includes("girone") && !l.toLowerCase().includes("gara") && !l.toLowerCase().includes("campo:")) {
        var teamLine = l.split(/del \d|alle|ore/i)[0].trim();
        var pair = splitTeamsSmart_(teamLine);
        if (pair) {
          res.squadraCasa = pair.casa;
          res.squadraOspite = pair.ospite;
          break;
        }
      }
    }

    var mCampo = text.match(/Campo:(.*?)\n/i);
    res.luogo = mCampo ? normalizeText_(mCampo[1]) : "Vedi mail";

    res.codA = text.match(/CODICE ATTIVAZIONE REFERTO:\s*(\d+)/i) ? text.match(/CODICE ATTIVAZIONE REFERTO:\s*(\d+)/i)[1] : (text.match(/REFERTO:\s*(\d+)/) ? text.match(/REFERTO:\s*(\d+)/)[1] : "");
    res.codF = text.match(/CODICE FIRMA REFERTO:\s*(\d+)/i) ? text.match(/CODICE FIRMA REFERTO:\s*(\d+)/i)[1] : (text.match(/FIRMA REFERTO:\s*(\d+)/) ? text.match(/FIRMA REFERTO:\s*(\d+)/)[1] : "");

    res.arb1 = extractField(text, "Primo arbitro:") || extractField(text, "I Arbitro:");
    res.arb2 = extractField(text, "II Arbitro:") || extractField(text, "Secondo arbitro:");

    res.categoria = text.match(/gara \d+\s+(.*?)\s+-/i) ? normalizeText_(text.match(/gara \d+\s+(.*?)\s+-/i)[1]) : "Territoriale";
    return res;
  } catch (e) { return null; }
}

function parseRegionaleStandard(html) {
  var text = cleanEmail(html);
  var res = { codA: "-", codF: "-" };
  try {
    // Numero, categoria, data e ora dalla riga della gara. Ancorarsi a quella riga evita
    // di pescare l'ora di invio ("Date: ... alle ore 21:24") nelle mail inoltrate.
    var mLine = text.match(/Gara numero\s+(\d+)\s+(.*?)\s+del\s+(\d{1,2})\s*[-\/]\s*(\d{1,2})\s*[-\/]\s*(\d{4})(?:\s+(\d{1,2})[:.](\d{2}))?/i);
    if (mLine) {
      res.numeroGara = mLine[1];
      res.categoria = normalizeText_(mLine[2]);
      res.data = pad2_(mLine[3]) + "/" + pad2_(mLine[4]) + "/" + mLine[5];
      if (mLine[6]) res.ora = pad2_(mLine[6]) + ":" + mLine[7];
    } else {
      var matchGara = text.match(/Gara numero (\d+)/i);
      if (!matchGara) return null;
      res.numeroGara = matchGara[1];

      var mData = text.match(/(\d{2}\s*[-\/]\s*\d{2}\s*[-\/]\s*\d{4})/);
      if (mData) res.data = mData[1].replace(/\s/g, "").replace(/-/g, "/");

      var mCat = text.match(/Gara numero \d+\s+(.*?)\s+del/i);
      if (mCat) res.categoria = normalizeText_(mCat[1]);
    }

    if (!res.ora) {
      var mOra = text.match(/(\d{2}[:\.]\d{2})/);
      if (mOra) res.ora = mOra[1].replace(".", ":");
    }

    // I link (Maps, "Luogo evento", tracciamento) non sono righe di dati.
    var lines = text.split('\n').map(s => normalizeText_(s))
      .filter(s => s.length > 0 && !/^<?https?:\/\//i.test(s));

    // Riga dell'impianto: contiene un CAP isolato (\b\d{5}\b evita di agganciare
    // i numeri di telefono dei dirigenti, che sono cifre consecutive piu' lunghe).
    var capIndex = lines.findIndex(l => /\b\d{5}\b/.test(l) && !/^Gara numero/i.test(l));
    if (capIndex > 1) {
      // Nel corpo in testo semplice il CAP sta su una riga da solo: l'impianto e' quella sopra.
      var capAlone = /^\d{5}$/.test(lines[capIndex]);
      var impIdx = capAlone ? capIndex - 1 : capIndex;
      if (impIdx > 0) {
        var pair = splitTeamsSmart_(lines[impIdx - 1]);
        if (pair) {
          res.squadraCasa = pair.casa;
          res.squadraOspite = pair.ospite;
        }
        res.luogo = capAlone ? lines[impIdx] + " " + lines[capIndex] : lines[capIndex];
      }
    }

    // Fallback: prima riga con " - " dopo "Gara numero", l'impianto e' quella successiva.
    if (!res.squadraCasa) {
      var gIdx = lines.findIndex(l => /Gara numero \d+/i.test(l));
      for (var j = gIdx + 1; j < lines.length; j++) {
        if (/^(Ruolo|Ufficiali|Contatto)/i.test(lines[j])) continue;
        if (lines[j].indexOf(" - ") !== -1) {
          var pair2 = splitTeamsSmart_(lines[j]);
          if (pair2) {
            res.squadraCasa = pair2.casa;
            res.squadraOspite = pair2.ospite;
            if (!res.luogo && lines[j + 1]) res.luogo = lines[j + 1];
          }
          break;
        }
      }
    }

    // Il "°" dei ruoli a volte arriva illeggibile, quindi non lo pretendo.
    var mA1 = text.match(/\b1\s*[^\s\w]{0,3}\s*arbitro:\s*([^\n\r]+)/i);
    var mA2 = text.match(/\b2\s*[^\s\w]{0,3}\s*arbitro:\s*([^\n\r]+)/i);
    res.arb1 = mA1 ? normalizeText_(mA1[1]) : "";
    res.arb2 = mA2 ? normalizeText_(mA2[1]) : "";

    // "Ruolo: Delegato" (o 1° arbitro, osservatore, segnapunti...) -> colonna L.
    var mRuolo = text.match(/^Ruolo:\s*(.+)$/im);
    if (mRuolo) res.ruolo = normalizeText_(mRuolo[1]);
    return res;
  } catch (e) { return null; }
}

/**
 * Designazioni dei Comitati Territoriali FIPAV (sistema web-fipav / SOL arbitri,
 * es. Milano Monza Lecco). Formato a righe fisse e senza codici referto:
 *
 *   COGNOME NOME
 *   Designazione a svolgere le funzioni di 1° Arbitro.
 *   Gara n. 77 - 1^ Divisione - Femminile.
 *   Del 23/01/2026 (Venerdi) ore 21.15.
 *   SQUADRA CASA - SQUADRA OSPITE.
 *   <impianto>. <via>. <comune>.
 */
function parseComitatoFipav(html) {
  var text = cleanEmail(html);
  if (!/Designazione a svolgere le funzioni/i.test(text)) return null;

  var lines = text.split('\n').map(l => normalizeText_(l)).filter(l => l !== "");
  var res = { codA: "-", codF: "-" };

  // Riga data/ora: e' l'ancora della mail, tutto il resto si posiziona rispetto a lei.
  var delIdx = -1;
  for (var i = 0; i < lines.length; i++) {
    var mDel = lines[i].match(/^Del\s+(\d{1,2}\/\d{1,2}\/\d{4}).*?\bore\s*(\d{1,2}[:.]\d{2})/i);
    if (mDel) {
      res.data = mDel[1];
      res.ora = mDel[2].replace(".", ":");
      delIdx = i;
      break;
    }
  }
  if (delIdx === -1) return null;

  // "Gara n. 77 - 1^ Divisione - Femminile."
  for (var j = 0; j < delIdx; j++) {
    var mGara = lines[j].match(/^Gara n\.?\s*([A-Z0-9]+)\b\s*(?:-\s*(.*))?$/i);
    if (mGara) {
      res.numeroGara = mGara[1];
      if (mGara[2]) res.categoria = stripTrailingDot_(mGara[2]);
      break;
    }
  }
  if (!res.numeroGara) return null;

  var pair = splitTeamsSmart_(stripTrailingDot_(lines[delIdx + 1] || ""));
  if (pair) {
    res.squadraCasa = pair.casa;
    res.squadraOspite = pair.ospite;
  }

  // Impianto: le righe successive (nome palestra, via, comune) fino alla prima riga di servizio.
  var luogo = [];
  for (var k = delIdx + 2; k < lines.length && luogo.length < 4; k++) {
    var l = lines[k];
    if (/^(Telefono|Note|Si rammenta|Per Confermare|La commissione|Si autorizza|email generata|Da:|Date:|Subject:|To:)/i.test(l)) break;
    if (/^[-\s.>*]/.test(l)) break;
    if (/^<?https?:\/\//i.test(l)) continue; // link Maps aggiunti da Gmail nel testo semplice
    var v = stripTrailingDot_(l);
    if (v) luogo.push(v);
  }
  res.luogo = luogo.length > 0 ? luogo.join(", ") : "Vedi mail";

  // Ruolo e nome del designato: il nome e' la riga subito sopra la designazione.
  var rIdx = lines.findIndex(l => /Designazione a svolgere le funzioni/i.test(l));
  var ruolo = "";
  if (rIdx !== -1) {
    var mRole = lines[rIdx].match(/funzioni di\s+(.+)$/i);
    if (mRole) ruolo = stripTrailingDot_(mRole[1]);
  }

  var nome = "";
  if (rIdx > 0) {
    var cand = stripTrailingDot_(lines[rIdx - 1]);
    if (cand.indexOf("@") === -1 && /^[A-Za-zÀ-ÿ'\s.]{4,60}$/.test(cand)) nome = cand;
  }

  // Il ruolo va in colonna L. Le colonne arbitri restano vuote: questa mail dice solo
  // il TUO ruolo, non chi sono gli altri ufficiali (nome del destinatario incluso in
  // "nome", usato solo per capire di chi parla la mail).
  if (ruolo) res.ruolo = ruolo;

  return res;
}

// Ruoli non arbitrali citati nelle designazioni nazionali.
var FIPAV_NAZ_ROLES_ = ["Addetto VideoCheck", "Assistente ?VideoCheck", "Segnapunti", "Refertista", "Commissario[^0-9]*", "Giudice di linea ?\\d?"];

/**
 * Le colonne della riga squadre sono tagliate a 23 caratteri e seguite dalla sigla
 * della provincia, separata da 2+ spazi ("CUS CAGLIARI   CA").
 * ponytail: larghezza 23 dedotta dai campioni reali; se FIPAV cambia il tracciato
 * basta aggiornare la costante qui sotto.
 */
function cleanTeamCell_(v) {
  var raw = String(v == null ? "" : v).replace(/\s+$/, "");
  var name = normalizeText_(raw.split(/\s{2,}/)[0]);
  // Sigla provincia troncata dal taglio a 23 caratteri ("VOLLEY PARELLA TORINO T").
  if (raw.length >= 23 && /\s[A-Z]$/.test(name)) name = name.replace(/\s[A-Z]$/, "");
  return name;
}

/**
 * Designazioni FIPAV nazionali (CED Fipav, tessonline@federvolley.it, serie A/B):
 *
 *   Signor/a GIANNINI GABRIELE
 *   Le comunichiamo che è stato/a designato/a per la gara:
 *   Serie B1  /F numero 08176 Del 09.05.2026 Ore 16.00
 *   VOLLEY PARELLA TORINO T - ACROBATICA ALESSANDRIA  AL
 *   a TORINO            TO
 *   Campo PALESTRA S.E. MANZONI - CAMPO CENTRALE
 *   CORSO SVIZZERA, 63
 *   Sono designati: 1° Arbitro   GIANNINI GABRIELE
 *                   2° Arbitro   ROSI DARIO
 */
function parseFipavNazionale(html) {
  var lines = cleanEmail(html, true).split('\n');

  // Riga gara: "<categoria> numero <nr> Del <data> Ore <ora>". E' anche la firma del formato.
  var gIdx = -1, mGara = null;
  for (var i = 0; i < lines.length; i++) {
    mGara = lines[i].match(/^(.*?)\s+numero\s+([A-Z0-9]+)\s+Del\s+(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})\s+Ore\s+(\d{1,2})[.:](\d{2})/i);
    if (mGara) { gIdx = i; break; }
  }
  if (gIdx === -1) return null;

  var res = { codA: "-", codF: "-" };
  res.categoria = normalizeText_(mGara[1]);
  res.numeroGara = mGara[2];
  res.data = pad2_(mGara[3]) + "/" + pad2_(mGara[4]) + "/" + mGara[5];
  res.ora = pad2_(mGara[6]) + ":" + mGara[7];

  // Squadre: riga subito successiva.
  var sides = String(lines[gIdx + 1] || "").split(/\s+-\s+/);
  if (sides.length >= 2) {
    res.squadraCasa = cleanTeamCell_(sides[0]);
    res.squadraOspite = cleanTeamCell_(sides.slice(1).join(" - "));
  }

  // Impianto: riga "Campo ...", indirizzo sotto, citta' e provincia sopra ("a TORINO TO").
  var luogo = [];
  for (var c = gIdx + 1; c < lines.length; c++) {
    if (!/^Campo\s+\S/i.test(lines[c])) continue;
    luogo.push(normalizeText_(lines[c].replace(/^Campo\s+/i, "")));
    for (var a = c + 1; a < lines.length; a++) {
      var indirizzo = normalizeText_(lines[a]);
      if (!indirizzo) break;
      if (/^<?https?:\/\//i.test(indirizzo)) continue; // link Maps aggiunti da Gmail
      luogo.push(indirizzo);
      break;
    }
    for (var b = c - 1; b >= 0 && b >= c - 3; b--) {
      var mCity = normalizeText_(lines[b]).match(/^a\s+(.+)$/i);
      if (mCity) { luogo.push(normalizeText_(mCity[1])); break; }
    }
    break;
  }
  res.luogo = luogo.length > 0 ? luogo.join(", ") : "Vedi mail";

  // Ufficiali. Il "°" arriva spesso come carattere illeggibile, quindi non lo pretendo.
  var mSig = lines.join("\n").match(/Signor\/a\s+(.+)/i);
  var ioNome = mSig ? normalizeText_(mSig[1]) : "";
  var mioRuolo = "";
  var arbRe = /^(?:Sono designati:\s*)?([12])\s*[^\s\w]{0,3}\s*Arbitro\s+(.+)$/i;
  var altriRe = new RegExp("^(" + FIPAV_NAZ_ROLES_.join("|") + ")\\s+(.+)$", "i");

  for (var r = gIdx; r < lines.length; r++) {
    var line = normalizeText_(lines[r]);
    var mArb = line.match(arbRe);
    if (mArb) {
      var nomeArb = normalizeText_(mArb[2]);
      if (mArb[1] === "1") res.arb1 = nomeArb;
      else res.arb2 = nomeArb;
      if (ioNome && nomeArb.toUpperCase() === ioNome.toUpperCase()) mioRuolo = mArb[1] + "° Arbitro";
      continue;
    }
    var mAltro = line.match(altriRe);
    if (mAltro && ioNome && normalizeText_(mAltro[2]).toUpperCase() === ioNome.toUpperCase()) {
      mioRuolo = normalizeText_(mAltro[1]);
    }
  }

  if (mioRuolo) res.ruolo = mioRuolo;

  return res;
}

function parseSpostamentoFipavOnline(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    var matchGara = text.match(/gara:\s*([A-Z0-9\s\-]+)\s*e['\s]/i);
    res.numeroGara = matchGara ? normalizeText_(matchGara[1]) : null;

    var matchData = text.match(/al giorno:\s*\w+\s*(\d{2}\/\d{2}\/\d{4})/i);
    res.data = matchData ? matchData[1] : null;

    var matchOra = text.match(/ore:\s*(\d{1,2}[.:]\d{2})/i);
    res.ora = matchOra ? matchOra[1].replace('.', ':') : null;

    var lines = text.split('\n').map(s => normalizeText_(s));
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].toLowerCase().includes("ore:")) {
        if (lines[i + 1] && lines[i + 1].trim() !== "" && !lines[i + 1].toLowerCase().includes("mail generata")) {
          res.luogo = lines[i + 1].trim();
          break;
        }
      }
    }
    return res.numeroGara ? res : null;
  } catch (e) { return null; }
}

function parseSpostamentoTBT(html) {
  var text = cleanEmail(html);
  var res = {};
  try {
    // 1. Estrazione Numero Gara
    var matchGara = text.match(/modifiche alla gara n\.\s*(\d+)/i) || 
                    text.match(/Numero:\s*(\d+)/i) ||
                    text.match(/gara n\.\s*(\d+)/i);
    if (!matchGara) return null;
    res.numeroGara = matchGara[1];

    // 2. Estrazione Data e Ora Attuale (formato YYYY-MM-DD o DD/MM/YYYY)
    // Cerchiamo la riga "Data attuale:"
    // Attenzione: cleanEmail/normalizeText_ trasformano "2026-11-22" in "2026 - 11 - 22",
    // quindi i separatori vanno letti con \s*-\s*.
    var mIso = text.match(/Data attuale:\s*(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})\s*(\d{1,2}[:.]\d{2})/i);
    var mIta = text.match(/Data attuale:\s*(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})\s*(\d{1,2}[:.]\d{2})/i);

    if (mIso) {
      res.data = pad2_(mIso[3]) + "/" + pad2_(mIso[2]) + "/" + mIso[1];
      res.ora = mIso[4].replace(".", ":");
    } else if (mIta) {
      res.data = pad2_(mIta[1]) + "/" + pad2_(mIta[2]) + "/" + mIta[3];
      res.ora = mIta[4].replace(".", ":");
    }

    // 3. Estrazione Impianto Attuale
    var matchImpianto = text.match(/Impianto attuale:\s*([^\n\r]+)/i);
    if (matchImpianto) {
      res.luogo = normalizeText_(matchImpianto[1]);
    }

    return res.numeroGara ? res : null;
  } catch (e) {
    return null;
  }
}

var QP_CHARS_ = {
  "=C3=A0": "à", "=C3=A8": "è", "=C3=A9": "é", "=C3=AC": "ì", "=C3=B2": "ò",
  "=C3=B9": "ù", "=C2=B0": "°", "=C2=A0": " ", "=E2=80=99": "'", "=E2=80=93": "-",
  "=20": " ", "=3D": "=", "=EF=BF=BD": ""
};

var HTML_ENTITIES_ = {
  "&nbsp;": " ", "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"',
  "&apos;": "'", "&#39;": "'", "&#039;": "'", "&#176;": "°", "&#8217;": "'"
};

function cleanEmail(html, keepColumns) {
  var s = String(html || "")
    // I soft break Quoted-Printable vanno via per primi: possono spezzare anche i tag ("<br=\n>").
    .replace(/=\r?\n/g, '');
  Object.keys(QP_CHARS_).forEach(function(k) { s = s.split(k).join(QP_CHARS_[k]); });
  s = s
    // Via blocchi non testuali col loro contenuto: altrimenti il CSS finisce nel corpo
    // e "margin-bottom" diventa una riga con " - " dentro, cioe' una finta riga squadre.
    .replace(/<(style|script|head|title)\b[^>]*>[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6]|table)>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  Object.keys(HTML_ENTITIES_).forEach(function(k) { s = s.split(k).join(HTML_ENTITIES_[k]); });
  s = s
    .replace(/&#(\d{2,5});/g, function(_, n) { return String.fromCharCode(parseInt(n, 10)); })
    .replace(/\uFFFD/g, '')  // caratteri non decodificabili (le "è" e i "°" delle mail FIPAV)
    .replace(/\r/g, '\n')
    .replace(/\n\s*\n/g, '\n');
  // Rimuove i marcatori di citazione delle mail inoltrate (">" a inizio riga).
  // keepColumns: le mail a colonne fisse usano la spaziatura multipla come separatore
  // di campo, quindi si tolgono solo gli spazi ai bordi della riga.
  s = s.split('\n').map(function(line) {
    var l = String(line).replace(/^[>\s]+/, '');
    return keepColumns ? l.replace(/[\u2012\u2013\u2014\u2212]/g, "-").replace(/\s+$/, "") : normalizeText_(l);
  }).join('\n');
  return s;
}

function stripTrailingDot_(v) {
  return String(v == null ? "" : v).replace(/\s*\.\s*$/, "").trim();
}

function extractField(text, label) {
  var safeLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var m = text.match(new RegExp(safeLabel + "\\s*([^\\n\\r]+)", "i"));
  return m ? normalizeText_(m[1]) : "";
}

function parseDateCell(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
    if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  }
  return null;
}

function parseTimeCell(v) {
  if (v == null) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return { hh: v.getHours(), mm: v.getMinutes() };
  const m = String(v).trim().match(/^(\d{1,2})[:\.](\d{2})$/);
  if (!m) return null;
  return { hh: parseInt(m[1], 10), mm: parseInt(m[2], 10) };
}

function getOrCreateCalendar() {
  var c = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  return c.length > 0 ? c[0] : CalendarApp.createCalendar(CALENDAR_NAME, { timeZone: "Europe/Rome" });
}

function callGeminiAI(html, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`;
  const prompt = `Analizza HTML designazione volley. Estrai dati gara. Rispondi SOLO JSON: {"data":"DD/MM/YYYY","ora":"HH:MM","luogo":"...","squadraCasa":"...","squadraOspite":"...","categoria":"...","numeroGara":"...","codA":"...","codF":"...","arb1":"...","arb2":"..."}. Testo: ${html}`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { response_mime_type: "application/json", temperature: 0.1 } }),
      muteHttpExceptions: true
    });
    return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { return null; }
}

// ===================== BOT TELEGRAM (opzionale) =====================

/**
 * Chiamata all'API del bot. Senza token configurato non fa nulla: il bot e' opzionale.
 */
function telegramApi_(method, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) return null;
  var options = { muteHttpExceptions: true };
  if (payload) {
    options.method = "post";
    options.contentType = "application/json";
    options.payload = JSON.stringify(payload);
  }
  try {
    var resp = UrlFetchApp.fetch("https://api.telegram.org/bot" + token + "/" + method, options);
    var json = JSON.parse(resp.getContentText() || "{}");
    return json.ok ? json.result : null;
  } catch (e) {
    return null;
  }
}

function escapeTelegram_(v) {
  return String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Invia al solo utente abbinato. Torna true se il messaggio e' partito.
 */
function sendTelegram_(text) {
  var chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  if (!chatId || !text) return false;
  var res = telegramApi_("sendMessage", {
    chat_id: chatId,
    text: text,
    parse_mode: "HTML",
    disable_web_page_preview: true
  });
  return !!res;
}

/** Una gara come la vede Telegram: stessi dati dell'evento in calendario. */
function telegramGaraText_(row) {
  var d = parseDateCell(row[0]);
  var t = parseTimeCell(row[1]);
  var quando = (d ? pad2_(d.getDate()) + "/" + pad2_(d.getMonth() + 1) + "/" + d.getFullYear() : String(row[0] || "")) +
               (t ? " ore " + pad2_(t.hh) + ":" + pad2_(t.mm) : "");
  var lines = ["<b>" + escapeTelegram_(row[3]) + " - " + escapeTelegram_(row[4]) + "</b>", escapeTelegram_(quando)];
  if (row[5]) lines.push(escapeTelegram_(row[5]));
  if (row[11]) lines.push("Ruolo: " + escapeTelegram_(row[11]));
  if (row[2]) lines.push("📍 " + escapeTelegram_(row[2]));
  if (row[7] && row[7] !== "-") lines.push("Cod. attivazione: " + escapeTelegram_(row[7]));
  if (row[8] && row[8] !== "-") lines.push("Cod. firma: " + escapeTelegram_(row[8]));
  return lines.join("\n");
}

/** Messaggio dopo una sincronizzazione automatica che ha prodotto qualcosa. */
function telegramNotifySync_(addedRows, updatedRows) {
  if (!PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID')) return;
  var blocchi = [];
  if (addedRows && addedRows.length > 0) {
    blocchi.push("<b>🆕 Nuove designazioni (" + addedRows.length + ")</b>");
    addedRows.forEach(function(r) { blocchi.push(telegramGaraText_(r)); });
  }
  if (updatedRows && updatedRows.length > 0) {
    blocchi.push("<b>🔄 Gare aggiornate (" + updatedRows.length + ")</b>");
    updatedRows.forEach(function(r) { blocchi.push(telegramGaraText_(r)); });
  }
  if (blocchi.length > 0) sendTelegram_(blocchi.join("\n\n"));
}

function telegramReminderLabel_(instant, start) {
  var g0 = new Date(instant.getFullYear(), instant.getMonth(), instant.getDate());
  var g1 = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  var giorni = Math.round((g1.getTime() - g0.getTime()) / 86400000);
  if (giorni <= 0) return "🏐 Oggi";
  if (giorni === 1) return "⏰ Domani";
  return "⏰ Tra " + giorni + " giorni";
}

/**
 * I promemoria che cadono nella finestra (windowStart, now]. Usa gli stessi minuti di
 * buildReminderOverrides_, quindi Telegram e Calendario restano allineati per costruzione.
 */
function telegramDueReminders_(data, windowStart, now, extraReminder) {
  var due = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[6]) continue;
    var dOnly = parseDateCell(row[0]);
    var time = parseTimeCell(row[1]);
    if (!dOnly || !time) continue;
    var start = new Date(dOnly.getFullYear(), dOnly.getMonth(), dOnly.getDate(), time.hh, time.mm);
    if (start.getTime() < now.getTime()) continue;
    buildReminderOverrides_(start, dOnly, extraReminder).forEach(function(o) {
      var instant = new Date(start.getTime() - o.minutes * 60000);
      if (instant.getTime() > windowStart.getTime() && instant.getTime() <= now.getTime()) {
        due.push({ row: row, start: start, instant: instant });
      }
    });
  }
  return due;
}

/**
 * Trigger ogni 15 minuti: rispedisce su Telegram i promemoria del calendario appena scaduti.
 * ponytail: tick a finestra invece di un trigger per promemoria, perche' Apps Script ne
 * consente al massimo 20 per progetto. Precisione +/- 15 minuti, abbastanza per un avviso.
 */
function telegramReminderTick() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TELEGRAM_CHAT_ID')) return;

  var now = new Date();
  var last = parseInt(props.getProperty('TELEGRAM_LAST_TICK') || "0", 10);
  // Alla prima esecuzione guarda indietro 15 minuti; dopo un fermo prolungato massimo 2 ore,
  // cosi' non arriva una raffica di promemoria vecchi.
  var floor = now.getTime() - 7200000;
  var windowStart = new Date(Math.max(last || (now.getTime() - 900000), floor));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Arbitro") || ss.getSheets()[0];
  var due = telegramDueReminders_(sheet.getDataRange().getValues(), windowStart, now,
                                  props.getProperty('EXTRA_REMINDER_MINUTES'));

  if (due.length > 0) {
    var blocchi = due.map(function(d) {
      return telegramReminderLabel_(d.instant, d.start) + "\n" + telegramGaraText_(d.row);
    });
    sendTelegram_(blocchi.join("\n\n"));
  }
  props.setProperty('TELEGRAM_LAST_TICK', String(now.getTime()));
}

function ensureTelegramTrigger_() {
  var attivo = !!PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  var triggers = ScriptApp.getProjectTriggers();
  var presente = triggers.some(function(t) { return t.getHandlerFunction() === 'telegramReminderTick'; });
  if (attivo && !presente) {
    ScriptApp.newTrigger('telegramReminderTick').timeBased().everyMinutes(15).create();
  } else if (!attivo && presente) {
    triggers.forEach(function(t) {
      if (t.getHandlerFunction() === 'telegramReminderTick') ScriptApp.deleteTrigger(t);
    });
  }
}

/** Configurazione: token, generazione della key e abbinamento con /start. */
function setTelegramBot() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var stato = props.getProperty('TELEGRAM_CHAT_ID') ? "Bot abbinato ✅" :
              (props.getProperty('TELEGRAM_BOT_TOKEN') ? "Token presente, abbinamento da completare" : "Non configurato");

  var res = ui.prompt('Bot Telegram (opzionale)',
    'Stato: ' + stato + '\n\n' +
    'Serve un bot tuo, si crea in un minuto:\n' +
    '1) Su Telegram apri la chat con @BotFather\n' +
    '2) Scrivi /newbot\n' +
    '3) Dai un nome al bot (es. AutoCalendar Rossi)\n' +
    '4) Scegli uno username che finisca per "bot" (es. autocalendar_rossi_bot)\n' +
    '5) BotFather risponde con il token: copialo e incollalo qui sotto\n\n' +
    'Incolla il token (lascia vuoto e premi OK per disattivare le notifiche Telegram):',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() != ui.Button.OK) return;

  var token = (res.getResponseText() || "").trim();
  if (!token) {
    props.deleteProperty('TELEGRAM_BOT_TOKEN');
    props.deleteProperty('TELEGRAM_CHAT_ID');
    props.deleteProperty('TELEGRAM_PAIR_KEY');
    ensureTelegramTrigger_();
    ui.alert('Notifiche Telegram disattivate.');
    return;
  }

  props.setProperty('TELEGRAM_BOT_TOKEN', token);
  var me = telegramApi_("getMe");
  if (!me || !me.username) {
    props.deleteProperty('TELEGRAM_BOT_TOKEN');
    ui.alert('Token non valido',
      'Telegram non ha riconosciuto il token.\n\nControlla di aver copiato tutta la riga che ti ha dato @BotFather ' +
      '(è lunga, nella forma 1234567890:AA...). Se l\'hai perso, su @BotFather scrivi /token e scegli il tuo bot.',
      ui.ButtonSet.OK);
    return;
  }

  var key = "AC-" + Utilities.getUuid().replace(/-/g, "").substring(0, 6).toUpperCase();
  props.setProperty('TELEGRAM_PAIR_KEY', key);

  var conferma = ui.alert('Abbinamento',
    'Bot riconosciuto: @' + me.username + '\n\n1) Apri la chat con @' + me.username +
    '\n2) Scrivi /start\n3) Incolla in chat questa key:\n\n' + key +
    '\n\nFatto questo premi OK per completare l\'abbinamento.',
    ui.ButtonSet.OK_CANCEL);
  if (conferma != ui.Button.OK) return;

  if (telegramPair_()) {
    sendTelegram_("✅ <b>AutoCalendar abbinato</b>\nDa ora ricevi qui le nuove designazioni e i promemoria delle gare.");
    ensureTelegramTrigger_();
    ui.alert('✅ Abbinamento completato', 'Ti ho inviato un messaggio di conferma su Telegram.', ui.ButtonSet.OK);
  } else {
    ui.alert('Key non trovata',
      'Non ho ancora visto la key nella chat del bot. Controlla di aver scritto /start e incollato:\n\n' + key +
      '\n\nPoi riapri "Configura Bot Telegram" e reincolla lo stesso token.', ui.ButtonSet.OK);
  }
}

/**
 * Cerca la key nei messaggi ricevuti dal bot e memorizza la chat da notificare.
 * ponytail: getUpdates viene chiamato SENZA offset, cosi' i messaggi restano in coda e
 * l'abbinamento si puo' ripetere quante volte serve; Telegram tiene gli ultimi 100.
 */
function telegramPair_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('TELEGRAM_PAIR_KEY');
  if (!key) return false;
  var needle = key.replace(/[^A-Z0-9]/gi, "").toUpperCase();

  var updates = telegramApi_("getUpdates", { limit: 100, timeout: 0, allowed_updates: ["message"] });
  if (!updates || updates.length === 0) return false;

  for (var i = updates.length - 1; i >= 0; i--) {
    var msg = updates[i].message;
    if (!msg || !msg.text || !msg.chat) continue;
    var testo = String(msg.text).replace(/[^A-Z0-9]/gi, "").toUpperCase();
    if (testo.indexOf(needle) !== -1) {
      props.setProperty('TELEGRAM_CHAT_ID', String(msg.chat.id));
      props.deleteProperty('TELEGRAM_PAIR_KEY');
      return true;
    }
  }
  return false;
}

function telegramTest() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TELEGRAM_CHAT_ID')) {
    ui.alert('Bot non abbinato', 'Configura prima il bot da "Configura Bot Telegram".', ui.ButtonSet.OK);
    return;
  }
  var ok = sendTelegram_("🔔 <b>Prova notifica</b>\nAutoCalendar sta parlando con la chat giusta.");
  ui.alert(ok ? '✅ Messaggio inviato su Telegram.' : '❌ Invio fallito. Controlla il token e rifai l\'abbinamento.');
}

// ===================== fine blocco Telegram =====================

function checkUpdatesAutomated() {
  const props = PropertiesService.getScriptProperties();
  let count = parseInt(props.getProperty('UPDATE_CHECK_COUNT') || "0", 10);
  count++;
  if (count >= 4) {
    const updateInfo = getUpdateAvailable();
    if (updateInfo.isNewer) {
      const myEmail = Session.getEffectiveUser().getEmail();
      if (myEmail) {
        GmailApp.sendEmail(myEmail, "🚀 Aggiornamento AutoCalendar", "", { htmlBody: updtMailBody });
      }
    }
    count = 0;
  }
  props.setProperty('UPDATE_CHECK_COUNT', count.toString());
}
