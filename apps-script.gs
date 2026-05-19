/**
 * CamPùfood — Endpoint Apps Script per il form di contatto.
 *
 * SETUP (5 minuti):
 * 1) Apri il Google Sheet dove vuoi i lead (oppure creane uno nuovo).
 * 2) Menu: Estensioni → Apps Script. Cancella il codice esistente.
 * 3) Incolla TUTTO questo file e salva (icona dischetto). Dai un nome al progetto.
 * 4) (Opzionale) Cambia SHEET_NAME se vuoi un'altra scheda.
 * 5) Clicca "Distribuisci" → "Nuova distribuzione" → seleziona tipo "App web".
 *    - Descrizione: "CamPufood form"
 *    - Esegui come: "Me"
 *    - Chi ha accesso: "Chiunque"  (necessario per ricevere POST dal sito)
 *    - Clicca "Distribuisci" e autorizza l'accesso quando richiesto.
 * 6) Copia l'URL "App web" (https://script.google.com/macros/s/AKfycb.../exec).
 * 7) In index.html sostituisci la costante FORM_ENDPOINT con quell'URL.
 *
 * NOTE:
 * - Quando aggiorni il codice, fai "Distribuisci → Gestisci distribuzioni → ✎ → Versione: Nuova versione"
 *   per mantenere lo stesso URL. Se crei una "Nuova distribuzione" ottieni un URL diverso.
 * - Se vuoi anche ricevere un'email ad ogni lead, valorizza NOTIFY_EMAIL qui sotto.
 */

const SHEET_NAME   = 'Lead';                     // nome della scheda dove scrivere
const NOTIFY_EMAIL = '';                         // es: 'tu@campufood.it'  — vuoto = nessuna mail

const HEADERS = ['Data', 'Nome', 'Locale', 'Email', 'Telefono', 'Tipologia', 'Messaggio', 'Sorgente'];

function doPost(e) {
  try {
    const data = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : (e && e.parameter ? e.parameter : {});

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.nome      || '',
      data.locale    || '',
      data.email     || '',
      data.telefono  || '',
      data.tipologia || '',
      data.messaggio || '',
      data._source   || 'web'
    ]);

    if (NOTIFY_EMAIL) {
      const subject = `Nuovo lead CamPùfood — ${data.locale || data.nome || 'senza nome'}`;
      const body = [
        `Nome: ${data.nome || '-'}`,
        `Locale: ${data.locale || '-'}`,
        `Email: ${data.email || '-'}`,
        `Telefono: ${data.telefono || '-'}`,
        `Tipologia: ${data.tipologia || '-'}`,
        '',
        'Messaggio:',
        data.messaggio || '-'
      ].join('\n');
      try { MailApp.sendEmail(NOTIFY_EMAIL, subject, body); } catch (_) {}
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Endpoint GET solo per verificare che la Web App sia attiva (apri l'URL nel browser).
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, hint: 'POST JSON per registrare un lead' }))
    .setMimeType(ContentService.MimeType.JSON);
}
