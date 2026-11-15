/**
 * Zusagen der Hochzeitsseite in ein Google Sheet schreiben.
 *
 * Einrichtung:
 * 1. Neues Google Sheet anlegen, z. B. "Hochzeit Zusagen".
 * 2. Erweiterungen > Apps Script öffnen, diesen Code einfügen, speichern.
 * 3. Bereitstellen > Neue Bereitstellung > Typ "Web-App".
 *    Ausführen als: Ich. Zugriff: Jeder.
 * 4. Die angezeigte Web-App-URL in index.html bei RSVP_ENDPOINT eintragen.
 *
 * Jede Antwort ergibt eine Zeile pro Person (Erwachsene und Kinder),
 * mit gemeinsamer Antwort-ID, damit Haushalte zusammen bleiben.
 */
const SHEET_ID = "1zYOC-IX--6HA-ekeFgOxqkn54kqb3YHDI3ODQu_2j8g";
const SHEET_NAME = "Antworten";
const HEADER = [
  "Zeitpunkt", "Antwort-ID", "Zusage", "Typ", "Name", "Alter",
  "Abend", "Übernachtung", "Allergien", "Nachricht"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    const now = new Date();
    const id = Utilities.getUuid().slice(0, 8);
    const clean = v => String(v == null ? "" : v).slice(0, 500);

    const rows = [];
    (data.erwachsene || []).slice(0, 10).forEach(name => {
      rows.push([now, id, clean(data.zusage), "Erwachsen", clean(name), "",
        clean(data.abend), clean(data.uebernachtung), clean(data.allergien), clean(data.nachricht)]);
    });
    (data.kinder || []).slice(0, 10).forEach(kind => {
      rows.push([now, id, clean(data.zusage), "Kind", clean(kind.name), clean(kind.alter),
        clean(data.abend), clean(data.uebernachtung), clean(data.allergien), clean(data.nachricht)]);
    });
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
    }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
