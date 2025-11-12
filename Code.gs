/**
 * ======================================
 * Gmail → PDF → Drive + Sheet
 * Direkt mit GmailApp.getAttachment()
 * ======================================
 */

const SCRIPT_CONFIG = {
  TARGET_MAILBOX_EMAIL: "me",
  GMAIL_SEARCH_QUERY: "label:Name",
  DRIVE_FOLDER_ID: "",
  SPREADSHEET_ID: "",
  SHEET_NAME: "Posteingangsbuch",
  FILE_PREFIX: "PE_",
};

const PDF_MIME_TYPE = "application/pdf";

function processEmailsAndSavePdfs() {
  Logger.log("=== Skript gestartet ===");

  try {
    const folder = DriveApp.getFolderById(SCRIPT_CONFIG.DRIVE_FOLDER_ID);
    Logger.log(`Drive-Ordner gefunden: ${folder.getName()}`);

    const sheet = getOrCreateSpreadsheet(SCRIPT_CONFIG.SPREADSHEET_ID, SCRIPT_CONFIG.SHEET_NAME);
    let nextId = getNextIterativeId(sheet, SCRIPT_CONFIG.FILE_PREFIX.length);
    Logger.log(`Nächste fortlaufende ID: ${nextId}`);

    const threads = GmailApp.search(SCRIPT_CONFIG.GMAIL_SEARCH_QUERY, 0, 100);
    Logger.log(`Gefundene Threads: ${threads.length}`);

    const logData = [];

    threads.forEach((thread) => {
      const messages = thread.getMessages();

      messages.forEach((msg) => {
        const attachments = msg.getAttachments();

        attachments.forEach((att) => {
          if (att.getContentType() === PDF_MIME_TYPE) {
            const iterativeId = nextId.toString().padStart(3, "0");
            const newFileName = `${SCRIPT_CONFIG.FILE_PREFIX}${iterativeId}_${att.getName()}`;

            try {
              const savedFile = folder.createFile(att.copyBlob()).setName(newFileName);
              Logger.log(`PDF gespeichert: ${newFileName}`);

              logData.push([
                iterativeId,
                newFileName,
                msg.getDate(),
                msg.getSubject(),
                msg.getFrom(),
                savedFile.getUrl()
              ]);

              nextId++;
            } catch (e) {
              Logger.log(`Fehler beim Speichern der PDF '${att.getName()}': ${e}`);
            }
          }
        });
      });

      // Label entfernen
      const labelName = SCRIPT_CONFIG.GMAIL_SEARCH_QUERY.replace("label:", "");
      const label = GmailApp.getUserLabelByName(labelName);
      if (label) thread.removeLabel(label);
    });

    if (logData.length > 0) {
      appendLogData(sheet, logData);
      Logger.log(`Protokolliert: ${logData.length} Zeilen`);
    } else {
      Logger.log("Keine PDFs gespeichert.");
    }

    PropertiesService.getScriptProperties().setProperty("nextId", nextId.toString());
    Logger.log("=== Skript beendet ===");

  } catch (err) {
    Logger.log(`!!! Skriptfehler: ${err}`);
  }
}

/**
 * Spreadsheet holen oder erstellen
 */
function getOrCreateSpreadsheet(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet '${sheetName}' erstellt`);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID","Datei-Name","Datum","Betreff","Absender","Drive Link"]);
    sheet.getRange("A1:F1").setFontWeight("bold");
  }
  return sheet;
}

/**
 * Nächste ID ermitteln
 */
function getNextIterativeId(sheet, prefixLength) {
  const props = PropertiesService.getScriptProperties();
  let nextId = props.getProperty("nextId");
  if (nextId) return parseInt(nextId);

  if (sheet.getLastRow() > 1) {
    const fileNames = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat(); // Spalte 2: Datei-Name
    const lastId = fileNames.reduce((maxId, name) => {
      const idEndIndex = name.indexOf("_", prefixLength);
      if (idEndIndex > prefixLength) {
        const idStr = name.substring(prefixLength, idEndIndex);
        const curr = parseInt(idStr, 10);
        return Math.max(maxId, isNaN(curr) ? 0 : curr);
      }
      return maxId;
    }, 0);
    return lastId + 1;
  }
  return 1;
}

/**
 * Daten ins Sheet anhängen
 */
function appendLogData(sheet, data) {
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}
