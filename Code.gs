/**
 * ======================================
 * Gmail → PDF → Drive + Sheet
 * Direkt mit GmailApp.getAttachment()
 * ======================================
 */

const SCRIPT_CONFIG = {
  GMAIL_SEARCH_QUERY: "label:Name",
  DRIVE_FOLDER_ID: "",
  SPREADSHEET_ID: "",
  SHEET_NAME: "Posteingangsbuch",
  FILE_PREFIX: "PE_",
};

function processEmailsAndSavePdfs() {
  const lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds for other executions to finish
    lock.waitLock(30000);
  } catch (e) { 
    Logger.log("Could not obtain lock: " + e);
    return; 
  }

  try {
    const folder = DriveApp.getFolderById(SCRIPT_CONFIG.DRIVE_FOLDER_ID);
    // Force open by ID every time to refresh the connection
    const ss = SpreadsheetApp.openById(SCRIPT_CONFIG.SPREADSHEET_ID);
    const sheet = getOrCreateSheet(ss, SCRIPT_CONFIG.SHEET_NAME);

    const processedKeys = getProcessedKeysFromSheet(sheet);
    let nextId = getNextIterativeId(sheet);

    // Limit threads to prevent "Server Error" timeouts
    const threads = GmailApp.search(SCRIPT_CONFIG.GMAIL_SEARCH_QUERY, 0, 10).reverse();
    const labelName = SCRIPT_CONFIG.GMAIL_SEARCH_QUERY.replace("label:", "");
    const label = GmailApp.getUserLabelByName(labelName);

    if (threads.length === 0) return;

    threads.forEach((thread) => {
      const threadId = thread.getId().trim();
      const messages = thread.getMessages();
      let threadLogData = [];

      messages.forEach((msg) => {
        msg.getAttachments().forEach((att) => {
          // Safety: ensure it's a PDF
          if (att.getContentType() !== "application/pdf") return;

          const rawFileName = att.getName().trim();
          const bytes = att.getBytes();

          // MD5 Hash calculation
          const attHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
            .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0"))
            .join("");

          const combinedKey = attHash + "||" + threadId;
          if (processedKeys.has(combinedKey)) {
            Logger.log("Skipping duplicate: " + rawFileName);
            return;
          }

          try {
            const iterativeId = nextId.toString().padStart(3, "0");
            const newFileName = `${SCRIPT_CONFIG.FILE_PREFIX}${iterativeId}_${rawFileName}`;
            const savedFile = folder.createFile(att.copyBlob()).setName(newFileName);

            threadLogData.push([
              iterativeId,        // A
              newFileName,        // B
              msg.getDate(),      // C
              msg.getSubject(),   // D
              msg.getFrom(),      // E
              savedFile.getUrl(), // F
              attHash,            // G
              threadId            // H
            ]);

            processedKeys.add(combinedKey);
            nextId++;
          } catch (err) {
            Logger.log("File save error: " + err);
          }
        });
      });

      if (threadLogData.length > 0) {
        appendLogData(sheet, threadLogData);
      }

      // Remove label only after successful processing
      if (label) thread.removeLabel(label);
    });

  } catch (globalError) {
    Logger.log("Global Error: " + globalError.message);
  } finally {
    // Crucial: Always flush and release
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * Robust Sheet Getter
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID", "Datei-Name", "Datum", "Betreff", "Absender", "Drive Link", "Att-Hash", "Thread-ID"]);
    SpreadsheetApp.flush(); // Force creation before continuing
  }
  return sheet;
}

function getProcessedKeysFromSheet(sheet) {
  const keys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  const data = sheet.getRange(2, 7, lastRow - 1, 2).getValues();
  data.forEach(row => {
    if (row[0] && row[1]) keys.add(row[0].toString().trim() + "||" + row[1].toString().trim());
  });
  return keys;
}

function getNextIterativeId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastId = sheet.getRange(lastRow, 1).getValue();
    if (!isNaN(lastId) && lastId !== "") return parseInt(lastId, 10) + 1;
  }
  return 1;
}

function appendLogData(sheet, data) {
  // Use a fresh reference to the sheet to prevent "Sheet not found" errors
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}
