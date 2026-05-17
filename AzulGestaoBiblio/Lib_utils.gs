var CLIENT_SPREADSHEET_ID = "";

function setSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId || typeof spreadsheetId !== "string") {
    throw new Error("spreadsheetId invalide : " + JSON.stringify(spreadsheetId));
  }

  CLIENT_SPREADSHEET_ID = spreadsheetId.trim();
}

function getSS() {
  if (CLIENT_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CLIENT_SPREADSHEET_ID);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  throw new Error("Aucun Google Sheet lie. spreadsheetId manquant.");
}

function invalidateProductsCache() {
  return true;
}
