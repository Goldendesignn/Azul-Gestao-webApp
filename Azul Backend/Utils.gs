let CLIENT_SPREADSHEET_ID = "";

function setClientSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId || typeof spreadsheetId !== "string") {
    throw new Error("spreadsheetId invalide : " + JSON.stringify(spreadsheetId));
  }

  CLIENT_SPREADSHEET_ID = spreadsheetId.trim();

  if (
    typeof AzulGestaoBiblio !== "undefined" &&
    typeof AzulGestaoBiblio.setSpreadsheetId === "function"
  ) {
    AzulGestaoBiblio.setSpreadsheetId(CLIENT_SPREADSHEET_ID);
  }
}

function getSpreadsheetBinding() {
  return CLIENT_SPREADSHEET_ID || "";
}

function getSS(spreadsheetId) {
  const id = spreadsheetId || CLIENT_SPREADSHEET_ID;

  if (!id) {
    throw new Error("Aucun spreadsheetId disponible");
  }

  return SpreadsheetApp.openById(id);
}
