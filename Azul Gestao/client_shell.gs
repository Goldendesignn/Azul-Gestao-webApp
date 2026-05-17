
function abrirERP() {
  const html = getHtmlFromBackend_("main");

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html)
      .setWidth(1400)
      .setHeight(900),
    "Azul Gestão"
  );
}

function getHtmlFromBackend_(page) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  const res = UrlFetchApp.fetch(BACKEND_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      action: "getHtml",
      spreadsheetId: spreadsheetId,
      data: { page: page }
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(res.getContentText());

  if (!result.ok) {
    throw new Error(result.message);
  }

  return result.data.html;
}