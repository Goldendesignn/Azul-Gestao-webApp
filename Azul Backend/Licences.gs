const LICENCE_SHEET_ID = "1vLdedHfSWlYu055Z4bVuLVPUowSnRGX2yO24OmvyqhM";
const LICENCE_SHEET_NAME = "feuille1";

function getLicenceSheet_() {
  const ss = SpreadsheetApp.openById(LICENCE_SHEET_ID);
  const sh = ss.getSheetByName(LICENCE_SHEET_NAME);

  if (!sh) {
    throw new Error("Feuille licence introuvable : " + LICENCE_SHEET_NAME);
  }

  return sh;
}


function getSpreadsheetIdByLicence(cle) {
  cle = String(cle || "").trim().toUpperCase().replace(/\s+/g, "");

  if (!cle) return "";

  const sh = getLicenceSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowCle = String(data[i][0] || "").trim().toUpperCase().replace(/\s+/g, "");
    const rowStatus = String(data[i][5] || "").trim().toLowerCase();
    const rowSpreadsheetId = String(data[i][6] || "").trim();

    if (rowCle === cle && rowStatus === "active" && rowSpreadsheetId) {
      return rowSpreadsheetId;
    }
  }

  return "";
}

function saveUtilisateur(nom, numero, email, licence, statut) {
  const sh = getLicenceSheet_();
  const data = sh.getDataRange().getValues();

  const cle = String(licence || "").trim().toUpperCase().replace(/\s+/g, "");

  if (!cle) {
    return {
      ok: false,
      message: "Chave de licença obrigatória."
    };
  }

  for (let i = 1; i < data.length; i++) {
    const rowCle = String(data[i][0] || "").trim().toUpperCase().replace(/\s+/g, "");
    const rowStatus = String(data[i][5] || "").trim().toLowerCase();
    const rowSpreadsheetId = String(data[i][6] || "").trim();

    if (rowCle === cle) {
      if (rowStatus !== "active") {
        return {
          ok: false,
          message: "Licença inativa. Solicite uma nova chave."
        };
      }

      let spreadsheetId = rowSpreadsheetId;

      if (!spreadsheetId) {
        const clientName = String(nom || data[i][2] || "Cliente").trim() || "Cliente";
        const ss = SpreadsheetApp.create("Azul Gestao - " + clientName + " - " + cle);
        spreadsheetId = ss.getId();

        sh.getRange(i + 1, 7).setValue(spreadsheetId);
      }

      sh.getRange(i + 1, 2).setValue(new Date());
      sh.getRange(i + 1, 3).setValue(nom || "");
      sh.getRange(i + 1, 4).setValue(email || "");
      sh.getRange(i + 1, 5).setValue(numero || "");
      sh.getRange(i + 1, 6).setValue("active");

      setClientSpreadsheetId(spreadsheetId);

      AzulGestaoBiblio.garantirPrimeiraInicializacao({
        stockMode: "boutique"
      });

      return {
        ok: true,
        message: "Licença ativada.",
        spreadsheetId: spreadsheetId
      };
    }
  }

  return {
    ok: false,
    message: "Chave de licença inválida."
  };
}

function verification() {
  const spreadsheetId = CLIENT_SPREADSHEET_ID;
  const sh = getLicenceSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowStatut = String(data[i][5] || "").trim().toLowerCase();
    const rowId = String(data[i][6] || "").trim();

    if (rowId === spreadsheetId && rowStatut === "active") {
      return {
        ok: true,
        message: "Licence active"
      };
    }
  }

  return {
    ok: false,
    message: "Licence non active ou non trouvée"
  };
}

function getOrCreateSpreadsheetIdByLicence(cle) {
  cle = String(cle || "").trim().toUpperCase().replace(/\s+/g, "");

  if (!cle) {
    throw new Error("Chave de licenca obrigatoria.");
  }

  const sh = getLicenceSheet_();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowCle = String(data[i][0] || "").trim().toUpperCase().replace(/\s+/g, "");
    const rowStatus = String(data[i][5] || "").trim().toLowerCase();
    const rowSpreadsheetId = String(data[i][6] || "").trim();
    const rowName = String(data[i][2] || "Cliente").trim() || "Cliente";

    if (rowCle === cle) {
      if (rowStatus === "inactive" && rowSpreadsheetId) {
        throw new Error("Licenca desativada. Contacte o suporte Azul Gestao.");
      }

      if (rowSpreadsheetId) {
        return rowSpreadsheetId;
      }

      const ss = SpreadsheetApp.create("Azul Gestao - " + rowName + " - " + cle);
      const newSpreadsheetId = ss.getId();

      sh.getRange(i + 1, 2).setValue(new Date());
      sh.getRange(i + 1, 6).setValue("active");
      sh.getRange(i + 1, 7).setValue(newSpreadsheetId);

      return newSpreadsheetId;
    }
  }

  throw new Error("Chave de licenca invalida.");
}
