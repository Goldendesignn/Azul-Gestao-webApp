const ROUTES = {
  garantirPrimeiraInicializacao: data => AzulGestaoBiblio.garantirPrimeiraInicializacao(data),

  getDashboardData: data => AzulGestaoBiblio.getDashboardData(data),
  getProducts: data => AzulGestaoBiblio.getProducts(data),
  saveProductProfile: data => AzulGestaoBiblio.saveProductProfile(data),
  activarModoEdicaoPOS: data => AzulGestaoBiblio.activarModoEdicaoPOS(data),

  registarVenda: data => AzulGestaoBiblio.registarVenda(data),
  getVentes: data => AzulGestaoBiblio.getVentes(data),

  getFornecedorNames: data => AzulGestaoBiblio.getFornecedorNames(data),
  registarFornecedor: data => AzulGestaoBiblio.registarFornecedor(data),
  registarPagamentoForn: data => AzulGestaoBiblio.registarPagamentoForn(data),
  getResumoDettes: data => AzulGestaoBiblio.getResumoDettes(data),

  registarPagamentoClient: data => AzulGestaoBiblio.registarPagamentoClient(data),
  getClientFicheData: data => AzulGestaoBiblio.getClientFicheData(data),
  getClientDebt: data => AzulGestaoBiblio.getClientDebt(data),
  getFournDebt: data => AzulGestaoBiblio.getFournDebt(data),

  getStockArmazem: data => AzulGestaoBiblio.getStockArmazem(data),
  registarTransferencia: data => AzulGestaoBiblio.registarTransferencia(data),
  transferirTudo: data => AzulGestaoBiblio.transferirTudo(data),

  registarTresorerie: data => AzulGestaoBiblio.registarTresorerie(data),
  getTresorerie: data => AzulGestaoBiblio.getTresorerie(data),
  getComptabiliteData: data => AzulGestaoBiblio.getComptabiliteData(data),

  registarDepense: data => AzulGestaoBiblio.registarDepense(data),
  getDepenseDashboard: data => AzulGestaoBiblio.getDepenseDashboard(data),
  getHistoriqueDepenses: data => AzulGestaoBiblio.getHistoriqueDepenses(data),

  getRevendeurNames: data => AzulGestaoBiblio.getRevendeurNames(data),
  getRevendeurDetail: data => AzulGestaoBiblio.getRevendeurDetail(data),
  registarConsignacao: data => AzulGestaoBiblio.registarConsignacao(data),
  getConsignationsOpen: data => AzulGestaoBiblio.getConsignationsOpen(data),
  getConsignationsByRevendeur: data => AzulGestaoBiblio.getConsignationsByRevendeur(data),
  confirmerPaiementConsignation: data => AzulGestaoBiblio.confirmerPaiementConsignation(data),
  confirmerPaiementConsignations: data => AzulGestaoBiblio.confirmerPaiementConsignations(data),
  retornarConsignacao: data => AzulGestaoBiblio.retornarConsignacao(data),
  retornarConsignacoes: data => AzulGestaoBiblio.retornarConsignacoes(data),
  getHistoriqueConsignations: data => AzulGestaoBiblio.getHistoriqueConsignations(data),

  registarAchatMultiple: data => AzulGestaoBiblio.registarAchatMultiple(data),
  registarAchat: data => AzulGestaoBiblio.registarAchat(data),

  inicializarFeuilles: data => AzulGestaoBiblio.inicializarFeuilles(data),
  reinicializarInventaire: data => AzulGestaoBiblio.reinicializarInventaire(data),
  protegerFeuilles: data => AzulGestaoBiblio.protegerFeuilles(data),
  activarModoEdicao: data => AzulGestaoBiblio.activarModoEdicao(data),
  reinicializarTudo: data => AzulGestaoBiblio.reinicializarTudo(data),

  verification: data => verification(),
  saveUtilisateur: data => saveUtilisateur(data.nom, data.numero, data.email, data.licence, data.statut)
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  let spreadsheetId = String(params.spreadsheetId || "").trim();
  const cle = String(params.cle || params.licence || "").trim().toUpperCase().replace(/\s+/g, "");

  if (!spreadsheetId && cle) {
    spreadsheetId = getSpreadsheetIdByLicence(cle);
  }

  if (spreadsheetId && spreadsheetId !== "login") {
    setClientSpreadsheetId(spreadsheetId);

    const licence = verification(spreadsheetId); // ← passe l'ID directement

    if (licence && licence.ok === true) {
      // ✅ Direct vers POS_Core — ZERO init ici
      const template = HtmlService.createTemplateFromFile("POS_Core");
      template.spreadsheetId = spreadsheetId;
      return template.evaluate()
        .setTitle("Azul Gestao")
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
    }
  }

  const template = HtmlService.createTemplateFromFile("POS_login");
  template.spreadsheetId = "";
  return template.evaluate()
    .setTitle("Activation Azul Gestao")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

function apiCall(req) {
  try {
    req = req || {};

    const action = req.action;
    const spreadsheetId = String(req.spreadsheetId || "").trim();
    const data = req.data || {};

    if (!spreadsheetId && action !== "saveUtilisateur") {
      throw new Error("spreadsheetId manquant");
    }


    if (!ROUTES[action]) {
      throw new Error("Action inconnue : " + action);
    }

    if (spreadsheetId && spreadsheetId !== "login") {
      setClientSpreadsheetId(spreadsheetId);
    }

    return {
      ok: true,
      data: ROUTES[action](data)
    };

  } catch (err) {
    return {
      ok: false,
      message: err.message,
      stack: err.stack
    };
  }
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || "{}");
    return jsonResponse(apiCall(req));
  } catch (err) {
    return jsonResponse({
      ok: false,
      message: err.message,
      stack: err.stack
    });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
