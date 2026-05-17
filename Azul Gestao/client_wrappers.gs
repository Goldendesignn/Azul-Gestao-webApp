const BACKEND_URL = "https://script.google.com/macros/s/AKfycbxNHgshSfHC7jcTROmGNzDbQEz-j-zAesUaNd0QqxFN1iR-ZgxeAWXrNGR0LoBC6xsA/exec";

/* =========================
   BACKEND : LICENCE + HTML
========================= */

function callBackend(action, data) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  const response = UrlFetchApp.fetch(BACKEND_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      action: action,
      spreadsheetId: spreadsheetId,
      data: data || {}
    }),
    muteHttpExceptions: true
  });

  const text = response.getContentText();

  if (text.trim().startsWith("<")) {
    throw new Error("Backend retourne HTML au lieu de JSON. Vérifie URL /exec et redéploiement.");
  }

  const result = JSON.parse(text);

  if (!result.ok) {
    throw new Error(result.message || "Erreur backend");
  }

  return result.data;
}

function verification(data) {
  return callBackend("verification", data);
}

function saveUtilisateur(nom, numero, email, licence, statut) {
  return callBackend("saveUtilisateur", {
    nom: nom,
    numero: numero,
    email: email,
    licence: licence,
    statut: statut
  });
}

function getHtmlFromBackend_(page) {
  const result = callBackend("getHtml", { page: page });

  if (!result || !result.html) {
    throw new Error("HTML introuvable : " + page);
  }

  return result.html;
}

/* =========================
   ERP : BIBLIOTHÈQUE
========================= */

function getProducts(data) {
  return AzulGestaoBiblio.getProducts(data);
}

function saveProductProfile(data) {
  return AzulGestaoBiblio.saveProductProfile(data);
}

function getDashboardData(data) {
  return AzulGestaoBiblio.getDashboardData(data);
}

function registarVenda(data) {
  return AzulGestaoBiblio.registarVenda(data);
}

function getVentes(data) {
  return AzulGestaoBiblio.getVentes(data);
}

function registarAchatMultiple(data) {
  return AzulGestaoBiblio.registarAchatMultiple(data);
}

function getFornecedorNames(data) {
  return AzulGestaoBiblio.getFornecedorNames(data);
}

function registarFornecedor(data) {
  return AzulGestaoBiblio.registarFornecedor(data);
}

function registarPagamentoForn(data) {
  return AzulGestaoBiblio.registarPagamentoForn(data);
}

function getResumoDettes(data) {
  return AzulGestaoBiblio.getResumoDettes(data);
}

function registarPagamentoClient(data) {
  return AzulGestaoBiblio.registarPagamentoClient(data);
}

function getClientFicheData(data) {
  return AzulGestaoBiblio.getClientFicheData(data);
}

function getClientDebt(clientName) {
  return AzulGestaoBiblio.getClientDebt({
    clientName: clientName
  });
}

function getStockArmazem(data) {
  return AzulGestaoBiblio.getStockArmazem(data);
}

function registarTransferencia(data) {
  return AzulGestaoBiblio.registarTransferencia(data);
}

function transferirTudo(data) {
  return AzulGestaoBiblio.transferirTudo(data);
}

function registarTresorerie(data) {
  return AzulGestaoBiblio.registarTresorerie(data);
}

function getTresorerie(data) {
  return AzulGestaoBiblio.getTresorerie(data);
}

function getComptabiliteData(data) {
  return AzulGestaoBiblio.getComptabiliteData(data);
}

function registarDepense(data) {
  return AzulGestaoBiblio.registarDepense(data);
}

function getDepenseDashboard(data) {
  return AzulGestaoBiblio.getDepenseDashboard(data);
}

function getHistoriqueDepenses(data) {
  return AzulGestaoBiblio.getHistoriqueDepenses(data);
}

function getRevendeurNames(data) {
  return AzulGestaoBiblio.getRevendeurNames(data);
}

function getRevendeurDetail(data) {
  return AzulGestaoBiblio.getRevendeurDetail(data);
}

function registarConsignacao(data) {
  return AzulGestaoBiblio.registarConsignacao(data);
}

function getConsignationsOpen(data) {
  return AzulGestaoBiblio.getConsignationsOpen(data);
}

function getConsignationsByRevendeur(data) {
  return AzulGestaoBiblio.getConsignationsByRevendeur(data);
}

function confirmerPaiementConsignation(data) {
  return AzulGestaoBiblio.confirmerPaiementConsignation(data);
}

function confirmerPaiementConsignations(data) {
  return AzulGestaoBiblio.confirmerPaiementConsignations(data);
}

function retornarConsignacao(data) {
  return AzulGestaoBiblio.retornarConsignacao(data);
}

function retornarConsignacoes(data) {
  return AzulGestaoBiblio.retornarConsignacoes(data);
}

function getHistoriqueConsignations(data) {
  return AzulGestaoBiblio.getHistoriqueConsignations(data);
}
function getFournDebt(fournisseur) {
  return AzulGestaoBiblio.getFournDebt({
    fournName: fournisseur
  });
}
/* =========================
   MENU : UI CÔTÉ CLIENT
========================= */

function inicializarFeuilles() {
  const res = AzulGestaoBiblio.inicializarFeuilles({});
  SpreadsheetApp.getUi().alert(res.message || "Feuilles initialisées.");
}

function reinicializarInventaire() {
  const ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();

  var response = ui.alert(
    'Alterar modo de stock',
    'Clique SIM para: apenas loja (as compras entram diretamente no stock da loja)\n\n' +
    'Clique NAO para: loja + armazem (as compras entram primeiro no armazem)',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response !== ui.Button.YES && response !== ui.Button.NO) {
    ui.alert('Alteracao cancelada. O modo de stock nao foi alterado.');
    return;
  }

  var mode = response === ui.Button.YES ? 'boutique' : 'armazem';

  props.setProperty('stockMode', mode);

  const res = AzulGestaoBiblio.reinicializarInventaire({ mode: mode });

  ui.alert(res.message || 'Inventario atualizado.');
}


function protegerFeuilles() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Proteger folhas",
    "Deseja ativar a proteção das folhas agora?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const res = AzulGestaoBiblio.protegerFeuilles({});
  ui.alert(res.message || "Protection activée.");
}

function activarModoEdicao() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Modo de edição",
    "Permitir edição direta por 1 minuto?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const res = AzulGestaoBiblio.activarModoEdicao({});
  ui.alert(res.message || "Modo de edição ativo.");
}

function reinicializarTudo() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "ATENÇÃO",
    "Isto vai apagar TODOS os dados. Esta ação não pode ser anulada.\n\nTem certeza?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var props = PropertiesService.getDocumentProperties();
  var stockModeSalvo = props.getProperty('stockMode') || 'boutique';

  if (typeof AzulGestaoBiblio.setSpreadsheetId === 'function') {
    AzulGestaoBiblio.setSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet().getId());
  }

  const res = AzulGestaoBiblio.reinicializarTudo({
    stockMode: stockModeSalvo
  });

  ui.alert(res.message || "Dados apagados.");
}


/* =========================
   TESTS
========================= */

function testAzulLib() {
  Logger.log(AzulGestaoBiblio.ping());
}

function testProductsLib() {
  Logger.log(JSON.stringify(AzulGestaoBiblio.getProducts({})));
}

function testBackend() {
  Logger.log(callBackend("verification", {}));
}