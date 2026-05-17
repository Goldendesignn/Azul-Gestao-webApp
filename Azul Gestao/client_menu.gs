function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Azul Gestao")
    .addItem("Abrir POS", "abrirPOS")
    .addToUi();
}

function abrirPOS() {
  try {
    const res = verification();

    if (res && res.ok === false) {
      showLicensePopup();
      return;
    }

    showMenupos();
    abrirInterfacePOS();

  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "Erreur connexion backend : " + e.message +
      "\n\nSe esta for a primeira vez que você usa o serviço, o Google solicitará permissões."
    );
  }
}

function abrirInterfacePOS() {
  const html = getHtmlFromBackend_("main");

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html)
      .setWidth(1400)
      .setHeight(900),
    "Azul Gestao"
  );
}

function showLicensePopup() {
  const html = getHtmlFromBackend_("login");

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html)
      .setWidth(500)
      .setHeight(450),
    "Activation licence"
  );
}

function showMenupos() {
  SpreadsheetApp.getUi()
    .createMenu("Azul Gestao")
    .addItem("Abrir POS", "abrirPOS")
    .addSeparator()
    .addItem("Inicializar folhas", "inicializarFeuilles")
    .addItem("Alterar modo de stock", "reinicializarInventaire")
    .addSeparator()
    .addItem("Proteger folhas", "protegerFeuilles")
    .addItem("Modo de edição (1 min)", "activarModoEdicao")
    .addSeparator()
    .addItem("APAGAR TODOS OS DADOS", "reinicializarTudo")
    .addToUi();
}

function ouvrirPOSApresActivation() {
  SpreadsheetApp.getUi().alert("✅ Licença ativada com sucesso! Bem-vindo à Azul Gestão.");

  showMenupos();
  abrirInterfacePOS();
}

function getHtmlFromBackend_(page) {
  const result = callBackend("getHtml", {
    page: page
  });

  if (!result || !result.html) {
    throw new Error("HTML introuvable pour la page : " + page);
  }

  return result.html;
}