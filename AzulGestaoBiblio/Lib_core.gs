// ============================================================
//  POS BACKEND - Apps Script
//
//  INSTALLATION :
//  1. Extensions > Apps Script dans Google Sheets
//  2. Colle ce code dans Codigo.gs
//  3. Cree un fichier HTML nomme "POS_Core" et colle le HTML
//  4. Sauvegarde (Ctrl+S) et recharge Google Sheets
//  5. Menu "Azul Gestão" > Initialize Sheets
// ============================================================

// ============================================================
//  MENU - cree le menu "Azul Gestão" dans Google Sheets
//  Pour changer le nom du menu : modifie 'Azul Gestão'
// ============================================================


// ============================================================
//  OUVRIR LE POS
//  setWidth/setHeight = taille de la fenetre en pixels
// ============================================================



// ============================================================
//  CREER TOUTES LES FEUILLES
//  Ne relance PAS si tu as deja des donnees !
// ============================================================

function inicializarFeuilles(data) {
  var SS = getSS();

  aplicarLocalizacaoPortuguesa_(SS);

  var feuilles = [
    { name: 'Compra', headers: ['Data','Fornecedor','Designaçao','Quantidade','Preço unitaro','Montante total','Categoria','Code','Variaçao','Photo','Preço de venda'] },
    { name: 'Vendas', headers: ['Data','Designationdds','Quantidade','Preço unitaro','Cash','Express','Cartao','Credito','Montante total','Preço de compra','Lucro','Origem','Vendedor','Cliente','Recibo No'] },
    { name: 'Transferencias', headers: ['Data','Designaçao','Quantidade','Nota'] },
    { name: 'Clientes', headers: ['Nome','Phone','Email','Primeira compra','Compra total'] },
    { name: 'Fornecedores', headers: ['Nome','Phone','Pais','Nota'] },
    { name: 'Divida Fornecedores', headers: ['Data','Fornecedor','Type','Designaçao','Montant Achat','Paiement','Solde','Estatuto'] },
    { name: 'Divida Clientes', headers: ['Data','Cliente','Designaçao','Total compra','Montante','Pago','Restante','Estatuto'] },
    { name: 'Despesas', headers: ['Data','Type','Descriçao','Rendimento','Rendimento'] },
    { name: 'Tesouraria', headers: ['Data','Type','Descriçao','Rendimento','Rendimento','Equilibrio'] },
    { name: 'Revendedores', headers: ['ID','Data','Revendeur','Designaçao','Quantidade','Preço unitaro','Montante total','Estatuto','data de pagamento','data de regresso','Recibo No','Pagamento'] },
    { name: 'Immobilizaçao', headers: ['Ativo','Data da compra','Valor','Depreciaçao anual','Valor atual'] }
  ];

  feuilles = folhasPortuguesas_(feuilles);

  feuilles.forEach(function(f) {
    var sheet = SS.getSheetByName(f.name);
    if (!sheet) sheet = SS.insertSheet(f.name);

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, f.headers.length).setValues([f.headers]);
    } else {
      normalizarCabecalhosPortugueses_(sheet, f.headers);
    }

    styliserFeuille_(sheet, f.headers.length);
  });

  data = data || {};
  criarInventaire(data.stockMode === 'armazem' ? 'armazem' : 'boutique');


  return {
    ok: true,
    message: 'Todas as folhas foram criadas e formatadas com sucesso!'
  };
}

function garantirPrimeiraInicializacao(data) {
  data = data || {};

  var ss = getSS();
  var systemSheet = ss.getSheetByName('_Azul_System');

  if (!systemSheet) {
    systemSheet = ss.insertSheet('_Azul_System');
    systemSheet.hideSheet();
    systemSheet.getRange(1, 1, 1, 2).setValues([['Chave', 'Valor']]);
  }

  var values = systemSheet.getDataRange().getValues();
  var initializedAt = '';
  var initializingAt = '';

  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '');
    var val = values[i][1];

    if (key === 'initializedAt') initializedAt = val;
    if (key === 'initializingAt') initializingAt = val;
  }

  var folhasObrigatorias = [
    'Compra',
    'Vendas',
    'Transferencias',
    'Clientes',
    'Fornecedores',
    'Divida Fornecedores',
    'Divida Clientes',
    'Despesas',
    'Tesouraria',
    'Revendedores',
    'Inventario'
  ];

  var faltaFolha = folhasObrigatorias.some(function(nome) {
    return !ss.getSheetByName(nome);
  });

  if (initializedAt && !faltaFolha) {
    return {
      ok: true,
      initialized: false,
      message: 'Sistema ja inicializado.'
    };
  }

  var now = new Date();

  if (initializingAt) {
    var started = new Date(initializingAt);
    var ageMs = now.getTime() - started.getTime();

    if (!isNaN(ageMs) && ageMs < 180000) {
      return {
        ok: true,
        initializing: true,
        message: 'Inicializacao em curso. Aguarde alguns segundos e atualize a pagina.'
      };
    }
  }

  systemSheet.appendRow(['initializingAt', now]);

  var result = inicializarFeuilles({
    stockMode: data.stockMode || 'boutique'
  });

  systemSheet.appendRow(['initializedAt', new Date()]);
  systemSheet.appendRow(['version', '1']);
  systemSheet.appendRow(['stockMode', data.stockMode || 'boutique']);

  return {
    ok: true,
    initialized: true,
    message: result.message || 'Sistema inicializado com sucesso.'
  };
}



function aplicarLocalizacaoPortuguesa_(ss) {
  if (!ss) return;
  try { ss.setSpreadsheetLocale('pt_AO'); } catch (e) {}
  try { ss.setSpreadsheetTimeZone('Africa/Luanda'); } catch (e2) {}
}

function folhasPortuguesas_(folhas) {
  var headersPorFolha = {
    'Compra': ['Data','Fornecedor','Designação','Quantidade','Preço unitário','Montante total','Categoria','Código','Variação','Foto','Preço de venda'],
    'Vendas': ['Data','Designação','Quantidade','Preço unitário','Cash','Express','Cartão','Crédito','Montante total','Preço de compra','Lucro','Origem','Vendedor','Cliente','Nº Recibo'],
    'Transferencias': ['Data','Designação','Quantidade','Nota'],
    'Clientes': ['Nome','Telefone','Email','Primeira compra','Compra total'],
    'Fornecedores': ['Nome','Telefone','País','Nota'],
    'Divida Fornecedores': ['Data','Fornecedor','Tipo','Designação','Montante compra','Pagamento','Saldo','Estatuto'],
    'Divida Clientes': ['Data','Cliente','Designação','Total compra','Montante','Pago','Restante','Estatuto'],
    'Despesas': ['Data','Tipo','Descrição','Entrada','Saída'],
    'Tesouraria': ['Data','Tipo','Descrição','Entrada','Saída','Saldo'],
    'Revendedores': ['ID','Data','Revendedor','Designação','Quantidade','Preço unitário','Montante total','Estatuto','Data de pagamento','Data de regresso','Nº Recibo','Pagamento'],
    'Immobilizaçao': ['Ativo','Data da compra','Valor','Depreciação anual','Valor atual']
  };
  return (folhas || []).map(function(f) {
    return { name: f.name, headers: headersPorFolha[f.name] || f.headers };
  });
}

function normalizarCabecalhosPortugueses_(sheet, headers) {
  if (!sheet || !headers || !headers.length) return;
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function styliserFeuille_(sheet, nbColonnes) {
  var maxRows = Math.max(sheet.getMaxRows(), 50);

  var headerRange = sheet.getRange(1, 1, 1, nbColonnes);

  headerRange
    .setFontWeight('bold')
    .setFontSize(13)
    .setFontColor('#ffffff')
    .setBackground('#0b3d91')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);

  var bodyRange = sheet.getRange(2, 1, maxRows - 1, nbColonnes);

  bodyRange
    .setFontSize(11)
    .setFontColor('#222222')
    .setVerticalAlignment('middle')
    .setWrap(false)
    .setBorder(true, true, true, true, true, true, '#d9e2f3', SpreadsheetApp.BorderStyle.SOLID);

  // Alternance manuelle des couleurs
  for (var r = 2; r <= maxRows; r++) {
    var color = r % 2 === 0 ? '#ffffff' : '#f5f8ff';
    sheet.getRange(r, 1, 1, nbColonnes).setBackground(color);
  }

  for (var col = 1; col <= nbColonnes; col++) {
    sheet.setColumnWidth(col, 130);
  }

  var headers = sheet.getRange(1, 1, 1, nbColonnes).getValues()[0];
  var titre = sheet.getRange(1, 1, 1, nbColonnes);
  titre
    .setWrap(true)

  headers.forEach(function(h, index) {
    var col = index + 1;
    var name = String(h).toLowerCase();
    var normalizedName = name.normalize ? name.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : name;

    if (
      name.includes('descri') ||
      name.includes('design') ||
      name.includes('nota') ||
      name.includes('cliente') ||
      name.includes('fornecedor') ||
      name.includes('revendeur')
    ) {
      sheet.setColumnWidth(col, 180);
    }

    if (
      name.includes('data') ||
      name.includes('code') ||
      name.includes('id') ||
      name.includes('status') ||
      name.includes('estatuto')
    ) {
      sheet.setColumnWidth(col, 120);
    }

    if (
      name.includes('mont') ||
      name.includes('preço') ||
      name.includes('lucro') ||
      name.includes('pago') ||
      name.includes('restante') ||
      name.includes('rendimento') ||
      name.includes('equilibrio')
    ) {
      sheet.setColumnWidth(col, 140);
    }

    aplicarFormatoColunaPortuguesa_(sheet, col, maxRows, name, normalizedName);
  });

  nettoyerColonnesVides_(sheet, nbColonnes);
}

function aplicarFormatoColunaPortuguesa_(sheet, col, maxRows, name, normalizedName) {
  name = String(name || '').toLowerCase();
  normalizedName = String(normalizedName || name).toLowerCase();
  var range = sheet.getRange(2, col, maxRows - 1, 1);

  if (normalizedName.includes('data') || normalizedName === 'primeira compra') {
    range.setNumberFormat('dd/mm/yyyy');
    return;
  }

  if (
    normalizedName.includes('quantidade') ||
    normalizedName.includes('stock') ||
    normalizedName.includes('entradas') ||
    normalizedName.includes('saidas')
  ) {
    range.setNumberFormat('#,##0');
    return;
  }

  if (
    normalizedName.includes('mont') ||
    normalizedName.includes('preco') ||
    normalizedName.includes('lucro') ||
    normalizedName.includes('pago') ||
    normalizedName.includes('restante') ||
    normalizedName.includes('rendimento') ||
    normalizedName.includes('equilibrio') ||
    normalizedName.includes('saldo') ||
    normalizedName === 'entrada' ||
    normalizedName === 'saida' ||
    normalizedName.includes('valor') ||
    normalizedName === 'compra total'
  ) {
    range.setNumberFormat('#,##0 "Kz"');
  }
}

function nettoyerColonnesVides_(sheet, nbColonnesUtiles) {
  var maxColumns = sheet.getMaxColumns();

  if (maxColumns > nbColonnesUtiles) {
    sheet.deleteColumns(nbColonnesUtiles + 1, maxColumns - nbColonnesUtiles);
  }
}


// ============================================================
//  CREER L'INVENTAIRE
//
//  DEUX MODES :
//
//  MODE 'boutique' (par defaut) :
//  --------------------------------
//  Les achats vont DIRECTEMENT en Stock Boutique.
//  Pas de depot separe, pas de transferts necessaires.
//
//  Colonnes : A=Designation, B=Supplier, C=Entries,
//             D=Exits, E=Stock Boutique, F=Unit Price,
//             G=Total Amount
//
//  Formule Stock Boutique = Achats - Ventes (SIMPLE)
//
//  MODE 'armazem' :
//  --------------------------------
//  Les achats vont d'abord en Armazem.
//  Tu transferes vers la Boutique manuellement.
//
//  Colonnes : A=Designation, B=Supplier, C=Entries,
//             D=Exits, E=Stock Boutique, F=Stock Armazem,
//             G=Stock Total, H=Unit Price, I=Total Amount
//
//  Formule Stock Boutique = Transferts - Ventes
//  Formule Stock Armazem  = Achats - Transferts
// ============================================================
function formulaPortuguesa_(formula) {
  return String(formula || '').replace(/,/g, ';');
}

function setFormulaInventario_(range, formula) {
  range.setFormula(formulaPortuguesa_(formula));
}

function criarInventaire(stockMode) {
  var SS = getSS();
  getConsignationsSheet_();

  // Par defaut : boutique si aucun parametre fourni
  stockMode = stockMode === 'armazem' ? 'armazem' : 'boutique';


  var sheet = SS.getSheetByName('Inventario');
  if (!sheet) sheet = SS.insertSheet('Inventario');
  sheet.clear();

  //  EN-TETES selon le mode 
  var headers;
  if (stockMode === 'boutique') {
    headers = ['Designaçao','Fornecedor','Entradas','Saidas','Stock de loja','Preço unitaro','Montante total','Categoria','Code','Variaçao','Photo','Preço de compra','Preço de venda','Principal fornecedor'];
  } else {
    headers = ['Designaçao','Fornecedor','Entradas','Saidas','Stock de loja','Stock Armazem','Total Stock','Preço unitaro','Montante total','Categoria','Code','Variaçao','Photo','Preço de compra','Preço de venda','Principal fornecedor'];
  }

  headers = stockMode === 'boutique'
    ? ['Designa\u00e7\u00e3o','Fornecedor','Entradas','Sa\u00eddas','Stock da loja','Pre\u00e7o unit\u00e1rio','Montante total','Categoria','C\u00f3digo','Varia\u00e7\u00e3o','Foto','Pre\u00e7o de compra','Pre\u00e7o de venda','Principal fornecedor']
    : ['Designa\u00e7\u00e3o','Fornecedor','Entradas','Sa\u00eddas','Stock da loja','Stock Armaz\u00e9m','Stock total','Pre\u00e7o unit\u00e1rio','Montante total','Categoria','C\u00f3digo','Varia\u00e7\u00e3o','Foto','Pre\u00e7o de compra','Pre\u00e7o de venda','Principal fornecedor'];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1f3864')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);

  // Bordure sur toute la plage de données (ex: 200 lignes)
  var nbLignes = 200; // Adaptez selon vos besoins
  sheet.getRange(1, 1, nbLignes, headers.length)
    .setBorder(true, true, true, true, true, true, '#1f3864', SpreadsheetApp.BorderStyle.SOLID);

  //  COLONNE A - Liste unique et triee des produits 
  // TRIM = supprime les espaces, UNIQUE = pas de doublons
  setFormulaInventario_(sheet.getRange('A2'),
    '=SORT(FILTER(UNIQUE(TRIM(Compra!C2:C)), UNIQUE(TRIM(Compra!C2:C))<>""))'
  );

  //  COLONNE B - Fournisseur 
  // Formule copiee sur 200 lignes (ARRAYFORMULA ne marche pas
  // bien avec INDEX/MATCH - il retourne toujours le 1er resultat)
  // TRIM+LOWER = ignore les espaces et differences de casse
  for (var row = 2; row <= 200; row++) {
    setFormulaInventario_(sheet.getRange('B' + row),
      '=IFERROR(INDEX(Compra!B:B, MATCH(TRIM(LOWER(A' + row + ')), ARRAYFORMULA(TRIM(LOWER(Compra!C:C))), 0)), "")'
    );
  }

  //  COLONNE C - Entries (total achete par produit) 
  // Formule copiee sur 200 lignes (ARRAYFORMULA ne fonctionne pas
  // bien avec SUMIF ligne par ligne)
  // SUMIF = additionne Achat!D quand Achat!C correspond au produit
  // SUMIF avec TRIM+LOWER pour ignorer espaces et casse
  // On utilise SUMPRODUCT car SUMIF ne supporte pas TRIM/LOWER directement
  // SUMPRODUCT multiplie deux tableaux : 1 si le produit correspond, x la quantite
  for (var rowC = 2; rowC <= 200; rowC++) {
    setFormulaInventario_(sheet.getRange('C' + rowC),
      '=IF(A' + rowC + '<>"", SUMPRODUCT((TRIM(LOWER(Compra!C2:C))=TRIM(LOWER(A' + rowC + ')))*(Compra!D2:D)), "")'
    );
  }

  //  COLONNE D - Exits (total vendu avec statut "interno") 
  // Formule copiee sur 200 lignes
  // SUMIFS = 2 conditions : produit ET statut="interno"
  // SUMPRODUCT avec TRIM+LOWER pour les sorties aussi
  var vendasProductNameFormula = 'REGEXREPLACE(Vendas!B2:B, " \\[[^\\]]*\\]$", "")';
  for (var rowD = 2; rowD <= 200; rowD++) {
    setFormulaInventario_(sheet.getRange('D' + rowD),
      '=IF(A' + rowD + '<>"", SUMPRODUCT((TRIM(LOWER(' + vendasProductNameFormula + '))=TRIM(LOWER(A' + rowD + ')))*(Vendas!L2:L="interno")*(Vendas!C2:C)), "")'
    );
  }

  if (stockMode === 'boutique') {
    // 
    //  MODE BOUTIQUE UNIQUEMENT
    //  Stock Boutique = Total Achete - Total Vendu
    //  Simple et direct, pas de transferts
    // 

    //  COLONNE E - Stock Boutique 
    // = Achats - Ventes
    // Formule copiee sur 200 lignes
    // SUMPRODUCT pour Stock Boutique aussi
    for (var rowE = 2; rowE <= 200; rowE++) {
      setFormulaInventario_(sheet.getRange('E' + rowE),
        '=IF(A' + rowE + '<>"", SUMPRODUCT((TRIM(LOWER(Compra!C2:C))=TRIM(LOWER(A' + rowE + ')))*(Compra!D2:D)) - SUMPRODUCT((TRIM(LOWER(' + vendasProductNameFormula + '))=TRIM(LOWER(A' + rowE + ')))*(Vendas!L2:L="interno")*(Vendas!C2:C)) - SUMPRODUCT((TRIM(LOWER(Revendedores!D2:D))=TRIM(LOWER(A' + rowE + ')))*((Revendedores!H2:H="Em curso")+(Revendedores!H2:H="En cours"))*(Revendedores!E2:E)), "")'
      );
    }

    //  COLONNE F - Unit Price 
    // Meme correction TRIM+LOWER que colonne B
    for (var rowF = 2; rowF <= 200; rowF++) {
      setFormulaInventario_(sheet.getRange('F' + rowF),
        '=IFERROR(INDEX(Compra!E:E, MATCH(TRIM(LOWER(A' + rowF + ')), ARRAYFORMULA(TRIM(LOWER(Compra!C:C))), 0)), "")'
      );
    }

    //  COLONNE G - Total Amount = Stock x Prix 
    for (var rowG = 2; rowG <= 200; rowG++) {
      setFormulaInventario_(sheet.getRange('G' + rowG),
        '=IF(A' + rowG + '<>"", E' + rowG + ' * F' + rowG + ', "")'
      );
    }

  } else {
    // 
    //  MODE BOUTIQUE + ARMAZEM
    //  Stock Boutique = Transferts vers Boutique - Ventes
    //  Stock Armazem  = Achats - Transferts vers Boutique
    // 

    //  COLONNE E - Stock Boutique 
    // = Transferts - Ventes, formule copiee sur 200 lignes
    for (var rowEa = 2; rowEa <= 200; rowEa++) {
      setFormulaInventario_(sheet.getRange('E' + rowEa),
        '=IF(A' + rowEa + '<>"", SUMPRODUCT((TRIM(LOWER(Transferencias!B2:B))=TRIM(LOWER(A' + rowEa + ')))*(Transferencias!C2:C)) - SUMPRODUCT((TRIM(LOWER(' + vendasProductNameFormula + '))=TRIM(LOWER(A' + rowEa + ')))*(Vendas!L2:L="interno")*(Vendas!C2:C)) - SUMPRODUCT((TRIM(LOWER(Revendedores!D2:D))=TRIM(LOWER(A' + rowEa + ')))*((Revendedores!H2:H="Em curso")+(Revendedores!H2:H="En cours"))*(Revendedores!E2:E)), "")'
      );
    }

    //  COLONNE F - Stock Armazem 
    // = Achats - Transferts, formule copiee sur 200 lignes
    for (var rowFa = 2; rowFa <= 200; rowFa++) {
      setFormulaInventario_(sheet.getRange('F' + rowFa),
        '=IF(A' + rowFa + '<>"", SUMPRODUCT((TRIM(LOWER(Compra!C2:C))=TRIM(LOWER(A' + rowFa + ')))*(Compra!D2:D)) - SUMPRODUCT((TRIM(LOWER(Transferencias!B2:B))=TRIM(LOWER(A' + rowFa + ')))*(Transferencias!C2:C)), "")'
      );
    }

    //  COLONNE G - Total Stock = Boutique + Armazem 
    for (var rowGa = 2; rowGa <= 200; rowGa++) {
      setFormulaInventario_(sheet.getRange('G' + rowGa),
        '=IF(A' + rowGa + '<>"", E' + rowGa + ' + F' + rowGa + ', "")'
      );
    }

    //  COLONNE H - Unit Price 
    for (var rowH = 2; rowH <= 200; rowH++) {
      setFormulaInventario_(sheet.getRange('H' + rowH),
        '=IFERROR(INDEX(Compra!E:E, MATCH(TRIM(LOWER(A' + rowH + ')), ARRAYFORMULA(TRIM(LOWER(Compra!C:C))), 0)), "")'
      );
    }

    //  COLONNE I - Total Amount 
    for (var rowI = 2; rowI <= 200; rowI++) {
      setFormulaInventario_(sheet.getRange('I' + rowI),
        '=IF(A' + rowI + '<>"", G' + rowI + ' * H' + rowI + ', "")'
      );
    }
  }
  headers.forEach(function(header, index) {
    var name = String(header || '').toLowerCase();
    var normalizedName = name.normalize ? name.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : name;
    aplicarFormatoColunaPortuguesa_(sheet, index + 1, nbLignes, name, normalizedName);
  });
  applyInventoryMetadataFormulas_(sheet, headers);
}
function normalizarCabecalho_(value) {
  var text = String(value || '').trim().toLowerCase();
  if (text.normalize) text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return text;
}

function indiceCabecalho_(headers, aliases) {
  var normalized = (headers || []).map(normalizarCabecalho_);
  for (var i = 0; i < aliases.length; i++) {
    var wanted = normalizarCabecalho_(aliases[i]);
    var index = normalized.indexOf(wanted);
    if (index >= 0) return index;
  }
  return -1;
}

function getAchatExpectedHeaders_() {
  return ['Data','Fornecedor','Designa\u00e7\u00e3o','Quantidade','Pre\u00e7o unit\u00e1rio','Montante total','Categoria','C\u00f3digo','Varia\u00e7\u00e3o','Foto','Pre\u00e7o de venda'];
  return ['Data','Fornecedor','Designaçao','Quantidade','Preço unitaro','Montante total','Categoria','Code','Variaçao','Photo','Preço de venda'];
}

function getAchatColumnMap_(headers) {
  return {
    date: indiceCabecalho_(headers, ['Data']),
    supplier: indiceCabecalho_(headers, ['Fornecedor']),
    designation: indiceCabecalho_(headers, ['Designa\u00e7\u00e3o','Designa\u00e7ao','DesignaÃ§ao','Designationdds']),
    quantity: indiceCabecalho_(headers, ['Quantidade']),
    unitPrice: indiceCabecalho_(headers, ['Pre\u00e7o unit\u00e1rio','Pre\u00e7o unitario','Pre\u00e7o unitaro','PreÃ§o unitaro']),
    totalAmount: indiceCabecalho_(headers, ['Montante total']),
    category: indiceCabecalho_(headers, ['Categoria']),
    code: indiceCabecalho_(headers, ['C\u00f3digo','Code']),
    variation: indiceCabecalho_(headers, ['Varia\u00e7\u00e3o','Varia\u00e7ao','VariaÃ§ao']),
    photo: indiceCabecalho_(headers, ['Foto','Photo']),
    targetMargin: indiceCabecalho_(headers, ['Pre\u00e7o de venda','PreÃ§o de venda','Margem alvo'])
  };
  var normalized = (headers || []).map(function(header) { return String(header || '').trim(); });
  return {
    date: normalized.indexOf('Data'),
    supplier: normalized.indexOf('Fornecedor'),
    designation: normalized.indexOf('Designaçao'),
    quantity: normalized.indexOf('Quantidade'),
    unitPrice: normalized.indexOf('Preço unitaro'),
    totalAmount: normalized.indexOf('Montante total'),
    category: normalized.indexOf('Categoria'),
    code: normalized.indexOf('Code'),
    variation: normalized.indexOf('Variaçao'),
    photo: normalized.indexOf('Photo'),
    targetMargin: normalized.indexOf('Preço de venda') >= 0 ? normalized.indexOf('Preço de venda') : normalized.indexOf('Margem alvo')
  };
}

function ensureAchatProductColumns_(sheet) {
  if (!sheet) return getAchatColumnMap_([]);

  var baseHeaders = ['Data','Fornecedor','Designaçao','Quantidade','Preço unitaro','Montante total'];
  var metaHeaders = ['Categoria','Code','Variaçao','Photo','Preço de venda'];
  baseHeaders = ['Data','Fornecedor','Designa\u00e7\u00e3o','Quantidade','Pre\u00e7o unit\u00e1rio','Montante total'];
  metaHeaders = ['Categoria','C\u00f3digo','Varia\u00e7\u00e3o','Foto','Pre\u00e7o de venda'];
  var lastColumn = Math.max(sheet.getLastColumn(), baseHeaders.length);
  if (sheet.getMaxColumns() < lastColumn) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastColumn - sheet.getMaxColumns());
  }

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var oldSalePriceIndex = headers.indexOf('Margem alvo');
  if (oldSalePriceIndex !== -1 && headers.indexOf('Preço de venda') === -1) {
    headers[oldSalePriceIndex] = 'Preço de venda';
  }

  for (var i = 0; i < baseHeaders.length; i++) {
    headers[i] = baseHeaders[i];
  }

  var occurrences = {};
  for (var h = 0; h < headers.length; h++) {
    var name = headers[h];
    if (!name) continue;
    occurrences[name] = occurrences[name] || [];
    occurrences[name].push(h);
  }

  metaHeaders.forEach(function(name) {
    var positions = occurrences[name] || [];
    if (positions.length > 1) {
      for (var d = 1; d < positions.length; d++) {
        headers[positions[d]] = name + ' Legacy ' + d;
      }
    }
  });

  metaHeaders.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      headers.push(name);
    }
  });

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');

  return getAchatColumnMap_(headers);
}
function firstDefinedValue_(values) {
  for (var i = 0; i < values.length; i++) {
    if (values[i] !== undefined && values[i] !== null && String(values[i]) !== '') return values[i];
  }
  return '';
}

function buildAchatRow_(map, supplier, line) {
  line = line || {};
  var width = 0;
  Object.keys(map).forEach(function(key) {
    width = Math.max(width, (map[key] || 0) + 1);
  });
  var row = new Array(width);
  for (var i = 0; i < row.length; i++) row[i] = '';

  var designation = firstDefinedValue_([line.prod, line.designation, line.name]);
  var quantity = firstDefinedValue_([line.qty, line.quantity]);
  var unitPrice = firstDefinedValue_([line.price, line.unitPrice, line.purchasePrice]);
  var totalAmount = line.total != null && line.total !== '' ? line.total : ((parseFloat(quantity) || 0) * (parseFloat(unitPrice) || 0));
  var category = firstDefinedValue_([line.category, line.categorie]);
  var code = firstDefinedValue_([line.code, line.productCode]);
  var variation = firstDefinedValue_([line.variation, line.variations]);
  var photo = firstDefinedValue_([line.photo, line.image, line.imageUrl]);
  var targetMargin = firstDefinedValue_([line.targetMargin, line.margin, line.marge]);

  row[map.date] = firstDefinedValue_([line.date]);
  row[map.supplier] = firstDefinedValue_([supplier, line.forn, line.supplier, line.mainSupplier]);
  row[map.designation] = designation;
  row[map.quantity] = quantity || 0;
  row[map.unitPrice] = unitPrice || 0;
  row[map.totalAmount] = totalAmount;
  if (map.category >= 0) row[map.category] = category;
  if (map.code >= 0) row[map.code] = code;
  if (map.variation >= 0) row[map.variation] = normalizeVariationValue_(variation);
  if (map.photo >= 0) row[map.photo] = photo;
  if (map.targetMargin >= 0) row[map.targetMargin] = targetMargin;
  return row;
}

function parseVariationList_(value) {
  return String(value || '')
    .split(/\s*[|,;]+\s*/)
    .map(function(entry) { return String(entry || '').trim(); })
    .filter(function(entry, index, list) { return entry && list.indexOf(entry) === index; });
}

function normalizeVariationValue_(value) {
  if (Array.isArray(value)) return parseVariationList_(value.join('|')).join(' | ');
  return parseVariationList_(value).join(' | ');
}

function columnToLetter_(column) {
  var temp = '';
  while (column > 0) {
    var remainder = (column - 1) % 26;
    temp = String.fromCharCode(65 + remainder) + temp;
    column = Math.floor((column - 1) / 26);
  }
  return temp;
}
function getProductMetaHeaders_() {
  return ['Categoria','C\u00f3digo','Varia\u00e7\u00e3o','Foto','Pre\u00e7o de compra','Pre\u00e7o de venda','Principal fornecedor'];
  return ['Categoria','Code','Variaçao','Photo','Preço de compra','Preço de venda','Principal fornecedor'];
}

function isArmazemInventory_(headers) {
  headers = headers || [];
  if (indiceCabecalho_(headers, ['Stock Armaz\u00e9m','Stock Armazem','Stock total','Total Stock']) !== -1) return true;
  return headers.indexOf('Stock Armazem') !== -1 || headers.indexOf('Total Stock') !== -1;
}

function ensureProductMetadataColumns_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var oldSalePriceIndex = headers.indexOf('Margem alvo');
  if (oldSalePriceIndex !== -1 && headers.indexOf('Preço de venda') === -1) {
    sheet.getRange(1, oldSalePriceIndex + 1).setValue('Preço de venda');
    headers[oldSalePriceIndex] = 'Preço de venda';
  }
  var renamedMetaHeaders = {
    'Code': 'C\u00f3digo',
    'Photo': 'Foto',
    'VariaÃ§ao': 'Varia\u00e7\u00e3o',
    'PreÃ§o de compra': 'Pre\u00e7o de compra',
    'PreÃ§o de venda': 'Pre\u00e7o de venda'
  };
  headers.forEach(function(header, index) {
    if (!renamedMetaHeaders[header]) return;
    headers[index] = renamedMetaHeaders[header];
    sheet.getRange(1, index + 1).setValue(headers[index]);
  });
  var metaHeaders = getProductMetaHeaders_();
  metaHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      var col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col)
        .setValue(header)
        .setFontWeight('bold')
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff');
      headers.push(header);
    }
  });
  return headers;
}


function getInventoryMetadataFormulas_(row, map) {
  function lookupFormula(compraColumn) {
    return formulaPortuguesa_('=IFERROR(INDEX(Compra!' + compraColumn + ':' + compraColumn + ', MATCH(TRIM(LOWER(A' + row + ')), ARRAYFORMULA(TRIM(LOWER(Compra!C:C))), 0)), "")');
  }

  var formulas = {};
  formulas[map.category] = lookupFormula('G');
  formulas[map.code] = lookupFormula('H');
  formulas[map.variation] = lookupFormula('I');
  formulas[map.photo] = lookupFormula('J');
  formulas[map.purchasePrice] = lookupFormula('E');
  formulas[map.targetMargin] = lookupFormula('K');
  formulas[map.mainSupplier] = lookupFormula('B');
  return formulas;
}

function applyInventoryMetadataFormulas_(sheet, headers) {
  if (!sheet) return false;
  headers = headers || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = getInventoryColumnMap_(headers);
  var metaColumns = [map.category, map.code, map.variation, map.photo, map.purchasePrice, map.targetMargin, map.mainSupplier]
    .filter(function(index) { return index >= 0; });
  var maxRow = Math.min(Math.max(sheet.getMaxRows(), sheet.getLastRow()), 200);
  if (maxRow < 2 || !metaColumns.length) return false;

  var currentFormulas = sheet.getRange(1, 1, maxRow, sheet.getLastColumn()).getFormulas();
  var changed = false;

  for (var row = 2; row <= maxRow; row++) {
    var expected = getInventoryMetadataFormulas_(row, map);
    metaColumns.forEach(function(index) {
      if (currentFormulas[row - 1][index] === expected[index]) return;
      setFormulaInventario_(sheet.getRange(row, index + 1), expected[index]);
      changed = true;
    });
  }

  return changed;
}
function getInventoryColumnMap_(headers) {
  var isArmazem = isArmazemInventory_(headers);
  return {
    isArmazem: isArmazem,
    name: 0,
    supplier: 1,
    entries: 2,
    exits: 3,
    stockBoutique: 4,
    stockStockage: indiceCabecalho_(headers, ['Stock Armazem']),
    stockArmazem: isArmazem ? 5 : -1,
    stock: isArmazem ? 6 : 4,
    price: isArmazem ? 7 : 5,
    total: isArmazem ? 8 : 6,
    category: indiceCabecalho_(headers, ['Categoria']),
    code: indiceCabecalho_(headers, ['C\u00f3digo','Code']),
    variation: indiceCabecalho_(headers, ['Varia\u00e7\u00e3o','Varia\u00e7ao','VariaÃ§ao']),
    photo: indiceCabecalho_(headers, ['Foto','Photo']),
    purchasePrice: indiceCabecalho_(headers, ['Pre\u00e7o de compra','PreÃ§o de compra']),
    targetMargin: indiceCabecalho_(headers, ['Pre\u00e7o de venda','PreÃ§o de venda','Margem alvo']),
    mainSupplier: indiceCabecalho_(headers, ['Principal fornecedor'])
  };
}


// ============================================================
//  CHANGER LE MODE DE STOCK
//  Recree seulement les formules, sans toucher aux donnees
//
//  UTILISATION :
//  Menu Azul Gestão > Change Stock Mode
//  OU : Apps Script > selectionne reinicializarInventaire > Run
// ============================================================
function reinicializarInventaire(data) {
  data = data || {};

  var mode = data.mode === 'armazem' ? 'armazem' : 'boutique';

  criarInventaire(mode);

  return {
    ok: true,
    mode: mode,
    message: 'Inventario atualizado! Modo: ' +
      (mode === 'boutique' ? 'Apenas loja' : 'Loja + armazem')
  };
}



// ============================================================
//  RECUPERER LES PRODUITS POUR LE POS
//
//  Lit l'Inventaire et retourne les produits pour la page
//  "Nova Venda" du POS.
//
//  Colonnes lues (indices 0-based) :
//  Mode boutique : E(4)=Stock Boutique, F(5)=Unit Price, E(4)=Total
//  Mode armazem  : E(4)=Stock Boutique, H(7)=Unit Price, G(6)=Total
//
//  Pour changer le seuil "stock faible" : modifie le 3
//  dans getDashboardData plus bas
// ============================================================

function corrigirFormulasStockVendas_() {
  var inv = getSS().getSheetByName('Inventario');
  if (!inv) return;

  var range = inv.getDataRange();
  var formulas = range.getFormulas();
  var changed = false;
  var vendasBaseFormula = 'TRIM(LOWER(REGEXREPLACE(Vendas!B2:B, " \\[[^\\]]*\\]$", "")))';
  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      var formula = formulas[r][c];
      if (formula && formula.indexOf('Vendas!H2:H="interno"') !== -1) {
        formula = formula.replace(/Vendas!H2:H="interno"/g, 'Vendas!L2:L="interno"');
        changed = true;
      }
      if (formula && formula.indexOf('TRIM(LOWER(Vendas!B2:B))') !== -1) {
        formula = formula.replace(/TRIM\(LOWER\(Vendas!B2:B\)\)/g, vendasBaseFormula);
        changed = true;
      }
      if (formula !== formulas[r][c]) {
        setFormulaInventario_(range.getCell(r + 1, c + 1), formula);
        changed = true;
      }
    }
  }
  if (changed) SpreadsheetApp.flush();
}
function getProducts() {
  corrigirFormulasStockVendas_();
  var SS = getSS();
  var inv = SS.getSheetByName('Inventario');
  if (!inv) return [];

  var headers = ensureProductMetadataColumns_(inv);
  if (applyInventoryMetadataFormulas_(inv, headers)) SpreadsheetApp.flush();
  var data = inv.getDataRange().getValues();
  if (data.length < 2) return [];

  var map = getInventoryColumnMap_(headers);
  var products = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][map.name];
    if (!name || name === '') continue;

    var purchasePrice = map.purchasePrice >= 0 ? (parseFloat(data[i][map.purchasePrice]) || 0) : 0;
    var unitPrice = parseFloat(data[i][map.price]) || 0;
    var salePrice = map.targetMargin >= 0 ? (parseFloat(data[i][map.targetMargin]) || 0) : 0;
  
    products.push({
      name: name,
      price: salePrice || unitPrice,
      entries: data[i][map.entries] || 0,
      exits: data[i][map.exits] || 0,
      salePrice: salePrice || unitPrice,
      stock: parseFloat(data[i][map.stock]) || 0,
      stockBoutique: parseFloat(data[i][map.stockBoutique]) || 0,
      stockage: parseFloat(data[i][map.stockStockage]) || 0,
      photo: map.photo >= 0 ? String(data[i][map.photo] || '') : '',
      category: map.category >= 0 ? String(data[i][map.category] || '') : '',
      code: map.code >= 0 ? String(data[i][map.code] || '') : '',
      variation: map.variation >= 0 ? String(data[i][map.variation] || '') : '',
      variations: map.variation >= 0 ? parseVariationList_(data[i][map.variation]) : [],
      purchasePrice: purchasePrice || unitPrice,
      targetMargin: salePrice,
      mainSupplier: map.mainSupplier >= 0 ? String(data[i][map.mainSupplier] || data[i][map.supplier] || '') : String(data[i][map.supplier] || ''),
      supplier: String(data[i][map.supplier] || '')
    });
  }
  return products;
}

function saveProductProfile(data) {
  data = data || {};
  var name = String(data.name || '').trim();
  if (!name) throw new Error('Produit introuvable.');

  var ss = getSS();
  var achat = ss.getSheetByName('Compra');
  var inv = ss.getSheetByName('Inventario');
  if (!achat) throw new Error('Achat sheet not found');

  var achatMap = ensureAchatProductColumns_(achat);
  var lastRow = achat.getLastRow();
  var targetRow = -1;

  if (lastRow > 1) {
    var values = achat.getRange(2, 1, lastRow - 1, achat.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][achatMap.designation] || '').trim().toLowerCase() === name.toLowerCase()) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) throw new Error("Produit non trouve dans la feuille Compra.");

  if (achatMap.category >= 0) achat.getRange(targetRow, achatMap.category + 1).setValue(String(data.category || '').trim());
  if (achatMap.code >= 0) achat.getRange(targetRow, achatMap.code + 1).setValue(String(data.code || '').trim());
  if (achatMap.photo >= 0) achat.getRange(targetRow, achatMap.photo + 1).setValue(String(data.photo || '').trim());
  if (achatMap.targetMargin >= 0) {
    achat.getRange(targetRow, achatMap.targetMargin + 1).setValue(data.targetMargin === '' || data.targetMargin == null ? '' : (parseFloat(data.targetMargin) || 0));
  }
  if (achatMap.unitPrice >= 0 && data.purchasePrice !== '' && data.purchasePrice != null) {
    achat.getRange(targetRow, achatMap.unitPrice + 1).setValue(parseFloat(data.purchasePrice) || 0);
  }
  if (achatMap.supplier >= 0 && data.mainSupplier) {
    achat.getRange(targetRow, achatMap.supplier + 1).setValue(String(data.mainSupplier || '').trim());
  }

  if (inv) {
    SpreadsheetApp.flush();
    var headers = ensureProductMetadataColumns_(inv);
    applyInventoryMetadataFormulas_(inv, headers);
  }

  invalidateProductsCache();
  return true;
}
function syncProductMetadataFromPurchase_(lines, supplier) {
  var inv = getSS().getSheetByName('Inventario');
  if (!inv) return;
  SpreadsheetApp.flush();
  var headers = ensureProductMetadataColumns_(inv);
  applyInventoryMetadataFormulas_(inv, headers);
}
function parseSheetDate(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  var text = value.toString().trim();
  var m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  }

  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  }

  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  return null;
}

function normalizeDashboardText_(value) {
  var text = String(value == null ? '' : value).toLowerCase().trim();
  try {
    text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  } catch (e) {}
  return text.replace(/\s+/g, ' ');
}

function dashboardTextMatches_(value, filter) {
  var search = normalizeDashboardText_(filter);
  if (!search) return true;
  var text = normalizeDashboardText_(value);
  return search.split(' ').every(function(part) {
    return !part || text.indexOf(part) >= 0;
  });
}

function getDashboardDateRange_(filters) {
  filters = filters || {};
  var now = new Date();
  var from = filters.from ? parseSheetDate(filters.from) : new Date(now.getFullYear(), now.getMonth(), 1);
  var to = filters.to ? parseSheetDate(filters.to) : now;
  if (!from) from = new Date(now.getFullYear(), now.getMonth(), 1);
  if (!to) to = now;
  from.setHours(0, 0, 0, 0);
  to.setHours(23, 59, 59, 999);
  if (from > to) {
    var tmp = from;
    from = new Date(to.getFullYear(), to.getMonth(), to.getDate());
    to = new Date(tmp.getFullYear(), tmp.getMonth(), tmp.getDate(), 23, 59, 59, 999);
  }
  return { from: from, to: to };
}

function getTodaySheetDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Luanda', 'dd/MM/yyyy');
}

function getTresorerieSheet_() {
  var sheet = getSS().getSheetByName('Tesouraria');
  if (!sheet) throw new Error('Tresorerie sheet not found');
  return sheet;
}

function getLastTresorerieBalance_(sheet) {
  sheet = sheet || getTresorerieSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  return parseFloat(sheet.getRange(lastRow, 6).getValue()) || 0;
}

function appendTresorerieMovement_(date, type, description, income, expense) {
  var sheet = getTresorerieSheet_();
  var inVal = Math.max(parseFloat(income) || 0, 0);
  var outVal = Math.max(parseFloat(expense) || 0, 0);
  var balance = getLastTresorerieBalance_(sheet) + inVal - outVal;
  var lastRow = sheet.getLastRow() + 1;

  sheet.getRange(lastRow, 1, 1, 6).setValues([[
    date || getTodaySheetDate_(),
    type || 'Mouvement',
    description || '',
    inVal,
    outVal,
    balance
  ]]);

  return balance;
}

function buildItemsSummary_(items) {
  if (!items || items.length === 0) return '';
  var names = items.map(function(item) { return item.name || item.prod || ''; }).filter(String);
  if (names.length === 0) return '';
  if (names.length === 1) return names[0];
  if (names.length === 2) return names[0] + ' + ' + names[1];
  return names[0] + ' +' + (names.length - 1);
}


// ============================================================
//  DONNEES DU DASHBOARD
//
//  Calcule les statistiques selon les filtres :
//  filters.from   = date debut
//  filters.to     = date fin
//  filters.prod   = filtre par produit
//  filters.forn   = filtre par fournisseur
//
//  Pour changer le seuil alerte stock faible : modifie le 3
//  Pour afficher plus de depenses : modifie le 5
// ============================================================
function getDashboardData(filters) {
  var SS = getSS();
  var ventes = SS.getSheetByName('Vendas');
  var inv = SS.getSheetByName('Inventario');
  var depensesSheet = SS.getSheetByName('Despesas');
  var tresorerieSheet = SS.getSheetByName('Tesouraria');

  // Après var tresorerieSheet = SS.getSheetByName('Tesouraria'); moi-meme
  var prodFournisseurMap = {};
  if (inv) {
    var invData = inv.getDataRange().getValues();
    for (var j = 1; j < invData.length; j++) {
      var nomProd = normalizeDashboardText_(invData[j][0]);
      var supplier = normalizeDashboardText_(invData[j][1]);
      if (nomProd) prodFournisseurMap[nomProd] = supplier;
    }
  }

  if (!ventes) return {};

  filters = filters || {};
  var range = getDashboardDateRange_(filters);
  var from = range.from;
  var to = range.to;
  var prodFilter = normalizeDashboardText_(filters.prod);
  var fornFilter = normalizeDashboardText_(filters.forn);

  var data = ventes.getDataRange().getValues();
  var headers = data.length ? data[0].map(function(h) { return normalizeDashboardText_(h); }) : [];
  function col(names, fallback) {
    for (var n = 0; n < names.length; n++) {
      var idx = headers.indexOf(normalizeDashboardText_(names[n]));
      if (idx >= 0) return idx;
    }
    return fallback;
  }
  function numCell(value) { return parseFloat(value) || 0; }

  var ixDate = col(['Data'], 0);
  var ixProd = col(['Designationdds', 'Designaçao', 'Produto', 'Product'], 1);
  var ixQty = col(['Quantidade', 'Qtd', 'Qty'], 2);
  var ixTotal = col(['Montante total', 'Total'], 8);
  var ixProfit = col(['Lucro'], 10);
  var ixPayment = col(['Pagamento'], -1);
  var ixCash = col(['Cash'], -1);
  var ixExpress = col(['Express'], -1);
  var ixCartao = col(['Cartao', 'Cartão'], -1);
  var ixCredito = col(['Credito', 'Crédito', 'Credit'], -1);
  var hasPaymentColumns = ixCash >= 0 || ixExpress >= 0 || ixCartao >= 0 || ixCredito >= 0;

  var vendasTotal = 0, vendasCount = 0, lucroTotal = 0;
  var pagamentos = { Cash: 0, Express: 0, Cartao: 0, Credito: 0 };
  var prodMap = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[ixDate]) continue;

    var rowDate = parseSheetDate(row[ixDate]);
    if (!rowDate) continue;
    if (rowDate < from || rowDate > to) continue;

    var prod = (row[ixProd] || '').toString();
    if (prodFilter && !dashboardTextMatches_(prod, prodFilter)) continue;

    // Remplace cette ligne vide (fornFilter n'était pas utilisé dans ventes)
    // par ceci, juste après le filtre prodFilter :
    // NOUVEAU - filtre fournisseur
    if (fornFilter) {
      var supplierDuProd = prodFournisseurMap[normalizeDashboardText_(getSaleInventoryName_(prod))] || prodFournisseurMap[normalizeDashboardText_(prod)] || '';
      if (!dashboardTextMatches_(supplierDuProd, fornFilter)) continue;
    }

    var montant  = numCell(row[ixTotal]);
    var benefice = numCell(row[ixProfit]);
    var qty      = numCell(row[ixQty]);

    vendasTotal += montant;
    vendasCount++;
    lucroTotal += benefice;

    if (hasPaymentColumns) {
      pagamentos.Cash += ixCash >= 0 ? numCell(row[ixCash]) : 0;
      pagamentos.Express += ixExpress >= 0 ? numCell(row[ixExpress]) : 0;
      pagamentos.Cartao += ixCartao >= 0 ? numCell(row[ixCartao]) : 0;
      pagamentos.Credito += ixCredito >= 0 ? numCell(row[ixCredito]) : 0;
    } else {
      var pay = (row[ixPayment] || 'Cash').toString();
      var repartition = repartirMontantParPaiement(pay, montant);
      pagamentos.Cash += repartition.Cash;
      pagamentos.Express += repartition.Express;
      pagamentos.Cartao += repartition.Cartao;
      pagamentos.Credito += repartition.Credito;
    }

    if (prod) {
      if (!prodMap[prod]) prodMap[prod] = { qty: 0, total: 0 };
      prodMap[prod].qty   += qty;
      prodMap[prod].total += montant;
    }
  }

  // Top 5 produits par total de ventes
  var topProdutos = Object.keys(prodMap).map(function(k) {
    return { name: k, qty: prodMap[k].qty, total: prodMap[k].total };
  }).sort(function(a, b) { return b.total - a.total; }).slice(0, 5);

  // Alertes stock faible - seuil = 3 unites
  // Modifie le 3 pour changer ce seuil
  var stockAlertas = [];
  if (inv) {
    var invData = inv.getDataRange().getValues();
    for (var j = 1; j < invData.length; j++) {
      var stockB   = parseFloat(invData[j][4]) || 0; // E = Stock Boutique (meme colonne dans les 2 modes)
      var nomeProd = invData[j][0];
      if (nomeProd && stockB <= 3 && stockB >= 0) {
        stockAlertas.push({
          name:  nomeProd,
          stock: stockB,
          level: stockB <= 1 ? 'critical' : 'warning'
        });
      }
    }
  }

  // Depenses
  var totalDepenses = 0, depensesCount = 0, depenses = [];
  if (depensesSheet && depensesSheet.getLastRow() > 1) {
    var depData = depensesSheet.getDataRange().getValues();
    for (var d = 1; d < depData.length; d++) {
      var depRow = depData[d];
      if (!depRow[0]) continue;
      var depDate = parseSheetDate(depRow[0]);
      if (!depDate) continue;
      if (depDate < from || depDate > to) continue;
      var depDesc = (depRow[2] || '').toString();
      if (fornFilter && !dashboardTextMatches_(depDesc, fornFilter)) continue;
      var depVal = parseFloat(depRow[4]) || 0;
      totalDepenses += depVal;
      depensesCount++;
      if (depenses.length < 5) {
        depenses.push({
          date:  depDate.toLocaleDateString('pt-PT'),
          desc:  depDesc || depRow[1] || '',
          valor: depVal
        });
      }
    }
  }

  var soldeTresorerie = 0;
  if (tresorerieSheet && tresorerieSheet.getLastRow() > 1) {
    soldeTresorerie = parseFloat(tresorerieSheet.getRange(tresorerieSheet.getLastRow(), 6).getValue()) || 0;
  }

  return {
    vendasHoje:      vendasTotal,
    vendasHojeCount: vendasCount,
    lucroMes:        lucroTotal,
    alertas:         stockAlertas.length,
    topProdutos:     topProdutos,
    pagamentos:      pagamentos,
    stockAlertas:    stockAlertas.slice(0, 5),
    totalDepenses:   totalDepenses,
    depensesCount:   depensesCount,
    depenses:        depenses,
    soldeTresorerie: soldeTresorerie
  };
}

function normalizarMetodoPagamento(method) {
  var payNorm = (method || '').toString().toLowerCase().replace(/[^a-z]/g, '');
  if (payNorm === 'express') return 'Express';
  if (payNorm === 'cartao' || payNorm === 'carte' || payNorm === 'card') return 'Cartao';
  if (payNorm === 'credito' || payNorm === 'credit') return 'Credito';
  return 'Cash';
}

function formatarMontantePagamento(valor) {
  var num = Math.round((parseFloat(valor) || 0) * 100) / 100;
  return num % 1 === 0 ? String(num) : String(num).replace(/\.?0+$/, '');
}

function formatarResumoPagamento(paymentLines) {
  if (!paymentLines || paymentLines.length === 0) return 'Cash';
  return paymentLines.map(function(p) {
    return normalizarMetodoPagamento(p.method) + ': ' + formatarMontantePagamento(p.montant);
  }).join(' + ');
}

function ensureFornecedorExists_(nome) {
  var name = String(nome || '').trim();
  if (!name) return;
  var sheet = getSS().getSheetByName('Fornecedores');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var target = name.toLowerCase();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === target) return;
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([[name, '', '', '']]);
}

function ensureSupplierDebtSheet_() {
  var sheet = getSS().getSheetByName('Divida Fornecedores');
  if (!sheet) throw new Error('Dette Fornecedores sheet not found');
  var headers = ['Data','Fornecedor','Type','Designaçao','Montant Achat','Paiement','Solde','Estatuto'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function ensureSupplierDebtSheet_() {
  var sheet = getSS().getSheetByName('Divida Fornecedores');
  if (!sheet) throw new Error('Divida Fornecedores sheet not found');
  var headers = ['Data','Fornecedor','Tipo','Designa\u00e7\u00e3o','Montante compra','Pagamento','Saldo','Estatuto'];
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function recomputeSupplierDebtSheet_(sheet) {
  sheet = sheet || ensureSupplierDebtSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var balances = {};
  var out = [];
  rows.forEach(function(row) {
    var supplier = String(row[1] || '').trim();
    var key = supplier.toLowerCase();
    if (!supplier) {
      out.push(['', '']);
      return;
    }
    var achat = parseFloat(row[4]) || 0;
    var paiement = parseFloat(row[5]) || 0;
    balances[key] = Math.max((balances[key] || 0) + achat - paiement, 0);
    out.push([balances[key], balances[key] > 0 ? 'Em curso' : 'Pago']);
  });
  sheet.getRange(2, 7, out.length, 2).setValues(out);
}

function getSupplierDebtBalance_(fournisseur) {
  var sheet = ensureSupplierDebtSheet_();
  recomputeSupplierDebtSheet_(sheet);
  var target = String(fournisseur || '').trim().toLowerCase();
  if (!target || sheet.getLastRow() <= 1) return 0;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][1] || '').trim().toLowerCase() === target) {
      return Math.max(parseFloat(rows[i][6]) || 0, 0);
    }
  }
  return 0;
}

function appendSupplierDebtRow_(date, fournisseur, type, designation, montantAchat, paiement) {
  var sheet = ensureSupplierDebtSheet_();
  var previous = getSupplierDebtBalance_(fournisseur);
  var achat = parseFloat(montantAchat) || 0;
  var paid = parseFloat(paiement) || 0;
  var solde = Math.max(previous + achat - paid, 0);
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 8).setValues([[
    date || getTodaySheetDate_(),
    fournisseur,
    type || '',
    designation || '',
    achat,
    paid,
    solde,
    solde > 0 ? 'Em curso' : 'Pago'
  ]]);
  return solde;
}
function splitPaymentTotals(paymentLines, totalVenda) {
  var totals = { Cash: 0, Express: 0, Cartao: 0, Credito: 0 };
  var lines = Array.isArray(paymentLines) ? paymentLines : [];
  lines.forEach(function(p) {
    var method = normalizarMetodoPagamento(p.method);
    totals[method] += parseFloat(p.montant) || 0;
  });
  var sum = totals.Cash + totals.Express + totals.Cartao + totals.Credito;
  if (Math.abs(sum - (parseFloat(totalVenda) || 0)) <= 0.01) return totals;
  return totals;
}

function roundMoney_(value) {
  return Math.round((parseFloat(value) || 0) * 100) / 100;
}

function getSaleInventoryName_(name) {
  return String(name || '').replace(/\s*\[[^\]]*\]\s*$/, '').trim();
}

function getPurchasePriceForSaleItem_(itemName, invData, invMap) {
  if (!invData || !invMap) return 0;
  var targetName = getSaleInventoryName_(itemName).toLowerCase();

  for (var i = 1; i < invData.length; i++) {
    var inventoryName = String(invData[i][invMap.name] || '').trim().toLowerCase();
    if (inventoryName !== targetName) continue;

    var metaPurchase = invMap.purchasePrice >= 0 ? (parseFloat(invData[i][invMap.purchasePrice]) || 0) : 0;
    var inventoryPrice = invMap.price >= 0 ? (parseFloat(invData[i][invMap.price]) || 0) : 0;
    return metaPurchase || inventoryPrice;
  }

  return 0;
}
function summarizePaymentColumns(row) {
  if (!row) return 'Cash';
  var paymentLines = [];
  var methods = [
    { key: 'Cash', value: parseFloat(row[4]) || 0 },
    { key: 'Express', value: parseFloat(row[5]) || 0 },
    { key: 'Cartao', value: parseFloat(row[6]) || 0 },
    { key: 'Credito', value: parseFloat(row[7]) || 0 }
  ];
  methods.forEach(function(entry) {
    if (entry.value > 0) paymentLines.push({ method: entry.key, montant: entry.value });
  });
  return formatarResumoPagamento(paymentLines);
}

function repartirMontantParPaiement(payText, rowMontant) {
  var repartition = { Cash: 0, Express: 0, Cartao: 0, Credito: 0 };
  var text = (payText || '').toString();
  var regex = /([A-Za-zA-y]+)\s*:\s*([0-9]+(?:[.,][0-9]+)?)/g;
  var parsed = [];
  var match;

  while ((match = regex.exec(text)) !== null) {
    var amount = parseFloat(String(match[2]).replace(',', '.')) || 0;
    if (amount <= 0) continue;
    parsed.push({
      method: normalizarMetodoPagamento(match[1]),
      montant: amount
    });
  }

  if (parsed.length > 0) {
    var totalPaiement = parsed.reduce(function(sum, p) { return sum + p.montant; }, 0);
    if (totalPaiement > 0) {
      parsed.forEach(function(p) {
        repartition[p.method] += (rowMontant * p.montant) / totalPaiement;
      });
      return repartition;
    }
  }

  repartition[normalizarMetodoPagamento(text)] += rowMontant || 0;
  return repartition;
}


// ============================================================
//  ENREGISTRER UNE VENTE
//
//  Recoit du POS :
//  vendaData.date      = date de la vente
//  vendaData.client    = nom client (ou "Anonimo")
//  vendaData.items     = liste produits {name, price, qty}
//  vendaData.paymentLines = liste des paiements multi-methode
//  vendaData.recibo    = numero du recu
//  vendaData.statut    = "interno" ou "Commande"
//
//  Colonnes ecrites dans Ventes (A a N) :
//  A=Date  B=Designationdds  C=Quantity  D=Unit Price
//  E=Cash  F=Express  G=Cartao  H=Credito
//  I=Total Amount  J=Purchase Price  K=Profit
//  L=Status  M=Client  N=Receipt No
// ============================================================
function registarVenda(vendaData) {
  Logger.log("vendaData");
  Logger.log(vendaData);
  var SS = getSS();
  var ventes = SS.getSheetByName('Vendas');
  if (!ventes) throw new Error('Ventes sheet not found');
  var inv = SS.getSheetByName('Inventario');
  var totalVenda = parseFloat(vendaData.total) || 0;

  if (!totalVenda && vendaData.items) {
    totalVenda = vendaData.items.reduce(function(sum, item) {
      return sum + ((parseFloat(item.price) || 0) * (parseFloat(item.qty) || 0));
    }, 0);
  }

  var paymentLines = Array.isArray(vendaData.paymentLines) ? vendaData.paymentLines.map(function(p) {
    return {
      method: normalizarMetodoPagamento(p.method),
      montant: parseFloat(p.montant) || 0
    };
  }).filter(function(p) {
    return p.montant > 0;
  }) : [];

  if (paymentLines.length === 0) {
    paymentLines = [{
      method: normalizarMetodoPagamento(vendaData.pagamento || 'Cash'),
      montant: totalVenda
    }];
  }

  var totalPaiement = paymentLines.reduce(function(sum, p) { return sum + p.montant; }, 0);
  if (Math.abs(totalPaiement - totalVenda) > 0.01) {
    throw new Error('Le total des paiements doit correspondre au total de la vente.');
  }

  var paymentTotals = splitPaymentTotals(paymentLines, totalVenda);
  var remainingPaymentTotals = {
    Cash: paymentTotals.Cash,
    Express: paymentTotals.Express,
    Cartao: paymentTotals.Cartao,
    Credito: paymentTotals.Credito
  };
  var invHeaders = [];
  var invData = [];
  var invMap = null;

  if (inv) {
    invHeaders = ensureProductMetadataColumns_(inv);
    invData = inv.getDataRange().getValues();
    invMap = getInventoryColumnMap_(invHeaders);
  }

  vendaData.items.forEach(function(item, itemIndex) {
    var saleProductName = String(item.name || '').trim();
    var inventoryName = getSaleInventoryName_(saleProductName);
    var pAchat = getPurchasePriceForSaleItem_(inventoryName, invData, invMap);
    var qty = parseFloat(item.qty) || 0;
    var price = parseFloat(item.price) || 0;
    var montant = qty * price;
    var benefice = (price - pAchat) * qty;
    var lastRow = ventes.getLastRow() + 1;
    var isLastItem = itemIndex === vendaData.items.length - 1;
    var ratio = totalVenda > 0 ? montant / totalVenda : 0;
    var linePaymentTotals = {};

    ['Cash', 'Express', 'Cartao', 'Credito'].forEach(function(method) {
      if (isLastItem) {
        linePaymentTotals[method] = roundMoney_(remainingPaymentTotals[method]);
        return;
      }

      linePaymentTotals[method] = roundMoney_(paymentTotals[method] * ratio);
      remainingPaymentTotals[method] = roundMoney_(remainingPaymentTotals[method] - linePaymentTotals[method]);
    });

    ventes.getRange(lastRow, 1, 1, 15).setValues([[
      vendaData.date,
      saleProductName || inventoryName,
      qty,
      price,
      linePaymentTotals.Cash,
      linePaymentTotals.Express,
      linePaymentTotals.Cartao,
      linePaymentTotals.Credito,
      montant,
      pAchat,
      benefice,
      vendaData.statut || 'interno',
      vendaData.vendedor ||"loja",
      vendaData.client,
      vendaData.recibo
    ]]);
  });

  atualizarCliente(
    vendaData.client,//nom
    vendaData.date,//date
    vendaData.items.reduce(function(sum, item) {
      return sum + ((parseFloat(item.price) || 0) * (parseFloat(item.qty) || 0));
    }, 0),//montant total acheter
    vendaData.items.reduce(function(sum, item) {
      return sum + (item.qty);
    }, 0),//nombre total achat
  );

  if (paymentTotals.Credito > 0) {
    registarDetteCliente(
      vendaData.client,//nom
      vendaData.date,//date
      paymentTotals.Credito,//dette
      vendaData.items.reduce(function(sum, item) {
        return sum + (item.qty);
      }, 0)//nombre total achat
    );
  }

  var totalRecebido = paymentTotals.Cash + paymentTotals.Express + paymentTotals.Cartao;
  if (totalRecebido > 0) {
    appendTresorerieMovement_(
      vendaData.date,
      'Venda',
      'Venda ' + (vendaData.recibo || ''),
      totalRecebido,
      0
    );
  }

  corrigirFormulasStockVendas_();
  invalidateProductsCache();
  return { success: true, recibo: vendaData.recibo };
}

function registarAchat(data) {
  var SS = getSS();
  var achat = SS.getSheetByName('Compra');
  if (!achat) throw new Error('Achat sheet not found');
  ensureFornecedorExists_(data.forn);

  var map = ensureAchatProductColumns_(achat);
  var row = buildAchatRow_(map, data.forn, {
    date: data.date,
    prod: data.prod,
    qty: data.qty,
    price: data.price,
    total: data.total,
    category: data.category,
    code: data.code,
    variation: data.variation,
    photo: data.photo,
    targetMargin: data.targetMargin
  });
  
  var lastRow = achat.getLastRow() + 1;
  achat.getRange(lastRow, 1, 1, row.length).setValues([row]);
  syncProductMetadataFromPurchase_([{ prod: data.prod, price: data.price, category: data.category, code: data.code, variation: data.variation, photo: data.photo, targetMargin: data.targetMargin }], data.forn);
  return true;
}


// ============================================================
//  ENREGISTRER UN ACHAT MULTIPLE
//  Recoit plusieurs lignes produits + plusieurs paiements
// ============================================================
function registarAchatMultiple(data) {
  var SS    = getSS();
  var achat = SS.getSheetByName('Compra');
  if (!achat) throw new Error('Achat sheet not found');
  ensureFornecedorExists_(data.forn);
  var map = ensureAchatProductColumns_(achat);
  var totalAchat = 0;

  // Ecrire toutes les lignes produits (batch)
  var rows = data.lines.map(function(l) {
    var lineTotal = (parseFloat(l.qty) || 0) * (parseFloat(l.price) || 0);
    totalAchat += lineTotal;
    return buildAchatRow_(map, data.forn, {
      date: l.date,
      prod: l.prod,
      qty: l.qty,
      price: l.price,
      total: lineTotal,
      category: l.category,
      code: l.code,
      variation: l.variation,
      photo: l.photo,
      targetMargin: l.targetMargin
    });
  });
  var firstRow = achat.getLastRow() + 1;
  achat.getRange(firstRow, 1, rows.length, rows[0].length).setValues(rows);
  syncProductMetadataFromPurchase_(data.lines, data.forn);

  // Si achat a credit -> enregistre dans Dette Fornecedores
  if (data.credit) {
    ensureSupplierDebtSheet_();
    data.lines.forEach(function(l) {
      appendSupplierDebtRow_(l.date, data.forn, 'Achat credit', l.prod, (parseFloat(l.qty) || 0) * (parseFloat(l.price) || 0), 0);
    });
    if (data.paiements && data.paiements.length > 0) {
      data.paiements.forEach(function(p) {
        if (!p.montant || p.montant <= 0) return;
        appendSupplierDebtRow_(p.date, data.forn, 'Paiement', '', 0, p.montant);
      });
    }
  }

  var montantPaye = data.credit
    ? (data.paiements || []).reduce(function(sum, p) { return sum + (parseFloat(p.montant) || 0); }, 0)
    : totalAchat;

  if (montantPaye > 0) {
    appendTresorerieMovement_(
      (data.lines[0] && data.lines[0].date) || data.date,
      'Compra',
      'Achat fournisseur ' + (data.forn || '') + ' - ' + buildItemsSummary_(data.lines),
      0,
      montantPaye
    );
  }

  invalidateProductsCache();
  return true;
}
//Paiement dette du client
function registarPagamentoClient(data) {
  var SS = getSS();
  var dette = SS.getSheetByName('Divida Clientes');
  if (!dette) throw new Error('Dette Clientes sheet not found');

  var montant = parseFloat(data.montant) || 0;

  if (!data.client || montant <= 0) {
    throw new Error('Cliente et montant obligatoires.');
  }

  var lastRow = dette.getLastRow();
  var detteActuelle = 0;

  if (lastRow > 1) {
    var rows = dette.getRange(2, 1, lastRow - 1, 8).getValues();

    for (var i = rows.length - 1; i >= 0; i--) {
      var nomClient = rows[i][1];

      if (
        String(nomClient).trim().toLowerCase() ===
        String(data.client).trim().toLowerCase()
      ) {
        detteActuelle = parseFloat(rows[i][6]) || 0; // G = restant
        break;
      }
    }
  }

  var resteDette = detteActuelle - montant;

  var newRow = dette.getLastRow() + 1;

  dette.getRange(newRow, 1, 1, 8).setValues([[
    data.date || getTodaySheetDate_(),
    data.client,
    'Paiement',
    0,
    0,
    montant,
    resteDette,
    resteDette <= 0 ? 'pago' : 'divida'
  ]]);

  appendTresorerieMovement_(
    data.date || getTodaySheetDate_(),
    'Paiement Client',
    'Paiement client ' + data.client + (data.note ? ' - ' + data.note : ''),
    montant,
    0
  );

  atualizarClienteDette_(data.client, -montant);

  return true;
}

function registarPagamentoForn(data) {
  var montant = parseFloat(data.montant) || 0;
  if (!data.forn || montant <= 0) throw new Error('Fornecedor et montant obligatoires.');
  ensureFornecedorExists_(data.forn);
  appendSupplierDebtRow_(data.date || getTodaySheetDate_(), data.forn, 'Paiement', data.note || '', 0, montant);

  appendTresorerieMovement_(
    data.date,
    'Paiement Fournisseur',
    'Paiement fournisseur ' + data.forn + (data.note ? ' - ' + data.note : ''),
    0,
    montant
  );

  return true;
}

function getResumoDettes() {
  var sheet = ensureSupplierDebtSheet_();
  if (!sheet || sheet.getLastRow() <= 1) return [];
  recomputeSupplierDebtSheet_(sheet);

  var data = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var forn = (data[i][1] || '').toString().trim();
    if (!forn) continue;

    if (!map[forn]) {
      map[forn] = { forn: forn, totalCompras: 0, totalPago: 0, saldo: 0, statut: 'Pago' };
    }

    map[forn].totalCompras += parseFloat(data[i][4]) || 0;
    map[forn].totalPago += parseFloat(data[i][5]) || 0;
  }

  return Object.keys(map).map(function(key) {
    map[key].saldo = Math.max(map[key].totalCompras - map[key].totalPago, 0);
    map[key].statut = map[key].saldo > 0 ? 'Em curso' : 'Pago';
    return map[key];
  }).sort(function(a, b) {
    return b.saldo - a.saldo;
  });
}

//MOI-MEME
function gettotaldettes(fournisseur){
  return getSupplierDebtBalance_(fournisseur);
}

function getFornecedorNames() {
  var sheet = getSS().getSheetByName('Fornecedores');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var seen = {};
  var names = [];
  data.forEach(function(row) {
    var name = String(row[0] || '').trim();
    var key = name.toLowerCase();
    if (!name || seen[key]) return;
    seen[key] = true;
    names.push(name);
  });
  return names.sort();
}

function registarTresorerie(data) {
  var montant = parseFloat(data.montant) || 0;
  if (montant <= 0) throw new Error('Montant invalide.');

  var mouvement = (data.mouvement || '').toString().toLowerCase();
  var income = mouvement === 'entrada' ? montant : 0;
  var expense = mouvement === 'saida' ? montant : 0;

  if (!income && !expense) throw new Error('Choisis Entrada ou Saida.');

  var balance = appendTresorerieMovement_(
    data.date,
    data.tipo || (income ? 'Entrada Manual' : 'Saida Manual'),
    data.desc || 'Mouvement manuel',
    income,
    expense
  );

  return { success: true, balance: balance };
}

function getTresorerie(params) {
  var sheet = getTresorerieSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { balance: 0, totalIn: 0, totalOut: 0, count: 0, entries: [] };
  }

  params = params || {};
  var from = params.from ? parseSheetDate(params.from) : null;
  var to = params.to ? parseSheetDate(params.to) : null;
  if (to) to.setHours(23, 59, 59, 999);
  var typeFilter = (params.type || '').toString().toLowerCase();
  var limit = parseInt(params.limit, 10) || 50;
  var data = sheet.getDataRange().getValues();
  var entries = [];
  var totalIn = 0;
  var totalOut = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var rowDate = parseSheetDate(row[0]);
    if (!rowDate) continue;
    if (from && rowDate < from) continue;
    if (to && rowDate > to) continue;

    var type = (row[1] || '').toString();
    if (typeFilter && type.toLowerCase().indexOf(typeFilter) < 0) continue;

    var income = parseFloat(row[3]) || 0;
    var expense = parseFloat(row[4]) || 0;
    totalIn += income;
    totalOut += expense;

    entries.push({
      date: rowDate.toLocaleDateString('pt-PT'),
      type: type,
      desc: row[2] || '',
      income: income,
      expense: expense,
      balance: parseFloat(row[5]) || 0
    });
  }

  entries.reverse();
  if (entries.length > limit) entries = entries.slice(0, limit);

  return {
    balance: getLastTresorerieBalance_(sheet),
    totalIn: totalIn,
    totalOut: totalOut,
    count: entries.length,
    entries: entries
  };
}


// ============================================================
//  ENREGISTRER UN TRANSFERT (Armazem -> Boutique)
//  Colonnes dans Transferts (A a D) :
//  A=Date  B=Designation  C=Quantity  D=Note
// ============================================================
function registarTransferencia(data) {
  var SS = getSS();
  var trans = SS.getSheetByName('Transferencias');
  if (!trans) throw new Error('Transferts sheet not found');
  var lastRow = trans.getLastRow() + 1;
  trans.getRange(lastRow, 1, 1, 4).setValues([[
    data.date,
    data.prod,
    data.qty,
    data.obs || ''
  ]]);
  return true;
}


// ============================================================
//  RECUPERER L'HISTORIQUE DES VENTES
//  params.from    = date debut
//  params.to      = date fin
//  params.search  = texte recherche (produit ou client)
// ============================================================
function getVentes(params) {
  var SS = getSS();
  var ventes = SS.getSheetByName('Vendas');
  if (!ventes) return [];

  params = params || {};
  var data = ventes.getDataRange().getValues();
  var headers = data.length ? data[0].map(function(h) { return normalizeDashboardText_(h); }) : [];
  var ixClient = headers.indexOf('cliente') >= 0 ? headers.indexOf('cliente') : 13;
  var ixRecibo = headers.indexOf('recibo no') >= 0 ? headers.indexOf('recibo no') : 14;
  var result = [];
  var from = params.from ? parseSheetDate(params.from) : null;
  var to = params.to ? parseSheetDate(params.to) : null;
  if (to) to.setHours(23, 59, 59, 999);
  var search = params.search ? params.search.toLowerCase() : '';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rowDate = parseSheetDate(row[0]);
    if (!rowDate) continue;
    if (from && rowDate < from) continue;
    if (to && rowDate > to) continue;
    var prod = (row[1] || '').toString().toLowerCase();
    var client = (row[ixClient] || '').toString().toLowerCase();
    if (search && !prod.includes(search) && !client.includes(search)) continue;
    result.push({
      date: rowDate.toLocaleDateString('pt-PT'),
      prod: row[1] || '',
      client: row[ixClient] || '',
      qty: row[2] || 0,
      punit: row[3] || 0,
      total: row[8] || 0,
      statut: row[11] || '',
      cash: row[4] || '',
      express: row[5] || '',
      cartao: row[6] || '',
      credito: row[7] || '',
      pay: summarizePaymentColumns(row),
      recibo: row[ixRecibo] || ''
    });
  }
  return result.reverse();
}

function getConsignationsSheet_() {
  var ss = getSS();
  var sheet = ss.getSheetByName('Revendedores');
  if (!sheet) {
    sheet = ss.insertSheet('Revendedores');
    sheet.getRange(1, 1, 1, 12).setValues([['ID','Data','Revendeur','Designaçao','Quantidade','Preço unitaro','Montante total','Estatuto','data de pagamento','data de regresso','Recibo No','Pagamento']]);
  }
  return sheet;
}

function generateConsignationId_() {
  return 'CON-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Luanda', 'yyMMdd-HHmmss');
}

function getConsignationRowsById_(consSheet, consignationId) {
  var data = consSheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString() === consignationId) {
      rows.push({ rowNumber: i + 1, values: data[i] });
    }
  }
  return rows;
}

function registarConsignacao(data) {
  var sheet = getConsignationsSheet_();
  var revendeur = (data.revendeur || '').toString().trim();
  var items = Array.isArray(data.items) ? data.items : [];
  if (!revendeur) throw new Error('Nom du revendeur obligatoire.');
  if (!items.length) throw new Error('Aucun article dans la consignation.');

  var consignationId = generateConsignationId_();
  var rows = [];

  items.forEach(function(item) {
    var qty = parseFloat(item.qty) || 0;
    var price = parseFloat(item.price) || 0;
    if (!item.name || qty <= 0 || price < 0) return;
    rows.push([
      consignationId,
      data.date || getTodaySheetDate_(),
      revendeur,
      item.name,
      qty,
      price,
      qty * price,
      "Em curso",
      '',
      '',
      '',
      ''
    ]);
  });

  if (!rows.length) throw new Error('Articles invalides dans la consignation.');

  var firstRow = sheet.getLastRow() + 1;
  sheet.getRange(firstRow, 1, rows.length, 12).setValues(rows);
  invalidateProductsCache();
  return { success: true, id: consignationId };
}

function confirmerPaiementConsignation(data) {
  var consSheet = getConsignationsSheet_();
  var consignationId = (data.id || '').toString().trim();
  if (!consignationId) throw new Error('ID de consignation obligatoire.');

  var rows = getConsignationRowsById_(consSheet, consignationId);
  if (!rows.length) throw new Error('Consignation introuvable.');

  var openRows = rows.filter(function(row) {
    return (row.values[7] || '').toString() === "Em curso";
  });
  if (!openRows.length) throw new Error('Cette consignation nest plus en cours.');

  var paymentLines = Array.isArray(data.paymentLines) ? data.paymentLines : [];
  var vendaData = {
    date: data.date || openRows[0].values[1] || getTodaySheetDate_(),
    client: openRows[0].values[2] || 'Revendeur',
    items: openRows.map(function(row) {
      return {
        name: row.values[3],
        qty: parseFloat(row.values[4]) || 0,
        price: parseFloat(row.values[5]) || 0
      };
    }),
    total: openRows.reduce(function(sum, row) { return sum + (parseFloat(row.values[6]) || 0); }, 0),
    pagamento: data.pagamento || formatarResumoPagamento(paymentLines),
    paymentLines: paymentLines,
    statut: 'Consignation',
    recibo: data.recibo || ('CONS-' + consignationId.replace(/[^0-9A-Z-]/gi, ''))
  };

  registarVenda(vendaData);

  openRows.forEach(function(row) {
    consSheet.getRange(row.rowNumber, 8, 1, 5).setValues([[
      'Pago',
      vendaData.date,
      '',
      vendaData.recibo,
      vendaData.pagamento || ''
    ]]);
  });

  invalidateProductsCache();
  return { success: true, recibo: vendaData.recibo };
}

function retornarConsignacao(data) {
  var consSheet = getConsignationsSheet_();
  var consignationId = (data.id || '').toString().trim();
  if (!consignationId) throw new Error('ID de consignation obligatoire.');

  var rows = getConsignationRowsById_(consSheet, consignationId);
  if (!rows.length) throw new Error('Consignation introuvable.');

  var hasOpen = false;
  rows.forEach(function(row) {
    if ((row.values[7] || '').toString() === "Em curso") {
      hasOpen = true;
      consSheet.getRange(row.rowNumber, 8, 1, 5).setValues([[
        'Retournee',
        '',
        data.date || getTodaySheetDate_(),
        '',
        ''
      ]]);
    }
  });

  if (!hasOpen) throw new Error('Aucune ligne en cours a retourner.');

  invalidateProductsCache();
  return { success: true };
}

function getConsignationsOpen() {
  var sheet = getConsignationsSheet_();
  if (sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if ((row[7] || '').toString() !== "Em curso") continue;

    var id = (row[0] || '').toString();
    if (!map[id]) {
      map[id] = {
        id: id,
        date: row[1],
        revendeur: row[2] || '',
        total: 0,
        qty: 0,
        items: []
      };
    }

    map[id].qty += parseFloat(row[4]) || 0;
    map[id].total += parseFloat(row[6]) || 0;
    map[id].items.push((row[3] || '') + ' x' + (row[4] || 0));
  }

  return Object.keys(map).map(function(id) {
    map[id].date = parseSheetDate(map[id].date) ? parseSheetDate(map[id].date).toLocaleDateString('pt-PT') : (map[id].date || '');
    return map[id];
  }).sort(function(a, b) {
    return b.id.localeCompare(a.id);
  });
}

function getRevendeurDetail(nom) {
  var sheet = getConsignationsSheet_();
  var reseller = (nom || '').toString().trim().toLowerCase();
  if (!reseller || sheet.getLastRow() <= 1) return null;

  var data = sheet.getDataRange().getValues();
  var groups = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if ((row[2] || '').toString().trim().toLowerCase() !== reseller) continue;
    var id = (row[0] || '').toString();
    if (!groups[id]) {
      groups[id] = {
        id: id,
        date: row[1],
        revendeur: row[2] || '',
        status: row[7] || "Em curso",
        total: 0,
        qty: 0,
        payDate: row[8] || '',
        returnDate: row[9] || '',
        recibo: row[10] || '',
        payment: row[11] || '',
        items: []
      };
    }
    groups[id].total += parseFloat(row[6]) || 0;
    groups[id].qty += parseFloat(row[4]) || 0;
    groups[id].items.push({
      prod: row[3] || '',
      qty: parseFloat(row[4]) || 0,
      price: parseFloat(row[5]) || 0,
      total: parseFloat(row[6]) || 0
    });
    if ((row[7] || '').toString() === "Em curso") groups[id].status = "Em curso";
    else if (groups[id].status !== "Em curso") groups[id].status = row[7] || groups[id].status;
  }

  var all = Object.keys(groups).map(function(id) {
    var g = groups[id];
    var parsed = parseSheetDate(g.date);
    g.date = parsed ? parsed.toLocaleDateString('pt-PT') : (g.date || '');
    return g;
  }).sort(function(a, b) {
    return b.id.localeCompare(a.id);
  });

  if (!all.length) return null;

  return {
    nom: all[0].revendeur,
    totalPossession: all.filter(function(g) { return g.status === "Em curso"; }).reduce(function(sum, g) { return sum + g.total; }, 0),
    openCount: all.filter(function(g) { return g.status === "Em curso"; }).length,
    ouvertes: all.filter(function(g) { return g.status === "Em curso"; }),
    historique: all
  };
}

function getRevendeurNames(data) {
  var sheet = getConsignationsSheet_();
  if (sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var seen = {};
  var names = [];

  for (var i = 1; i < data.length; i++) {
    var name = (data[i][2] || '').toString().trim();
    var key = name.toLowerCase();
    if (!name || seen[key]) continue;
    seen[key] = true;
    names.push(name);
  }

  return names.sort();
}

function getConsignationsByRevendeur(nom) {
  var detail = getRevendeurDetail(nom);
  return detail ? detail.ouvertes : [];
}

function confirmerPaiementConsignations(data) {
  var consSheet = getConsignationsSheet_();
  var ids = Array.isArray(data && data.ids) ? data.ids.map(function(id) {
    return (id || '').toString().trim();
  }).filter(String) : [];
  if (!ids.length) throw new Error('Choisis au moins une consignation.');

  var paymentLines = Array.isArray(data.paymentLines) ? data.paymentLines : [];
  var allOpenRows = [];
  var revendeur = '';

  ids.forEach(function(consignationId) {
    var rows = getConsignationRowsById_(consSheet, consignationId);
    if (!rows.length) throw new Error('Consignation introuvable: ' + consignationId);

    var openRows = rows.filter(function(row) {
      return (row.values[7] || '').toString() === "Em curso";
    });
    if (!openRows.length) throw new Error('Cette consignation n est plus en cours: ' + consignationId);

    var rowRevendeur = (openRows[0].values[2] || '').toString().trim();
    if (!revendeur) revendeur = rowRevendeur;
    if (revendeur.toLowerCase() !== rowRevendeur.toLowerCase()) {
      throw new Error('Tu ne peux payer que les consignations d un seul revendeur a la fois.');
    }

    Array.prototype.push.apply(allOpenRows, openRows);
  });

  var total = allOpenRows.reduce(function(sum, row) {
    return sum + (parseFloat(row.values[6]) || 0);
  }, 0);

  var recibo = data.recibo || ('CONS-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Luanda', 'yyMMdd-HHmmss'));
  var vendaData = {
    date: data.date || getTodaySheetDate_(),
    client: revendeur || 'Revendeur',
    items: allOpenRows.map(function(row) {
      return {
        name: row.values[3],
        qty: parseFloat(row.values[4]) || 0,
        price: parseFloat(row.values[5]) || 0
      };
    }),
    total: total,
    pagamento: data.pagamento || formatarResumoPagamento(paymentLines),
    paymentLines: paymentLines,
    provenance: 'interno',
    vendedor: "Revendedor",
    recibo: recibo
  };

  registarVenda(vendaData);

  allOpenRows.forEach(function(row) {
    consSheet.getRange(row.rowNumber, 8, 1, 5).setValues([[
      'Pago',
      vendaData.date,
      '',
      vendaData.recibo,
      vendaData.pagamento || ''
    ]]);
  });

  invalidateProductsCache();
  return { success: true, recibo: vendaData.recibo };
}

function confirmerPaiementConsignation(data) {
  return confirmerPaiementConsignations({
    ids: data && data.id ? [data.id] : [],
    date: data && data.date,
    pagamento: data && data.pagamento,
    paymentLines: data && data.paymentLines,
    recibo: data && data.recibo
  });
}

function retornarConsignacoes(data) {
  var consSheet = getConsignationsSheet_();
  var ids = Array.isArray(data && data.ids) ? data.ids.map(function(id) {
    return (id || '').toString().trim();
  }).filter(String) : [];
  if (!ids.length) throw new Error('Choisis au moins une consignation.');

  var changed = 0;
  ids.forEach(function(consignationId) {
    var rows = getConsignationRowsById_(consSheet, consignationId);
    if (!rows.length) throw new Error('Consignation introuvable: ' + consignationId);

    rows.forEach(function(row) {
      if ((row.values[7] || '').toString() === "Em curso") {
        changed++;
        consSheet.getRange(row.rowNumber, 8, 1, 5).setValues([[
          'Retournee',
          '',
          data.date || getTodaySheetDate_(),
          '',
          ''
        ]]);
      }
    });
  });

  if (!changed) throw new Error('Aucune ligne en cours a retourner.');

  invalidateProductsCache();
  return { success: true, count: changed };
}

function retornarConsignacao(data) {
  return retornarConsignacoes({
    ids: data && data.id ? [data.id] : [],
    date: data && data.date
  });
}

function getHistoriqueConsignations(filters) {
  var sheet = getConsignationsSheet_();
  if (sheet.getLastRow() <= 1) return [];

  filters = filters || {};
  var resellerFilter = (filters.revendeur || '').toString().trim().toLowerCase();
  var from = filters.from ? parseSheetDate(filters.from) : null;
  var to = filters.to ? parseSheetDate(filters.to) : null;
  if (to) to.setHours(23, 59, 59, 999);

  var data = sheet.getDataRange().getValues();
  var groups = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var revendeur = (row[2] || '').toString().trim();
    if (resellerFilter && revendeur.toLowerCase().indexOf(resellerFilter) < 0) continue;

    var id = (row[0] || '').toString();
    if (!id) continue;

    if (!groups[id]) {
      groups[id] = {
        id: id,
        date: row[1] || '',
        revendeur: revendeur,
        status: row[7] || "Em curso",
        payDate: row[8] || '',
        returnDate: row[9] || '',
        recibo: row[10] || '',
        payment: row[11] || '',
        total: 0,
        qty: 0,
        items: []
      };
    }

    groups[id].total += parseFloat(row[6]) || 0;
    groups[id].qty += parseFloat(row[4]) || 0;
    groups[id].items.push({
      prod: row[3] || '',
      qty: parseFloat(row[4]) || 0,
      price: parseFloat(row[5]) || 0,
      total: parseFloat(row[6]) || 0
    });

    if ((row[7] || '').toString() === "Em curso") groups[id].status = "Em curso";
    else if (groups[id].status !== "Em curso") groups[id].status = row[7] || groups[id].status;
    if (!groups[id].payDate && row[8]) groups[id].payDate = row[8];
    if (!groups[id].returnDate && row[9]) groups[id].returnDate = row[9];
    if (!groups[id].recibo && row[10]) groups[id].recibo = row[10];
    if (!groups[id].payment && row[11]) groups[id].payment = row[11];
  }

  return Object.keys(groups).map(function(id) {
    var g = groups[id];
    var baseDate = parseSheetDate(g.date);
    var payDate = parseSheetDate(g.payDate);
    var returnDate = parseSheetDate(g.returnDate);
    var statusText = (g.status || '').toString().toLowerCase();
    var actionDate = statusText.indexOf('pay') === 0 ? payDate : (statusText.indexOf('retour') === 0 ? returnDate : baseDate);

    if (from && actionDate && actionDate < from) return null;
    if (to && actionDate && actionDate > to) return null;

    g.date = baseDate ? baseDate.toLocaleDateString('pt-PT') : (g.date || '');
    g.payDate = payDate ? payDate.toLocaleDateString('pt-PT') : (g.payDate || '');
    g.returnDate = returnDate ? returnDate.toLocaleDateString('pt-PT') : (g.returnDate || '');
    g.actionDate = actionDate ? actionDate.toLocaleDateString('pt-PT') : '';
    g.itemsSummary = g.items.map(function(item) {
      return item.prod + ' x' + item.qty;
    }).join(', ');
    return g;
  }).filter(function(row) {
    return !!row;
  }).sort(function(a, b) {
    return (b.actionDate || '').localeCompare(a.actionDate || '') || b.id.localeCompare(a.id);
  });
}


// ============================================================
//  RECUPERER STOCK ARMAZEM (pour bouton "Tudo para Boutique")
// ============================================================
function getStockArmazem() {
  var SS = getSS();
  var inv = SS.getSheetByName('Inventario');
  if (!inv) return [];

  var data   = inv.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    // Stock Armazem = col F (index 5) en mode armazem
    var stockArmazem = parseFloat(data[i][5]) || 0;
    if (name && stockArmazem > 0) {
      result.push({ name: name, qty: stockArmazem });
    }
  }
  return result;
}

function invalidateProductsCache() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache) cache.remove('products');
  } catch (e) {
    // Aucun cache actif ou service indisponible: on ignore sans bloquer l'achat.
  }
  return true;
}


// ============================================================
//  TRANSFERER TOUT VERS BOUTIQUE
//  Appele depuis le POS uniquement (pas depuis Apps Script)
// ============================================================
function transferirTudo(data) {
  var SS = getSS();
  var trans = SS.getSheetByName('Transferencias');
  if (!trans) throw new Error('Transferts sheet not found');
  data.items.forEach(function(item) {
    var lastRow = trans.getLastRow() + 1;
    trans.getRange(lastRow, 1, 1, 4).setValues([[
      data.date, item.name, item.qty, 'Full transfer to Boutique'
    ]]);
  });
  return true;
}


// ============================================================
//  METTRE A JOUR LE CLIENT apres chaque vente
//  Appele automatiquement par registarVenda
// ============================================================
function isClienteAnonimo_(nome) {
  var value = String(nome || '').trim().toLowerCase();
  return !value || value === 'anonimo' || value === 'anonymous';
}

function findClienteRow_(clientes, nome) {
  if (!clientes || isClienteAnonimo_(nome)) return { row: -1, data: [] };
  var data = clientes.getDataRange().getValues();
  var target = String(nome || '').trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === target) {
      return { row: i + 1, data: data[i] };
    }
  }

  return { row: -1, data: [] };
}

function atualizarClienteDette_(nome, delta) {
  var SS = getSS();
  var clientes = SS.getSheetByName('Clientes');
  if (!clientes || isClienteAnonimo_(nome)) return;

  var montant = parseFloat(delta) || 0;
  var match = findClienteRow_(clientes, nome);

  if (match.row === -1) {
    var newRow = clientes.getLastRow() + 1;
    clientes.getRange(newRow, 1, 1, 7).setValues([[
      nome, '', '', getTodaySheetDate_(), 0, Math.max(montant, 0), 0
    ]]);
    return;
  }

  var currentDette = parseFloat(match.data[5]) || 0;
  clientes.getRange(match.row, 6).setValue(Math.max(currentDette + montant, 0));
}

function atualizarCliente(nome, date, totalAchat, nbreAchat) {
  var SS = getSS();
  var clientes = SS.getSheetByName('Clientes');
  if (!clientes || isClienteAnonimo_(nome)) return;

  var montant = parseFloat(totalAchat) || 0;
  var qty = parseFloat(nbreAchat) || 0;
  var match = findClienteRow_(clientes, nome);

  if (match.row === -1) {
    var lastRow = clientes.getLastRow() + 1;
    clientes.getRange(lastRow, 1, 1, 7).setValues([[
      nome, '', '', date, montant, 0, qty
    ]]);
    return;
  }

  var firstPurchase = match.data[3] || date;
  var currentTotal = parseFloat(match.data[4]) || 0;
  var currentQty = parseFloat(match.data[6]) || 0;

  clientes.getRange(match.row, 4, 1, 4).setValues([[
    firstPurchase,
    currentTotal + montant,
    parseFloat(match.data[5]) || 0,
    currentQty + qty
  ]]);
}
//MOI-MEME 
function registarDetteCliente(nome, date, montantDette, nbreAchat) {
  var SS = getSS();
  var clientes = SS.getSheetByName('Divida Clientes');

  if (!clientes || isClienteAnonimo_(nome)) return;

  var nouvelleDette = parseFloat(montantDette) || 0;
  if (nouvelleDette <= 0) return;
  var dettePrecedente = 0;
  var restantPrecedent = 0;

  var lastRow = clientes.getLastRow();

  if (lastRow > 1) {
    var rows = clientes.getRange(2, 1, lastRow - 1, 8).getValues();

    for (var i = rows.length - 1; i >= 0; i--) {
      var nomClient = rows[i][1];      // B = Client
      var designation = rows[i][2];    // C = Designation

      if (
        String(nomClient).trim().toLowerCase() === String(nome).trim().toLowerCase() &&
        String(designation).trim().toLowerCase() === 'divida'
      ) {
        dettePrecedente = parseFloat(rows[i][4]) || 0;   // E = Dette
        restantPrecedent = parseFloat(rows[i][6]) || 0;  // G = restant
        break;
      }
    }
  }

  var detteActuelle = dettePrecedente + nouvelleDette;
  var restantActuel = restantPrecedent + nouvelleDette;

  var newRow = clientes.getLastRow() + 1;

  clientes.getRange(newRow, 1, 1, 8).setValues([[
    date,
    nome,
    'divida',
    nbreAchat,
    detteActuelle,
    0,
    restantActuel,
    'divida'
  ]]);

  atualizarClienteDette_(nome, nouvelleDette);
}

function getClientDetail(params) {
  var SS = getSS();
  var clientes = SS.getSheetByName('Clientes');
  if (!clientes) return [];

  var data = clientes.getDataRange().getValues();
  var result = [];
  
  var search = params ? params.toLowerCase() : '';

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var name = (row[0] || '').toString().toLowerCase();;
 
    if (search && !name.includes(search) ) continue;
    result.push({
      name: row[0] || '',
      phone: row[1] || '',
      email: row[2] || 0,
      premierAchat: row[3] || 0,
      totalAchat: row[4] || 0,
      totaldette: row[5] || 0,
      totalqty: row[6] || 0,
    });
  }
  return result.reverse();
}

function normalizeClientKey_(value) {
  return String(value || '').trim().toLowerCase();
}

function getClientDebt(data) {
  var clientName = data.clientName || "";

  var sheet = getSS().getSheetByName('Divida Clientes');
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  var target = normalizeClientKey_(clientName);
  if (!target) return 0;

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    if (normalizeClientKey_(rows[i][1]) === target) {
      return Math.max(parseFloat(rows[i][6]) || 0, 0);
    }
  }

  return 0;
}

function getFournDebt(data) {
  var fournName = data.fournName || "";

  var sheet = getSS().getSheetByName('Divida Fornecedores');
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  var target = normalizeClientKey_(fournName);
  if (!target) return 0;

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  for (var i = rows.length - 1; i >= 0; i--) {
    if (normalizeClientKey_(rows[i][1]) === target) {
      return Math.max(parseFloat(rows[i][6]) || 0, 0);
    }
  }

  return 0;
}


function getClientFicheData(clientName) {
  var SS = getSS();
  var ventes = SS.getSheetByName('Vendas');
  var target = normalizeClientKey_(clientName);
  if (!target || isClienteAnonimo_(target)) return null;

  var name = String(clientName || '').trim();
  var totalAchat = 0;
  var totalQty = 0;
  var totaldette = 0;
  var receiptMap = {};
  var historique = [];
  var quantite = 0;

  if (ventes && ventes.getLastRow() > 1) {
    var data = ventes.getDataRange().getValues();
    var headers = data[0].map(function(h) { return normalizeDashboardText_(h); });
    var ixDate = headers.indexOf('data') >= 0 ? headers.indexOf('data') : 0;
    var ixProd = headers.indexOf('Designação') >= 0 ? headers.indexOf('Designação') : 1;
    var ixQty = headers.indexOf('Quantidade') >= 0 ? headers.indexOf('Quantidade') : 2;
    var ixPunit = headers.indexOf('preco unitaro') >= 0 ? headers.indexOf('preco unitaro') : 3;
    var ixCash = headers.indexOf('cash') >= 0 ? headers.indexOf('cash') : 4;
    var ixExpress = headers.indexOf('express') >= 0 ? headers.indexOf('express') : 5;
    var ixCartao = headers.indexOf('Cartão') >= 0 ? headers.indexOf('Cartão') : 6;
    var ixCredito = headers.indexOf('Crédito') >= 0 ? headers.indexOf('Crédito') : 7;
    var ixTotal = headers.indexOf('montante total') >= 0 ? headers.indexOf('montante total') : 8;
    var ixClient = headers.indexOf('cliente') >= 0 ? headers.indexOf('cliente') : 13;
    var ixRecibo = headers.indexOf('Nº Recibo') >= 0 ? headers.indexOf('Nº Recibo') : 14;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[ixDate]) continue;
      if (normalizeClientKey_(row[ixClient]) !== target) continue;
      
      var rowDate = parseSheetDate(row[ixDate]);
      var qty = parseFloat(row[ixQty]) || 0;
      var total = parseFloat(row[ixTotal]) || 0;
      var recibo = row[ixRecibo] || '';

      quantite += 1;
      totalAchat += total;
      totalQty += qty;
      totaldette += row[ixCredito];
      if (recibo) receiptMap[String(recibo)] = true;
      else receiptMap['row-' + i] = true;
      Logger.log(quantite);
      historique.push({
        date: rowDate ? rowDate.toLocaleDateString('pt-PT') : row[ixDate],
        rawTime: rowDate ? rowDate.getTime() : 0,
        prod: row[ixProd] || '',
        qty: qty,
        punit: parseFloat(row[ixPunit]) || 0,
        cash: parseFloat(row[ixCash]) || 0,
        express: parseFloat(row[ixExpress]) || 0,
        cartao: parseFloat(row[ixCartao]) || 0,
        credito: parseFloat(row[ixCredito]) || 0,
        total: total,
        pay: summarizePaymentColumns(row),
        recibo: recibo
      });
    }
  }

  historique.sort(function(a, b) {
    return (b.rawTime || 0) - (a.rawTime || 0);
  });

  var clientes = SS.getSheetByName('Clientes');
  if (clientes && clientes.getLastRow() > 1) {
    var match = findClienteRow_(clientes, name);
    if (match.row !== -1 && match.data[0]) name = match.data[0];
  }

  return {
    name: name,
    totalAchat: totalAchat,
    totalDette: parseFloat(totaldette) || 0,//getClientDebt(name),
    totalQty: totalQty,
    transactions: quantite,//Object.keys(receiptMap).length,
    historique: historique.slice(0, 100)
  };
}


// ============================================================
//  PROTECTION DES FEUILLES
//
//  COMMENT CA MARCHE :
//  - protegerFeuilles() verrouille toutes les feuilles
//    Personne ne peut modifier directement dans le tableur
//  - activarModoEdicao() deverrouille pendant 1 minute
//    Apres 1 minute, les feuilles se reverrouillent automatiquement
//
//  UTILISATION :
//  Menu "Azul Gestão" > "Protect All Sheets" pour verrouiller
//  Menu "Azul Gestão" > "Edit Mode (1 min)" pour deverrouiller
//  OU depuis le POS : bouton dans Definicoes
//
//   IMPORTANT : Seul le proprietaire du fichier peut
//  proteger/deproteger les feuilles. Les autres utilisateurs
//  ne pourront pas modifier meme avec le bouton.
// ============================================================

// ============================================================
//  SYSTEME DE PROTECTION PAR TRIGGER onEdit
//
//  Google Sheets ne peut pas bloquer le proprietaire via
//  les protections normales. La solution : un trigger onEdit
//  qui detecte et annule immediatement toute modification
//  directe dans les feuilles protegees.
//
//  COMMENT CA MARCHE :
//  - activerProtection() installe un trigger onEdit
//    Chaque modification directe est annulee automatiquement
//  - activarModoEdicaoPOS() desactive le trigger pendant 1 min
//    Apres 1 minute, le trigger se reinstalle
//
//  UTILISATION :
//  1. Menu "Azul Gestão" > "Protect All Sheets" pour activer
//  2. POS > Definicoes > "Activar Modo Edicao" pour 1 minute
// ============================================================

var FEUILLES_PROTEGEES = [
  'Compra', 'Vendas', 'Transferencias', 'Clientes',
  'Fornecedores', 'Divida Fornecedores', 'Divida Clientes',
  'Despesas', 'Tesouraria', 'Revendedores', 'Immobilizaçao'
  // Inventaire exclu car ses formules ont besoin de se recalculer
];

// Cle stockee dans les proprietes du script pour savoir
// si le mode edition est actif ou non
var PROP_KEY = 'editModeActive';

//  TRIGGER onEdit 
// Cette fonction est appelee automatiquement a CHAQUE
// modification dans le tableur.
// Si le mode edition n'est pas actif -> annule la modification
function onEditProtection(e) {
  // Verifier si le mode edition est actif
  var props = PropertiesService.getScriptProperties();
  var editMode = props.getProperty(PROP_KEY);

  if (editMode === 'true') return; // Mode edition actif -> laisser passer

  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();

  // Verifier si c'est une feuille protegee
  var estProtegee = FEUILLES_PROTEGEES.indexOf(sheetName) >= 0;
  if (!estProtegee) return; // Pas une feuille protegee -> laisser passer

  // Annuler la modification - restaurer l'ancienne valeur
  var range = e.range;
  var oldValue = e.oldValue;

  if (oldValue !== undefined) {
    range.setValue(oldValue); // Restaure l'ancienne valeur
  } else {
    range.clearContent(); // Nouvelle cellule -> effacer
  }

  // Afficher un message d'erreur
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Direct editing is blocked. Use the POS to make changes. Go to POS > Definicoes > Activar Modo Edicao to edit directly.',
    'Blocked',
    5
  );
}

//  ACTIVER LA PROTECTION 
// Installe le trigger onEdit et desactive le mode edition
function protegerFeuilles(data) {
  var SS = getSS();

  FEUILLES_PROTEGEES.forEach(function(name) {
    var sheet = SS.getSheetByName(name);
    if (!sheet) return;

    var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existing.forEach(function(p) {
      if (p.canEdit()) p.remove();
    });

    var protection = sheet.protect();
    protection.setDescription('Azul Gestão Lock');

    var editors = protection.getEditors();
    if (editors.length > 0) {
      protection.removeEditors(editors);
    }

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    protection.setWarningOnly(true);
  });

  PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'false');

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditProtection') {
      ScriptApp.deleteTrigger(t);
    }
  });

  return {
    ok: true,
    message: 'Protection activée ! Utilisez POS > Definições > Edit Mode pour modifier.'
  };
}

//  DESACTIVER LA PROTECTION (depuis le menu) 
function activarModoEdicao(data) {
  _activarEdicaoInterna();

  return {
    ok: true,
    message: 'Modo de edição ativo por 1 minuto. Será bloqueado automaticamente.'
  };
}

//  DESACTIVER LA PROTECTION (depuis le POS) 
function activarModoEdicaoPOS() {
  _activarEdicaoInterna();
  return { success: true, message: 'Edit mode active for 1 minute' };
}

//  FONCTION INTERNE 
function _activarEdicaoInterna() {
  // Activer le mode edition
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'true');

  // Supprimer anciens triggers de reverrouillage
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'reverrouiller') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Reverrouille automatiquement apres 1 minute
  ScriptApp.newTrigger('reverrouiller')
    .timeBased()
    .after(60 * 1000)
    .create();
}

//  REVERROUILLER (appele automatiquement apres 1 min) 
function reverrouiller() {
  // Desactiver le mode edition
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, 'false');

  // Supprimer ce trigger
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'reverrouiller') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Notifier dans le tableur
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sheets are now locked again.',
    'Locked',
    3
  );
}

// ============================================================
//  ENREGISTRER UNE DEPENSE
// ============================================================
function registarDepense(data) {
  var SS = getSS();
  var sheet = SS.getSheetByName('Despesas');
  if (!sheet) throw new Error('Depenses sheet not found');
  var montant = parseFloat(data.montant) || 0;
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 5).setValues([[
    data.date, data.tipo, data.desc, 0, montant
  ]]);
  appendTresorerieMovement_(data.date, 'Depense', data.tipo + ' - ' + (data.desc || ''), 0, montant);
  return true;
}
function getDepenseDashboard(filters) {
  filters = filters || {};
  var ss = getSS();
  var sheet = ss.getSheetByName('Despesas');
  if (!sheet || sheet.getLastRow() <= 1) {
    return { total: 0, count: 0, average: 0, max: 0, maxCategory: '', todayTotal: 0, byCategory: [], byDay: [] };
  }

  var tz = Session.getScriptTimeZone() || 'Africa/Luanda';
  var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var fromDate = filters.from ? parseSheetDate(filters.from) : null;
  var toDate = filters.to ? parseSheetDate(filters.to) : null;
  if (toDate) toDate.setHours(23, 59, 59, 999);
  var categoryFilter = (filters.category || '').toString().trim().toLowerCase();
  var data = sheet.getDataRange().getValues();
  var total = 0;
  var count = 0;
  var max = 0;
  var maxCategory = '';
  var todayTotal = 0;
  var byCategory = {};
  var byDay = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var date = parseSheetDate(row[0]);
    if (!date) continue;
    if (fromDate && date < fromDate) continue;
    if (toDate && date > toDate) continue;

    var category = (row[1] || '').toString().trim();
    if (categoryFilter && category.toLowerCase() !== categoryFilter) continue;

    var amount = parseFloat(row[4]) || 0;
    total += amount;
    count++;
    if (amount > max) {
      max = amount;
      maxCategory = category;
    }

    var dayKey = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
    var dayLabel = Utilities.formatDate(date, tz, 'dd/MM/yyyy');
    if (dayKey === todayKey) todayTotal += amount;
    byCategory[category || 'Autre'] = (byCategory[category || 'Autre'] || 0) + amount;
    if (!byDay[dayKey]) byDay[dayKey] = { date: dayLabel, total: 0 };
    byDay[dayKey].total += amount;
  }

  var categories = Object.keys(byCategory).map(function(name) {
    return { category: name, total: byCategory[name] };
  }).sort(function(a, b) { return b.total - a.total; }).slice(0, 8);

  var days = Object.keys(byDay).sort().map(function(key) {
    return byDay[key];
  });
  if (days.length > 10) days = days.slice(days.length - 10);

  return {
    total: total,
    count: count,
    average: count ? total / count : 0,
    max: max,
    maxCategory: maxCategory,
    todayTotal: todayTotal,
    byCategory: categories,
    byDay: days
  };
}

function getHistoriqueDepenses(filters) {
  filters = filters || {};
  var ss = getSS();
  var sheet = ss.getSheetByName('Despesas');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var tz = Session.getScriptTimeZone() || 'Africa/Luanda';
  var fromDate = filters.from ? parseSheetDate(filters.from) : null;
  var toDate = filters.to ? parseSheetDate(filters.to) : null;
  if (toDate) toDate.setHours(23, 59, 59, 999);
  var categoryFilter = (filters.category || '').toString().trim().toLowerCase();
  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var date = parseSheetDate(row[0]);
    if (!date) continue;
    if (fromDate && date < fromDate) continue;
    if (toDate && date > toDate) continue;

    var category = (row[1] || '').toString().trim();
    if (categoryFilter && category.toLowerCase() !== categoryFilter) continue;

    rows.push({
      rawDate: date,
      date: Utilities.formatDate(date, tz, 'dd/MM/yyyy'),
      category: category,
      description: (row[2] || '').toString(),
      amount: parseFloat(row[4]) || 0
    });
  }

  rows.sort(function(a, b) {
    return b.rawDate - a.rawDate;
  });

  return rows.map(function(row) {
    return {
      date: row.date,
      category: row.category,
      description: row.description,
      amount: row.amount
    };
  });
}


// ============================================================
//  ENREGISTRER UN FOURNISSEUR
// ============================================================
function getComptabiliteData(filters) {
  filters = filters || {};

  var ss = getSS();
  var today = new Date();
  var tz = Session.getScriptTimeZone() || 'Africa/Luanda';

  // Periode du filtre
  var fromDate = filters.from
    ? parseSheetDate(filters.from)
    : new Date(today.getFullYear(), today.getMonth(), 1);

  var toDate = filters.to
    ? parseSheetDate(filters.to)
    : today;

  if (toDate) toDate.setHours(23, 59, 59, 999);

  var typeFilter = (filters.type || '').toString().toLowerCase();

  // Convertit une valeur en nombre fiable
  function num(value) {
    return parseFloat(value) || 0;
  }

  // Verifie si une date est dans la periode filtree
  function inPeriod(value) {
    var d = parseSheetDate(value);
    if (!d) return false;
    if (fromDate && d < fromDate) return false;
    if (toDate && d > toDate) return false;
    return true;
  }

  // Affiche une date propre
  function fmtDate(value) {
    var d = parseSheetDate(value);
    return d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : String(value || '');
  }

  // Ajoute une ligne dans le journal comptable
  // debit  = ce que l'entreprise recoit ou possede
  // credit = ce que l'entreprise paie, doit, ou sort
  function addJournal(list, date, type, desc, debit, credit, source) {
    if (!inPeriod(date)) return;

    var hay = (type + ' ' + source + ' ' + desc).toLowerCase();
    if (typeFilter && hay.indexOf(typeFilter) === -1) return;

    list.push({
      rawDate: parseSheetDate(date) || new Date(0),
      date: fmtDate(date),
      type: type,
      desc: desc || '',
      debit: Math.max(num(debit), 0),
      credit: Math.max(num(credit), 0),

      // Compatibilite avec ton ancien affichage Entrada/Saida
      entree: Math.max(num(debit), 0),
      sortie: Math.max(num(credit), 0),

      source: source || ''
    });
  }

  // =========================
  // 1. VENTES
  // =========================
  var journal = [];

  var ventesTotal = 0;
  var ventesCount = 0;
  var coutVentes = 0;
  var beneficeBrut = 0;

  var ventes = ss.getSheetByName('Vendas');

  if (ventes && ventes.getLastRow() > 1) {
    var vData = ventes.getDataRange().getValues();

    for (var i = 1; i < vData.length; i++) {
      var v = vData[i];
      if (!v[0] || !inPeriod(v[0])) continue;

      var totalVenda = num(v[8]);   // Montante total
      var lucro = num(v[10]);       // Lucro
      var custo = totalVenda - lucro;

      ventesTotal += totalVenda;
      beneficeBrut += lucro;
      coutVentes += custo;
      ventesCount++;

      addJournal(
        journal,
        v[0],
        'Venda',
        String(v[1] || ''),
        totalVenda,
        0,
        'Vendas'
      );
    }
  }

  // =========================
  // Variables principales
  // =========================
  var achatsTotal = 0;
  var depensesTotal = 0;
  var comprasCreditoPeriodo = 0;
  var pagamentosFornecedoresPeriodo = 0;
  var dettesFournisseurs = 0;
  var clientesAReceber = 0;
  var journal = [];

  // =========================
  // 2. ACHATS
  // Important:
  // Les achats augmentent le stock.
  // Ils ne sont PAS automatiquement une sortie de caisse.
  // La sortie reelle vient de Tesouraria ou Pagamento Fornecedor.
  // =========================
  var achat = ss.getSheetByName('Compra');

  if (achat && achat.getLastRow() > 1) {
    var aData = achat.getDataRange().getValues();

    for (var a = 1; a < aData.length; a++) {
      var ar = aData[a];
      if (!ar[0] || !inPeriod(ar[0])) continue;

      var amount = num(ar[5]); // Montante total
      achatsTotal += amount;
    }
  }

  // =========================
  // 3. DEPENSES
  // Les depenses diminuent le resultat.
  // La vraie sortie de caisse est deja dans Tesouraria.
  // =========================
  var dep = ss.getSheetByName('Despesas');

  if (dep && dep.getLastRow() > 1) {
    var dData = dep.getDataRange().getValues();

    for (var d = 1; d < dData.length; d++) {
      var dr = dData[d];
      if (!dr[0] || !inPeriod(dr[0])) continue;

      var depAmount = num(dr[4]); // Saida
      depensesTotal += depAmount;
    }
  }

  // =========================
  // 4. TRESOURARIA
  // C'est ici qu'on lit les vraies entrees/sorties d'argent.
  // =========================
  var tresoBalance = 0;
  var fluxoEntradas = 0;
  var fluxoSaidas = 0;

  var treso = ss.getSheetByName('Tesouraria');

  if (treso && treso.getLastRow() > 1) {
    tresoBalance = num(treso.getRange(treso.getLastRow(), 6).getValue());

    var tData = treso.getDataRange().getValues();

    for (var t = 1; t < tData.length; t++) {
      var tr = tData[t];
      if (!tr[0] || !inPeriod(tr[0])) continue;

      var entrada = num(tr[3]);
      var saida = num(tr[4]);

      fluxoEntradas += entrada;
      fluxoSaidas += saida;

      addJournal(
        journal,
        tr[0],
        String(tr[1] || 'Movimento'),
        String(tr[2] || ''),
        entrada,
        saida,
        'Tesouraria'
      );
    }
  }

  // =========================
  // 5. DETTES FOURNISSEURS
  // On lit la feuille Divida Fornecedores.
  // La colonne G contient le solde restant.
  // =========================
  var dividaForn = ss.getSheetByName('Divida Fornecedores');

  if (dividaForn && dividaForn.getLastRow() > 1) {
    // Si ta fonction existe, elle recalcule les soldes avant lecture
    if (typeof recomputeSupplierDebtSheet_ === 'function') {
      recomputeSupplierDebtSheet_(dividaForn);
    }

    var dfData = dividaForn.getDataRange().getValues();
    var supplierBalances = {};

    for (var df = 1; df < dfData.length; df++) {
      var rowDf = dfData[df];

      var dateDf = rowDf[0];
      var supplier = String(rowDf[1] || '').trim();
      var tipoDf = String(rowDf[2] || '');
      var designationDf = String(rowDf[3] || '');
      var compraCredito = num(rowDf[4]);
      var pagamentoForn = num(rowDf[5]);
      var saldoForn = num(rowDf[6]);

      if (!supplier) continue;

      supplierBalances[supplier.toLowerCase()] = saldoForn;

      if (inPeriod(dateDf)) {
        if (compraCredito > 0) {
          comprasCreditoPeriodo += compraCredito;

          addJournal(
            journal,
            dateDf,
            'Compra a credito',
            supplier + (designationDf ? ' - ' + designationDf : ''),
            compraCredito,
            compraCredito,
            'Divida Fornecedores'
          );
        }

        if (pagamentoForn > 0) {
          pagamentosFornecedoresPeriodo += pagamentoForn;
        }
      }
    }

    Object.keys(supplierBalances).forEach(function(key) {
      dettesFournisseurs += supplierBalances[key];
    });
  }

  // =========================
  // 6. DETTES CLIENTS
  // Ce sont les montants que les clients doivent encore payer.
  // Colonne G = restant.
  // =========================
  var dividaClientes = ss.getSheetByName('Divida Clientes');

  if (dividaClientes && dividaClientes.getLastRow() > 1) {
    var dcData = dividaClientes.getDataRange().getValues();
    var clientBalances = {};

    for (var dc = 1; dc < dcData.length; dc++) {
      var rowDc = dcData[dc];

      var client = String(rowDc[1] || '').trim();
      var restantClient = Math.max(num(rowDc[6]), 0);

      if (!client) continue;

      clientBalances[client.toLowerCase()] = restantClient;
    }

    Object.keys(clientBalances).forEach(function(key) {
      clientesAReceber += clientBalances[key];
    });
  }

  // =========================
  // 7. STOCK
  // Valeur estimee du stock restant:
  // quantite restante x prix d'achat
  // =========================
  var stockValeur = 0;
  var inv = ss.getSheetByName('Inventario');

  if (inv && inv.getLastRow() > 1) {
    var invHeaders = ensureProductMetadataColumns_(inv);
    var invData = inv.getDataRange().getValues();
    var invMap = getInventoryColumnMap_(invHeaders);

    for (var s = 1; s < invData.length; s++) {
      var qty = num(invData[s][invMap.stock]);

      var cost = invMap.purchasePrice >= 0
        ? num(invData[s][invMap.purchasePrice]) || num(invData[s][invMap.price])
        : num(invData[s][invMap.price]);

      stockValeur += qty * cost;
    }
  }

  // =========================
  // 8. RESULTAT ET BILAN
  // =========================
  var resultatNet = beneficeBrut - depensesTotal;

  var actifTotal = tresoBalance + stockValeur + clientesAReceber;
  var passifTotal = dettesFournisseurs;
  var capitaisProprios = actifTotal - passifTotal;

  journal.sort(function(left, right) {
    return right.rawDate - left.rawDate;
  });

  journal = journal.map(function(row) {
    delete row.rawDate;
    return row;
  });

  return {
    period: {
      from: Utilities.formatDate(fromDate, tz, 'dd/MM/yyyy'),
      to: Utilities.formatDate(toDate, tz, 'dd/MM/yyyy')
    },

    resume: {
      ventes: ventesTotal,
      ventesCount: ventesCount,
      coutVentes: coutVentes,
      beneficeBrut: beneficeBrut,

      achats: achatsTotal,
      comprasCredito: comprasCreditoPeriodo,
      comprasPagas: Math.max(achatsTotal - comprasCreditoPeriodo, 0),
      pagamentosFornecedores: pagamentosFornecedoresPeriodo,

      depenses: depensesTotal,
      resultatNet: resultatNet,
      resultadoOperacional: resultatNet,

      marge: ventesTotal ? (beneficeBrut / ventesTotal) * 100 : 0
    },

    bilan: {
      tresorerie: tresoBalance,
      stock: stockValeur,
      clientesAReceber: clientesAReceber,

      // Les deux noms sont gardes pour compatibilite avec ton affichage actuel
      dettesFournisseurs: dettesFournisseurs,
      dividasFournisseurs: dettesFournisseurs,

      actifSimplifie: actifTotal,
      ativoTotal: actifTotal,

      passivo: passifTotal,
      capitaisProprios: capitaisProprios,

      resultatNet: resultatNet
    },

    flux: {
      entrees: fluxoEntradas,
      sorties: fluxoSaidas,
      saldo: fluxoEntradas - fluxoSaidas
    },

    journal: journal.slice(0, parseInt(filters.limit, 10) || 200)
  };
}

function registarFornecedor(data) {
  var SS = getSS();
  var sheet = SS.getSheetByName('Fornecedores');
  if (!sheet) throw new Error('Fornecedores sheet not found');
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 4).setValues([[
    data.nome, data.tel, data.pais, data.nota
  ]]);
  return true;
}

// ============================================================
//  REINITIALISER TOUTES LES DONNEES
//  ATTENTION : supprime TOUT (Achat, Ventes, Transferts, Clientes)
//  Cette action est IRREVERSIBLE
//  Utilisation : Menu Azul Gestão > RESET ALL DATA
// ============================================================
function reinicializarTudo(data) {
  var SS = getSS();

  [
    'Compra',
    'Vendas',
    'Transferencias',
    'Clientes',
    'Divida Fornecedores',
    'Divida Clientes',
    'Immobilizaçao',
    'Revendedores',
    'Tesouraria',
    'Despesas',
    'Fornecedores'
  ].forEach(function(name) {
    var sheet = SS.getSheetByName(name);

    if (sheet && sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  });

  var stockMode = data && data.stockMode === 'armazem' ? 'armazem' : 'boutique';

  criarInventaire(stockMode);

  return {
    ok: true,
    message: 'Dados apagados. Inventario atualizado no modo: ' +
      (stockMode === 'boutique' ? 'Apenas loja' : 'Loja + armazem')
  };
}


// function doGet() {
//   return HtmlService.createHtmlOutputFromFile('POS_Core');
// }
