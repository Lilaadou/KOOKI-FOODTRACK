// ================================================================
//  KOOKI — Google Apps Script
//  1. Coller dans Apps Script (Extensions > Apps Script)
//  2. Sauvegarder (Ctrl+S)
//  3. Selectionner initSheets dans le menu deroulant > Executer
//  4. Autoriser les permissions
//  5. Deployer > Nouvelle version > Application Web > Tout le monde
// ================================================================

// Noms des onglets Google Sheets
var NOM_STOCK          = 'Stock';
var NOM_PRODUCTION     = 'Production';
var NOM_PREMIX         = 'Premix';
var NOM_RECEPTION      = 'Reception';
var NOM_FACTURES       = 'Factures';
var NOM_SORTIES        = 'Sorties';
var NOM_MANQUANTS      = 'Manquants';
var NOM_INVENTAIRE     = 'Inventaire_Premix';
var NOM_CORRESPONDANCES = 'Correspondances';

// Ligne ou commencent vos donnees de production
var PROD_LIGNE_DEBUT = 2;

// ================================================================
//  INIT — creation des onglets
// ================================================================
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var onglets = [
    { nom: NOM_STOCK,      entetes: ['Code Lot', 'Saveur', 'Big 165g', 'Small 80g', 'Mise a jour'],                                          ligne: 1 },
    { nom: NOM_PRODUCTION, entetes: ['Date', 'Code Lot', 'Saveur', 'Code Premix', 'Producteur', 'Big 165g', 'Small 80g', 'Factures'],        ligne: PROD_LIGNE_DEBUT },
    { nom: NOM_PREMIX,     entetes: ['Date', 'Code Premix', 'Saveur', 'Sac', 'Producteur', 'Factures', 'Ingredients manquants', 'Resolu'],    ligne: 1 },
    // Fournisseurs : Foodflow, Bruyerre, Colruyt, Hygiena, Makady, De Notten Shop, Autre
    { nom: NOM_RECEPTION,  entetes: ['Date', 'N Facture', 'Fournisseur', 'Ajoute par'],                                                       ligne: 1 },
    { nom: NOM_FACTURES,   entetes: ['Date', 'N Facture', 'Fournisseur', 'Code Production'],                                                  ligne: 1 },
    { nom: NOM_SORTIES,    entetes: ['Date', 'Saveur', 'Destination', 'Code Lot', 'Big 165g', 'Small 80g', 'Boites', 'Livre par'],            ligne: 1 },
    { nom: NOM_MANQUANTS,  entetes: ['Date', 'Code Premix', 'Saveur', 'Sac', 'Ingredients manquants', 'Producteur', 'Resolu', 'Date resolu'], ligne: 1 },
    { nom: NOM_INVENTAIRE,     entetes: ['Saveur', 'Sac1 Statut', 'Sac1 Code', 'Sac1 Date', 'Sac2 Statut', 'Sac2 Code', 'Sac2 Date'],  ligne: 1 },
    { nom: NOM_CORRESPONDANCES, entetes: ['Date', 'N BL', 'N Facture Odoo', 'Fournisseur'],                                              ligne: 1 },
  ];

  var crees = [];

  for (var i = 0; i < onglets.length; i++) {
    var cfg = onglets[i];
    var sheet = ss.getSheetByName(cfg.nom);

    if (!sheet) {
      sheet = ss.insertSheet(cfg.nom);
      crees.push(cfg.nom);
    }

    var ligneEntete = (cfg.ligne === 1) ? 1 : cfg.ligne - 1;
    var cellule = sheet.getRange(ligneEntete, 1).getValue();

    if (cellule === '' || cellule === null) {
      sheet.getRange(ligneEntete, 1, 1, cfg.entetes.length).setValues([cfg.entetes]);
    }
  }

  // Remplir l'inventaire premix si vide
  var invSheet = ss.getSheetByName(NOM_INVENTAIRE);
  if (invSheet && invSheet.getLastRow() <= 1) {
    var saveurs = ['SUZZY','JACKY','BENNY','DAVY','MOLLY','WALLY','CHARLY','LENI','EMMY','WOODY','DANNY'];
    var lignes = [];
    for (var s = 0; s < saveurs.length; s++) {
      lignes.push([saveurs[s], 'Vide', '', '', 'Vide', '', '']);
    }
    invSheet.getRange(2, 1, lignes.length, 7).setValues(lignes);
  }

  var msg = crees.length > 0
    ? 'Onglets crees : ' + crees.join(', ')
    : 'Tous les onglets existent deja.';
  SpreadsheetApp.getUi().alert(msg);
}

// ================================================================
//  GET — lecture des donnees
// ================================================================
function doGet(e) {
  try {
    var action = e.parameter.action;
    var result;

    if      (action === 'getStock')              result = lireStock();
    else if (action === 'getProduction')         result = lireProduction();
    else if (action === 'getRecentInvoices')     result = lireReception();
    else if (action === 'getHistoriqueSorties')  result = lireSorties();
    else if (action === 'getPremix')             result = lirePremix();
    else if (action === 'getFactures')           result = lireFactures();
    else if (action === 'getManquants')          result = lireManquants();
    else if (action === 'getInventairePremix')   result = lireInventaire();
    else if (action === 'getCorrespondances')    result = lireCorrespondances();
    else if (action === 'getStockIngredients')   result = getStockIngredients_();
    else if (action === 'getSacsDispo')          result = getSacsDispo_();
    else if (action === 'getLivraisonsIngredients') result = getLivraisonsIngredients_();
    else result = { error: 'Action inconnue : ' + action };

    return reponseJSON(result);
  } catch (err) {
    return reponseJSON({ error: err.toString() });
  }
}

function lireStock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_STOCK);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      code:  r['Code Lot']  || '',
      saveur:r['Saveur']    || '',
      big:   Number(r['Big 165g'])   || 0,
      small: Number(r['Small 80g'])  || 0
    });
  }
  return result;
}

function lireProduction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_PRODUCTION);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, PROD_LIGNE_DEBUT);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var facsStr = String(r['Factures'] || '');
    result.push({
      date:  r['Date']        || '',
      code:  r['Code Lot']    || '',
      saveur:r['Saveur']      || '',
      pmx:   r['Code Premix'] || '',
      nom:   r['Producteur']  || '',
      big:   Number(r['Big 165g'])   || 0,
      small: Number(r['Small 80g'])  || 0,
      bls:   facsStr ? facsStr.split(',').map(function(s){ return s.trim(); }) : []
    });
  }
  return result;
}

function lireReception() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_RECEPTION);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    result.push({
      date:        r['Date']        || '',
      numero:      r['N Facture']   || '',
      fournisseur: r['Fournisseur'] || '',
      nom:         r['Ajoute par']  || ''
    });
  }
  return result;
}

function lireSorties() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_SORTIES);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    result.push({
      date:        r['Date']        || '',
      saveur:      r['Saveur']      || '',
      destination: r['Destination'] || '',
      lot:         r['Code Lot']    || '',
      big:         Number(r['Big 165g'])   || 0,
      small:       Number(r['Small 80g'])  || 0,
      boites:      r['Boites']      || '',
      producteur:  r['Livre par']   || ''
    });
  }
  return result;
}

function lirePremix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_PREMIX);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    var facsStr = String(r['Factures'] || '');
    var mqStr   = String(r['Ingredients manquants'] || '');
    result.push({
      date:             r['Date']        || '',
      codePremix:       r['Code Premix'] || '',
      saveur:           r['Saveur']      || '',
      sac:              r['Sac']         || '',
      nom:              r['Producteur']  || '',
      factures:         facsStr ? facsStr.split(',').map(function(s){ return s.trim(); }) : [],
      manquants:        mqStr   ? mqStr.split(',').map(function(s){ return s.trim(); })   : [],
      manquantsResolus: r['Resolu'] === 'Oui' || r['Resolu'] === true
    });
  }
  return result;
}

function lireFactures() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_FACTURES);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      ajouteLe:       r['Date']            || '',
      numero:         r['N Facture']        || '',
      fournisseur:    r['Fournisseur']      || '',
      codeProduction: r['Code Production']  || ''
    });
  }
  return result;
}

function lireManquants() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_MANQUANTS);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r['Resolu'] === 'Oui' || r['Resolu'] === true) continue;
    var mqStr = String(r['Ingredients manquants'] || '');
    var ings  = mqStr ? mqStr.split(',').map(function(s){ return { nom: s.trim() }; }) : [];
    result.push({
      date:        r['Date']        || '',
      code:        r['Code Premix'] || '',
      saveur:      r['Saveur']      || '',
      sac:         r['Sac']         || '',
      ingredients: ings,
      producteur:  r['Producteur']  || ''
    });
  }
  return result;
}

function lireInventaire() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_INVENTAIRE);
  if (!sheet) return {};
  var rows = lireOnglet(sheet, 2);
  var result = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var sav = r['Saveur'];
    if (!sav) continue;
    result[sav] = {
      1: { statut: r['Sac1 Statut'] === 'Disponible' ? 'dispo' : 'vide', codePremix: r['Sac1 Code'] || '', date: r['Sac1 Date'] || '' },
      2: { statut: r['Sac2 Statut'] === 'Disponible' ? 'dispo' : 'vide', codePremix: r['Sac2 Code'] || '', date: r['Sac2 Date'] || '' }
    };
  }
  return result;
}

function lireCorrespondances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_CORRESPONDANCES);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    result.push({
      date:        r['Date']            || '',
      bl:          r['N BL']            || '',
      facture:     r['N Facture Odoo']  || '',
      fournisseur: r['Fournisseur']     || ''
    });
  }
  return result;
}

// ================================================================
//  POST — ecriture des donnees
// ================================================================
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;

    if      (action === 'production')     result = sauverProduction(data);
    else if (action === 'premix')         result = sauverPremix(data);
    else if (action === 'reception')      result = sauverReception(data);
    else if (action === 'sortie')         result = sauverSortie(data);
    else if (action === 'correspondance')       result = sauverCorrespondance(data);
    else if (action === 'livraisonIngredients') result = livraisonIngredients_(data);
    else if (action === 'deduireIngredients')   result = deduireIngredients_(data);
    else if (action === 'corrigerStock')         result = corrigerStock_(data);
    else result = { success: false, error: 'Action inconnue : ' + action };

    return reponseJSON(result);
  } catch (err) {
    return reponseJSON({ success: false, error: err.toString() });
  }
}

function sauverProduction(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_PRODUCTION);
  if (!sheet) return { success: false, error: 'Onglet Production introuvable' };

  var facs = Array.isArray(d.factures) ? d.factures.join(', ') : (d.factures || '');
  sheet.appendRow([maintenant(), d.code, d.saveur, d.pmx, d.nom, Number(d.big) || 0, Number(d.small) || 0, facs]);

  majStock(ss, d.code, d.saveur, Number(d.big) || 0, Number(d.small) || 0);
  marquerSacUtilise(ss, d.pmx);

  return { success: true };
}

function sauverPremix(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_PREMIX);
  if (!sheet) return { success: false, error: 'Onglet Premix introuvable' };

  var facs = Array.isArray(d.factures) ? d.factures.join(', ') : '';
  var mqs  = Array.isArray(d.manquants) ? d.manquants.join(', ') : '';
  sheet.appendRow([maintenant(), d.codePremix, d.saveur, d.sac, d.nom, facs, mqs, 'Non']);

  marquerSacDisponible(ss, d.saveur, String(d.sac), d.codePremix);

  if (mqs) {
    var mqSheet = ss.getSheetByName(NOM_MANQUANTS);
    if (mqSheet) mqSheet.appendRow([maintenant(), d.codePremix, d.saveur, d.sac, mqs, d.nom, 'Non', '']);
  }

  return { success: true };
}

function sauverReception(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_RECEPTION);
  if (!sheet) return { success: false, error: 'Onglet Reception introuvable' };

  sheet.appendRow([maintenant(), d.numero, d.fournisseur, d.nom || '']);
  return { success: true };
}

function sauverCorrespondance(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_CORRESPONDANCES);
  if (!sheet) return { success: false, error: 'Onglet Correspondances introuvable' };

  var dernLigne = sheet.getLastRow();
  var data = dernLigne > 1 ? sheet.getRange(2, 2, dernLigne - 1, 1).getValues() : [];

  // Mettre à jour si le BL existe déjà
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === d.bl) {
      sheet.getRange(i + 2, 3).setValue(d.facture);
      sheet.getRange(i + 2, 4).setValue(d.fournisseur || '');
      sheet.getRange(i + 2, 1).setValue(maintenant());
      return { success: true, updated: true };
    }
  }
  // Sinon nouvelle ligne
  sheet.appendRow([maintenant(), d.bl, d.facture, d.fournisseur || '']);
  return { success: true, updated: false };
}

function sauverSortie(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_SORTIES);
  if (!sheet) return { success: false, error: 'Onglet Sorties introuvable' };

  var lots = Array.isArray(d.lots) ? d.lots : [];
  for (var i = 0; i < lots.length; i++) {
    var lot = lots[i];
    sheet.appendRow([maintenant(), d.saveur || '', d.destination || '', lot.code || '', Number(lot.big) || 0, Number(lot.small) || 0, '', '']);
    diminuerStock(ss, lot.code, Number(lot.big) || 0, Number(lot.small) || 0);
  }

  return { success: true };
}

// ================================================================
//  STOCK — mise a jour automatique
// ================================================================
function majStock(ss, codeLot, saveur, addBig, addSmall) {
  var sheet = ss.getSheetByName(NOM_STOCK);
  if (!sheet) return;
  var dernLigne = sheet.getLastRow();
  if (dernLigne < 2) {
    sheet.appendRow([codeLot, saveur, addBig, addSmall, maintenant()]);
    return;
  }
  var data = sheet.getRange(2, 1, dernLigne - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === codeLot) {
      sheet.getRange(i + 2, 3).setValue(data[i][2] + addBig);
      sheet.getRange(i + 2, 4).setValue(data[i][3] + addSmall);
      sheet.getRange(i + 2, 5).setValue(maintenant());
      return;
    }
  }
  sheet.appendRow([codeLot, saveur, addBig, addSmall, maintenant()]);
}

function diminuerStock(ss, codeLot, remBig, remSmall) {
  var sheet = ss.getSheetByName(NOM_STOCK);
  if (!sheet) return;
  var dernLigne = sheet.getLastRow();
  if (dernLigne < 2) return;
  var data = sheet.getRange(2, 1, dernLigne - 1, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === codeLot) {
      sheet.getRange(i + 2, 3).setValue(Math.max(0, (data[i][2] || 0) - remBig));
      sheet.getRange(i + 2, 4).setValue(Math.max(0, (data[i][3] || 0) - remSmall));
      sheet.getRange(i + 2, 5).setValue(maintenant());
      return;
    }
  }
}

// ================================================================
//  INVENTAIRE PREMIX — mise a jour automatique
// ================================================================
function marquerSacDisponible(ss, saveur, sac, codePremix) {
  var sheet = ss.getSheetByName(NOM_INVENTAIRE);
  if (!sheet) return;
  var dernLigne = sheet.getLastRow();
  if (dernLigne < 2) return;
  var data = sheet.getRange(2, 1, dernLigne - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === saveur) {
      var col = sac === '1' ? 2 : 5;
      sheet.getRange(i + 2, col).setValue('Disponible');
      sheet.getRange(i + 2, col + 1).setValue(codePremix);
      sheet.getRange(i + 2, col + 2).setValue(maintenant());
      return;
    }
  }
}

function marquerSacUtilise(ss, codePremix) {
  var sheet = ss.getSheetByName(NOM_INVENTAIRE);
  if (!sheet) return;
  var dernLigne = sheet.getLastRow();
  if (dernLigne < 2) return;
  var data = sheet.getRange(2, 1, dernLigne - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][2] === codePremix && data[i][1] === 'Disponible') {
      sheet.getRange(i + 2, 2).setValue('Vide');
      sheet.getRange(i + 2, 3).setValue('');
      sheet.getRange(i + 2, 4).setValue(maintenant());
      return;
    }
    if (data[i][5] === codePremix && data[i][4] === 'Disponible') {
      sheet.getRange(i + 2, 5).setValue('Vide');
      sheet.getRange(i + 2, 6).setValue('');
      sheet.getRange(i + 2, 7).setValue(maintenant());
      return;
    }
  }
}

// ================================================================
//  HELPERS
// ================================================================
function reponseJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function maintenant() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function lireOnglet(sheet, ligneDebut) {
  var dernLigne = sheet.getLastRow();
  if (dernLigne < ligneDebut) return [];
  var entetes = sheet.getRange(ligneDebut - 1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lignes  = sheet.getRange(ligneDebut, 1, dernLigne - ligneDebut + 1, entetes.length).getValues();
  var result  = [];
  for (var i = 0; i < lignes.length; i++) {
    var vide = true;
    for (var j = 0; j < lignes[i].length; j++) {
      if (lignes[i][j] !== '' && lignes[i][j] !== null) { vide = false; break; }
    }
    if (vide) continue;
    var obj = {};
    for (var k = 0; k < entetes.length; k++) {
      obj[entetes[k]] = lignes[i][k];
    }
    result.push(obj);
  }
  return result;
}

// ================================================================
//  MENU dans Google Sheets
// ================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('KOOKI App')
    .addItem('Initialiser les onglets', 'initSheets')
    .addItem('Initialiser les onglets ingredients', 'initIngredientSheets')
    .addSeparator()
    .addItem('Voir manquants non resolus', 'afficherManquants')
    .addItem('Voir stock actuel', 'afficherStock')
    .addToUi();
}

function afficherManquants() {
  var data = lireManquants();
  if (!data.length) {
    SpreadsheetApp.getUi().alert('Aucun ingredient manquant non resolu.');
    return;
  }
  var msg = '';
  for (var i = 0; i < data.length; i++) {
    var m = data[i];
    var ings = m.ingredients.map(function(x){ return x.nom; }).join(', ');
    msg += '- ' + m.code + ' (' + m.saveur + ' Sac ' + m.sac + ') : ' + ings + '\n';
  }
  SpreadsheetApp.getUi().alert(data.length + ' premix incomplet(s) :\n\n' + msg);
}

function afficherStock() {
  var data = lireStock();
  var totalBig = 0, totalSmall = 0;
  var msg = '';
  for (var i = 0; i < data.length; i++) {
    totalBig   += data[i].big;
    totalSmall += data[i].small;
    msg += '- ' + data[i].code + ' (' + data[i].saveur + ') : ' + data[i].big + ' Big / ' + data[i].small + ' Small\n';
  }
  SpreadsheetApp.getUi().alert('Stock total : ' + totalBig + ' Big / ' + totalSmall + ' Small\n\n' + msg);
}

// ================================================================
//  STOCK INGRÉDIENTS — noms des nouvelles feuilles
// ================================================================
var NOM_RECETTES           = 'Recettes';          // protegee, cachee
var NOM_STOCK_ING          = 'Stock_Ingredients';
var NOM_LIVRAISONS_ING     = 'Livraisons_Ingredients';

var INGREDIENTS = [
  'Sucre brun','Sucre blanc','Levure chimique','Bicarbonate','Maizena','Sel',
  'Chocolat blanc','Chocolat noir','Chocolat lait','Caramel','Vegan','Noix',
  'Poudre amande','Poudre cacao','Pistache','Pistache poudre','Coconut','Cacahuetes'
];

var SACS = [
  'SUZZY 1','SUZZY 2','WOODY 1','WOODY 2','EMMY 1','EMMY 2',
  'BENNY 1','BENNY 2','JACKY 1','JACKY 2','MOLLY 1','MOLLY 2',
  'CHARLY 1','CHARLY 2','DAVY 1','DAVY 2','WALLY 1','WALLY 2',
  'LENI 1','LENI 2','DANNY 1','DANNY 2'
];

// ================================================================
//  INIT FEUILLES INGREDIENTS
// ================================================================
// Appeler initIngredientSheets() depuis le menu KOOKI App
function initIngredientSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Feuille Stock_Ingredients
  var stockSheet = ss.getSheetByName(NOM_STOCK_ING);
  if (!stockSheet) {
    stockSheet = ss.insertSheet(NOM_STOCK_ING);
    stockSheet.getRange(1, 1, 1, 3).setValues([['Ingredient', 'Stock (g)', 'Derniere MAJ']]);
    var rows = INGREDIENTS.map(function(ing) { return [ing, 0, '']; });
    stockSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  // 2. Feuille Livraisons_Ingredients
  var livSheet = ss.getSheetByName(NOM_LIVRAISONS_ING);
  if (!livSheet) {
    livSheet = ss.insertSheet(NOM_LIVRAISONS_ING);
    var hLiv = ['Date', 'N BL', 'Fournisseur'].concat(INGREDIENTS);
    livSheet.getRange(1, 1, 1, hLiv.length).setValues([hLiv]);
  }

  // 3. Feuille Recettes (22 sacs x 18 ingredients, tout a 0)
  var recSheet = ss.getSheetByName(NOM_RECETTES);
  if (!recSheet) {
    recSheet = ss.insertSheet(NOM_RECETTES);
    var hRec = ['Sac'].concat(INGREDIENTS);
    recSheet.getRange(1, 1, 1, hRec.length).setValues([hRec]);
    var recRows = SACS.map(function(sac) {
      return [sac].concat(INGREDIENTS.map(function() { return 0; }));
    });
    recSheet.getRange(2, 1, recRows.length, hRec.length).setValues(recRows);

    // Proteger la feuille Recettes
    var protection = recSheet.protect().setDescription('Recettes - Protege');
    protection.setUnprotectedRanges([]);
    // Retirer tous les editeurs sauf le proprietaire
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    protection.removeEditors(protection.getEditors().filter(function(e) { return e.getEmail() !== me.getEmail(); }));
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    // Cacher la feuille
    recSheet.hideSheet();
  }

  SpreadsheetApp.getUi().alert('Feuilles ingredients initialisees. La feuille Recettes est protegee et cachee.');
}

// ================================================================
//  GET — stock ingredients et disponibilite sacs
// ================================================================
function getStockIngredients_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_STOCK_ING);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  return rows.map(function(r) {
    return { nom: r['Ingredient'] || '', stock: Number(r['Stock (g)']) || 0, maj: r['Derniere MAJ'] || '' };
  });
}

function getSacsDispo_() {
  // Calcule cote serveur : recettes jamais exposees au client
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var recSheet = ss.getSheetByName(NOM_RECETTES);
  var stSheet  = ss.getSheetByName(NOM_STOCK_ING);
  if (!recSheet || !stSheet) return SACS.map(function(s) { return { sac: s, ok: false, manquants: [] }; });

  // Lire recettes (lignes protegees)
  var recRows = lireOnglet(recSheet, 2);
  var stRows  = lireOnglet(stSheet,  2);

  // Index du stock par ingredient
  var stockMap = {};
  stRows.forEach(function(r) { stockMap[r['Ingredient']] = Number(r['Stock (g)']) || 0; });

  // Pour chaque sac, verifier si le stock est suffisant
  return recRows.map(function(r) {
    var sacNom    = r['Sac'];
    var manquants = [];
    INGREDIENTS.forEach(function(ing) {
      var besoin = Number(r[ing]) || 0;
      var dispo  = stockMap[ing]  || 0;
      if (besoin > 0 && dispo < besoin) {
        manquants.push({ nom: ing, deficit: besoin - dispo });
        // NE PAS retourner "besoin" pour cacher la recette
      }
    });
    return { sac: sacNom, ok: manquants.length === 0, manquants: manquants };
  });
}

// ================================================================
//  POST — livraison + deduction ingredients
// ================================================================
function livraisonIngredients_(d) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var stSheet  = ss.getSheetByName(NOM_STOCK_ING);
  var livSheet = ss.getSheetByName(NOM_LIVRAISONS_ING);
  if (!stSheet) return { success: false, error: 'Feuille Stock_Ingredients introuvable' };

  // Mettre a jour le stock
  var dernLigne = stSheet.getLastRow();
  var data = dernLigne > 1 ? stSheet.getRange(2, 1, dernLigne - 1, 3).getValues() : [];
  var stockMisAJour = {};

  for (var i = 0; i < data.length; i++) {
    var ing = data[i][0];
    var qte = Number(d.livraison && d.livraison[ing] !== undefined ? d.livraison[ing] : 0);
    if (qte > 0) {
      var newStock = (Number(data[i][1]) || 0) + qte;
      stSheet.getRange(i + 2, 2).setValue(newStock);
      stSheet.getRange(i + 2, 3).setValue(maintenant());
      stockMisAJour[ing] = newStock;
    }
  }

  // Enregistrer dans le journal des livraisons
  if (livSheet) {
    var row = [maintenant(), d.bl || '', d.fournisseur || ''];
    INGREDIENTS.forEach(function(ing) { row.push(Number(d.livraison && d.livraison[ing] || 0)); });
    livSheet.appendRow(row);
  }

  return { success: true, stockMisAJour: stockMisAJour };
}

function deduireIngredients_(d) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var recSheet = ss.getSheetByName(NOM_RECETTES);
  var stSheet  = ss.getSheetByName(NOM_STOCK_ING);
  if (!recSheet || !stSheet) return { success: false, error: 'Feuilles manquantes' };

  var sacNom = d.sacNom;
  // Trouver la ligne de la recette
  var recRows  = lireOnglet(recSheet, 2);
  var recette  = recRows.find(function(r) { return r['Sac'] === sacNom; });
  if (!recette) return { success: false, error: 'Sac ' + sacNom + ' introuvable dans les recettes' };

  // Lire stock actuel
  var dernLigne = stSheet.getLastRow();
  var data = dernLigne > 1 ? stSheet.getRange(2, 1, dernLigne - 1, 3).getValues() : [];
  var stockMisAJour = {};

  for (var i = 0; i < data.length; i++) {
    var ing    = data[i][0];
    var besoin = Number(recette[ing]) || 0;
    if (besoin > 0) {
      var newStock = Math.max(0, (Number(data[i][1]) || 0) - besoin);
      stSheet.getRange(i + 2, 2).setValue(newStock);
      stSheet.getRange(i + 2, 3).setValue(maintenant());
      stockMisAJour[ing] = newStock;
    }
  }

  return { success: true, stockMisAJour: stockMisAJour };
}

// ================================================================
//  Ajouter dans doGet et doPost
// ================================================================
// IMPORTANT : Remplacer les fonctions doGet et doPost existantes
// par ces nouvelles versions qui incluent les nouvelles actions

function doGet_v2(e) {
  try {
    var action = e.parameter.action;
    var result;
    if      (action === 'getStock')              result = lireStock();
    else if (action === 'getProduction')         result = lireProduction();
    else if (action === 'getRecentInvoices')     result = lireReception();
    else if (action === 'getHistoriqueSorties')  result = lireSorties();
    else if (action === 'getPremix')             result = lirePremix();
    else if (action === 'getFactures')           result = lireFactures();
    else if (action === 'getManquants')          result = lireManquants();
    else if (action === 'getInventairePremix')   result = lireInventaire();
    else if (action === 'getCorrespondances')    result = lireCorrespondances();
    else if (action === 'getStockIngredients')   result = getStockIngredients_();
    else if (action === 'getSacsDispo')          result = getSacsDispo_();
    else if (action === 'getLivraisonsIngredients') result = getLivraisonsIngredients_();
    else result = { error: 'Action inconnue : ' + action };
    return reponseJSON(result);
  } catch (err) {
    return reponseJSON({ error: err.toString() });
  }
}

function doPost_v2(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;
    if      (action === 'production')           result = sauverProduction(data);
    else if (action === 'premix')               result = sauverPremix(data);
    else if (action === 'reception')            result = sauverReception(data);
    else if (action === 'sortie')               result = sauverSortie(data);
    else if (action === 'correspondance')       result = sauverCorrespondance(data);
    else if (action === 'livraisonIngredients') result = livraisonIngredients_(data);
    else if (action === 'deduireIngredients')   result = deduireIngredients_(data);
    else if (action === 'corrigerStock')         result = corrigerStock_(data);
    else result = { success: false, error: 'Action inconnue : ' + action };
    return reponseJSON(result);
  } catch (err) {
    return reponseJSON({ success: false, error: err.toString() });
  }
}

// ================================================================
//  NOUVELLES ACTIONS — historique livraisons + correction stock
// ================================================================

function getLivraisonsIngredients_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOM_LIVRAISONS_ING);
  if (!sheet) return [];
  var rows = lireOnglet(sheet, 2);
  // Retourner seulement date, BL, fournisseur (pas les quantites)
  return rows.map(function(r) {
    return { date: r['Date'] || '', bl: r['N BL'] || '', fournisseur: r['Fournisseur'] || '' };
  }).reverse(); // plus recentes en premier
}

function corrigerStock_(d) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var stSheet = ss.getSheetByName(NOM_STOCK_ING);
  if (!stSheet) return { success: false, error: 'Feuille Stock_Ingredients introuvable' };

  var dernLigne = stSheet.getLastRow();
  var data = dernLigne > 1 ? stSheet.getRange(2, 1, dernLigne - 1, 3).getValues() : [];

  for (var i = 0; i < data.length; i++) {
    var ing = data[i][0];
    if (d.corrections && d.corrections[ing] !== undefined) {
      stSheet.getRange(i + 2, 2).setValue(Number(d.corrections[ing]));
      stSheet.getRange(i + 2, 3).setValue(maintenant() + ' (correction manuelle)');
    }
  }
  return { success: true };
}
