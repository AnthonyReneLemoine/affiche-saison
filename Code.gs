function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Créateur d\'affiches A3');
}


function previewPoster(payload) {
  var data = JSON.parse(payload);
  data.category = inferCategory_(data);
  data = normalizePosterData(data);

  var html = buildPosterHtml(data);

  // Ajustements CSS pour l'aperçu (pas d'impact sur le PDF)
  html = html.replace('</head>', '<style>html,body{margin:0;padding:0;background:#e9e9e9;} .poster{transform:scale(0.38);transform-origin:top left;box-shadow:0 12px 40px rgba(0,0,0,.25);} body{padding:16px;} </style></head>');

  return { ok: true, html: html };
}

function createPoster(payload) {
  var sync = { sheetOk: false, driveOk: false, sheetMsg: '', driveMsg: '' };

  var data = JSON.parse(payload);
  data.category = inferCategory_(data);
  data = normalizePosterData(data);

  var html = buildPosterHtml(data);

  var pdfFile;
  try {
    var pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
    var dateForName = formatDateForFilename(data.dateDay || '');
    var titleForName = sanitizeFilename(data.title || 'sans-titre');
    pdfBlob.setName('HER-' + dateForName + '-' + titleForName + '.pdf');
    var targetFolder = ensureAfficheFolder();
    pdfFile = targetFolder.createFile(pdfBlob);
    sync.driveOk = true;
  } catch (e) {
    sync.driveOk = false;
    sync.driveMsg = (e && e.message) ? e.message : String(e);
    throw e;
  }

  var entryRecord;
  try {
    var entryId = data.entryId || '';
    entryRecord = savePosterEntry(entryId, data, pdfFile ? pdfFile.getUrl() : '');
    sync.sheetOk = true;
  } catch (e2) {
    sync.sheetOk = false;
    sync.sheetMsg = (e2 && e2.message) ? e2.message : String(e2);
    // Nettoyage si le PDF a été créé mais pas la ligne
    try { if (pdfFile) pdfFile.setTrashed(true); } catch (e3) {}
    throw e2;
  }

  return {
    entryId: entryRecord.id,
    pdfUrl: pdfFile.getUrl(),
    sync: sync
  };
}

function listPosterEntries() {
  var sheet = ensureAfficheSheet();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  var headers = values[0];
  var tz = Session.getScriptTimeZone();
  var out = [];

  var categoryCol = getHeaderIndex_(headers, 'category', 21);

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var entry = {};

    for (var c = 0; c < headers.length; c++) {
      var header = headers[c];
      var value = row[c];

      if (value instanceof Date) {
        if (header === 'createdAt' || header === 'updatedAt') {
          entry[header] = Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm');
        } else {
          entry[header] = Utilities.formatDate(value, tz, 'dd/MM/yyyy');
        }
      } else {
        entry[header] = (value !== null && value !== undefined) ? String(value) : '';
      }
    }

    if (!entry.id) continue;

    // Réparation : anciennes lignes avec category vide
    var inferred = inferCategory_(entry);
    var current = String(entry.category || '').trim().toLowerCase();
    if (!current) {
      entry.category = inferred;
      if (categoryCol >= 0) {
        try {
          sheet.getRange(r + 1, categoryCol + 1).setValue(inferred);
        } catch (e) {}
      }
      try { SpreadsheetApp.flush(); } catch(e2) {}
    }

    out.push(entry);
  }

  return out.reverse();
}



function ensureAfficheSpreadsheet() {
  var sheetName = 'AFFICHE HERMINE A3';
  var propertyKey = 'AFFICHE_HERMINE_A3_SHEET_ID';
  var storedId = PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (error) {
      PropertiesService.getScriptProperties().deleteProperty(propertyKey);
    }
  }
  var files = DriveApp.getFilesByName(sheetName);
  if (files.hasNext()) {
    var existing = files.next();
    PropertiesService.getScriptProperties().setProperty(propertyKey, existing.getId());
    return SpreadsheetApp.open(existing);
  }
  var created = SpreadsheetApp.create(sheetName);
  PropertiesService.getScriptProperties().setProperty(propertyKey, created.getId());
  var targetFolder = ensureAfficheFolder();
  var createdFile = DriveApp.getFileById(created.getId());
  targetFolder.addFile(createdFile);
  DriveApp.getRootFolder().removeFile(createdFile);
  return created;
}

function ensureAfficheSheet() {
  var sheetName = 'AFFICHE HERMINE A3';
  var spreadsheet = ensureAfficheSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  ensureSheetHeaders(sheet);
  return sheet;
}

function ensureSheetHeaders(sheet) {
  // En-têtes canoniques (ordre attendu par l'app)
  var headers = [
    'id',
    'createdAt',
    'updatedAt',
    'title',
    'titleLine2',
    'subtitle',
    'subtitleText',
    'dateDay',
    'dateTime',
    'dateText',
    'footer',
    'tarifs',
    'tariffTier',
    'footerInfoAuto',
    'footerInfo',
    'footerVenueAuto',
    'footerVenue',
    'credit',
    'hermineChoice',
    'pdfUrl',
    'category'
  ];

  // Assure les en-têtes sur les colonnes canoniques
  var existing = sheet.getLastColumn() >= headers.length
    ? sheet.getRange(1, 1, 1, headers.length).getValues()[0]
    : [];
  var needsUpdate = existing.length !== headers.length || existing.some(function(value, index) {
    return value !== headers[index];
  });
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // Corrige les doublons (ex: 2 colonnes "category") et sécurise l'indexation
  try {
    sanitizeDuplicateHeaders_(sheet, headers);
  } catch (e) {
    // On ne bloque pas l'app si la feuille est protégée ou si une correction échoue.
  }
}


/**
 * Renomme les en-têtes en doublon au-delà des colonnes canoniques.
 * Cas particulier : si une 2e colonne "category" existe, on rapatrie sa valeur
 * dans la colonne canonique "category" lorsque celle-ci est vide.
 */
function sanitizeDuplicateHeaders_(sheet, canonicalHeaders) {
  var lastCol = sheet.getLastColumn();
  if (lastCol <= canonicalHeaders.length) return;

  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h || '').trim(); });
  var canonicalLen = canonicalHeaders.length;

  var canonicalIndexByName = {};
  for (var i = 0; i < canonicalHeaders.length; i++) {
    canonicalIndexByName[canonicalHeaders[i]] = i; // 0-based
  }

  var seen = {};
  // D'abord, marque les en-têtes canoniques comme vus
  for (var c = 0; c < canonicalLen; c++) {
    var name = String(headerRow[c] || '').trim();
    if (!name) continue;
    seen[name] = 1;
  }

  // Spécifique : si une colonne "category" en trop existe, rapatrier si besoin
  var canonicalCategoryCol = canonicalIndexByName['category']; // 0-based
  var extraCategoryCols = [];
  for (var c2 = canonicalLen; c2 < lastCol; c2++) {
    if (String(headerRow[c2] || '').trim() === 'category') extraCategoryCols.push(c2);
  }
  if (extraCategoryCols.length > 0 && canonicalCategoryCol !== undefined) {
    var nRows = sheet.getLastRow();
    if (nRows > 1) {
      var canonRange = sheet.getRange(2, canonicalCategoryCol + 1, nRows - 1, 1);
      var canonVals = canonRange.getValues();
      // Pour la première colonne "category" extra, on rapatrie les valeurs manquantes
      var extraCol = extraCategoryCols[0];
      var extraRange = sheet.getRange(2, extraCol + 1, nRows - 1, 1);
      var extraVals = extraRange.getValues();
      var changed = false;
      for (var r = 0; r < canonVals.length; r++) {
        var canon = String(canonVals[r][0] || '').trim();
        var extra = String(extraVals[r][0] || '').trim();
        if (!canon && extra) {
          canonVals[r][0] = extra;
          changed = true;
        }
      }
      if (changed) {
        canonRange.setValues(canonVals);
      }
    }
  }

  // Renommage des colonnes en trop si elles dupliquent un en-tête canonique ou déjà vu
  var updates = [];
  for (var c3 = canonicalLen; c3 < lastCol; c3++) {
    var h = String(headerRow[c3] || '').trim();
    if (!h) continue;

    var base = h;
    var isDup = (canonicalIndexByName[base] !== undefined) || (seen[base] !== undefined);

    if (isDup) {
      var suffix = 1;
      var proposed = 'legacy_' + base;
      // Évite de créer un nouveau doublon
      while (canonicalIndexByName[proposed] !== undefined || seen[proposed] !== undefined) {
        suffix++;
        proposed = 'legacy_' + base + '_' + suffix;
      }
      updates.push({ col: c3 + 1, value: proposed });
      seen[proposed] = 1;
    } else {
      seen[base] = 1;
    }
  }

  if (updates.length) {
    updates.forEach(function(u){
      sheet.getRange(1, u.col).setValue(u.value);
    });
  }
}

/**
 * Renvoie l'index (0-based) de l'en-tête recherché dans la zone canonique si possible,
 * sinon retombe sur la dernière occurrence trouvée.
 */
function getHeaderIndex_(headers, name, canonicalLen) {
  if (!headers || !headers.length) return -1;
  var n = String(name || '').trim();
  var canonLen = canonicalLen || 0;

  // 1) Priorité : position canonique si elle correspond
  if (canonLen > 0) {
    for (var i = 0; i < Math.min(canonLen, headers.length); i++) {
      if (String(headers[i] || '').trim() === n) return i;
    }
  }

  // 2) Sinon : dernière occurrence
  var idx = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j] || '').trim() === n) idx = j;
  }
  return idx;
}

function ensureAfficheFolder() {
  var folderName = 'AFFICHE HERMINE A3';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function savePosterEntry(entryId, data, pdfUrl) {
  var sheet = ensureAfficheSheet();
  var now = new Date();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var payload = {
    id: entryId || Utilities.getUuid(),
    createdAt: now,
    updatedAt: now,
    title: data.title || '',
    titleLine2: data.titleLine2 || '',
    subtitle: data.subtitle || '',
    subtitleText: data.subtitleText || '',
    dateDay: data.dateDay || '',
    dateTime: data.dateTime || '',
    dateText: data.dateText || '',
    footer: data.footer || '',
    tarifs: data.tarifs || '',
    tariffTier: data.tariffTier || '',
    footerInfoAuto: isTruthy_(data.footerInfoAuto),
    footerInfo: data.footerInfo || '',
    footerVenueAuto: isTruthy_(data.footerVenueAuto),
    footerVenue: data.footerVenue || '',
    credit: data.credit || '',
    hermineChoice: data.hermineChoice || '',
    pdfUrl: pdfUrl || '',
    category: inferCategory_(data)
  };

  var rowIndex = findEntryRow(sheet, payload.id);
  if (rowIndex) {
    var createdAtCol = headers.indexOf('createdAt');
    if (createdAtCol >= 0) {
      payload.createdAt = sheet.getRange(rowIndex, createdAtCol + 1).getValue();
    }
    payload.updatedAt = now;
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(function(header) {
      return (payload.hasOwnProperty(header) ? payload[header] : '');
    })]);
  } else {
    sheet.appendRow(headers.map(function(header) {
      return (payload.hasOwnProperty(header) ? payload[header] : '');
    }));
  }
  try { SpreadsheetApp.flush(); } catch (e) {}
  return payload;
}

function findEntryRow(sheet, entryId) {
  if (!entryId) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }
  var searchId = String(entryId).trim();
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchId) {
      return i + 2;
    }
  }
  return 0;
}

function deletePosterEntry(entryId) {
  var sync = { sheetOk: false, driveOk: false, sheetMsg: '', driveMsg: '' };
  if (!entryId) throw new Error('ID manquant');
  var sheet = ensureAfficheSheet();
  var rowIndex = findEntryRow(sheet, entryId);
  if (!rowIndex) throw new Error('Entrée introuvable');

  sync.driveOk = true;
  // Récupérer l'URL du PDF pour supprimer le fichier du Drive
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var pdfColIndex = headers.indexOf('pdfUrl');
  if (pdfColIndex >= 0) {
    var pdfUrl = sheet.getRange(rowIndex, pdfColIndex + 1).getValue();
    if (pdfUrl) {
      try {
        var fileIdMatch = pdfUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileIdMatch) {
          DriveApp.getFileById(fileIdMatch[1]).setTrashed(true);
        }
      } catch (e) {
        sync.driveOk = false;
        sync.driveMsg = (e && e.message) ? e.message : String(e);
        // Le fichier n'existe peut-être plus, on continue
      }
    }
  }

  // Supprimer la ligne du sheet
  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  sync.sheetOk = true;
  return { ok: true, sync: sync };
}




function listPosterEntriesEx() {
  var sync = { sheetOk: false, driveOk: false, sheetMsg: '', driveMsg: '' };
  try {
    // Vérifie l'accès au Sheet (lecture)
    var sheet = ensureAfficheSheet();
    sheet.getLastRow();
    sync.sheetOk = true;
  } catch (e) {
    sync.sheetOk = false;
    sync.sheetMsg = (e && e.message) ? e.message : String(e);
  }

  try {
    // Vérifie l'accès au Drive (dossier)
    ensureAfficheFolder().getName();
    sync.driveOk = true;
  } catch (e2) {
    sync.driveOk = false;
    sync.driveMsg = (e2 && e2.message) ? e2.message : String(e2);
  }

  var entries = [];
  try {
    entries = listPosterEntries();
  } catch (e3) {
    sync.sheetOk = false;
    sync.sheetMsg = (e3 && e3.message) ? e3.message : String(e3);
  }

  return { entries: entries, sync: sync };
}

function togglePosterCategory(entryId) {
  try {
    if (!entryId) return { ok: false, error: 'ID manquant' };

    var sheet = ensureAfficheSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowIndex = findEntryRow(sheet, entryId);
    if (!rowIndex) return { ok: false, error: 'Entrée introuvable pour id=' + entryId };

    var categoryCol = getHeaderIndex_(headers, 'category', 21);
    if (categoryCol < 0) return { ok: false, error: 'Colonne category introuvable' };

    var rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    var entry = {};
    for (var c = 0; c < headers.length; c++) {
      entry[headers[c]] = rowValues[c];
    }

    var currentRaw = String(entry.category || '').trim().toLowerCase();
    var inferred = inferCategory_(entry);
    var current = currentRaw || inferred;
    var next = (current === 'spectacle') ? 'atelier' : 'spectacle';

    sheet.getRange(rowIndex, categoryCol + 1).setValue(next);

    var updatedAtCol = headers.indexOf('updatedAt');
    if (updatedAtCol >= 0) {
      sheet.getRange(rowIndex, updatedAtCol + 1).setValue(new Date());
    }

    SpreadsheetApp.flush();

    return { ok: true, newCategory: next };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}


function inferCategory_(data) {
  var raw = String((data && data.category) || '').trim().toLowerCase();
  if (raw === 'spectacle') return 'spectacle';
  if (raw === 'atelier') return 'atelier';

  if (!data) return 'atelier';

  var hasTier = String(data.tariffTier || '').trim().length > 0;
  var hasCredit = String(data.credit || '').trim().length > 0;
  var hasInfo = String(data.footerInfo || '').trim().length > 0;
  var hasVenue = String(data.footerVenue || '').trim().length > 0;
  var hasInfoAuto = isTruthy_(data.footerInfoAuto);
  var hasVenueAuto = isTruthy_(data.footerVenueAuto);

  // Heuristique legacy : ancien "footer" unique commençant par une ligne de tarifs
  var footerFirst = String((data.footer || '').split(/\r?\n/)[0] || '').trim().toLowerCase();
  var footerLooksSpectacle = footerFirst.indexOf('tarif unique') === 0 || footerFirst.indexOf('tarifs de') === 0;

  if (hasTier || hasCredit || hasInfo || hasVenue || hasInfoAuto || hasVenueAuto || footerLooksSpectacle) {
    return 'spectacle';
  }
  return 'atelier';
}


function normalizePosterData(data) {
  if (!data) return data;
  var category = inferCategory_(data);
  data.category = category;
  if (category !== 'spectacle') return data;

  // Tarifs (ligne 1)
  var tarifsLine = getSpectacleTarifLine_(data);
  data.tarifs = tarifsLine;

  // Blocs bas de page (hors tarifs)
  var infoText = '';
  var venueText = '';

  var hasBlocks = (data.footerInfo != null) || (data.footerVenue != null) || (data.footerInfoAuto != null) || (data.footerVenueAuto != null);

  if (hasBlocks) {
    infoText = isTruthy_(data.footerInfoAuto) ? buildSpectacleFooterInfoAuto_() : String(data.footerInfo || '');
    venueText = isTruthy_(data.footerVenueAuto) ? buildSpectacleFooterVenueAuto_() : String(data.footerVenue || '');
  } else {
    // Compat (anciennes versions) : footerAuto + footer unique
    var footerRest = isTruthy_(data.footerAuto) ? buildSpectacleFooterRestAuto_() : String(data.footer || '');
    footerRest = stripLeadingSpectacleTarifLine_(footerRest);

    // Split en 2 blocs si possible (séparateur: ligne vide)
    var rest = normalizeLineBreaks_(footerRest);
    var parts = rest.split('\n\n');
    infoText = (parts[0] || '').trim();
    venueText = (parts.slice(1).join('\n\n') || '').trim();
  }

  infoText = normalizeLineBreaks_(infoText);
  venueText = normalizeLineBreaks_(venueText);

  // Anti-duplication : supprimer toute ligne "Tarif(s)..." si elle a été collée dans un bloc
  infoText = stripTarifLinesAnywhere_(infoText);
  venueText = stripTarifLinesAnywhere_(venueText);

  infoText = normalizeFooterBlock_(infoText);
  venueText = normalizeFooterBlock_(venueText);

  // Reconstruire le footer complet
  var footerRest2 = '';
  if (infoText && venueText) footerRest2 = infoText + '\n\n' + venueText;
  else footerRest2 = infoText || venueText || '';

  if (String(tarifsLine || '').trim()) {
    data.footer = String(tarifsLine).trim() + (footerRest2 ? ('\n' + footerRest2) : '');
  } else {
    data.footer = footerRest2;
  }

  // Stockage
  data.footerInfo = infoText;
  data.footerVenue = venueText;

  return data;
}


function getSpectacleTarifLine_(data) {
  var tier = String(data.tariffTier || '').trim().toUpperCase();
    var map = {
    'A+': 'Tarifs de 10 à 22€',
    'A': 'Tarifs de 8 à 17€',
    'B': 'Tarifs de 6 à 13€',
    'C': 'Tarif unique de 5€',
    'ENTREE_LIBRE': 'Entrée libre',
    'GRATUIT': 'Gratuit'
  };
  if (tier && map[tier]) {
    return map[tier];
  }

    var t = String(data.tarifs || '').trim();
  if (t) {
    var lowT = t.toLowerCase();
    if (lowT.indexOf('entrée libre') === 0 || lowT.indexOf('entree libre') === 0) return 'Entrée libre';
    if (lowT.indexOf('gratuit') === 0) return 'Gratuit';
    if (lowT.indexOf('tarif unique') === 0 || lowT.indexOf('tarifs de') === 0) {
      return t;
    }
  }

  var firstLine = String((data.footer || '').split(/\r?\n/)[0] || '').trim();
  var lowFirst = firstLine.toLowerCase();
  if (lowFirst.indexOf('entrée libre') === 0 || lowFirst.indexOf('entree libre') === 0) {
    return 'Entrée libre';
  }
  if (lowFirst.indexOf('gratuit') === 0) {
    return 'Gratuit';
  }
  if (lowFirst.indexOf('tarif unique') === 0 || lowFirst.indexOf('tarifs de') === 0) {
    return firstLine;
  }

  return map['C'];
}



function buildSpectacleFooterRestAuto_() {
  var lines = [];
  lines.push('Infos > 02 97 48 29 40 et lhermine.bzh');
  lines.push('Réservations > billetterie.lhermine.bzh');
  lines.push('');
  lines.push("espace culturel l'hermine");
  lines.push('rue du Père Coudrin / 56370 Sarzeau');
  return lines.join('\n');
}

function normalizeLineBreaks_(text) {
  var s = String(text || '');
  // Pas de regex ici pour éviter les caractères de contrôle en source
  s = s.split('\r\n').join('\n');
  s = s.split('\r').join('\n');
  return s;
}

function buildSpectacleFooterInfoAuto_() {
  return [
    'Infos > 02 97 48 29 40 et lhermine.bzh',
    'Réservations > billetterie.lhermine.bzh'
  ].join('\n');
}

function buildSpectacleFooterVenueAuto_() {
  return [
    "espace culturel l'hermine",
    'rue du Père Coudrin / 56370 Sarzeau'
  ].join('\n');
}

function stripTarifLinesAnywhere_(text) {
  var t = normalizeLineBreaks_(text);
  var lines = t.split('\n');
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || '');
    var low = line.trim().toLowerCase();
    if (!low) { out.push(line); continue; }
        if (low.indexOf('tarif unique') === 0) continue;
    if (low.indexOf('tarifs de') === 0) continue;
    if (low === 'entrée libre' || low === 'entree libre') continue;
    if (low === 'gratuit') continue;
    out.push(line);
  }
  return out.join('\n');
}

function normalizeFooterBlock_(text) {
  var t = normalizeLineBreaks_(text);
  var lines = t.split('\n');

  while (lines.length && String(lines[0] || '').trim() === '') lines.shift();
  while (lines.length && String(lines[lines.length - 1] || '').trim() === '') lines.pop();

  return lines.join('\n').trimEnd ? lines.join('\n').trimEnd() : lines.join('\n').replace(/\s+$/, '');
}


function stripLeadingSpectacleTarifLine_(text) {
  var t = String(text || '').replace(/\r\n/g, '\n');
  var lines = t.split('\n');
  if (!lines.length) return '';
  var first = (lines[0] || '').trim().toLowerCase();
  if (first.indexOf('tarif unique') === 0 || first.indexOf('tarifs de') === 0 || first === 'entrée libre' || first === 'entree libre' || first === 'gratuit') {
    lines.shift();
    while (lines.length && lines[0] === '') lines.shift();
    return lines.join('\n');
  }
  return t;
}


function isTruthy_(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  var s = String(value).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'on';
}


function buildPosterHtml(data) {
  var rawCategory = (data.category || '').toLowerCase();
  var category = (rawCategory === 'spectacle') ? 'spectacle' : 'atelier';
  var isSpectacle = category === 'spectacle';

  var hasTitleLine2 = !!data.titleLine2;
  var titleFontSize = calculateTitleFontSize(data.title, hasTitleLine2);
  var titleLine2 = data.titleLine2
    ? '<span class="title-line">' + escapeHtml(data.titleLine2) + '</span>'
    : '';
  var titleHtml = '<span class="title-line">' + escapeHtml(data.title) + '</span>' + titleLine2;

  var formattedDate = formatLongDate(data.dateDay || '');
  if (isSpectacle && formattedDate) {
    formattedDate = formattedDate.charAt(0).toUpperCase() + formattedDate.slice(1);
  }
  var dateDisplay = escapeHtml(formattedDate);
  if ((data.dateTime || '').trim()) {
    dateDisplay += ' / ' + escapeHtml(data.dateTime.trim());
  }

  var topLogo = data.topLogo || '';
  var bottomLogo = data.bottomLogo || '';
  var mainImage = data.mainImage || '';
  var credit = (data.credit || '').trim();
  var footerText = data.footer || '';
  var fontVagRounded = data.fontVagRounded || '';

  var mainImgTag = mainImage ? '<img class="main-photo" src="' + mainImage + '" alt="Photo" />' : '';
  var fontFaceStyle = fontVagRounded
    ? '@font-face { font-family: "VAG Rounded Std"; font-style: normal; font-weight: 700; src: url("' + fontVagRounded + '"); }'
    : '';

  // Contenu (atelier vs spectacle)
  var optionalSubtitle = '';
  var optionalDescription = '';
  var optionalDateText = '';

  if (isSpectacle) {
    // Fusion : "subtitle" + "dateText" (anciens champs compagnie / mention) dans une seule zone.
    var mergedCaption = '';
    if (data.subtitle) mergedCaption = String(data.subtitle);
    if (data.dateText) mergedCaption = mergedCaption ? (mergedCaption + '\n' + String(data.dateText)) : String(data.dateText);
    mergedCaption = normalizeLineBreaks_(mergedCaption);

    if (mergedCaption.trim()) {
      optionalSubtitle = '<p class="sp-line sp-company">' + nl2brHtml_(escapeHtml(mergedCaption)) + '</p>';
    }

    if (data.subtitleText) {
      var genre = normalizeLineBreaks_(String(data.subtitleText));
      optionalDescription = '<p class="sp-line sp-genre">' + nl2brHtml_(escapeHtml(genre)) + '</p>';
    }
  } else {
    optionalSubtitle = data.subtitle
      ? '<h2 class="subtitle">' + escapeHtml(data.subtitle) + '</h2>'
      : '';
    optionalDescription = data.subtitleText
      ? '<p class="description">' + nl2brHtml_(escapeHtml(normalizeLineBreaks_(String(data.subtitleText)))) + '</p>'
      : '';
    optionalDateText = data.dateText
      ? '<p class="date-text">' + nl2brHtml_(escapeHtml(normalizeLineBreaks_(String(data.dateText)))) + '</p>'
      : '';
  }

  // Footer (atelier: dernier ligne en lien / spectacle: conserve les retours à la ligne)
  var footerMain = '';
  var footerLink = '';
  if (isSpectacle) {
    footerMain = '<div class="footer-pre">' + escapeHtml(footerText || '') + '</div>';
  } else {
    var footerParts = buildFooterHtmlParts(footerText);
    footerMain = footerParts.main;
    footerLink = footerParts.link;
  }

  return (
    '<!DOCTYPE html>' +
    '<html lang="fr">' +
    '<head>' +
    '<meta charset="utf-8">' +
    '<style>' +
    fontFaceStyle +
    '@page { size: 841.89pt 1190.55pt; margin: 0; }' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'html, body { width: 841.89pt; height: 1190.55pt; font-family: "VAG Rounded Std", "Arial Rounded MT Bold", "Helvetica Neue", Helvetica, Arial, sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }' +
    '.poster { width: 841.89pt; height: 1190.55pt; display: flex; flex-direction: column; }' +
    '.image-section { width: 841.89pt; flex: 1 1 auto; position: relative; overflow: hidden; }' +
    '.main-photo { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; }' +
    '.top-logo { position: absolute; top: 15mm; left: 15mm; width: 67mm; }' +
    '.photo-credit { position: absolute; bottom: 5mm; right: 15mm; font-size: 10pt; color: #fff; text-shadow: 0 0 3px rgba(0,0,0,0.7); text-align: right; }' +
    '.content { width: 841.89pt; background: #f2d20a; padding: 15mm 20mm; }' +
    '.title { font-size: 60pt; font-weight: 700; margin: 0 0 3mm 0; line-height: 1.05; width: 70%;' +
    (hasTitleLine2 ? '' : ' white-space: nowrap;') +
    ' }' +
    '.title-line { display: block; }' +
    '.subtitle { font-size: 24pt; font-weight: 700; margin: 0 0 3mm 0; line-height: 1.15; }' +
    '.description { font-size: 24pt; font-weight: 700; margin: 0 0 3mm 0; line-height: 1.3; }' +
    '.date { font-size: 46pt; font-weight: 700; margin: 0; line-height: 1.1; }' +
    '.date-text { font-size: 18pt; font-weight: 400; margin: 3mm 0 0 0; line-height: 1.3; }' +
    '.footer { margin-top: 10mm; display: flex; justify-content: space-between; align-items: flex-end; }' +
    '.footer .info { font-size: 18pt; line-height: 1; font-weight: 700; }' +
    '.footer-pre { white-space: pre-line; }' +
    '.footer .info .footer-line { display: block; }' +
    '.footer .info .footer-blank { display: block; }' +
    '.footer .info .footer-link { display: block; margin-top: 3mm; }' +
    '.bottom-logo { width: 40mm; }' +
    '.spectacle .date { font-size: 46pt; }' +
    '.spectacle .footer .info { font-size: 16pt; line-height: 1.15; }' +
    '.spectacle .bottom-logo { width: 55mm; }' +
    '.spectacle .sp-line { font-size: 24pt; font-weight: 700; margin: 0 0 3mm 0; line-height: 1.15; }' +
    '</style>' +
    '</head>' +
    '<body class="' + category + '">' +
    '<div class="poster">' +
    '<div class="image-section">' +
    mainImgTag +
    (topLogo ? '<img class="top-logo" src="' + topLogo + '" alt="Logo" />' : '') +
    (credit ? '<div class="photo-credit">' + escapeHtml(credit) + '</div>' : '') +
    '</div>' +
    '<div class="content">' +
    '<h1 class="title">' + titleHtml + '</h1>' +
    optionalSubtitle +
    optionalDescription +
    '<p class="date">' + dateDisplay + '</p>' +
    optionalDateText +
    '<div class="footer">' +
    '<div class="info">' + footerMain + footerLink + '</div>' +
    (bottomLogo ? '<img class="bottom-logo" src="' + bottomLogo + '" alt="Logo Sarzeau" />' : '') +
    '</div>' +
    '</div>' +
    '</div>' +
    '</body>' +
    '</html>'
  );
}


function escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function nl2brHtml_(text) {
  return String(text || '').split('\n').join('<br>');
}

function formatDateForFilename(input) {
  var trimmed = (input || '').trim();
  if (!trimmed) return '00-00-00';
  var date;
  var isoMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    date = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
  } else {
    var match = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
    if (match) {
      var year = parseInt(match[3], 10);
      if (year < 100) year += 2000;
      date = new Date(year, parseInt(match[2], 10) - 1, parseInt(match[1], 10));
    } else if (!isNaN(new Date(trimmed).getTime())) {
      date = new Date(trimmed);
    }
  }
  if (!date || isNaN(date.getTime())) return '00-00-00';
  var yy = String(date.getFullYear()).slice(-2);
  var mm = ('0' + (date.getMonth() + 1)).slice(-2);
  var dd = ('0' + date.getDate()).slice(-2);
  return yy + '-' + mm + '-' + dd;
}

function sanitizeFilename(text) {
  return text
    .replace(/[\/\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 80);
}

function formatLongDate(input) {
  var trimmed = (input || '').trim();
  if (!trimmed) return '';
  var match = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (!match) {
    return trimmed;
  }
  var day = parseInt(match[1], 10);
  var month = parseInt(match[2], 10) - 1;
  var year = parseInt(match[3], 10);
  if (year < 100) {
    year += 2000;
  }
  var date = new Date(year, month, day);
  if (isNaN(date.getTime())) {
    return trimmed;
  }
  var dayNames = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
  var monthNames = [
    'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
    'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'
  ];
  return dayNames[date.getDay()] + ' ' + day + ' ' + monthNames[date.getMonth()];
}

function calculateTitleFontSize(title, hasTitleLine2) {
  if (hasTitleLine2) {
    return 50;
  }
  var trimmed = (title || '').trim();
  if (!trimmed) {
    return 50;
  }
  var length = trimmed.length;
  if (length <= 20) {
    return 50;
  }
  if (length <= 28) {
    return 48;
  }
  if (length <= 36) {
    return 46;
  }
  if (length <= 44) {
    return 44;
  }
  if (length <= 52) {
    return 42;
  }
  return 40;
}

function buildFooterHtmlParts(footerText) {
  var lines = (footerText || '').split(/\r?\n/).map(function(line) {
    return line.trim();
  }).filter(function(line) {
    return line.length > 0;
  });
  if (lines.length === 0) {
    return { main: '', link: '' };
  }

  var lastLine = lines[lines.length - 1];
  var useLink = lastLine.indexOf('►') === 0;
  var mainLines = useLink ? lines.slice(0, -1) : lines;
  var linkLine = useLink ? lastLine : '';

  var mainHtml = mainLines.length
    ? mainLines.map(function(line) { return escapeHtml(line); }).join('<br>')
    : '';
  var linkHtml = linkLine
    ? '<span class="footer-link">' + escapeHtml(linkLine) + '</span>'
    : '';

  return {
    main: mainHtml ? mainHtml + (linkHtml ? '<br>' : '') : '',
    link: linkHtml
  };
}


