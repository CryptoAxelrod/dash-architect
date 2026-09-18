/*
 * Dash Architect for Google Sheets — the only file in this add-on that
 * touches SpreadsheetApp/DriveApp/Utilities. Everything deterministic
 * (engine/, render/) and everything account-related (addin/auth.js) is
 * loaded unchanged by addin-sheets/sidebar.html and addin-sheets/dashboard.html
 * from the same hosted origin the Excel add-in uses — see those files'
 * <script> tags. This file is the Apps Script analog of addin/excel-io.js +
 * addin/dashboard-io.js's "touches the host document" halves, called from
 * both pages via google.script.run instead of Excel.run.
 *
 * Two UI surfaces, like the Excel add-in (SPEC.md §14.5), for the same
 * reason: a sidebar is too narrow (~300-360px) for a 1180px dashboard
 * canvas. Unlike Excel's dialog<->parent messaging (needed because an
 * Office dialog cannot touch the workbook at all), the sidebar and the
 * dashboard dialog here never talk to each other directly — there is no
 * supported way for two separate HtmlService surfaces to message one
 * another, and building one blind would be unverifiable in this
 * environment (see SPEC.md §15.1). Instead:
 *   - the sidebar hands off a small descriptor (which range, which theme,
 *     which saved role overrides — never the analysis itself) to
 *     `openDashboardDialog`, which inlines it into the dialog's own HTML at
 *     render time via HtmlTemplate (`<?!= initialPayloadJson ?>` in
 *     dashboard.html) — a one-way, one-time handoff, not a channel;
 *   - after that, the dialog is fully self-sufficient: it re-reads the
 *     range and reclassifies itself, and every later action (refresh,
 *     change data range, place on sheet, review column roles) calls this
 *     file directly, exactly like the sidebar does. Nothing here is
 *     Excel's protocol re-implemented — it's not needed.
 *
 * Persistent storage note (CLAUDE.md's determinism/zero-network rules do
 * not apply to this file — it's the /addin-equivalent host seam, same as
 * excel-io.js/dashboard-io.js are for Excel): dashboards are kept in a
 * hidden helper sheet (HIDDEN_SHEET_NAME), one row per dashboard, not
 * PropertiesService — a saved layoutSpec snapshot can run to "10s of KB"
 * (SPEC.md §14.3), well past PropertiesService's 9KB-per-value / 500KB-
 * total caps, while a single cell comfortably holds up to 50,000
 * characters.
 */

var HIDDEN_SHEET_NAME = '_DashArchitectData';
var SHAPE_NAME_PREFIX = 'DashArchitectDashboard_';
var IMAGE_SCALE = 0.5; // matches addin/dashboard-io.js's CONFIG.imageScale
var FORMAT_SAMPLE_ROWS = 200; // matches addin-sheets/sheets-io.js's CONFIG.formatSampleRows

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dash Architect')
    .addItem('Open', 'showSidebar')
    .addToUi();
}

// Workspace add-on homepage trigger (appsscript.json's addOns.common) — the
// side-panel "open" card some Sheets add-on surfaces show instead of (or
// alongside) the classic menu item above. Same entry point either way.
function onHomepage() {
  showSidebar();
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Dash Architect'))
    .addSection(CardService.newCardSection().addWidget(
      CardService.newTextParagraph().setText('Opened in the sidebar — you can close this panel.')
    ))
    .build();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('Dash Architect');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the live/frozen dashboard in its own modeless (non-blocking —
 * needed for "Change data range" to work) dialog, wide enough for the
 * 1180px canvas. `payload` is intentionally tiny — see this file's header
 * comment for why this is a one-time handoff, not a channel:
 *   - live: {mode:'live', source, roleOverrides, title, theme, palette, shapeName, isPro}
 *   - frozen: {mode:'frozen', layoutSpec, title, theme, palette}
 */
function openDashboardDialog(payload) {
  var tmpl = HtmlService.createTemplateFromFile('dashboard');
  // Escaping every "<" (not just the ones spelling "</script") is
  // deliberate: a title or sheet name is arbitrary user/cell content
  // (CLAUDE.md never assumes sheet data is trusted), and this string is
  // inlined RAW (`<?!= ?>`, not HTML-escaped) into dashboard.html's own
  // <script> tag — "<" is the only character that lets that content break
  // out of it. Replacing it with its unicode escape keeps the resulting
  // JS string value identical while the literal "<" never appears in the
  // HTML source for the browser's HTML parser to act on.
  tmpl.initialPayloadJson = JSON.stringify(payload).replace(/</g, '\\u003c');
  var html = tmpl.evaluate().setWidth(1200).setHeight(820);
  SpreadsheetApp.getUi().showModelessDialog(html, (payload && payload.title) || 'Dash Architect');
}

/* ---------- range selection ---------- */

/**
 * Cheap, address-only read — polled every ~600ms by sidebar.html while the
 * user is actively picking (see its startPicking/pollSelection), to fake
 * the live address feedback Office.js's DocumentSelectionChanged event
 * gives Excel's task pane for free. Apps Script has no push equivalent —
 * a server-side onSelectionChange simple trigger exists, but it cannot
 * proactively notify an already-open sidebar, only a client pull can.
 * @returns {{address:string, sheetName:string, a1:string, rows:number, cols:number}|null}
 */
function getSelectionInfo() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) return null;
  return describeRange_(range);
}

function describeRange_(range) {
  var sheet = range.getSheet();
  return {
    address: sheet.getName() + '!' + range.getA1Notation(),
    sheetName: sheet.getName(),
    a1: range.getA1Notation(),
    rows: range.getNumRows(),
    cols: range.getNumColumns(),
  };
}

/**
 * The confirmed pick — Excel's parallel (addin/excel-io.js#detectSource)
 * also detects whether the selection sits inside an Excel Table. Sheets
 * has no equivalent Apps Script API for its own newer "Tables" feature as
 * of this writing, so this add-on only ever produces a plain range source
 * — a deliberate scope reduction (SPEC.md §15 documents Excel's own
 * Table-detection code the same way it documents everything Sheets-side
 * that's different by necessity, not by omission).
 * @returns {{kind:'range', address:string, sheetName:string, a1:string, rows:number, cols:number, headers:Array<string>}|null}
 */
function detectSource() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) return null;
  var info = describeRange_(range);
  var headerRow = range.getSheet().getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns());
  var headers = headerRow.getValues()[0].map(function (v) { return v == null ? '' : String(v); });
  // Body = every selected row after the first (header) row. Range#isBlank()
  // is a cheap native "is there anything at all in here" check — no
  // getValues() round trip into the Apps Script runtime, unlike
  // readRangeRaw's full read (only done once Generate is actually clicked)
  // — see sidebar.html's finishPicking insufficient-data gate, which is
  // what this feeds. info.rows<=1 (no body rows at all) is already caught
  // there before detectSource ever runs, so this is never actually called
  // with nothing to check.
  var hasData = info.rows <= 1 ? false : !range.getSheet().getRange(range.getRow() + 1, range.getColumn(), info.rows - 1, info.cols).isBlank();
  return { kind: 'range', address: info.address, sheetName: info.sheetName, a1: info.a1, rows: info.rows, cols: info.cols, headers: headers, hasData: hasData };
}

/**
 * Reads `source` fresh (never a range captured once at pick time — a
 * renamed sheet or a grown selection is picked up naturally) and returns
 * plain arrays for addin-sheets/sheets-io.js#toEngineInput to convert —
 * this file never converts to engine input itself, matching the pure/host
 * split addin/excel-io.js uses. numberFormats are sampled the same way
 * (first FORMAT_SAMPLE_ROWS rows only) since a column's format is
 * realistically uniform and reading it for a 200,000-row range is where
 * the real cost is; values are read for the whole range in one call —
 * there is no per-call marshaling overhead here worth chunking against
 * the way Excel's context.sync() has (see addin-sheets/sheets-io.js's doc
 * comment).
 * @param {{sheetName:string, a1:string}} source
 * @returns {{values, numberFormatSample, totalRows, cols, address, timing}}
 */
function readRangeRaw(source) {
  var t0 = Date.now();
  var sheet = SpreadsheetApp.getActive().getSheetByName(source.sheetName);
  if (!sheet) throw new Error('Sheet "' + source.sheetName + '" no longer exists — it may have been deleted or renamed.');
  var range = sheet.getRange(source.a1);
  // Date-formatted cells come back from getValues() as real JS Date
  // objects — google.script.run's client<->server bridge only reliably
  // marshals JSON-safe values (Number/String/Boolean/null/plain
  // Array/Object), not Date, so those are converted to a plain
  // 'yyyy-MM-dd' string here, server-side, before this ever crosses the
  // bridge. addin-sheets/sheets-format-bridge.js#cellToCanonicalText
  // already passes a plain string straight through, so this needs no
  // matching change on the client side.
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var values = range.getValues().map(function (row) {
    return row.map(function (cell) {
      return cell instanceof Date ? Utilities.formatDate(cell, tz, 'yyyy-MM-dd') : cell;
    });
  });
  var tValues = Date.now();

  var sampleRows = Math.min(FORMAT_SAMPLE_ROWS, range.getNumRows());
  var numberFormatSample = sampleRows > 0
    ? sheet.getRange(range.getRow(), range.getColumn(), sampleRows, range.getNumColumns()).getNumberFormats()
    : [];
  var tFormat = Date.now();

  return {
    values: values,
    numberFormatSample: numberFormatSample,
    totalRows: range.getNumRows(),
    cols: range.getNumColumns(),
    address: source.sheetName + '!' + source.a1,
    timing: { valuesMs: tValues - t0, formatMs: tFormat - tValues, totalMs: tFormat - t0 },
  };
}

function getWorkbookName() {
  return SpreadsheetApp.getActive().getName();
}

/* ---------- dashboard storage: hidden helper sheet, one row per dashboard ---------- */

function ensureDataSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HIDDEN_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['shapeName', 'json']]);
    sheet.hideSheet();
  }
  return sheet;
}

function findDashboardRow_(sheet, shapeName) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var names = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] === shapeName) return i + 2; // 1-indexed, +1 for the header row
  }
  return -1;
}

/** @param {object} payload plain object (addin/dashboard-io.js#buildStoragePayload's shape, built client-side) */
function saveDashboardSettings(shapeName, payload) {
  var sheet = ensureDataSheet_();
  var json = JSON.stringify(payload);
  var row = findDashboardRow_(sheet, shapeName);
  if (row === -1) sheet.appendRow([shapeName, json]);
  else sheet.getRange(row, 2).setValue(json);
}

/** @returns {Array<{shapeName:string, payload:object}>} */
function listDashboards() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    try {
      out.push({ shapeName: rows[i][0], payload: JSON.parse(rows[i][1]) });
    } catch (e) { /* a corrupted row shouldn't take the whole list down */ }
  }
  return out;
}

/** @returns {object|null} */
function loadDashboardSettings(shapeName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) return null;
  var row = findDashboardRow_(sheet, shapeName);
  if (row === -1) return null;
  var json = sheet.getRange(row, 2).getValue();
  try { return JSON.parse(json); } catch (e) { return null; }
}

function deleteDashboardSettings_(shapeName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(HIDDEN_SHEET_NAME);
  if (!sheet) return;
  var row = findDashboardRow_(sheet, shapeName);
  if (row !== -1) sheet.deleteRow(row);
}

/* ---------- dashboard image: Sheets has no shape ".name" the way Excel
   does, so this add-on uses OverGridImage#setAltTextTitle as the identity
   Excel's shape.name plays — set once at insert, read back later to find/
   replace/delete the right picture across every sheet. ---------- */

function findDashboardImage_(shapeName) {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var images = sheets[s].getImages();
    for (var i = 0; i < images.length; i++) {
      if (images[i].getAltTextTitle() === shapeName) return { image: images[i], sheet: sheets[s] };
    }
  }
  return null;
}

function deleteExistingImage_(shapeName) {
  var found = findDashboardImage_(shapeName);
  if (!found) return null;
  found.image.remove();
  return found.sheet;
}

function decodePng_(pngBase64) {
  return Utilities.newBlob(Utilities.base64Decode(pngBase64), 'image/png', shapeNameSafe_() + '.png');
}
function shapeNameSafe_() { return 'dashboard'; }

function addDashboardImage_(sheet, pngBase64, canvasSize, shapeName) {
  var blob = decodePng_(pngBase64);
  var image = sheet.insertImage(blob, 1, 1);
  image.setWidth(Math.round(canvasSize.width * IMAGE_SCALE));
  image.setHeight(Math.round(canvasSize.height * IMAGE_SCALE));
  image.setAltTextTitle(shapeName);
  return image;
}

/**
 * Places pngBase64 as a picture identified by `shapeName`. If one with
 * that identity already exists (any sheet) it's replaced on that SAME
 * sheet; a brand-new name goes on the active sheet — mirrors
 * addin/dashboard-io.js#placeDashboardImage exactly.
 */
function placeDashboardImage_(pngBase64, canvasSize, shapeName) {
  var existingSheet = deleteExistingImage_(shapeName);
  var targetSheet = existingSheet || SpreadsheetApp.getActiveSheet();
  addDashboardImage_(targetSheet, pngBase64, canvasSize, shapeName);
}

/** Never creates a picture — a no-op if `shapeName` has none yet (a draft). */
function replaceDashboardImageIfExists_(pngBase64, canvasSize, shapeName) {
  var existingSheet = deleteExistingImage_(shapeName);
  if (!existingSheet) return false;
  addDashboardImage_(existingSheet, pngBase64, canvasSize, shapeName);
  return true;
}

/**
 * Full "generate" step: place the image and save its settings, then read
 * the settings back to confirm the write actually stuck — same reasoning
 * as addin/dashboard-io.js#generateAndPlace's own verify-by-reading-back.
 * @param {{pngBase64:string, canvasSize:object, payload:object, shapeName?:string}} args `payload` already built client-side via addin/dashboard-io.js#buildStoragePayload
 * @returns {{shapeName:string}}
 */
function generateAndPlace(args) {
  var shapeName = args.shapeName || (SHAPE_NAME_PREFIX + Date.now());
  placeDashboardImage_(args.pngBase64, args.canvasSize, shapeName);
  saveDashboardSettings(shapeName, args.payload);
  var verify = loadDashboardSettings(shapeName);
  if (!verify) throw new Error('Settings for "' + shapeName + '" did not persist — the picture is on the sheet, but nothing was saved to reopen it from the list.');
  return { shapeName: shapeName };
}

/* ---------- chunked PNG upload: google.script.run's own bridge — not
   just the value types it can carry (see readRangeRaw's Date note above)
   — also has a practical size ceiling per call. A rasterized dashboard
   PNG as base64 can run into the multiple-MB range, comfortably past it,
   which showed up as "Place image on sheet" doing nothing: the call
   itself never completed, success or failure. dashboard.html's
   host.placeDashboard sends the base64 string as many small chunks
   instead of one huge argument — CacheService.getScriptCache() (100KB
   per value) buffers them here until generateAndPlaceFromUpload
   reassembles and hands off to generateAndPlace above, unchanged. ---------- */
var PNG_UPLOAD_TTL_SEC = 300;

function beginPngUpload() {
  return Utilities.getUuid();
}

function uploadPngChunk(uploadId, index, chunk) {
  CacheService.getScriptCache().put('pngUpload_' + uploadId + '_' + index, chunk, PNG_UPLOAD_TTL_SEC);
}

/** @param {{canvasSize:object, payload:object, shapeName?:string}} args same as generateAndPlace, minus pngBase64 */
function generateAndPlaceFromUpload(uploadId, totalChunks, args) {
  var cache = CacheService.getScriptCache();
  var parts = [];
  for (var i = 0; i < totalChunks; i++) {
    var key = 'pngUpload_' + uploadId + '_' + i;
    var part = cache.get(key);
    if (part == null) throw new Error('Upload incomplete (chunk ' + (i + 1) + ' of ' + totalChunks + ' missing or expired) — try Place image on sheet again.');
    parts.push(part);
    cache.remove(key);
  }
  return generateAndPlace({ pngBase64: parts.join(''), canvasSize: args.canvasSize, payload: args.payload, shapeName: args.shapeName });
}

/** No picture yet — see addin/dashboard-io.js#saveGeneratedDraft's doc comment for why Generate and Place stay separate actions. */
function saveGeneratedDraft(shapeName, payload) {
  saveDashboardSettings(shapeName, payload);
}

/**
 * Always updates the stored title (and, when given, the frozen
 * layoutSpec's own title); only touches the actual on-sheet picture when
 * one already exists under this name — mirrors
 * addin/dashboard-io.js#renameDashboard, including never resurrecting a
 * deleted picture on a plain rename.
 * @param {{layoutSpec?:object, pngBase64?:string, canvasSize?:object}} [snapshot]
 */
function renameDashboard(shapeName, newTitle, snapshot) {
  var existing = loadDashboardSettings(shapeName);
  if (!existing) throw new Error('Settings for "' + shapeName + '" not found — it may have been deleted.');
  existing.title = newTitle;
  if (snapshot && snapshot.layoutSpec) existing.layoutSpec = snapshot.layoutSpec;
  saveDashboardSettings(shapeName, existing);
  if (snapshot && snapshot.pngBase64 && snapshot.canvasSize) {
    replaceDashboardImageIfExists_(snapshot.pngBase64, snapshot.canvasSize, shapeName);
  }
}

/** Deletes a dashboard entirely: its picture (if any — a draft has none) and its settings row. */
function deleteDashboard(shapeName) {
  deleteExistingImage_(shapeName);
  deleteDashboardSettings_(shapeName);
}

/**
 * Dashboard-named pictures with no matching settings row — the signature
 * left behind when the spreadsheet was closed (or the tab reloaded)
 * without the settings write completing. Mirrors
 * addin/dashboard-io.js#countOrphanedDashboardShapes.
 */
function countOrphanedDashboardShapes() {
  var saved = {};
  listDashboards().forEach(function (d) { saved[d.shapeName] = true; });
  var count = 0;
  SpreadsheetApp.getActive().getSheets().forEach(function (sheet) {
    sheet.getImages().forEach(function (image) {
      var title = image.getAltTextTitle();
      if (title && title.indexOf(SHAPE_NAME_PREFIX) === 0 && !saved[title]) count++;
    });
  });
  return count;
}
