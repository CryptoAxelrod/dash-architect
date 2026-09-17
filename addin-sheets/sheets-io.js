/*
 * Reading a Sheets range into the shape engine/index.js#analyzeTable
 * expects — the Google Sheets analog of addin/excel-io.js. Same split for
 * the same reason (CLAUDE.md forbids Apps Script/Sheets awareness inside
 * /engine, and this is the seam): `toEngineInput` is a plain function over
 * plain arrays, fully testable in Node (tests/sheets-io.test.js).
 *
 * Unlike excel-io.js there is no chunked-read half in this file at all:
 * that chunking exists in the Excel version to keep each Office.js
 * context.sync() round trip small, a real, documented constraint of the
 * Excel JS API's marshaling. Reading a Sheets range is a single
 * Range#getValues()/getNumberFormats() call from Apps Script server code
 * (addin-sheets/Code.gs#readRangeRaw) — there is no equivalent per-call
 * overhead to chunk against. That server call already returns exactly the
 * `{values, numberFormats, totalRows, cols, address, timing}` shape this
 * file's `toEngineInput` consumes, so nothing here ever touches
 * SpreadsheetApp.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./sheets-format-bridge'));
  } else {
    root.DashSheetsIo = factory(root.DashSheetsFormatBridge);
  }
})(typeof self !== 'undefined' ? self : this, function (FormatBridge) {
  'use strict';

  const CONFIG = {
    // Rows sampled to infer each column's number format — mirrors
    // addin/excel-io.js's CONFIG.formatSampleRows for the same reason: a
    // column's format is realistically uniform, and reading it for every
    // row of a very large range is needless work.
    formatSampleRows: 200,
  };

  /**
   * Infers each column's dominant number-format kind from the sampled
   * rows — same "first non-null kind wins" approach as
   * addin/excel-io.js#inferColumnFormats.
   * @param {Array<Array<string>>} numberFormatSample rows x cols of Sheets number-format patterns
   */
  function inferColumnFormats(numberFormatSample, cols) {
    const kinds = new Array(cols).fill(null);
    for (let c = 0; c < cols; c++) {
      for (let r = 0; r < numberFormatSample.length; r++) {
        const kind = FormatBridge.sheetsNumberFormatKind(numberFormatSample[r][c]);
        if (kind) {
          kinds[c] = kind;
          break;
        }
      }
    }
    return kinds;
  }

  /**
   * Pure: raw Sheets-shaped arrays -> {headers, dataRows} ready for
   * `engine.analyzeTable(headers, dataRows, ',')`. Row 0 is the header
   * row, matching every other entry point into the engine.
   *
   * @param {{values, numberFormatSample, totalRows, cols}} raw from Code.gs#readRangeRaw
   */
  function toEngineInput(raw) {
    const columnFormats = inferColumnFormats(raw.numberFormatSample, raw.cols);
    const headers = raw.values[0].map((v) => (v == null ? '' : String(v)));
    const dataRows = new Array(Math.max(0, raw.totalRows - 1));
    for (let r = 1; r < raw.totalRows; r++) {
      const row = new Array(raw.cols);
      for (let c = 0; c < raw.cols; c++) {
        row[c] = FormatBridge.cellToCanonicalText(raw.values[r][c], columnFormats[c]);
      }
      dataRows[r - 1] = row;
    }
    return { headers, dataRows, columnFormats };
  }

  return { CONFIG, inferColumnFormats, toEngineInput };
});
