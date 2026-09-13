/*
 * Reading a worksheet range into the shape engine/index.js#analyzeTable
 * expects. Two layers on purpose:
 *
 *  - `readRangeRaw` is the only function in this file that touches
 *    Office.js. It exists to get `values`/`valueTypes` (chunked, for large
 *    ranges) and a `numberFormat` *sample* (first CONFIG.formatSampleRows
 *    rows only — reading it for the full range is what times out on large
 *    sheets) off the Excel object model, nothing else.
 *  - `toEngineInput` is a plain function over plain arrays: no Office.js,
 *    fully testable in Node (tests/excel-io.test.js) by handing it a fake
 *    "raw" object shaped like readRangeRaw's return value.
 *
 * CLAUDE.md requires /addin to make no network calls and forbids Office.js
 * inside /engine — this file is the seam: everything Excel-flavored stays
 * on this side of it, and engine/index.js#analyzeTable never learns Excel
 * exists.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./format-bridge'));
  } else {
    root.DashAddinExcelIo = factory(root.DashAddinFormatBridge);
  }
})(typeof self !== 'undefined' ? self : this, function (FormatBridge) {
  'use strict';

  const CONFIG = {
    // Rows per values/valueTypes sync() round trip. Excel.js has no hard
    // documented ceiling here, but every round trip has fixed overhead and
    // every extra row adds two cells' worth of payload (values +
    // valueTypes) to marshal across the JS<->host boundary; this is a
    // starting point, not a measured optimum — see the timing note below.
    chunkRows: 5000,
    // Rows sampled to infer each column's number format. Reading
    // numberFormat for a 200,000-row range is what actually times out;
    // a column's format is realistically uniform, so a sample this size
    // is not a meaningful accuracy trade against reading everything.
    formatSampleRows: 200,
  };

  function now() {
    return typeof performance !== 'undefined' && performance.now ? performance.now() : Date.now();
  }

  /**
   * The only Office.js-touching function in this file. Reads `range` in
   * row chunks and returns plain arrays — no engine calls, no DOM.
   *
   * @param {Excel.RequestContext} context
   * @param {Excel.Range} range an already-obtained range (e.g. `workbook.getSelectedRange()`), not yet loaded
   * @returns {Promise<{values, valueTypes, numberFormatSample, totalRows, cols, address, timing}>}
   */
  async function readRangeRaw(context, range) {
    range.load('rowCount, columnCount, address');
    await context.sync();
    const totalRows = range.rowCount;
    const cols = range.columnCount;
    const t0 = now();

    const values = new Array(totalRows);
    const valueTypes = new Array(totalRows);
    for (let start = 0; start < totalRows; start += CONFIG.chunkRows) {
      const n = Math.min(CONFIG.chunkRows, totalRows - start);
      const chunk = range.getCell(start, 0).getResizedRange(n - 1, cols - 1);
      chunk.load('values, valueTypes');
      await context.sync();
      for (let i = 0; i < n; i++) {
        values[start + i] = chunk.values[i];
        valueTypes[start + i] = chunk.valueTypes[i];
      }
    }
    const tValues = now();

    const sampleRows = Math.min(CONFIG.formatSampleRows, totalRows);
    let numberFormatSample = [];
    if (sampleRows > 0) {
      const sampleRange = range.getCell(0, 0).getResizedRange(sampleRows - 1, cols - 1);
      sampleRange.load('numberFormat');
      await context.sync();
      numberFormatSample = sampleRange.numberFormat;
    }
    const tFormat = now();

    return {
      values,
      valueTypes,
      numberFormatSample,
      totalRows,
      cols,
      address: range.address,
      timing: { valuesMs: Math.round(tValues - t0), formatMs: Math.round(tFormat - tValues), totalMs: Math.round(tFormat - t0) },
    };
  }

  /**
   * Infers each column's dominant number-format kind from the sampled
   * rows. Takes the first non-null kind seen in the sample — a real
   * column's format is essentially always uniform top to bottom, so this
   * is a cheap, deterministic pick rather than a full vote.
   */
  function inferColumnFormats(numberFormatSample, cols) {
    const kinds = new Array(cols).fill(null);
    for (let c = 0; c < cols; c++) {
      for (let r = 0; r < numberFormatSample.length; r++) {
        const kind = FormatBridge.excelNumberFormatKind(numberFormatSample[r][c]);
        if (kind) {
          kinds[c] = kind;
          break;
        }
      }
    }
    return kinds;
  }

  /**
   * Pure: raw Office.js-shaped arrays -> {headers, dataRows} ready for
   * `engine.analyzeTable(headers, dataRows, ',')`. Row 0 is the header
   * row, matching every other entry point into the engine.
   *
   * @param {{values, valueTypes, numberFormatSample, totalRows, cols}} raw from readRangeRaw
   */
  function toEngineInput(raw) {
    const columnFormats = inferColumnFormats(raw.numberFormatSample, raw.cols);
    const headers = raw.values[0].map((v) => (v == null ? '' : String(v)));
    const dataRows = new Array(Math.max(0, raw.totalRows - 1));
    for (let r = 1; r < raw.totalRows; r++) {
      const row = new Array(raw.cols);
      for (let c = 0; c < raw.cols; c++) {
        row[c] = FormatBridge.cellToCanonicalText(raw.values[r][c], raw.valueTypes[r][c], columnFormats[c]);
      }
      dataRows[r - 1] = row;
    }
    return { headers, dataRows, columnFormats };
  }

  /**
   * Convenience wrapper: read + convert in one call.
   * @returns {Promise<{headers, dataRows, columnFormats, timing, address, totalRows, cols}>}
   */
  async function readRangeForEngine(context, range) {
    const raw = await readRangeRaw(context, range);
    const { headers, dataRows, columnFormats } = toEngineInput(raw);
    return { headers, dataRows, columnFormats, timing: raw.timing, address: raw.address, totalRows: raw.totalRows, cols: raw.cols };
  }

  return { CONFIG, readRangeRaw, inferColumnFormats, toEngineInput, readRangeForEngine };
});
