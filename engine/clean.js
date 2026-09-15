/*
 * Cleans a raw headers/dataRows pair before it reaches engine/values.js —
 * dropping rows/columns that carry no real data and a leftover exported
 * index column, so a messy source doesn't have to be tidied by hand before
 * it produces a sensible dashboard. Runs on the raw string cells, before
 * any locale/type parsing.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineClean = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const UNNAMED_HEADER = /^unnamed/i;
  const PLAIN_INTEGER = /^-?\d+$/;

  function isBlankCell(cell) {
    return cell == null || cell.trim() === '';
  }

  /** Duplicate header text gets " (2)", " (3)", … — same convention as a duplicate dashboard title (addin/taskpane.js#computeDefaultTitle). */
  function uniqueHeaders(headers) {
    const seen = new Map();
    return headers.map((name) => {
      const count = seen.get(name) || 0;
      seen.set(name, count + 1);
      return count === 0 ? name : `${name} (${count + 1})`;
    });
  }

  function dropEmptyColumns(headers, dataRows) {
    const keepIndex = headers.map((_, colIndex) => dataRows.some((row) => !isBlankCell(row[colIndex])));
    const headersOut = headers.filter((_, i) => keepIndex[i]);
    const dataRowsOut = dataRows.map((row) => row.filter((_, i) => keepIndex[i]));
    return { headers: headersOut, dataRows: dataRowsOut };
  }

  function dropEmptyRows(dataRows) {
    return dataRows.filter((row) => !row.every(isBlankCell));
  }

  /**
   * A leftover index column from a pandas/Excel export: unnamed (or blank
   * header), every value a plain integer, strictly ascending, starting at 0
   * or 1, and dense (no huge gaps — `last - first < 2 * count` rules out an
   * ordinary ascending numeric column that just happens to start low).
   */
  function looksLikeExportedIndex(header, colIndex, dataRows) {
    if (!(header.trim() === '' || UNNAMED_HEADER.test(header.trim()))) return false;
    if (!dataRows.length) return false;

    const values = new Array(dataRows.length);
    let prev = -Infinity;
    for (let i = 0; i < dataRows.length; i++) {
      const raw = (dataRows[i][colIndex] || '').trim();
      if (!PLAIN_INTEGER.test(raw)) return false;
      const n = parseInt(raw, 10);
      if (n <= prev) return false;
      prev = n;
      values[i] = n;
    }
    const first = values[0];
    const last = values[values.length - 1];
    if (first !== 0 && first !== 1) return false;
    return last - first < 2 * values.length;
  }

  function dropExportedIndexColumns(headers, dataRows) {
    const keepIndex = headers.map((header, colIndex) => !looksLikeExportedIndex(header, colIndex, dataRows));
    const headersOut = headers.filter((_, i) => keepIndex[i]);
    const dataRowsOut = dataRows.map((row) => row.filter((_, i) => keepIndex[i]));
    return { headers: headersOut, dataRows: dataRowsOut };
  }

  /**
   * @param {string[]} headers
   * @param {string[][]} dataRows rows of raw string cells, aligned to `headers`
   * @returns {{headers:string[], dataRows:string[][]}}
   */
  function cleanTable(headers, dataRows) {
    let h = uniqueHeaders(headers);
    let rows = dataRows;

    ({ headers: h, dataRows: rows } = dropEmptyColumns(h, rows));
    rows = dropEmptyRows(rows);
    ({ headers: h, dataRows: rows } = dropExportedIndexColumns(h, rows));

    return { headers: h, dataRows: rows };
  }

  return { cleanTable };
});
