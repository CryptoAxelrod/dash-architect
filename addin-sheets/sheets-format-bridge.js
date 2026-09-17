/*
 * Pure translation between Google Sheets' native cell representation
 * (value + numberFormat pattern, as read via SpreadsheetApp/Apps Script)
 * and the canonical text shape engine/index.js#analyzeTable already knows
 * how to read — the exact same text shape addin/format-bridge.js produces
 * for Excel, so nothing downstream (engine/values.js, engine/roles.js) has
 * to learn a second input format. See addin/format-bridge.js's own doc
 * comment for why routing through text is the right seam.
 *
 * This is a separate module from addin/format-bridge.js, not a thin
 * wrapper around it, for one concrete reason: Sheets' own generated
 * currency pattern quotes the symbol itself (e.g. `"$"#,##0.00`), where
 * Excel's does not (`$#,##0.00`, or `[$€-407] #,##0.00`). Excel's
 * classifier strips quoted literals *before* checking for a currency
 * symbol (correctly, for its own decorative-suffix case like `0.00"m"`) —
 * applied to a Sheets currency pattern, that would strip the quoted `"$"`
 * first and misclassify it as a plain number. So the quoted-currency check
 * has to run *before* the generic quote-strip here, which is the one rule
 * genuinely different from Excel's, not a full reimplementation.
 *
 * No Office.js, no Apps Script globals (SpreadsheetApp, etc.) — plain
 * string/number transforms, loadable in Node (tests), in a browser (the
 * Sheets add-on's own dialog page), and inside an Apps Script project's
 * server-side V8 runtime alike, via the same UMD wrapper the rest of this
 * codebase already uses (addin/format-bridge.js, addin/dashboard-io.js).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashSheetsFormatBridge = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const CURRENCY_SYMBOLS = /[$€£¥₽₹]/;
  const QUOTED_CURRENCY = /"[^"]*[$€£¥₽₹][^"]*"/;
  const BRACKETED = /\[[^\]]*\]/g; // conditional/color sections, e.g. [RED], [>=1000000]
  const QUOTED = /"[^"]*"/g; // literal text inside a format code, e.g. 0.00"m", or a quoted currency symbol
  const DATE_LETTERS = /[ymdhs]/i;

  // Formats Sheets itself uses to mean "nothing special" — a plain number
  // or explicit plain text, same role Excel's 'General'/'@' play.
  const PLAIN_FORMATS = new Set(['', 'General number', '@', '@STRING@']);

  /**
   * Classifies a Sheets number-format pattern the same way
   * addin/format-bridge.js#excelNumberFormatKind classifies an Excel one —
   * 'date' | 'percentage' | 'currency' | null. See the module doc comment
   * above for the one rule ordering difference (quoted currency symbol
   * checked before the generic quote-strip).
   */
  function sheetsNumberFormatKind(fmt) {
    if (!fmt || PLAIN_FORMATS.has(fmt)) return null;
    if (QUOTED_CURRENCY.test(fmt)) return 'currency';
    const stripped = fmt.replace(BRACKETED, '').replace(QUOTED, '');
    if (stripped.indexOf('%') !== -1) return 'percentage';
    if (CURRENCY_SYMBOLS.test(stripped)) return 'currency';
    if (DATE_LETTERS.test(stripped)) return 'date';
    return null;
  }

  /**
   * A date-formatted cell read via Range#getValues() comes back as a real
   * JS Date object (Apps Script's own marshaling), not a serial number the
   * way Excel's Office.js does — so there is no serial-number math to do
   * here, unlike addin/format-bridge.js#excelSerialToEpochMs. The one
   * documented simplification carried over from the Excel side is the
   * same one: time-of-day is dropped, since engine/aggregate.js's time
   * bucketing never uses better than day granularity anyway. Uses the
   * Date's local getters (the timezone Apps Script/the browser already
   * resolved it into when marshaling the value across the client/server
   * boundary — see addin-sheets/sheets-io.js), not UTC getters, since
   * there is no serial number here to anchor a UTC calculation to.
   */
  function isoDateFromLocalDate(date) {
    const pad2 = (n) => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
  }

  // Sheets' own formula-error strings — a cell showing one of these reads
  // back from getValues() as this literal string, not a thrown exception.
  // Treated as blank, the same way Excel's 'Error' valueType is (see
  // addin/format-bridge.js#cellToCanonicalText) rather than as literal
  // garbage text fed to the engine.
  const ERROR_TEXT = /^#(REF|DIV\/0|VALUE|NAME|NULL|NUM|N\/A|ERROR)!?\??$/;

  /**
   * @param {*} value one cell from Range#getValues() — string, number,
   *   boolean, Date, or '' for a blank cell.
   * @param {'date'|'percentage'|'currency'|null} formatKind this column's inferred format (see addin-sheets/sheets-io.js)
   * @returns {string} canonical text, ready for engine/values.js#parseCell
   */
  function cellToCanonicalText(value, formatKind) {
    if (value == null || value === '') return '';
    if (value instanceof Date) return isoDateFromLocalDate(value);
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value === 'string') return ERROR_TEXT.test(value) ? '' : value;
    if (typeof value === 'number') {
      if (formatKind === 'percentage') return `${value * 100}%`;
      if (formatKind === 'currency') return `$${value}`;
      return String(value);
    }
    return String(value);
  }

  return { sheetsNumberFormatKind, isoDateFromLocalDate, cellToCanonicalText };
});
