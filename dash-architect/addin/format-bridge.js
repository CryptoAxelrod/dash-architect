/*
 * Pure translation between Excel's native cell representation
 * (value + valueType + numberFormat, as read via Office.js) and the
 * canonical text shape engine/index.js#analyzeTable already knows how to
 * read (the same shape a CSV row parses into). No Office.js/DOM here —
 * this file is a plain string/number transform and is fully testable in
 * Node without an Excel host (see tests/format-bridge.test.js).
 *
 * Why route through text instead of teaching the engine a second, richer
 * input format: the engine's classification rules are already written and
 * tested against that text-shaped input (engine/values.js) — turning an
 * Excel percent/currency/date cell into the same "48.5%" / "$1,234.56" /
 * "2025-01-04" text a well-formed CSV would contain lets addin/excel-io.js
 * hand the exact same array-of-arrays to engine.analyzeTable that the CSV
 * path already uses, with no new engine code and no new rules to keep in
 * sync with the CSV ones. See SPEC.md's addin section.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashAddinFormatBridge = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const CURRENCY_SYMBOLS = /[$€£¥₽₹]/;
  const BRACKETED = /\[[^\]]*\]/g; // Excel format "sections", e.g. [$€-407], [Red], [>=100]
  const QUOTED = /"[^"]*"/g; // literal text inside a format code, e.g. 0.00"m"
  const DATE_LETTERS = /[ymdhs]/i;

  /**
   * Classifies an Excel number-format string the way engine/roles.js rule 1
   * classifies a CSV column's inferred format — 'date' | 'percentage' |
   * 'currency' | null (no special format: plain number, or "General"/"@").
   * "@" is deliberately not special-cased here: a cell explicitly formatted
   * as Text already comes back from Office.js as a JS string in `values`,
   * so cellToCanonicalText's plain string passthrough already does the
   * right thing for it (including preserving leading zeros).
   */
  function excelNumberFormatKind(fmt) {
    if (!fmt || fmt === 'General' || fmt === '@') return null;
    const stripped = fmt.replace(BRACKETED, '').replace(QUOTED, '');
    if (stripped.indexOf('%') !== -1) return 'percentage';
    if (CURRENCY_SYMBOLS.test(fmt) || /\[\$[^\]]*\]/.test(fmt)) return 'currency';
    if (DATE_LETTERS.test(stripped)) return 'date';
    return null;
  }

  /**
   * Excel's date serial number (days since a fictional Dec 31 1899, with a
   * deliberate off-by-one for a fictitious Feb 29 1900, preserved for
   * backward compatibility with Lotus 1-2-3) to a UTC epoch in ms.
   *
   * Two known simplifications, acceptable for a thin adapter:
   *  - assumes the 1900 date system (the default; the legacy 1904 system,
   *    mostly old Mac-authored files, isn't detected — Office.js has no
   *    documented way to read a workbook's date system from an add-in).
   *  - drops time-of-day (a fractional serial is floored to its date) —
   *    engine/aggregate.js's time bucketing operates at day granularity or
   *    coarser anyway, so intraday precision isn't used downstream.
   */
  function excelSerialToEpochMs(serial) {
    return Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000;
  }

  function isoDateFromSerial(serial) {
    const d = new Date(excelSerialToEpochMs(serial));
    const pad2 = (n) => String(n).padStart(2, '0');
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  }

  /**
   * @param {*} value one cell from `range.values`
   * @param {string} valueType the matching entry from `range.valueTypes`
   * @param {'date'|'percentage'|'currency'|null} formatKind this column's inferred format (see addin/excel-io.js)
   * @returns {string} canonical text, ready for engine/values.js#parseCell
   */
  function cellToCanonicalText(value, valueType, formatKind) {
    if (valueType === 'Empty' || valueType === 'Error' || value == null) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value === 'number') {
      if (formatKind === 'date') return isoDateFromSerial(value);
      if (formatKind === 'percentage') return `${value * 100}%`;
      if (formatKind === 'currency') return `$${value}`;
      return String(value);
    }
    return String(value);
  }

  return { excelNumberFormatKind, excelSerialToEpochMs, isoDateFromSerial, cellToCanonicalText };
});
