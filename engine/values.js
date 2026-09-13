/*
 * Per-cell value parsing and format inference.
 *
 * CSV text carries no real Excel number-format metadata, so this module
 * infers the closest equivalent from the text itself: a trailing "%" is
 * treated the same way an explicit percentage cell format would be (rule 1
 * in engine/roles.js), a leading zero on a digit string is treated as a
 * hint that the cell was explicitly formatted as text (rule 1d), and so on.
 * When real Excel format strings become available (reading .xlsx directly),
 * this inference layer is what gets replaced/extended — the rest of the
 * pipeline (profile.js, roles.js) only cares about the resulting
 * {kind, numericValue, ...} shape, not where it came from.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineValues = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const SPACE_CHARS = /[\s  ]/g; // plain space, nbsp, narrow nbsp (thousands separators)
  const MINUS_CHARS = /[−‐-―]/g; // unicode minus/dashes -> ascii '-'

  const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
  const EU_DATE = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/;
  const CURRENCY_SYMBOLS = '$€₽£';
  const CURRENCY_RE = new RegExp(
    `^([${CURRENCY_SYMBOLS}])\\s?([0-9.,\\s\\u00A0\\u202F−-]+)$|^([0-9.,\\s\\u00A0\\u202F−-]+)\\s?([${CURRENCY_SYMBOLS}])$`
  );
  const LEADING_ZERO_INT = /^0\d+$/;

  /** Number-locale style used to read grouping/decimal separators. */
  function localeStyleForDelimiter(delimiter) {
    // Files that use ';' as the field delimiter reserve ',' for decimals
    // (European convention) — see fixtures/07_ru_locale_text_numbers.csv.
    return delimiter === ';' ? 'euro' : 'us';
  }

  function normalizeMinus(s) {
    return s.replace(MINUS_CHARS, '-');
  }

  /** @returns {number} NaN if `raw` (with separators already stripped of spaces) isn't a plain number in this locale */
  function parsePlainNumber(raw, localeStyle) {
    let s = normalizeMinus(raw.trim());
    if (s === '') return NaN;
    if (localeStyle === 'euro') {
      s = s.replace(SPACE_CHARS, ''); // thousands separator
      if (!/^-?\d+(,\d+)?$/.test(s)) return NaN;
      s = s.replace(',', '.');
    } else {
      s = s.replace(SPACE_CHARS, '');
      // strip thousands commas: groups of exactly 3 digits after a comma
      if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
        s = s.replace(/,/g, '');
      } else if (!/^-?\d+(\.\d+)?$/.test(s)) {
        return NaN;
      }
    }
    return Number(s);
  }

  function tryDate(raw) {
    const s = raw.trim();
    let m = ISO_DATE.exec(s);
    if (m) {
      const [, y, mo, d] = m;
      const t = Date.UTC(+y, +mo - 1, +d);
      if (!Number.isNaN(t)) return t;
    }
    m = EU_DATE.exec(s);
    if (m) {
      const [, d, mo, y] = m;
      if (+mo >= 1 && +mo <= 12 && +d >= 1 && +d <= 31) {
        const t = Date.UTC(+y, +mo - 1, +d);
        if (!Number.isNaN(t)) return t;
      }
    }
    return null;
  }

  function tryPercentage(raw, localeStyle) {
    const s = raw.trim();
    if (!s.endsWith('%')) return null;
    const num = parsePlainNumber(s.slice(0, -1), localeStyle);
    if (Number.isNaN(num)) return null;
    return num / 100;
  }

  function tryCurrency(raw, localeStyle) {
    const m = CURRENCY_RE.exec(raw.trim());
    if (!m) return null;
    const numeric = m[2] !== undefined ? m[2] : m[3];
    const num = parsePlainNumber(numeric, localeStyle);
    if (Number.isNaN(num)) return null;
    return num;
  }

  /**
   * Parses one raw CSV cell into a typed value.
   * @returns {{raw:string, empty:boolean, kind:'empty'|'date'|'percentage'|'currency'|'number'|'text', numericValue:number|null, isInteger:boolean}}
   */
  function parseCell(raw, localeStyle) {
    const s = raw == null ? '' : String(raw);
    if (s.trim() === '') {
      return { raw: s, empty: true, kind: 'empty', numericValue: null, isInteger: false };
    }

    const dateValue = tryDate(s);
    if (dateValue !== null) {
      return { raw: s, empty: false, kind: 'date', numericValue: dateValue, isInteger: true };
    }

    const pct = tryPercentage(s, localeStyle);
    if (pct !== null) {
      return { raw: s, empty: false, kind: 'percentage', numericValue: pct, isInteger: Number.isInteger(pct * 100) };
    }

    const currency = tryCurrency(s, localeStyle);
    if (currency !== null) {
      return { raw: s, empty: false, kind: 'currency', numericValue: currency, isInteger: Number.isInteger(currency) };
    }

    // A leading zero on an otherwise-plain integer string ("00512") can't be
    // a real number written by a spreadsheet — it only survives if the cell
    // was formatted/typed as text. Treat it as text-on-a-numeric-looking
    // column (rule 1d), not as a number.
    if (!LEADING_ZERO_INT.test(s.trim())) {
      const plain = parsePlainNumber(s, localeStyle);
      if (!Number.isNaN(plain)) {
        return { raw: s, empty: false, kind: 'number', numericValue: plain, isInteger: Number.isInteger(plain) };
      }
    }

    return { raw: s, empty: false, kind: 'text', numericValue: null, isInteger: false };
  }

  return { localeStyleForDelimiter, parseCell, parsePlainNumber };
});
