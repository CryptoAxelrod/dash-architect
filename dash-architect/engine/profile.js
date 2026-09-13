/*
 * Column profiling: turns a column of parsed cells (engine/values.js) into
 * the summary statistics that engine/roles.js classifies on.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineProfile = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const NUMERIC_KINDS = new Set(['date', 'percentage', 'currency', 'number']);
  const DIGITS_ONLY = /^\d+$/;

  function mean(values) {
    if (!values.length) return null;
    return values.reduce((a, b) => a + b, 0) / values.length;
  }

  function minOf(values) {
    let m = values[0];
    for (let i = 1; i < values.length; i++) if (values[i] < m) m = values[i];
    return m;
  }
  function maxOf(values) {
    let m = values[0];
    for (let i = 1; i < values.length; i++) if (values[i] > m) m = values[i];
    return m;
  }

  function monotonicity(orderedValues) {
    let increasing = true;
    let decreasing = true;
    let sawPair = false;
    let prev = null;
    for (const v of orderedValues) {
      if (v == null) continue;
      if (prev != null) {
        sawPair = true;
        if (v < prev) increasing = false;
        if (v > prev) decreasing = false;
      }
      prev = v;
    }
    if (!sawPair) return 'none';
    if (increasing && decreasing) return 'constant';
    if (increasing) return 'increasing';
    if (decreasing) return 'decreasing';
    return 'none';
  }

  /**
   * @param {string} name column header
   * @param {Array<{raw:string, empty:boolean, kind:string, numericValue:number|null, isInteger:boolean}>} cells one entry per row, in row order
   * @param {number} rowCount total number of data rows (== cells.length)
   * @returns {{profile: object, numericValues: Array<number|null>}}
   */
  function buildColumnProfile(name, cells, rowCount) {
    const nonEmpty = cells.filter((c) => !c.empty);
    const emptyCount = cells.length - nonEmpty.length;
    const emptyRatio = rowCount ? emptyCount / rowCount : 0;

    const numericValues = cells.map((c) => (!c.empty && NUMERIC_KINDS.has(c.kind) ? c.numericValue : null));

    let valueType = 'empty';
    let cellFormat = null;
    let looksNumericButTextFormatted = false;

    if (nonEmpty.length > 0) {
      const firstKind = nonEmpty[0].kind;
      const allSameKind = nonEmpty.every((c) => c.kind === firstKind);
      const kind = allSameKind ? firstKind : 'text';

      if (kind === 'date') {
        valueType = 'date';
        cellFormat = 'date';
      } else if (kind === 'percentage') {
        valueType = 'number';
        cellFormat = 'percentage';
      } else if (kind === 'currency') {
        valueType = 'number';
        cellFormat = 'currency';
      } else if (kind === 'number') {
        valueType = 'number';
        cellFormat = null;
      } else {
        valueType = 'text';
        cellFormat = null;
      }

      if (valueType === 'text') {
        // A column that is only ever digit strings (leading zeros included)
        // looks like it was deliberately formatted as text over a numeric
        // column — e.g. postal codes, padded IDs. See engine/values.js and
        // rule 1d in engine/roles.js.
        looksNumericButTextFormatted = nonEmpty.every((c) => DIGITS_ONLY.test(c.raw.trim()));
        if (looksNumericButTextFormatted) cellFormat = 'text-on-numeric';
      }
    }

    const numericSample = numericValues.filter((v) => v != null);
    const hasNumeric = numericSample.length > 0;

    const uniqueKeys = new Set(
      nonEmpty.map((c) => (NUMERIC_KINDS.has(c.kind) ? `n:${c.numericValue}` : `t:${c.raw.trim()}`))
    );
    const uniqueCount = uniqueKeys.size;
    const uniqueRatio = rowCount ? uniqueCount / rowCount : 0;

    const textLengths = nonEmpty.map((c) => c.raw.trim().length);

    const profile = {
      name,
      valueType, // 'number' | 'date' | 'text' | 'empty'
      cellFormat, // 'date' | 'percentage' | 'currency' | 'text-on-numeric' | null
      rowCount,
      uniqueCount,
      uniqueRatio,
      emptyCount,
      emptyRatio,
      // Not Math.min(...arr)/Math.max(...arr): spreading a large array as
      // call arguments overflows the stack well under the 200k-row target
      // (verified — it throws around ~120k in this engine's testing).
      min: hasNumeric ? minOf(numericSample) : null,
      max: hasNumeric ? maxOf(numericSample) : null,
      isInteger: hasNumeric ? numericSample.every((v) => Number.isInteger(v)) : null,
      mean: hasNumeric ? mean(numericSample) : null,
      monotonic: monotonicity(numericValues),
      avgTextLength: textLengths.length ? mean(textLengths) : null,
      looksNumericButTextFormatted,
    };

    return { profile, numericValues };
  }

  return { buildColumnProfile };
});
