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

  // A column doesn't need every single cell to agree on kind to be typed as
  // that kind — a handful of genuinely unreadable cells (a stray typo, a
  // "N/A") shouldn't collapse an otherwise-clean date/number column to
  // text. Below the threshold, the column still falls back to 'text' as
  // before. Dates get a looser threshold than numbers/percentages/currency
  // because a free-text date column has more ways for one cell to look
  // like something else (a stray number, a note) than a numeric column
  // does. Deliberately NOT keyed off the column's header name/language
  // (unlike some prior-art tools) — a threshold that only kicks in for
  // English/Russian header keywords would just move the same gap
  // elsewhere for any other language.
  const DATE_MATCH_THRESHOLD = 0.9;
  const NUMERIC_MATCH_THRESHOLD = 0.95;

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

  // Direction plus step shape, in one pass: `maxAbsStep`/`stepCount` (over
  // consecutive non-equal values, row order, nulls skipped) exist so
  // engine/roles.js's identifier rule can tell "increments like a counter"
  // (small, bounded steps) apart from "happens to be sorted by this value"
  // (steps as large and irregular as the data itself) — direction alone
  // can't distinguish those two, and a table sorted by a metric is exactly
  // as monotonic as a real ID column.
  function monotonicity(orderedValues) {
    let increasing = true;
    let decreasing = true;
    let sawPair = false;
    let prev = null;
    let maxAbsStep = 0;
    let stepCount = 0;
    for (const v of orderedValues) {
      if (v == null) continue;
      if (prev != null) {
        sawPair = true;
        if (v < prev) increasing = false;
        if (v > prev) decreasing = false;
        const step = Math.abs(v - prev);
        if (step > 0) {
          stepCount++;
          if (step > maxAbsStep) maxAbsStep = step;
        }
      }
      prev = v;
    }
    let direction = 'none';
    if (sawPair) {
      if (increasing && decreasing) direction = 'constant';
      else if (increasing) direction = 'increasing';
      else if (decreasing) direction = 'decreasing';
    }
    return { direction, maxAbsStep, stepCount };
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
      const matchRatio = nonEmpty.filter((c) => c.kind === firstKind).length / nonEmpty.length;
      const threshold = firstKind === 'date' ? DATE_MATCH_THRESHOLD : NUMERIC_MATCH_THRESHOLD;
      const kind = NUMERIC_KINDS.has(firstKind) && matchRatio >= threshold ? firstKind : 'text';

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
    const monotonicInfo = monotonicity(numericValues);

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
      monotonic: monotonicInfo.direction,
      // Largest |step| between consecutive non-equal values in row order,
      // and how many such steps exist — see monotonicity()'s doc comment.
      // null/0 (not just 0) when there's no direction at all, so a caller
      // can tell "no monotonic run to measure" apart from "ran, step was 0."
      monotonicMaxStep: monotonicInfo.direction === 'none' ? null : monotonicInfo.maxAbsStep,
      monotonicStepCount: monotonicInfo.stepCount,
      avgTextLength: textLengths.length ? mean(textLengths) : null,
      looksNumericButTextFormatted,
    };

    return { profile, numericValues };
  }

  return { buildColumnProfile };
});
