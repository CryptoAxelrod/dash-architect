/*
 * Shared value formatting for both renderers (render/dom.js, render/svg.js)
 * so the same number reads identically in the interactive DOM and the
 * static SVG snapshot. Pure JS, no dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashRenderFormat = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function formatNumber(v, decimals) {
    if (v == null || Number.isNaN(v)) return '—';
    const d = decimals == null ? (Number.isInteger(v) ? 0 : 2) : decimals;
    return v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
  }

  /** Abbreviated form for axis labels/tight spaces: 1.2k, 3.4M. */
  function shortNumber(v) {
    if (v == null || Number.isNaN(v)) return '—';
    const sign = v < 0 ? '-' : '';
    const a = Math.abs(v);
    if (a >= 1e9) return sign + trimZero(a / 1e9) + 'B';
    if (a >= 1e6) return sign + trimZero(a / 1e6) + 'M';
    if (a >= 1e3) return sign + trimZero(a / 1e3) + 'k';
    return sign + trimZero(a);
  }

  function trimZero(v) {
    return (Math.round(v * 10) / 10).toString();
  }

  function formatPercent(v, decimals) {
    if (v == null || Number.isNaN(v)) return '—';
    return v.toFixed(decimals == null ? 1 : decimals) + '%';
  }

  /**
   * @param {{cellFormat:?string, aggregation:?string, valueScale:?string}} col a kpi/line/bar widget carries these fields directly
   * @param {number|null} value
   * @param {{negativeStyle?: 'color'|'parens', short?: boolean}} [opts]
   */
  function formatMeasureValue(col, value, opts) {
    opts = opts || {};
    if (value == null || Number.isNaN(value)) return '—';

    if (col.aggregation === 'weighted') {
      const pctValue = col.valueScale === 'percent100' ? value : value * 100;
      return wrapNegative(pctValue < 0, formatPercent(Math.abs(pctValue)), opts);
    }

    const numFn = opts.short ? shortNumber : (v) => formatNumber(v);
    if (col.cellFormat === 'currency') {
      return wrapNegative(value < 0, '$' + numFn(Math.abs(value)), opts);
    }
    return wrapNegative(value < 0, numFn(Math.abs(value)), opts);
  }

  function wrapNegative(isNegative, absText, opts) {
    if (!isNegative) return absText;
    if (opts.negativeStyle === 'parens') return `(${absText})`;
    return `−${absText}`; // U+2212 minus sign, matches reference/dashboard.html
  }

  function esc(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }

  // Shared by both renderers so a KPI value shrinks to fit its card the
  // same way whether it's drawn as an SVG <text> (render/svg.js, which also
  // has to run in Node, so no canvas/DOM text measurement there) or a real
  // DOM element (render/dom.js) — a character-count estimate calibrated for
  // tabular-nums digits/currency/percent strings, not a precise metric.
  function estimateTextWidth(str, fontSize) {
    return str.length * fontSize * 0.58;
  }
  function fitFontSize(str, maxWidth, baseSize, minSize) {
    let size = baseSize;
    while (size > minSize && estimateTextWidth(str, size) > maxWidth) size -= 1;
    return size;
  }

  return { formatNumber, shortNumber, formatPercent, formatMeasureValue, esc, estimateTextWidth, fitFontSize };
});
