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

  // Lightens/darkens a hex color by mixing it toward another (e.g. the
  // theme's own paper color) — used for a chart's "muted" fill (every bar
  // but the highlighted one) from a palette color that, unlike the old
  // per-theme accent/accentSoft pair, has no hand-authored soft variant of
  // its own. `ratio` is how much of `towardHex` to mix in (0 = unchanged, 1 = towardHex).
  function mixHex(hex, towardHex, ratio) {
    const a = parseInt(hex.slice(1), 16);
    const b = parseInt(towardHex.slice(1), 16);
    const mix = (shift) => {
      const av = (a >> shift) & 0xff;
      const bv = (b >> shift) & 0xff;
      return Math.round(av + (bv - av) * ratio);
    };
    const r = mix(16), g = mix(8), bch = mix(0);
    return `#${[r, g, bch].map((v) => v.toString(16).padStart(2, '0')).join('')}`;
  }

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

  // SVG <text> has no CSS text-overflow — used by render/svg.js's KPI label
  // (render/dom.js gets the same effect from plain overflow:hidden +
  // text-overflow:ellipsis instead, since it's real DOM). Same
  // estimateTextWidth heuristic as fitFontSize, for the same reason: this
  // has to run in Node too, no canvas/DOM text measurement available.
  function truncateToWidth(str, maxWidth, fontSize) {
    if (estimateTextWidth(str, fontSize) <= maxWidth) return str;
    let end = str.length;
    while (end > 0 && estimateTextWidth(str.slice(0, end) + '…', fontSize) > maxWidth) end -= 1;
    return end > 0 ? str.slice(0, end) + '…' : '…';
  }

  /**
   * Per-column table width, "autofit"-style: a column sized for its own
   * typical content (header + a sample of formatted values — the current
   * page is plenty, values in one column are rarely wildly different
   * lengths) rather than every column splitting the width evenly or every
   * column being squeezed down to fit no matter how many there are. A
   * narrow column (a handful of short codes, say) stays narrow — floored at
   * minChars so it's never *less* readable than that — and a column with
   * genuinely long values caps out at maxChars rather than growing forever.
   * If the columns' natural widths add up to more than `totalWidth`, this
   * does NOT shrink them to fit — the table is meant to scroll horizontally
   * for the overflow (render/dom.js wraps it in its own scroll container)
   * rather than cram every column into an unreadable sliver. Only when
   * there's room to *spare* does the extra get handed out, so a table with
   * few columns still ends up using the full available width.
   * @param {Array<{header:string, samples:string[]}>} columns
   * @param {number} totalWidth
   * @param {{fontSize?:number, minChars?:number, maxChars?:number, cellPadding?:number}} [opts]
   * @returns {number[]} one width per column
   */
  function computeColumnWidths(columns, totalWidth, opts) {
    opts = opts || {};
    const fontSize = opts.fontSize || 13;
    const cellPadding = opts.cellPadding != null ? opts.cellPadding : 24;
    const minWidth = estimateTextWidth('x'.repeat(opts.minChars || 5), fontSize) + cellPadding;
    const maxWidth = estimateTextWidth('x'.repeat(opts.maxChars || 18), fontSize) + cellPadding;
    if (!columns.length) return [];

    const natural = columns.map((col) => {
      const longest = [col.header, ...col.samples].reduce((a, s) => Math.max(a, (s || '').length), 0);
      return estimateTextWidth('x'.repeat(longest), fontSize) + cellPadding;
    });
    const widths = natural.map((w) => Math.min(maxWidth, Math.max(minWidth, w)));
    const sum = widths.reduce((a, b) => a + b, 0);

    // More columns than comfortably fit — let the table scroll horizontally
    // for the rest instead of shrinking every column below its own natural
    // width (that was the previous behavior, and it truncated far more
    // aggressively than any single column's own content actually needed).
    if (sum >= totalWidth) return widths;

    // Room to spare: water-fill it onto columns whose OWN content actually
    // wants more than the floor (natural[i] > minWidth) — a column that's
    // narrow because its content genuinely is narrow (an ID column of
    // 2-digit numbers, say) stays exactly that narrow instead of being
    // stretched just because a neighbor needs the room. This is also what
    // makes a table with few (but substantial) columns end up using the
    // full available width: each grows up to maxChars rather than stopping
    // at its own shorter sample.
    let remaining = totalWidth - sum;
    let growable = widths.map((_, i) => i).filter((i) => natural[i] > minWidth && widths[i] < maxWidth);
    while (remaining > 0.5 && growable.length) {
      const share = remaining / growable.length;
      let used = 0;
      growable = growable.filter((i) => {
        const add = Math.min(maxWidth - widths[i], share);
        widths[i] += add;
        used += add;
        return widths[i] < maxWidth - 0.01;
      });
      if (used < 0.5) break;
      remaining -= used;
    }
    return widths;
  }

  return { formatNumber, shortNumber, formatPercent, formatMeasureValue, esc, estimateTextWidth, fitFontSize, truncateToWidth, computeColumnWidths, mixHex };
});
