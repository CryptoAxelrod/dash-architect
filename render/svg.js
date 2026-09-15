/*
 * Static SVG renderer: layoutSpec + theme -> one self-contained <svg> string.
 * Hand-rolled markup only (rects/paths/text) — no chart library, no
 * html2canvas-style screenshotting — so the output is pixel-predictable and
 * can be re-rendered identically for a PNG export. See CLAUDE.md.
 *
 * Pure string building: no DOM APIs, so this also runs in Node.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./format'));
  } else {
    root.DashRenderSvg = factory(root.DashRenderFormat);
  }
})(typeof self !== 'undefined' ? self : this, function (Format) {
  'use strict';

  const esc = Format.esc;

  function niceStep(x) {
    if (!(x > 0)) return 1;
    const e = Math.pow(10, Math.floor(Math.log10(x)));
    const f = x / e;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * e;
  }

  function valueFmt(widget) {
    return (v, opts) => Format.formatMeasureValue(widget, v, opts);
  }

  // --- tiny SVG builders ---------------------------------------------

  function rect(x, y, w, h, attrs) {
    return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(0, w).toFixed(1)}" height="${Math.max(0, h).toFixed(1)}"${attrStr(attrs)}/>`;
  }
  function line(x1, y1, x2, y2, attrs) {
    return `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}"${attrStr(attrs)}/>`;
  }
  function text(x, y, str, attrs) {
    return `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}"${attrStr(attrs)}>${esc(str)}</text>`;
  }
  function g(children, transform) {
    return `<g${transform ? ` transform="${transform}"` : ''}>${children}</g>`;
  }

  const fitFontSize = Format.fitFontSize;
  const truncateToWidth = Format.truncateToWidth;
  function attrStr(attrs) {
    if (!attrs) return '';
    return Object.entries(attrs)
      .filter(([, v]) => v != null && v !== false)
      .map(([k, v]) => ` ${k}="${v === true ? '' : esc(v)}"`)
      .join('');
  }

  // --- hatch pattern defs (print theme) --------------------------------

  const HATCH_DEFS = {
    diagonal: (id, color) => `<pattern id="${id}" width="6" height="6" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="6" stroke="${color}" stroke-width="2"/></pattern>`,
    diagonal2: (id, color) => `<pattern id="${id}" width="6" height="6" patternTransform="rotate(-45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="6" stroke="${color}" stroke-width="2"/></pattern>`,
    cross: (id, color) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="6" y2="6" stroke="${color}" stroke-width="1.3"/><line x1="6" y1="0" x2="0" y2="6" stroke="${color}" stroke-width="1.3"/></pattern>`,
    horizontal: (id, color) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><line x1="0" y1="3" x2="6" y2="3" stroke="${color}" stroke-width="2"/></pattern>`,
    vertical: (id, color) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><line x1="3" y1="0" x2="3" y2="6" stroke="${color}" stroke-width="2"/></pattern>`,
    dots: (id, color) => `<pattern id="${id}" width="6" height="6" patternUnits="userSpaceOnUse"><circle cx="1.5" cy="1.5" r="1.5" fill="${color}"/></pattern>`,
  };

  function collectHatchDefs(theme, count) {
    if (!theme.chart.textures) return { defs: '', fill: (i) => theme.chart.seriesColors[i % theme.chart.seriesColors.length] };
    const names = theme.chart.hatches || Object.keys(HATCH_DEFS);
    let defs = '';
    const ids = [];
    for (let i = 0; i < count; i++) {
      const name = names[i % names.length];
      const id = `hatch-${name}-${i}`;
      defs += HATCH_DEFS[name](id, theme.color.ink);
      ids.push(id);
    }
    return { defs: `<defs>${defs}</defs>`, fill: (i) => `url(#${ids[i]})` };
  }

  // --- markers (print theme point shapes) ------------------------------

  // Axis-aligned bounding box of an annular sector, for donut-slice hit
  // regions — a rectangle, not the wedge shape itself (that's the point of
  // asking for rectangles: cheap, uniform hit-testing across chart types,
  // at the cost of a little overlap between adjacent slim slices). The
  // outer arc's own bounding box always contains the inner arc's, since
  // both sweep the same angles at a smaller radius, so only the outer
  // radius needs sampling: the two endpoints plus any cardinal direction
  // (right/down/left/up) the sweep passes through, which is where an arc's
  // bounding box can extend past its own endpoints.
  function arcBoundingRect(cx, cy, outerR, a0, a1) {
    const norm = (a) => ((a % (2 * Math.PI)) + 2 * Math.PI) % (2 * Math.PI);
    const pts = [
      [cx + outerR * Math.cos(a0), cy + outerR * Math.sin(a0)],
      [cx + outerR * Math.cos(a1), cy + outerR * Math.sin(a1)],
    ];
    const na0 = norm(a0);
    const na1 = norm(a1);
    for (const a of [0, Math.PI / 2, Math.PI, (3 * Math.PI) / 2]) {
      const na = norm(a);
      const within = na0 <= na1 ? na >= na0 && na <= na1 : na >= na0 || na <= na1;
      if (within) pts.push([cx + outerR * Math.cos(a), cy + outerR * Math.sin(a)]);
    }
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
  }

  function markerPath(shape, cx, cy, size) {
    const r = size / 2;
    if (shape === 'square') return rect(cx - r, cy - r, r * 2, r * 2, {});
    if (shape === 'diamond') return `<path d="M ${cx} ${cy - r} L ${cx + r} ${cy} L ${cx} ${cy + r} L ${cx - r} ${cy} Z"/>`;
    if (shape === 'triangle') return `<path d="M ${cx} ${cy - r} L ${cx + r} ${cy + r} L ${cx - r} ${cy + r} Z"/>`;
    return `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="${r.toFixed(1)}"/>`;
  }

  // --- card chrome (charts & table only — KPIs are unboxed everywhere,
  // matching reference/dashboard.html, which never puts a border around a
  // KPI, only around its chart/table `.panel`) ---------------------------

  function panelChrome(rect_, theme, title) {
    const c = theme.color;
    let out = '';
    if (theme.card.style === 'panel') {
      out += rect(rect_.x, rect_.y, rect_.w, rect_.h, { fill: c.panel, stroke: c.panelBorder, 'stroke-width': theme.card.borderWidth, rx: theme.radius.lg });
    } else {
      out += line(rect_.x, rect_.y + rect_.h, rect_.x + rect_.w, rect_.y + rect_.h, { stroke: c.rule, 'stroke-width': theme.card.borderWidth });
    }
    if (title) {
      out += text(rect_.x + pad(theme), rect_.y + pad(theme) + theme.type.panelTitle * 0.8, title, {
        fill: c.ink, 'font-size': theme.type.panelTitle, 'font-weight': 650, 'font-family': theme.font.family,
      });
    }
    return out;
  }

  function pad(theme) {
    return theme.card.style === 'panel' ? theme.spacing.lg : theme.spacing.sm;
  }

  // --- header ------------------------------------------------------------

  function renderHeader(widget, theme) {
    const c = theme.color;
    const r = widget.rect;
    let out = text(r.x, r.y + theme.type.h1, widget.title || '', {
      fill: c.ink, 'font-size': theme.type.h1, 'font-weight': 700, 'font-family': theme.font.family,
    });
    if (widget.subtitle) {
      out += text(r.x, r.y + theme.type.h1 + theme.type.subtitle + 6, widget.subtitle, {
        fill: c.muted, 'font-size': theme.type.subtitle, 'font-family': theme.font.family,
      });
    }
    return g(out);
  }

  // --- filter bar (snapshot: shows current state, not interactive) ------

  function renderFilterBar(widget, theme) {
    const c = theme.color;
    const r = widget.rect;
    const h = 30;
    const cy = r.y + r.h / 2;
    let x = r.x;
    let out = '';

    for (const f of widget.filters) {
      const active = f.active || [];
      const label = active.length === 0 ? `${f.column}: All` : active.length === 1 ? `${f.column}: ${active[0]}` : `${f.column}: ${active.length} selected`;
      const w = Math.min(220, 16 + label.length * 6.4);
      const filled = active.length > 0;
      out += rect(x, cy - h / 2, w, h, { fill: filled ? c.accent : 'none', stroke: filled ? c.accent : c.rule, 'stroke-width': theme.card.borderWidth, rx: h / 2 });
      out += text(x + w / 2, cy + 4, label, { fill: filled ? c.panel : c.ink, 'font-size': 12.5, 'text-anchor': 'middle', 'font-family': theme.font.family });
      x += w + theme.spacing.sm;
    }
    if (widget.overflow && widget.overflow.length) {
      const label = `+${widget.overflow.length} more`;
      const w = 16 + label.length * 6.4;
      out += rect(x, cy - h / 2, w, h, { fill: 'none', stroke: c.rule, 'stroke-dasharray': '3 2', rx: h / 2 });
      out += text(x + w / 2, cy + 4, label, { fill: c.muted, 'font-size': 12.5, 'text-anchor': 'middle', 'font-family': theme.font.family });
    }
    return g(out);
  }

  // --- KPI cards (never boxed — see panelChrome comment) -----------------

  function renderKpi(widget, theme) {
    const c = theme.color;
    const r = widget.rect;
    const isHero = widget.variant === 'hero';
    const valueStr = Format.formatMeasureValue(widget, widget.value, { negativeStyle: theme.negativeStyle });
    const valueSize = fitFontSize(valueStr, r.w, isHero ? theme.type.kpiHero : theme.type.kpiValue, isHero ? 24 : 14);
    const negative = widget.value != null && widget.value < 0;

    // SVG <text> doesn't clip on its own (unlike render/dom.js's real DOM,
    // which gets overflow:hidden) — truncate the string itself so a long
    // measure name can't visually run into the next card.
    const label = truncateToWidth(widget.label, r.w, theme.type.kpiLabel);
    let out = text(r.x, r.y + theme.type.kpiLabel, label, { fill: c.muted, 'font-size': theme.type.kpiLabel, 'font-family': theme.font.family });
    out += text(r.x, r.y + theme.type.kpiLabel + valueSize * 0.95, valueStr, {
      fill: negative && theme.negativeStyle === 'color' ? c.negative : c.ink,
      'font-size': valueSize, 'font-weight': 750, 'font-family': theme.font.family,
    });
    const captionSize = theme.type.kpiLabel - 1.5;
    const caption = truncateToWidth(
      widget.approximate ? 'Unweighted average — no verified weight base' : `${widget.count.toLocaleString('en-US')} row${widget.count === 1 ? '' : 's'}`,
      r.w, captionSize
    );
    out += text(r.x, r.y + theme.type.kpiLabel + valueSize + 16, caption, {
      fill: c.faint, 'font-size': captionSize, 'font-family': theme.font.family,
    });
    if (!isHero) {
      // secondary cards are divided by a vertical rule, not boxed — matches
      // reference/dashboard.html's `.kpi-row>div+div{border-left}`
      out = line(r.x - 7, r.y - 2, r.x - 7, r.y + r.h - 2, { stroke: c.rule }) + out;
    }
    return g(out);
  }

  // Static counterpart of render/dom.js#renderEmptyState — same message,
  // no button (a flat picture can't open the mapping screen; the task pane
  // is where that lives). Shown instead of the KPI/chart block when
  // classification found no measure column at all — see
  // engine/layout.js#buildSkeleton.
  function renderEmptyState(widget, theme) {
    const c = theme.color;
    const r = widget.rect;
    let out = panelChrome(r, theme, null);
    const cx = r.x + r.w / 2;
    const midY = r.y + r.h / 2;
    out += text(cx, midY - 10, 'No measures to show', {
      fill: c.ink, 'font-size': theme.type.panelTitle, 'font-weight': 650, 'text-anchor': 'middle', 'font-family': theme.font.family,
    });
    out += text(cx, midY + 14, "None of this data's columns were classified as a measure —", {
      fill: c.muted, 'font-size': theme.type.subtitle, 'text-anchor': 'middle', 'font-family': theme.font.family,
    });
    out += text(cx, midY + 14 + theme.type.subtitle + 4, 'fix the column role on the mapping screen.', {
      fill: c.muted, 'font-size': theme.type.subtitle, 'text-anchor': 'middle', 'font-family': theme.font.family,
    });
    return g(out);
  }

  // --- charts --------------------------------------------------------

  function plotArea(rect_, theme, hasTitle) {
    const top = rect_.y + (hasTitle ? theme.type.panelTitle + pad(theme) * 1.8 : pad(theme));
    return { x: rect_.x + pad(theme) + 34, y: top, w: rect_.w - pad(theme) * 2 - 34, h: rect_.y + rect_.h - pad(theme) - 22 - top };
  }

  function yScale(values, plot) {
    const vals = values.filter((v) => v != null);
    let hi = Math.max(0, ...vals);
    let lo = Math.min(0, ...vals);
    if (hi === lo) hi = lo + 1;
    const step = niceStep((hi - lo) / 4);
    const top = Math.ceil(hi / step - 1e-9) * step;
    const bot = Math.floor(lo / step + 1e-9) * step;
    return { top, bot, step, y: (v) => plot.y + ((top - v) / (top - bot)) * plot.h };
  }

  function renderGrid(scale, plot, theme, tickFmt) {
    let out = '';
    for (let t = scale.bot; t <= scale.top + scale.step / 2; t += scale.step) {
      const yy = scale.y(t);
      const isZero = Math.abs(t) < scale.step / 1e6;
      out += line(plot.x, yy, plot.x + plot.w, yy, { stroke: isZero ? theme.chart.axisColor : theme.chart.gridColor, 'stroke-dasharray': isZero ? null : '2 3' });
      out += text(plot.x - 8, yy + 4, tickFmt(Math.abs(t) < scale.step / 1e6 ? 0 : t), {
        fill: theme.chart.axisColor, 'font-size': theme.type.axis, 'text-anchor': 'end', 'font-family': theme.font.family,
      });
    }
    return out;
  }

  function thinLabels(n, maxLabels) {
    const every = Math.max(1, Math.ceil(n / maxLabels));
    const show = new Set();
    for (let i = 0; i < n; i += every) show.add(i);
    show.add(n - 1);
    return show;
  }

  function renderLineChart(widget, theme) {
    const plot = plotArea(widget.rect, theme, true);
    const fmt = valueFmt(widget);
    const points = widget.points;
    const values = points.map((p) => p.value);
    const scale = yScale(values, plot);
    const tickFmt = widget.aggregation === 'weighted' ? (v) => Format.formatMeasureValue(widget, v, { short: true }) : (v) => Format.formatMeasureValue(widget, v, { short: true });

    let out = panelChrome(widget.rect, theme, widget.title);
    out += renderGrid(scale, plot, theme, tickFmt);

    const step = points.length > 1 ? plot.w / (points.length - 1) : 0;
    const xy = points.map((p, i) => [plot.x + step * i, p.value == null ? null : scale.y(p.value)]);

    // Hit regions: a full-height vertical lane per point (not just the dot)
    // — much easier to hover accurately than a few-pixel marker. Adjacent
    // lanes tile edge-to-edge; the two end lanes are half-width, clamped to
    // the plot bounds rather than overhanging it.
    //
    // dimension is deliberately always null here, not widget.timeColumn:
    // a point's `category` is a *bucket* label ("Jan 2024" — days/weeks/
    // months rolled up by engine/aggregate.js#bucketByTime), not a value
    // that appears anywhere in the time column itself (whose raw values
    // are per-row epoch milliseconds). Wiring it up as a filter dimension
    // like a bar/sector's would make render/dom.js send that label to
    // Aggregate.filterRowIndices, which compares it for equality against
    // the real per-row epoch values — never matches anything, silently
    // filtering the whole dashboard to zero rows. Hover/tooltip still work
    // (they don't need `dimension`); only the click-to-filter wiring
    // requires it, and time genuinely isn't click-filterable this way.
    const regions = [];
    const half = (step || plot.w) / 2;
    points.forEach((p, i) => {
      if (p.value == null) return;
      const rx = Math.max(plot.x, xy[i][0] - half);
      const rx2 = Math.min(plot.x + plot.w, xy[i][0] + half);
      regions.push({ rect: { x: rx, y: plot.y, w: rx2 - rx, h: plot.h }, value: p.value, category: p.label, dimension: null });
    });

    const segs = [];
    let cur = [];
    xy.forEach(([x, y]) => {
      if (y == null) {
        if (cur.length) segs.push(cur);
        cur = [];
      } else cur.push([x, y]);
    });
    if (cur.length) segs.push(cur);

    // The chart's own series color, not theme.color.accent — a palette
    // choice must actually change what a single-series chart looks like
    // (see render/palettes.js's doc comment on why this is a separate
    // token from the theme's brand/UI accent).
    const seriesColor = theme.chart.seriesColors[0];
    for (const seg of segs) {
      const d = seg.map(([x, y], i) => `${i === 0 ? 'M' : 'L'} ${x.toFixed(1)} ${y.toFixed(1)}`).join(' ');
      out += `<path d="${d}" fill="none" stroke="${seriesColor}" stroke-width="${theme.chart.lineWidth}" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
    if (theme.chart.markers) {
      const shape = theme.chart.markers[0];
      xy.forEach(([x, y]) => {
        if (y != null) out += `<g fill="${seriesColor}">${markerPath(shape, x, y, 6)}</g>`;
      });
    }

    const labelIdx = thinLabels(points.length, Math.max(2, Math.floor(plot.w / 56)));
    points.forEach((p, i) => {
      if (!labelIdx.has(i)) return;
      out += text(xy[i][0], widget.rect.y + widget.rect.h - 8, p.label, {
        fill: theme.chart.axisColor, 'font-size': theme.type.axis, 'text-anchor': 'middle', 'font-family': theme.font.family,
      });
    });

    return { markup: g(out), regions };
  }

  function renderBarChart(widget, theme, horizontal) {
    const plot = plotArea(widget.rect, theme, true);
    const bars = widget.bars;
    const values = bars.map((b) => b.value);
    let out = panelChrome(widget.rect, theme, widget.title);
    const regions = [];

    const { defs, fill } = collectHatchDefs(theme, 1);
    out = defs + out;

    // Palette color, not theme.color.accent — see renderLineChart's comment.
    // The non-peak bars get a version mixed 62% toward the paper color: the
    // old accent/accentSoft pair was two hand-picked shades of the same
    // theme green; a palette color has no such hand-authored soft twin, so
    // it's derived instead (Format.mixHex) rather than doubling every
    // palette's color count for a fill that's always "the series color,
    // muted" and never independently art-directed.
    const seriesColor = theme.chart.seriesColors[0];
    const softColor = Format.mixHex(seriesColor, theme.color.paper, 0.62);

    const peakIdx = values.reduce((best, v, i) => (v != null && (best < 0 || v > values[best]) ? i : best), -1);

    if (!horizontal) {
      const scale = yScale(values, plot);
      out += renderGrid(scale, plot, theme, (v) => Format.formatMeasureValue(widget, v, { short: true }));
      const band = plot.w / bars.length;
      const bw = Math.max(6, Math.min(band * 0.62, 56));
      bars.forEach((b, i) => {
        if (b.value == null) return;
        const cx = plot.x + band * i + band / 2;
        const y0 = scale.y(0);
        const y1 = scale.y(b.value);
        const top = Math.min(y0, y1);
        const h = Math.max(Math.abs(y0 - y1), 1.5);
        const isPeak = i === peakIdx;
        out += rect(cx - bw / 2, top, bw, h, { fill: isPeak ? seriesColor : theme.chart.textures ? fill(0) : softColor, stroke: theme.chart.textures ? theme.color.ink : null, 'stroke-width': theme.chart.textures ? 1 : null, rx: theme.chart.barRadius });
        if (isPeak) {
          out += text(cx, top - 6, Format.formatMeasureValue(widget, b.value, { short: true }), { fill: theme.color.ink, 'font-size': theme.type.axis, 'font-weight': 650, 'text-anchor': 'middle', 'font-family': theme.font.family });
        }
        regions.push({ rect: { x: cx - bw / 2, y: top, w: bw, h }, value: b.value, category: b.category, dimension: widget.dimensionColumn || null });
      });
      const labelIdx = thinLabels(bars.length, Math.max(2, Math.floor(plot.w / 70)));
      bars.forEach((b, i) => {
        if (!labelIdx.has(i)) return;
        const cx = plot.x + band * i + band / 2;
        out += text(cx, widget.rect.y + widget.rect.h - 8, truncate(b.category, 14), {
          fill: theme.chart.axisColor, 'font-size': theme.type.axis, 'text-anchor': 'middle', 'font-family': theme.font.family,
        });
      });
    } else {
      // horizontal: categories on Y, values on X
      const labelW = Math.min(plot.w * 0.32, 120);
      const barPlot = { x: plot.x + labelW, y: plot.y, w: plot.w - labelW, h: plot.h };
      const vals = values.filter((v) => v != null);
      const hi = Math.max(0, ...vals) || 1;
      const xScale = (v) => barPlot.x + (v / hi) * barPlot.w;
      const band = barPlot.h / bars.length;
      const bh = Math.max(8, Math.min(band * 0.6, 34));
      out += line(barPlot.x, plot.y, barPlot.x, plot.y + plot.h, { stroke: theme.chart.axisColor });
      bars.forEach((b, i) => {
        const cy = plot.y + band * i + band / 2;
        out += text(plot.x - 8 + labelW, cy + 4, truncate(b.category, 16), { fill: theme.color.ink, 'font-size': theme.type.axis, 'text-anchor': 'end', 'font-family': theme.font.family });
        if (b.value == null) return;
        const isPeak = i === peakIdx;
        const w = xScale(b.value) - barPlot.x;
        out += rect(barPlot.x, cy - bh / 2, w, bh, { fill: isPeak ? seriesColor : theme.chart.textures ? fill(0) : softColor, stroke: theme.chart.textures ? theme.color.ink : null, 'stroke-width': theme.chart.textures ? 1 : null, rx: theme.chart.barRadius });
        out += text(barPlot.x + w + 6, cy + 4, Format.formatMeasureValue(widget, b.value, { short: true }), { fill: theme.color.muted, 'font-size': theme.type.axis, 'font-family': theme.font.family });
        regions.push({ rect: { x: barPlot.x, y: cy - bh / 2, w, h: bh }, value: b.value, category: b.category, dimension: widget.dimensionColumn || null });
      });
    }

    return { markup: g(out), regions };
  }

  function truncate(s, n) {
    return s.length > n ? s.slice(0, n - 1) + '…' : s;
  }

  function renderGroupedBarChart(widget, theme) {
    const plot = plotArea(widget.rect, theme, true);
    const { categories, series } = widget;
    const allValues = series.flatMap((s) => s.values);
    const scale = yScale(allValues, plot);
    let out = panelChrome(widget.rect, theme, widget.title);
    const { defs, fill } = collectHatchDefs(theme, series.length);
    out = defs + out;
    out += renderGrid(scale, plot, theme, (v) => Format.formatMeasureValue(widget, v, { short: true }));
    const regions = [];

    const band = plot.w / categories.length;
    const groupW = band * 0.72;
    const barW = groupW / series.length;
    categories.forEach((cat, ci) => {
      const gx = plot.x + band * ci + (band - groupW) / 2;
      series.forEach((s, si) => {
        const v = s.values[ci];
        if (v == null) return;
        const y0 = scale.y(0);
        const y1 = scale.y(v);
        const top = Math.min(y0, y1);
        const h = Math.max(Math.abs(y0 - y1), 1.5);
        out += rect(gx + si * barW, top, barW - 2, h, {
          fill: theme.chart.textures ? fill(si) : theme.chart.seriesColors[si % theme.chart.seriesColors.length],
          stroke: theme.chart.textures ? theme.color.ink : null, 'stroke-width': theme.chart.textures ? 1 : null,
          rx: theme.chart.barRadius,
        });
        regions.push({ rect: { x: gx + si * barW, y: top, w: barW - 2, h }, value: v, category: cat, dimension: widget.dimensionColumn || null, series: s.name });
      });
      out += text(plot.x + band * ci + band / 2, widget.rect.y + widget.rect.h - 8, truncate(cat, 12), {
        fill: theme.chart.axisColor, 'font-size': theme.type.axis, 'text-anchor': 'middle', 'font-family': theme.font.family,
      });
    });

    out += renderLegend(series.map((s, i) => ({ label: s.name, swatch: theme.chart.textures ? fill(i) : theme.chart.seriesColors[i % theme.chart.seriesColors.length] })), widget.rect, theme);
    return { markup: g(out), regions };
  }

  // Right-aligned, on the title's own baseline — a left-aligned legend
  // would sit on top of the panel title, which starts at the same x.
  function renderLegend(items, rect_, theme) {
    const itemWidths = items.map((it) => 14 + it.label.length * 6.2 + 18);
    const totalW = itemWidths.reduce((a, b) => a + b, 0);
    let x = rect_.x + rect_.w - pad(theme) - totalW;
    const y = rect_.y + pad(theme) + theme.type.panelTitle * 0.8 - 5;
    let out = '';
    items.forEach((it, i) => {
      out += rect(x, y - 9, 10, 10, { fill: it.swatch, stroke: theme.chart.textures ? theme.color.ink : null, 'stroke-width': theme.chart.textures ? 1 : null });
      out += text(x + 16, y, it.label, { fill: theme.color.muted, 'font-size': theme.type.axis, 'font-family': theme.font.family });
      x += itemWidths[i];
    });
    return out;
  }

  function renderDonutChart(widget, theme) {
    const r = widget.rect;
    let out = panelChrome(r, theme, widget.title);
    const { defs, fill } = collectHatchDefs(theme, widget.slices.length);
    out = defs + out;

    const total = widget.slices.reduce((a, s) => a + (s.value || 0), 0) || 1;
    const cx = r.x + r.w * 0.32;
    const cy = r.y + r.h / 2 + 6;
    const outerR = Math.min(r.w * 0.24, r.h * 0.36);
    const innerR = outerR * 0.58;

    let angle = -Math.PI / 2;
    const legendItems = [];
    const regions = [];
    widget.slices.forEach((s, i) => {
      const frac = (s.value || 0) / total;
      const a0 = angle;
      const a1 = angle + frac * Math.PI * 2;
      angle = a1;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const p0o = [cx + outerR * Math.cos(a0), cy + outerR * Math.sin(a0)];
      const p1o = [cx + outerR * Math.cos(a1), cy + outerR * Math.sin(a1)];
      const p0i = [cx + innerR * Math.cos(a1), cy + innerR * Math.sin(a1)];
      const p1i = [cx + innerR * Math.cos(a0), cy + innerR * Math.sin(a0)];
      const d = `M ${p0o[0].toFixed(1)} ${p0o[1].toFixed(1)} A ${outerR.toFixed(1)} ${outerR.toFixed(1)} 0 ${large} 1 ${p1o[0].toFixed(1)} ${p1o[1].toFixed(1)} L ${p0i[0].toFixed(1)} ${p0i[1].toFixed(1)} A ${innerR.toFixed(1)} ${innerR.toFixed(1)} 0 ${large} 0 ${p1i[0].toFixed(1)} ${p1i[1].toFixed(1)} Z`;
      const swatch = theme.chart.textures ? fill(i) : theme.chart.seriesColors[i % theme.chart.seriesColors.length];
      out += `<path d="${d}" fill="${swatch}" stroke="${theme.color.panel === theme.color.paper ? theme.color.ink : theme.color.panel}" stroke-width="${theme.chart.textures ? 1.5 : 1}"/>`;
      legendItems.push({ label: `${s.label} (${Math.round(frac * 100)}%)`, swatch });
      regions.push({ rect: arcBoundingRect(cx, cy, outerR, a0, a1), value: s.value, category: s.label, dimension: widget.dimensionColumn || null });
    });

    out += text(cx, cy + 4, Format.shortNumber(total), { fill: theme.color.ink, 'font-size': theme.type.kpiValue * 0.7, 'font-weight': 700, 'text-anchor': 'middle', 'font-family': theme.font.family });

    let lx = r.x + r.w * 0.58;
    let ly = r.y + pad(theme) * 2.6;
    for (const it of legendItems) {
      out += rect(lx, ly - 9, 10, 10, { fill: it.swatch, stroke: theme.chart.textures ? theme.color.ink : null });
      out += text(lx + 16, ly, it.label, { fill: theme.color.muted, 'font-size': theme.type.axis, 'font-family': theme.font.family });
      ly += 20;
    }

    return { markup: g(out), regions };
  }

  /**
   * @returns {{markup:string, regions:Array<{rect:{x,y,w,h}, value:number|null, category:string, dimension:?string, series?:string}>}}
   *   `markup` is the same SVG fragment as before this method started
   *   returning an object — the PNG export path (renderDashboardSvg) uses
   *   only that field, so the exported picture is unaffected by hit
   *   regions. `regions` is for render/dom.js's interactive overlay.
   */
  function renderChart(widget, theme) {
    if (widget.type === 'line') return renderLineChart(widget, theme);
    if (widget.type === 'bar') return renderBarChart(widget, theme, false);
    if (widget.type === 'horizontalBar') return renderBarChart(widget, theme, true);
    if (widget.type === 'groupedBar') return renderGroupedBarChart(widget, theme);
    if (widget.type === 'donut') return renderDonutChart(widget, theme);
    return { markup: '', regions: [] };
  }

  // --- table ------------------------------------------------------------

  function formatCell(col, value, theme) {
    if (value == null) return '—';
    // See render/dom.js's formatCell — same year_like exception, kept in
    // sync since this is the separate static/export renderer.
    if (col.granularity === 'year') return String(value);
    if (col.role === 'time' || col.cellFormat === 'date') {
      const d = new Date(value);
      const MONTHS_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${MONTHS_[d.getUTCMonth()]} ${d.getUTCDate()}, ${d.getUTCFullYear()}`;
    }
    if (col.role === 'measure') return Format.formatMeasureValue(col, value, { negativeStyle: theme.negativeStyle });
    return String(value);
  }

  // A table widget is either "live" (`rowIndices` into rows the caller
  // resolves through `dataColumns`) or "frozen" (`rows`: literal values
  // already baked in by engine/layout.js#buildStorageSnapshot, no external
  // lookup needed or possible). Both renderers accept either shape.
  function tableRowValues(widget, dataColumns, pageSize) {
    if (widget.rows) return { rows: widget.rows, total: widget.shownRows };
    const idx = widget.rowIndices.slice(0, pageSize);
    return { rows: idx.map((i) => widget.columns.map((col) => dataColumns.get(col.name).values[i])), total: widget.totalRows };
  }

  function renderTable(widget, theme, dataColumns) {
    const r = widget.rect;
    const c = theme.color;
    let out = panelChrome(r, theme, null);
    const innerX = r.x + pad(theme);
    const innerW = r.w - pad(theme) * 2;
    const headerY = r.y + pad(theme);
    const rowH = theme.table.rowHeight;
    const visibleRows = Math.max(1, Math.floor((r.h - pad(theme) * 2 - theme.table.headerSize - 12 - 26) / rowH));

    const cols = widget.columns;
    const { rows: pageRows, total } = tableRowValues(widget, dataColumns, visibleRows);
    // Same per-column "autofit" as render/dom.js's live table (see
    // Format.computeColumnWidths) instead of splitting the width evenly —
    // a static picture needs this even more, since there's no hover here to
    // fall back on for whatever gets truncated.
    const cellStrings = pageRows.map((rowValues) => cols.map((col, ci) => formatCell(col, rowValues[ci], theme)));
    const colWidths = Format.computeColumnWidths(
      cols.map((col, ci) => ({ header: col.name, samples: cellStrings.map((row) => row[ci]) })),
      innerW, { fontSize: theme.table.cellSize }
    );
    const colX = [];
    for (let i = 0, acc = innerX; i < colWidths.length; i++) { colX.push(acc); acc += colWidths[i]; }

    cols.forEach((col, i) => {
      const x = colX[i];
      const cw = colWidths[i];
      const isNum = col.role === 'measure';
      const label = truncateToWidth(col.name + (widget.sort.key === col.name ? (widget.sort.dir === 'asc' ? ' ↑' : ' ↓') : ''), cw, theme.table.headerSize);
      out += text(isNum ? x + cw : x, headerY + theme.table.headerSize, label, {
        fill: c.muted, 'font-size': theme.table.headerSize, 'font-weight': 650, 'text-anchor': isNum ? 'end' : 'start', 'font-family': theme.font.family,
      });
    });
    out += line(innerX, headerY + theme.table.headerSize + 8, innerX + innerW, headerY + theme.table.headerSize + 8, { stroke: c.rule });

    let y = headerY + theme.table.headerSize + 8;
    pageRows.forEach((rowValues, ri) => {
      y += rowH;
      cols.forEach((col, ci) => {
        const x = colX[ci];
        const cw = colWidths[ci];
        const isNum = col.role === 'measure';
        const raw = rowValues[ci];
        const str = truncateToWidth(cellStrings[ri][ci], cw, theme.table.cellSize);
        const isNegative = isNum && raw != null && raw < 0 && theme.negativeStyle === 'color';
        out += text(isNum ? x + cw : x, y - rowH / 2 + 5, str, {
          fill: isNegative ? c.negative : c.ink, 'font-size': theme.table.cellSize, 'text-anchor': isNum ? 'end' : 'start', 'font-family': theme.font.family,
        });
      });
      if (ri < pageRows.length - 1) out += line(innerX, y, innerX + innerW, y, { stroke: c.ruleSoft });
    });

    out += text(innerX, r.y + r.h - pad(theme), `${pageRows.length} of ${total.toLocaleString('en-US')} rows`, {
      fill: c.faint, 'font-size': theme.table.headerSize, 'font-family': theme.font.family,
    });

    return g(out);
  }

  // --- top level ----------------------------------------------------

  function renderWidget(widget, theme, dataColumns) {
    if (widget.type === 'header') return renderHeader(widget, theme);
    if (widget.type === 'filterBar') return renderFilterBar(widget, theme);
    if (widget.type === 'kpi') return renderKpi(widget, theme);
    if (widget.type === 'table') return renderTable(widget, theme, dataColumns);
    if (widget.type === 'emptyState') return renderEmptyState(widget, theme);
    return renderChart(widget, theme).markup; // hit regions are dom.js's concern, not the flattened PNG's
  }

  /**
   * @param {{canvas:{width,height}, widgets:Array}} layoutSpec
   * @param {object} theme one of render/themes.js THEMES
   * @param {Map<string,object>} [dataColumns] name -> column (for the table's raw cell values); required if the spec has a table widget
   */
  function renderDashboardSvg(layoutSpec, theme, dataColumns) {
    const { width, height } = layoutSpec.canvas;
    const c = theme.color;
    // Every widget renderer above draws in absolute canvas coordinates
    // (reading widget.rect.x/y directly, since panelChrome/plotArea etc.
    // need the rect for gridlines and text baselines too) — so widgets are
    // concatenated as-is, with no additional positioning wrapper.
    let body = rect(0, 0, width, height, { fill: c.paper });
    for (const w of layoutSpec.widgets) body += renderWidget(w, theme, dataColumns);
    // theme.font.family is a CSS font stack containing literal double quotes
    // ("Segoe UI", ...) — must go through esc() or it breaks this attribute
    // as invalid XML. innerHTML's lenient HTML parser hides that; loading
    // the string as a standalone image/svg+xml document (the PNG export
    // path, render/share.js) does not — it just fails to load.
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" font-family="${esc(theme.font.family)}">${body}</svg>`;
  }

  return { renderDashboardSvg, renderChart, niceStep, truncate };
});
