/*
 * Layout engine: turns classified columns (engine/roles.js) into a
 * `layoutSpec` — a flat list of widgets with fixed pixel rects, computed
 * once and shared by both renderers (render/dom.js, render/svg.js). See
 * SPEC.md §10-11: composition and widget position/size are the same across
 * every theme — only paint (color/type/spacing tokens) differs per theme,
 * so none of that lives here.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./aggregate'));
  } else {
    root.DashEngineLayout = factory(root.DashEngineAggregate);
  }
})(typeof self !== 'undefined' ? self : this, function (Aggregate) {
  'use strict';

  const LAYOUT = {
    width: 1180, // default canvasWidth (widgetConfig.canvasWidth) — the static export always uses this unless the live view was a different width when placed; see buildSkeleton
    margin: { top: 32, right: 28, bottom: 40, left: 28 },
    gap: 28, // vertical gap between major sections
    header: { height: 60 },
    // One chip-row per filterable dimension (render/dom.js#renderFilterBar) —
    // sized for the worst case (maxVisible rows) so it never overlaps the
    // widget below it; most datasets show fewer than maxVisible and just
    // leave the rest of the box empty, same tradeoff every other fixed
    // heuristic constant in this file makes. maxVisible dropped from 5 to 4
    // (2026-09) — 5 rows routinely overflowed this fixed height; the rest
    // are reachable via widgetConfig.enabledDimensions (settings panel's
    // new "Filters" section) instead of just disappearing.
    filterBar: { height: 150, maxVisible: 4 },
    // Reserves this many px at the right edge of whatever row it's anchored
    // to (the filter bar's, or the header's if there's no filter bar) —
    // see the `showWatermark` block in buildSkeleton. A fixed reserved
    // width (not "measure the text and hug it") means this file never
    // needs to know actual rendered text width, which depends on font
    // metrics only /render knows. Sized for "Made with Dash Architect" at
    // 1.5x theme.type.axis in the largest theme (light/dark, axis 11.5) —
    // Format.estimateTextWidth puts that around 240px; 260 leaves margin.
    watermark: { width: 260 },
    kpi: {
      // Dropped from 6 to 4 (2026-09) — a long measure name in a 5th/6th
      // secondary card had too little width to avoid overlapping its
      // neighbor even with the truncation added alongside this (render/dom.js
      // and render/svg.js's renderKpi) — 1 hero + 3 secondary is what
      // LAYOUT.overview.kpiFraction's column width actually accommodates
      // cleanly. The rest are reachable via widgetConfig.enabledKpis.
      maxCards: 4,
      heroHeight: 150,
      secondaryHeight: 108,
      innerGap: 16, // between hero and the secondary row
      cardGap: 14, // between secondary cards
    },
    overview: { gap: 32, kpiFraction: 5 / 12 },
    // maxCount dropped from 4 to 3 (2026-09) — with 4, a remainder of 3
    // charts below the overview row forced a 2-column grid whose last row
    // held a single chart at half width, an orphaned-looking row. A
    // remainder of at most 2 (maxCount 3) never produces that: it's either
    // one full-width chart or an even 2-chart row.
    // horizontalBarThreshold/longLabelChars: same thresholds and same idea
    // as automated-data-analyst's autovis.py (HORIZONTAL_BAR_THRESHOLD,
    // LONG_LABEL_CHARACTERS) — a bar chart with more than a handful of
    // categories, or whose category labels are long, reads better as
    // horizontal bars (labels get room to the left instead of rotating or
    // truncating under each bar). Only the default when nothing overrides
    // it — a settings-panel chart-type pick always wins.
    chart: { height: 264, gridGap: 20, maxCount: 3, maxDimensionCardinality: 12, horizontalBarThreshold: 6, longLabelChars: 12 },
    // Shown instead of the KPI/chart block when classification produced
    // zero measure columns at all — see buildSkeleton below.
    emptyState: { height: 160 },
    table: { referenceRowHeight: 40, headerHeight: 40, footerHeight: 48, maxReferenceRows: 12, minReferenceRows: 3 },
  };

  function contentWidth(canvasWidth) {
    return canvasWidth - LAYOUT.margin.left - LAYOUT.margin.right;
  }

  // --- column selection (role-driven, no per-dataset special-casing) ---

  function selectColumns(columns) {
    const dimensions = columns.filter((c) => c.decision.role === 'dimension');
    const measures = columns.filter((c) => c.decision.role === 'measure');
    const time = columns.find((c) => c.decision.role === 'time') || null;
    const tableColumns = columns.filter((c) => c.decision.role !== 'excluded');
    return { dimensions, measures, time, tableColumns, primaryMeasure: measures[0] || null };
  }

  // Returns EVERY chart-eligible candidate, not just the first
  // LAYOUT.chart.maxCount — capping happens later, in buildSkeleton, AFTER
  // widgetConfig.enabledCharts is applied. Capping here first would mean a
  // 4th+ eligible dimension's chart could never appear in the settings
  // panel's "Charts" section at all, making it impossible to ever pick it
  // instead of one of the first maxCount — the whole point of that section.
  function planCharts(sel) {
    const plan = [];
    if (sel.time && sel.primaryMeasure) plan.push({ kind: 'line', time: sel.time });
    for (const dim of sel.dimensions) {
      // chartEligible === false is a user override (mapping screen: "filter
      // only") — the column still drives the filter bar via sel.dimensions
      // below, it just never gets a chart of its own. Absent/true (auto-
      // classified dimensions) keeps today's behavior.
      if (dim.decision.chartEligible === false) continue;
      if (dim.profile.uniqueCount <= LAYOUT.chart.maxDimensionCardinality) plan.push({ kind: 'bar', dimension: dim });
    }
    // Stable, order-based ids assigned right away — enabledCharts/
    // chartOverrides (settings-panel "Charts" section) key off these, and
    // need something to reference before enable-filtering can even run.
    return plan.map((p, i) => Object.assign({ id: `chart-${i}` }, p));
  }

  // A chart's "form of data" bounds which types it can become: a time
  // series only ever has one axis worth of buckets, so it stays a line —
  // swapping to a dimension-shaped type would need re-bucketing into
  // categories that don't exist. A dimension-by-measure chart (the 'bar'
  // plan kind) is the same {category, value} shape read three ways.
  const CHART_TYPE_OPTIONS = { line: ['line'], bar: ['bar', 'horizontalBar', 'donut'] };

  // Shared by chart overrides and KPI overrides — the settings panel's
  // Sum/Avg/Min/Max/Count pick, see engine/aggregate.js#aggregateMeasure.
  const VALID_AGGREGATIONS = ['sum', 'avg', 'min', 'max', 'count'];

  // Stands in for a real measure column when a table has zero measure-role
  // columns at all (every column is dimension/date/identifier/text) — see
  // buildSkeleton's `hasAnyMeasure` handling below. Never actually read as
  // a column: engine/aggregate.js#aggregateMeasure returns rowIndices.length
  // for a 'count' override before ever touching its `column` argument, and
  // every widget built from this always carries aggregationOverride:'count'
  // explicitly for that exact reason — there is no native aggregation to
  // fall back to for a column that was never real.
  const ROW_COUNT_MEASURE = { name: 'Row count' };

  // Applies a settings-panel chart override (type / dimensionColumn /
  // measureColumn / aggregation) on top of one auto-planned chart —
  // silently ignoring anything that doesn't resolve against the *current*
  // classification (a stale override surviving a Refresh onto data that
  // dropped the column it named), same "reconcile, don't crash" spirit as
  // engine/reconcile.js.
  function applyChartOverride(plan, sel, override) {
    if (!override) return plan;
    const next = Object.assign({}, plan);
    if (plan.kind === 'bar') {
      if (override.type && CHART_TYPE_OPTIONS.bar.includes(override.type)) next.overrideType = override.type;
      if (override.dimensionColumn) {
        const dim = sel.dimensions.find((d) => d.name === override.dimensionColumn && d.profile.uniqueCount <= LAYOUT.chart.maxDimensionCardinality);
        if (dim) next.dimension = dim;
      }
    }
    if (override.measureColumn) {
      const measure = sel.measures.find((m) => m.name === override.measureColumn);
      if (measure) next.measure = measure;
    }
    if (override.aggregation && VALID_AGGREGATIONS.includes(override.aggregation)) next.overrideAggregation = override.aggregation;
    return next;
  }

  // widgetConfig.kpiAggregations — settings panel's per-measure Sum/Avg/
  // Min/Max pick for a KPI card, keyed by measure name (a KPI has no
  // separate "chart id" the way a chart does — its own column name already
  // is a stable, unique key).
  function kpiAggregationOverride(cfg, measureName) {
    const picked = cfg.kpiAggregations && cfg.kpiAggregations[measureName];
    return picked && VALID_AGGREGATIONS.includes(picked) ? picked : null;
  }

  // --- skeleton: widget ids/types/rects and the static (data-independent)
  // parts of each widget. Positions never change with filters or theme. ---

  function buildSkeleton(columns, rowCount, widgetConfig) {
    const cfg = Object.assign(
      { enabledKpis: null, enabledCharts: null, chartOverrides: null, enabledTableColumns: null, enabledDimensions: null, kpiAggregations: null, showCharts: true, showTable: true, showFilters: true, showWatermark: false, canvasWidth: LAYOUT.width },
      widgetConfig
    );
    const sel = selectColumns(columns);
    // Not the same thing as "chartPlan/kpiMeasures ended up empty" below —
    // that can also happen when the user has simply unchecked every KPI and
    // chart in the settings panel, which is a deliberate, temporary display
    // choice, not a data problem. `hasAnyMeasure` is measured before any of
    // cfg's filtering, so it only tracks whether classification itself
    // produced zero measure-role columns at all — a table of pure text/date/
    // identifier columns (e.g. "Product Name, Category, Status, Region"),
    // which is a real, unremarkable dataset shape, not an edge case.
    // Falls back to ROW_COUNT_MEASURE + a forced 'count' aggregation rather
    // than showing nothing: a plain "how many rows" KPI (and a "count by
    // X" chart for any chart-eligible dimension) is correct and useful
    // for such a table, needs no numeric column at all, and is exactly
    // what a person would do by hand — see engine/aggregate.js#aggregateMeasure.
    const hasAnyMeasure = sel.measures.length > 0;
    const chartSel = hasAnyMeasure ? sel : Object.assign({}, sel, { primaryMeasure: ROW_COUNT_MEASURE });
    let chartPlan = cfg.showCharts ? planCharts(chartSel) : [];
    if (cfg.enabledCharts) chartPlan = chartPlan.filter((p) => cfg.enabledCharts.includes(p.id));
    // Cap AFTER enabledCharts, not inside planCharts — see planCharts' own
    // comment. Default (no enabledCharts set) keeps today's behavior: the
    // first LAYOUT.chart.maxCount eligible, in column order.
    chartPlan = chartPlan.slice(0, LAYOUT.chart.maxCount);
    if (cfg.chartOverrides) chartPlan = chartPlan.map((p) => applyChartOverride(p, sel, cfg.chartOverrides[p.id]));
    // Forced, not just defaulted: sum/avg/min/max are meaningless with no
    // real values behind them, so a stale chartOverrides.aggregation from
    // before this table lost its only measure (or never had one) must not
    // win here the way applyChartOverride would normally let it.
    if (!hasAnyMeasure) chartPlan = chartPlan.map((p) => Object.assign({}, p, { measure: ROW_COUNT_MEASURE, overrideAggregation: 'count' }));
    const kpiCandidates = cfg.enabledKpis ? sel.measures.filter((m) => cfg.enabledKpis.includes(m.name)) : sel.measures;
    const kpiMeasures = hasAnyMeasure ? kpiCandidates.slice(0, LAYOUT.kpi.maxCards) : [ROW_COUNT_MEASURE];
    // Same enabled/all pattern as kpiCandidates above — settings panel's new
    // "Filters" section (addin/dashboard-dialog.js) — filtered BEFORE the
    // maxVisible slice, so a deliberately-picked set of 4 is exactly what
    // shows, not just "the first 4 in column order" among the enabled ones.
    const dimCandidates = cfg.enabledDimensions ? sel.dimensions.filter((d) => cfg.enabledDimensions.includes(d.name)) : sel.dimensions;
    const filterDims = cfg.showFilters ? dimCandidates.slice(0, LAYOUT.filterBar.maxVisible) : [];
    const overflowDims = cfg.showFilters ? dimCandidates.slice(LAYOUT.filterBar.maxVisible) : [];

    const W = contentWidth(cfg.canvasWidth);
    const x0 = LAYOUT.margin.left;
    let y = LAYOUT.margin.top;
    const widgets = [];

    widgets.push({ id: 'header', type: 'header', rect: { x: x0, y, w: W, h: LAYOUT.header.height } });
    y += LAYOUT.header.height + LAYOUT.gap;

    if (filterDims.length || overflowDims.length) {
      widgets.push({
        id: 'filters',
        type: 'filterBar',
        rect: { x: x0, y, w: W, h: LAYOUT.filterBar.height },
        filters: filterDims.map((d) => ({ column: d.name })),
        overflow: overflowDims.map((d) => d.name),
      });
      y += LAYOUT.filterBar.height + LAYOUT.gap;
    }

    // Anchored to the filter bar's row when there is one (the rightmost
    // widget.watermark.width px of it — filter chips rarely come close to
    // using the full row width, see render/*.js's renderFilterBar), else
    // to the header row, which is always present. Either way this is
    // widgets[widgets.length - 1] at this exact point: header alone was
    // just pushed if showFilters/filterDims left the block above skipped,
    // filters last if not.
    if (cfg.showWatermark) {
      const host = widgets[widgets.length - 1];
      widgets.push({
        id: 'watermark',
        type: 'watermark',
        rect: { x: host.rect.x + host.rect.w - LAYOUT.watermark.width, y: host.rect.y, w: LAYOUT.watermark.width, h: host.rect.h },
      });
    }

    const hasKpis = kpiMeasures.length > 0;
    const hasChart1 = chartPlan.length > 0;
    if (hasKpis || hasChart1) {
      const secondary = kpiMeasures.slice(1);
      const kpiColH = LAYOUT.kpi.heroHeight + (secondary.length ? LAYOUT.kpi.innerGap + LAYOUT.kpi.secondaryHeight : 0);
      const rowH = hasChart1 ? Math.max(kpiColH, LAYOUT.chart.height) : kpiColH;

      let kpiW = W;
      let chartX = x0;
      let chartW = W;
      if (hasKpis && hasChart1) {
        kpiW = Math.round(W * LAYOUT.overview.kpiFraction - LAYOUT.overview.gap / 2);
        chartX = x0 + kpiW + LAYOUT.overview.gap;
        chartW = W - kpiW - LAYOUT.overview.gap;
      }

      if (hasKpis) {
        const hero = kpiMeasures[0];
        widgets.push({
          id: 'kpi-hero',
          type: 'kpi',
          variant: 'hero',
          column: hero.name,
          // Forced 'count' for the same reason as chartPlan above — there's
          // no real measure a cfg.kpiAggregations entry could have been
          // saved against, and sum/avg/min/max mean nothing here anyway.
          aggregationOverride: hasAnyMeasure ? kpiAggregationOverride(cfg, hero.name) : 'count',
          rect: { x: x0, y, w: kpiW, h: LAYOUT.kpi.heroHeight },
        });
        if (secondary.length) {
          const secY = y + LAYOUT.kpi.heroHeight + LAYOUT.kpi.innerGap;
          const cardW = (kpiW - (secondary.length - 1) * LAYOUT.kpi.cardGap) / secondary.length;
          secondary.forEach((m, i) => {
            widgets.push({
              id: `kpi-${m.name}`,
              type: 'kpi',
              variant: 'secondary',
              column: m.name,
              aggregationOverride: kpiAggregationOverride(cfg, m.name),
              rect: { x: x0 + i * (cardW + LAYOUT.kpi.cardGap), y: secY, w: cardW, h: LAYOUT.kpi.secondaryHeight },
            });
          });
        }
      }

      if (hasChart1) {
        widgets.push(chartWidgetSkeleton(chartPlan[0], chartSel.primaryMeasure, chartPlan[0].id, { x: chartX, y, w: chartW, h: LAYOUT.chart.height }));
      }

      y += rowH + LAYOUT.gap;
    }
    // No `else if (!hasAnyMeasure)` branch anymore — kpiMeasures always has
    // at least ROW_COUNT_MEASURE when there's no real measure, so hasKpis
    // is always true and the branch above always runs instead. The
    // 'emptyState' widget type, its renderers (render/dom.js, render/svg.js)
    // and the mapping-reopen plumbing they trigger (onOpenMapping) are left
    // in place — nothing currently constructs one, but nothing about this
    // change makes them wrong, only unreachable from here.

    const restCharts = chartPlan.slice(1);
    if (restCharts.length) {
      const cols = restCharts.length >= 2 ? 2 : 1;
      const rows = Math.ceil(restCharts.length / cols);
      const cellW = (W - (cols - 1) * LAYOUT.chart.gridGap) / cols;
      restCharts.forEach((plan, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);
        const rect = {
          x: x0 + col * (cellW + LAYOUT.chart.gridGap),
          y: y + row * (LAYOUT.chart.height + LAYOUT.chart.gridGap),
          w: cellW,
          h: LAYOUT.chart.height,
        };
        widgets.push(chartWidgetSkeleton(plan, chartSel.primaryMeasure, plan.id, rect));
      });
      y += rows * LAYOUT.chart.height + (rows - 1) * LAYOUT.chart.gridGap + LAYOUT.gap;
    }

    if (cfg.showTable) {
      // Which columns show, not whether the table shows at all — a wide
      // source can have more table-eligible columns than fit comfortably,
      // and cfg.enabledTableColumns (settings panel's "Table" section) lets
      // the user drop the ones they don't need without hiding the whole
      // table. `null` (never configured) keeps every non-excluded column,
      // same as before this existed.
      const tableColumns = cfg.enabledTableColumns
        ? sel.tableColumns.filter((c) => cfg.enabledTableColumns.includes(c.name))
        : sel.tableColumns;
      const referenceRows = Math.max(Math.min(rowCount, LAYOUT.table.minReferenceRows), Math.min(rowCount, LAYOUT.table.maxReferenceRows));
      const tableH = LAYOUT.table.headerHeight + LAYOUT.table.footerHeight + referenceRows * LAYOUT.table.referenceRowHeight;
      widgets.push({
        id: 'table',
        type: 'table',
        rect: { x: x0, y, w: W, h: tableH },
        columns: tableColumns.map((c) => ({ name: c.name, role: c.decision.role, cellFormat: c.profile.cellFormat, aggregation: c.decision.aggregation, valueScale: c.decision.valueScale || null, granularity: c.decision.granularity || null })),
        defaultSort: { key: (sel.time || tableColumns[0] || {}).name, dir: 'asc' },
      });
      y += tableH;
    }

    return { canvas: { width: cfg.canvasWidth, height: y + LAYOUT.margin.bottom }, widgets, sel, chartPlan };
  }

  // Default bar orientation when nothing overrides it — same rule and same
  // thresholds as automated-data-analyst's autovis.py: a category axis with
  // more than a handful of values, or with long labels, reads better as
  // horizontal bars (labels get room to the left instead of crowding
  // together under each bar). profile.avgTextLength is already computed by
  // engine/profile.js — no new measurement needed.
  function defaultBarType(dimension) {
    const p = dimension.profile;
    if (p.uniqueCount > LAYOUT.chart.horizontalBarThreshold) return 'horizontalBar';
    if (p.avgTextLength != null && p.avgTextLength > LAYOUT.chart.longLabelChars) return 'horizontalBar';
    return 'bar';
  }

  function chartWidgetSkeleton(plan, primaryMeasure, id, rect) {
    const measure = plan.measure || primaryMeasure;
    if (plan.kind === 'line') {
      return { id, type: 'line', rect, title: `${measure.name} over time`, timeColumn: plan.time.name, measureColumn: measure.name, aggregationOverride: plan.overrideAggregation || null };
    }
    const type = plan.overrideType || defaultBarType(plan.dimension);
    const verb = type === 'donut' ? 'share of' : 'by';
    return { id, type, rect, title: `${measure.name} ${verb} ${plan.dimension.name}`, dimensionColumn: plan.dimension.name, measureColumn: measure.name, aggregationOverride: plan.overrideAggregation || null };
  }

  // --- data fill: reads the skeleton's static fields, computes numbers
  // for the given row subset. Re-run on every filter change. ---

  function fillWidgetData(widget, columns, rowIndices, columnsByName, meta, allRowIndices) {
    if (widget.type === 'header') return Object.assign({}, widget, { title: meta.title, subtitle: meta.subtitle });

    if (widget.type === 'filterBar') {
      // Deliberately the *unfiltered* universe, not `rowIndices` (which is
      // already narrowed by whatever filters are currently active) — a
      // dimension's own chip row must keep showing every value it ever
      // had, or picking one collapses the row down to just that value
      // with no way back to the rest. `active` still reflects the current
      // selection; only which chips exist is unaffected by filtering.
      const universe = allRowIndices || rowIndices;
      return Object.assign({}, widget, {
        filters: widget.filters.map((f) => ({
          column: f.column,
          values: Aggregate.distinctValues(columnsByName.get(f.column), universe),
          active: meta.activeFilters && meta.activeFilters[f.column] ? [...meta.activeFilters[f.column]] : [],
        })),
      });
    }

    if (widget.type === 'kpi') {
      // `col` is undefined for a widget built from ROW_COUNT_MEASURE (see
      // buildSkeleton) — it was never a real column, so it's not in
      // columnsByName at all. widget.aggregationOverride is always exactly
      // 'count' in that case (never left to default), which is what keeps
      // aggregateMeasure from ever touching `col`, so every fallback below
      // just needs to not dereference it either.
      const col = columnsByName.get(widget.column);
      const agg = Aggregate.aggregateMeasure(col, rowIndices, columnsByName, widget.aggregationOverride);
      return Object.assign({}, widget, {
        label: col ? col.name : widget.column,
        value: agg.value,
        approximate: agg.approximate,
        // The EFFECTIVE aggregation actually used (override, if any) — not
        // always the classified default. `aggregationOverridden` tells the
        // renderer whether to actually SHOW that as a label — render/dom.js
        // #renderKpi does, whenever it's true, so Min/Avg/Max don't look
        // indistinguishable from an unlabeled Sum.
        aggregation: widget.aggregationOverride || (col && col.decision.aggregation),
        aggregationOverridden: !!widget.aggregationOverride,
        cellFormat: col ? col.profile.cellFormat : null,
        valueScale: (col && col.decision.valueScale) || null,
        count: rowIndices.length,
      });
    }

    if (widget.type === 'line') {
      const timeCol = columnsByName.get(widget.timeColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const points = Aggregate.bucketByTime(timeCol, measureCol, rowIndices, columnsByName, widget.aggregationOverride);
      return Object.assign({}, widget, { points, cellFormat: measureCol ? measureCol.profile.cellFormat : null, aggregation: widget.aggregationOverride || (measureCol && measureCol.decision.aggregation), aggregationOverridden: !!widget.aggregationOverride, valueScale: (measureCol && measureCol.decision.valueScale) || null });
    }

    // 'bar' and 'horizontalBar' are the same {category, value} data, just
    // drawn on different axes (render/svg.js#renderBarChart's `horizontal`
    // flag) — only 'donut' reshapes it (label/value, no per-bar rect math).
    if (widget.type === 'bar' || widget.type === 'horizontalBar') {
      const dimCol = columnsByName.get(widget.dimensionColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const bars = Aggregate.groupByDimension(dimCol, measureCol, rowIndices, columnsByName, widget.aggregationOverride);
      return Object.assign({}, widget, { bars, cellFormat: measureCol ? measureCol.profile.cellFormat : null, aggregation: widget.aggregationOverride || (measureCol && measureCol.decision.aggregation), aggregationOverridden: !!widget.aggregationOverride, valueScale: (measureCol && measureCol.decision.valueScale) || null });
    }

    if (widget.type === 'donut') {
      const dimCol = columnsByName.get(widget.dimensionColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const bars = Aggregate.groupByDimension(dimCol, measureCol, rowIndices, columnsByName, widget.aggregationOverride);
      const slices = bars.map((b) => ({ label: b.category, value: b.value }));
      return Object.assign({}, widget, { slices, cellFormat: measureCol ? measureCol.profile.cellFormat : null, aggregation: widget.aggregationOverride || (measureCol && measureCol.decision.aggregation), aggregationOverridden: !!widget.aggregationOverride, valueScale: (measureCol && measureCol.decision.valueScale) || null });
    }

    if (widget.type === 'table') {
      const sortKey = (meta.sort && meta.sort.key) || widget.defaultSort.key;
      const sortDir = (meta.sort && meta.sort.dir) || widget.defaultSort.dir;
      const sortCol = columnsByName.get(sortKey);
      const sorted = sortCol ? sortRowIndices(sortCol, rowIndices, sortDir) : rowIndices.slice();
      return Object.assign({}, widget, {
        sort: { key: sortKey, dir: sortDir },
        totalRows: rowIndices.length,
        rowIndices: sorted,
      });
    }

    return widget;
  }

  function sortRowIndices(column, rowIndices, dir) {
    const mul = dir === 'desc' ? -1 : 1;
    const values = column.values;
    return rowIndices.slice().sort((a, b) => {
      const va = values[a];
      const vb = values[b];
      if (va == null && vb == null) return 0;
      if (va == null) return 1;
      if (vb == null) return -1;
      const c = typeof va === 'string' ? va.localeCompare(vb) : va - vb;
      return c * mul;
    });
  }

  /**
   * @param {{rowCount:number, columns:Array}} analysis result of engine/index.js#analyzeTable
   * @param {{title?:string, subtitle?:string, widgetConfig?:{enabledKpis:?string[], showCharts:boolean, showTable:boolean, showFilters:boolean, showWatermark:boolean}}} [meta]
   * @returns {{canvas:object, widgets:Array}} the full layoutSpec, unfiltered
   */
  function buildLayoutSpec(analysis, meta) {
    const { columns, rowCount } = analysis;
    const fillMeta = Object.assign({ title: 'Dashboard', subtitle: null, activeFilters: {}, sort: null, widgetConfig: null }, meta);
    const skeleton = buildSkeleton(columns, rowCount, fillMeta.widgetConfig);
    const columnsByName = Aggregate.byName(columns);
    const rowIndices = Aggregate.allRowIndices(rowCount);

    return {
      canvas: skeleton.canvas,
      widgets: skeleton.widgets.map((w) => fillWidgetData(w, columns, rowIndices, columnsByName, fillMeta, rowIndices)),
    };
  }

  /**
   * Recomputes every widget's data (not its rect) for a new filter/sort
   * state, reusing the same skeleton positions — see module doc comment.
   * @param {Array} skeletonWidgets the `widgets` array from a prior buildLayoutSpec/recompute call (rects + static fields are reused as-is)
   */
  function recompute(analysis, skeletonWidgets, meta) {
    const { columns, rowCount } = analysis;
    const columnsByName = Aggregate.byName(columns);
    const allRowIndices = Aggregate.allRowIndices(rowCount);
    const rowIndices = Aggregate.filterRowIndices(columns, rowCount, meta && meta.activeFilters);
    const fillMeta = Object.assign({ title: 'Dashboard', subtitle: null, activeFilters: {}, sort: null }, meta);
    return skeletonWidgets.map((w) => fillWidgetData(w, columns, rowIndices, columnsByName, fillMeta, allRowIndices));
  }

  /**
   * Produces a fully self-contained, JSON-serializable copy of a layoutSpec
   * suitable for long-term storage (addin/dashboard-io.js writes this into
   * `workbook.settings`) or for handing to a viewer that has no access to
   * the original data — e.g. a colleague who opens the workbook, has the
   * add-in, and clicks the placed picture without the source range being
   * re-read or re-classified ("без пересчёта").
   *
   * The only widget that points at external data is `table` (`rowIndices`
   * into `columns`, resolved through the caller's live column values on
   * every render) — every other widget already carries its own numbers.
   * This bakes the table's current page of rows into literal values instead
   * of indices, exactly as many rows as the static PNG shows
   * (LAYOUT.table.maxReferenceRows), and drops `rowIndices`. A restored
   * dashboard's table therefore shows that one frozen page — sortable
   * client-side, but not re-filterable or pageable beyond it, since doing
   * either would require the original row data, which this snapshot
   * deliberately does not carry. See render/dom.js `mountFrozen`.
   *
   * @param {{canvas:object, widgets:Array}} layoutSpec a live spec from buildLayoutSpec/recompute
   * @param {Array} columns the same `columns` passed to buildLayoutSpec (for resolving table cell values)
   */
  function buildStorageSnapshot(layoutSpec, columns) {
    const columnsByName = Aggregate.byName(columns);
    const widgets = layoutSpec.widgets.map((w) => {
      if (w.type !== 'table') return w;
      const capped = w.rowIndices.slice(0, LAYOUT.table.maxReferenceRows);
      const rows = capped.map((i) => w.columns.map((c) => columnsByName.get(c.name).values[i]));
      const frozen = Object.assign({}, w, { rows, shownRows: capped.length });
      delete frozen.rowIndices;
      return frozen;
    });
    return { canvas: layoutSpec.canvas, widgets };
  }

  return { LAYOUT, selectColumns, planCharts, CHART_TYPE_OPTIONS, VALID_AGGREGATIONS, ROW_COUNT_MEASURE, applyChartOverride, defaultBarType, buildLayoutSpec, recompute, buildStorageSnapshot };
});
