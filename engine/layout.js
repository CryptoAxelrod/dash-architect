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
    width: 1180,
    margin: { top: 32, right: 28, bottom: 40, left: 28 },
    gap: 28, // vertical gap between major sections
    header: { height: 60 },
    // One chip-row per filterable dimension (render/dom.js#renderFilterBar) —
    // sized for the worst case (maxVisible rows) so it never overlaps the
    // widget below it; most datasets show fewer than maxVisible and just
    // leave the rest of the box empty, same tradeoff every other fixed
    // heuristic constant in this file makes.
    filterBar: { height: 150, maxVisible: 5 },
    kpi: {
      maxCards: 6,
      heroHeight: 150,
      secondaryHeight: 108,
      innerGap: 16, // between hero and the secondary row
      cardGap: 14, // between secondary cards
    },
    overview: { gap: 32, kpiFraction: 5 / 12 },
    chart: { height: 264, gridGap: 20, maxCount: 4, maxDimensionCardinality: 12 },
    table: { referenceRowHeight: 40, headerHeight: 40, footerHeight: 48, maxReferenceRows: 12, minReferenceRows: 3 },
  };

  function contentWidth() {
    return LAYOUT.width - LAYOUT.margin.left - LAYOUT.margin.right;
  }

  // --- column selection (role-driven, no per-dataset special-casing) ---

  function selectColumns(columns) {
    const dimensions = columns.filter((c) => c.decision.role === 'dimension');
    const measures = columns.filter((c) => c.decision.role === 'measure');
    const time = columns.find((c) => c.decision.role === 'time') || null;
    const tableColumns = columns.filter((c) => c.decision.role !== 'excluded');
    return { dimensions, measures, time, tableColumns, primaryMeasure: measures[0] || null };
  }

  function planCharts(sel) {
    const plan = [];
    if (sel.time && sel.primaryMeasure) plan.push({ kind: 'line', time: sel.time });
    for (const dim of sel.dimensions) {
      if (plan.length >= LAYOUT.chart.maxCount) break;
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
    return plan.slice(0, LAYOUT.chart.maxCount).map((p, i) => Object.assign({ id: `chart-${i}` }, p));
  }

  // A chart's "form of data" bounds which types it can become: a time
  // series only ever has one axis worth of buckets, so it stays a line —
  // swapping to a dimension-shaped type would need re-bucketing into
  // categories that don't exist. A dimension-by-measure chart (the 'bar'
  // plan kind) is the same {category, value} shape read three ways.
  const CHART_TYPE_OPTIONS = { line: ['line'], bar: ['bar', 'horizontalBar', 'donut'] };

  // Applies a settings-panel chart override (type / dimensionColumn /
  // measureColumn) on top of one auto-planned chart — silently ignoring
  // anything that doesn't resolve against the *current* classification
  // (a stale override surviving a Refresh onto data that dropped the
  // column it named), same "reconcile, don't crash" spirit as
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
    return next;
  }

  // --- skeleton: widget ids/types/rects and the static (data-independent)
  // parts of each widget. Positions never change with filters or theme. ---

  function buildSkeleton(columns, rowCount, widgetConfig) {
    const cfg = Object.assign({ enabledKpis: null, enabledCharts: null, chartOverrides: null, showCharts: true, showTable: true, showFilters: true }, widgetConfig);
    const sel = selectColumns(columns);
    let chartPlan = cfg.showCharts ? planCharts(sel) : [];
    if (cfg.enabledCharts) chartPlan = chartPlan.filter((p) => cfg.enabledCharts.includes(p.id));
    if (cfg.chartOverrides) chartPlan = chartPlan.map((p) => applyChartOverride(p, sel, cfg.chartOverrides[p.id]));
    const kpiCandidates = cfg.enabledKpis ? sel.measures.filter((m) => cfg.enabledKpis.includes(m.name)) : sel.measures;
    const kpiMeasures = kpiCandidates.slice(0, LAYOUT.kpi.maxCards);
    const filterDims = cfg.showFilters ? sel.dimensions.slice(0, LAYOUT.filterBar.maxVisible) : [];
    const overflowDims = cfg.showFilters ? sel.dimensions.slice(LAYOUT.filterBar.maxVisible) : [];

    const W = contentWidth();
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
              rect: { x: x0 + i * (cardW + LAYOUT.kpi.cardGap), y: secY, w: cardW, h: LAYOUT.kpi.secondaryHeight },
            });
          });
        }
      }

      if (hasChart1) {
        widgets.push(chartWidgetSkeleton(chartPlan[0], sel.primaryMeasure, chartPlan[0].id, { x: chartX, y, w: chartW, h: LAYOUT.chart.height }));
      }

      y += rowH + LAYOUT.gap;
    }

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
        widgets.push(chartWidgetSkeleton(plan, sel.primaryMeasure, plan.id, rect));
      });
      y += rows * LAYOUT.chart.height + (rows - 1) * LAYOUT.chart.gridGap + LAYOUT.gap;
    }

    if (cfg.showTable) {
      const referenceRows = Math.max(Math.min(rowCount, LAYOUT.table.minReferenceRows), Math.min(rowCount, LAYOUT.table.maxReferenceRows));
      const tableH = LAYOUT.table.headerHeight + LAYOUT.table.footerHeight + referenceRows * LAYOUT.table.referenceRowHeight;
      widgets.push({
        id: 'table',
        type: 'table',
        rect: { x: x0, y, w: W, h: tableH },
        columns: sel.tableColumns.map((c) => ({ name: c.name, role: c.decision.role, cellFormat: c.profile.cellFormat, aggregation: c.decision.aggregation, valueScale: c.decision.valueScale || null })),
        defaultSort: { key: (sel.time || sel.tableColumns[0] || {}).name, dir: 'asc' },
      });
      y += tableH;
    }

    return { canvas: { width: LAYOUT.width, height: y + LAYOUT.margin.bottom }, widgets, sel, chartPlan };
  }

  function chartWidgetSkeleton(plan, primaryMeasure, id, rect) {
    const measure = plan.measure || primaryMeasure;
    if (plan.kind === 'line') {
      return { id, type: 'line', rect, title: `${measure.name} over time`, timeColumn: plan.time.name, measureColumn: measure.name };
    }
    const type = plan.overrideType || 'bar';
    const verb = type === 'donut' ? 'share of' : 'by';
    return { id, type, rect, title: `${measure.name} ${verb} ${plan.dimension.name}`, dimensionColumn: plan.dimension.name, measureColumn: measure.name };
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
      const col = columnsByName.get(widget.column);
      const agg = Aggregate.aggregateMeasure(col, rowIndices, columnsByName);
      return Object.assign({}, widget, {
        label: col.name,
        value: agg.value,
        approximate: agg.approximate,
        aggregation: col.decision.aggregation,
        cellFormat: col.profile.cellFormat,
        valueScale: col.decision.valueScale || null,
        count: rowIndices.length,
      });
    }

    if (widget.type === 'line') {
      const timeCol = columnsByName.get(widget.timeColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const points = Aggregate.bucketByTime(timeCol, measureCol, rowIndices, columnsByName);
      return Object.assign({}, widget, { points, cellFormat: measureCol.profile.cellFormat, aggregation: measureCol.decision.aggregation, valueScale: measureCol.decision.valueScale || null });
    }

    // 'bar' and 'horizontalBar' are the same {category, value} data, just
    // drawn on different axes (render/svg.js#renderBarChart's `horizontal`
    // flag) — only 'donut' reshapes it (label/value, no per-bar rect math).
    if (widget.type === 'bar' || widget.type === 'horizontalBar') {
      const dimCol = columnsByName.get(widget.dimensionColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const bars = Aggregate.groupByDimension(dimCol, measureCol, rowIndices, columnsByName);
      return Object.assign({}, widget, { bars, cellFormat: measureCol.profile.cellFormat, aggregation: measureCol.decision.aggregation, valueScale: measureCol.decision.valueScale || null });
    }

    if (widget.type === 'donut') {
      const dimCol = columnsByName.get(widget.dimensionColumn);
      const measureCol = columnsByName.get(widget.measureColumn);
      const bars = Aggregate.groupByDimension(dimCol, measureCol, rowIndices, columnsByName);
      const slices = bars.map((b) => ({ label: b.category, value: b.value }));
      return Object.assign({}, widget, { slices, cellFormat: measureCol.profile.cellFormat, aggregation: measureCol.decision.aggregation, valueScale: measureCol.decision.valueScale || null });
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
   * @param {{title?:string, subtitle?:string, widgetConfig?:{enabledKpis:?string[], showCharts:boolean, showTable:boolean, showFilters:boolean}}} [meta]
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

  return { LAYOUT, selectColumns, planCharts, CHART_TYPE_OPTIONS, applyChartOverride, buildLayoutSpec, recompute, buildStorageSnapshot };
});
