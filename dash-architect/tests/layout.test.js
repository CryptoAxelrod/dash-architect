'use strict';
/*
 * Tests for the layout/aggregation engine (engine/layout.js,
 * engine/aggregate.js) — the composition rules from SPEC.md §10: which
 * widgets a table produces, the caps on each, and that filtering recomputes
 * data without moving anything. Uses fixtures/01_sales_timeseries.csv,
 * whose composition was hand-verified against the rendered dashboard.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { analyzeCsv, buildLayoutSpec, recomputeLayout, Aggregate, Layout } = require('../engine');

const FIXTURES_DIR = path.join(__dirname, '..', 'fixtures');

function widgetsByType(spec) {
  const byType = {};
  for (const w of spec.widgets) (byType[w.type] ||= []).push(w);
  return byType;
}

test('sales_timeseries: composition matches the role-derived caps', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '01_sales_timeseries.csv'), 'utf8');
  const analysis = analyzeCsv(csvText);
  const spec = buildLayoutSpec(analysis, { title: 'Store sales, 2025' });
  const byType = widgetsByType(spec);

  // 4 dimensions (Region, Manager, Category, Channel), all <= 5 -> all shown as filters
  assert.equal(byType.filterBar[0].filters.length, 4);
  assert.equal(byType.filterBar[0].overflow.length, 0);

  // 5 measures (Units, Revenue, Gross Profit, Margin, Avg Unit Price) -> 1 hero + 4 secondary, all <= 6 cap
  assert.equal(byType.kpi.length, 5);
  assert.equal(byType.kpi.filter((w) => w.variant === 'hero').length, 1);
  assert.equal(byType.kpi[0].column, 'Units'); // first measure column, in column order

  // 1 time column + 4 qualifying dimensions -> capped at 4 charts total (1 line + 3 bar)
  assert.equal(byType.line.length, 1);
  assert.equal(byType.bar.length, 3);
  assert.equal(byType.line.length + byType.bar.length, 4);

  // every widget rect stays within the canvas
  for (const w of spec.widgets) {
    assert.ok(w.rect.x >= 0 && w.rect.y >= 0);
    assert.ok(w.rect.x + w.rect.w <= spec.canvas.width + 0.5);
  }
});

test('Margin KPI uses the verified weight base, not a naive row-average', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '01_sales_timeseries.csv'), 'utf8');
  const analysis = analyzeCsv(csvText);
  const spec = buildLayoutSpec(analysis);
  const margin = spec.widgets.find((w) => w.id === 'kpi-Margin');

  assert.equal(margin.approximate, false);
  // sum(Gross Profit) / sum(Revenue) over all 900 rows, not mean(Margin per row)
  const cols = Aggregate.byName(analysis.columns);
  const rows = Aggregate.allRowIndices(analysis.rowCount);
  const expected =
    cols.get('Gross Profit').values.reduce((a, v, i) => a + v, 0) / cols.get('Revenue').values.reduce((a, v) => a + v, 0);
  assert.ok(Math.abs(margin.value - expected) < 1e-9);
});

test('filtering recomputes widget data without moving any widget', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '01_sales_timeseries.csv'), 'utf8');
  const analysis = analyzeCsv(csvText);
  const spec = buildLayoutSpec(analysis);
  const before = spec.widgets.map((w) => ({ id: w.id, rect: w.rect }));

  const after = recomputeLayout(analysis, spec.widgets, { activeFilters: { Region: new Set(['Central']) } });
  const afterRects = after.map((w) => ({ id: w.id, rect: w.rect }));
  assert.deepStrictEqual(afterRects, before);

  const heroBefore = spec.widgets.find((w) => w.id === 'kpi-hero');
  const heroAfter = after.find((w) => w.id === 'kpi-hero');
  assert.equal(heroAfter.count, 189); // rows where Region === 'Central'
  assert.notEqual(heroAfter.value, heroBefore.value);
});

test('a weighted measure with no verified base falls back to a flagged mean (05_percent_without_base.csv)', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '05_percent_without_base.csv'), 'utf8');
  const analysis = analyzeCsv(csvText);
  const spec = buildLayoutSpec(analysis);
  const conversionRate = spec.widgets.find((w) => w.column === 'Conversion Rate');

  assert.equal(conversionRate.approximate, true);
  const col = Aggregate.byName(analysis.columns).get('Conversion Rate');
  const rows = Aggregate.allRowIndices(analysis.rowCount);
  const expectedMean = col.values.reduce((a, v) => a + v, 0) / col.values.length;
  assert.ok(Math.abs(conversionRate.value - expectedMean) < 1e-9);
});

test('buildStorageSnapshot: table becomes self-contained (no rowIndices), everything else unchanged', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '01_sales_timeseries.csv'), 'utf8');
  const analysis = analyzeCsv(csvText);
  const spec = buildLayoutSpec(analysis, { title: 'Store sales, 2025' });
  const snapshot = Layout.buildStorageSnapshot(spec, analysis.columns);

  // round-trips through JSON exactly like addin/dashboard-io.js's workbook.settings write would
  const restored = JSON.parse(JSON.stringify(snapshot));

  const liveTable = spec.widgets.find((w) => w.type === 'table');
  const frozenTable = restored.widgets.find((w) => w.type === 'table');
  assert.equal(frozenTable.rowIndices, undefined);
  assert.equal(frozenTable.shownRows, Layout.LAYOUT.table.maxReferenceRows);
  assert.equal(frozenTable.rows.length, Layout.LAYOUT.table.maxReferenceRows);

  // the embedded rows are the actual cell values for the first N sorted row indices, in column order
  const cols = Aggregate.byName(analysis.columns);
  const expectedFirstRow = liveTable.columns.map((c) => cols.get(c.name).values[liveTable.rowIndices[0]]);
  assert.deepStrictEqual(frozenTable.rows[0], expectedFirstRow);

  // every non-table widget is untouched (still carries its own numbers, no external lookup)
  for (const w of restored.widgets) {
    if (w.type === 'table') continue;
    const live = spec.widgets.find((x) => x.id === w.id);
    assert.deepStrictEqual(w, JSON.parse(JSON.stringify(live)));
  }

  // small enough to comfortably fit in a workbook.settings entry
  assert.ok(JSON.stringify(snapshot).length < 50000);
});
