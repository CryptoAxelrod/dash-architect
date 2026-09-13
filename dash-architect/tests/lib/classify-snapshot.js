'use strict';
/*
 * Shared plumbing for the classifier snapshot tests (tests/classify.test.js):
 * discovering fixtures, summarizing a column's classification, and diffing
 * two summaries down to just the fields that actually changed. Kept
 * separate from the test file so `node --test` and a future standalone
 * "what changed" script can't drift apart on what counts as a divergence.
 */
const fs = require('node:fs');
const path = require('node:path');

const FIXTURES_DIR = path.join(__dirname, '..', '..', 'fixtures');
const SNAPSHOTS_DIR = path.join(__dirname, '..', 'snapshots');

const COMPARED_FIELDS = ['valueType', 'cellFormat', 'role', 'aggregation', 'confidence', 'rule', 'weightBase', 'needsWeightBase'];

/** Every `.csv` fixture directly under /fixtures (not fixtures/xlsx — the engine has no xlsx reader), sorted. */
function listFixtureCsvFiles() {
  return fs
    .readdirSync(FIXTURES_DIR)
    .filter((f) => f.toLowerCase().endsWith('.csv'))
    .sort();
}

function snapshotPathFor(fixtureFile) {
  return path.join(SNAPSHOTS_DIR, fixtureFile.replace(/\.csv$/i, '.json'));
}

/** @param {{columns: Array}} analysis engine/index.js#analyzeTable (or analyzeCsv) result */
function summarizeColumns(analysis) {
  return analysis.columns.map(({ name, profile, decision }) => ({
    name,
    valueType: profile.valueType,
    cellFormat: profile.cellFormat,
    role: decision.role,
    aggregation: decision.aggregation,
    confidence: decision.confidence,
    rule: decision.rule,
    weightBase: decision.weightBase || null,
    needsWeightBase: !!decision.needsWeightBase,
  }));
}

/**
 * Compares two column-summary arrays by column name (not position — robust
 * to reordering) and returns only what differs: per changed column, only
 * the fields that actually changed; columns added/removed get one entry
 * each rather than a field-by-field dump.
 * @returns {Array<{column:string, field:string, before:*, after:*}>}
 */
function diffColumns(expected, actual) {
  const diffs = [];
  const expectedByName = new Map(expected.map((c) => [c.name, c]));
  const actualByName = new Map(actual.map((c) => [c.name, c]));

  for (const [name, exp] of expectedByName) {
    const act = actualByName.get(name);
    if (!act) {
      diffs.push({ column: name, field: '(column)', before: 'present', after: 'missing' });
      continue;
    }
    for (const field of COMPARED_FIELDS) {
      const b = exp[field];
      const a = act[field];
      if (JSON.stringify(b) !== JSON.stringify(a)) diffs.push({ column: name, field, before: b, after: a });
    }
  }
  for (const name of actualByName.keys()) {
    if (!expectedByName.has(name)) diffs.push({ column: name, field: '(column)', before: 'missing', after: 'present' });
  }
  return diffs;
}

/** Groups diffColumns' flat list back into one line per column for compact reading. */
function formatDiff(fixtureFile, diffs) {
  const byColumn = new Map();
  for (const d of diffs) {
    if (!byColumn.has(d.column)) byColumn.set(d.column, []);
    byColumn.get(d.column).push(d);
  }
  const lines = [`${fixtureFile} — ${byColumn.size} column(s) diverged from the snapshot:`];
  for (const [column, fieldDiffs] of byColumn) {
    const parts = fieldDiffs.map((d) => `${d.field} ${JSON.stringify(d.before)} → ${JSON.stringify(d.after)}`);
    lines.push(`  ${column}: ${parts.join(', ')}`);
  }
  return lines.join('\n');
}

module.exports = { FIXTURES_DIR, SNAPSHOTS_DIR, listFixtureCsvFiles, snapshotPathFor, summarizeColumns, diffColumns, formatDiff };
