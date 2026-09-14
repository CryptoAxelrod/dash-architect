'use strict';
/*
 * Snapshot tests for column classification (engine/roles.js), one per file
 * in /fixtures — the reference JSON pins role/aggregation/confidence/rule
 * for every column. Per CLAUDE.md: any change to a classification rule
 * must come with updated snapshots here.
 *
 * Run: npm test (or `node --test tests`) — one command, and a failing
 * fixture prints only the columns that diverged (name + the specific
 * fields that changed), not a full object dump — see tests/lib/classify-snapshot.js.
 *
 * After an intentional rule/threshold change, regenerate and inspect the
 * diff yourself before trusting it:
 *   UPDATE_SNAPSHOTS=1 node --test tests/classify.test.js
 *   git diff tests/snapshots/
 *
 * fixtures/06_messy_headers.csv (no header-row detection yet — see
 * SPEC.md §6), fixtures/10_large_volume.csv (60k rows), and
 * fixtures/11_minimal_two_columns.csv (5 rows, Date + a numeric column
 * whose 5 values happen to all be distinct — Rule 2's uniqueRatio > 0.95
 * identifier check has no row-count floor, so it currently misclassifies
 * that column as an excluded identifier rather than a measure; see the
 * diagnosis in the conversation this fixture came from) are snapshotted
 * like every other fixture: whatever the classifier currently does on them
 * is what's pinned, "correct" or not — that's the point of a snapshot. A
 * future fix to the identifier rule's row-count handling will need this
 * snapshot regenerated, which is exactly the point: CLAUDE.md §6 requires
 * that update to happen visibly, in the same change, not silently.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { analyzeCsv } = require('../engine');
const { FIXTURES_DIR, listFixtureCsvFiles, snapshotPathFor, summarizeColumns, diffColumns, formatDiff } = require('./lib/classify-snapshot');

for (const fixture of listFixtureCsvFiles()) {
  test(`column classification snapshot: ${fixture}`, () => {
    const csvText = fs.readFileSync(path.join(FIXTURES_DIR, fixture), 'utf8');
    const actual = summarizeColumns(analyzeCsv(csvText));
    const snapshotPath = snapshotPathFor(fixture);

    if (process.env.UPDATE_SNAPSHOTS) {
      fs.writeFileSync(snapshotPath, JSON.stringify(actual, null, 2) + '\n');
      return;
    }

    if (!fs.existsSync(snapshotPath)) {
      assert.fail(`${fixture} — no snapshot yet at ${path.relative(process.cwd(), snapshotPath)}. Generate one: UPDATE_SNAPSHOTS=1 node --test tests/classify.test.js`);
    }

    const expected = JSON.parse(fs.readFileSync(snapshotPath, 'utf8'));
    const diffs = diffColumns(expected, actual);
    assert.ok(diffs.length === 0, diffs.length ? formatDiff(fixture, diffs) : undefined);
  });
}

test('large volume fixture: analyzes 60k rows without error and stays fast', () => {
  const csvText = fs.readFileSync(path.join(FIXTURES_DIR, '10_large_volume.csv'), 'utf8');
  const start = Date.now();
  const result = analyzeCsv(csvText);
  const elapsedMs = Date.now() - start;

  assert.equal(result.rowCount, 60000);
  assert.equal(result.columns.length, 6);
  assert.ok(elapsedMs < 5000, `expected under 5s, took ${elapsedMs}ms`);
});
