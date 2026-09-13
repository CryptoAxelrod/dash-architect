'use strict';
/*
 * Regression test for a real crash found while verifying the 200,000-row
 * target from the addin brief: engine/profile.js used to compute a
 * column's min/max via Math.min(...values)/Math.max(...values), which
 * throws "Maximum call stack size exceeded" once the array is large enough
 * to spread as call arguments (observed well under 200k rows in this
 * engine). Fixed with a plain loop (minOf/maxOf) — this test pins that a
 * large numeric column no longer crashes and computes the right min/max.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const { buildColumnProfile } = require('../engine/profile');
const { parseCell } = require('../engine/values');

test('buildColumnProfile does not crash on a very large numeric column and computes correct min/max', () => {
  const n = 200000;
  const cells = new Array(n);
  for (let i = 0; i < n; i++) cells[i] = parseCell(String(i), 'us');
  // plant a known min and max away from the ends, where a subtly-wrong
  // partial implementation would still happen to get the extremes right
  cells[12345] = parseCell('-999999', 'us');
  cells[54321] = parseCell('999999', 'us');

  const { profile } = buildColumnProfile('N', cells, n);
  assert.equal(profile.min, -999999);
  assert.equal(profile.max, 999999);
});
