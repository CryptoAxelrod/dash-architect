'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { localeStyleForDelimiter, parseCell } = require('../engine/values');

test('locale style follows the CSV delimiter', () => {
  assert.equal(localeStyleForDelimiter(','), 'us');
  assert.equal(localeStyleForDelimiter(';'), 'euro');
});

test('empty cell', () => {
  assert.deepEqual(parseCell('', 'us'), { raw: '', empty: true, kind: 'empty', numericValue: null, isInteger: false });
  assert.equal(parseCell('   ', 'us').kind, 'empty');
});

test('ISO date', () => {
  const c = parseCell('2025-01-01', 'us');
  assert.equal(c.kind, 'date');
  assert.equal(c.numericValue, Date.UTC(2025, 0, 1));
});

test('European d.m.yyyy date', () => {
  const c = parseCell('07.04.2024', 'us');
  assert.equal(c.kind, 'date');
  assert.equal(c.numericValue, Date.UTC(2024, 3, 7));
});

test('an invalid calendar date falls back to text, not a broken date', () => {
  const c = parseCell('32.13.2024', 'us');
  assert.equal(c.kind, 'text');
});

test('percentage sign normalizes to a 0-1 fraction', () => {
  const c = parseCell('50%', 'us');
  assert.equal(c.kind, 'percentage');
  assert.equal(c.numericValue, 0.5);
  assert.equal(c.isInteger, true);
});

test('ru-locale percentage (comma decimal)', () => {
  const c = parseCell('48,5%', 'euro');
  assert.equal(c.kind, 'percentage');
  assert.equal(c.numericValue, 0.485);
});

test('currency symbol is stripped', () => {
  const c = parseCell('$1,234.56', 'us');
  assert.equal(c.kind, 'currency');
  assert.equal(c.numericValue, 1234.56);
});

test('plain integer', () => {
  const c = parseCell('512', 'us');
  assert.equal(c.kind, 'number');
  assert.equal(c.numericValue, 512);
  assert.equal(c.isInteger, true);
});

test('negative decimal', () => {
  const c = parseCell('-6.63', 'us');
  assert.equal(c.kind, 'number');
  assert.equal(c.numericValue, -6.63);
  assert.equal(c.isInteger, false);
});

test('ru-locale number (space thousands, comma decimal)', () => {
  const c = parseCell('925 408,28', 'euro');
  assert.equal(c.kind, 'number');
  assert.equal(c.numericValue, 925408.28);
});

test('a leading zero on a digit string is treated as text, not a number', () => {
  const c = parseCell('00512', 'us');
  assert.equal(c.kind, 'text');
  assert.equal(c.numericValue, null);
});

test('plain digits with no leading zero stay numeric', () => {
  assert.equal(parseCell('512', 'us').kind, 'number');
  assert.equal(parseCell('0', 'us').kind, 'number'); // a bare "0" isn't a leading-zero pattern
});

test('free text stays text', () => {
  const c = parseCell('Feature request: bulk edit', 'us');
  assert.equal(c.kind, 'text');
});
