'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const FormatBridge = require('../addin/format-bridge');

test('excelNumberFormatKind: percentage', () => {
  assert.equal(FormatBridge.excelNumberFormatKind('0.00%'), 'percentage');
  assert.equal(FormatBridge.excelNumberFormatKind('0%'), 'percentage');
});

test('excelNumberFormatKind: currency, bare and bracketed', () => {
  assert.equal(FormatBridge.excelNumberFormatKind('$#,##0.00'), 'currency');
  assert.equal(FormatBridge.excelNumberFormatKind('[$€-407] #,##0.00'), 'currency');
  assert.equal(FormatBridge.excelNumberFormatKind('"$"#,##0.00'), 'currency');
});

test('excelNumberFormatKind: date', () => {
  assert.equal(FormatBridge.excelNumberFormatKind('m/d/yyyy'), 'date');
  assert.equal(FormatBridge.excelNumberFormatKind('yyyy-mm-dd'), 'date');
  assert.equal(FormatBridge.excelNumberFormatKind('h:mm:ss AM/PM'), 'date');
});

test('excelNumberFormatKind: no special format', () => {
  assert.equal(FormatBridge.excelNumberFormatKind('General'), null);
  assert.equal(FormatBridge.excelNumberFormatKind('@'), null);
  assert.equal(FormatBridge.excelNumberFormatKind('#,##0.00'), null);
  assert.equal(FormatBridge.excelNumberFormatKind(''), null);
  assert.equal(FormatBridge.excelNumberFormatKind(null), null);
});

test('excelNumberFormatKind: a quoted decorative letter is not a date', () => {
  // e.g. 0.00"m" to display "5.00m" for millions — the "m" is literal text, not a month token
  assert.equal(FormatBridge.excelNumberFormatKind('0.00"m"'), null);
});

test('excelSerialToEpochMs / isoDateFromSerial: known reference point', () => {
  // 45292 is the well-known Excel serial for 2024-01-01 (1900 date system)
  assert.equal(FormatBridge.excelSerialToEpochMs(45292), Date.UTC(2024, 0, 1));
  assert.equal(FormatBridge.isoDateFromSerial(45292), '2024-01-01');
});

test('isoDateFromSerial: drops time-of-day (documented simplification)', () => {
  assert.equal(FormatBridge.isoDateFromSerial(45292.75), '2024-01-01');
});

test('cellToCanonicalText: empty and error cells become empty text', () => {
  assert.equal(FormatBridge.cellToCanonicalText(0, 'Empty', null), '');
  assert.equal(FormatBridge.cellToCanonicalText('#DIV/0!', 'Error', null), '');
  assert.equal(FormatBridge.cellToCanonicalText(null, 'Empty', null), '');
});

test('cellToCanonicalText: strings pass through untouched (preserves leading zeros)', () => {
  assert.equal(FormatBridge.cellToCanonicalText('00512', 'String', null), '00512');
  assert.equal(FormatBridge.cellToCanonicalText('North', 'String', null), 'North');
});

test('cellToCanonicalText: booleans', () => {
  assert.equal(FormatBridge.cellToCanonicalText(true, 'Boolean', null), 'TRUE');
  assert.equal(FormatBridge.cellToCanonicalText(false, 'Boolean', null), 'FALSE');
});

test('cellToCanonicalText: numbers by inferred format kind', () => {
  assert.equal(FormatBridge.cellToCanonicalText(1234.5, 'Double', null), '1234.5');
  assert.equal(FormatBridge.cellToCanonicalText(45292, 'Double', 'date'), '2024-01-01');
  assert.equal(FormatBridge.cellToCanonicalText(0.485, 'Double', 'percentage'), '48.5%');
  assert.equal(FormatBridge.cellToCanonicalText(1234.56, 'Double', 'currency'), '$1234.56');
  assert.equal(FormatBridge.cellToCanonicalText(-1234.56, 'Double', 'currency'), '$-1234.56');
});

test('cellToCanonicalText round-trips through engine/values.js#parseCell', () => {
  const Values = require('../engine/values');
  const us = Values.localeStyleForDelimiter(',');

  let text = FormatBridge.cellToCanonicalText(0.485, 'Double', 'percentage');
  let parsed = Values.parseCell(text, us);
  assert.equal(parsed.kind, 'percentage');
  assert.ok(Math.abs(parsed.numericValue - 0.485) < 1e-9);

  text = FormatBridge.cellToCanonicalText(1234.56, 'Double', 'currency');
  parsed = Values.parseCell(text, us);
  assert.equal(parsed.kind, 'currency');
  assert.ok(Math.abs(parsed.numericValue - 1234.56) < 1e-9);

  text = FormatBridge.cellToCanonicalText(45292, 'Double', 'date');
  parsed = Values.parseCell(text, us);
  assert.equal(parsed.kind, 'date');
  assert.equal(parsed.numericValue, Date.UTC(2024, 0, 1));
});
