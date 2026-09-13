'use strict';
/*
 * Tests for the pure half of addin/excel-io.js — toEngineInput and
 * inferColumnFormats operate on plain arrays shaped like Office.js's
 * range.values/valueTypes/numberFormat, no Excel host required. The other
 * half of that file (readRangeRaw) calls Office.js directly and can only
 * be exercised inside a real Excel add-in — see SPEC.md's addin section
 * for what could and couldn't be verified in this environment.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const { analyzeTable } = require('../engine');
const ExcelIo = require('../addin/excel-io');

function fakeRaw({ headers, rows, formats }) {
  const cols = headers.length;
  const values = [headers, ...rows];
  const valueTypes = values.map((row) => row.map((v) => (v === '' || v == null ? 'Empty' : typeof v === 'string' ? 'String' : typeof v === 'boolean' ? 'Boolean' : 'Double')));
  const numberFormatSample = [headers.map(() => 'General'), ...rows.map(() => formats)];
  return { cols, totalRows: values.length, values, valueTypes, numberFormatSample };
}

test('toEngineInput: infers a per-column format from the sample and converts every row', () => {
  const raw = fakeRaw({
    headers: ['Date', 'Category', 'Rate'],
    rows: [
      [45292, 'North', 0.15],
      [45293, 'South', 0.22],
      [45294, 'North', 0.18],
    ],
    formats: ['m/d/yyyy', 'General', '0.00%'],
  });

  const { headers, dataRows, columnFormats } = ExcelIo.toEngineInput(raw);
  assert.deepStrictEqual(headers, ['Date', 'Category', 'Rate']);
  assert.deepStrictEqual(columnFormats, ['date', null, 'percentage']);
  assert.deepStrictEqual(dataRows, [
    ['2024-01-01', 'North', '15%'],
    ['2024-01-02', 'South', '22%'],
    ['2024-01-03', 'North', '18%'],
  ]);
});

test('toEngineInput output feeds engine.analyzeTable and classifies as expected', () => {
  const raw = fakeRaw({
    headers: ['Date', 'Category', 'Rate'],
    rows: [
      [45292, 'North', 0.15],
      [45293, 'South', 0.22],
      [45294, 'North', 0.18],
      [45295, 'East', 0.09],
      [45296, 'South', 0.31],
    ],
    formats: ['m/d/yyyy', 'General', '0.00%'],
  });

  const { headers, dataRows } = ExcelIo.toEngineInput(raw);
  const analysis = analyzeTable(headers, dataRows, ',');
  const byName = new Map(analysis.columns.map((c) => [c.name, c]));

  assert.equal(byName.get('Date').decision.role, 'time');
  assert.equal(byName.get('Date').decision.rule, 'format_date');
  assert.equal(byName.get('Rate').decision.role, 'measure');
  assert.equal(byName.get('Rate').decision.aggregation, 'weighted');
  assert.equal(byName.get('Rate').profile.cellFormat, 'percentage');
});

test('inferColumnFormats: an empty/blank sample leaves a column unclassified (falls through to a plain number)', () => {
  const raw = fakeRaw({
    headers: ['Plain'],
    rows: [[42], [43]],
    formats: ['General'],
  });
  assert.deepStrictEqual(ExcelIo.inferColumnFormats(raw.numberFormatSample, 1), [null]);
});

test('toEngineInput: Empty/Error cells become blank text, not literal garbage', () => {
  const raw = {
    cols: 2,
    totalRows: 3,
    values: [
      ['Name', 'Score'],
      ['Alice', 10],
      ['', '#DIV/0!'],
    ],
    valueTypes: [
      ['String', 'String'],
      ['String', 'Double'],
      ['Empty', 'Error'],
    ],
    numberFormatSample: [
      ['General', 'General'],
      ['General', 'General'],
      ['General', 'General'],
    ],
  };
  const { dataRows } = ExcelIo.toEngineInput(raw);
  assert.deepStrictEqual(dataRows, [
    ['Alice', '10'],
    ['', ''],
  ]);
});
