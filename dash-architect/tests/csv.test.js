'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { parseCsv } = require('../engine/csv');

test('detects comma delimiter and splits fields', () => {
  const { delimiter, rows } = parseCsv('a,b,c\n1,2,3\n');
  assert.equal(delimiter, ',');
  assert.deepEqual(rows, [
    ['a', 'b', 'c'],
    ['1', '2', '3'],
  ]);
});

test('detects semicolon delimiter (ru-locale files)', () => {
  const { delimiter, rows } = parseCsv('Дата;Выручка\n04.07.2025;925 408,28\n');
  assert.equal(delimiter, ';');
  assert.deepEqual(rows, [
    ['Дата', 'Выручка'],
    ['04.07.2025', '925 408,28'],
  ]);
});

test('strips a leading UTF-8 BOM', () => {
  const { rows } = parseCsv('﻿a,b\n1,2\n');
  assert.equal(rows[0][0], 'a');
});

test('handles quoted fields with an embedded delimiter and escaped quotes', () => {
  const { rows } = parseCsv('name,note\n"Smith, John","said ""hi"""\n');
  assert.deepEqual(rows, [
    ['name', 'note'],
    ['Smith, John', 'said "hi"'],
  ]);
});

test('handles quoted fields spanning multiple lines', () => {
  const { rows } = parseCsv('a,b\n"line1\nline2",2\n');
  assert.deepEqual(rows, [
    ['a', 'b'],
    ['line1\nline2', '2'],
  ]);
});

test('drops trailing blank lines', () => {
  const { rows } = parseCsv('a,b\n1,2\n\n\n');
  assert.equal(rows.length, 2);
});

test('parses a file with no trailing newline', () => {
  const { rows } = parseCsv('a,b\n1,2');
  assert.deepEqual(rows, [
    ['a', 'b'],
    ['1', '2'],
  ]);
});
