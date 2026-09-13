'use strict';
/*
 * Tests for the pure (naming/key/payload) half of addin/dashboard-io.js.
 * Everything else in that file calls context.workbook (shapes, settings)
 * and can only be exercised inside a real Excel add-in — untestable here.
 * See SPEC.md's addin section.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const DashboardIo = require('../addin/dashboard-io');

test('makeShapeName / isDashboardShapeName round-trip and reject unrelated names', () => {
  const name = DashboardIo.makeShapeName(123);
  assert.equal(name, 'DashArchitectDashboard_123');
  assert.ok(DashboardIo.isDashboardShapeName(name));
  assert.ok(!DashboardIo.isDashboardShapeName('Picture 1'));
  assert.ok(!DashboardIo.isDashboardShapeName(undefined));
});

test('settingsKeyForShape / shapeNameFromSettingsKey round-trip', () => {
  const name = DashboardIo.makeShapeName('abc');
  const key = DashboardIo.settingsKeyForShape(name);
  assert.equal(key, 'dashArchitect:DashArchitectDashboard_abc');
  assert.ok(DashboardIo.isDashboardSettingsKey(key));
  assert.ok(!DashboardIo.isDashboardSettingsKey('someUnrelatedSetting'));
  assert.equal(DashboardIo.shapeNameFromSettingsKey(key), name);
});

test('buildStoragePayload carries the snapshot, mapping and source address as given', () => {
  const payload = DashboardIo.buildStoragePayload({
    snapshot: { canvas: { width: 1180, height: 1600 }, widgets: [] },
    mapping: [{ name: 'Region', role: 'dimension' }],
    sourceAddress: "'Sheet1'!A1:E301",
    title: 'Store sales, 2025',
  });
  assert.equal(payload.version, 1);
  assert.equal(payload.sourceAddress, "'Sheet1'!A1:E301");
  assert.equal(payload.title, 'Store sales, 2025');
  assert.deepStrictEqual(payload.mapping, [{ name: 'Region', role: 'dimension' }]);
  assert.deepStrictEqual(payload.layoutSpec, { canvas: { width: 1180, height: 1600 }, widgets: [] });
  assert.ok(!Number.isNaN(Date.parse(payload.generatedAt)));
});
