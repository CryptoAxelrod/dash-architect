/*
 * Engine entry point: wires csv.js -> values.js -> profile.js -> roles.js
 * into one deterministic `analyzeTable` call.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(
      require('./csv'),
      require('./values'),
      require('./profile'),
      require('./roles'),
      require('./aggregate'),
      require('./layout')
    );
  } else {
    root.DashEngine = factory(
      root.DashEngineCsv,
      root.DashEngineValues,
      root.DashEngineProfile,
      root.DashEngineRoles,
      root.DashEngineAggregate,
      root.DashEngineLayout
    );
  }
})(typeof self !== 'undefined' ? self : this, function (Csv, Values, Profile, Roles, Aggregate, Layout) {
  'use strict';

  /**
   * @param {string[]} headers column names, in column order
   * @param {string[][]} dataRows rows of raw string cells, aligned to `headers` (no header row)
   * @param {string} [delimiter=','] the delimiter the source CSV used — controls number-locale parsing (see engine/values.js)
   * @returns {{rowCount:number, columns: Array<{name:string, profile:object, decision:object, values:Array<number|string|null>}>}}
   */
  function analyzeTable(headers, dataRows, delimiter) {
    const localeStyle = Values.localeStyleForDelimiter(delimiter || ',');
    const rowCount = dataRows.length;

    const entries = headers.map((name, colIndex) => {
      const cells = dataRows.map((row) => Values.parseCell(row[colIndex], localeStyle));
      const { profile, numericValues } = Profile.buildColumnProfile(name, cells, rowCount);
      const decision = Roles.classifyColumn(profile);
      // resolveWeightBases (engine/roles.js) reads `numericValues` by that exact
      // name on every entry — keep it intact rather than renaming in place.
      return { name, profile, decision, numericValues, cells };
    });

    Roles.resolveWeightBases(entries);

    return {
      rowCount,
      columns: entries.map((e) => ({
        name: e.name,
        profile: e.profile,
        decision: e.decision,
        // Typed, per-row values for downstream aggregation/layout (engine/aggregate.js,
        // engine/layout.js): numbers/dates use the already-parsed numeric value, text
        // columns keep the trimmed raw string, empty cells are null either way.
        values: e.profile.valueType === 'text' ? e.cells.map((c) => (c.empty ? null : c.raw.trim())) : e.numericValues,
      })),
    };
  }

  /**
   * Convenience wrapper: parses raw CSV text and analyzes it in one call.
   * Assumes the first row is the header row.
   * @param {string} csvText
   */
  function analyzeCsv(csvText) {
    const { delimiter, rows } = Csv.parseCsv(csvText);
    const [headers, ...dataRows] = rows;
    return analyzeTable(headers || [], dataRows, delimiter);
  }

  return {
    analyzeTable,
    analyzeCsv,
    CONFIG: Roles.CONFIG,
    Aggregate,
    Layout,
    buildLayoutSpec: Layout.buildLayoutSpec,
    recomputeLayout: Layout.recompute,
  };
});
