/*
 * Pure helpers for carrying dashboard UI state (filters/sort/KPI picks)
 * across a change of underlying data — a Refresh (same source, newer rows)
 * or a Change data range onto a different source. Deliberately dumb about
 * *why* the data changed; the caller decides whether to use these at all.
 *
 * Structure comparison is by column NAME only (see SPEC.md) — `role` is
 * excluded on purpose: it can legitimately differ between two analyses of
 * the same shape (a user override) without the underlying report having
 * changed, and `profile.valueType` is exactly as volatile (one stray text
 * cell in an otherwise-numeric column of new rows flips it) — neither is a
 * safe signal for "should I throw away the user's filters." A per-column
 * valueType note for the UI (informational only) does NOT go through here.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineReconcile = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** @returns {boolean} true iff both column-name sets are identical (order-independent) */
  function columnsStructureMatches(oldNames, newNames) {
    if (oldNames.length !== newNames.length) return false;
    const b = new Set(newNames);
    if (b.size !== oldNames.length) return false;
    return oldNames.every((n) => b.has(n));
  }

  /**
   * Adapts a previously-serialized dialog state ({activeFilters, sort,
   * theme, widgetConfig} — see addin/dialog-messaging.js) to a (possibly
   * different) analysis: drops filter values and sort/KPI references to
   * columns that no longer exist or no longer play the required role.
   * Never throws, never invents new state — a name simply absent from
   * `rawState` stays absent.
   *
   * @param {{columns:Array}} analysis the analysis the caller is about to mount/remount with
   * @param {{activeFilters?, sort?, theme?, palette?, widgetConfig?}} rawState
   * @returns {{activeFilters, sort, theme, palette, widgetConfig}|null} null if rawState itself was null/undefined
   */
  // Mirrors engine/aggregate.js's BLANK sentinel — a blank/missing cell is
  // filterable as '(blank)' there, so a saved '(blank)' filter selection
  // must match against still-blank rows here too, not just literal values
  // (this module has no dependency on Aggregate — see its doc comment — so
  // the sentinel is duplicated rather than imported).
  const BLANK = '(blank)';

  function reconcileDashboardState(analysis, rawState) {
    if (!rawState) return null;
    const byName = new Map(analysis.columns.map((c) => [c.name, c]));

    const activeFilters = {};
    for (const [col, rawValues] of Object.entries(rawState.activeFilters || {})) {
      const column = byName.get(col);
      if (!column || column.decision.role !== 'dimension' || !Array.isArray(rawValues) || !rawValues.length) continue;
      const stillPresent = new Set(column.values.map((v) => (v == null ? BLANK : v)));
      const kept = rawValues.filter((v) => stillPresent.has(v));
      if (kept.length) activeFilters[col] = kept;
    }

    let sort = rawState.sort || null;
    if (sort && !byName.has(sort.key)) sort = null;

    let widgetConfig = rawState.widgetConfig || null;
    if (widgetConfig && Array.isArray(widgetConfig.enabledKpis)) {
      const kept = widgetConfig.enabledKpis.filter((name) => {
        const c = byName.get(name);
        return c && c.decision.role === 'measure';
      });
      widgetConfig = Object.assign({}, widgetConfig, { enabledKpis: kept.length ? kept : null });
    }
    if (widgetConfig && Array.isArray(widgetConfig.enabledTableColumns)) {
      const kept = widgetConfig.enabledTableColumns.filter((name) => {
        const c = byName.get(name);
        return c && c.decision.role !== 'excluded';
      });
      widgetConfig = Object.assign({}, widgetConfig, { enabledTableColumns: kept.length ? kept : null });
    }
    // chartOverrides' chart ids are positional (engine/layout.js#planCharts
    // assigns them fresh every build) so a stale id just silently stops
    // matching anything on the next build — nothing to reconcile there.
    // The dimension/measure *names* inside a surviving override are a real
    // dangling reference risk after a Refresh/Change-range drops a column,
    // so those get the same drop-if-gone treatment as enabledKpis.
    if (widgetConfig && widgetConfig.chartOverrides) {
      const cleaned = {};
      for (const [chartId, override] of Object.entries(widgetConfig.chartOverrides)) {
        const next = {};
        if (override.type) next.type = override.type;
        if (override.dimensionColumn) {
          const c = byName.get(override.dimensionColumn);
          if (c && c.decision.role === 'dimension') next.dimensionColumn = override.dimensionColumn;
        }
        if (override.measureColumn) {
          const c = byName.get(override.measureColumn);
          if (c && c.decision.role === 'measure') next.measureColumn = override.measureColumn;
        }
        if (Object.keys(next).length) cleaned[chartId] = next;
      }
      widgetConfig = Object.assign({}, widgetConfig, { chartOverrides: Object.keys(cleaned).length ? cleaned : null });
    }

    // 'ocean' mirrors render/palettes.js#DEFAULT_PALETTE as a literal —
    // engine stays free of any dependency on /render (CLAUDE.md §2's spirit
    // extended to the render layer too), same reason 'light' above isn't
    // imported from render/themes.js either.
    return { activeFilters, sort, theme: rawState.theme || 'light', palette: rawState.palette || 'ocean', widgetConfig };
  }

  /**
   * Per-column value-type drift between two analyses, for shared column
   * names only — informational (a UI banner), never a basis for resetting
   * anything. See the module doc comment for why.
   * @returns {string[]} names of columns whose profile.valueType differs
   */
  function diffValueTypes(oldAnalysis, newAnalysis) {
    const oldByName = new Map(oldAnalysis.columns.map((c) => [c.name, c.profile.valueType]));
    const changed = [];
    for (const col of newAnalysis.columns) {
      const oldType = oldByName.get(col.name);
      if (oldType != null && oldType !== col.profile.valueType) changed.push(col.name);
    }
    return changed;
  }

  return { columnsStructureMatches, reconcileDashboardState, diffValueTypes };
});
