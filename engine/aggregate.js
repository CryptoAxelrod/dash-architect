/*
 * Pure aggregation/filtering helpers shared by engine/layout.js (initial
 * build) and the DOM renderer (recompute on a filter change). Everything
 * here operates on the `{name, profile, decision, values}` column shape
 * produced by engine/index.js#analyzeTable — no DOM, no I/O.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineAggregate = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function byName(columns) {
    const map = new Map();
    for (const c of columns) map.set(c.name, c);
    return map;
  }

  function allRowIndices(rowCount) {
    const idx = new Array(rowCount);
    for (let i = 0; i < rowCount; i++) idx[i] = i;
    return idx;
  }

  /**
   * @param {Array} columns
   * @param {number} rowCount
   * @param {Object<string,Set<string>>} activeFilters column name -> set of allowed values (dimension columns only). A column absent, or with an empty set, is not filtered.
   * @returns {number[]} row indices that pass every active filter
   */
  function filterRowIndices(columns, rowCount, activeFilters) {
    const entries = Object.entries(activeFilters || {}).filter(([, set]) => set && set.size > 0);
    if (!entries.length) return allRowIndices(rowCount);

    const cols = byName(columns);
    const checks = entries.map(([name, set]) => ({ values: cols.get(name).values, set }));

    const out = [];
    for (let i = 0; i < rowCount; i++) {
      let ok = true;
      for (const check of checks) {
        if (!check.set.has(check.values[i])) {
          ok = false;
          break;
        }
      }
      if (ok) out.push(i);
    }
    return out;
  }

  function sum(values, rowIndices) {
    let total = 0;
    let any = false;
    for (const i of rowIndices) {
      const v = values[i];
      if (v == null) continue;
      total += v;
      any = true;
    }
    return any ? total : null;
  }

  function mean(values, rowIndices) {
    let total = 0;
    let n = 0;
    for (const i of rowIndices) {
      const v = values[i];
      if (v == null) continue;
      total += v;
      n++;
    }
    return n ? total / n : null;
  }

  /**
   * Aggregates one measure column over a row subset, honoring its
   * classified aggregation strategy (see engine/roles.js). Returns a value
   * already scaled to match how the column's own row values are displayed
   * (see SPEC.md §8): a weighted measure with a resolved base is computed
   * from the raw numerator/denominator sums, never from the percent values
   * themselves; one without a base falls back to a flagged, unweighted mean.
   *
   * @returns {{value: number|null, approximate: boolean}}
   */
  function aggregateMeasure(column, rowIndices, columnsByName) {
    const { decision, values } = column;

    if (decision.aggregation === 'sum') return { value: sum(values, rowIndices), approximate: false };
    if (decision.aggregation === 'avg') return { value: mean(values, rowIndices), approximate: false };

    if (decision.aggregation === 'weighted') {
      if (decision.weightBase) {
        const num = columnsByName.get(decision.weightBase.numerator).values;
        const den = columnsByName.get(decision.weightBase.denominator).values;
        const sumNum = sum(num, rowIndices);
        const sumDen = sum(den, rowIndices);
        if (sumDen == null || sumDen === 0) return { value: null, approximate: false };
        const ratio = sumNum / sumDen;
        return { value: decision.valueScale === 'percent100' ? ratio * 100 : ratio, approximate: false };
      }
      // No verified numerator/denominator pair: an unweighted mean of the
      // row values is a labeled approximation, not a silent guess — see
      // SPEC.md §8 and CLAUDE.md's "don't guess" principle.
      return { value: mean(values, rowIndices), approximate: true };
    }

    return { value: null, approximate: false };
  }

  /**
   * Distinct dimension values present in a row subset, with counts.
   * @returns {Array<{value:string, count:number}>} sorted alphabetically (matches reference/dashboard.html's category chip order)
   */
  function distinctValues(column, rowIndices) {
    const counts = new Map();
    for (const i of rowIndices) {
      const v = column.values[i];
      if (v == null) continue;
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    return [...counts.entries()]
      .map(([value, count]) => ({ value, count }))
      .sort((a, b) => a.value.localeCompare(b.value));
  }

  /**
   * Buckets a row subset by a dimension column's distinct values, each
   * bucket carrying the primary measure aggregated over just its rows.
   * Sorted by aggregated value, descending (ties broken alphabetically) —
   * the common "what stands out" ranked-bar convention; there is no
   * reference precedent for a categorical bar chart to follow instead.
   */
  function groupByDimension(dimensionColumn, measureColumn, rowIndices, columnsByName) {
    const groups = new Map();
    for (const i of rowIndices) {
      const v = dimensionColumn.values[i];
      if (v == null) continue;
      if (!groups.has(v)) groups.set(v, []);
      groups.get(v).push(i);
    }
    const out = [...groups.entries()].map(([category, idx]) => ({
      category,
      ...aggregateMeasure(measureColumn, idx, columnsByName),
      count: idx.length,
    }));
    out.sort((a, b) => (b.value || 0) - (a.value || 0) || a.category.localeCompare(b.category));
    return out;
  }

  // --- time bucketing --------------------------------------------------

  const DAY_MS = 86400000;

  /** Deterministic bucket-unit choice from the date span — see SPEC.md §10. */
  function chooseTimeUnit(minEpoch, maxEpoch) {
    const spanDays = (maxEpoch - minEpoch) / DAY_MS;
    if (spanDays <= 14) return 'day';
    if (spanDays <= 120) return 'week';
    if (spanDays <= 900) return 'month';
    if (spanDays <= 3650) return 'quarter';
    return 'year';
  }

  function bucketKeyAndLabel(epoch, unit) {
    const d = new Date(epoch);
    const y = d.getUTCFullYear();
    const m = d.getUTCMonth(); // 0-based
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (unit === 'day') {
      const key = Math.floor(epoch / DAY_MS);
      return { key, sortKey: key, label: `${MONTHS[m]} ${d.getUTCDate()}` };
    }
    if (unit === 'week') {
      const dow = (d.getUTCDay() + 6) % 7; // Monday = 0
      const weekStart = epoch - dow * DAY_MS;
      const wd = new Date(weekStart);
      const key = Math.floor(weekStart / DAY_MS);
      return { key, sortKey: key, label: `${MONTHS[wd.getUTCMonth()]} ${wd.getUTCDate()}` };
    }
    if (unit === 'month') {
      return { key: `${y}-${m}`, sortKey: y * 12 + m, label: `${MONTHS[m]} ${y}` };
    }
    if (unit === 'quarter') {
      const q = Math.floor(m / 3) + 1;
      return { key: `${y}-Q${q}`, sortKey: y * 4 + q, label: `Q${q} ${y}` };
    }
    return { key: y, sortKey: y, label: String(y) };
  }

  /**
   * Buckets a row subset along a `time` column, aggregating the primary
   * measure per bucket. A `year_like` time column (see engine/roles.js
   * rule 3) has no real calendar to bucket — its own integer values are
   * the buckets, one per distinct year.
   */
  function bucketByTime(timeColumn, measureColumn, rowIndices, columnsByName) {
    const isYearLike = timeColumn.decision.granularity === 'year';
    const unit = isYearLike ? null : chooseTimeUnit(timeColumn.profile.min, timeColumn.profile.max);
    const buckets = new Map();

    for (const i of rowIndices) {
      const v = timeColumn.values[i];
      if (v == null) continue;
      const { key, sortKey, label } = isYearLike ? { key: v, sortKey: v, label: String(v) } : bucketKeyAndLabel(v, unit);
      if (!buckets.has(key)) buckets.set(key, { sortKey, label, idx: [] });
      buckets.get(key).idx.push(i);
    }

    return [...buckets.values()]
      .sort((a, b) => (a.sortKey > b.sortKey ? 1 : a.sortKey < b.sortKey ? -1 : 0))
      .map((b) => ({ label: b.label, ...aggregateMeasure(measureColumn, b.idx, columnsByName), count: b.idx.length }));
  }

  return {
    byName,
    allRowIndices,
    filterRowIndices,
    aggregateMeasure,
    distinctValues,
    groupByDimension,
    chooseTimeUnit,
    bucketByTime,
  };
});
