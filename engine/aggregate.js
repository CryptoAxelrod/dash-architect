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

  // Sentinel category a blank/missing dimension value is bucketed under —
  // in filter chips, chart grouping, and filter matching alike — instead of
  // silently dropping those rows from every breakdown. Never appears in
  // `column.values` itself (that stays `null`); only introduced at the
  // point a raw value is turned into a bucket/filter key. Table cells still
  // render a plain "—" for a blank (render/dom.js#formatCell) — that's a
  // per-row display choice, unrelated to this aggregation-level bucketing.
  const BLANK = '(blank)';

  function bucketKey(rawValue) {
    return rawValue == null ? BLANK : rawValue;
  }

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
        if (!check.set.has(bucketKey(check.values[i]))) {
          ok = false;
          break;
        }
      }
      if (ok) out.push(i);
    }
    return out;
  }

  // A blank measure cell counts as 0 (not "skip this row") — it still
  // contributes to the row count an average divides by, so a mix of real
  // values and blanks pulls the average down rather than quietly averaging
  // over only the populated rows.
  function sum(values, rowIndices) {
    if (!rowIndices.length) return null;
    let total = 0;
    for (const i of rowIndices) {
      const v = values[i];
      total += v == null ? 0 : v;
    }
    return total;
  }

  function mean(values, rowIndices) {
    if (!rowIndices.length) return null;
    let total = 0;
    for (const i of rowIndices) {
      const v = values[i];
      total += v == null ? 0 : v;
    }
    return total / rowIndices.length;
  }

  // Unlike sum/mean, a blank cell is simply excluded here rather than
  // counted as 0 — a missing value isn't a real "low" (for min) or "high"
  // (for max) reading, and treating it as one would silently invent a
  // minimum/maximum that never actually occurred in the data.
  function min(values, rowIndices) {
    let m = null;
    for (const i of rowIndices) {
      const v = values[i];
      if (v == null) continue;
      if (m == null || v < m) m = v;
    }
    return m;
  }

  function max(values, rowIndices) {
    let m = null;
    for (const i of rowIndices) {
      const v = values[i];
      if (v == null) continue;
      if (m == null || v > m) m = v;
    }
    return m;
  }

  /**
   * Aggregates one measure column over a row subset. Normally honors the
   * column's classified aggregation strategy (see engine/roles.js), but
   * `overrideAggregation` — a user's explicit Sum/Avg/Min/Max pick from the
   * settings panel (widgetConfig.kpiAggregations / a chart's
   * chartOverrides[id].aggregation) — takes full priority over that when
   * given, bypassing the weighted numerator/denominator logic entirely.
   * Returns a value already scaled to match how the column's own row values
   * are displayed (see SPEC.md §8): a weighted measure with a resolved base
   * is computed from the raw numerator/denominator sums, never from the
   * percent values themselves; one without a base falls back to a flagged,
   * unweighted mean.
   *
   * @param {string} [overrideAggregation] 'sum' | 'avg' | 'min' | 'max'
   * @returns {{value: number|null, approximate: boolean}}
   */
  function aggregateMeasure(column, rowIndices, columnsByName, overrideAggregation) {
    const { decision, values } = column;
    const agg = overrideAggregation || decision.aggregation;
    // Only a *deviation* from a verified weighted ratio counts as an
    // approximation — overriding an already-plain sum measure to show its
    // average instead is just a different, equally exact view of the same
    // values, not a guess.
    const approximateOverride = !!overrideAggregation && decision.aggregation === 'weighted' && overrideAggregation !== 'weighted';

    if (agg === 'sum') return { value: sum(values, rowIndices), approximate: approximateOverride };
    if (agg === 'avg') return { value: mean(values, rowIndices), approximate: approximateOverride };
    if (agg === 'min') return { value: min(values, rowIndices), approximate: approximateOverride };
    if (agg === 'max') return { value: max(values, rowIndices), approximate: approximateOverride };

    if (agg === 'weighted') {
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
   * @returns {Array<{value:string, count:number}>} sorted by count, descending
   *   (ties broken alphabetically for determinism) — the filter bar's only
   *   consumer caps this at a fixed number of chips before "+N", so the
   *   ones kept visible need to be the most-represented values, not
   *   whichever happen to sort first alphabetically.
   */
  function distinctValues(column, rowIndices) {
    const counts = new Map();
    for (const i of rowIndices) {
      const v = bucketKey(column.values[i]);
      counts.set(v, (counts.get(v) || 0) + 1);
    }
    return [...counts.entries()]
      .map(([value, count]) => ({ value, count }))
      .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
  }

  /**
   * Buckets a row subset by a dimension column's distinct values, each
   * bucket carrying the primary measure aggregated over just its rows.
   * Sorted by aggregated value, descending (ties broken alphabetically) —
   * the common "what stands out" ranked-bar convention; there is no
   * reference precedent for a categorical bar chart to follow instead.
   */
  function groupByDimension(dimensionColumn, measureColumn, rowIndices, columnsByName, overrideAggregation) {
    const groups = new Map();
    for (const i of rowIndices) {
      const v = bucketKey(dimensionColumn.values[i]);
      if (!groups.has(v)) groups.set(v, []);
      groups.get(v).push(i);
    }
    const out = [...groups.entries()].map(([category, idx]) => ({
      category,
      ...aggregateMeasure(measureColumn, idx, columnsByName, overrideAggregation),
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
  function bucketByTime(timeColumn, measureColumn, rowIndices, columnsByName, overrideAggregation) {
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
      .map((b) => ({ label: b.label, ...aggregateMeasure(measureColumn, b.idx, columnsByName, overrideAggregation), count: b.idx.length }));
  }

  return {
    BLANK,
    bucketKey,
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
