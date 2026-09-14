/*
 * Column role classification.
 *
 * Rules run in a fixed order; the first one that matches wins. Thresholds
 * live in CONFIG below — nowhere else in this file — so tuning them never
 * means hunting through the rule bodies.
 *
 * Pure JS, no dependencies, no Office.js/Excel/DOM — see CLAUDE.md. Any
 * change to a rule below must come with updated snapshots in /tests.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashEngineRoles = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const CONFIG = {
    identifier: {
      uniqueRatioThreshold: 0.95, // integer column above this unique ratio looks like an id
      // uniqueRatio is noise below this many rows — see Rule 2's comment.
      // Doesn't gate the monotonic-sequence signal, which isn't statistical.
      minRows: 30,
      nameKeywords: ['id', 'код', 'артикул', 'номер', '№', 'sku', 'индекс'],
    },
    year: {
      min: 1900,
      max: 2100,
      maxUnique: 50,
    },
    fraction: {
      lowRange: { min: 0, max: 1 },
      highRangeWithName: { min: 0, max: 100 },
      nameKeywords: ['%', 'процент', 'доля', 'конверсия', 'маржа', 'rate', 'share', 'ratio'],
    },
    perUnit: {
      nameKeywords: ['цена', 'стоимость за', 'средн', 'на единицу', 'per'],
    },
    lowCardinality: {
      maxUnique: 25,
      maxUniqueRatio: 0.05,
    },
    weightBase: {
      // 0.5 percentage points, absolute, on the 0-1 fraction scale (i.e. a
      // ratio of 0.0550 matches a stated value of 0.0549). Absolute rather
      // than relative: a *relative* 0.5% is unforgivingly tight on small
      // percentages (a value rounded to 4 decimals, like 0.0056, can be off
      // by 0.8% relative from a single rounding step alone), which threw
      // out genuine matches during testing against fixtures/04_percent_with_base.csv.
      tolerance: 0.005,
      minComparableRows: 5,
      // Tried first, in order, before falling back to brute-force pairing
      // across every numeric column. A hint still has to pass the row-wise
      // numeric check to be accepted — matching a name alone is never enough.
      hints: [
        { subject: /conversion|конверси/i, numerator: /order|purchase|conversion|заказ|покупк/i, denominator: /visit|session|click|визит|сесси|клик/i },
        { subject: /margin|марж/i, numerator: /profit|прибыл/i, denominator: /revenue|sales|выручк|продаж/i },
        { subject: /ctr|click.?through/i, numerator: /click|клик/i, denominator: /impression|показ/i },
        { subject: /retention|удержани/i, numerator: /retain|remain|удерж|оставш/i, denominator: /total|start|начальн/i },
        { subject: /churn|отток/i, numerator: /churn|lost|ушедш|отток/i, denominator: /total|start|начальн/i },
      ],
    },
  };

  const CONFIDENCE = { FORMAT: 'high', HEURISTIC: 'medium', DEFAULT: 'low' };

  function nameLower(name) {
    return String(name).toLowerCase();
  }

  function matchesAny(name, keywords) {
    const lower = nameLower(name);
    const hit = keywords.find((k) => lower.includes(k.toLowerCase()));
    return hit || null;
  }

  function pct(ratio, digits) {
    return `${(ratio * 100).toFixed(digits == null ? 0 : digits)}%`;
  }

  function decide(role, aggregation, confidence, rule, reason, extra) {
    return Object.assign({ role, aggregation, confidence, rule, reason }, extra || {});
  }

  /**
   * @param {object} profile from engine/profile.js
   * @returns {object} {role, aggregation, confidence, rule, reason, ...}
   */
  function classifyColumn(profile) {
    const { name, valueType, cellFormat, isInteger, uniqueRatio, uniqueCount, min, max, rowCount, monotonic } = profile;

    // Rule 0 (not in the spec's numbered list): a column with no data at all
    // can't be classified by content — exclude it outright rather than
    // guessing.
    if (valueType === 'empty') {
      return decide('excluded', null, CONFIDENCE.DEFAULT, 'empty_column', 'column has no values — nothing to classify');
    }

    // --- Rule 1: cell format wins over every name-based heuristic below,
    // Rule 2 (identifier) included — each branch here returns unconditionally,
    // so a currency/percentage-formatted column can never fall through to be
    // reclassified as an identifier no matter what its name or uniqueness
    // look like. Keep every branch below an unconditional `return`; an
    // early-exit that merely fell through on some condition would silently
    // break this guarantee. ---
    if (cellFormat === 'date') {
      return decide('time', null, CONFIDENCE.FORMAT, 'format_date', 'date-formatted column — treated as a time dimension', {
        granularity: 'day',
      });
    }
    if (cellFormat === 'percentage') {
      return decide('measure', 'weighted', CONFIDENCE.FORMAT, 'format_percentage',
        'percentage-formatted column — treated as a weighted measure', { needsWeightBase: true, weightBase: null, valueScale: 'fraction' });
    }
    if (cellFormat === 'currency') {
      return decide('measure', 'sum', CONFIDENCE.FORMAT, 'format_currency', 'currency-formatted column — treated as a summable measure');
    }
    if (cellFormat === 'text-on-numeric') {
      return decide('dimension', null, CONFIDENCE.FORMAT, 'format_text_on_numeric',
        'text-formatted column with numeric-looking values (e.g. leading zeros) — treated as a dimension, not a measure');
    }

    // --- Rule 2: looks like an identifier ---
    //
    // Three independent signals, any one of which is enough:
    //
    //  a) name alone (`idKeyword`) — a column literally called "ID"/"код"/
    //     "SKU"/etc. is treated as an identifier regardless of its values or
    //     type (text SKUs included) — unchanged from before.
    //  b) a strictly monotonic integer sequence (`monotonicSequence`) — row
    //     order climbing (or falling) end to end, gaps allowed, is a
    //     structural property of sequentially-assigned IDs that a handful of
    //     rows is already enough evidence for, unlike uniqueness (see (c)).
    //     No row-count floor: this signal isn't statistical, so it isn't
    //     noisy on a small table the way a bare ratio is. Explicitly not
    //     triggered by a year-shaped range: a fiscal-year column sorted
    //     chronologically (2024, 2024, 2024, 2025, ...) is exactly as
    //     "monotonic, gaps allowed, repeats allowed" as a real ID sequence
    //     — Rule 3 below is the more specific, correct call on that shape
    //     and must not lose to this rule just because it runs second.
    //  c) high uniqueness (`idByUniquenessAndName`) — kept only as a
    //     fallback for a non-sequential integer column, and only combined
    //     with BOTH a size floor and a matching name: uniqueRatio alone is
    //     not informative below ~30 rows (five distinct amounts out of five
    //     rows is just what money looks like, not evidence of an identifier
    //     — see fixtures/11_minimal_two_columns.csv), and without a name
    //     match a purely-unique integer column (a postal code, an ad
    //     impression count) is at least as likely to be a real measure as
    //     an id. Note this branch's name requirement means it can only ever
    //     fire alongside (a) — it exists to keep the "how unique, how many
    //     rows" bar explicit and independently tunable, not because it adds
    //     a case (a) doesn't already cover.
    const idKeyword = matchesAny(name, CONFIG.identifier.nameKeywords);
    const looksLikeYearRange =
      min != null && max != null && min >= CONFIG.year.min && max <= CONFIG.year.max && uniqueCount < CONFIG.year.maxUnique;
    const monotonicSequence =
      valueType === 'number' && isInteger && !looksLikeYearRange && (monotonic === 'increasing' || monotonic === 'decreasing');
    const idByUniquenessAndName =
      valueType === 'number' && isInteger && !!idKeyword &&
      rowCount >= CONFIG.identifier.minRows &&
      uniqueRatio > CONFIG.identifier.uniqueRatioThreshold;

    if (idKeyword || monotonicSequence || idByUniquenessAndName) {
      const reason = idKeyword
        ? `column name contains "${idKeyword}" — looks like an identifier`
        : monotonicSequence
          ? `integer column, values ${monotonic} in row order — looks like a sequential identifier`
          : `integer column, ${pct(uniqueRatio)} unique across ${rowCount} rows with a matching name — looks like an identifier`;
      return decide('excluded', null, CONFIDENCE.HEURISTIC, 'identifier', reason, {
        matchedKeyword: idKeyword || null,
      });
    }

    // --- Rule 3: looks like a year ---
    if (
      valueType === 'number' &&
      isInteger &&
      min != null &&
      max != null &&
      min >= CONFIG.year.min &&
      max <= CONFIG.year.max &&
      uniqueCount < CONFIG.year.maxUnique
    ) {
      return decide('time', null, CONFIDENCE.HEURISTIC, 'year_like',
        `integer values between ${CONFIG.year.min}-${CONFIG.year.max} with ${uniqueCount} unique values — looks like a year`,
        { granularity: 'year' });
    }

    // --- Rule 4: looks like a share/percentage ---
    if (valueType === 'number' && min != null && max != null) {
      const inLowRange = !isInteger && min >= CONFIG.fraction.lowRange.min && max <= CONFIG.fraction.lowRange.max;
      const fractionKeyword = matchesAny(name, CONFIG.fraction.nameKeywords);
      const inHighRangeWithName =
        fractionKeyword && min >= CONFIG.fraction.highRangeWithName.min && max <= CONFIG.fraction.highRangeWithName.max;

      if (inLowRange || inHighRangeWithName) {
        const reason = inLowRange
          ? 'decimal values between 0 and 1 — looks like a share/rate, needs weighted aggregation'
          : `values between 0 and 100 with a rate-like name ("${fractionKeyword}") — looks like a percentage, needs weighted aggregation`;
        return decide('measure', 'weighted', CONFIDENCE.HEURISTIC, 'fraction_like', reason, {
          needsWeightBase: true,
          weightBase: null,
          valueScale: inLowRange ? 'fraction' : 'percent100',
        });
      }
    }

    // --- Rule 5: looks like a per-unit value ---
    const perUnitKeyword = valueType === 'number' ? matchesAny(name, CONFIG.perUnit.nameKeywords) : null;
    if (perUnitKeyword) {
      return decide('measure', 'avg', CONFIDENCE.HEURISTIC, 'per_unit',
        `column name contains "${perUnitKeyword}" — looks like a per-unit value, averaged rather than summed`);
    }

    // --- Rule 6: low cardinality, non-numeric -> dimension ---
    if (valueType !== 'number') {
      const cap = Math.min(CONFIG.lowCardinality.maxUnique, rowCount * CONFIG.lowCardinality.maxUniqueRatio);
      if (uniqueCount <= cap) {
        return decide('dimension', null, CONFIDENCE.DEFAULT, 'low_cardinality',
          `${uniqueCount} unique values out of ${rowCount} rows — low cardinality, looks like a dimension`);
      }
    }

    // --- Rule 7: remaining numbers -> measure/sum ---
    if (valueType === 'number') {
      return decide('measure', 'sum', CONFIDENCE.DEFAULT, 'default_measure',
        'numeric column not matched by an earlier rule — defaulting to a summable measure');
    }

    // --- Rule 8: remaining text -> text, table-only ---
    return decide('text', null, CONFIDENCE.DEFAULT, 'default_text',
      'free-form text column, kept for the table only');
  }

  // ---------------------------------------------------------------------
  // Weighted-percentage base resolution (cross-column, runs after every
  // column has an initial decision — see engine/index.js).
  // ---------------------------------------------------------------------

  function normalizedTarget(numericValues, valueScale) {
    if (valueScale === 'percent100') return numericValues.map((v) => (v == null ? null : v / 100));
    return numericValues;
  }

  function checkRatioMatch(target, numeratorValues, denominatorValues, cfg) {
    let compared = 0;
    for (let i = 0; i < target.length; i++) {
      const t = target[i];
      const num = numeratorValues[i];
      const den = denominatorValues[i];
      if (t == null || num == null || den == null || den === 0) continue;
      const ratio = num / den;
      if (Math.abs(ratio - t) > cfg.tolerance) return false;
      compared++;
    }
    return compared >= cfg.minComparableRows;
  }

  function orderedCandidatePairs(subjectName, candidates, cfg) {
    const pairs = [];
    const seen = new Set();
    function addPair(a, b) {
      const key = a.name + ' ' + b.name;
      if (seen.has(key)) return;
      seen.add(key);
      pairs.push([a, b]);
    }

    for (const hint of cfg.hints) {
      if (!hint.subject.test(subjectName)) continue;
      const numerators = candidates.filter((c) => hint.numerator.test(c.name));
      const denominators = candidates.filter((c) => hint.denominator.test(c.name));
      for (const num of numerators) for (const den of denominators) if (num !== den) addPair(num, den);
    }
    for (const num of candidates) for (const den of candidates) if (num !== den) addPair(num, den);
    return pairs;
  }

  /**
   * Mutates `decisions` in place: for every entry needing a weight base,
   * looks for a numerator/denominator pair among the other numeric columns
   * whose row-by-row ratio matches this column's values within tolerance.
   * A pair is only ever recorded once it has passed that numeric check —
   * a name-based hint alone never sets weightBase.
   *
   * @param {Array<{name:string, profile:object, decision:object, numericValues:Array<number|null>}>} entries
   */
  function resolveWeightBases(entries) {
    const numericPool = entries.filter((e) => e.profile.valueType === 'number');

    for (const entry of entries) {
      if (!entry.decision.needsWeightBase) continue;

      const target = normalizedTarget(entry.numericValues, entry.decision.valueScale);
      const candidates = numericPool.filter((c) => c !== entry);
      const pairs = orderedCandidatePairs(entry.name, candidates, CONFIG.weightBase);

      for (const [num, den] of pairs) {
        if (checkRatioMatch(target, num.numericValues, den.numericValues, CONFIG.weightBase)) {
          entry.decision.weightBase = { numerator: num.name, denominator: den.name };
          entry.decision.needsWeightBase = false;
          break;
        }
      }
    }
  }

  return { CONFIG, classifyColumn, resolveWeightBases };
});
