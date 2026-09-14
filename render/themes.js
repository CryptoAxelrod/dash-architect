/*
 * Theme tokens. A theme is *only* a set of values — palette, font stack,
 * spacing scale, radii, chart series palette, line weight, a textures
 * flag — read by both renderers. It never changes layout or which widgets
 * exist (engine/layout.js owns that, identically for every theme).
 *
 * `light` is extracted from reference/dashboard.html, values unchanged
 * (see SPEC.md §3.1) — it's the default. `dark`, `report` and `print` each
 * recompute contrast and the chart palette for their own background rather
 * than inverting/reusing `light`'s numbers — see the comment on each.
 *
 * No network calls: Archivo (the reference's Google Font) needs a network
 * fetch, which fights the "pixel predictability" this renderer exists for
 * (a slow/blocked font load would silently change the static SVG export).
 * All four themes use the same system-font fallback stack instead.
 *
 * Pure JS, no dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./palettes'));
  } else {
    root.DashRenderThemes = factory(root.DashRenderPalettes);
  }
})(typeof self !== 'undefined' ? self : this, function (Palettes) {
  'use strict';

  const FONT_STACK = '"Segoe UI", "Helvetica Neue", Arial, system-ui, sans-serif';

  // `chart.seriesColors` below is each theme's *default* series palette —
  // render/palettes.js's "Ocean", the same values, single source of
  // truth. A caller that wants a different palette overrides this via
  // Palettes.withPalette(theme, paletteId) rather than this file growing a
  // second, redundant color list per theme — see palettes.js's doc comment
  // for why palette and theme are separate tokens in the first place.
  const LIGHT_SERIES = Palettes.PALETTES.ocean.light;

  const light = {
    id: 'light',
    name: 'Light',
    color: {
      paper: '#EEF1EC', panel: '#FAFBF9', panelBorder: '#D2D9D3', ink: '#17231F',
      muted: '#5B6964', faint: '#8A968F', rule: '#D2D9D3', ruleSoft: '#E4E9E5',
      accent: '#1D6F4C', accentDark: '#14523A', accentMid: '#6FA388', accentSoft: '#C6DDD0',
      negative: '#B23A24', negativeSoft: '#E7B9AE', focus: '#1D6F4C',
    },
    negativeStyle: 'color',
    font: { family: FONT_STACK },
    radius: { sm: 6, md: 10, lg: 12, pill: 999 },
    spacing: { xs: 4, sm: 8, md: 14, lg: 20 },
    card: { style: 'panel', shadow: false, borderWidth: 1 },
    table: { rowHeight: 40, headerSize: 13, cellSize: 14 },
    type: { h1: 31, subtitle: 14.5, panelTitle: 17, kpiHero: 56, kpiValue: 26, kpiLabel: 14, axis: 11.5, tableCell: 14, tableHeader: 13 },
    chart: {
      seriesColors: LIGHT_SERIES, lineWidth: 2, barRadius: 3,
      gridColor: '#D2D9D3', axisColor: '#5B6964', textures: false, dashPatterns: null, markers: null,
    },
  };

  // Not an inversion of `light`: a dark, green-tinted charcoal (continuing
  // the "paper" mood rather than jumping to neutral gray or near-black),
  // with the accent brightened for legibility and every series color pulled
  // up in lightness and pulled down in saturation from `light`'s — a flat
  // hue-inversion would either wash out against the dark ground or, at full
  // saturation, read as neon, which is the exact effect being avoided.
  const DARK_SERIES = Palettes.PALETTES.ocean.dark;

  const dark = {
    id: 'dark',
    name: 'Dark',
    color: {
      paper: '#101A16', panel: '#16211C', panelBorder: '#2A3A33', ink: '#E7EFE9',
      muted: '#8FA69A', faint: '#62766C', rule: '#2A3A33', ruleSoft: '#212E28',
      accent: '#4FAE7C', accentDark: '#79C79A', accentMid: '#3D8563', accentSoft: '#1F3B2E',
      negative: '#E2836D', negativeSoft: '#4A2620', focus: '#4FAE7C',
    },
    negativeStyle: 'color',
    font: { family: FONT_STACK },
    radius: { sm: 6, md: 10, lg: 12, pill: 999 },
    spacing: { xs: 4, sm: 8, md: 14, lg: 20 },
    card: { style: 'panel', shadow: false, borderWidth: 1 },
    table: { rowHeight: 40, headerSize: 13, cellSize: 14 },
    type: { h1: 31, subtitle: 14.5, panelTitle: 17, kpiHero: 56, kpiValue: 26, kpiLabel: 14, axis: 11.5, tableCell: 14, tableHeader: 13 },
    chart: {
      seriesColors: DARK_SERIES, lineWidth: 2, barRadius: 3,
      gridColor: '#26332D', axisColor: '#8FA69A', textures: false, dashPatterns: null, markers: null,
    },
  };

  // Dense corporate mode: white page, flat sections divided by a thin rule
  // instead of bordered/shadowed cards, tighter type and spacing so more
  // rows/cards fit before scrolling — "for pasting into a report to
  // management." Same accent hue as `light`, just slightly less saturated
  // to read as restrained rather than a website's bright chart green.
  const REPORT_SERIES = Palettes.PALETTES.ocean.report;

  const report = {
    id: 'report',
    name: 'Report',
    color: {
      paper: '#FFFFFF', panel: '#FFFFFF', panelBorder: '#DADEDC', ink: '#16201B',
      muted: '#5C6864', faint: '#8B948F', rule: '#DADEDC', ruleSoft: '#EDEFEE',
      accent: '#1B6146', accentDark: '#123D2C', accentMid: '#5E8C76', accentSoft: '#D8E4DD',
      negative: '#A23522', negativeSoft: '#F0DAD3', focus: '#1B6146',
    },
    negativeStyle: 'color',
    font: { family: FONT_STACK },
    radius: { sm: 2, md: 4, lg: 6, pill: 999 },
    spacing: { xs: 2, sm: 6, md: 10, lg: 14 },
    card: { style: 'flat', shadow: false, borderWidth: 1 },
    table: { rowHeight: 28, headerSize: 11, cellSize: 12 },
    type: { h1: 22, subtitle: 12, panelTitle: 13, kpiHero: 34, kpiValue: 20, kpiLabel: 11.5, axis: 10, tableCell: 12, tableHeader: 11 },
    chart: {
      seriesColors: REPORT_SERIES, lineWidth: 1.5, barRadius: 1,
      gridColor: '#EDEFEE', axisColor: '#5C6864', textures: false, dashPatterns: null, markers: null,
    },
  };

  // Print/PDF: high contrast, near-black on white, minimal fills. Series
  // are told apart by hatch fill + line dash + point-marker shape as the
  // primary signal, color as a secondary one — so the dashboard still
  // reads correctly in black-and-white print and for colorblind viewers.
  // Negative numbers use accounting-style parentheses rather than color.
  const PRINT_SERIES = Palettes.PALETTES.ocean.print;

  const print = {
    id: 'print',
    name: 'Print',
    color: {
      paper: '#FFFFFF', panel: '#FFFFFF', panelBorder: '#111111', ink: '#000000',
      muted: '#333333', faint: '#555555', rule: '#111111', ruleSoft: '#CCCCCC',
      accent: '#155A3E', accentDark: '#0E3D2A', accentMid: '#3F7A5D', accentSoft: '#DCEAE2',
      negative: '#000000', negativeSoft: '#DDDDDD', focus: '#000000',
    },
    negativeStyle: 'parens',
    font: { family: FONT_STACK },
    radius: { sm: 0, md: 2, lg: 2, pill: 999 },
    spacing: { xs: 2, sm: 6, md: 10, lg: 14 },
    card: { style: 'flat', shadow: false, borderWidth: 1.5 },
    table: { rowHeight: 26, headerSize: 11, cellSize: 11 },
    type: { h1: 20, subtitle: 11, panelTitle: 12, kpiHero: 30, kpiValue: 18, kpiLabel: 10.5, axis: 9.5, tableCell: 11, tableHeader: 10.5 },
    chart: {
      seriesColors: PRINT_SERIES, lineWidth: 2, barRadius: 0,
      gridColor: '#CCCCCC', axisColor: '#000000', textures: true,
      dashPatterns: ['none', '6 3', '2 2', '8 2 2 2', '6 2 2 2 2 2', '10 3 2 3'],
      markers: ['circle', 'square', 'triangle', 'diamond', 'circle', 'square'],
      hatches: ['diagonal', 'cross', 'dots', 'horizontal', 'diagonal2', 'vertical'],
    },
  };

  const THEMES = { light, dark, report, print };

  return { THEMES, DEFAULT_THEME: 'light' };
});
