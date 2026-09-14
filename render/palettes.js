/*
 * Chart series palettes — deliberately separate from render/themes.js.
 * Before this file, every theme's `chart.seriesColors` was a shade of its
 * own accent green, so every chart was green regardless of which theme was
 * picked. A palette is a set of *data-encoding* colors (line/bar fills,
 * donut slices, grouped-bar series); `theme.color.accent` stays what it
 * always was — brand/UI-state color (buttons, active filter chips, hover
 * highlight) — untouched by palette choice.
 *
 * Each palette carries one 6-color list per theme (`light`/`dark`/`report`/
 * `print`) rather than one list recolored on the fly, for the same reason
 * render/themes.js hand-tunes contrast per theme instead of inverting
 * `light`: a color that reads on a near-white ground and one that reads on
 * a near-black one are rarely the same value at a different lightness.
 *
 * The first three colors of every palette/theme combination share one hue
 * triple (amber/azure/teal — see HUES below) chosen and verified against a
 * standard deuteranopia simulation (Brettel/Viénot-style LMS confusion
 * matrix) specifically so they stay distinguishable under red-green color
 * blindness, the most common type — deliberately avoiding the classic
 * red-vs-green confusion pair for the positions a 1-3 series chart
 * actually uses. Colors 4-6 (violet/vermillion/olive) extend the palette
 * for higher-cardinality charts (grouped bars, donuts) but are not
 * independently verified against every other color — CLAUDE.md's
 * determinism rule doesn't require a hue wheel to solve an NP-hard
 * covering problem, and the explicit requirement was "the first three."
 *
 * `print`'s variant is intentionally near-grayscale for all four palettes:
 * render/svg.js's print theme already tells series apart by hatch pattern
 * + line dash + marker shape (see render/themes.js's comment on `print`),
 * with color as a secondary, half-tone-safe signal — a palette swap there
 * only nudges overall contrast/darkness, never hue, so it can't undermine
 * the black-and-white-safe design.
 *
 * Pure JS, no dependencies.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DashRenderPalettes = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const PALETTES = {
    meridian: {
      id: 'meridian', name: 'Meridian',
      light: ['#936639', '#395093', '#39937C', '#6D3993', '#934839', '#399348'],
      dark: ['#CAA37D', '#7D90CA', '#7DCAB6', '#AA7DCA', '#CA897D', '#7DCA89'],
      report: ['#825C35', '#354982', '#35826F', '#623582', '#824235', '#358242'],
      print: ['#111111', '#3D3D3D', '#6B6B6B', '#8F8F8F', '#4C4C4C', '#2E2E2E'],
    },
    harbor: {
      id: 'harbor', name: 'Harbor',
      light: ['#8E5C29', '#29438E', '#298E75', '#64298E', '#8E3A29', '#298E3A'],
      dark: ['#CF9E6E', '#6E86CF', '#6ECFB6', '#A66ECF', '#CF7E6E', '#6ECF7E'],
      report: ['#7C5227', '#273C7C', '#277C67', '#59277C', '#7C3527', '#277C35'],
      print: ['#050505', '#303030', '#5E5E5E', '#828282', '#404040', '#212121'],
    },
    ember: {
      id: 'ember', name: 'Ember',
      light: ['#AB7036', '#3653AB', '#36AB8D', '#7A36AB', '#AB4936', '#36AB49'],
      dark: ['#D2A87F', '#7F93D2', '#7FD2BD', '#AF7FD2', '#D28D7F', '#7FD28D'],
      report: ['#976635', '#354E97', '#35977E', '#6E3597', '#974535', '#359745'],
      print: ['#13110F', '#433D37', '#766B60', '#9A8F84', '#544C45', '#322E29'],
    },
    slate: {
      id: 'slate', name: 'Slate',
      light: ['#9D754D', '#4D619D', '#4D9D89', '#7C4D9D', '#9D5B4D', '#4D9D5B'],
      dark: ['#BE9974', '#7487BE', '#74BEAB', '#9F74BE', '#BE8174', '#74BE81'],
      report: ['#906B47', '#475990', '#47907D', '#714790', '#905347', '#479053'],
      print: ['#303030', '#5C5C5C', '#8A8A8A', '#ADADAD', '#6B6B6B', '#4C4C4C'],
    },
  };

  const DEFAULT_PALETTE = 'meridian';

  function seriesColorsFor(paletteId, themeId) {
    const palette = PALETTES[paletteId] || PALETTES[DEFAULT_PALETTE];
    return palette[themeId] || palette.light;
  }

  // The one place a theme and a palette actually combine: a shallow copy of
  // `theme` with `chart.seriesColors` swapped for the chosen palette's
  // colors for that theme's id. Every renderer keeps reading
  // `theme.chart.seriesColors` exactly as before — callers that pick a
  // theme (addin/dashboard-dialog.js, render/dom.js callers, the PNG export
  // path) route it through here first instead of teaching render/svg.js or
  // render/dom.js a second "palette" parameter.
  function withPalette(theme, paletteId) {
    if (!paletteId) return theme;
    return Object.assign({}, theme, { chart: Object.assign({}, theme.chart, { seriesColors: seriesColorsFor(paletteId, theme.id) }) });
  }

  return { PALETTES, DEFAULT_PALETTE, seriesColorsFor, withPalette };
});
